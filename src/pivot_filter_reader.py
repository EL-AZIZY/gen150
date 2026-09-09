from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl.utils.cell import range_boundaries

from .text_utils import canonical_text


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"m": _MAIN_NS}

PivotFilterSelections = dict[tuple[str, int, str], tuple[Any, ...]]


def read_pivot_filter_selections(path: Path) -> PivotFilterSelections:
    """Read visible page-filter items directly from the workbook package."""
    selections: PivotFilterSelections = {}
    with ZipFile(path) as archive:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        workbook_rels = _relationships(archive, "xl/_rels/workbook.xml.rels")
        caches: dict[str, list[tuple[str, tuple[Any, ...]]]] = {}
        archive_names = set(archive.namelist())

        sheets = workbook.find("m:sheets", _NS)
        if sheets is None:
            return selections
        for sheet in sheets:
            relationship_id = sheet.get(f"{{{_OFFICE_REL_NS}}}id")
            if not relationship_id or relationship_id not in workbook_rels:
                continue
            sheet_path = _resolve("xl/workbook.xml", workbook_rels[relationship_id])
            sheet_rels_path = _rels_path(sheet_path)
            if sheet_rels_path not in archive_names:
                continue
            for relationship_id, target in _relationships(archive, sheet_rels_path).items():
                if not relationship_id.startswith("pivotTable:"):
                    continue
                pivot_path = _resolve(sheet_path, target)
                pivot = ElementTree.fromstring(archive.read(pivot_path))
                location = pivot.find("m:location", _NS)
                if location is None or not location.get("ref"):
                    continue
                min_column, _, _, _ = range_boundaries(location.get("ref"))
                cache_path = _pivot_cache_path(archive, pivot_path)
                if cache_path not in caches:
                    caches[cache_path] = _cache_fields(archive, cache_path)
                cache_fields = caches[cache_path]
                pivot_fields = pivot.find("m:pivotFields", _NS)
                page_fields = pivot.find("m:pageFields", _NS)
                if pivot_fields is None or page_fields is None:
                    continue
                pivot_field_list = list(pivot_fields)
                for page_field in page_fields:
                    field_index = int(page_field.get("fld", "-1"))
                    if not (0 <= field_index < len(pivot_field_list)):
                        continue
                    if not (0 <= field_index < len(cache_fields)):
                        continue
                    field_name, shared_items = cache_fields[field_index]
                    visible = _visible_items(pivot_field_list[field_index], shared_items)
                    if visible:
                        key = (sheet.get("name", ""), min_column, canonical_text(field_name))
                        selections[key] = visible
    return selections


def _relationships(archive: ZipFile, path: str) -> dict[str, str]:
    root = ElementTree.fromstring(archive.read(path))
    relationships: dict[str, str] = {}
    for relationship in root:
        relationship_id = relationship.get("Id", "")
        target = relationship.get("Target")
        if not target:
            continue
        if relationship.get("Type", "").endswith("/pivotTable"):
            relationship_id = f"pivotTable:{relationship_id}"
        relationships[relationship_id] = target
    return relationships


def _resolve(source_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    parts: list[str] = []
    for part in (PurePosixPath(source_path).parent / target).parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part not in {"", "."}:
            parts.append(part)
    return "/".join(parts)


def _rels_path(source_path: str) -> str:
    source = PurePosixPath(source_path)
    return str(source.parent / "_rels" / f"{source.name}.rels")


def _pivot_cache_path(archive: ZipFile, pivot_path: str) -> str:
    relationships = _relationships(archive, _rels_path(pivot_path))
    for relationship_id, target in relationships.items():
        if not relationship_id.startswith("pivotTable:"):
            return _resolve(pivot_path, target)
    raise ValueError(f"Pivot table cache relationship not found: {pivot_path}")


def _cache_fields(archive: ZipFile, cache_path: str) -> list[tuple[str, tuple[Any, ...]]]:
    root = ElementTree.fromstring(archive.read(cache_path))
    fields = root.find("m:cacheFields", _NS)
    if fields is None:
        return []
    result: list[tuple[str, tuple[Any, ...]]] = []
    for field in fields:
        shared = field.find("m:sharedItems", _NS)
        values = tuple(_shared_value(item) for item in shared) if shared is not None else ()
        result.append((field.get("name", ""), values))
    return result


def _shared_value(item: ElementTree.Element) -> Any:
    kind = item.tag.rsplit("}", 1)[-1]
    if kind == "m":
        return None
    value = item.get("v")
    if kind != "n" or value is None:
        return value
    try:
        number = Decimal(value)
    except InvalidOperation:
        return value
    return int(number) if number == number.to_integral_value() else number


def _visible_items(
    pivot_field: ElementTree.Element, shared_items: tuple[Any, ...]
) -> tuple[Any, ...]:
    items = pivot_field.find("m:items", _NS)
    if items is None:
        return ()
    visible: list[Any] = []
    for item in items:
        if item.get("h") == "1" or item.get("t") == "default" or item.get("x") is None:
            continue
        index = int(item.get("x", "-1"))
        if 0 <= index < len(shared_items) and shared_items[index] is not None:
            visible.append(shared_items[index])
    return tuple(visible)
