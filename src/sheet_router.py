from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterator

from openpyxl.workbook.workbook import Workbook


def normalize_sheet_name(value: str) -> str:
    return " ".join(str(value).strip().split()).casefold()


def slugify_name(value: str) -> str:
    ascii_value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", ascii_value.casefold()).strip("_")


def build_sheet_index(workbook: Workbook) -> dict[str, str]:
    index: dict[str, str] = {}
    for original in workbook.sheetnames:
        normalized = normalize_sheet_name(original)
        if normalized in index:
            raise ValueError(
                f"Ambiguous workbook sheet names after normalization: {index[normalized]!r}, {original!r}"
            )
        index[normalized] = original
    return index


def find_sheet_name(workbook: Workbook, configured_name: str) -> str | None:
    return build_sheet_index(workbook).get(normalize_sheet_name(configured_name))


@dataclass(frozen=True)
class RoutedSheet:
    family_name: str
    configured_name: str
    real_name: str
    family_config: dict[str, Any]
    sheet_config: dict[str, Any]


class SheetRouter:
    def __init__(
        self,
        workbook: Workbook,
        families: dict[str, Any],
        workbook_config: dict[str, Any],
        logger: Any,
    ) -> None:
        self.workbook = workbook
        self.families = families
        self.workbook_config = workbook_config
        self.logger = logger
        self.sheet_index = build_sheet_index(workbook)

    def route(self) -> Iterator[RoutedSheet]:
        processed: set[str] = set()
        configured: set[str] = set()
        structures = {
            normalize_sheet_name(name): value
            for name, value in self.workbook_config.get("sheets", {}).items()
        }

        for family_name, family_config in self.families.items():
            if not family_config.get("enabled", True):
                continue
            family_sheets = family_config["family_sheets"]
            self.logger.info(
                "Family %s: %d configured sheets", family_name, len(family_sheets)
            )
            for configured_name in family_sheets:
                normalized = normalize_sheet_name(configured_name)
                configured.add(normalized)
                real_name = self.sheet_index.get(normalized)
                if real_name is None:
                    self.logger.warning("Configured sheet not found: %s", configured_name)
                    continue
                worksheet = self.workbook[real_name]
                if worksheet.sheet_state != "visible":
                    self.logger.info("Sheet skipped: hidden: %s", real_name)
                    continue
                if normalized in processed:
                    raise ValueError(f"Sheet routed more than once: {real_name}")
                sheet_config = structures.get(normalized)
                if not sheet_config:
                    self.logger.warning("Configured sheet has no structure: %s", real_name)
                    continue
                if sheet_config.get("family") != family_name:
                    raise ValueError(
                        f"Sheet {real_name!r} has family {sheet_config.get('family')!r}, expected {family_name!r}"
                    )
                processed.add(normalized)
                yield RoutedSheet(
                    family_name=family_name,
                    configured_name=configured_name,
                    real_name=real_name,
                    family_config=family_config,
                    sheet_config=sheet_config,
                )

        for real_name in self.workbook.sheetnames:
            normalized = normalize_sheet_name(real_name)
            if normalized not in configured:
                self.logger.info("Sheet skipped: not configured in any family: %s", real_name)

