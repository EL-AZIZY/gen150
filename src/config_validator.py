from __future__ import annotations

from typing import Any

from openpyxl.utils.cell import column_index_from_string

from .config_loader import ConfigError
from .sheet_router import normalize_sheet_name


ALLOWED_POLICIES = {"warning", "reject"}


def validate_configuration(configs: dict[str, dict[str, Any]], registry: dict[str, type]) -> None:
    settings = configs["settings"].get("excel", {})
    for key in ("error_policy", "missing_formula_cache_policy"):
        value = str(settings.get(key, "warning")).lower()
        if value not in ALLOWED_POLICIES:
            raise ConfigError(f"excel.{key} must be warning or reject")

    families = configs["families"].get("families")
    if not isinstance(families, dict) or not families:
        raise ConfigError("families.yml must define a non-empty 'families' mapping")

    owners: dict[str, str] = {}
    for family_name, family in families.items():
        if family_name not in registry:
            raise ConfigError(f"No registered parser for family {family_name!r}")
        sheets = family.get("family_sheets")
        if not isinstance(sheets, list) or not sheets:
            raise ConfigError(f"Family {family_name!r} must define non-empty family_sheets")
        expected_parser = registry[family_name].__name__
        if family.get("parser") != expected_parser:
            raise ConfigError(
                f"Family {family_name!r} declares {family.get('parser')!r}; expected {expected_parser!r}"
            )
        for sheet in sheets:
            normalized = normalize_sheet_name(str(sheet))
            if not normalized:
                raise ConfigError(f"Family {family_name!r} contains an empty sheet name")
            if normalized in owners:
                raise ConfigError(
                    f"Sheet {sheet!r} belongs to both {owners[normalized]!r} and {family_name!r}"
                )
            owners[normalized] = family_name

    defaults = configs["workbooks"].get("defaults", {})
    _validate_sheet_structures(defaults.get("sheets", {}), families)
    for workbook in configs["workbooks"].get("workbooks", []):
        if not workbook.get("file_name"):
            raise ConfigError("Every workbooks entry must have file_name")
        _validate_sheet_structures(workbook.get("sheets", {}), families, partial=True)


def _validate_sheet_structures(
    sheets: dict[str, Any], families: dict[str, Any], partial: bool = False
) -> None:
    if not isinstance(sheets, dict):
        raise ConfigError("Workbook sheets configuration must be a mapping")
    allowed = {
        normalize_sheet_name(sheet): family_name
        for family_name, family in families.items()
        for sheet in family["family_sheets"]
    }
    for sheet_name, structure in sheets.items():
        normalized = normalize_sheet_name(sheet_name)
        declared_family = structure.get("family")
        if normalized not in allowed:
            raise ConfigError(f"Workbook structure declares unowned sheet {sheet_name!r}")
        if declared_family != allowed[normalized]:
            raise ConfigError(
                f"Sheet {sheet_name!r} belongs to {allowed[normalized]!r}, not {declared_family!r}"
            )
        limits = structure.get("limits")
        if limits and (
            int(limits["first_row"]) > int(limits["last_row"])
            or int(limits["first_column"]) > int(limits["last_column"])
        ):
            raise ConfigError(f"Invalid limits for sheet {sheet_name!r}")
        label_column = structure.get("label_column")
        if label_column:
            column_index_from_string(str(label_column))

