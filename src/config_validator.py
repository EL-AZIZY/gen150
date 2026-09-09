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
        enabled_families = workbook.get("enabled_families")
        if enabled_families is not None:
            if not isinstance(enabled_families, list) or not enabled_families:
                raise ConfigError("enabled_families must be a non-empty list")
            unknown = set(enabled_families) - set(families)
            if unknown:
                raise ConfigError(f"Unknown enabled_families: {sorted(unknown)!r}")
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
        if declared_family == "production_marque_variation_tpr":
            tables = structure.get("tables")
            if not isinstance(tables, list) or not tables:
                raise ConfigError(f"Sheet {sheet_name!r} requires Variation TPR tables")
            for table in tables:
                for key in ("header_row", "data_start_row", "data_end_row"):
                    if key not in table:
                        raise ConfigError(f"Sheet {sheet_name!r} table requires {key}")
                if int(table["data_start_row"]) > int(table["data_end_row"]):
                    raise ConfigError(f"Sheet {sheet_name!r} has invalid table rows")
                for key in ("ville_column", "variation_column"):
                    column_index_from_string(str(table[key]))
                value_columns = table.get("value_columns", [])
                if len(value_columns) != 2:
                    raise ConfigError(f"Sheet {sheet_name!r} requires two period columns")
                for column in value_columns:
                    column_index_from_string(str(column))
