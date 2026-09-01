from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from typing import Any

from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

from ..date_utils import extract_reporting_date
from ..numeric_utils import percentage_decimal, to_decimal
from ..text_utils import canonical_text, clean_text, contains_total
from .base_parser import BaseParser, ParseResult


class ApsfActivityParser(BaseParser):
    HEADER = (
        "organisme",
        "date_reporting",
        "section",
        "type_pret",
        "categorie",
        "produit",
        "annee_reference",
        "annee_comparaison",
        "montant_reference_mad",
        "montant_comparaison_mad",
        "dossiers_reference",
        "dossiers_comparaison",
        "variation_montant_mad",
        "variation_montant_pct",
        "variation_dossiers_unites",
        "variation_dossiers_pct",
        "ordre_section",
        "ordre_ligne",
    )

    def process(self) -> ParseResult:
        limits = self.config["limits"]
        title_area = self.config.get(
            "title_area",
            {
                "first_row": limits["first_row"],
                "last_row": min(int(limits["last_row"]), 15),
                "first_column": limits["first_column"],
                "last_column": limits["last_column"],
            },
        )
        date_reporting = extract_reporting_date(self.values_in_area(title_area))
        reference_year, comparison_year = self._extract_years(title_area)
        organisme = str(self.config.get("organisme") or self.ws.title).strip()
        label_column = column_index_from_string(self.config.get("label_column", "A"))
        data_columns = {
            name: column_index_from_string(column)
            for name, column in self.config["data_columns"].items()
        }

        hierarchy = self.config["hierarchy"]
        type_markers = self._alias_map(hierarchy.get("types", {}))
        category_markers = self._alias_map(hierarchy.get("categories", {}))
        product_markers = self._alias_map(hierarchy.get("products", {}))
        skip_labels = {canonical_text(item) for item in self.config.get("skip_labels", [])}
        unit_tokens = {canonical_text(item) for item in self.config.get("unit_tokens", [])}
        section_aliases = self._alias_map(self.config.get("sections", {}))

        current_type = ""
        current_category = ""
        current_product = ""
        current_section = ""
        section_order: dict[str, int] = {}
        line_order: defaultdict[str, int] = defaultdict(int)
        rows: list[dict[str, Any]] = []

        for row_number in range(int(limits["first_row"]), int(limits["last_row"]) + 1):
            raw_label = self.value(row_number, label_column)
            label = clean_text(raw_label)
            canonical = canonical_text(label)
            if not canonical:
                continue

            raw_values = {
                name: self.value(row_number, column) for name, column in data_columns.items()
            }
            has_data = any(value not in (None, "") for value in raw_values.values())

            if canonical in type_markers:
                current_type = type_markers[canonical]
                current_category = ""
                current_product = ""
                continue
            if contains_total(label) or canonical in skip_labels:
                continue
            if any(token and token in canonical for token in unit_tokens):
                continue
            if canonical in section_aliases:
                current_section = section_aliases[canonical]
                continue
            if canonical in category_markers:
                current_category = category_markers[canonical]
                current_product = ""
                if not has_data:
                    continue
            if not has_data:
                # A non-numeric presentation line is treated as a section only when
                # explicitly enabled. This prevents accidental sections from titles/units.
                if self.config.get("infer_sections", False) and row_number > int(
                    title_area["last_row"]
                ):
                    current_section = canonical
                continue

            category = current_category
            if canonical in category_markers:
                product = ""
            elif canonical in product_markers:
                product = product_markers[canonical]
                current_product = product
            elif current_product == "LOA":
                product = canonical if canonical.startswith("LOA_") else f"LOA_{canonical}"
                category = category or "VEHICULES"
            else:
                product = canonical
            section = current_section or str(self.config.get("default_section", "ACTIVITE"))
            if section not in section_order:
                section_order[section] = len(section_order) + 1
            line_order[section] += 1

            row = {
                "organisme": organisme,
                "date_reporting": date_reporting,
                "section": section,
                "type_pret": current_type,
                "categorie": category,
                "produit": product,
                "annee_reference": reference_year,
                "annee_comparaison": comparison_year,
                "ordre_section": section_order[section],
                "ordre_ligne": line_order[section],
            }
            for field, value in raw_values.items():
                cell = self.ws.cell(row=row_number, column=data_columns[field])
                if field.endswith("_pct"):
                    row[field] = percentage_decimal(value, cell.number_format)
                else:
                    row[field] = to_decimal(value)
            rows.append({field: row.get(field) for field in self.HEADER})

        return ParseResult(self.HEADER, rows, self.warnings)

    def _extract_years(self, title_area: dict[str, int]) -> tuple[int | str, int | str]:
        configured_reference = self.config.get("annee_reference")
        configured_comparison = self.config.get("annee_comparaison")
        if configured_reference and configured_comparison:
            return int(configured_reference), int(configured_comparison)
        year_cells = self.config.get("year_cells", {})
        if year_cells.get("reference") and year_cells.get("comparison"):
            reference = self._year_from_coordinate(year_cells["reference"])
            comparison = self._year_from_coordinate(year_cells["comparison"])
            if reference is None or comparison is None:
                raise ValueError(
                    f"Invalid activity years in {year_cells['reference']}/{year_cells['comparison']}"
                )
            return reference, comparison
        years: list[int] = []
        for value in self.values_in_area(title_area):
            year = self.first_year(value)
            if year is not None and year not in years:
                years.append(year)
        if len(years) < 2:
            raise ValueError(
                "Reference and comparison years not found; configure annee_reference and "
                "annee_comparaison or include both in title_area"
            )
        return years[0], years[1]

    def _year_from_coordinate(self, coordinate: str) -> int | None:
        row, column = coordinate_to_tuple(str(coordinate))
        value = self.value(row, column)
        if isinstance(value, (int, float)) and 1900 <= int(value) <= 2100:
            return int(value)
        return self.first_year(value)

    @staticmethod
    def _alias_map(config: dict[str, Any]) -> dict[str, str]:
        result: dict[str, str] = {}
        for target, aliases in config.items():
            result[canonical_text(target)] = target
            for alias in aliases or []:
                result[canonical_text(alias)] = target
        return result
