from __future__ import annotations

from decimal import Decimal
from typing import Any

from openpyxl.utils.cell import column_index_from_string

from ..date_utils import extract_reporting_date, parse_period
from ..numeric_utils import percentage_decimal, scaled_decimal
from ..text_utils import canonical_text, clean_text, contains_total
from .base_parser import BaseParser, ParseResult


class ApsfCompetitorSummaryParser(BaseParser):
    HEADER = (
        "date_reporting",
        "section",
        "organisme",
        "categorie",
        "periode_reference",
        "periode_comparaison",
        "valeur_reference_mdh",
        "valeur_comparaison_mdh",
        "evolution_pct",
        "marche_reference_mdh",
        "marche_comparaison_mdh",
        "evolution_marche_pct",
        "ratio_reference_pct",
        "ratio_comparaison_pct",
        "evolution_ratio_points",
        "marche_ratio_reference_pct",
        "marche_ratio_comparaison_pct",
        "evolution_marche_ratio_points",
    )
    RATIO_SECTION = "CES_SUR_ENCOURS_BRUT"

    def process(self) -> ParseResult:
        title_area = self.config.get(
            "title_area",
            {"first_row": 1, "last_row": 8, "first_column": 1, "last_column": 12},
        )
        date_reporting = extract_reporting_date(self.values_in_area(title_area))
        label_column = column_index_from_string(self.config["label_column"])
        market_label = canonical_text(self.config.get("market_label", "Marché"))
        market_share_label = canonical_text(
            self.config.get("market_share_label", "Part de marché")
        )
        rows: list[dict[str, Any]] = []

        for section, section_config in self.config["sections"].items():
            start = int(section_config["data_start_row"])
            end = int(section_config["data_end_row"])
            for category, columns in self.config["category_blocks"].items():
                reference_column, comparison_column, evolution_column = (
                    column_index_from_string(column) for column in columns
                )
                period_row = int(section_config["period_row"])
                reference_period = parse_period(self.value(period_row, reference_column))
                comparison_period = parse_period(self.value(period_row, comparison_column))
                market_values: tuple[Any, Any, Any] | None = None
                organization_values: list[tuple[str, tuple[Any, Any, Any]]] = []

                for row_number in range(start, end + 1):
                    label = clean_text(self.value(row_number, label_column))
                    canonical = canonical_text(label)
                    if not canonical or contains_total(label) or market_share_label in canonical:
                        continue
                    values = self._read_triplet(
                        row_number,
                        reference_column,
                        comparison_column,
                        evolution_column,
                        section == self.RATIO_SECTION,
                    )
                    if canonical == market_label:
                        market_values = values
                        continue
                    if any(value not in (None, "") for value in values):
                        organization_values.append((label, values))

                if organization_values and market_values is None and self.config.get(
                    "require_market_context", True
                ):
                    raise ValueError(
                        f"Market row missing for section={section}, category={category}"
                    )
                market_values = market_values or (None, None, None)
                for organisme, values in organization_values:
                    rows.append(
                        self._build_row(
                            date_reporting,
                            section,
                            organisme,
                            category,
                            reference_period,
                            comparison_period,
                            values,
                            market_values,
                        )
                    )

        return ParseResult(self.HEADER, rows, self.warnings)

    def _read_triplet(
        self,
        row: int,
        reference_column: int,
        comparison_column: int,
        evolution_column: int,
        is_ratio: bool,
    ) -> tuple[Any, Any, Any]:
        cells = [
            self.ws.cell(row=row, column=reference_column),
            self.ws.cell(row=row, column=comparison_column),
            self.ws.cell(row=row, column=evolution_column),
        ]
        values = [self.value(row, cell.column) for cell in cells]
        if is_ratio:
            return tuple(
                percentage_decimal(value, cell.number_format)
                for value, cell in zip(values, cells)
            )
        return (
            scaled_decimal(values[0], cells[0].number_format),
            scaled_decimal(values[1], cells[1].number_format),
            percentage_decimal(values[2], cells[2].number_format),
        )

    def _build_row(
        self,
        date_reporting: str,
        section: str,
        organisme: str,
        category: str,
        reference_period: str,
        comparison_period: str,
        values: tuple[Any, Any, Any],
        market_values: tuple[Any, Any, Any],
    ) -> dict[str, Any]:
        row = {field: None for field in self.HEADER}
        row.update(
            {
                "date_reporting": date_reporting,
                "section": section,
                "organisme": organisme,
                "categorie": category,
                "periode_reference": reference_period,
                "periode_comparaison": comparison_period,
            }
        )
        if section == self.RATIO_SECTION:
            row.update(
                {
                    "ratio_reference_pct": values[0],
                    "ratio_comparaison_pct": values[1],
                    "evolution_ratio_points": values[2],
                    "marche_ratio_reference_pct": market_values[0],
                    "marche_ratio_comparaison_pct": market_values[1],
                    "evolution_marche_ratio_points": market_values[2],
                }
            )
        else:
            row.update(
                {
                    "valeur_reference_mdh": values[0],
                    "valeur_comparaison_mdh": values[1],
                    "evolution_pct": values[2],
                    "marche_reference_mdh": market_values[0],
                    "marche_comparaison_mdh": market_values[1],
                    "evolution_marche_pct": market_values[2],
                }
            )
        return row

