from __future__ import annotations

from typing import Any

from openpyxl.utils.cell import column_index_from_string

from ..date_utils import extract_reporting_date, normalize_reporting_date
from ..numeric_utils import percentage_decimal, scaled_decimal
from ..text_utils import canonical_text, clean_text, contains_total
from .base_parser import BaseParser, ParseResult


class ApsfCompetitorQuarterlyParser(BaseParser):
    HEADER = (
        "date_reporting",
        "section",
        "organisme",
        "categorie",
        "annee",
        "t1_mdh",
        "t2_mdh",
        "t3_mdh",
        "t4_mdh",
        "marche_t1_mdh",
        "marche_t2_mdh",
        "marche_t3_mdh",
        "marche_t4_mdh",
        "t1_pct",
        "t2_pct",
        "t3_pct",
        "t4_pct",
        "marche_t1_pct",
        "marche_t2_pct",
        "marche_t3_pct",
        "marche_t4_pct",
    )
    EXPECTED_QUARTERS = ("T1", "T2", "T3", "T4")
    RATIO_SECTION = "CES_SUR_ENCOURS_BRUT"

    def process(self) -> ParseResult:
        date_reporting = self._date_reporting()
        year = self._configured_year()
        label_column = column_index_from_string(self.config["label_column"])
        market_label = canonical_text(self.config.get("market_label", "Marché"))
        market_share_label = canonical_text(
            self.config.get("market_share_label", "Part de marché")
        )
        rows: list[dict[str, Any]] = []

        for section, section_config in self.config["sections"].items():
            start = int(section_config["data_start_row"])
            end = int(section_config["data_end_row"])
            period_row = int(section_config["period_row"])
            for category, configured_columns in self.config["category_blocks"].items():
                columns = self._validated_quarter_columns(
                    section, category, configured_columns, period_row
                )
                market_values: tuple[Any, Any, Any, Any] | None = None
                organization_values: list[
                    tuple[str, tuple[Any, Any, Any, Any]]
                ] = []

                for row_number in range(start, end + 1):
                    label = clean_text(self.value(row_number, label_column))
                    canonical = canonical_text(label)
                    if (
                        not canonical
                        or contains_total(label)
                        or market_share_label in canonical
                    ):
                        continue
                    values = self._read_quarters(
                        row_number, columns, section == self.RATIO_SECTION
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
                market_values = market_values or (None, None, None, None)
                for organisme, values in organization_values:
                    rows.append(
                        self._build_row(
                            date_reporting,
                            section,
                            organisme,
                            category,
                            year,
                            values,
                            market_values,
                        )
                    )

        return ParseResult(self.HEADER, rows, self.warnings)

    def _date_reporting(self) -> str:
        configured = self.config.get("date_reporting")
        if configured not in (None, ""):
            return normalize_reporting_date(configured)
        title_area = self.config.get(
            "title_area",
            {"first_row": 1, "last_row": 8, "first_column": 1, "last_column": 12},
        )
        try:
            return extract_reporting_date(self.values_in_area(title_area))
        except ValueError:
            # B2 in the supplied quarterly sheet is a caption without a date.
            # The explicit reporting year uses December 31 as a convention;
            # other families keep their existing date extraction behavior.
            return f"{self._configured_year():04d}-12-31"

    def _configured_year(self) -> int:
        value = self.config.get("annee")
        try:
            year = int(value)
        except (TypeError, ValueError) as exc:
            raise ValueError("Quarterly configuration must define a valid 'annee'") from exc
        if not 1900 <= year <= 2100:
            raise ValueError(f"Quarterly year out of range: {year}")
        return year

    def _validated_quarter_columns(
        self,
        section: str,
        category: str,
        configured_columns: list[str],
        period_row: int,
    ) -> tuple[int, int, int, int]:
        if len(configured_columns) != 4:
            raise ValueError(
                f"Quarterly block must contain four columns for "
                f"section={section}, category={category}"
            )
        columns = tuple(
            column_index_from_string(column) for column in configured_columns
        )
        actual = tuple(
            canonical_text(self.value(period_row, column)) for column in columns
        )
        if actual != self.EXPECTED_QUARTERS:
            raise ValueError(
                f"Invalid quarterly headers for section={section}, category={category}: "
                f"expected {self.EXPECTED_QUARTERS}, got {actual}"
            )
        return columns

    def _read_quarters(
        self,
        row: int,
        columns: tuple[int, int, int, int],
        is_ratio: bool,
    ) -> tuple[Any, Any, Any, Any]:
        cells = [self.ws.cell(row=row, column=column) for column in columns]
        values = [self.value(row, cell.column) for cell in cells]
        converter = percentage_decimal if is_ratio else scaled_decimal
        return tuple(
            converter(value, cell.number_format)
            for value, cell in zip(values, cells)
        )

    def _build_row(
        self,
        date_reporting: str,
        section: str,
        organisme: str,
        category: str,
        year: int,
        values: tuple[Any, Any, Any, Any],
        market_values: tuple[Any, Any, Any, Any],
    ) -> dict[str, Any]:
        row = {field: None for field in self.HEADER}
        row.update(
            {
                "date_reporting": date_reporting,
                "section": section,
                "organisme": organisme,
                "categorie": category,
                "annee": year,
            }
        )
        if section == self.RATIO_SECTION:
            for index, value in enumerate(values, start=1):
                row[f"t{index}_pct"] = value
            for index, value in enumerate(market_values, start=1):
                row[f"marche_t{index}_pct"] = value
        else:
            for index, value in enumerate(values, start=1):
                row[f"t{index}_mdh"] = value
            for index, value in enumerate(market_values, start=1):
                row[f"marche_t{index}_mdh"] = value
        return row
