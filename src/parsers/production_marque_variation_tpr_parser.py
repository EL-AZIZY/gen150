from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

from openpyxl.utils.cell import column_index_from_string

from ..numeric_utils import percentage_decimal, to_decimal
from ..text_utils import clean_text
from .base_parser import BaseParser, ParseResult


class ProductionMarqueVariationTprParser(BaseParser):
    HEADER = ("ville", "axe", "valeur")

    def process(self) -> ParseResult:
        tables = self.config.get("tables")
        if not isinstance(tables, list) or not tables:
            raise ValueError("Variation TPR requires at least one configured table")

        prepared: list[tuple[dict[str, Any], list[int], list[date]]] = []
        all_periods: set[date] = set()
        for table in tables:
            header_row = int(table["header_row"])
            value_columns = [
                column_index_from_string(str(value)) for value in table["value_columns"]
            ]
            if len(value_columns) != 2:
                raise ValueError("Variation TPR table requires exactly two period columns")
            periods = [self._period(self.value(header_row, column)) for column in value_columns]
            if periods[0] == periods[1]:
                raise ValueError("Variation TPR period headers must be different")
            prepared.append((table, value_columns, periods))
            all_periods.update(periods)

        if len(all_periods) != 2:
            raise ValueError("Variation TPR tables must share the same two periods")
        current_period = max(all_periods)
        previous_period = min(all_periods)
        current_header = self._period_label(current_period)
        previous_header = self._period_label(previous_period)
        rows: list[dict[str, Any]] = []
        for table, value_columns, periods in prepared:
            rows.extend(self._table_rows(
                table, value_columns, periods,
                current_period, previous_period, current_header, previous_header,
            ))
        return ParseResult(self.HEADER, rows, self.warnings)

    def _table_rows(
        self, table: dict[str, Any], value_columns: list[int], periods: list[date],
        current_period: date, previous_period: date,
        current_header: str, previous_header: str,
    ) -> list[dict[str, Any]]:
        current_index = periods.index(current_period)
        previous_index = periods.index(previous_period)
        ville_column = column_index_from_string(str(table["ville_column"]))
        variation_column = column_index_from_string(str(table["variation_column"]))
        rows: list[dict[str, Any]] = []
        for row in range(int(table["data_start_row"]), int(table["data_end_row"]) + 1):
            ville = self.value(row, ville_column)
            values = [self.value(row, column) for column in value_columns]
            variation = self.value(row, variation_column)
            if all(value in (None, "") for value in (ville, *values, variation)):
                if table.get("skip_empty_rows", False):
                    continue
                break

            current_column = value_columns[current_index]
            previous_column = value_columns[previous_index]
            ville_value = clean_text(ville)
            measures = (
                (current_header, self._two_decimals(percentage_decimal(
                    values[current_index], self.ws.cell(row, current_column).number_format
                ))),
                (previous_header, self._two_decimals(percentage_decimal(
                    values[previous_index], self.ws.cell(row, previous_column).number_format
                ))),
                ("Variation", self._two_decimals(to_decimal(variation))),
            )
            rows.extend(
                {"ville": ville_value, "axe": axe, "valeur": valeur}
                for axe, valeur in measures
            )
        return rows

    @staticmethod
    def _period(value: Any) -> date:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        raise ValueError(f"Variation TPR period header is not an Excel date: {value!r}")

    @staticmethod
    def _period_label(value: date) -> str:
        months = (
            "janv", "févr", "mars", "avr", "mai", "juin",
            "juil", "août", "sept", "oct", "nov", "déc",
        )
        return f"{months[value.month - 1]}-{value.year % 100:02d}"

    @staticmethod
    def _two_decimals(value: Decimal | str | None) -> Decimal | str | None:
        if isinstance(value, Decimal) and value.is_finite():
            return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return value
