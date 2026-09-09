from __future__ import annotations

from decimal import Decimal
from typing import Any

from ..numeric_utils import to_decimal
from ..text_utils import canonical_text, clean_text
from .base_parser import BaseParser, ParseResult


class FlatTableParser(BaseParser):
    """Read a contiguous table from column A without changing source text."""

    def process(self) -> ParseResult:
        header_row = self._header_row()
        names: list[str] = []
        for column in range(1, self.ws.max_column + 1):
            value = self.value(header_row, column)
            if value in (None, ""):
                break
            name = canonical_text(value).lower()
            if not name or name in names:
                raise ValueError(f"Empty or duplicate normalized column: {value!r}")
            names.append(name)
        if not names:
            raise ValueError("Flat table header is empty")
        self.HEADER = tuple(names)
        rows: list[dict[str, Any]] = []
        for row_number in range(header_row + 1, self.ws.max_row + 1):
            values = [self.value(row_number, c) for c in range(1, len(names) + 1)]
            if all(value in (None, "") for value in values):
                if self.config.get("stop_on_empty_row", True):
                    break
                continue
            row = dict(zip(names, values))
            for name in ("nombre", "mois", "annee"):
                if name not in row:
                    continue
                value = to_decimal(row[name])
                if name in ("mois", "annee") and isinstance(value, Decimal):
                    if not value.is_finite() or value != value.to_integral_value():
                        raise ValueError(f"Row {row_number}: {name} must be an integer")
                    value = int(value)
                row[name] = value
            rows.append(row)
        return ParseResult(self.HEADER, rows, self.warnings)

    def _header_row(self) -> int:
        configured = self.config.get("header_row", "first_non_empty")
        if configured not in (None, "", "first_non_empty"):
            return int(configured)
        for row in range(1, self.ws.max_row + 1):
            if any(clean_text(self.value(row, column)) for column in range(1, self.ws.max_column + 1)):
                return row
        raise ValueError("Flat table header is empty")
