from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from ..numeric_utils import EXCEL_ERRORS


class SheetRejected(RuntimeError):
    """Raised when a configured reject policy is triggered."""


@dataclass
class ParseResult:
    header: tuple[str, ...]
    rows: list[dict[str, Any]]
    warnings: list[str] = field(default_factory=list)


@dataclass
class ParserContext:
    worksheet: Worksheet
    formula_worksheet: Worksheet
    sheet_config: dict[str, Any]
    family_config: dict[str, Any]
    error_policy: str
    missing_formula_cache_policy: str
    logger: Any


class BaseParser:
    HEADER: tuple[str, ...] = ()

    def __init__(self, context: ParserContext) -> None:
        self.context = context
        self.ws = context.worksheet
        self.formula_ws = context.formula_worksheet
        self.config = context.sheet_config
        self.warnings: list[str] = []
        self._audited: set[tuple[int, int]] = set()

    def process(self) -> ParseResult:
        raise NotImplementedError

    def value(self, row: int, column: int) -> Any:
        coordinates = (row, column)
        cached_cell = self.ws.cell(row=row, column=column)
        if coordinates not in self._audited:
            self._audited.add(coordinates)
            formula_cell = self.formula_ws.cell(row=row, column=column)
            cached = cached_cell.value
            formula = formula_cell.value
            if isinstance(cached, str) and cached.upper() in EXCEL_ERRORS:
                self._issue(
                    f"Excel error detected: {self.ws.title}!{cached_cell.coordinate}: {cached}",
                    self.context.error_policy,
                )
            if isinstance(formula, str) and formula.startswith("="):
                formula_errors = [error for error in EXCEL_ERRORS if error in formula.upper()]
                if formula_errors:
                    self._issue(
                        f"Excel formula error reference detected: {self.ws.title}!{cached_cell.coordinate}: "
                        f"{','.join(sorted(formula_errors))}",
                        self.context.error_policy,
                    )
                if cached is None:
                    self._issue(
                        f"Missing cached formula value: {self.ws.title}!{cached_cell.coordinate}",
                        self.context.missing_formula_cache_policy,
                    )
        return cached_cell.value

    def values_in_area(self, area: dict[str, int]) -> list[Any]:
        values: list[Any] = []
        for row in range(int(area["first_row"]), int(area["last_row"]) + 1):
            for column in range(
                int(area["first_column"]), int(area["last_column"]) + 1
            ):
                values.append(self.value(row, column))
        return values

    def _issue(self, message: str, policy: str) -> None:
        if policy == "reject":
            raise SheetRejected(message)
        self.warnings.append(message)
        self.context.logger.warning(message)

    @staticmethod
    def first_year(value: object) -> int | None:
        match = re.search(r"\b(20\d{2})\b", str(value or ""))
        return int(match.group(1)) if match else None

