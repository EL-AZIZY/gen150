from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Any, Callable, Iterator

from ..numeric_utils import to_decimal
from ..text_utils import canonical_text
from .base_parser import BaseParser


@dataclass(frozen=True)
class PivotBlock:
    header_row: int
    label_column: int
    value_columns: tuple[int, ...]


def find_cells(parser: BaseParser, matches: Callable[[Any], bool]) -> list[tuple[int, int]]:
    return [
        (cell.row, cell.column)
        for row in parser.ws.iter_rows()
        for cell in row
        if matches(cell.value)
    ]


def discover_blocks(parser: BaseParser, labels: list[str]) -> list[PivotBlock]:
    """Find contiguous value headers, independently of block positions/widths."""
    accepted = {canonical_text(label) for label in labels}
    anchors = find_cells(parser, lambda value: canonical_text(value) in accepted)
    blocks: list[PivotBlock] = []
    for row, column in anchors:
        parser.value(row, column)
        columns: list[int] = []
        for current in range(column + 1, parser.ws.max_column + 1):
            value = parser.value(row, current)
            if value in (None, "") or canonical_text(value) in accepted:
                break
            columns.append(current)
        # VILLE can also be a period filter. It is a table anchor only when
        # followed by source measure headers, never a city filter value.
        if canonical_text(parser.value(row, column)) == "VILLE":
            if not any(canonical_text(parser.value(row, c)) in {"AIVAM", "SOFAC"} for c in columns):
                continue
        if columns:
            blocks.append(PivotBlock(row, column, tuple(columns)))
    return blocks


def period_context(parser: BaseParser, block: PivotBlock, lookback: int = 8) -> dict[str, Any]:
    context: dict[str, Any] = {"annee": None, "ville": None, "mois": None}
    for row in range(block.header_row - 1, max(0, block.header_row - lookback - 1), -1):
        for column in range(block.label_column, block.value_columns[-1]):
            key = canonical_text(parser.value(row, column)).lower()
            if key not in context or context[key] is not None:
                continue
            value = parser.value(row, column + 1)
            converted = to_decimal(value)
            if isinstance(converted, Decimal) and converted.is_finite() and converted == converted.to_integral_value():
                value = int(converted)
            context[key] = value
    return context


def block_axis(parser: BaseParser, block: PivotBlock) -> Any:
    """Read the first non-empty cell above a block in its label column."""
    for row in range(1, block.header_row):
        value = parser.value(row, block.label_column)
        if value not in (None, "") and canonical_text(value):
            return value
    return None


def block_rows(
    parser: BaseParser, block: PivotBlock, *, end_row: int | None = None,
    stop_at_total: bool = False, blank_label_policy: str = "stop",
) -> Iterator[tuple[Any, list[Any], bool]]:
    """Walk each block separately, preserving labels and all cached measures."""
    if blank_label_policy not in {"stop", "keep_with_values"}:
        raise ValueError(f"Invalid blank_label_policy: {blank_label_policy!r}")
    end = parser.ws.max_row if end_row is None else end_row
    for row in range(block.header_row + 1, end + 1):
        label = parser.value(row, block.label_column)
        values = [parser.value(row, column) for column in block.value_columns]
        if label in (None, ""):
            if blank_label_policy == "stop" and any(value not in (None, "") for value in values):
                parser._issue(
                    f"Pivot block at {parser.ws.title}!R{block.header_row}C{block.label_column} "
                    f"stopped at row {row}: empty label with cached measures; "
                    "blank_label_policy=stop", "warning",
                )
            if blank_label_policy == "stop" or all(value in (None, "") for value in values):
                break
        total = canonical_text(label) == "TOTAL_GENERAL"
        yield label, values, total
        if total and stop_at_total:
            break
