from __future__ import annotations

from typing import Any

from ..numeric_utils import to_decimal
from ..text_utils import canonical_text, clean_text
from .base_parser import BaseParser, ParseResult
from .pivot_utils import block_rows, discover_blocks, find_cells, period_context


class ProductionMarqueConcessionParser(BaseParser):
    HEADER = ("groupe_concessionnaire", "annee", "marque", "aivam", "sofac", "total", "row_type")

    def process(self) -> ParseResult:
        markers = find_cells(self, lambda value: clean_text(value).upper().startswith("TPR "))
        blocks = discover_blocks(self, ["Étiquettes de lignes"])
        if not markers:
            raise ValueError("No TPR concession group found")
        rows: list[dict[str, Any]] = []
        assigned: set[tuple[int, int]] = set()
        for marker_row, marker_column in markers:
            group = clean_text(self.value(marker_row, marker_column)).split(" ", 1)[1]
            end = min((r for r, _ in markers if r > marker_row), default=self.ws.max_row + 1)
            right = min((c for r, c in markers if r == marker_row and c > marker_column),
                        default=self.ws.max_column + 1)
            candidates = [b for b in blocks if marker_row <= b.header_row < end
                          and marker_column < b.label_column < right]
            if not candidates:
                raise ValueError(f"No pivot table for concession group {group!r}")
            for block in candidates:
                key = (block.header_row, block.label_column)
                if key in assigned:
                    raise ValueError(f"Concession pivot assigned twice: {key}")
                assigned.add(key)
                headers = {canonical_text(self.value(block.header_row, c)): i
                           for i, c in enumerate(block.value_columns)}
                if not {"AIVAM", "SOFAC", "TOTAL_GENERAL"} <= headers.keys():
                    raise ValueError(f"Missing AIVAM/SOFAC/Total général for {group!r}")
                year = period_context(self, block, int(self.config.get("context_lookback", 8)))["annee"]
                if not isinstance(year, int):
                    raise ValueError(f"Missing integer year for concession group {group!r}")
                found_total = False
                for label, values, total in block_rows(self, block, end_row=end - 1, stop_at_total=True):
                    rows.append({
                        "groupe_concessionnaire": group, "annee": year, "marque": label,
                        "aivam": to_decimal(values[headers["AIVAM"]]),
                        "sofac": to_decimal(values[headers["SOFAC"]]),
                        "total": to_decimal(values[headers["TOTAL_GENERAL"]]),
                        "row_type": "total" if total else "marque",
                    })
                    found_total = total
                if not found_total:
                    raise ValueError(f"Missing Total général for concession group {group!r}")
        return ParseResult(self.HEADER, rows, self.warnings)
