from __future__ import annotations

from typing import Any

from ..numeric_utils import to_decimal
from ..text_utils import canonical_text, clean_text
from .base_parser import BaseParser, ParseResult
from .pivot_utils import block_axis, block_rows, discover_blocks, period_context


class ProductionMarquePivotStagingParser(BaseParser):
    HEADER = (
        "sheet", "bloc_index", "axe", "periodicite", "annee", "mois",
        "dimension", "marque", "ville", "organisme", "valeur",
    )

    def process(self) -> ParseResult:
        blocks = discover_blocks(self, self.config.get("anchor_labels", ["Étiquettes de lignes"]))
        if not blocks:
            raise ValueError("No pivot table anchors found")
        rows: list[dict[str, Any]] = []
        for index, block in enumerate(blocks, start=1):
            period = period_context(self, block, int(self.config.get("context_lookback", 8)))
            selected_months = self.context.pivot_filter_selections.get(
                (self.ws.title, block.label_column, "MOIS")
            )
            if canonical_text(period["mois"]) == "PLUSIEURS_ELEMENTS" and selected_months:
                period["mois"] = ",".join(str(month) for month in selected_months)
            axis = block_axis(self, block)
            axis_key = canonical_text(axis)
            if axis_key.startswith("MOIS"):
                periodicity = "MENSUEL"
            elif axis_key.startswith("ANNEE"):
                periodicity = "ANNUEL"
            else:
                raise ValueError(
                    f"Unknown synthesis axis at {self.ws.title}!"
                    f"R{block.header_row}C{block.label_column}: {axis!r}"
                )
            dimension = (
                "VILLE"
                if canonical_text(self.value(block.header_row, block.label_column)) == "VILLE"
                else "MARQUE"
            )
            headers = [self.value(block.header_row, c) for c in block.value_columns]
            for label, values, _total in block_rows(
                self, block, blank_label_policy=self.config.get("blank_label_policy", "stop"),
            ):
                for header, value in zip(headers, values):
                    rows.append({
                        "sheet": self.ws.title, "bloc_index": index,
                        "axe": clean_text(axis), "periodicite": periodicity,
                        "annee": period["annee"], "mois": period["mois"],
                        "dimension": dimension,
                        "marque": None if dimension == "VILLE" else label,
                        "ville": label if dimension == "VILLE" else period["ville"],
                        "organisme": clean_text(header),
                        "valeur": to_decimal(value),
                    })
        return ParseResult(self.HEADER, rows, self.warnings)
