from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


@dataclass
class WorkbookPair:
    values: Workbook
    formulas: Workbook

    def close(self) -> None:
        self.values.close()
        self.formulas.close()


def load_workbook_pair(path: Path) -> WorkbookPair:
    common = {"filename": path, "read_only": False, "keep_links": True}
    return WorkbookPair(
        values=load_workbook(data_only=True, **common),
        formulas=load_workbook(data_only=False, **common),
    )

