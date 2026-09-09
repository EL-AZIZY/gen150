from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any, Iterable

from .numeric_utils import format_decimal
from .sheet_router import slugify_name


def decimal_places_for_field(field: str) -> int | None:
    if field.endswith(("_pct", "_points")):
        return 1
    if field.startswith("dossiers_") or field.endswith(("_mad", "_mdh", "_unites")):
        return 0
    return None


class Exporter:
    def __init__(self, output_directory: Path, separator: str, encoding: str) -> None:
        self.output_directory = output_directory
        self.separator = separator
        self.encoding = encoding
        for name in ("success", "rejected", "logs", "manifests"):
            (output_directory / name).mkdir(parents=True, exist_ok=True)

    def csv_name(self, workbook_name: str, sheet_name: str) -> str:
        return f"{slugify_name(Path(workbook_name).stem)}__{slugify_name(sheet_name)}.csv"

    def write_csv(
        self,
        workbook_name: str,
        sheet_name: str,
        header: Iterable[str],
        rows: list[dict[str, Any]],
        rejected: bool = False,
    ) -> Path:
        folder = "rejected" if rejected else "success"
        csv_name = self.csv_name(workbook_name, sheet_name)
        path = self.output_directory / folder / csv_name
        if rejected:
            (self.output_directory / "success" / csv_name).unlink(missing_ok=True)
        else:
            (self.output_directory / "rejected" / csv_name).unlink(missing_ok=True)
            self._rejection_path(workbook_name, sheet_name).unlink(missing_ok=True)
        header_tuple = tuple(header)
        with path.open("w", encoding=self.encoding, newline="") as stream:
            writer = csv.DictWriter(
                stream, fieldnames=header_tuple, delimiter=self.separator, extrasaction="raise"
            )
            writer.writeheader()
            for row in rows:
                writer.writerow(
                    {
                        field: format_decimal(
                            row.get(field), decimal_places_for_field(field)
                        )
                        for field in header_tuple
                    }
                )
        return path

    def write_rejection(self, workbook_name: str, sheet_name: str, reasons: list[str]) -> Path:
        csv_name = self.csv_name(workbook_name, sheet_name)
        (self.output_directory / "success" / csv_name).unlink(missing_ok=True)
        path = self._rejection_path(workbook_name, sheet_name)
        path.write_text(
            json.dumps({"workbook": workbook_name, "sheet": sheet_name, "reasons": reasons},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return path

    def _rejection_path(self, workbook_name: str, sheet_name: str) -> Path:
        return self.output_directory / "rejected" / (
            f"{slugify_name(Path(workbook_name).stem)}__{slugify_name(sheet_name)}.json"
        )

    def write_manifest(self, workbook_name: str, content: dict[str, Any]) -> Path:
        path = self.output_directory / "manifests" / (
            f"{slugify_name(Path(workbook_name).stem)}.json"
        )
        path.write_text(json.dumps(content, ensure_ascii=False, indent=2), encoding="utf-8")
        return path
