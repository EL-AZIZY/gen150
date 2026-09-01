from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Settings:
    input_directory: Path
    output_directory: Path
    csv_separator: str = ";"
    csv_encoding: str = "utf-8-sig"
    excel_error_policy: str = "warning"
    missing_formula_cache_policy: str = "warning"
    log_level: str = "INFO"

    @classmethod
    def from_mapping(cls, data: dict[str, Any], base_dir: Path) -> "Settings":
        paths = data.get("paths", {})
        csv = data.get("csv", {})
        excel = data.get("excel", {})
        logging = data.get("logging", {})
        return cls(
            input_directory=_resolve(base_dir, paths.get("input", "input")),
            output_directory=_resolve(base_dir, paths.get("output", "output")),
            csv_separator=str(csv.get("separator", ";")),
            csv_encoding=str(csv.get("encoding", "utf-8-sig")),
            excel_error_policy=str(excel.get("error_policy", "warning")).lower(),
            missing_formula_cache_policy=str(
                excel.get("missing_formula_cache_policy", "warning")
            ).lower(),
            log_level=str(logging.get("level", "INFO")).upper(),
        )


def _resolve(base_dir: Path, value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base_dir / path).resolve()

