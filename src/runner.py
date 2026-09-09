from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from .config_loader import ConfigError, load_all, select_workbook_config
from .config_validator import validate_configuration
from .exporter import Exporter
from .parsers import FAMILY_REGISTRY
from .parsers.base_parser import ParserContext, SheetRejected
from .pivot_filter_reader import read_pivot_filter_selections
from .settings import Settings
from .sheet_router import SheetRouter
from .validators import validate_rows
from .workbook_loader import load_workbook_pair


LOGGER = logging.getLogger("gen150")


def configure_logging(output_directory: Path, level: str) -> None:
    log_directory = output_directory / "logs"
    log_directory.mkdir(parents=True, exist_ok=True)
    log_path = log_directory / f"gen150_{datetime.now():%Y%m%d}.log"
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    LOGGER.setLevel(getattr(logging, level, logging.INFO))
    LOGGER.handlers.clear()
    for handler in (logging.StreamHandler(), logging.FileHandler(log_path, encoding="utf-8")):
        handler.setFormatter(formatter)
        LOGGER.addHandler(handler)


def discover_inputs(input_argument: Path | None, default_directory: Path) -> list[Path]:
    target = (input_argument or default_directory).resolve()
    if target.is_file():
        if target.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"Unsupported input file: {target}")
        return [target]
    if not target.exists():
        raise ValueError(f"Input path does not exist: {target}")
    return sorted(
        path
        for path in target.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
    )


def process_workbook(
    path: Path,
    families: dict[str, Any],
    workbook_configs: dict[str, Any],
    settings: Settings,
    exporter: Exporter,
) -> dict[str, Any]:
    LOGGER.info("Workbook loaded: %s", path)
    workbook_config = select_workbook_config(workbook_configs, path.name)
    pair = load_workbook_pair(path)
    manifest: dict[str, Any] = {
        "generator": "GEN150",
        "workbook": path.name,
        "started_at": datetime.now(timezone.utc).isoformat(),
        "sheets": [],
    }
    try:
        pivot_filter_selections = read_pivot_filter_selections(path)
        router = SheetRouter(pair.values, families, workbook_config, LOGGER)
        for routed in router.route():
            parser_class = FAMILY_REGISTRY[routed.family_name]
            LOGGER.info(
                "Processing sheet %s with %s", routed.real_name, parser_class.__name__
            )
            entry: dict[str, Any] = {
                "sheet": routed.real_name,
                "family": routed.family_name,
                "parser": parser_class.__name__,
            }
            try:
                parser = parser_class(
                    ParserContext(
                        worksheet=pair.values[routed.real_name],
                        formula_worksheet=pair.formulas[routed.real_name],
                        sheet_config=routed.sheet_config,
                        family_config=routed.family_config,
                        error_policy=settings.excel_error_policy,
                        missing_formula_cache_policy=settings.missing_formula_cache_policy,
                        logger=LOGGER,
                        pivot_filter_selections=pivot_filter_selections,
                    )
                )
                result = parser.process()
                validation_errors = validate_rows(
                    routed.family_name, result.header, result.rows,
                    layout=routed.sheet_config.get("layout", "summary"),
                )
                if validation_errors:
                    rejected_path = exporter.write_csv(
                        path.name,
                        routed.real_name,
                        result.header,
                        result.rows,
                        rejected=True,
                    )
                    exporter.write_rejection(path.name, routed.real_name, validation_errors)
                    entry.update(
                        status="rejected",
                        rows=len(result.rows),
                        output=str(rejected_path),
                        errors=validation_errors,
                        warnings=result.warnings,
                    )
                    LOGGER.error("Sheet rejected: %s", routed.real_name)
                else:
                    csv_path = exporter.write_csv(
                        path.name, routed.real_name, result.header, result.rows
                    )
                    entry.update(
                        status="success",
                        rows=len(result.rows),
                        output=str(csv_path),
                        warnings=result.warnings,
                    )
                    LOGGER.info("CSV generated: %s", csv_path)
            except (SheetRejected, ValueError, KeyError) as exc:
                rejection_path = exporter.write_rejection(
                    path.name, routed.real_name, [str(exc)]
                )
                entry.update(status="rejected", rows=0, output=str(rejection_path), errors=[str(exc)])
                LOGGER.error("Sheet rejected: %s: %s", routed.real_name, exc)
            manifest["sheets"].append(entry)
    finally:
        pair.close()
    manifest["finished_at"] = datetime.now(timezone.utc).isoformat()
    manifest["success_count"] = sum(
        entry["status"] == "success" for entry in manifest["sheets"]
    )
    manifest["rejected_count"] = sum(
        entry["status"] == "rejected" for entry in manifest["sheets"]
    )
    exporter.write_manifest(path.name, manifest)
    return manifest


def run(
    config_dir: Path,
    input_argument: Path | None = None,
    output_argument: Path | None = None,
) -> int:
    configs = load_all(config_dir)
    validate_configuration(configs, FAMILY_REGISTRY)
    settings = Settings.from_mapping(configs["settings"], config_dir.parent)
    if output_argument:
        settings = Settings(
            input_directory=settings.input_directory,
            output_directory=output_argument.resolve(),
            csv_separator=settings.csv_separator,
            csv_encoding=settings.csv_encoding,
            excel_error_policy=settings.excel_error_policy,
            missing_formula_cache_policy=settings.missing_formula_cache_policy,
            log_level=settings.log_level,
        )
    configure_logging(settings.output_directory, settings.log_level)
    LOGGER.info("GEN150 started")
    exporter = Exporter(
        settings.output_directory, settings.csv_separator, settings.csv_encoding
    )
    paths = discover_inputs(input_argument, settings.input_directory)
    if not paths:
        LOGGER.warning("No Excel workbook found")
        return 0
    manifests = [
        process_workbook(
            path,
            configs["families"]["families"],
            configs["workbooks"],
            settings,
            exporter,
        )
        for path in paths
    ]
    return 1 if any(manifest["rejected_count"] for manifest in manifests) else 0


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="GEN150 - convert configured APSF Excel sheets to database-ready CSV files"
    )
    parser.add_argument(
        "input", nargs="?", type=Path, help="Excel workbook or directory (defaults to settings.yml)"
    )
    parser.add_argument(
        "--config-dir", type=Path, default=Path(__file__).parents[1] / "config"
    )
    parser.add_argument("--output-dir", type=Path)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    try:
        return run(args.config_dir.resolve(), args.input, args.output_dir)
    except (ConfigError, ValueError, OSError) as exc:
        print(f"GEN150 error: {exc}", file=sys.stderr)
        return 2

