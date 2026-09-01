from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import yaml

from .sheet_router import normalize_sheet_name


class ConfigError(ValueError):
    """Raised when GEN150 configuration is invalid or incomplete."""


def load_yaml(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as stream:
            content = yaml.safe_load(stream) or {}
    except FileNotFoundError as exc:
        raise ConfigError(f"Configuration file not found: {path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigError(f"Invalid YAML in {path}: {exc}") from exc
    if not isinstance(content, dict):
        raise ConfigError(f"Top-level YAML value must be a mapping: {path}")
    return content


def load_all(config_dir: Path) -> dict[str, dict[str, Any]]:
    return {
        "settings": load_yaml(config_dir / "settings.yml"),
        "families": load_yaml(config_dir / "families.yml"),
        "workbooks": load_yaml(config_dir / "workbooks.yml"),
    }


def select_workbook_config(config: dict[str, Any], file_name: str) -> dict[str, Any]:
    defaults = deepcopy(config.get("defaults", {}))
    matches = [
        item
        for item in config.get("workbooks", [])
        if normalize_sheet_name(str(item.get("file_name", "")))
        == normalize_sheet_name(file_name)
    ]
    if len(matches) > 1:
        raise ConfigError(f"Several workbook configurations match {file_name!r}")
    if not matches:
        if not defaults:
            raise ConfigError(f"No workbook configuration found for {file_name!r}")
        return defaults
    selected = matches[0]
    if selected.get("inherit_defaults", True):
        return deep_merge(defaults, {k: v for k, v in selected.items() if k != "file_name"})
    return deepcopy(selected)


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result

