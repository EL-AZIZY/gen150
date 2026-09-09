from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any


EXCEL_ERRORS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!"}
_SCIENTIFIC = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)[Ee][+-]?\d+$")


def to_decimal(value: Any) -> Decimal | str | None:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        if stripped.upper() == "NS":
            return "NS"
        if stripped.upper() in EXCEL_ERRORS:
            return stripped.upper()
        normalized = stripped.replace("\u00a0", "").replace(" ", "").replace(",", ".")
        try:
            return Decimal(normalized)
        except InvalidOperation:
            return stripped
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    return str(value)


def format_decimal(value: Any, decimal_places: int | None = None) -> str:
    converted = to_decimal(value)
    if converted is None:
        return ""
    if isinstance(converted, str):
        return converted
    if decimal_places is not None and converted.is_finite():
        quantum = Decimal(1).scaleb(-decimal_places)
        converted = converted.quantize(quantum, rounding=ROUND_HALF_UP)
        if converted == 0:
            converted = abs(converted)
        return format(converted, f".{decimal_places}f")
    if converted == 0:
        return "0"
    rendered = format(converted, "f")
    if "." in rendered:
        rendered = rendered.rstrip("0").rstrip(".")
    return rendered


def scaled_decimal(value: Any, number_format: str | None) -> Decimal | str | None:
    converted = to_decimal(value)
    if not isinstance(converted, Decimal):
        return converted
    if has_thousands_scaling(number_format):
        return converted / Decimal(1000)
    return converted


def percentage_decimal(value: Any, number_format: str | None) -> Decimal | str | None:
    converted = to_decimal(value)
    if not isinstance(converted, Decimal):
        return converted
    if number_format and "%" in number_format:
        return converted * Decimal(100)
    return converted


def has_thousands_scaling(number_format: str | None) -> bool:
    if not number_format:
        return False
    primary = str(number_format).split(";")[0]
    primary = re.sub(r'"[^"]*"', "", primary)
    primary = re.sub(r"\\.", "", primary)
    return bool(re.search(r"[0#?],(?:\s*[^0#?.,%]|\s*$)", primary))


def is_scientific_notation(value: object) -> bool:
    return bool(_SCIENTIFIC.fullmatch(str(value).strip()))
