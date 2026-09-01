from __future__ import annotations

import calendar
import re
import unicodedata
from datetime import date, datetime


MONTHS = {
    "JANVIER": 1,
    "JAN": 1,
    "FEVRIER": 2,
    "FEV": 2,
    "MARS": 3,
    "MAR": 3,
    "AVRIL": 4,
    "AVR": 4,
    "MAI": 5,
    "JUIN": 6,
    "JUILLET": 7,
    "JUIL": 7,
    "AOUT": 8,
    "SEPTEMBRE": 9,
    "SEPT": 9,
    "OCTOBRE": 10,
    "OCT": 10,
    "NOVEMBRE": 11,
    "NOV": 11,
    "DECEMBRE": 12,
    "DEC": 12,
}


def _ascii_upper(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(char for char in text if not unicodedata.combining(char)).upper()


def extract_reporting_date(values: list[object]) -> str:
    for value in values:
        if isinstance(value, (datetime, date)):
            return date(
                value.year,
                value.month,
                calendar.monthrange(value.year, value.month)[1],
            ).isoformat()
    joined = " ".join(str(value) for value in values if value not in (None, ""))
    normalized = _ascii_upper(joined)
    match = re.search(
        r"\b(" + "|".join(sorted(MONTHS, key=len, reverse=True)) + r")[\s./-]+(20\d{2})\b",
        normalized,
    )
    if not match:
        raise ValueError("Reporting month and year not found in configured title area")
    month = MONTHS[match.group(1)]
    year = int(match.group(2))
    return date(year, month, calendar.monthrange(year, month)[1]).isoformat()


def parse_period(value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return f"{value.year:04d}-{value.month:02d}"
    text = _ascii_upper(value).strip()
    match = re.search(
        r"\b(" + "|".join(sorted(MONTHS, key=len, reverse=True)) + r")[\s./-]+(\d{2}|20\d{2})\b",
        text,
    )
    if not match:
        raise ValueError(f"Unsupported period: {value!r}")
    year = int(match.group(2))
    if year < 100:
        year += 2000
    return f"{year:04d}-{MONTHS[match.group(1)]:02d}"
