from __future__ import annotations

import re
import unicodedata


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\u00a0", " ").strip().split())


def canonical_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", clean_text(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.upper().replace("’", "'")
    return re.sub(r"[^A-Z0-9]+", "_", text).strip("_")


def contains_total(value: object) -> bool:
    return bool(re.search(r"(?:^|_)TOTAL(?:_|$)", canonical_text(value)))
