from __future__ import annotations

from collections import defaultdict
from typing import Any, Iterable

from .numeric_utils import EXCEL_ERRORS, is_scientific_notation
from .text_utils import canonical_text, contains_total


ACTIVITY_HEADER = (
    "organisme", "date_reporting", "section", "type_pret", "categorie", "produit",
    "annee_reference", "annee_comparaison", "montant_reference_mad",
    "montant_comparaison_mad", "dossiers_reference", "dossiers_comparaison",
    "variation_montant_mad", "variation_montant_pct", "variation_dossiers_unites",
    "variation_dossiers_pct", "ordre_section", "ordre_ligne",
)
SUMMARY_HEADER = (
    "date_reporting", "section", "organisme", "categorie", "periode_reference",
    "periode_comparaison", "valeur_reference_mdh", "valeur_comparaison_mdh",
    "evolution_pct", "marche_reference_mdh", "marche_comparaison_mdh",
    "evolution_marche_pct", "ratio_reference_pct", "ratio_comparaison_pct",
    "evolution_ratio_points", "marche_ratio_reference_pct",
    "marche_ratio_comparaison_pct", "evolution_marche_ratio_points",
)
QUARTERLY_HEADER = (
    "date_reporting", "section", "organisme", "categorie", "annee",
    "t1_mdh", "t2_mdh", "t3_mdh", "t4_mdh",
    "marche_t1_mdh", "marche_t2_mdh", "marche_t3_mdh", "marche_t4_mdh",
    "t1_pct", "t2_pct", "t3_pct", "t4_pct",
    "marche_t1_pct", "marche_t2_pct", "marche_t3_pct", "marche_t4_pct",
)

FAMILY_HEADERS = {
    "apsf_activity": ACTIVITY_HEADER,
    "apsf_competitor_summary": SUMMARY_HEADER,
    "apsf_competitor_quarterly": QUARTERLY_HEADER,
}


def validate_rows(family: str, header: Iterable[str], rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    expected = FAMILY_HEADERS.get(family)
    if expected is None:
        return [f"Unknown family validator: {family}"]
    header_tuple = tuple(header)
    if header_tuple != expected:
        errors.append(f"Header mismatch: expected {expected!r}, got {header_tuple!r}")
    for index, row in enumerate(rows, start=1):
        if tuple(row.keys()) != expected:
            missing = set(expected) - set(row)
            extra = set(row) - set(expected)
            errors.append(f"Row {index}: unexpected schema; missing={missing}, extra={extra}")
        if contains_total(row.get("organisme", "")) or contains_total(row.get("produit", "")):
            errors.append(f"Row {index}: TOTAL line exported")
        for field, value in row.items():
            if is_scientific_notation(value):
                errors.append(f"Row {index}, {field}: scientific notation")
    if family == "apsf_activity":
        errors.extend(_validate_activity(rows))
    elif family == "apsf_competitor_summary":
        errors.extend(_validate_summary(rows))
    elif family == "apsf_competitor_quarterly":
        errors.extend(_validate_quarterly(rows))
    return errors


def _validate_activity(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    grouped: defaultdict[tuple[Any, Any], list[int]] = defaultdict(list)
    keys: set[tuple[Any, ...]] = set()
    for index, row in enumerate(rows, start=1):
        grouped[(row["organisme"], row["section"])].append(int(row["ordre_ligne"]))
        key = (
            row["organisme"], row["date_reporting"], row["section"], row["type_pret"],
            row["categorie"], row["produit"], row["ordre_ligne"],
        )
        if key in keys:
            errors.append(f"Row {index}: duplicate business key")
        keys.add(key)
    for group, order in grouped.items():
        if order != list(range(1, len(order) + 1)):
            errors.append(f"Non-contiguous ordre_ligne for {group!r}: {order!r}")
    return errors


def _validate_summary(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    keys: set[tuple[Any, ...]] = set()
    market_contexts: defaultdict[tuple[Any, ...], set[tuple[Any, ...]]] = defaultdict(set)
    monetary = {
        "PRODUCTION_NETTE", "ENCOURS_SAIN", "CREANCES_EN_SOUFFRANCE", "ENCOURS_BRUT"
    }
    money_fields = (
        "valeur_reference_mdh", "valeur_comparaison_mdh", "evolution_pct",
        "marche_reference_mdh", "marche_comparaison_mdh", "evolution_marche_pct",
    )
    ratio_fields = (
        "ratio_reference_pct", "ratio_comparaison_pct", "evolution_ratio_points",
        "marche_ratio_reference_pct", "marche_ratio_comparaison_pct",
        "evolution_marche_ratio_points",
    )
    for index, row in enumerate(rows, start=1):
        organism = canonical_text(row["organisme"])
        if organism == "MARCHE" or "PART_DE_MARCHE" in organism:
            errors.append(f"Row {index}: market technical line exported")
        key = (
            row["date_reporting"], row["section"], organism, row["categorie"],
            row["periode_reference"], row["periode_comparaison"],
        )
        if key in keys:
            errors.append(f"Row {index}: duplicate business key")
        keys.add(key)
        group = (
            row["date_reporting"], row["section"], row["categorie"],
            row["periode_reference"], row["periode_comparaison"],
        )
        market_contexts[group].add(
            tuple(
                row[field]
                for field in (
                    "marche_reference_mdh", "marche_comparaison_mdh",
                    "evolution_marche_pct", "marche_ratio_reference_pct",
                    "marche_ratio_comparaison_pct", "evolution_marche_ratio_points",
                )
            )
        )
        if row["section"] in monetary:
            if any(row[field] not in (None, "") for field in ratio_fields):
                errors.append(f"Row {index}: monetary section contains ratio values")
        elif row["section"] == "CES_SUR_ENCOURS_BRUT":
            if any(row[field] not in (None, "") for field in money_fields):
                errors.append(f"Row {index}: ratio section contains monetary values")
        else:
            errors.append(f"Row {index}: unexpected section {row['section']!r}")
    for group, contexts in market_contexts.items():
        if len(contexts) > 1:
            errors.append(f"Inconsistent market propagation for {group!r}")
    return errors


def _validate_quarterly(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    keys: set[tuple[Any, ...]] = set()
    market_contexts: defaultdict[tuple[Any, ...], set[tuple[Any, ...]]] = defaultdict(set)
    monetary = {
        "PRODUCTION_NETTE", "ENCOURS_SAIN", "CREANCES_EN_SOUFFRANCE", "ENCOURS_BRUT"
    }
    money_fields = (
        "t1_mdh", "t2_mdh", "t3_mdh", "t4_mdh",
        "marche_t1_mdh", "marche_t2_mdh", "marche_t3_mdh", "marche_t4_mdh",
    )
    ratio_fields = (
        "t1_pct", "t2_pct", "t3_pct", "t4_pct",
        "marche_t1_pct", "marche_t2_pct", "marche_t3_pct", "marche_t4_pct",
    )
    market_fields = (
        "marche_t1_mdh", "marche_t2_mdh", "marche_t3_mdh", "marche_t4_mdh",
        "marche_t1_pct", "marche_t2_pct", "marche_t3_pct", "marche_t4_pct",
    )
    for index, row in enumerate(rows, start=1):
        organism = canonical_text(row["organisme"])
        if organism == "MARCHE" or "PART_DE_MARCHE" in organism:
            errors.append(f"Row {index}: market technical line exported")
        key = (
            row["date_reporting"], row["section"], organism, row["categorie"], row["annee"]
        )
        if key in keys:
            errors.append(f"Row {index}: duplicate business key")
        keys.add(key)
        group = (
            row["date_reporting"], row["section"], row["categorie"], row["annee"]
        )
        market_contexts[group].add(tuple(row[field] for field in market_fields))
        if row["section"] in monetary:
            if any(row[field] not in (None, "") for field in ratio_fields):
                errors.append(f"Row {index}: monetary section contains ratio values")
        elif row["section"] == "CES_SUR_ENCOURS_BRUT":
            if any(row[field] not in (None, "") for field in money_fields):
                errors.append(f"Row {index}: ratio section contains monetary values")
        else:
            errors.append(f"Row {index}: unexpected section {row['section']!r}")
    for group, contexts in market_contexts.items():
        if len(contexts) > 1:
            errors.append(f"Inconsistent market propagation for {group!r}")
    return errors
