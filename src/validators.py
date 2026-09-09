from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
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
STAGING_HEADER = (
    "date_reporting", "sheet", "section", "category_label_raw",
    "period_label_raw", "row_label_raw", "row_type", "value",
    "number_format", "source_row", "source_column",
)

QUARTERLY_HEADER = (
    "date_reporting", "section", "organisme", "categorie", "annee",
    "t1_mdh", "t2_mdh", "t3_mdh", "t4_mdh",
    "marche_t1_mdh", "marche_t2_mdh", "marche_t3_mdh", "marche_t4_mdh",
    "t1_pct", "t2_pct", "t3_pct", "t4_pct",
    "marche_t1_pct", "marche_t2_pct", "marche_t3_pct", "marche_t4_pct",
)

DETAILS_HEADER = ("marque", "nombre", "ville", "mois", "annee", "source")
PENETRATION_HEADER = (
    "marque", "ville", "periodicite", "dimension",
    "axe", "sous_axe", "valeur",
)

CONCESSION_HEADER = ("groupe_concessionnaire", "annee", "marque", "aivam", "sofac", "total", "row_type")
PIVOT_STAGING_HEADER = (
    "sheet", "bloc_index", "axe", "periodicite", "annee", "mois",
    "dimension", "marque", "ville", "organisme", "valeur",
)
VARIATION_TPR_HEADER = ("ville", "axe", "valeur")

FAMILY_HEADERS = {
    "production_marque_concession": CONCESSION_HEADER,
    "production_marque_pivot_staging": PIVOT_STAGING_HEADER,
    "production_marque_variation_tpr": VARIATION_TPR_HEADER,
    "apsf_competitor_quarterly": QUARTERLY_HEADER,
    "production_marque_details": DETAILS_HEADER,
    "production_marque_penetration": PENETRATION_HEADER,
    "apsf_activity": ACTIVITY_HEADER,
    "apsf_competitor_summary": SUMMARY_HEADER,
}


def validate_rows(family: str, header: Iterable[str], rows: list[dict[str, Any]], *, layout: str = "summary") -> list[str]:
    errors: list[str] = []
    expected = FAMILY_HEADERS.get(family)
    if expected is None:
        return [f"Unknown family validator: {family}"]
    if family == "apsf_competitor_summary":
        if layout not in {"summary", "staging"}:
            return [f"Unknown competitor layout: {layout}"]
        if layout == "staging":
            expected = STAGING_HEADER
    header_tuple = tuple(header)
    if family == "production_marque_details":
        if layout not in {"summary", "details", "flat"}:
            return [f"Unknown flat table layout: {layout}"]
        if len(set(header_tuple)) != len(header_tuple) or any(not name for name in header_tuple):
            errors.append("Flat table header requires non-empty unique columns")
        if layout in {"summary", "details"} and not set(DETAILS_HEADER) <= set(header_tuple):
            errors.append("Details header requires all required fields")
        expected = header_tuple
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
            if family.startswith("production_marque_"):
                # Decimal representations are formatted by Exporter, while
                # source text is allowed to resemble scientific notation.
                continue
            if is_scientific_notation(value):
                errors.append(f"Row {index}, {field}: scientific notation")
    if family == "apsf_activity":
        errors.extend(_validate_activity(rows))
    elif family == "apsf_competitor_summary":
        errors.extend(_validate_staging(rows) if layout == "staging" else _validate_summary(rows))
    elif family == "apsf_competitor_quarterly":
        errors.extend(_validate_quarterly(rows))
    elif family == "production_marque_details" and layout in {"summary", "details"}:
        errors.extend(_validate_details(rows))
    elif family == "production_marque_penetration":
        errors.extend(_validate_penetration(rows))
    elif family == "production_marque_concession":
        errors.extend(_validate_concession(rows))
    elif family == "production_marque_pivot_staging":
        errors.extend(_validate_pivot_staging(rows))
    elif family == "production_marque_variation_tpr":
        errors.extend(_validate_variation_tpr(rows))
    return errors


def _validate_details(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    for index, row in enumerate(rows, start=1):
        for field in ("nombre", "mois", "annee"):
            value = row.get(field)
            if value in (None, "NS") or (isinstance(value, str) and value in EXCEL_ERRORS):
                continue
            expected = Decimal if field == "nombre" else int
            if not isinstance(value, expected) or isinstance(value, bool):
                errors.append(f"Row {index}: {field} must be {expected.__name__}")
            elif isinstance(value, Decimal) and not value.is_finite():
                errors.append(f"Row {index}: {field} must be finite")
    return errors


def _validate_penetration(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    for index, row in enumerate(rows, start=1):
        dimension = row.get("dimension")
        if dimension not in {"MARQUE", "VILLE"}:
            errors.append(f"Row {index}: invalid dimension")
        label = row.get("ville" if dimension == "VILLE" else "marque")
        if label in (None, "") or canonical_text(label) == "TOTAL_GENERAL":
            errors.append(f"Row {index}: missing label or exported total")
        if dimension == "VILLE" and row.get("marque") is not None:
            errors.append(f"Row {index}: VILLE dimension must not contain a marque")
        if row.get("periodicite") not in {"MENSUEL", "ANNUEL"}:
            errors.append(f"Row {index}: invalid periodicite")
        if row.get("axe") in (None, "") or row.get("sous_axe") in (None, ""):
            errors.append(f"Row {index}: missing penetration axis or sub-axis")
        value = row.get("valeur")
        if value not in (None, "NS") and not (isinstance(value, str) and value in EXCEL_ERRORS):
            if not isinstance(value, Decimal) or not value.is_finite():
                errors.append(f"Row {index}: valeur must be a finite Decimal")
    return errors


def _validate_concession(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    for index, row in enumerate(rows, start=1):
        if not row.get("groupe_concessionnaire") or not isinstance(row.get("annee"), int):
            errors.append(f"Row {index}: missing concession group or integer year")
        expected = "total" if canonical_text(row.get("marque")) == "TOTAL_GENERAL" else "marque"
        if not row.get("marque") or row.get("row_type") != expected:
            errors.append(f"Row {index}: invalid concession row_type or marque")
        errors.extend(_validate_pivot_measures(index, row, ("aivam", "sofac", "total")))
    return errors


def _validate_pivot_staging(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    for index, row in enumerate(rows, start=1):
        if not row.get("sheet") or not isinstance(row.get("bloc_index"), int) or row["bloc_index"] < 1:
            errors.append(f"Row {index}: invalid sheet or bloc_index")
        dimension = row.get("dimension")
        if dimension not in {"MARQUE", "VILLE"}:
            errors.append(f"Row {index}: invalid synthesis dimension")
        if not row.get("organisme"):
            errors.append(f"Row {index}: missing synthesis organisme")
        if row.get("periodicite") not in {"MENSUEL", "ANNUEL"} or not row.get("axe"):
            errors.append(f"Row {index}: invalid synthesis axis or periodicite")
        if dimension == "VILLE" and row.get("marque") is not None:
            errors.append(f"Row {index}: VILLE dimension must not contain a marque")
        errors.extend(_validate_pivot_measures(index, row, ("valeur",)))
    return errors


def _validate_variation_tpr(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    for index, row in enumerate(rows, start=1):
        if not row.get("ville"):
            errors.append(f"Row {index}: missing Variation TPR city")
        if not row.get("axe"):
            errors.append(f"Row {index}: missing Variation TPR axis")
        errors.extend(_validate_pivot_measures(index, row, ("valeur",)))
    return errors


def _validate_pivot_measures(index: int, row: dict[str, Any], fields: tuple[str, ...]) -> list[str]:
    errors: list[str] = []
    for field in fields:
        value = row.get(field)
        if value in (None, "NS") or (isinstance(value, str) and value in EXCEL_ERRORS):
            continue
        if not isinstance(value, Decimal) or not value.is_finite():
            errors.append(f"Row {index}: {field} must be a finite Decimal")
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


def _validate_staging(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    keys: set[tuple[Any, ...]] = set()
    for index, row in enumerate(rows, start=1):
        key = (
            row["sheet"], row["section"], row["category_label_raw"],
            row["period_label_raw"], row["row_label_raw"], row["source_row"],
            row["source_column"],
        )
        if key in keys:
            errors.append(f"Row {index}: duplicate staging key")
        keys.add(key)
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
