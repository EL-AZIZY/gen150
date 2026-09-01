import logging
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from src.parsers.apsf_activity_parser import ApsfActivityParser
from src.parsers.base_parser import ParserContext, SheetRejected
from src.exporter import Exporter
from src.runner import process_workbook
from src.settings import Settings
from src.validators import ACTIVITY_HEADER, validate_rows


def activity_config():
    return {
        "limits": {"first_row": 1, "last_row": 117, "first_column": 1, "last_column": 9},
        "title_area": {"first_row": 1, "last_row": 2, "first_column": 1, "last_column": 9},
        "label_column": "A",
        "data_columns": {
            "montant_reference_mad": "B", "montant_comparaison_mad": "C",
            "dossiers_reference": "D", "dossiers_comparaison": "E",
            "variation_montant_mad": "F", "variation_montant_pct": "G",
            "variation_dossiers_unites": "H", "variation_dossiers_pct": "I",
        },
        "default_section": "ACTIVITE",
        "infer_sections": False,
        "skip_labels": [],
        "unit_tokens": ["unité"],
        "sections": {},
        "hierarchy": {
            "types": {"PRETS_AFFECTES": ["PRÊTS AFFECTÉS"]},
            "categories": {"VEHICULES": ["VÉHICULES"]},
            "products": {"CREDIT_CLASSIQUE": ["CRÉDIT CLASSIQUE"]},
        },
    }


def context(values, formulas, policy="warning"):
    return ParserContext(
        worksheet=values.active,
        formula_worksheet=formulas.active,
        sheet_config=activity_config(),
        family_config={},
        error_policy=policy,
        missing_formula_cache_policy="warning",
        logger=logging.getLogger("test"),
    )


class ActivityParserTests(unittest.TestCase):
    def make_books(self):
        values = Workbook()
        formulas = Workbook()
        for workbook in (values, formulas):
            sheet = workbook.active
            sheet.title = "SOFAC"
            sheet["A1"] = "STATISTIQUES D'ACTIVITE A FIN MARS 2026"
            sheet["A2"] = "Comparaison 2025"
            sheet["A3"] = "PRÊTS AFFECTÉS"
            sheet["A4"] = "VÉHICULES"
            sheet["A5"] = "CRÉDIT CLASSIQUE"
            for column, value in zip("BCDEFGHI", [100, 90, 10, 9, 10, 0.1, 1, 0]):
                sheet[f"{column}5"] = value
            sheet["G5"].number_format = "0.0%"
            sheet["I5"].number_format = "0.0%"
            sheet["A6"] = "TOTAL"
            sheet["B6"] = 999
            sheet["A7"] = "SOUS-TOTAL VÉHICULES"
            sheet["B7"] = 999
            sheet["A118"] = "OUTSIDE"
            sheet["B118"] = 999
            sheet["J5"] = 999
        return values, formulas

    def test_hierarchy_limits_totals_and_header(self):
        values, formulas = self.make_books()
        result = ApsfActivityParser(context(values, formulas)).process()
        self.assertEqual(result.header, ACTIVITY_HEADER)
        self.assertEqual(len(result.rows), 1)
        row = result.rows[0]
        self.assertEqual(row["date_reporting"], "2026-03-31")
        self.assertEqual(row["type_pret"], "PRETS_AFFECTES")
        self.assertEqual(row["categorie"], "VEHICULES")
        self.assertEqual(row["produit"], "CREDIT_CLASSIQUE")
        self.assertEqual(str(row["variation_montant_pct"]), "10.0")
        self.assertEqual(str(row["variation_dossiers_pct"]), "0")
        self.assertEqual(validate_rows("apsf_activity", result.header, result.rows), [])

    def test_year_cells_override_unrelated_years_in_title_area(self):
        values, formulas = self.make_books()
        for workbook in (values, formulas):
            workbook.active["B7"] = 2025
            workbook.active["C7"] = 2024
        config = activity_config()
        config["year_cells"] = {"reference": "B7", "comparison": "C7"}
        parser_context = context(values, formulas)
        parser_context.sheet_config = config
        result = ApsfActivityParser(parser_context).process()
        self.assertEqual(result.rows[0]["annee_reference"], 2025)
        self.assertEqual(result.rows[0]["annee_comparaison"], 2024)

    def test_excel_error_warning_is_not_changed_to_zero(self):
        values, formulas = self.make_books()
        values.active["B5"] = "#REF!"
        formulas.active["B5"] = "=#REF!"
        result = ApsfActivityParser(context(values, formulas)).process()
        self.assertEqual(result.rows[0]["montant_reference_mad"], "#REF!")
        self.assertTrue(result.warnings)

    def test_excel_error_reject_policy(self):
        values, formulas = self.make_books()
        values.active["B5"] = "#VALUE!"
        with self.assertRaises(SheetRejected):
            ApsfActivityParser(context(values, formulas, "reject")).process()

    def test_integration_routes_parses_exports_and_manifests(self):
        values, _ = self.make_books()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook_path = root / "APSF Mars 2026.xlsx"
            values.save(workbook_path)
            output = root / "output"
            settings = Settings(root, output)
            exporter = Exporter(output, ";", "utf-8-sig")
            families = {
                "apsf_activity": {
                    "parser": "ApsfActivityParser",
                    "enabled": True,
                    "family_sheets": ["SOFAC"],
                }
            }
            workbook_configs = {
                "defaults": {"sheets": {"SOFAC": {"family": "apsf_activity", **activity_config()}}},
                "workbooks": [],
            }
            manifest = process_workbook(
                workbook_path, families, workbook_configs, settings, exporter
            )
            self.assertEqual(manifest["success_count"], 1)
            self.assertEqual(manifest["rejected_count"], 0)
            csv_path = Path(manifest["sheets"][0]["output"])
            self.assertTrue(csv_path.exists())
            self.assertEqual(
                csv_path.read_text(encoding="utf-8-sig").splitlines()[0],
                ";".join(ACTIVITY_HEADER),
            )
            self.assertTrue((output / "manifests" / "apsf_mars_2026.json").exists())


if __name__ == "__main__":
    unittest.main()
