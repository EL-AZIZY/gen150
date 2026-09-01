from decimal import Decimal
from datetime import datetime
import logging
import unittest

from openpyxl import Workbook

from src.parsers.apsf_competitor_summary_parser import ApsfCompetitorSummaryParser
from src.parsers.base_parser import ParserContext
from src.validators import SUMMARY_HEADER, validate_rows


def make_context(section="PRODUCTION_NETTE"):
    values = Workbook()
    formulas = Workbook()
    for workbook in (values, formulas):
        sheet = workbook.active
        sheet.title = "Syn prin concurents"
        sheet["A1"] = "STATISTIQUES APSF A FIN MARS 2026"
        sheet["E13"] = "Mar-26"
        sheet["F13"] = "Mar-25"
        sheet["B16"] = "Marché"
        sheet["E16"] = 2_000_000
        sheet["F16"] = 1_800_000
        sheet["G16"] = 0.111
        sheet["B17"] = "SOFAC"
        sheet["E17"] = 1_000_000
        sheet["F17"] = 900_000
        sheet["G17"] = 0.1
        sheet["B18"] = "Part de marché"
        sheet["E18"] = 0.5
        for row in (16, 17):
            sheet[f"E{row}"].number_format = "#,##0,"
            sheet[f"F{row}"].number_format = "#,##0,"
            sheet[f"G{row}"].number_format = "0.0%"
        if section == "CES_SUR_ENCOURS_BRUT":
            sheet["E16"] = 0.2
            sheet["F16"] = 0.18
            sheet["G16"] = 0.02
            sheet["E17"] = 0.12
            sheet["F17"] = 0.10
            sheet["G17"] = 0.02
            for row in (16, 17):
                for column in "EFG":
                    sheet[f"{column}{row}"].number_format = "0.0%"
    config = {
        "title_area": {"first_row": 1, "last_row": 2, "first_column": 1, "last_column": 5},
        "label_column": "B",
        "market_label": "Marché",
        "market_share_label": "Part de marché",
        "require_market_context": True,
        "category_blocks": {"VEHICULES": ["E", "F", "G"]},
        "sections": {
            section: {
                "period_row": 13,
                "data_start_row": 16,
                "data_end_row": 18,
            }
        },
    }
    return ParserContext(
        worksheet=values.active,
        formula_worksheet=formulas.active,
        sheet_config=config,
        family_config={},
        error_policy="warning",
        missing_formula_cache_policy="warning",
        logger=logging.getLogger("test"),
    )


class CompetitorSummaryTests(unittest.TestCase):
    def test_excel_date_cell_produces_month_end_reporting_date(self):
        context = make_context()
        context.worksheet["A1"] = datetime(2026, 3, 1)
        context.formula_worksheet["A1"] = datetime(2026, 3, 1)
        result = ApsfCompetitorSummaryParser(context).process()
        self.assertEqual(result.rows[0]["date_reporting"], "2026-03-31")

    def test_market_propagation_periods_scaling_and_ignored_rows(self):
        result = ApsfCompetitorSummaryParser(make_context()).process()
        self.assertEqual(result.header, SUMMARY_HEADER)
        self.assertEqual(len(result.rows), 1)
        row = result.rows[0]
        self.assertEqual(row["organisme"], "SOFAC")
        self.assertEqual(row["periode_reference"], "2026-03")
        self.assertEqual(row["periode_comparaison"], "2025-03")
        self.assertEqual(row["valeur_reference_mdh"], Decimal("1000"))
        self.assertEqual(row["marche_reference_mdh"], Decimal("2000"))
        self.assertEqual(row["evolution_pct"], Decimal("10.0"))
        self.assertEqual(validate_rows("apsf_competitor_summary", result.header, result.rows), [])

    def test_ratio_section_populates_only_ratio_fields(self):
        result = ApsfCompetitorSummaryParser(make_context("CES_SUR_ENCOURS_BRUT")).process()
        row = result.rows[0]
        self.assertIsNone(row["valeur_reference_mdh"])
        self.assertEqual(row["ratio_reference_pct"], Decimal("12.00"))
        self.assertEqual(row["marche_ratio_reference_pct"], Decimal("20.0"))
        self.assertEqual(validate_rows("apsf_competitor_summary", result.header, result.rows), [])

    def test_periods_are_read_category_by_category(self):
        context = make_context()
        for sheet in (context.worksheet, context.formula_worksheet):
            sheet["I13"] = "Dec-25"
            sheet["J13"] = "Dec-22"
            sheet["I16"] = 500_000
            sheet["J16"] = 400_000
            sheet["K16"] = 0.25
            sheet["I17"] = 200_000
            sheet["J17"] = 100_000
            sheet["K17"] = 1
            for row in (16, 17):
                sheet[f"I{row}"].number_format = "#,##0,"
                sheet[f"J{row}"].number_format = "#,##0,"
                sheet[f"K{row}"].number_format = "0.0%"
        context.sheet_config["category_blocks"]["GLOBAL"] = ["I", "J", "K"]
        result = ApsfCompetitorSummaryParser(context).process()
        rows_by_category = {row["categorie"]: row for row in result.rows}
        self.assertEqual(rows_by_category["VEHICULES"]["periode_reference"], "2026-03")
        self.assertEqual(rows_by_category["GLOBAL"]["periode_reference"], "2025-12")
        self.assertEqual(rows_by_category["GLOBAL"]["periode_comparaison"], "2022-12")

    def test_duplicate_business_key_is_rejected(self):
        result = ApsfCompetitorSummaryParser(make_context()).process()
        errors = validate_rows(
            "apsf_competitor_summary", result.header, result.rows + [result.rows[0].copy()]
        )
        self.assertTrue(any("duplicate business key" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
