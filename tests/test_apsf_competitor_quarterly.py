from decimal import Decimal
import logging
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from src.exporter import Exporter
from src.parsers.apsf_competitor_quarterly_parser import (
    ApsfCompetitorQuarterlyParser,
)
from src.parsers.base_parser import ParserContext
from src.runner import process_workbook
from src.settings import Settings
from src.validators import QUARTERLY_HEADER, validate_rows


def quarterly_config(section: str = "PRODUCTION_NETTE") -> dict:
    return {
        "date_reporting": "2026-12-31",
        "annee": 2026,
        "label_column": "B",
        "market_label": "Marché",
        "market_share_label": "Part de marché",
        "require_market_context": True,
        "category_blocks": {"VEHICULES": ["E", "F", "G", "H"]},
        "sections": {
            section: {
                "section_row": 9,
                "category_row": 11,
                "period_row": 13,
                "data_start_row": 16,
                "data_end_row": 18,
            }
        },
    }


def make_context(
    section: str = "PRODUCTION_NETTE",
    include_market: bool = True,
    quarter_headers: tuple[str, str, str, str] = ("T1", "T2", "T3", "T4"),
) -> ParserContext:
    values = Workbook()
    formulas = Workbook()
    for workbook in (values, formulas):
        sheet = workbook.active
        sheet.title = "Syn prin concurents par trim"
        for column, header in zip("EFGH", quarter_headers):
            sheet[f"{column}13"] = header
        if include_market:
            sheet["B16"] = "Marché"
        sheet["B17"] = "SOFAC"
        sheet["B18"] = "Part de marché"
        if section == "CES_SUR_ENCOURS_BRUT":
            market_values = (0.20, 0.21, 0.22, 0.23)
            organization_values = (0.10, 0.11, 0.12, 0.13)
            for column, market, organization in zip(
                "EFGH", market_values, organization_values
            ):
                sheet[f"{column}16"] = market
                sheet[f"{column}17"] = organization
                sheet[f"{column}18"] = 0.5
                for row in (16, 17, 18):
                    sheet[f"{column}{row}"].number_format = "0.0%"
        else:
            market_values = (2_000_000, 2_100_000, 2_200_000, 2_300_000)
            organization_values = (1_000_000, 1_100_000, 1_200_000, 1_300_000)
            for column, market, organization in zip(
                "EFGH", market_values, organization_values
            ):
                sheet[f"{column}16"] = market
                sheet[f"{column}17"] = organization
                sheet[f"{column}18"] = 0.5
                sheet[f"{column}16"].number_format = "#,##0,"
                sheet[f"{column}17"].number_format = "#,##0,"
                sheet[f"{column}18"].number_format = "0.0%"
    return ParserContext(
        worksheet=values.active,
        formula_worksheet=formulas.active,
        sheet_config=quarterly_config(section),
        family_config={},
        error_policy="warning",
        missing_formula_cache_policy="warning",
        logger=logging.getLogger("test"),
    )


class CompetitorQuarterlyTests(unittest.TestCase):
    def test_nominal_sheet_and_market_propagation(self):
        result = ApsfCompetitorQuarterlyParser(make_context()).process()
        self.assertEqual(result.header, QUARTERLY_HEADER)
        self.assertEqual(len(result.rows), 1)
        row = result.rows[0]
        self.assertEqual(row["date_reporting"], "2026-12-31")
        self.assertEqual(row["annee"], 2026)
        self.assertEqual(row["organisme"], "SOFAC")
        self.assertEqual(row["t1_mdh"], Decimal("1000"))
        self.assertEqual(row["t4_mdh"], Decimal("1300"))
        self.assertEqual(row["marche_t1_mdh"], Decimal("2000"))
        self.assertEqual(row["marche_t4_mdh"], Decimal("2300"))
        self.assertEqual(
            validate_rows("apsf_competitor_quarterly", result.header, result.rows), []
        )

    def test_market_share_row_is_ignored(self):
        result = ApsfCompetitorQuarterlyParser(make_context()).process()
        self.assertEqual([row["organisme"] for row in result.rows], ["SOFAC"])

    def test_missing_market_context_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Market row missing"):
            ApsfCompetitorQuarterlyParser(
                make_context(include_market=False)
            ).process()

    def test_ratio_section_populates_only_percentage_fields(self):
        result = ApsfCompetitorQuarterlyParser(
            make_context("CES_SUR_ENCOURS_BRUT")
        ).process()
        row = result.rows[0]
        self.assertIsNone(row["t1_mdh"])
        self.assertIsNone(row["marche_t1_mdh"])
        self.assertEqual(row["t1_pct"], Decimal("10.0"))
        self.assertEqual(row["t4_pct"], Decimal("13.00"))
        self.assertEqual(row["marche_t1_pct"], Decimal("20.0"))
        self.assertEqual(
            validate_rows("apsf_competitor_quarterly", result.header, result.rows), []
        )

    def test_invalid_quarter_headers_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "Invalid quarterly headers"):
            ApsfCompetitorQuarterlyParser(
                make_context(quarter_headers=("T1", "T2", "T4", "T3"))
            ).process()

    def test_end_to_end_excel_csv_manifest(self):
        parser_context = make_context()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook_path = root / "APSF quarterly.xlsx"
            parser_context.worksheet.parent.save(workbook_path)
            output = root / "output"
            settings = Settings(root, output)
            exporter = Exporter(output, ";", "utf-8-sig")
            families = {
                "apsf_competitor_quarterly": {
                    "parser": "ApsfCompetitorQuarterlyParser",
                    "enabled": True,
                    "family_sheets": ["Syn prin concurents par trim"],
                }
            }
            workbook_configs = {
                "defaults": {
                    "sheets": {
                        "Syn prin concurents par trim": {
                            "family": "apsf_competitor_quarterly",
                            **quarterly_config(),
                        }
                    }
                },
                "workbooks": [],
            }
            manifest = process_workbook(
                workbook_path, families, workbook_configs, settings, exporter
            )
            self.assertEqual(manifest["success_count"], 1)
            self.assertEqual(manifest["rejected_count"], 0)
            entry = manifest["sheets"][0]
            self.assertEqual(entry["parser"], "ApsfCompetitorQuarterlyParser")
            csv_path = Path(entry["output"])
            self.assertEqual(
                csv_path.read_text(encoding="utf-8-sig").splitlines()[0],
                ";".join(QUARTERLY_HEADER),
            )
            self.assertTrue(
                (output / "manifests" / "apsf_quarterly.json").exists()
            )


if __name__ == "__main__":
    unittest.main()

