from copy import deepcopy
from pathlib import Path
import unittest

from src.config_loader import ConfigError, load_all, select_workbook_config
from src.config_validator import validate_configuration
from src.parsers import FAMILY_REGISTRY


class ConfigurationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.configs = load_all(Path(__file__).parents[1] / "config")

    def test_project_configuration_is_valid(self):
        validate_configuration(self.configs, FAMILY_REGISTRY)

    def test_family_sheets_are_mandatory(self):
        configs = deepcopy(self.configs)
        del configs["families"]["families"]["apsf_activity"]["family_sheets"]
        with self.assertRaises(ConfigError):
            validate_configuration(configs, FAMILY_REGISTRY)

    def test_workbook_defaults_and_overrides_are_merged(self):
        selected = select_workbook_config(
            self.configs["workbooks"],
            "SCC_Statistiques APSF Globales Mars 2026.xlsx",
        )
        self.assertIn("GLOBAL", selected["sheets"])
        self.assertIn("Syn prin concurents", selected["sheets"])

    def test_sheet_cannot_belong_to_two_families(self):
        configs = deepcopy(self.configs)
        configs["families"]["families"]["apsf_competitor_summary"]["family_sheets"].append(
            " rci "
        )
        with self.assertRaises(ConfigError):
            validate_configuration(configs, FAMILY_REGISTRY)


if __name__ == "__main__":
    unittest.main()
