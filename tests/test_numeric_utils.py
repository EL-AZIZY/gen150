from decimal import Decimal
import unittest

from src.numeric_utils import (
    format_decimal,
    has_thousands_scaling,
    percentage_decimal,
    scaled_decimal,
    to_decimal,
)


class NumericUtilsTests(unittest.TestCase):
    def test_decimal_rendering_has_no_scientific_notation(self):
        cases = {
            Decimal("781065.8000"): "781065.8",
            Decimal("1034282.0"): "1034282",
            Decimal("20.8900"): "20.89",
            Decimal("0"): "0",
            None: "",
            "NS": "NS",
            "#REF!": "#REF!",
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(format_decimal(value), expected)

    def test_empty_zero_ns_and_errors_are_distinct(self):
        self.assertIsNone(to_decimal(""))
        self.assertEqual(to_decimal(0), Decimal("0"))
        self.assertEqual(to_decimal("NS"), "NS")
        self.assertEqual(to_decimal("#DIV/0!"), "#DIV/0!")

    def test_excel_scaling_and_percent_formats(self):
        self.assertTrue(has_thousands_scaling("#,##0,"))
        self.assertFalse(has_thousands_scaling("#,##0"))
        self.assertEqual(scaled_decimal(1_234_000, "#,##0,"), Decimal("1234"))
        self.assertEqual(percentage_decimal(0.2089, "0.00%"), Decimal("20.8900"))


if __name__ == "__main__":
    unittest.main()

