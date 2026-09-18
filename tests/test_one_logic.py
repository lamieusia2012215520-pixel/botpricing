import unittest
from one_logic import (
    one_is_discount_charge,
    one_should_include_charge,
    one_api_add_charge,
    one_api_charge_equipment_key,
    excel_formula_from_parts,
)


class OneLogicTests(unittest.TestCase):
    def test_discount_detection(self):
        self.assertTrue(one_is_discount_charge("Special Promotion Service"))
        self.assertTrue(one_is_discount_charge("special promotion service [discount]"))
        self.assertTrue(one_is_discount_charge("Promotion Discount"))
        self.assertTrue(one_is_discount_charge("", {"chargeCode": "DISC"}))
        self.assertTrue(one_is_discount_charge("", {"chargeType": "discount"}))
        self.assertTrue(one_is_discount_charge("", {"isDiscount": True}))
        self.assertFalse(one_is_discount_charge("Basic Ocean Freight"))
        self.assertFalse(one_is_discount_charge("Carrier Security Surcharge"))

    def test_should_include_charge_buenaventura_case(self):
        # Basic Ocean Freight
        self.assertTrue(one_should_include_charge("Basic Ocean Freight", group="basicOceanFreightCharges"))
        # Special Promotion Service (Discount)
        self.assertTrue(one_should_include_charge("Special Promotion Service", group="basicOceanFreightCharges"))
        # Carrier Security Surcharge (included in originCharges)
        self.assertTrue(one_should_include_charge("Carrier Security Surcharge", group="originCharges"))
        # Local origin excluded
        self.assertFalse(one_should_include_charge("Doc Fee (Origin)", group="originCharges"))
        self.assertFalse(one_should_include_charge("Seal Fee", group="originCharges"))
        self.assertFalse(one_should_include_charge("Terminal Handling Charge (L)", group="originCharges"))
        # Destination excluded
        self.assertFalse(one_should_include_charge("Doc Fee (Dest)", group="destinationCharges"))
        self.assertFalse(one_should_include_charge("Container Release Order", group="destinationCharges"))
        self.assertFalse(one_should_include_charge("Terminal Handling Charge (D)", group="destinationCharges"))

    def test_buenaventura_hcm_formula_and_price(self):
        """
        Verify user's exact case for HCM - BUENAVENTURA:
        Basic Ocean Freight: 6200
        Special Promotion Service: -200
        Carrier Security Surcharge: +15
        Total = 6200 - 200 + 15 = 6015
        Formula = "=6200-200+15"
        """
        final_prices = {"DRY 20": 0.0, "DRY 40": 0.0, "DRY 40H": 0.0}
        formula_parts = {"DRY 20": [], "DRY 40": [], "DRY 40H": []}

        # 1. Basic Ocean Freight: 6200
        charge_bof = {
            "chargeName": "Basic Ocean Freight",
            "equipmentIsoCode": "45G1",
        }
        one_api_add_charge(final_prices, formula_parts, charge_bof, raw_amount_usd=6200.0)

        # 2. Special Promotion Service: 200 (discount -> should be -200)
        charge_promo = {
            "chargeName": "Special Promotion Service",
            "equipmentIsoCode": "45G1",
        }
        one_api_add_charge(final_prices, formula_parts, charge_promo, raw_amount_usd=200.0)

        # 3. Carrier Security Surcharge: 15
        charge_css = {
            "chargeName": "Carrier Security Surcharge",
            "equipmentIsoCode": "45G1",
        }
        one_api_add_charge(final_prices, formula_parts, charge_css, raw_amount_usd=15.0)

        self.assertEqual(final_prices["DRY 40H"], 6015.0)
        formula = excel_formula_from_parts(formula_parts["DRY 40H"])
        self.assertEqual(formula, "=6200-200+15")

    def test_etd_dates_formatting(self):
        from datetime import date, datetime
        from bot_cli import format_etd_dates_excel

        # User's exact case: 26, 30-Sep & 2-Oct
        dates = [date(2026, 9, 26), date(2026, 9, 30), date(2026, 10, 2)]
        self.assertEqual(format_etd_dates_excel(dates), "26, 30-Sep & 2-Oct")

        # Same month: 26, 28, 30-Sep
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26), date(2026, 9, 28), date(2026, 9, 30)]), "26, 28, 30-Sep")

        # 1 in first month, 2 in second month: 26-Sep & 2, 6-Oct
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26), date(2026, 10, 2), date(2026, 10, 6)]), "26-Sep & 2, 6-Oct")

        # 3 different months: 26-Sep, 2-Oct & 5-Nov
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26), date(2026, 10, 2), date(2026, 11, 5)]), "26-Sep, 2-Oct & 5-Nov")

        # 2 dates: 26-Sep & 30-Sep
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26), date(2026, 9, 30)]), "26-Sep & 30-Sep")

        # 1 date: 26-Sep
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26)]), "26-Sep")


if __name__ == "__main__":
    unittest.main()
