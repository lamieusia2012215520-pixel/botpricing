import unittest
from datetime import datetime

from cosco_elines_ui import (
    dropdown_has_selected_option,
    is_new_elines_result_card_text,
    is_elines_auth_page,
    is_elines_booking_page,
    is_no_matching_ocean_freight_message,
    parse_new_elines_card_schedule,
    parse_premium_service_row,
    select_preferred_premium_service,
)


class CoscoELinesNewUiTests(unittest.TestCase):
    def test_recognises_new_result_card_even_without_view_premium(self):
        card_text = (
            "Sep06 (Sun) 14 Days Sep20 (Sun) "
            "CY Cutoff: 2026-09-05 (Sat) 03:00 "
            "Service: VTS Vessel / Voyage: INCEDA 052S COSCO ADEN 143S "
            "TBC Book Now"
        )

        self.assertTrue(is_new_elines_result_card_text(card_text))
        self.assertFalse(is_new_elines_result_card_text("Legacy Flash Sale USD 1,925"))

    def test_recognises_equipment_already_selected_from_multi_select_tags(self):
        selected = [
            "20GP - 20' General Purpose Container",
            "40GP - 40' General Purpose Container",
            "40HQ - 40' Hi-Cube Container",
        ]

        self.assertTrue(dropdown_has_selected_option(selected, "20GP", prefix=True))
        self.assertTrue(dropdown_has_selected_option(selected, "40HQ", prefix=True))
        self.assertFalse(dropdown_has_selected_option(selected, "45HQ", prefix=True))
        self.assertTrue(dropdown_has_selected_option(["General"], "General"))

    def test_accepts_only_the_new_booking_request_page_as_a_ready_elines_page(self):
        self.assertTrue(
            is_elines_booking_page(
                "https://elines.coscoshipping.com/ebusiness/bookingrequest/?from=dashboard"
            )
        )
        self.assertFalse(
            is_elines_booking_page(
                "https://elines.coscoshipping.com/ebusiness/aczoneSpotBooking/"
            )
        )

    def test_recognises_elines_login_notice_and_sso_as_auth_pages(self):
        self.assertTrue(
            is_elines_auth_page(
                "https://elines.coscoshipping.com/ebusiness/notice/loginPlease/"
            )
        )
        self.assertTrue(
            is_elines_auth_page("https://exiamfw.lines.coscoshipping.com/auth/login")
        )
        self.assertFalse(
            is_elines_auth_page(
                "https://elines.coscoshipping.com/ebusiness/bookingrequest/"
            )
        )

    def test_recognises_the_new_immediate_no_price_message(self):
        message = (
            "Sorry, no matching ocean freight products were found based on your "
            "search criteria. You can try switching the search method above "
            "to continue completing your booking!"
        )

        self.assertTrue(is_no_matching_ocean_freight_message(message))
        self.assertFalse(is_no_matching_ocean_freight_message("Search Results"))

    def test_parses_flash_sale_prices_and_remaining_stock_from_premium_modal_row(self):
        row = parse_premium_service_row(
            "Flash Sale Remaining Stock: 20 TEU "
            "20GP From USD1,925 40GP From USD3,850 40HQ From USD3,850",
            "Flash Sale",
        )

        self.assertEqual("Flash Sale", row["service"])
        self.assertEqual(20, row["remaining_stock"])
        self.assertEqual(
            {"20GP": 1925.0, "40GP": 3850.0, "40HQ": 3850.0},
            row["rates"],
        )

    def test_prefers_flash_sale_row_when_modal_has_both_services(self):
        standard = parse_premium_service_row(
            "Standard Service TBC 20GP From USD2,025 40HQ From USD4,050 "
            "40GP From USD4,050",
            "Standard Service",
        )
        flash = parse_premium_service_row(
            "Flash Sale Remaining Stock: 20 TEU 20GP From USD1,925 "
            "40GP From USD3,850 40HQ From USD3,850",
            "Flash Sale",
        )

        self.assertEqual("Flash Sale", select_preferred_premium_service([standard, flash])["service"])

    def test_parses_new_card_schedule_for_existing_golden_rule_engine(self):
        schedule = parse_new_elines_card_schedule(
            "Aug13 (Thu) 24 Days Sep06 (Sun) CY Cutoff: 2026-08-12 (Wed) 03:00 "
            "Service: VTS Vessel / Voyage: INCEDA 052S COSCO ADEN 143S TBC",
            now=datetime(2026, 8, 10),
        )

        self.assertEqual(datetime(2026, 8, 13), schedule["etd_dt"])
        self.assertEqual(24, schedule["tt_days"])
        self.assertEqual("TBC", schedule["space"])


if __name__ == "__main__":
    unittest.main()
