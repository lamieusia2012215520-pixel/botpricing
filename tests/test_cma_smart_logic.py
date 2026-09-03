import unittest
from datetime import datetime

from cma_logic import (
    classify_cma_card_texts,
    cma_date_input_matches,
    dedupe_cma_card_summaries,
    parse_cma_comparable_price,
)


class CMASmartLogicTests(unittest.TestCase):
    def test_excel_formula_price_can_be_compared_between_port_and_ramp(self):
        self.assertEqual(2011.0, parse_cma_comparable_price("=1911+150-50"))
        self.assertEqual(1911.0, parse_cma_comparable_price("1,911"))
        self.assertEqual(float("inf"), parse_cma_comparable_price("No Offer"))

    def test_existing_departure_date_is_recognised_across_cma_formats(self):
        target = datetime(2026, 8, 31)

        self.assertTrue(cma_date_input_matches("31-Aug-2026", target))
        self.assertTrue(cma_date_input_matches("2026/08/31", target))
        self.assertFalse(cma_date_input_matches("30-Aug-2026", target))

    def test_one_loaded_card_makes_snapshot_partial_instead_of_waiting_for_all(self):
        state = classify_cma_card_texts([
            "12 Aug per 20ST 1,911 USD",
            "Loading price...",
        ])

        self.assertEqual("PARTIAL", state)

    def test_all_resolved_cards_make_snapshot_ready(self):
        state = classify_cma_card_texts([
            "12 Aug per 20ST 1,911 USD",
            "No offer for some equipments MODIFY TEU",
            "SOLD OUT",
        ])

        self.assertEqual("READY", state)

    def test_all_loading_cards_stay_pending(self):
        self.assertEqual(
            "LOADING",
            classify_cma_card_texts(["Loading...", "Please wait"]),
        )

    def test_duplicate_cards_are_removed_before_opening_details(self):
        etd = datetime(2026, 8, 31)
        cards = [
            {"date": etd, "price": 3225, "transit": 18, "ts_port": "SINGAPORE", "card_idx": 0},
            {"date": etd, "price": 3225, "transit": 18, "ts_port": "SINGAPORE", "card_idx": 1},
            {"date": etd, "price": 3225, "transit": 16, "ts_port": "SINGAPORE", "card_idx": 2},
        ]

        unique = dedupe_cma_card_summaries(cards)

        self.assertEqual([0, 2], [item["card_idx"] for item in unique])


if __name__ == "__main__":
    unittest.main()
