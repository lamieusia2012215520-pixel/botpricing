import unittest
from datetime import datetime


try:
    from emc_greenx_logic import (
        build_quote_detail_payload,
        hydrate_quote_with_detail,
        quote_departure_is_on_or_before,
    )
except ImportError:
    # Keep the RED phase as an assertion failure with a useful message instead
    # of an import-time test collection error while the helper is being added.
    build_quote_detail_payload = None
    hydrate_quote_with_detail = None
    quote_departure_is_on_or_before = None


class EmcGreenxDetailLogicTests(unittest.TestCase):
    def test_builds_exact_detail_payload_without_container_fields(self):
        self.assertIsNotNone(
            build_quote_detail_payload,
            "EMC must build the GreenX detail request before deciding a schedule has no price",
        )

        payload = build_quote_detail_payload(
            quote_result={"uuid": "quote-uuid", "rqstNo": "request-42"},
            schedule={"seq": "1", "listSeq": "16"},
            search={
                "rct": "VNHCM",
                "dly": "INNXV",
                "etdDate": "09/01/2026",
                "cntr_20sd": "1",
                "cntr_40sd": "1",
                "cntr_40sh": "1",
            },
        )

        self.assertEqual(
            {
                "uuid": "quote-uuid",
                "rct": "VNHCM",
                "dly": "INNXV",
                "etdDate": "09/01/2026",
                "final2sdCnt": "1",
                "final4sdCnt": "1",
                "final4shCnt": "1",
                "rqstNo": "request-42",
                "seq": "1",
                "listSeq": "16",
            },
            payload,
        )
        self.assertNotIn("cntr_20sd", payload)
        self.assertNotIn("cntr_40sd", payload)
        self.assertNotIn("cntr_40sh", payload)

    def test_hydrates_waiting_schedule_with_priced_detail_without_losing_route(self):
        self.assertIsNotNone(
            hydrate_quote_with_detail,
            "EMC must combine the list route with the priced detail response",
        )

        summary = {
            "seq": "1",
            "listSeq": "16",
            "inventory": {"status": "W"},
            "legInfo": [{"rtemp2Depdate": "20260907", "rtemp2VslName": "EVER OUTWIT"}],
        }
        detail = {
            "inventory": {"status": "P"},
            "contract": {"of": [{"chrgRevUnit": "2SD", "chrgPrice": "600"}]},
            "freeTime": {"dly": [{"dmdtType": "DM", "dmdtFreedays": "14"}]},
        }

        hydrated = hydrate_quote_with_detail(summary, detail)

        self.assertEqual("P", hydrated["inventory"]["status"])
        self.assertEqual("600", hydrated["contract"]["of"][0]["chrgPrice"])
        self.assertEqual(summary["legInfo"], hydrated["legInfo"])
        self.assertEqual("14", hydrated["freeTime"]["dly"][0]["dmdtFreedays"])

    def test_only_hydrates_summary_cards_inside_max_etd_window(self):
        self.assertIsNotNone(
            quote_departure_is_on_or_before,
            "EMC must avoid requesting detail for schedules past the 21-day ETD cap",
        )

        max_etd = datetime(2026, 9, 15)
        inside = {"legInfo": [{"rtemp2Depdate": "20260907"}]}
        outside = {"legInfo": [{"rtemp2Depdate": "20260920"}]}

        self.assertTrue(quote_departure_is_on_or_before(inside, max_etd))
        self.assertFalse(quote_departure_is_on_or_before(outside, max_etd))


if __name__ == "__main__":
    unittest.main()
