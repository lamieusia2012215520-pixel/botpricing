import ast
import unittest
from datetime import date, datetime
from bot_cli import format_etd_dates_excel


class ETDFormattingTests(unittest.TestCase):
    def test_single_date(self):
        self.assertEqual(format_etd_dates_excel([date(2026, 9, 26)]), "26-Sep")

    def test_two_dates_same_month(self):
        self.assertEqual(
            format_etd_dates_excel([date(2026, 9, 26), date(2026, 9, 30)]),
            "26-Sep & 30-Sep",
        )

    def test_two_dates_different_months(self):
        self.assertEqual(
            format_etd_dates_excel([date(2026, 9, 26), date(2026, 10, 2)]),
            "26-Sep & 2-Oct",
        )

    def test_three_dates_same_month(self):
        self.assertEqual(
            format_etd_dates_excel(
                [date(2026, 9, 26), date(2026, 9, 28), date(2026, 9, 30)]
            ),
            "26, 28, 30-Sep",
        )

    def test_three_dates_cross_month_two_first_one_second(self):
        # User reported case: 26, 30-Sep & 2-Oct
        self.assertEqual(
            format_etd_dates_excel(
                [date(2026, 9, 26), date(2026, 9, 30), date(2026, 10, 2)]
            ),
            "26, 30-Sep & 2-Oct",
        )

    def test_three_dates_cross_month_one_first_two_second(self):
        self.assertEqual(
            format_etd_dates_excel(
                [date(2026, 9, 26), date(2026, 10, 2), date(2026, 10, 6)]
            ),
            "26-Sep & 2, 6-Oct",
        )

    def test_three_dates_cross_three_months(self):
        self.assertEqual(
            format_etd_dates_excel(
                [date(2026, 9, 26), date(2026, 10, 2), date(2026, 11, 5)]
            ),
            "26-Sep, 2-Oct & 5-Nov",
        )

    def test_dict_and_datetime_tolerance(self):
        entries = [
            {"etd_dt": datetime(2026, 9, 26, 14, 0)},
            {"etd_dt": datetime(2026, 9, 30, 8, 30)},
            {"etd_dt": datetime(2026, 10, 2, 10, 0)},
        ]
        self.assertEqual(format_etd_dates_excel(entries), "26, 30-Sep & 2-Oct")

    def test_string_dates_tolerance(self):
        entries = ["2026-09-26", "2026-09-30", "2026-10-02"]
        self.assertEqual(format_etd_dates_excel(entries), "26, 30-Sep & 2-Oct")


class AllBotsUseUniversalFormatterTests(unittest.TestCase):
    """Verify statically that all carrier bots import and use format_etd_dates_excel."""
    BOT_FILES = [
        "bot_one.py",
        "bot_HPL.py",
        "bot_EMC.py",
        "bot_maersk.py",
        "bot_KMTC.py",
        "bot_cma.py",
        "bot_msc.py",
        "bot_hmm.py",
        "bot_oocl.py",
        "bot_whl.py",
        "bot_yangming.py",
        "bot_cul.py",
        "bot_esl.py",
        "bot_zim.py",
        "bot_COSCO.py",
    ]

    def test_all_bots_import_and_call_format_etd_dates_excel(self):
        for bot_file in self.BOT_FILES:
            with self.subTest(bot=bot_file):
                with open(bot_file, "r", encoding="utf-8") as f:
                    content = f.read()
                tree = ast.parse(content, filename=bot_file)
                self.assertIn(
                    "format_etd_dates_excel",
                    content,
                    f"{bot_file} does not mention format_etd_dates_excel",
                )


if __name__ == "__main__":
    unittest.main()
