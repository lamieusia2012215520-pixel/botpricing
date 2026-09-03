import unittest

from bot_oocl import OOCLCombinedBot, OOCLSetupTabsError


class FakeOOCLBot(OOCLCombinedBot):
    def __init__(self, setup_outcomes, restart_ok=True):
        self.setup_outcomes = list(setup_outcomes)
        self.restart_ok = restart_ok
        self.setup_attempts = 0
        self.stop_calls = 0
        self.restart_calls = 0

    def setup_tabs(self):
        self.setup_attempts += 1
        outcome = self.setup_outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome

    def _stop_setup_page_load(self):
        self.stop_calls += 1

    def _restart_oocl_edge_for_setup(self):
        self.restart_calls += 1
        return self.restart_ok


class _NoTabDriver:
    window_handles = []


class FatalRunBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _NoTabDriver()

    def launch_edge_if_needed(self):
        return True

    def init_browser(self):
        return True

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def check_and_login(self):
        return True

    def setup_tabs_with_recovery(self):
        raise OOCLSetupTabsError("renderer recovery exhausted")


class OOCLSetupRecoveryTests(unittest.TestCase):
    def test_retries_current_session_after_setup_failure(self):
        bot = FakeOOCLBot([False, True])

        self.assertTrue(bot.setup_tabs_with_recovery())
        self.assertEqual(bot.setup_attempts, 2)
        self.assertEqual(bot.stop_calls, 1)
        self.assertEqual(bot.restart_calls, 0)

    def test_restarts_dedicated_edge_after_current_session_retry_fails(self):
        bot = FakeOOCLBot([False, False, True])

        self.assertTrue(bot.setup_tabs_with_recovery())
        self.assertEqual(bot.setup_attempts, 3)
        self.assertEqual(bot.stop_calls, 1)
        self.assertEqual(bot.restart_calls, 1)

    def test_raises_fatal_error_when_all_setup_recovery_attempts_fail(self):
        bot = FakeOOCLBot([False, False, False])

        with self.assertRaises(OOCLSetupTabsError):
            bot.setup_tabs_with_recovery()
        self.assertEqual(bot.setup_attempts, 3)
        self.assertEqual(bot.restart_calls, 1)

    def test_run_propagates_fatal_setup_error_for_nonzero_process_exit(self):
        with self.assertRaises(OOCLSetupTabsError):
            FatalRunBot().run()


if __name__ == "__main__":
    unittest.main()
