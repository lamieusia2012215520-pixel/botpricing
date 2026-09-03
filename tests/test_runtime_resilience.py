import threading
import time
import unittest

from bot_runtime_utils import (
    is_transient_webdriver_error,
    switch_to_live_window,
    wait_for_terminal_enter,
)


class _FakeSwitchTo:
    def __init__(self, driver):
        self.driver = driver

    def window(self, handle):
        self.driver.switch_calls.append(handle)
        if handle in self.driver.rejected_handles:
            raise RuntimeError("target window already closed")
        self.driver.current = handle


class _FakeDriver:
    def __init__(self, handles, rejected_handles=()):
        self.window_handles = list(handles)
        self.rejected_handles = set(rejected_handles)
        self.switch_calls = []
        self.current = None
        self.switch_to = _FakeSwitchTo(self)


class RuntimeResilienceTests(unittest.TestCase):
    def test_manual_enter_returns_without_waiting_for_timeout(self):
        self.assertTrue(
            wait_for_terminal_enter(
                lambda prompt: "",
                timeout_seconds=0.5,
                prompt="continue",
            )
        )

    def test_manual_enter_stops_at_deadline_when_input_is_still_blocked(self):
        never = threading.Event()
        started = time.monotonic()

        result = wait_for_terminal_enter(
            lambda prompt: never.wait(5),
            timeout_seconds=0.03,
            prompt="continue",
        )

        self.assertFalse(result)
        self.assertLess(time.monotonic() - started, 0.5)

    def test_stale_preferred_window_falls_back_to_a_live_handle(self):
        driver = _FakeDriver(["live-a", "live-b"])

        selected = switch_to_live_window(driver, preferred_handle="stale")

        self.assertEqual("live-a", selected)
        self.assertEqual(["live-a"], driver.switch_calls)

    def test_rejected_first_window_tries_the_next_live_handle(self):
        driver = _FakeDriver(["bad", "good"], rejected_handles={"bad"})

        selected = switch_to_live_window(driver)

        self.assertEqual("good", selected)
        self.assertEqual(["bad", "good"], driver.switch_calls)

    def test_transient_browser_errors_are_classified_for_retry(self):
        for message in (
            "Message: renderer timeout",
            "Message: script timeout",
            "no such window: target window already closed",
            "disconnected: not connected to DevTools",
        ):
            with self.subTest(message=message):
                self.assertTrue(is_transient_webdriver_error(RuntimeError(message)))

        self.assertFalse(is_transient_webdriver_error(RuntimeError("No rates found")))


if __name__ == "__main__":
    unittest.main()
