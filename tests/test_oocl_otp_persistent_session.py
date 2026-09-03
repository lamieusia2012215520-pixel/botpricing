import unittest
from unittest.mock import patch
import os
import tempfile

import main as main_module
from bot_oocl import (
    NEW_UI_URL,
    OOCL_LOGIN_MANUAL_WAIT_SECONDS,
    OOCLCombinedBot,
    OOCLManualLoginTimeout,
)


class _FakeSwitchTo:
    def __init__(self, driver):
        self.driver = driver

    def window(self, handle):
        self.driver.active_handle = handle


class _FakeServiceProcess:
    def __init__(self):
        self.kill_calls = 0
        self.wait_calls = 0

    def poll(self):
        return None

    def kill(self):
        self.kill_calls += 1

    def wait(self, timeout=None):
        self.wait_calls += 1


class _FakeService:
    def __init__(self):
        self.stop_calls = 0
        self.process = _FakeServiceProcess()

    def stop(self):
        self.stop_calls += 1


class _FakeDriver:
    def __init__(self, tabs):
        self.tabs = dict(tabs)
        self.active_handle = next(iter(self.tabs))
        self.switch_to = _FakeSwitchTo(self)
        self.get_calls = []
        self.quit_calls = 0
        self.close_calls = 0
        self.service = _FakeService()

    @property
    def window_handles(self):
        return list(self.tabs)

    @property
    def current_window_handle(self):
        return self.active_handle

    @property
    def current_url(self):
        return self.tabs[self.active_handle]

    @current_url.setter
    def current_url(self, value):
        self.tabs[self.active_handle] = value

    def get(self, url):
        self.get_calls.append(url)
        self.current_url = url

    def execute_script(self, script, *args):
        return None

    def quit(self):
        self.quit_calls += 1

    def close(self):
        self.close_calls += 1


class _OtpFlowBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _FakeDriver({"otp": "https://exiamfw.home.oocl.com/otp"})
        self.wait = object()

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _is_oocl_app_ready(self, success_url_keywords=None):
        return "/digital/" in self.driver.current_url


class _OtherWorkspaceOtpFlowBot(_OtpFlowBot):
    def __init__(self):
        super().__init__()
        self.driver.current_url = "https://freightsmart.oocl.com/ui/my-quotation"

    def _is_oocl_app_ready(self, success_url_keywords=None):
        url = self.driver.current_url
        return "freightsmart.oocl.com/ui/" in url or "/digital/" in url


class _BodyElement:
    text = "We sent a code to your email. Enter the code to continue."


class _EmailCodeDriver:
    def find_element(self, by, value):
        return _BodyElement()

    def find_elements(self, by, value):
        return []


class _ExistingSessionBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _FakeDriver({
            "signed-in": "https://freightsmart.oocl.com/digital/",
        })
        self.clean_calls = 0

    def _ensure_session_alive(self):
        return True

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _find_authenticated_oocl_tab(self):
        return "signed-in"

    def _focus_tab(self, handle):
        self.driver.switch_to.window(handle)

    def clean_tabs_and_open_fresh(self):
        self.clean_calls += 1
        return False


class _ExistingOtpSessionBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _FakeDriver({"otp": "https://exiamfw.home.oocl.com/otp"})
        self.clean_calls = 0
        self.manual_calls = 0

    def _ensure_session_alive(self):
        return True

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _find_authenticated_oocl_tab(self):
        return None

    def _find_existing_manual_login_tab(self):
        return "otp"

    def _focus_tab(self, handle):
        self.driver.switch_to.window(handle)

    def _wait_manual_login_done(self, reason="MFA/CAPTCHA", success_url_keywords=None, input_func=None):
        self.manual_calls += 1
        return True

    def clean_tabs_and_open_fresh(self):
        self.clean_calls += 1
        return False


class _WorkspaceStateBot(OOCLCombinedBot):
    def __init__(self, ready_by_handle):
        self.driver = _FakeDriver({
            "spot": "https://freightsmart.oocl.com/digital/",
            "quote": "https://freightsmart.oocl.com/ui/my-quotation",
        })
        self.ready_by_handle = dict(ready_by_handle)

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _is_oocl_app_ready(self, success_url_keywords=None):
        return self.ready_by_handle.get(self.driver.active_handle, False)


class _OtpAndUsernameBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _FakeDriver({"otp": "https://exiamfw.home.oocl.com/otp"})

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _is_oocl_app_ready(self, success_url_keywords=None):
        return False

    def _oocl_manual_login_check_present(self):
        return True

    def _find_visible_xpath(self, xpaths, timeout=0, clickable=False):
        if isinstance(xpaths, str) and "password" in xpaths:
            return None
        return object()


class _LoginFailureBot(OOCLCombinedBot):
    def launch_edge_if_needed(self):
        return True

    def init_browser(self):
        self.driver = _FakeDriver({"login": "https://freightsmart.oocl.com/login"})
        return True

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _authenticated_workspace_tabs_ready(self):
        return False

    def check_and_login(self):
        return False


class _ManualTimeoutDuringLoginBot(OOCLCombinedBot):
    def __init__(self):
        self.driver = _FakeDriver({"otp": "https://exiamfw.home.oocl.com/otp"})

    def _abort_if_oocl_blocked(self, context=""):
        return None

    def _is_oocl_app_ready(self, success_url_keywords=None):
        return False

    def _wait_for_oocl_login_or_app(self, timeout=None, success_url_keywords=None):
        return "MANUAL"

    def _wait_manual_login_done(self, reason="MFA/CAPTCHA", success_url_keywords=None, input_func=None):
        raise OOCLManualLoginTimeout("manual OTP expired")


class OOCLOtpPersistentSessionTests(unittest.TestCase):
    def test_default_otp_manual_window_is_fifteen_minutes(self):
        self.assertEqual(900, OOCL_LOGIN_MANUAL_WAIT_SECONDS)

    def test_otp_timeout_is_fatal_and_removes_parent_input_marker(self):
        bot = _OtpFlowBot()
        with tempfile.TemporaryDirectory() as folder:
            marker_path = os.path.join(folder, "oocl.wait")
            with (
                patch("bot_oocl.OOCL_MANUAL_LOGIN_WAIT_PATH", marker_path),
                patch("bot_oocl.wait_for_terminal_enter", return_value=False),
            ):
                with self.assertRaises(OOCLManualLoginTimeout):
                    bot._wait_manual_login_done("OTP", input_func=lambda prompt: "")

            self.assertFalse(os.path.exists(marker_path))

    def test_manual_timeout_is_not_swallowed_by_login_retry_loop(self):
        with self.assertRaises(OOCLManualLoginTimeout):
            _ManualTimeoutDuringLoginBot()._do_login()

    def test_otp_state_has_priority_over_visible_username_control(self):
        bot = _OtpAndUsernameBot()

        self.assertEqual("MANUAL", bot._wait_for_oocl_login_or_app(timeout=0.1))

    def test_login_failure_is_fatal_instead_of_merging_stale_workbook_data(self):
        with self.assertRaises(RuntimeError):
            _LoginFailureBot().run()

    def test_otp_wait_is_silent_until_enter_then_navigates_to_app(self):
        bot = _OtpFlowBot()
        prompts = []

        def press_enter(prompt):
            prompts.append(prompt)
            self.assertEqual([], bot.driver.get_calls)
            return ""

        self.assertTrue(bot._wait_manual_login_done("OTP", input_func=press_enter))
        self.assertEqual(1, len(prompts))
        self.assertEqual([NEW_UI_URL], bot.driver.get_calls)

    def test_enter_from_another_signed_in_page_returns_to_espot_workspace(self):
        bot = _OtherWorkspaceOtpFlowBot()

        self.assertTrue(bot._wait_manual_login_done("OTP", input_func=lambda prompt: ""))
        self.assertEqual([NEW_UI_URL], bot.driver.get_calls)

    def test_email_code_wording_is_detected_as_manual_login(self):
        bot = OOCLCombinedBot.__new__(OOCLCombinedBot)
        bot.driver = _EmailCodeDriver()

        self.assertTrue(bot._oocl_manual_login_check_present())

    def test_existing_authenticated_session_skips_sso_cleanup_and_login(self):
        bot = _ExistingSessionBot()

        self.assertTrue(bot.check_and_login())
        self.assertEqual(0, bot.clean_calls)
        self.assertEqual("signed-in", bot.driver.active_handle)

    def test_existing_otp_page_is_preserved_instead_of_requesting_a_new_code(self):
        bot = _ExistingOtpSessionBot()

        self.assertTrue(bot.check_and_login())
        self.assertEqual(1, bot.manual_calls)
        self.assertEqual(0, bot.clean_calls)

    def test_url_only_tabs_do_not_bypass_login_when_not_authenticated(self):
        bot = _WorkspaceStateBot({"spot": False, "quote": False})

        self.assertFalse(bot._authenticated_workspace_tabs_ready())

    def test_two_authenticated_workspace_tabs_can_skip_login(self):
        bot = _WorkspaceStateBot({"spot": True, "quote": True})

        self.assertTrue(bot._authenticated_workspace_tabs_ready())

    def test_finish_keeps_edge_and_only_stops_webdriver_service(self):
        bot = _OtpFlowBot()
        driver = bot.driver
        service_process = driver.service.process

        bot.keep_browser_on_finish()

        self.assertEqual(1, service_process.kill_calls)
        self.assertEqual(1, service_process.wait_calls)
        self.assertEqual(0, driver.service.stop_calls)
        self.assertEqual(0, driver.quit_calls)
        self.assertEqual(0, driver.close_calls)
        self.assertIsNone(bot.driver)
        self.assertIsNone(bot.wait)

    def test_main_marks_oocl_as_persistent_edge(self):
        self.assertTrue(main_module.should_keep_bot_edge("bot_oocl.py"))
        self.assertTrue(main_module.should_keep_bot_edge("bot_oocl.py#W1"))
        self.assertTrue(main_module.should_keep_bot_edge("bot_cosco.py"))
        self.assertTrue(main_module.should_keep_bot_edge("bot_hpl.py"))
        self.assertFalse(main_module.should_keep_bot_edge("bot_cma.py"))
        self.assertIn("edge_oocl", main_module.EDGE_KEEP_PATTERNS)
        self.assertIn("remote-debugging-port=9527", main_module.EDGE_KEEP_PATTERNS)

    def test_main_does_not_consume_enter_while_whl_or_oocl_waits_for_user(self):
        self.assertIn(
            main_module.OOCL_MANUAL_LOGIN_WAIT_PATH,
            main_module.MANUAL_INPUT_WAIT_PATHS,
        )
        self.assertIn(
            main_module.WHL_MANUAL_CAPTCHA_WAIT_PATH,
            main_module.MANUAL_INPUT_WAIT_PATHS,
        )

    @patch("subprocess.run")
    def test_main_cleanup_stops_oocl_driver_but_not_oocl_edge(self, run_mock):
        run_mock.return_value.returncode = 0

        self.assertTrue(main_module.kill_specific_bot_edge("bot_oocl.py", bot_pid=321))

        command = run_mock.call_args.args[0]
        powershell = command[-1]
        self.assertIn("msedgedriver.exe", powershell)
        self.assertNotIn("name = 'msedge.exe'", powershell)


if __name__ == "__main__":
    unittest.main()
