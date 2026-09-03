"""Small runtime helpers shared by the Selenium carrier bots.

The functions in this module deliberately do not touch a browser unless their
name says so.  In particular, ``wait_for_terminal_enter`` leaves Selenium
completely idle while a user solves OTP/CAPTCHA in the visible browser.
"""

from __future__ import annotations

import threading


def wait_for_terminal_enter(input_func, timeout_seconds, prompt=""):
    """Wait for one terminal input with a hard deadline.

    ``input()`` has no portable timeout.  Run it in a daemon thread so the bot
    can enforce a deadline without polling or otherwise touching the browser.
    The carrier process exits after a failed manual check, so a timed-out daemon
    input thread cannot keep that process alive.
    """

    completed = threading.Event()
    result = {"error": None}

    def read_input():
        try:
            input_func(prompt)
        except BaseException as exc:  # input can raise EOFError/KeyboardInterrupt
            result["error"] = exc
        finally:
            completed.set()

    worker = threading.Thread(
        target=read_input,
        name="manual-verification-input",
        daemon=True,
    )
    worker.start()

    try:
        timeout = max(0.0, float(timeout_seconds))
    except (TypeError, ValueError):
        timeout = 0.0
    if not completed.wait(timeout):
        return False

    error = result["error"]
    if isinstance(error, KeyboardInterrupt):
        raise error
    return error is None


def switch_to_live_window(driver, preferred_handle=None):
    """Select a usable browser window, tolerating stale stored handles."""

    try:
        handles = list(driver.window_handles)
    except Exception as exc:
        raise RuntimeError("Không đọc được danh sách browser window") from exc

    candidates = []
    if preferred_handle and preferred_handle in handles:
        candidates.append(preferred_handle)
    candidates.extend(handle for handle in handles if handle not in candidates)

    last_error = None
    for handle in candidates:
        try:
            driver.switch_to.window(handle)
            return handle
        except Exception as exc:
            last_error = exc

    raise RuntimeError("Không còn browser window nào sử dụng được") from last_error


def is_transient_webdriver_error(error):
    """Return True for browser/session failures worth retrying once."""

    message = f"{type(error).__name__}: {error}".lower()
    markers = (
        "renderer timeout",
        "script timeout",
        "no such window",
        "target window already closed",
        "web view not found",
        "disconnected",
        "not connected to devtools",
        "chrome not reachable",
        "edge not reachable",
        "invalid session id",
        "tab crashed",
        "timed out receiving message from renderer",
    )
    return any(marker in message for marker in markers)
