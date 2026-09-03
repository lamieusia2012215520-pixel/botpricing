"""Pure HPL state helpers used by the Selenium bot and regression tests."""

from __future__ import annotations

import base64
import json
import re
import time
from typing import Iterable


def _normalise_port(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def hpl_selected_port_matches(
    current_value: object,
    port_name: object,
    aliases: Iterable[object] = (),
    field_classes: object = "",
) -> bool:
    """Return True only for a value backed by a selected Quasar option.

    Text typed into a Quasar select remains in the input even when no location
    object was selected.  In that state Quasar adds ``q-select--empty`` and/or
    ``q-field--error``; accepting that text would submit an invalid route.
    """

    classes = {part.lower() for part in str(field_classes or "").split()}
    if "q-select--empty" in classes or "q-field--error" in classes:
        return False

    current = _normalise_port(current_value)
    if not current:
        return False

    for candidate in (port_name, *tuple(aliases or ())):
        token = _normalise_port(candidate)
        if token and (token in current or current in token):
            return True
    return False


def jwt_is_expired(token: object, now: float | None = None, skew_seconds: int = 30) -> bool:
    """Check a JWT ``exp`` claim without exposing or validating the token.

    Missing/malformed tokens return False: this helper is used to detect the
    known stale-session case, not to decide whether a user is authenticated.
    """

    raw = str(token or "").strip()
    try:
        payload_part = raw.split(".")[1]
        payload_part += "=" * (-len(payload_part) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_part).decode("utf-8"))
        expiry = float(payload["exp"])
    except (IndexError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return False

    current = time.time() if now is None else float(now)
    return expiry <= current + max(0, int(skew_seconds))

