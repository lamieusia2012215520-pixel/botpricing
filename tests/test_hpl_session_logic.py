import base64
import json
import unittest

from hpl_logic import hpl_selected_port_matches, jwt_is_expired


def _fake_jwt(exp):
    def part(value):
        raw = json.dumps(value, separators=(",", ":")).encode("utf-8")
        return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")

    return f"{part({'alg': 'none'})}.{part({'exp': exp})}.signature"


class HPLSessionLogicTests(unittest.TestCase):
    def test_expired_hpl_token_is_detected_with_clock_skew(self):
        self.assertTrue(jwt_is_expired(_fake_jwt(1_000), now=980, skew_seconds=30))
        self.assertFalse(jwt_is_expired(_fake_jwt(1_100), now=980, skew_seconds=30))

    def test_missing_or_malformed_token_is_not_misclassified_as_expired(self):
        self.assertFalse(jwt_is_expired("", now=1_000))
        self.assertFalse(jwt_is_expired("not-a-jwt", now=1_000))

    def test_existing_selected_hcm_alias_is_reused(self):
        self.assertTrue(
            hpl_selected_port_matches(
                "Ho Chi Minh City",
                "VNSGN",
                aliases=["Ho Chi Minh", "Ho Chi Minh City"],
                field_classes="q-field q-select q-field--outlined",
            )
        )

    def test_unselected_raw_text_is_not_reused(self):
        self.assertFalse(
            hpl_selected_port_matches(
                "VNSGN",
                "VNSGN",
                aliases=["Ho Chi Minh City"],
                field_classes="q-field q-select q-select--empty q-field--error",
            )
        )


if __name__ == "__main__":
    unittest.main()
