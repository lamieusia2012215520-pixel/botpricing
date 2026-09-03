"""Small, side-effect-free helpers for COSCO E-Lines' current booking UI.

The Selenium bot owns browser actions.  Keeping the text/URL interpretation
here makes the production decisions testable without launching Edge.
"""

from __future__ import annotations

import re
from datetime import datetime
from urllib.parse import urlparse


ELINES_BOOKING_PATH = "/ebusiness/bookingrequest"


def is_new_elines_result_card_text(text: str) -> bool:
    """Recognise a current booking-request result card from stable labels.

    ``View Premium`` is deliberately not required: COSCO also renders current-
    UI schedule cards that only have ``Book Now`` and therefore carry no
    premium freight price for the bot to collect.
    """
    normalized = " ".join(str(text or "").split()).upper()
    has_vessel_label = "VESSEL / VOYAGE" in normalized or "VESSEL/VOYAGE" in normalized
    has_action = "BOOK NOW" in normalized or "VIEW PREMIUM" in normalized
    return "CY CUTOFF" in normalized and has_vessel_label and has_action


def dropdown_has_selected_option(selected_labels, option_text: str, prefix: bool = False) -> bool:
    """Return whether an Element Plus select already contains an option.

    Multi-select controls display selections as tags while their input value is
    empty, so callers must inspect these labels before clicking a toggle.
    """
    target = " ".join(str(option_text or "").split()).upper()
    if not target:
        return False
    for label in selected_labels or []:
        value = " ".join(str(label or "").split()).upper()
        if value == target or (prefix and value.startswith(target)):
            return True
    return False


def is_elines_booking_page(url: str) -> bool:
    """Return whether *url* is the authenticated new E-Lines booking page."""
    try:
        parsed = urlparse(url or "")
    except ValueError:
        return False
    return (
        parsed.netloc.lower().endswith("elines.coscoshipping.com")
        and parsed.path.rstrip("/").lower() == ELINES_BOOKING_PATH
    )


def is_elines_auth_page(url: str) -> bool:
    """Recognise E-Lines' login notice and external SSO pages."""
    try:
        parsed = urlparse(url or "")
    except ValueError:
        return False
    host = parsed.netloc.lower()
    path = parsed.path.rstrip("/").lower()
    return (
        (host.endswith("elines.coscoshipping.com") and "loginplease" in path)
        or "exiamfw" in host
        or "/auth/" in path
    )


def is_no_matching_ocean_freight_message(text: str) -> bool:
    """Recognise both legacy and new E-Lines' definitive no-product notices."""
    normalized = " ".join(str(text or "").split()).upper()
    return (
        "NO PRODUCTS MATCHING YOUR CRITERIA WERE FOUND" in normalized
        or "NO MATCHING OCEAN FREIGHT PRODUCTS WERE FOUND" in normalized
    )


def parse_premium_service_row(text: str, service: str) -> dict:
    """Extract container prices and Flash Sale stock from one modal service row."""
    normalized = " ".join(str(text or "").split())
    rates: dict[str, float] = {}
    for container, raw_price in re.findall(
        r"\b(20GP|40GP|40HQ)\b\s*(?:FROM\s*)?USD\s*([0-9][0-9,]*(?:\.\d+)?)",
        normalized,
        flags=re.IGNORECASE,
    ):
        rates[container.upper()] = float(raw_price.replace(",", ""))

    stock_match = re.search(r"REMAINING\s+STOCK\s*:\s*(\d+)\s*TEU\b", normalized, re.I)
    return {
        "service": " ".join(str(service or "").split()),
        "rates": rates,
        "remaining_stock": int(stock_match.group(1)) if stock_match else None,
    }


def select_preferred_premium_service(rows: list[dict]) -> dict | None:
    """Use Flash Sale where present; otherwise retain the standard service."""
    usable = [row for row in rows if row and row.get("rates")]
    if not usable:
        return None
    for row in usable:
        if str(row.get("service", "")).strip().upper() == "FLASH SALE":
            return row
    return usable[0]


def parse_new_elines_card_schedule(text: str, now: datetime | None = None) -> dict:
    """Parse the displayed departure day, transit days, and space label on a card."""
    normalized = " ".join(str(text or "").split())
    now = now or datetime.now()
    date_match = re.search(
        r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*([0-3]?\d)\b",
        normalized,
        flags=re.IGNORECASE,
    )
    days_match = re.search(r"\b(\d+)\s*DAYS?\b", normalized, flags=re.IGNORECASE)
    if not date_match or not days_match:
        raise ValueError(f"Cannot parse E-Lines card schedule: {normalized!r}")

    month = date_match.group(1).title()
    day = int(date_match.group(2))
    etd = datetime.strptime(f"{now.year}{month}{day}", "%Y%b%d")
    if (etd - now).days < -60:
        etd = etd.replace(year=etd.year + 1)

    upper = normalized.upper()
    if re.search(r"\bTIGHT\b", upper):
        space = "TIGHT"
    elif re.search(r"\bTBC\b", upper):
        space = "TBC"
    else:
        space = "TBC"

    return {"etd_dt": etd, "tt_days": int(days_match.group(1)), "space": space}
