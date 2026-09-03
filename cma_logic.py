"""Pure helpers for fast and deterministic CMA result-card processing."""

from __future__ import annotations

import re
from datetime import datetime


_CMA_PRICE_RE = re.compile(r"[\d\s,]+\s*(?:USD|US\$|\$|EUR)\b", re.IGNORECASE)
_CMA_TERMINAL_CARD_MARKERS = (
    "SOLD OUT",
    "NO OFFER",
    "MODIFY TEU",
    "NO SPACE",
    "FULLY BOOKED",
)


def classify_cma_card_texts(card_texts):
    """Classify a result snapshot without requiring every card to have a price."""

    texts = [" ".join(str(text or "").upper().split()) for text in card_texts]
    texts = [text for text in texts if text]
    if not texts:
        return "EMPTY"

    resolved = [
        bool(_CMA_PRICE_RE.search(text))
        or any(marker in text for marker in _CMA_TERMINAL_CARD_MARKERS)
        for text in texts
    ]
    if all(resolved):
        return "READY"
    if any(resolved):
        return "PARTIAL"
    return "LOADING"


def dedupe_cma_card_summaries(cards):
    """Remove exact duplicate offers while preserving the first DOM card index."""

    unique = []
    seen = set()
    for card in cards:
        date_value = card.get("date")
        if hasattr(date_value, "isoformat"):
            date_value = date_value.isoformat()
        key = (
            date_value,
            card.get("price"),
            card.get("transit"),
            " ".join(str(card.get("ts_port") or "").upper().split()),
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(card)
    return unique


def cma_date_input_matches(raw_value, target_date):
    """Accept the date formats used by CMA's visible/hidden departure input."""

    value = str(raw_value or "").strip()
    if not value:
        return False
    candidates = [value]
    candidates.extend(
        match.group(0)
        for match in re.finditer(
            r"\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4}|\d{1,2}[-/][A-Za-z]{3}[-/]\d{4}",
            value,
        )
    )
    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%b-%Y",
        "%d/%b/%Y",
    )
    for candidate in candidates:
        for date_format in formats:
            try:
                parsed = datetime.strptime(candidate, date_format)
            except ValueError:
                continue
            return parsed.date() == target_date.date()
    return False


def parse_cma_comparable_price(value):
    """Return a numeric CMA price for PORT/RAMP comparison.

    CMA writes simple additive Excel formulas (for example ``=1911+150-50``)
    so calling ``float`` directly loses otherwise valid prices.
    """

    if isinstance(value, (int, float)):
        return float(value)
    text = str(value or "").strip().replace(",", "")
    if not text:
        return float("inf")
    if text.startswith("="):
        expression = text[1:]
        if not re.fullmatch(
            r"\s*[+-]?\d+(?:\.\d+)?(?:\s*[+-]\s*\d+(?:\.\d+)?)*\s*",
            expression,
        ):
            return float("inf")
        terms = re.findall(r"[+-]?\s*\d+(?:\.\d+)?", expression)
        return sum(float(term.replace(" ", "")) for term in terms)
    try:
        return float(text)
    except ValueError:
        return float("inf")
