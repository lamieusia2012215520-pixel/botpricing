"""Pure helpers for GreenX's two-step quote API.

GreenX returns a lightweight schedule list first.  Price, inventory status and
free time only exist in a second request for the selected list item.
"""

from datetime import datetime


def build_quote_detail_payload(quote_result, schedule, search):
    """Build GreenX_GetQuoteResultDetail payload exactly as the current UI does."""
    quote_result = quote_result or {}
    schedule = schedule or {}
    search = search or {}
    return {
        "uuid": quote_result.get("uuid"),
        "rct": search.get("rct"),
        "dly": search.get("dly"),
        "etdDate": search.get("etdDate"),
        "final2sdCnt": search.get("cntr_20sd"),
        "final4sdCnt": search.get("cntr_40sd"),
        "final4shCnt": search.get("cntr_40sh"),
        "rqstNo": quote_result.get("rqstNo"),
        "seq": schedule.get("seq"),
        "listSeq": schedule.get("listSeq"),
    }


def hydrate_quote_with_detail(schedule, detail):
    """Keep the route from the list response and add the priced detail fields."""
    hydrated = dict(schedule or {})
    detail = detail or {}
    for key in ("inventory", "contract", "freeTime", "mainLine"):
        if key in detail:
            hydrated[key] = detail[key]
    return hydrated


def quote_departure_is_on_or_before(schedule, max_etd):
    """Return whether a lightweight GreenX card is within the ETD cap."""
    try:
        leg_info = (schedule or {}).get("legInfo") or []
        etd_raw = leg_info[0].get("rtemp2Depdate") if leg_info else None
        etd = datetime.strptime(str(etd_raw), "%Y%m%d")
        return etd <= max_etd
    except (AttributeError, TypeError, ValueError):
        return False
