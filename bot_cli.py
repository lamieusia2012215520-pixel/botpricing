import os
import sys
from datetime import date as date_type
from datetime import datetime, timedelta


MAX_ETD_DAYS = int(os.environ.get("MAX_ETD_DAYS", "21"))


def max_etd_date(days=None):
    days = MAX_ETD_DAYS if days is None else int(days)
    return (datetime.now() + timedelta(days=days)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )


def max_etd_date_only(days=None):
    return max_etd_date(days).date()


def etd_within_max(etd, days=None):
    if etd is None:
        return False
    limit_day = max_etd_date(days).date()
    if isinstance(etd, datetime):
        return etd.date() <= limit_day
    if isinstance(etd, date_type):
        return etd <= limit_day
    return etd <= limit_day


def parse_date_offset_days(default=7, env_key="DATE_OFFSET_DAYS", argv=None):
    """
    Shared lightweight parser for standalone bot scripts.

    Supports:
      python bot_x.py --date +2
      python bot_x.py --date=+2

    main.py still passes DATE_OFFSET_DAYS via environment; direct CLI --date
    takes priority when present.
    """
    args = list(sys.argv[1:] if argv is None else argv)
    raw = None

    for idx, arg in enumerate(args):
        lower = str(arg or "").lower()
        if lower == "--date" and idx + 1 < len(args):
            raw = args[idx + 1]
            break
        if lower.startswith("--date="):
            raw = str(arg).split("=", 1)[1]
            break

    if raw is None:
        raw = os.environ.get(env_key, str(default))

    try:
        value = int(str(raw or default).strip().lstrip("+"))
    except Exception:
        value = int(default)

    value = max(0, value)
    os.environ[env_key] = str(value)
    return value
