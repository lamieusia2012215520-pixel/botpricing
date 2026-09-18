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


def _to_date(item):
    if item is None:
        return None
    if hasattr(item, "date"):
        try:
            return item.date()
        except TypeError:
            pass
    if isinstance(item, date_type):
        return item
    if isinstance(item, dict):
        for k in ("etd_dt", "etd", "etd_date", "date", "dt"):
            val = item.get(k)
            if val is not None:
                d = _to_date(val)
                if d:
                    return d
    if isinstance(item, str):
        item_str = item.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d-%b-%Y", "%d %b %Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(item_str, fmt).date()
            except Exception:
                pass
    return None


def format_etd_dates_excel(dates_or_entries):
    """
    Format list of ETDs according to company standard:
    - 1 date: '26-Sep'
    - 2 dates: '26-Sep & 30-Sep' or '26-Sep & 2-Oct'
    - 3 dates same month: '26, 28, 30-Sep'
    - 3 dates across 2 months: '26, 30-Sep & 2-Oct' or '26-Sep & 2, 6-Oct'
    - 3 dates across 3 months: '26-Sep, 2-Oct & 5-Nov'
    """
    if not dates_or_entries:
        return ""
    dates = []
    for item in dates_or_entries:
        d = _to_date(item)
        if d:
            dates.append(d)
    if not dates:
        return ""
    ordered = sorted(dates)
    if len(ordered) == 1:
        return f"{ordered[0].day}-{ordered[0].strftime('%b')}"
    if len(ordered) == 2:
        return f"{ordered[0].day}-{ordered[0].strftime('%b')} & {ordered[1].day}-{ordered[1].strftime('%b')}"

    groups = []
    for dt in ordered:
        key = (dt.year, dt.month)
        if not groups or groups[-1][0] != key:
            groups.append((key, [dt]))
        else:
            groups[-1][1].append(dt)

    parts = []
    for _, group_dates in groups:
        month = group_dates[-1].strftime("%b")
        if len(group_dates) == 1:
            parts.append(f"{group_dates[0].day}-{month}")
        else:
            days = ", ".join(str(d.day) for d in group_dates)
            parts.append(f"{days}-{month}")

    if len(parts) == 1:
        return parts[0]
    return f"{', '.join(parts[:-1])} & {parts[-1]}"

