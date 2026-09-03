import os
import json
import re
import urllib.request


CHINA_TOKENS = {
    "CHINA", "CN", "CHN", "HONG KONG", "HK",
    "SHANGHAI", "NINGBO", "QINGDAO", "XINGANG", "TIANJIN", "YANTIAN",
    "SHEKOU", "NANSHA", "XIAMEN", "DALIAN", "SHENZHEN", "GUANGZHOU",
}

JAPAN_TOKENS = {
    "JAPAN", "JP", "JPN",
    "TOKYO", "YOKOHAMA", "NAGOYA", "OSAKA", "KOBE", "MOJI", "HAKATA",
}

EUROPE_TOKENS = {
    "EUROPE", "EU", "GERMANY", "DE", "DEU", "FRANCE", "FR", "FRA",
    "NETHERLANDS", "NL", "NLD", "BELGIUM", "BE", "BEL", "SPAIN", "ES",
    "ESP", "ITALY", "IT", "ITA", "UNITED KINGDOM", "UK", "GB", "GBR",
    "PORTUGAL", "PT", "PRT", "SWEDEN", "SE", "SWE", "DENMARK", "DK",
    "DNK", "FINLAND", "FI", "FIN", "POLAND", "PL", "POL", "NORWAY",
    "NO", "NOR", "AUSTRIA", "AT", "AUT", "SWITZERLAND", "CH", "CHE",
    "GREECE", "GR", "GRC", "IRELAND", "IE", "IRL", "CZECH", "CZ",
    "HUNGARY", "HU", "HUN", "SLOVENIA", "SI", "CROATIA", "HR",
    "ROMANIA", "RO", "BULGARIA", "BG", "HAMBURG", "BREMERHAVEN",
    "ROTTERDAM", "ANTWERP", "FELIXSTOWE", "SOUTHAMPTON", "LE HAVRE",
    "DUNKERQUE",
    "FOS", "GENOA", "GENOVA", "LA SPEZIA", "NAPLES", "NAPOLI",
    "VALENCIA", "BARCELONA", "ALGECIRAS", "PIRAEUS", "KOPER", "RIJEKA",
    "GDANSK", "GDYNIA", "BASEL", "BASLE",
}

AMERICAS_TOKENS = {
    "AMERICA", "AMERICAS", "NORTH AMERICA", "SOUTH AMERICA", "LATIN AMERICA",
    "UNITED STATES", "USA", "US", "CANADA", "CA", "MEXICO", "MX",
    "BRAZIL", "BR", "ARGENTINA", "AR", "CHILE", "CL", "COLOMBIA", "CO",
    "PERU", "PE", "ECUADOR", "EC", "PANAMA", "PA", "COSTA RICA",
    "GUATEMALA", "HONDURAS", "EL SALVADOR", "NICARAGUA", "DOMINICAN",
    "PUERTO RICO", "JAMAICA", "HAITI", "TRINIDAD", "VENEZUELA", "URUGUAY",
    "PARAGUAY", "BOLIVIA", "MANZANILLO", "LOS ANGELES", "LONG BEACH",
    "NEW YORK", "SAVANNAH", "HOUSTON", "VANCOUVER", "TORONTO", "MONTREAL",
}

AFRICA_TOKENS = {
    "AFRICA", "NIGERIA", "NG", "LAGOS", "APAPA", "GHANA", "TEMA",
    "SOUTH AFRICA", "DURBAN", "CAPE TOWN", "EGYPT", "ALEXANDRIA",
    "KENYA", "MOMBASA", "TANZANIA", "DAR ES SALAAM", "MOROCCO",
    "ALGERIA", "TUNISIA", "SENEGAL", "DAKAR", "COTE DIVOIRE",
    "IVORY COAST", "ABIDJAN", "CAMEROON", "DOUALA", "ANGOLA", "LUANDA",
    "MOZAMBIQUE", "MAPUTO", "DJIBOUTI", "ETHIOPIA",
}

MANIFEST_CODES = {"AFS", "AFR", "ENS", "AMS"}


def _norm(value):
    return re.sub(r"\s+", " ", str(value or "").upper()).strip()


def _has_token(text, token):
    token = _norm(token)
    if not token:
        return False
    if len(token) <= 3:
        return bool(re.search(rf"(?<![A-Z0-9]){re.escape(token)}(?![A-Z0-9])", text))
    return token in text


def _has_any(text, tokens):
    return any(_has_token(text, token) for token in tokens)


def is_china_destination(country="", pod=""):
    """Return True for China/Hong Kong country names, codes, or known POD names."""
    text = _norm(f"{country} {pod}")
    return bool(text and _has_any(text, CHINA_TOKENS))


_CHARGE_FX_CACHE = {"USD": 1.0}
_CHARGE_FX_FALLBACK = {
    "EUR": 1.08,
    "CHF": 1.12,
    "AUD": 0.66,
    "GBP": 1.27,
    "JPY": 0.0068,
    "CNY": 0.14,
    "VND": 0.00004,
    "THB": 0.028,
    "NZD": 0.60,
}


def charge_amount_to_usd(amount, currency="USD", rate_getter=None):
    """Convert a separately shown charge to USD, using a carrier FX getter when available."""
    value = float(amount or 0)
    base = str(currency or "USD").strip().upper()
    if not base or base == "USD":
        return value

    if rate_getter is not None:
        try:
            try:
                raw_rate = rate_getter(base, "USD")
            except TypeError:
                raw_rate = rate_getter(base)
            rate = float(raw_rate)
            if rate > 0:
                return value * rate
        except Exception:
            pass

    if base not in _CHARGE_FX_CACHE:
        env_value = os.environ.get(f"CHARGE_{base}_TO_USD", "").strip()
        try:
            rate = float(env_value)
        except (TypeError, ValueError):
            try:
                with urllib.request.urlopen(
                    f"https://api.exchangerate-api.com/v4/latest/{base}", timeout=3
                ) as response:
                    rate = float(json.loads(response.read().decode("utf-8"))["rates"]["USD"])
            except Exception:
                rate = _CHARGE_FX_FALLBACK.get(base)
        if rate is None:
            raise ValueError(f"No {base}->USD exchange rate for origin THC")
        _CHARGE_FX_CACHE[base] = rate
    return value * _CHARGE_FX_CACHE[base]


def get_manifest_code(country="", pod=""):
    text = _norm(f"{country} {pod}")
    if not text:
        return ""
    if _has_any(text, CHINA_TOKENS):
        return "AFS"
    if _has_any(text, JAPAN_TOKENS):
        return "AFR"
    if _has_any(text, EUROPE_TOKENS):
        return "ENS"
    if _has_any(text, AMERICAS_TOKENS) or _has_any(text, AFRICA_TOKENS):
        return "AMS"
    return ""


def normalize_remark_text(remark):
    text = str(remark or "")
    text = re.sub(r"\bTELEX\b", "TLX", text, flags=re.IGNORECASE)
    text = re.sub(r"\bINCL\s+O\.THC\b", "INCLUDED O.THC", text, flags=re.IGNORECASE)
    return text


def apply_manifest_rule(remark, country="", pod="", keep_existing_if_unknown=True):
    text = normalize_remark_text(remark)
    code = get_manifest_code(country, pod)
    if not code:
        return text if keep_existing_if_unknown else _strip_manifest_codes(text)
    text = _strip_manifest_codes(text)
    return f"{text}, {code}" if text else code


def _strip_manifest_codes(remark):
    text = str(remark or "")
    for code in MANIFEST_CODES:
        text = re.sub(rf"(?i)(,\s*)?\b{code}\b", "", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",\s*,", ",", text)
    return text.strip(" ,")


def build_subject_remark(othc_included=False, country="", pod="", ows=False, extra_items=None):
    items = ["BILL", "SEAL", "TLX"]
    if not othc_included:
        items.insert(0, "THC")
    if extra_items:
        for item in extra_items:
            item = str(item or "").strip().upper()
            if item and item not in items:
                items.append(item)
    manifest = get_manifest_code(country, pod)
    if manifest and manifest not in items:
        items.append(manifest)
    if ows and "OWS" not in items:
        items.append("OWS")
    remark = "SUBJECT TO " + ", ".join(items)
    if othc_included:
        remark = "INCLUDED O.THC, " + remark
    return normalize_remark_text(remark)
