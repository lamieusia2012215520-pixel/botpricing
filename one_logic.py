import re

ONE_ORIGIN_LOCAL_EXCLUDED_KEYWORDS = (
    "terminal handling charge (l)",
    "terminal handling charge at origin",
    "terminal handling charge at port of loading",
    "thc/l",
    "thc at origin",
    "doc fee (origin)",
    "document fee",
    "bill of lading",
    "b/l fee",
    "bl fee",
    "seal fee",
    "seal charge",
    "entry summary declaration surcharge",
    "advanced manifest",
    "manifest declaration",
    "manifest fee",
    "customs manifest submission fee",
    "ams",
    "ens",
    "afs",
    "afr",
    "heavy surcharge",
    "heavy lift",
    "overweight",
    "ows",
)

ONE_ORIGIN_LOCAL_EXCLUDED_CODES = {
    "THC", "THCL", "OTHC", "DOC", "BLF", "BLC", "SEAL", "SLF",
    "AMS", "ENS", "AFS", "AFR", "EST", "OWS", "HWC", "HLC", "OOG",
}

ONE_DISCOUNT_KEYWORDS = (
    "special promotion service",
    "special promotion",
    "promotion",
    "discount",
    "rebate",
    "incentive",
)

ONE_DISCOUNT_CODES = {
    "DSC", "DISC", "PROM", "PRM", "SPS",
}


def one_fee_name(charge_or_text):
    if isinstance(charge_or_text, dict):
        return str(charge_or_text.get("chargeName") or "").strip().lower()
    return str(charge_or_text or "").strip().lower()


def one_fee_code(charge):
    if isinstance(charge, dict):
        return str(charge.get("chargeCode") or "").strip().upper()
    return ""


def one_is_discount_charge(fee_name, charge=None):
    fee = one_fee_name(fee_name)
    if any(k in fee for k in ONE_DISCOUNT_KEYWORDS):
        return True
    if isinstance(charge, dict):
        code = str(charge.get("chargeCode") or "").strip().upper()
        if code in ONE_DISCOUNT_CODES:
            return True
        for k in ("chargeType", "chargeCategory", "category", "rateType"):
            val = str(charge.get(k) or "").lower()
            if "discount" in val or "promotion" in val:
                return True
        if charge.get("isDiscount") is True or charge.get("discount") is True or charge.get("discountFlag") is True:
            return True
    return False


def one_is_origin_local_charge(fee_name, charge_code=""):
    fee = one_fee_name(fee_name)
    code = str(charge_code or "").strip().upper()
    if one_is_discount_charge(fee):
        return False
    if any(k in fee for k in ONE_ORIGIN_LOCAL_EXCLUDED_KEYWORDS):
        return True
    return bool(code and code in ONE_ORIGIN_LOCAL_EXCLUDED_CODES)


def one_is_origin_thc_charge(fee_name, charge_code=""):
    fee = one_fee_name(fee_name)
    code = str(charge_code or "").strip().upper()
    return (
        code in {"THC", "THCL", "OTHC"}
        or "terminal handling charge (l)" in fee
        or "terminal handling charge at origin" in fee
        or "terminal handling charge at port of loading" in fee
        or "thc/l" in fee
        or "thc at origin" in fee
    )


def one_is_ows_charge(fee_name, charge_code=""):
    """OWS/Heavy chỉ dùng để tạo remark, không được cộng vào giá."""
    fee = one_fee_name(fee_name)
    code = str(charge_code or "").strip().upper()
    return (
        code in {"OWS", "HWC", "HLC", "HEA"}
        or "heavy surcharge" in fee
        or "heavy weight" in fee
        or "heavy lift" in fee
        or "overweight" in fee
        or re.search(r"\bows\b", fee) is not None
    )


def one_should_include_charge(fee_name, group="", charge_code="", include_origin_thc=False):
    fee = one_fee_name(fee_name)
    group = str(group or "").strip()
    code = str(charge_code or "").strip().upper()
    if not fee:
        return False
    # OWS có thể nằm trong premiumCharges/freightCharges chứ không chỉ
    # originCharges. Luôn loại khỏi tổng, nhưng caller vẫn bật cờ remark.
    if one_is_ows_charge(fee, code):
        return False
    if group == "destinationCharges":
        return False
    if group == "promotionCharges" or one_is_discount_charge(fee):
        return True
    if group in {"basicOceanFreightCharges", "freightCharges", "premiumCharges"}:
        return True
    if group == "originCharges":
        if include_origin_thc and one_is_origin_thc_charge(fee, code):
            return True
        return not one_is_origin_local_charge(fee, code)
    if not group:
        if include_origin_thc and one_is_origin_thc_charge(fee, code):
            return True
        return not one_is_origin_local_charge(fee, code)
    return True


def one_api_charge_equipment_key(charge):
    iso = str(charge.get("equipmentIsoCode") or "").upper()
    if iso in {"22G1", "42G1", "45G1"}:
        return {"22G1": "DRY 20", "42G1": "DRY 40", "45G1": "DRY 40H"}[iso]
    text = " ".join(
        str(charge.get(k) or "")
        for k in (
            "equipmentDisplayName", "equipmentName", "equipmentType",
            "equipmentSize", "equipmentONECntrTpSz", "containerType",
            "cntrTpSz", "chargeUnit", "unit", "description"
        )
    ).upper()
    if any(x in text for x in ["45G1", "DRY 40H", "40H", "40 HC", "40HQ", "HIGH CUBE"]):
        return "DRY 40H"
    if any(x in text for x in ["42G1", "DRY 40", "40GP", "40 GP", "40'"]):
        return "DRY 40"
    if any(x in text for x in ["22G1", "DRY 20", "20GP", "20 GP", "20'"]):
        return "DRY 20"
    return ""


def one_api_add_charge(final_prices, formula_parts, charge, raw_amount_usd=None, group=""):
    key = one_api_charge_equipment_key(charge)
    raw_amount = float(raw_amount_usd if raw_amount_usd is not None else 0.0)
    if abs(raw_amount) < 1e-9:
        return

    fee_name = str(charge.get("chargeName") or "").strip().lower()
    is_discount = (group == "promotionCharges") or one_is_discount_charge(fee_name, charge)
    amount = -abs(raw_amount) if is_discount else raw_amount

    target_keys = [key] if key else list(final_prices.keys())
    for target_key in target_keys:
        final_prices[target_key] += amount
        formula_parts[target_key].append(amount)


def excel_formula_from_parts(parts):
    tokens = []
    for raw in parts:
        try:
            amount = float(raw)
        except (TypeError, ValueError):
            continue
        if abs(amount) < 1e-9:
            continue
        sign = "-" if amount < 0 else "+"
        amount = abs(amount)
        if abs(amount - round(amount)) < 1e-9:
            text = str(int(round(amount)))
        else:
            text = f"{amount:.2f}".rstrip("0").rstrip(".")
        tokens.append((sign if tokens else ("-" if sign == "-" else "")) + text)
    return "=" + "".join(tokens) if tokens else None
