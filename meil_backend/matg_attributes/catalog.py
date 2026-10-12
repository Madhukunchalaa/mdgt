"""
CMT Code / Catalog Number / Sequence rules, all in ONE place.

Business rules (change them here if the client answers differently):

  * Every VALUE of a characteristic has a 2-digit numeric "CMT Code".
  * A material's Catalog Number = the CMT codes of its selected values,
    joined in characteristic SEQUENCE order (e.g. 01 01 01 00 01 01 01).
  * Sequence must be a positive whole number, unique per material group.
"""
import logging
import re

logger = logging.getLogger(__name__)

# ---- Assumptions pending client confirmation ------------------------------
CMT_CODE_PATTERN = re.compile(r"[0-9]{2}")        # ASCII 0-9 only (\d would accept other scripts)
NO_VALUE_CODE = "00"                      # filler for a characteristic that has no value selected
REQUIRE_UNIQUE_CODES_PER_CHARACTERISTIC = True
# Characteristics never shown in descriptions are also left out of the catalog number
HIDDEN_FROM_DESCRIPTION = ["bu", "eil code", "toyo code", "invesca code"]
# ---------------------------------------------------------------------------


def _norm(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def _key(name):
    return _norm(name).replace(" ", "")


def _raw_value(value):
    """Item attributes store either 'text' or {'value': 'text', 'uom': 'mm'}."""
    if isinstance(value, dict):
        value = value.get("value", "")
    return "" if value is None else str(value).strip()


# ---------------------------------------------------------------------------
# CMT code validation
# ---------------------------------------------------------------------------
def clean_cmt_code(raw, pad_single_digit=False):
    """Return a valid 2-digit code string, or raise ValueError."""
    if raw is None or isinstance(raw, bool):
        raise ValueError("is required")
    text = str(raw).strip()
    if pad_single_digit and len(text) == 1 and text in "0123456789":
        text = text.zfill(2)          # Excel turns "01" into 1
    if not CMT_CODE_PATTERN.fullmatch(text):
        raise ValueError("must be exactly 2 digits (0-9), e.g. 01")
    return text


def validate_cmt_codes(possible_values, cmt_codes, pad_single_digit=False):
    """
    Check that EVERY possible value has a valid 2-digit code.
    Returns (clean_codes_dict, errors_list). Codes for values that are not in
    possible_values are dropped.
    """
    errors = []
    if cmt_codes is None:
        cmt_codes = {}
    if not isinstance(cmt_codes, dict):
        return {}, ["cmt_codes must be an object like {\"FLANGE\": \"01\"}"]

    lookup = {_norm(k): v for k, v in cmt_codes.items()}
    clean = {}
    seen = {}
    for value in possible_values or []:
        raw = lookup.get(_norm(value))
        try:
            code = clean_cmt_code(raw, pad_single_digit)
        except ValueError as exc:
            errors.append(f"CMT Code for '{value}' {exc}")
            continue
        if REQUIRE_UNIQUE_CODES_PER_CHARACTERISTIC and code in seen:
            errors.append(f"CMT Code {code} is used for both '{seen[code]}' and '{value}'")
            continue
        seen[code] = value
        clean[value] = code
    return clean, errors


# ---------------------------------------------------------------------------
# Sequence rules
# ---------------------------------------------------------------------------
def parse_sequence(raw):
    """Positive whole number or ValueError."""
    if raw is None or isinstance(raw, bool) or str(raw).strip() == "":
        raise ValueError("Sequence is required")
    try:
        number = float(str(raw).strip())
    except ValueError:
        raise ValueError("Sequence must be a whole number")
    if number != int(number) or int(number) <= 0:
        raise ValueError("Sequence must be a positive whole number (1, 2, 3 ...)")
    return int(number)


def find_sequence_conflict(mgrp, sequence, exclude_id=None):
    """The ACTIVE characteristic already using this sequence in the group, else None."""
    from .models import MatgAttributeItem
    qs = MatgAttributeItem.objects.filter(mgrp_code=mgrp, print_priority=sequence, is_deleted=False)
    if exclude_id is not None:
        qs = qs.exclude(id=exclude_id)
    return qs.first()


def duplicate_sequences(mgrp=None):
    """Existing bad data: [{'mgrp_code', 'sequence', 'attributes': [...]}, ...]"""
    from .models import MatgAttributeItem
    qs = MatgAttributeItem.objects.filter(is_deleted=False, print_priority__isnull=False)
    if mgrp is not None:
        qs = qs.filter(mgrp_code=mgrp)
    groups = {}
    for a in qs.order_by("mgrp_code_id", "print_priority", "id"):
        groups.setdefault((a.mgrp_code_id, a.print_priority), []).append(a.attribute_name)
    return [
        {"mgrp_code": g, "sequence": s, "attributes": names}
        for (g, s), names in groups.items() if len(names) > 1
    ]


# ---------------------------------------------------------------------------
# Ordering + Catalog Number
# ---------------------------------------------------------------------------
def ordered_characteristics(mgrp):
    from .models import MatgAttributeItem
    chars = MatgAttributeItem.objects.filter(mgrp_code=mgrp, is_deleted=False)
    return sorted(chars, key=lambda a: (a.print_priority is None, a.print_priority or 0, a.id))


def order_attributes(mgrp, attributes):
    """Return item attributes as a dict ordered by characteristic sequence."""
    if not isinstance(attributes, dict) or mgrp is None:
        return attributes
    by_key = {_key(k): k for k in attributes}
    ordered = {}
    for char in ordered_characteristics(mgrp):
        original = by_key.pop(_key(char.attribute_name), None)
        if original is not None:
            ordered[original] = attributes[original]
    for original in by_key.values():          # BU etc. (not in the table) go last
        ordered[original] = attributes[original]
    return ordered


def _code_for_value(char, value):
    codes = {_norm(k): v for k, v in (char.cmt_codes or {}).items()}
    candidates = [value]
    for uom in (char.uom or "").split(","):          # "2 inch" -> "2"
        uom = uom.strip()
        if uom and _norm(value).endswith(" " + _norm(uom)):
            candidates.append(value[: -len(uom)].strip())
    for cand in candidates:
        if _norm(cand) in codes:
            return codes[_norm(cand)]
    return None


def catalog_breakdown(mgrp, attributes):
    """
    Explain how the catalog number is built.
    Returns {'catalog_number': str|None, 'complete': bool, 'parts': [...], 'warnings': [...]}
    """
    parts, warnings = [], []
    if mgrp is None:
        return {"catalog_number": None, "complete": False, "parts": parts, "warnings": warnings}
    attributes = attributes if isinstance(attributes, dict) else {}
    by_key = {_key(k): v for k, v in attributes.items()}

    for char in ordered_characteristics(mgrp):
        if _norm(char.attribute_name) in HIDDEN_FROM_DESCRIPTION:
            continue
        value = _raw_value(by_key.get(_key(char.attribute_name)))
        if not value:
            code, status = NO_VALUE_CODE, "no value"
        else:
            found = _code_for_value(char, value)
            if found:
                code, status = found, "ok"
            else:
                code, status = NO_VALUE_CODE, "no code"
                warnings.append(f"'{value}' has no CMT Code in {char.attribute_name}")
        parts.append({
            "sequence": char.print_priority, "characteristic": char.attribute_name,
            "value": value, "code": code, "status": status,
        })
    # A catalog number is only produced when it is trustworthy:
    #   - a characteristic with NO value selected  -> filler code 00 (as in the client's sample)
    #   - a selected value that has NO CMT code    -> incomplete: no number, warning instead
    #   - nothing selected at all                  -> no number
    incomplete = any(p["status"] == "no code" for p in parts)
    nothing_selected = all(p["status"] == "no value" for p in parts)
    number = None if (incomplete or nothing_selected) else "".join(p["code"] for p in parts)
    return {"catalog_number": number, "complete": number is not None,
            "parts": parts, "warnings": warnings}


def build_catalog_number(mgrp, attributes):
    return catalog_breakdown(mgrp, attributes)["catalog_number"]


def refresh_group_catalog_numbers(mgrp):
    """Recalculate Catalog Number for every active item of a material group."""
    from itemmaster.models import ItemMaster
    count = 0
    for item in ItemMaster.objects.filter(mgrp_code=mgrp, is_deleted=False).iterator():
        number = build_catalog_number(mgrp, item.attributes)
        if number != item.catalog_number:
            ItemMaster.objects.filter(pk=item.pk).update(catalog_number=number)
            count += 1
    return count
