"""Compares live InternationalFees/get-by-code data against
international_fee_expected.csv. No phases (confirmed). Live fields ARE
fee-after-scholarship directly -- NO subtraction needed (unlike domestic
CUET data), confirmed by exact match against real P132 data."""
import csv

CAT_TO_LIVE_FIELD = {"A": "feeAbove90", "B": "feeAbove75", "C": "feeAbove60"}
ZONE_TO_LIVE_KEY = {"saarc": "saarc", "zone1": "nonSAARC", "zone2": "other"}


def values_equal(a, b):
    """Compare numeric values without treating 30000 and 30000.0 as different."""
    if a in ("", None) and b in ("", None):
        return True

    try:
        return float(a) == float(b)
    except (ValueError, TypeError):
        return str(a).strip() == str(b).strip()


def load_expected(path="lpu_monitor/config/international_fee_expected.csv"):
    lookup = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            lookup[row["official_code"]] = row
    return lookup


def check_international_fee(official_code, live_json, expected_lookup):
    expected = expected_lookup.get(official_code)
    if expected is None:
        return {"code": official_code, "status": "NO_EXPECTED_DATA"}

    mismatches = []
    for zone_key, live_key in ZONE_TO_LIVE_KEY.items():
        live_zone = live_json.get(live_key, {}) or {}
        expected_base = expected.get(f"base_{zone_key}", "")
        live_base = live_zone.get("tuitionFee", "")

        if (
            expected_base not in ("", None)
            and not values_equal(expected_base, live_base)
        ):
            mismatches.append({
                "field": f"base_{zone_key}",
                "expected": expected_base,
                "live": live_base
            })

        for cat, live_field in CAT_TO_LIVE_FIELD.items():
            expected_val = expected.get(f"{zone_key}_{cat}", "")
            live_val = live_zone.get(live_field, "")

            if expected_val in ("", None):
                continue  # no ground truth for this cell -- skip, not a mismatch

            if not values_equal(expected_val, live_val):
                mismatches.append({
                    "field": f"{zone_key}_{cat}",
                    "expected": expected_val,
                    "live": live_val,
                })

    return {
        "code": official_code,
        "status": "CHECKED",
        "mismatches": mismatches
    }