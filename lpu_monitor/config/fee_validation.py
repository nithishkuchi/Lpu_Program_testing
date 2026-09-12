"""
fee_validation.py

Compares live ScholarshipCUET data (per programme) against fee_expected.csv
(the ground-truth join built by build_fee_expected.py).

Category-to-live-field-prefix mapping (confirmed from the Excel header text:
"Category A for 60% Scholarship... or 98% marks or above"):
    A -> 98,  B -> 90,  C -> 80,  D -> 70

Phase handling is fully DYNAMIC, matching fee_excel_parser.py and
build_fee_expected.py's dynamic phase-count detection: this file does not
hardcode "phases 1-4 are fine, phase 5 is unsupported." Instead, it checks
whether fee_expected.csv actually has data for whatever phase is live -- so
if a future Excel adds phase 5 (or 6), this works automatically with no code
change. If the live site shows a phase the Excel doesn't have data for yet,
that's flagged generically.
"""
import csv
import re

CATEGORY_TO_PREFIX = {"A": "98", "B": "90", "C": "80", "D": "70"}


def load_fee_expected(path="lpu_monitor/config/fee_expected.csv"):
    lookup = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            lookup[row["official_code"]] = row
    return lookup


def unwrap_cuet_response(cuet_json):
    """
    ScholarshipCUET returns a JSON array containing one object, e.g. [{...}],
    not a plain object -- confirmed from a real 502-retry-then-success capture.
    Unwrap it here so the rest of this file can just work with a plain dict.
    """
    if isinstance(cuet_json, list):
        return cuet_json[0] if cuet_json else {}
    return cuet_json


def determine_active_phase(cuet_json):
    """
    Returns the active phase number, or None if undetermined. Scans ALL
    isPhaseN keys present in the live response dynamically -- does not
    assume phases stop at 4, or that 5 is special. Prefers openFeeTab (the
    default-shown tab) when multiple isPhaseN are true.
    """
    cuet_json = unwrap_cuet_response(cuet_json)
    phase_flags = {}
    for key, value in cuet_json.items():
        m = re.fullmatch(r"isPhase(\d+)", key)
        if m:
            phase_flags[int(m.group(1))] = value

    open_tab = cuet_json.get("openFeeTab")
    if open_tab and phase_flags.get(open_tab):
        return open_tab

    for phase in sorted(phase_flags):
        if phase_flags[phase]:
            return phase
    return None


def check_fee(official_code, cuet_json, fee_expected_lookup):
    cuet_json = unwrap_cuet_response(cuet_json)
    expected_row = fee_expected_lookup.get(official_code)
    if expected_row is None:
        return {"code": official_code, "status": "NO_EXPECTED_DATA"}

    active_phase = determine_active_phase(cuet_json)
    if active_phase is None:
        return {"code": official_code, "status": "NO_ACTIVE_PHASE_FOUND"}

    # Generic check: does fee_expected.csv actually have data for this phase?
    # Works for whatever phase count the current Excel covers -- no hardcoded
    # "only phase 5 is special" assumption.
    sample_field = f"phase{active_phase}_A"
    if sample_field not in expected_row:
        return {
            "code": official_code,
            "status": f"PHASE_{active_phase}_NOT_COVERED_BY_EXCEL",
            "active_phase": active_phase,
        }

    base_fee = float(expected_row["base_fee"])
    mismatches = []

    # Pull all 4 category raw live values for this phase first
    live_values = {}
    for category, prefix in CATEGORY_TO_PREFIX.items():
        live_values[category] = cuet_json.get(f"p{prefix}{active_phase}")

    non_null = [v for v in live_values.values() if v is not None]
    # If every category has the same raw value, this programme has no real
    # merit-scholarship differentiation (e.g. LPU RISE-only programmes) --
    # the raw value IS the final fee, not a discount to subtract.
    flat_no_scholarship = len(set(non_null)) <= 1 if non_null else True

    for category, prefix in CATEGORY_TO_PREFIX.items():
        live_field = f"p{prefix}{active_phase}"
        live_value = live_values[category]
        if live_value is None:
            mismatches.append({"category": category, "phase": active_phase,
                                "issue": f"live field {live_field} missing entirely"})
            continue

        live_fee_after = live_value if flat_no_scholarship else (base_fee - live_value)

        expected_field = f"phase{active_phase}_{category}"
        expected_raw = expected_row.get(expected_field, "")
        expected_value = float(expected_raw) if expected_raw not in ("", None) else base_fee

        if expected_value != live_fee_after:
            mismatches.append({
                "category": category, "phase": active_phase,
                "expected_fee_after_scholarship": expected_value,
                "live_computed_fee_after_scholarship": live_fee_after,
                "live_raw_value": live_value,
                "mode": "flat" if flat_no_scholarship else "scholarship",
            })

    return {"code": official_code, "status": "CHECKED", "active_phase": active_phase, "mismatches": mismatches}


# ---------------------------------------------------------------------------
# PhD-specific validation (separate sheet, no phases, only A/B/C categories,
# split by Full Time / Part Time -- see phd_fee_parser.py)
# ---------------------------------------------------------------------------

def load_phd_fee_expected(path="lpu_monitor/config/phd_fee_master.csv"):
    with open(path, encoding="utf-8") as f:
        row = next(csv.DictReader(f))  # single row file
    return {k: float(v) for k, v in row.items()}


PHD_CATEGORY_TO_PREFIX = {"A": "98", "B": "90", "C": "80"}
PHD_MODE_SUFFIX = {"Full Time": "ft", "Part Time": "pt"}


def normalize_mode(mode):
    """
    phd_codes.csv has inconsistent mode text across rows -- some say
    'Full Time', others 'Full Time Mode', 'full time', or even just 'Part'.
    Normalize by substring match instead of requiring an exact string, so
    every variant resolves correctly without hand-editing the source CSV.
    """
    if not mode:
        return None
    mode_lower = mode.strip().lower()
    if "full" in mode_lower:
        return "Full Time"
    if "part" in mode_lower:
        return "Part Time"
    return None


def check_fee_phd(official_code, mode, cuet_json, phd_fee_expected):
    cuet_json = unwrap_cuet_response(cuet_json)
    base_fee = phd_fee_expected["base_fee"]

    normalized_mode = normalize_mode(mode)
    mode_suffix = PHD_MODE_SUFFIX.get(normalized_mode)
    if mode_suffix is None:
        return {"code": official_code, "status": "UNKNOWN_MODE", "mode": mode}

    mismatches = []
    for category, prefix in PHD_CATEGORY_TO_PREFIX.items():
        live_field = f"p{prefix}1"  # PhD is always phase 1
        live_value = cuet_json.get(live_field)
        if live_value is None:
            mismatches.append({"category": category, "issue": f"live field {live_field} missing entirely"})
            continue

        live_fee_after = base_fee - live_value

        expected_key = f"cat{category}_{mode_suffix}"
        expected_value = phd_fee_expected.get(expected_key)
        if expected_value is None:
            mismatches.append({"category": category, "issue": f"no expected value for {expected_key}"})
            continue

        if expected_value != live_fee_after:
            mismatches.append({
                "category": category,
                "expected_fee_after_scholarship": expected_value,
                "live_computed_fee_after_scholarship": live_fee_after,
                "live_raw_scholarship_amount": live_value,
            })

    return {"code": official_code, "mode": normalized_mode, "status": "CHECKED", "mismatches": mismatches}


if __name__ == "__main__":
    import json
    fee_expected = load_fee_expected()
    with open("lpu_monitor/config/live_capture_cuet_P102.json") as f:
        p102_cuet = json.load(f)
    result = check_fee("P102", p102_cuet, fee_expected)
    print(result)