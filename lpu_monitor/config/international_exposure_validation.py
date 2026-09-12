"""
international_exposure_validation.py

Compares live GetProgramFeeDetail data (already fetched for fee validation
-- same endpoint, no new call needed) against the two confirmed ground-truth
columns in the Final Programme List Excel:
    "Availability of International Exposure (Summer/Winter School Option)"
    "Availability of Semester Abroad Option"

Extraction is tag-agnostic by design: strips ALL HTML first, then checks
for the plain-text titles directly. Confirmed bug in an earlier version:
regex requiring a specific tag structure (<strong><em>...) missed
programmes using different markup (<strong> alone) for the same content --
tag-agnostic substring matching avoids that failure mode entirely.

Also handles duplicate array entries in GetProgramFeeDetail (2-3 identical
copies, confirmed since early in this project) automatically -- since this
only checks TEXT PRESENCE (not a specific numeric value), duplicates don't
need special coalescing the way fee amounts do; searching the whole
stringified response is sufficient and safe.
"""
import csv
import re
import xlrd

INTL_EXPOSURE_TITLE = "International Exposure Option"
SEMESTER_EXCHANGE_TITLE = "Semester Exchange"


def strip_html(text):
    return re.sub(r"<[^>]+>", "", text)


def parse_yes_no(raw_value):
    """Handles qualified values like 'Yes (Only Semester 1/Freshmen students
    eligible)' or 'Yes (Equivalent to General B.Tech EEE/ECE)' -- confirmed
    real values in the source file, not just plain Yes/No. Treats any value
    starting with 'Yes' as True; keeps the full qualifier text separately
    so it's not silently discarded."""
    raw = str(raw_value).strip()
    is_yes = raw.lower().startswith("yes")
    qualifier = raw[3:].strip(" ()") if is_yes and len(raw) > 3 else ""
    return is_yes, qualifier


def load_expected_exposure(xls_path):
    """OfficialCode -> {intl_exposure, intl_exposure_qualifier,
    semester_abroad, semester_abroad_qualifier}"""
    wb = xlrd.open_workbook(xls_path)
    sheet = wb.sheet_by_index(0)
    header = sheet.row_values(0)

    code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
    intl_col = next(i for i, h in enumerate(header) if "international exposure" in h.lower())
    exch_col = next(i for i, h in enumerate(header) if "semester abroad" in h.lower())

    lookup = {}
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        code = row[code_col]
        if not code:
            continue
        intl_yes, intl_qual = parse_yes_no(row[intl_col])
        exch_yes, exch_qual = parse_yes_no(row[exch_col])
        lookup[code] = {
            "intl_exposure": intl_yes,
            "intl_exposure_qualifier": intl_qual,
            "semester_abroad": exch_yes,
            "semester_abroad_qualifier": exch_qual,
        }
    return lookup


def check_international_exposure(official_code, live_fee_json, expected_lookup):
    expected = expected_lookup.get(official_code)
    if expected is None:
        return {"code": official_code, "status": "NO_EXPECTED_DATA"}

    import json
    text = strip_html(json.dumps(live_fee_json))
    live_intl = INTL_EXPOSURE_TITLE in text
    live_exch = SEMESTER_EXCHANGE_TITLE in text

    mismatches = []
    if expected["intl_exposure"] != live_intl:
        mismatches.append({
            "field": "International Exposure Option",
            "expected": expected["intl_exposure"],
            "live": live_intl,
        })
    if expected["semester_abroad"] != live_exch:
        mismatches.append({
            "field": "Semester Exchange",
            "expected": expected["semester_abroad"],
            "live": live_exch,
        })

    return {"code": official_code, "status": "CHECKED", "mismatches": mismatches}