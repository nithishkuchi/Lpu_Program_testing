"""
map_programmes.py

Maps every valid non-PhD programme in the "Final Programme List" workbook to its
correct schedule bucket in the "Important Dates" workbook, and writes an output
workbook with the matched dates attached.

Nothing about programme names, programme count, or type letters (A/B/C/D...) is
hardcoded. The script discovers:
  - which sheet is the "Programme list" sheet (by name match)
  - which columns are "Type of Programme", "Programme Name", "LPU NEST Test Code"
    (by header text match, case/whitespace-insensitive)
  - which sheets in the Important Dates workbook are ignorable (PET, TABS)
  - which sheets represent which bucket (Type C, Type B generic, Type B MBA,
    Type B UG Design, Type B BTECH interview, Clinical Psychology, and any future
    "Type <letter>" sheet)

===========================================================================
CONFIG YOU WILL LIKELY EDIT LATER  (search for "CONFIG:" to find these fast)
===========================================================================
"""

import re
import sys
import argparse
from collections import OrderedDict, defaultdict

import xlrd
import openpyxl


# ---------------------------------------------------------------------------
# CONFIG: programmes to skip outright (exceptions), matched as a case-insensitive
# substring of the Programme Name. Add / remove entries here as needed.
# ---------------------------------------------------------------------------
SKIP_PROGRAMME_NAME_SUBSTRINGS = [
    "software product engineering",
]

# ---------------------------------------------------------------------------
# CONFIG: type letters that should be treated as a DIFFERENT letter before any
# other logic runs. Use this for known mislabeling in the source file.
# Format: {"a": "b"} means "any programme typed as Type A is treated as Type B".
# Set to {} once the source file is corrected.
# ---------------------------------------------------------------------------
TYPE_LETTER_OVERRIDE = {
    "a": "b",
}

# ---------------------------------------------------------------------------
# CONFIG: Important Dates sheet names to ignore completely (case-insensitive
# substring match against the sheet name).
# ---------------------------------------------------------------------------
IGNORE_SHEET_SUBSTRINGS = ["pet", "tabs"]

# ---------------------------------------------------------------------------
# CONFIG: Final Programme List - which sheet holds the working programme list.
# ---------------------------------------------------------------------------
PROGRAMME_LIST_SHEET_SUBSTRING = "programme list"  # falls back to "program list"

# Header text we look for (case/whitespace-insensitive substring match) to find
# the right columns, regardless of exact column position.
HEADER_TYPE_OF_PROGRAMME = "type of programme"
HEADER_PROGRAMME_NAME = "programme name"
HEADER_NEST_TEST_CODE = "lpu nest test code"
HEADER_PROGRAMME_CODE = "programme code"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def norm(s):
    """Normalize a string: collapse whitespace/newlines, strip, lowercase."""
    if s is None:
        return ""
    s = str(s).replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def norm_keep_case(s):
    if s is None:
        return ""
    s = str(s).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------------------
# Step 1: Load Important Dates workbook -> buckets
# ---------------------------------------------------------------------------
def is_phase_label_row(row_values):
    """A row like [None, 'Phase 5'] or [None, 'Schedule I', 'Schedule II'] that
    labels columns rather than holding a data particular."""
    first = row_values[0]
    if first not in (None, ""):
        return False
    rest = [v for v in row_values[1:] if v not in (None, "")]
    if not rest:
        return False
    return all(re.search(r"phase|schedule", str(v), re.I) for v in rest)


def load_bucket_from_sheet(ws):
    """Parse a schedule sheet into an ordered list of (particular, value) using
    the right-most data column (treated as the 'current' phase)."""
    rows = list(ws.iter_rows(values_only=True))
    header_row_idx = None
    for i, row in enumerate(rows):
        if row and norm(row[0]) in ("particulars", "activity"):
            header_row_idx = i
            break
    if header_row_idx is None:
        # Fallback: assume first row is header
        header_row_idx = 0

    data_start = header_row_idx + 1
    if data_start < len(rows) and is_phase_label_row(rows[data_start]):
        data_start += 1

    # right-most column actually used in this sheet
    last_col = ws.max_column - 1  # 0-indexed

    particulars = []
    for row in rows[data_start:]:
        if not row:
            continue
        label = row[0]
        if label in (None, ""):
            continue
        value = row[last_col] if last_col < len(row) else None
        particulars.append((norm_keep_case(label), value))
    return particulars


FILLER_WORDS = {
    "all", "tests", "test", "programme", "programmes", "program", "programs",
    "the", "for", "and", "only", "offered",
}


def extract_bucket_keywords(sheet_name, letter):
    """Pulls the meaningful, distinguishing words out of a Type-lettered sheet's
    own name - e.g. 'Type B (MBA)' -> {'mba'}, 'TypeB (UG Design)' -> {'ug','design'},
    'Type B (BTECH_INTERVIEW)' -> {'btech','interview'}, 'Type B (all tests)' -> {}
    (empty set = this is the generic/base bucket for that letter).
    Nothing here is a hardcoded subject name - it's derived from whatever sheet
    names actually exist in the Important Dates workbook, so a brand-new sheet
    like 'Type A (Something New)' or 'Type C (Something Else)' is picked up
    automatically without touching this function.
    """
    n = norm(sheet_name)
    n = re.sub(r"type\s*" + re.escape(letter) + r"\b", " ", n)
    n = re.sub(r"[()_\-.]", " ", n)
    words = [w for w in n.split() if w not in FILLER_WORDS and len(w) > 1]
    return set(words)


def load_important_dates(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    buckets = {}  # bucket_key -> {"sheet","rows","letter","keywords","is_generic"}

    for sheet_name in wb.sheetnames:
        n = norm(sheet_name)
        if any(kw in n for kw in IGNORE_SHEET_SUBSTRINGS):
            continue

        ws = wb[sheet_name]

        if re.match(r"^ph\.?\s*d\.?$", n):
            buckets["phd"] = {"sheet": sheet_name, "rows": load_bucket_from_sheet(ws)}
            continue

        if "clinical" in n and "psych" in n:
            # Not a "Type X" pattern sheet - explicitly named for one specific
            # programme identity (Bachelor's Clinical Psychology). Kept as its
            # own special case since it isn't part of the letter/keyword system.
            buckets["clinical_psychology"] = {
                "sheet": sheet_name,
                "rows": load_bucket_from_sheet(ws),
            }
            continue

        m = re.search(r"type\s*([a-z])\b", n)
        if not m:
            continue  # not a "Type X..." sheet and not clinical psychology -> unused
        letter = m.group(1)
        keywords = extract_bucket_keywords(sheet_name, letter)
        is_generic = len(keywords) == 0
        suffix = "generic" if is_generic else "_".join(sorted(keywords))
        bucket_key = f"type_{letter}__{suffix}"
        buckets[bucket_key] = {
            "sheet": sheet_name,
            "rows": load_bucket_from_sheet(ws),
            "letter": letter,
            "keywords": keywords,
            "is_generic": is_generic,
        }

    return buckets


def find_generic_bucket(buckets, letter):
    for key, b in buckets.items():
        if b.get("letter") == letter and b.get("is_generic"):
            return key
    return None


def find_specialty_buckets(buckets, letter):
    return [
        (key, b) for key, b in buckets.items()
        if b.get("letter") == letter and not b.get("is_generic")
    ]


# ---------------------------------------------------------------------------
# Step 2: Load Final Programme List -> rows + dynamic column detection
# ---------------------------------------------------------------------------
def find_programme_list_sheet(wb):
    for name in wb.sheet_names():
        n = norm(name)
        if PROGRAMME_LIST_SHEET_SUBSTRING in n or "program list" in n:
            return name
    raise ValueError("Could not find the Programme List sheet by name.")


def find_header_columns(sheet):
    header_row = [sheet.cell_value(0, c) for c in range(sheet.ncols)]
    header_map = {}
    for c, h in enumerate(header_row):
        header_map[c] = norm(h)

    def find_col(*keywords):
        for c, h in header_map.items():
            if all(kw in h for kw in keywords):
                return c
        return None

    col_type = find_col(HEADER_TYPE_OF_PROGRAMME)
    col_name = find_col(HEADER_PROGRAMME_NAME)
    col_nest = find_col(HEADER_NEST_TEST_CODE)
    col_code = find_col(HEADER_PROGRAMME_CODE)

    if col_type is None or col_name is None:
        raise ValueError(
            f"Could not locate required columns. Found type_col={col_type}, "
            f"name_col={col_name}, nest_col={col_nest}. Check header text config."
        )
    return header_row, col_type, col_name, col_nest, col_code


def find_phd_sheet(wb):
    for name in wb.sheet_names():
        n = norm(name)
        if re.search(r"ph\.?\s*d\b", n) and "abbrev" not in n:
            return name
    return None


def find_name_and_code_columns(sheet):
    header_row = [sheet.cell_value(0, c) for c in range(sheet.ncols)]
    col_name, col_code = None, None
    for c, h in enumerate(header_row):
        hn = norm(h)
        if col_name is None and HEADER_PROGRAMME_NAME in hn:
            col_name = c
        if col_code is None and HEADER_PROGRAMME_CODE in hn:
            col_code = c
    if col_name is None:
        raise ValueError("Could not find 'Programme Name' column in PhD sheet.")
    return header_row, col_name, col_code


def load_phd_list(path):
    """Loads the PhD sheet from the Final Programme List workbook, if present."""
    wb = xlrd.open_workbook(path)
    sheet_name = find_phd_sheet(wb)
    if sheet_name is None:
        return None
    sh = wb.sheet_by_name(sheet_name)
    header_row, col_name, col_code = find_name_and_code_columns(sh)
    rows = []
    for r in range(1, sh.nrows):
        row_values = [sh.cell_value(r, c) for c in range(sh.ncols)]
        rows.append(row_values)
    return {
        "sheet_name": sheet_name,
        "header": header_row,
        "col_name": col_name,
        "col_code": col_code,
        "rows": rows,
    }


def load_final_list(path):
    wb = xlrd.open_workbook(path)
    sheet_name = find_programme_list_sheet(wb)
    sh = wb.sheet_by_name(sheet_name)
    header_row, col_type, col_name, col_nest, col_code = find_header_columns(sh)

    rows = []
    for r in range(1, sh.nrows):
        row_values = [sh.cell_value(r, c) for c in range(sh.ncols)]
        rows.append(row_values)

    return {
        "sheet_name": sheet_name,
        "header": header_row,
        "col_type": col_type,
        "col_name": col_name,
        "col_nest": col_nest,
        "col_code": col_code,
        "rows": rows,
    }


# ---------------------------------------------------------------------------
# Step 3: Classification logic
# ---------------------------------------------------------------------------
def classify_row(programme_name, type_text, nest_code, buckets):
    """Returns (status, bucket_key_or_none, detail)
    status is one of: MAPPED, SKIP_NO_TYPE, SKIP_EXCEPTION, UNMAPPED_NO_SHEET,
                       SKIP_UNRECOGNIZED_TYPE

    Fully dynamic across type letters (A/B/C/D/...) and across whatever
    specialty sheets exist in the Important Dates workbook - no subject or
    process keyword (MBA, Design, Interview, etc.) is hardcoded here. They are
    all derived at load time from the Important Dates sheet names themselves
    (see extract_bucket_keywords). Adding a brand-new sheet there is enough to
    teach the script a new bucket; nothing in this function needs editing.
    """
    name_n = norm(programme_name)
    type_n = norm(type_text)
    nest_n = norm(nest_code)

    if type_n in ("", "-", "na", "n/a"):
        return "SKIP_NO_TYPE", None, type_text

    for bad in SKIP_PROGRAMME_NAME_SUBSTRINGS:
        if bad in name_n:
            return "SKIP_EXCEPTION", None, f"Matched skip-list entry: '{bad}'"

    # Explicitly-named, non-"Type X" sheet: Bachelor's Clinical Psychology.
    # Checked by programme identity, not by type letter, since the dates sheet
    # is written for exactly one specific degree level.
    if "clinical" in name_n and "psych" in name_n:
        is_bachelor = any(k in name_n for k in ["bachelor", "b.clin", "b clin", "bclin", "b.sc"])
        if is_bachelor:
            if "clinical_psychology" in buckets:
                return "MAPPED", "clinical_psychology", None
            return "UNMAPPED_NO_SHEET", None, "No Clinical Psychology sheet found"
        # not the Bachelor's -> fall through to normal letter-based matching below

    m = re.search(r"type\s*([a-z])\b", type_n)
    if not m:
        return "SKIP_UNRECOGNIZED_TYPE", None, type_text

    letter = m.group(1)
    letter = TYPE_LETTER_OVERRIDE.get(letter, letter)

    specialty = find_specialty_buckets(buckets, letter)

    # Step A: subject-style signals (e.g. MBA, UG Design) live in the NEST Test
    # Code - require ALL of a bucket's keywords to appear there together, so a
    # generic code like plain "B.Tech" never accidentally satisfies a
    # multi-keyword bucket such as {"btech","interview"}.
    nest_matches = [
        (key, len(b["keywords"])) for key, b in specialty
        if b["keywords"] and all(kw in nest_n for kw in b["keywords"])
    ]
    if nest_matches:
        nest_matches.sort(key=lambda x: -x[1])  # most specific (most keywords) wins
        return "MAPPED", nest_matches[0][0], None

    # Step B: process-style signals (e.g. "+ Interview") live in the Type of
    # Programme text itself - a single matching keyword there is decisive,
    # since that field is what was explicitly typed for this exact programme.
    type_matches = [
        (key, len(b["keywords"])) for key, b in specialty
        if b["keywords"] and any(kw in type_n for kw in b["keywords"])
    ]
    if type_matches:
        type_matches.sort(key=lambda x: -x[1])
        return "MAPPED", type_matches[0][0], None

    # Step C: fall back to the generic bucket for this letter, if one exists.
    generic_key = find_generic_bucket(buckets, letter)
    if generic_key:
        return "MAPPED", generic_key, None

    return "UNMAPPED_NO_SHEET", None, f"No 'Type {letter.upper()}' sheet found"


# ---------------------------------------------------------------------------
# Step 4: Build output workbook
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Step 4: Build output CSV(s)
#
# Output columns, exactly as requested:
#   Programme Name | Programme Code | Matched Bucket (Dates Sheet) | Dates
# "Dates" is a JSON dictionary: {side_heading: date_value, ...} taken directly
# from the matched Important Dates sheet's own particular labels - so the keys
# are whatever the Important Dates sheet calls them, nothing renamed/hardcoded.
# ---------------------------------------------------------------------------
import csv
import json


def build_date_dict(bucket):
    """side-heading -> date value, straight from the matched sheet."""
    d = OrderedDict()
    for particular, value in bucket["rows"]:
        if value is None:
            value = ""
        elif not isinstance(value, str):
            value = str(value)
        d[particular] = value
    return d


def classify_all(rows, col_type, col_name, col_nest, col_code, buckets):
    """Runs classify_row over every row, returns (mapped, review) lists."""
    mapped, review = [], []
    for row_values in rows:
        type_text = row_values[col_type] if col_type is not None else ""
        prog_name = row_values[col_name]
        nest_code = row_values[col_nest] if col_nest is not None else ""
        prog_code = row_values[col_code] if col_code is not None else ""

        status, bucket_key, detail = classify_row(prog_name, type_text, nest_code, buckets)

        if status == "MAPPED":
            bucket = buckets[bucket_key]
            mapped.append({
                "name": prog_name,
                "code": prog_code,
                "bucket_sheet": bucket["sheet"],
                "dates": build_date_dict(bucket),
            })
        elif status == "SKIP_NO_TYPE":
            continue  # blank/'-' rows, fully ignored
        else:
            review.append({
                "name": prog_name,
                "code": prog_code,
                "status": status,
                "detail": detail,
            })
    return mapped, review


def write_mapped_csv(path, mapped_rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Programme Name", "Programme Code", "Matched Bucket", "Dates"])
        for r in mapped_rows:
            writer.writerow([
                norm_keep_case(r["name"]),
                norm_keep_case(r["code"]),
                r["bucket_sheet"],
                json.dumps(r["dates"], ensure_ascii=False),
            ])


def write_review_csv(path, review_rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Programme Name", "Programme Code", "Issue", "Detail"])
        for r in review_rows:
            writer.writerow([
                norm_keep_case(r["name"]),
                norm_keep_case(r["code"]),
                r["status"],
                r["detail"],
            ])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--final-list", required=True, help="Path to Final Programme List (.xls/.xlsx)")
    ap.add_argument("--important-dates", required=True, help="Path to Important Dates workbook (.xlsx)")
    ap.add_argument("--out", required=True, help="Output .csv path (main mapped output)")
    ap.add_argument("--review-out", default=None,
                     help="Optional .csv path for unmapped/flagged rows (defaults next to --out)")
    args = ap.parse_args()

    review_out = args.review_out or re.sub(r"\.csv$", "_needs_review.csv", args.out)

    buckets = load_important_dates(args.important_dates)
    final_data = load_final_list(args.final_list)
    phd_data = load_phd_list(args.final_list)

    mapped, review = classify_all(
        final_data["rows"], final_data["col_type"], final_data["col_name"],
        final_data["col_nest"], final_data["col_code"], buckets,
    )

    phd_mapped, phd_review = [], []
    if phd_data is not None:
        for row_values in phd_data["rows"]:
            name = row_values[phd_data["col_name"]]
            code = row_values[phd_data["col_code"]] if phd_data["col_code"] is not None else ""
            if name in (None, ""):
                continue
            if "phd" in buckets:
                phd_mapped.append({
                    "name": name,
                    "code": code,
                    "bucket_sheet": buckets["phd"]["sheet"],
                    "dates": build_date_dict(buckets["phd"]),
                })
            else:
                phd_review.append({
                    "name": name, "code": code,
                    "status": "UNMAPPED_NO_SHEET",
                    "detail": "No 'Ph.D.' sheet found in Important Dates workbook",
                })

    all_mapped = mapped + phd_mapped
    all_review = review + phd_review

    write_mapped_csv(args.out, all_mapped)
    write_review_csv(review_out, all_review)

    print("Buckets discovered:")
    for k, b in buckets.items():
        print(f"  {k:25s} -> sheet '{b['sheet']}'  ({len(b['rows'])} particulars)")
    print()
    print(f"Non-PhD: mapped={len(mapped)}  needs_review={len(review)}")
    print(f"PhD:     mapped={len(phd_mapped)}  needs_review={len(phd_review)}")
    print(f"Total:   mapped={len(all_mapped)}  needs_review={len(all_review)}")
    print()
    print("Wrote:", args.out)
    print("Wrote:", review_out)


if __name__ == "__main__":
    main()