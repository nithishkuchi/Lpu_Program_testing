# """
# build_fee_expected.py

# Joins two already-finished files into the ONE file the comparison step actually
# needs:

#   fee_row_mapping.csv  (OfficialCode -> EligibilityBase, FeeRowLabel)   [297 rows]
#   fee_master.csv        (EligibilityBase, FeeRowLabel -> phase amounts) [~236 rows]

#                               |
#                               v

#   fee_expected.csv      (OfficialCode -> phase amounts directly)       [297 rows]

# This is a pure lookup/join -- no fuzzy matching, no ambiguity, since the
# mapping file is already fully resolved. If a code's (EligibilityBase,
# FeeRowLabel) pair isn't found in fee_master.csv, that's flagged loudly rather
# than silently skipped -- it means the two files are out of sync (e.g. a stray
# whitespace difference in FeeRowLabel), and needs fixing before trusting any
# downstream comparison.
# """
# import csv
# import re


# def normalize_label(text):
#     """
#     Handles real issues found in the source Excel and in how the mapping CSV
#     transcribes the same labels:
#     1. Genuine footnotes (e.g. "Note:- For Int. Credit Transfer...") --
#        strip from "note" onward, case-insensitive, wherever it appears.
#     2. Multi-name cells using a newline as a separator (e.g. "Zoology/ \n
#        M.Sc. (Zoology)") -- do NOT discard this, it's real content, not a
#        footnote. Flatten newlines to spaces instead of deleting anything.
#     3. Punctuation/spacing differences between how the master file and the
#        mapping CSV write the same label (slashes, double spaces, etc.) --
#        strip ALL non-alphanumeric characters rather than trying to guess a
#        consistent separator style, then lowercase and collapse whitespace.
#     """
#     text = re.split(r"note\s*:", text, flags=re.IGNORECASE)[0]
#     text = re.sub(r"[^a-z0-9]+", " ", text.lower())
#     return text.strip()


# # Same mismatch already solved once in fee_row_matcher.py: the Programme List
# # Excel's "Eligibility Base" column text doesn't match the Fee Excel's actual
# # sheet names. Normalize here too, so the join works regardless of which
# # wording the mapping CSV happens to use.
# ELIGIBILITY_BASE_TO_SHEET = {
#     "After 10th": "After 10th",
#     "After 10+2 (12th)": "After 10+2",
#     "Lateral Entry/Other Programmes(After ITI etc)": "Lateral Entry",
#     "Lateral Entry/Other Programmes(After 3 yrs. Diploma)": "Lateral Entry",
#     "After Graduation": "After Graduation",
# }


# def resolve_sheet_name(eligibility_base_raw):
#     key = eligibility_base_raw.strip()
#     for k, v in ELIGIBILITY_BASE_TO_SHEET.items():
#         if k.strip() == key:
#             return v
#     return key  # fall back to the raw value -- will simply not match, and get flagged as unmatched


# def load_fee_master(path="lpu_monitor/config/fee_master.csv"):
#     """Returns dict keyed by (eligibility_base, normalized_row_label) -> full row dict."""
#     lookup = {}
#     for_collision_check = {}  # key -> original row_label, to detect collisions
#     with open(path, encoding="utf-8") as f:
#         for row in csv.DictReader(f):
#             key = (row["eligibility_base"].strip(), normalize_label(row["row_label"]))
#             if key in for_collision_check and for_collision_check[key] != row["row_label"]:
#                 raise ValueError(
#                     f"COLLISION: two different row_labels normalized to the same key {key}:\n"
#                     f"  1: {for_collision_check[key]!r}\n"
#                     f"  2: {row['row_label']!r}\n"
#                     f"Normalization is too aggressive -- fix normalize_label() before trusting output."
#                 )
#             for_collision_check[key] = row["row_label"]
#             lookup[key] = row
#     return lookup


# def build_fee_expected(
#     mapping_path="lpu_monitor/config/fee_row_mapping.csv",
#     fee_master_path="lpu_monitor/config/fee_master.csv",
#     out_path="lpu_monitor/config/fee_expected.csv",
# ):
#     fee_master = load_fee_master(fee_master_path)
#     phase_cat_fields = [f"phase{p}_{c}" for p in [1, 2, 3, 4,5] for c in "ABCD"]

#     matched_rows = []
#     unmatched = []

#     with open(mapping_path, encoding="utf-8") as f:
#         for row in csv.DictReader(f):
#             code = row["official_code"].strip()
#             sheet_name = resolve_sheet_name(row["eligibility_base"])
#             key = (sheet_name, normalize_label(row["row_label"]))
#             master_row = fee_master.get(key)

#             if master_row is None:
#                 unmatched.append({"official_code": code, "eligibility_base": row["eligibility_base"], "row_label": row["row_label"].strip()})
#                 continue

#             out_row = {"official_code": code, "eligibility_base": sheet_name, "base_fee": master_row["base_fee"]}
#             for field in phase_cat_fields:
#                 out_row[field] = master_row.get(field, "")
#             matched_rows.append(out_row)

#     fieldnames = ["official_code", "eligibility_base", "base_fee"] + phase_cat_fields
#     with open(out_path, "w", newline="", encoding="utf-8") as f:
#         writer = csv.DictWriter(f, fieldnames=fieldnames)
#         writer.writeheader()
#         for row in matched_rows:
#             writer.writerow(row)

#     return len(matched_rows), unmatched


# if __name__ == "__main__":
#     matched_count, unmatched = build_fee_expected()
#     print(f"Matched {matched_count} codes -> lpu_monitor/config/fee_expected.csv")
#     if unmatched:
#         print(f"\n!!! {len(unmatched)} codes did NOT match fee_master.csv -- fix these before trusting the output:")
#         for u in unmatched:
#             print(" ", u)
#     else:
#         print("All codes matched cleanly. No unmatched rows.")
"""
build_fee_expected.py

Joins two already-finished files into the ONE file the comparison step actually
needs:

  fee_row_mapping.csv  (OfficialCode -> EligibilityBase, FeeRowLabel)   [297 rows]
  fee_master.csv        (EligibilityBase, FeeRowLabel -> phase amounts) [~236 rows]

                              |
                              v

  fee_expected.csv      (OfficialCode -> phase amounts directly)       [297 rows]

This is a pure lookup/join -- no fuzzy matching, no ambiguity, since the
mapping file is already fully resolved. If a code's (EligibilityBase,
FeeRowLabel) pair isn't found in fee_master.csv, that's flagged loudly rather
than silently skipped -- it means the two files are out of sync (e.g. a stray
whitespace difference in FeeRowLabel), and needs fixing before trusting any
downstream comparison.
"""
import csv
import re


def normalize_label(text):
    """
    Handles real issues found in the source Excel and in how the mapping CSV
    transcribes the same labels:
    1. Genuine footnotes (e.g. "Note:- For Int. Credit Transfer...") --
       strip from "note" onward, case-insensitive, wherever it appears.
    2. Multi-name cells using a newline as a separator (e.g. "Zoology/ \n
       M.Sc. (Zoology)") -- do NOT discard this, it's real content, not a
       footnote. Flatten newlines to spaces instead of deleting anything.
    3. Punctuation/spacing differences between how the master file and the
       mapping CSV write the same label (slashes, double spaces, etc.) --
       strip ALL non-alphanumeric characters rather than trying to guess a
       consistent separator style, then lowercase and collapse whitespace.
    """
    text = re.split(r"note\s*:", text, flags=re.IGNORECASE)[0]
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return text.strip()


# Same mismatch already solved once in fee_row_matcher.py: the Programme List
# Excel's "Eligibility Base" column text doesn't match the Fee Excel's actual
# sheet names. Normalize here too, so the join works regardless of which
# wording the mapping CSV happens to use.
ELIGIBILITY_BASE_TO_SHEET = {
    "After 10th": "After 10th",
    "After 10+2 (12th)": "After 10+2",
    "Lateral Entry/Other Programmes(After ITI etc)": "Lateral Entry",
    "Lateral Entry/Other Programmes(After 3 yrs. Diploma)": "Lateral Entry",
    "After Graduation": "After Graduation",
}


def resolve_sheet_name(eligibility_base_raw):
    key = eligibility_base_raw.strip()
    for k, v in ELIGIBILITY_BASE_TO_SHEET.items():
        if k.strip() == key:
            return v
    return key  # fall back to the raw value -- will simply not match, and get flagged as unmatched


def get_phase_cat_fields(fee_master_path):
    """
    Reads however many phase columns actually exist in fee_master.csv's
    header, instead of assuming a fixed count -- matches fee_excel_parser.py's
    dynamic phase detection, so this stays correct even if a 5th/6th phase
    appears in a future file without any code changes needed here.
    """
    with open(fee_master_path, encoding="utf-8") as f:
        header = next(csv.reader(f))
    phase_fields = [h for h in header if re.match(r"phase\d+_[A-D]$", h)]
    return sorted(phase_fields, key=lambda h: (int(re.search(r"\d+", h).group()), h[-1]))


def load_fee_master(path="lpu_monitor/config/fee_master.csv"):
    """Returns dict keyed by (eligibility_base, normalized_row_label) -> full row dict."""
    lookup = {}
    for_collision_check = {}  # key -> original row_label, to detect collisions
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            key = (row["eligibility_base"].strip(), normalize_label(row["row_label"]))
            if key in for_collision_check and for_collision_check[key] != row["row_label"]:
                raise ValueError(
                    f"COLLISION: two different row_labels normalized to the same key {key}:\n"
                    f"  1: {for_collision_check[key]!r}\n"
                    f"  2: {row['row_label']!r}\n"
                    f"Normalization is too aggressive -- fix normalize_label() before trusting output."
                )
            for_collision_check[key] = row["row_label"]
            lookup[key] = row
    return lookup


def build_fee_expected(
    mapping_path="lpu_monitor/config/fee_row_mapping.csv",
    fee_master_path="lpu_monitor/config/fee_master.csv",
    out_path="lpu_monitor/config/fee_expected.csv",
):
    fee_master = load_fee_master(fee_master_path)
    phase_cat_fields = get_phase_cat_fields(fee_master_path)

    matched_rows = []
    unmatched = []

    with open(mapping_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            code = row["official_code"].strip()
            sheet_name = resolve_sheet_name(row["eligibility_base"])
            key = (sheet_name, normalize_label(row["row_label"]))
            master_row = fee_master.get(key)

            if master_row is None:
                unmatched.append({"official_code": code, "eligibility_base": row["eligibility_base"], "row_label": row["row_label"].strip()})
                continue

            out_row = {"official_code": code, "eligibility_base": sheet_name, "base_fee": master_row["base_fee"]}
            for field in phase_cat_fields:
                out_row[field] = master_row.get(field, "")
            matched_rows.append(out_row)

    fieldnames = ["official_code", "eligibility_base", "base_fee"] + phase_cat_fields
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in matched_rows:
            writer.writerow(row)

    return len(matched_rows), unmatched


if __name__ == "__main__":
    matched_count, unmatched = build_fee_expected()
    print(f"Matched {matched_count} codes -> lpu_monitor/config/fee_expected.csv")
    if unmatched:
        print(f"\n!!! {len(unmatched)} codes did NOT match fee_master.csv -- fix these before trusting the output:")
        for u in unmatched:
            print(" ", u)
    else:
        print("All codes matched cleanly. No unmatched rows.")