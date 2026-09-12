import csv
import re
import os

ELIGIBILITY_BASE_TO_SHEET = {
    "After 10th": "After 10th",
    "After 10+2 (12th)": "After 10+2",
    "Lateral Entry/Other Programmes(After ITI etc)": "Lateral Entry",
    "Lateral Entry/Other Programmes(After 3 yrs. Diploma)": "Lateral Entry",
    "After Graduation": "After Graduation",
}

def resolve_sheet_name(eligibility_base_raw):
    key = eligibility_base_raw.strip()
    if key in ELIGIBILITY_BASE_TO_SHEET: return ELIGIBILITY_BASE_TO_SHEET[key]
    return key

def normalize_label(text):
    if not text: return ""
    text = str(text)
    text = re.sub(r"[\s\xa0]+", " ", text)
    text = re.split(r"note\s*:", text, flags=re.IGNORECASE)[0]
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    text = re.sub(r"\s+", " ", text).strip()
    return text

def load_eligibility_master(path="lpu_monitor/config/eligibility_master.csv"):
    lookup = {}; for_collision_check = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            key = (row["eligibility_base"].strip(), normalize_label(row["row_label"]))
            if key in for_collision_check and for_collision_check[key] != row["row_label"]:
                raise ValueError(f"COLLISION: {key}\n1: {for_collision_check[key]!r}\n2: {row['row_label']!r}")
            for_collision_check[key] = row["row_label"]; lookup[key] = row
    return lookup

def build_eligibility_expected(mapping_path="lpu_monitor/config/eligibility_row_mapping.csv", master_path="lpu_monitor/config/eligibility_master.csv", out_path="lpu_monitor/config/eligibility_expected.csv"):
    master = load_eligibility_master(master_path); matched_rows = []; unmatched = []
    with open(mapping_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            code = row.get("official_code", row.get("OfficialCode", "")).strip()
            raw_base = row.get("eligibility_base", row.get("EligibilityBase", "")).strip()
            label = row.get("row_label", row.get("MatchedLabel", "")).strip()
            sheet_name = resolve_sheet_name(raw_base); key = (sheet_name, normalize_label(label))
            master_row = master.get(key)
            if master_row is None: unmatched.append({"official_code": code, "eligibility_base": raw_base, "row_label": label}); continue
            out_row = {"official_code": code, "eligibility_base": sheet_name, "row_label": label, "eligibility_indian": master_row.get("eligibility_indian", ""), "eligibility_international": master_row.get("eligibility_international", ""), "relaxation": master_row.get("relaxation", ""), "remarks": master_row.get("remarks", ""), "equivalent_qualification": master_row.get("equivalent_qualification", "")}
            matched_rows.append(out_row)
    fieldnames = ["official_code", "eligibility_base", "row_label", "eligibility_indian", "eligibility_international", "relaxation", "remarks", "equivalent_qualification"]
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames); writer.writeheader()
        for row in matched_rows: writer.writerow(row)
    return len(matched_rows), unmatched

if __name__ == "__main__":
    matched_count, unmatched = build_eligibility_expected()
    print(f"Matched {matched_count} codes -> lpu_monitor/config/eligibility_expected.csv")
    if unmatched:
        print(f"\n!!! {len(unmatched)} codes did NOT match eligibility_master.csv:")
        for u in unmatched[:15]: print(" ", u)
        if len(unmatched) > 15: print(f" ... and {len(unmatched) - 15} more.")
    else:
        print("All codes matched cleanly. No unmatched rows.")