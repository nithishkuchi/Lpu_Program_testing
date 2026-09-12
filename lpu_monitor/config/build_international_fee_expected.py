"""Joins fee_row_mapping.csv (REUSED from domestic pipeline -- same row
labels) with international_fee_master.csv -> international_fee_expected.csv"""
import csv, re

ELIGIBILITY_BASE_TO_SHEET = {
    "After 10th": "After 10th", "After 10+2 (12th)": "After 10+2",
    "Lateral Entry/Other Programmes(After ITI etc)": "Lateral Entry",
    "Lateral Entry/Other Programmes(After 3 yrs. Diploma)": "Lateral Entry",
    "After Graduation": "After Graduation",
}

def normalize_label(text):
    text = re.split(r"note\s*:", text, flags=re.IGNORECASE)[0]
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()

def resolve_sheet(raw):
    return ELIGIBILITY_BASE_TO_SHEET.get(raw.strip(), raw.strip())

def load_master(path="lpu_monitor/config/international_fee_master.csv"):
    lookup = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            key = (row["eligibility_base"].strip(), normalize_label(row["row_label"]))
            lookup[key] = row
    return lookup

def build(mapping_path="lpu_monitor/config/international_fee_row_mapping.csv",
          master_path="lpu_monitor/config/international_fee_master.csv",
          out_path="lpu_monitor/config/international_fee_expected.csv"):
    master = load_master(master_path)
    fields = ["base_saarc","base_zone1","base_zone2","saarc_A","saarc_B","saarc_C",
              "zone1_A","zone1_B","zone1_C","zone2_A","zone2_B","zone2_C"]
    matched, unmatched = [], []
    with open(mapping_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            code = row["official_code"].strip()
            sheet = resolve_sheet(row["eligibility_base"])
            key = (sheet, normalize_label(row["row_label"]))
            m = master.get(key)
            if m is None:
                unmatched.append(code); continue
            out = {"official_code": code}
            for fld in fields:
                out[fld] = m.get(fld, "")
            matched.append(out)
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["official_code"] + fields)
        w.writeheader()
        for r in matched:
            w.writerow(r)
    print(f"Matched {len(matched)} -> {out_path}; unmatched: {len(unmatched)}")
    if unmatched:
        print("Unmatched codes:", unmatched)
    return len(matched), unmatched

if __name__ == "__main__":
    build()