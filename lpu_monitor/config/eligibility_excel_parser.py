"""
lpu_monitor/config/eligibility_parser.py
Parses the eligibility Excel → eligibility_master.csv
"""
import pandas as pd
import csv
import re
import os
import glob

def normalize_label(text):
    if not text or pd.isna(text):
        return ""
    text = str(text)
    text = re.sub(r"[\s\xa0]+", " ", text)
    text = re.split(r"note\s*:", text, flags=re.IGNORECASE)[0]
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    text = re.sub(r"\s+", " ", text).strip()
    return text

def is_category_header(label):
    label = str(label).strip()
    if not label:
        return True
    label_lower = label.lower()

    degree_keywords = [
        "diploma", "b.tech", "m.tech", "b.sc", "m.sc", "bba", "mba", "b.com", "m.com",
        "bca", "mca", "b.a", "m.a", "b.pharm", "m.pharm", "bpt", "mpt", "bfa", "mfa",
        "b.plan", "m.plan", "ll.b", "ll.m", "ph.d", "bhmct", "dmlt", "b.design",
        "m.design", "bachelor", "master", "integrated", "b.ed", "m.ed", "bpes",
        "bpa", "mpa", "bsw", "msw", "mfm", "d.litt", "d.sc", "ll.d", "m.phil"
    ]
    if any(kw in label_lower for kw in degree_keywords):
        return False

    section_keywords = [
        "engineering", "management", "architecture", "medical", "computer applications",
        "agriculture", "food technology", "nutrition", "biotechnology", "microbiology",
        "forensic", "zoology", "physics", "chemistry", "mathematics", "commerce",
        "economics", "hotel", "law", "planning", "interior", "product", "interaction",
        "fashion", "multimedia", "fine arts", "journalism", "performing", "psychology",
        "arts", "sociology", "governance", "library", "education", "physical",
        "pharmaceutical", "physiotherapy", "ayurvedic", "environmental", "nanotechnology",
        "bioinformatics", "biochemistry", "botany", "geography", "history",
        "political science", "public administration", "indian languages", "other programmes",
        "after graduation", "after 10th", "after 10+2", "lateral entry"
    ]
    if label_lower in section_keywords or label_lower.replace(" ", "") in section_keywords:
        return True

    return False

def parse_sheet(df, eligibility_base):
    rows = []
    header_idx = None
    for i, row in df.iterrows():
        row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
        if "eligibility" in row_str and ("indian" in row_str or "programme" in row_str or "program" in row_str):
            header_idx = i
            break
    if header_idx is None:
        return rows

    df = df.iloc[header_idx:].reset_index(drop=True)
    df.columns = [str(c).strip().lower() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    name_col = next((col for col in df.columns if "programme" in col or "program" in col), None)
    indian_col = next((col for col in df.columns if "indian" in col), None)
    intl_col = next((col for col in df.columns if "international" in col), None)
    relax_col = next((col for col in df.columns if "relax" in col), None)
    remarks_col = next((col for col in df.columns if "remark" in col), None)
    equiv_col = next((col for col in df.columns if col.strip().startswith("equivalent")), None)

    if name_col is None:
        return rows

    def _is_blank(val):
        """Treat None, empty, whitespace-only, and bare '*' as blank cells."""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return True
        s = str(val).strip()
        return s == "" or s == "*"

    # Carry-forward: track RAW (un-normalized) text per column.
    last_elig = {
        "indian": "", "international": "",
        "relaxation": "", "remarks": "", "equivalent": ""
    }

    for _, row in df.iterrows():
        label = row.get(name_col)
        if _is_blank(label):
            continue
        label_str = str(label).strip()
        if is_category_header(label_str):
            continue

        # Read RAW values — do NOT normalize eligibility text.
        indian_raw = "" if _is_blank(row.get(indian_col)) else str(row.get(indian_col)).strip()
        intl_raw = "" if _is_blank(row.get(intl_col)) else str(row.get(intl_col)).strip()
        relax_raw = "" if _is_blank(row.get(relax_col)) else str(row.get(relax_col)).strip()
        remarks_raw = "" if _is_blank(row.get(remarks_col)) else str(row.get(remarks_col)).strip()
        equiv_raw = "" if _is_blank(row.get(equiv_col)) else str(row.get(equiv_col)).strip()

        # Per-column carry-forward on raw values
        use_indian = indian_raw or last_elig["indian"]
        use_intl = intl_raw or last_elig["international"]
        use_relax = relax_raw or last_elig["relaxation"]
        use_remarks = remarks_raw or last_elig["remarks"]
        use_equiv = equiv_raw or last_elig["equivalent"]

        if not use_indian and not use_intl:
            continue

        if indian_raw:   last_elig["indian"] = indian_raw
        if intl_raw:     last_elig["international"] = intl_raw
        if relax_raw:    last_elig["relaxation"] = relax_raw
        if remarks_raw:  last_elig["remarks"] = remarks_raw
        if equiv_raw:    last_elig["equivalent"] = equiv_raw

        record = {
            "eligibility_base": eligibility_base,
            "row_label": normalize_label(label_str),   # ONLY row_label is normalized
            "eligibility_indian": use_indian,           # raw text, % intact
            "eligibility_international": use_intl,
            "relaxation": use_relax,
            "remarks": use_remarks,
            "equivalent_qualification": use_equiv,
        }
        rows.append(record)
    return rows

def build_eligibility_lookup(xls_path, out_csv="lpu_monitor/config/eligibility_master.csv"):
    try:
        if xls_path.lower().endswith('.xlsx'):
            xls = pd.ExcelFile(xls_path, engine='openpyxl')
        else:
            xls = pd.ExcelFile(xls_path, engine='xlrd')
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return 0

    all_rows = []
    for sheet_name in xls.sheet_names:
        category = None
        name_lower = sheet_name.lower().strip()
        if "diploma" in name_lower or "after 10th" in name_lower:
            category = "After 10th"
        elif "10+2" in name_lower or "after 10+2" in name_lower or "12th" in name_lower:
            category = "After 10+2"
        elif "lateral" in name_lower:
            category = "Lateral Entry"
        elif "after graduation" in name_lower:
            category = "After Graduation"

        if category:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            records = parse_sheet(df, category)
            all_rows.extend(records)
            print(f"  -> Parsed {len(records)} rows from sheet: '{sheet_name}' (Mapped to: '{category}')")

    fieldnames = [
        "eligibility_base", "row_label", "eligibility_indian", "eligibility_international",
        "relaxation", "remarks", "equivalent_qualification"
    ]
    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, restval="")
        writer.writeheader()
        for row in all_rows:
            writer.writerow(row)
    print(f"\nSuccessfully built eligibility master with {len(all_rows)} total rows into {out_csv}")
    return len(all_rows)

if __name__ == "__main__":
    import sys
    def find_single_file(folder, pattern="*.xls*"):
        files = glob.glob(os.path.join(folder, pattern))
        if len(files) == 0:
            raise FileNotFoundError(f"No file found in {folder}")
        if len(files) > 1:
            raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
        return files[0]
    if len(sys.argv) > 1:
        xls_path = sys.argv[1]
    else:
        xls_path = find_single_file("lpu_monitor/incoming/eligibility_excel/")
    count = build_eligibility_lookup(xls_path)