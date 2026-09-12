"""international_fee_excel_parser.py -- same structure as fee_excel_parser.py
but NO phases (confirmed: single flat table year-round) and 3 zones x 3
categories instead of 4 phases x 4 categories."""
import xlrd, csv, re, glob, os

STANDARD_SHEETS = ["After 10th", "After 10+2", "Lateral Entry", "After Graduation"]
ZONES = ["saarc", "zone1", "zone2"]
CATS = ["A", "B", "C"]

def extract_fee_number(v):
    if v == "" or v is None: return None
    if isinstance(v, (int, float)): return float(v)
    m = re.search(r"[\d,]+", str(v))
    return float(m.group(0).replace(",", "")) if m else None

def parse_sheet(sheet):
    rows = []
    for r in range(4, sheet.nrows):
        label = sheet.cell_value(r, 0)
        base_saarc = extract_fee_number(sheet.cell_value(r, 1))
        if not label or base_saarc is None:
            continue
        rec = {"row_label": str(label).strip(),
               "base_saarc": base_saarc,
               "base_zone1": extract_fee_number(sheet.cell_value(r, 2)),
               "base_zone2": extract_fee_number(sheet.cell_value(r, 3))}
        col = 4
        for zone in ZONES:
            for cat in CATS:
                rec[f"{zone}_{cat}"] = extract_fee_number(sheet.cell_value(r, col))
                col += 1
        rows.append(rec)
    return rows

def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) != 1:
        raise ValueError(f"Expected 1 file in {folder}, found {len(files)}")
    return files[0]

def build(xls_path, out_csv="lpu_monitor/config/international_fee_master.csv"):
    wb = xlrd.open_workbook(xls_path)
    all_rows = []
    for sname in STANDARD_SHEETS:
        sheet = wb.sheet_by_name(sname)
        for rec in parse_sheet(sheet):
            rec["eligibility_base"] = sname
            all_rows.append(rec)
    fields = ["eligibility_base", "row_label", "base_saarc", "base_zone1", "base_zone2"] + \
             [f"{z}_{c}" for z in ZONES for c in CATS]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, restval="")
        w.writeheader()
        for r in all_rows:
            w.writerow(r)
    print(f"Parsed {len(all_rows)} rows -> {out_csv}")
    return len(all_rows)

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else find_single_file("lpu_monitor/incoming/international_fee_excel/")
    build(path)