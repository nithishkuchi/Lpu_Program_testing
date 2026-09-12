import xlrd
import csv
import glob
import os

def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) == 0:
        raise FileNotFoundError(f"No file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
    return files[0]

SRC_XLS = find_single_file("lpu_monitor/incoming/programme_list/")
OUT_CSV = "lpu_monitor/config/annexure_lookup.csv"

ANNEXURE_COLUMNS = {
    "A": "Annexure-A (Needbased)\n2026",
    "B": "Annexure-B (Orphan Scholarship)*\n2026",
    "C": "Annexure-C (Person with Certain Disability) - Including Blind\n2026",
    "D": "Annexure-D (LPU Rakshak Samman Anudan)\n2026",
    "E": "Annexure-E (LPURISE)\n2026",
    "F": "Annexure-F (LPU Shikshak Samman Anudan)\n2026",
    "G": "Annexure-G (LPU Staff Fee Waiver)\n2026",
    "H": "Annexure-H (Cultural, RandD, Co-curricular, Social Service and Bravery Awards)\n2026",
    "I": "Annexure-I (Sports)\n2026",
    "J": "Annexure-J\n(Startup, Innovation, and Entrepreneurship)\n2026",
}
ANNEXURE_TITLE_MATCH = {
    "A": "Need Based",
    "B": "Orphan",
    "C": "Certain Disability",
    "D": "Rakshak",
    "E": "LPU RISE",
    "F": "Shikshak",
    # G has no public-facing tab
    "H": "Sports, Cultural",
    "I": "Sports, Cultural",
    "J": "Startup",
}

def build_lookup():
    wb = xlrd.open_workbook(SRC_XLS)
    sheet = wb.sheet_by_index(0)  # "Programme list"
    header = sheet.row_values(0)
    code_idx = header.index("Programme code 2015 onwards")
    col_idx = {k: header.index(v) for k, v in ANNEXURE_COLUMNS.items()}

    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["OfficialCode"] + list(ANNEXURE_COLUMNS.keys()))
        for r in range(1, sheet.nrows):
            row = sheet.row_values(r)
            code = row[code_idx]
            if not code:
                continue
            writer.writerow([code] + [row[col_idx[k]] for k in ANNEXURE_COLUMNS])

if __name__ == "__main__":
    build_lookup()