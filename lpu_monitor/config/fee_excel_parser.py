
 
"""
fee_excel_parser.py
Parses the National Fee and Scholarship Excel (old .xls format) into a clean,
flat CSV -- one row per "programme group" label, per sheet -- with ALL active
phases kept as columns (nothing deleted). Phase selection happens later, at
comparison time, based on the live isPhaseN flags -- not here.

The number of phase-blocks is DETECTED from the sheet's column count, not
hardcoded -- LPU adds phase columns as admission progresses through the year
(sometimes just Phase 1 exists, later Phase 1+2, etc.), so a fixed phase
count would need manual updates every time a new phase appears. This adapts
automatically.

Handles the 4 "standard" sheets (After 10th, After 10+2, Lateral Entry,
After Graduation), which all share the same layout. The Ph.D. sheet has a
different structure entirely and is handled separately by phd_fee_parser.py.
"""
import xlrd
import csv
import re

STANDARD_SHEETS = ["After 10th", "After 10+2", "Lateral Entry", "After Graduation"]
CATEGORIES = ["A", "B", "C", "D"]


def detect_phase_count(sheet):
    """
    Instead of hardcoding how many phase-blocks exist (which changes as
    admission progresses through the year, and isn't fixed at 4 or 5 forever),
    compute it from the sheet's actual column count. Each phase block is
    4 columns (A/B/C/D); columns 0-1 are the label and base fee.
    """
    return (sheet.ncols - 2) // 4


def extract_fee_number(cell_value):
    """
    Cell values look like: 56000.0 (a real number) or '56000\n(60%)' (a string
    with a trailing percentage note). Extract just the numeric fee in both cases.
    Returns None if the cell is blank or unparseable (e.g. merged/section-header rows).
    """
    if cell_value == "" or cell_value is None:
        return None
    if isinstance(cell_value, (int, float)):
        return float(cell_value)
    match = re.search(r"[\d,]+", str(cell_value))
    if not match:
        return None
    return float(match.group(0).replace(",", ""))


def parse_sheet(sheet):
    num_phases = detect_phase_count(sheet)
    rows = []
    for r in range(4, sheet.nrows):  # data starts after the 4 header rows (0-3)
        label = sheet.cell_value(r, 0)
        base_fee = extract_fee_number(sheet.cell_value(r, 1))
        if not label or base_fee is None:
            continue  # skip blank rows and section-header rows like "ENGINEERING"

        record = {
            "row_label": label.strip(),
            "base_fee": base_fee,
        }
        for phase in range(1, num_phases + 1):
            start_col = 2 + (phase - 1) * 4
            for i, cat in enumerate(CATEGORIES):
                fee = extract_fee_number(sheet.cell_value(r, start_col + i))
                record[f"phase{phase}_{cat}"] = fee
        rows.append(record)
    return rows, num_phases


def build_fee_lookup(xls_path, out_csv="lpu_monitor/config/fee_master.csv"):
    wb = xlrd.open_workbook(xls_path)
    all_rows = []
    max_phases_seen = 0
    for sheet_name in STANDARD_SHEETS:
        sheet = wb.sheet_by_name(sheet_name)
        records, num_phases = parse_sheet(sheet)
        max_phases_seen = max(max_phases_seen, num_phases)
        for record in records:
            record["eligibility_base"] = sheet_name
            all_rows.append(record)

    fieldnames = ["eligibility_base", "row_label", "base_fee"] + [
        f"phase{p}_{c}" for p in range(1, max_phases_seen + 1) for c in CATEGORIES
    ]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, restval="")
        writer.writeheader()
        for row in all_rows:
            writer.writerow(row)
    print(f"Detected {max_phases_seen} phase(s) in this file.")
    return len(all_rows)


if __name__ == "__main__":
    import sys, glob, os

    def find_single_file(folder, pattern="*.xls"):
        files = glob.glob(os.path.join(folder, pattern))
        if len(files) == 0:
            raise FileNotFoundError(f"No file found in {folder}")
        if len(files) > 1:
            raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
        return files[0]

    if len(sys.argv) > 1:
        xls_path = sys.argv[1]
    else:
        xls_path = find_single_file("lpu_monitor/incoming/fee_excel/")

    count = build_fee_lookup(xls_path)
    print(f"Parsed {count} programme-group rows into lpu_monitor/config/fee_master.csv")
