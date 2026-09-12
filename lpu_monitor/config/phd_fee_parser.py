"""
phd_fee_parser.py

Parses the PhD sheet of the National Fee and Scholarship Excel INTO A CSV --
same principle as fee_excel_parser.py for the other 4 sheets: nothing about
the actual fee numbers is hardcoded in this file. Every value is read from
the Excel at run time.

Two things made deliberately robust to future changes, per feedback on the
first draft of this task:
1. Sheet name matched by REGEX (contains "ph.d" / "phd", case-insensitive),
   not an exact string -- the sheet is titled "Ph.D. Autumn Term (Aug-2026)"
   this year; the year/term text will change next cycle, and an exact match
   would silently break.
2. No per-programme-code matching needed here at all (unlike the other 4
   sheets) -- this sheet has exactly ONE row that applies to every PhD
   programme regardless of discipline. So the "many codes -> one row"
   mapping step from the other sheets doesn't apply; every PhD code just
   uses this same single row, differentiated only by Full Time vs Part Time.

Source-file lookup: this parses the SAME uploaded Excel as fee_excel_parser.py
(just a different sheet within it), so it must find that file the same way --
by looking inside lpu_monitor/incoming/fee_excel/, not by a hardcoded
filename. A hardcoded name here would silently go stale the moment the file
is renamed or a new year's file is uploaded, even though fee_excel_parser.py
would keep working fine.
"""
import xlrd
import re
import csv


def find_phd_sheet(wb):
    for name in wb.sheet_names():
        if re.search(r"ph\.?\s*d", name, re.IGNORECASE):
            return wb.sheet_by_name(name)
    raise ValueError(
        "No sheet matching a PhD-like name found. Sheet names present: "
        f"{wb.sheet_names()} -- check if the PhD sheet was renamed to something "
        "that no longer contains 'ph.d' or 'phd'."
    )


def forward_fill(values):
    """Merged header cells (e.g. 'Category A' spanning 2 columns) only store
    text in the first cell of the merge; xlrd returns '' for the rest.
    Forward-fill so every column knows which category it belongs to."""
    filled = []
    last = ""
    for v in values:
        if v not in ("", None):
            last = v
        filled.append(last)
    return filled


def parse_phd_sheet(xls_path, out_csv="lpu_monitor/config/phd_fee_master.csv"):
    wb = xlrd.open_workbook(xls_path)
    sheet = find_phd_sheet(wb)

    mode_row = sheet.row_values(1)                       # Full Time / Part Time per column
    category_row = forward_fill(sheet.row_values(2))      # Category A/B/C (merged -> forward-filled)

    # Find the data row dynamically: first row (after the header rows) with
    # a label AND a numeric base fee in column 1 -- not a fixed row number,
    # in case a row gets inserted/removed above it in a future file.
    data_row_idx = None
    for r in range(4, sheet.nrows):
        label = sheet.cell_value(r, 0)
        base_fee = sheet.cell_value(r, 1)
        if label and isinstance(base_fee, (int, float)):
            data_row_idx = r
            break
    if data_row_idx is None:
        raise ValueError(
            "Could not find the PhD data row (expected a row with a text label "
            "in column A and a numeric fee in column B). Sheet structure may "
            "have changed -- inspect manually before proceeding."
        )

    data_row = sheet.row_values(data_row_idx)
    base_fee = data_row[1]

    fields = {"base_fee": base_fee}
    for col in range(2, sheet.ncols):
        mode = str(mode_row[col]).strip() if mode_row[col] else ""
        category = str(category_row[col]).strip() if category_row[col] else ""
        value = data_row[col] if col < len(data_row) else ""
        if value == "" or not category:
            continue

        mode_lower = mode.lower()
        if "full time" in mode_lower or "part time" in mode_lower:
            cat_letter = category.lower().split("category")[-1].strip()[:1].upper() if "category" in category.lower() else "?"
            mode_suffix = "ft" if "full time" in mode_lower else "pt"
            field_name = f"cat{cat_letter}_{mode_suffix}"
        else:
            # Not a Full Time/Part Time column at all -- this is the distinct
            # female-candidates/alumni flat rate, regardless of what the
            # forward-filled category text says (forward-fill wrongly leaks
            # "Category C" into this column since it's the last non-blank
            # value to its left -- this column is NOT actually category-based).
            field_name = "female_alumni_flat"

        fields[field_name] = value

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(fields.keys()))
        writer.writeheader()
        writer.writerow(fields)

    return fields


def find_single_file(folder, pattern="*.xls"):
    """Same lookup fee_excel_parser.py uses: find the one Excel file the UI
    saved into the incoming folder, rather than assuming a fixed filename.
    Kept as an exact duplicate of that function (not imported) so this script
    still runs standalone via `python -m lpu_monitor.config.phd_fee_parser`
    without depending on fee_excel_parser's module being importable first."""
    import glob
    import os

    files = glob.glob(os.path.join(folder, pattern))
    if len(files) == 0:
        raise FileNotFoundError(f"No file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
    return files[0]


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        xls_path = sys.argv[1]
    else:
        xls_path = find_single_file("lpu_monitor/incoming/fee_excel/")

    result = parse_phd_sheet(xls_path)
    print(f"Parsed PhD fee row from {xls_path} (nothing hardcoded):")
    for k, v in result.items():
        print(f"  {k}: {v}")