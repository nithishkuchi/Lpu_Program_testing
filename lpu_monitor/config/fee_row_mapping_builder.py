"""
fee_row_mapping_builder.py

Wraps the teammate-provided matching engine (programme_fee_matcher.py --
kept completely unmodified) with:
  1. Auto-discovery of the two input files from the incoming/ folders
     (no hardcoded filenames -- same pattern as every other parser in
     this project).
  2. .xls -> .xlsx conversion using pure Python (pandas + xlrd + openpyxl)
     -- avoids requiring LibreOffice to be installed on the machine
     running this, which the original script needed as an external
     dependency.
  3. Writing output in the exact column format build_fee_expected.py
     already expects (official_code, eligibility_base, row_label) --
     build_fee_expected.py's existing resolve_sheet_name() already
     handles converting the raw Eligibility Base text, so no extra
     normalization is needed here.
"""
import glob
import os
import csv
import pandas as pd
import sys

sys.path.insert(0, os.path.dirname(__file__))
import programme_fee_matcher as matcher

INCOMING_FEE = "lpu_monitor/incoming/fee_excel"
INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"
WORKDIR = "lpu_monitor/config/_xlsx_cache"


def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) == 0:
        raise FileNotFoundError(f"No file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
    return files[0]


def convert_xls_to_xlsx(xls_path, out_dir):
    """Pure-Python .xls -> .xlsx conversion, no LibreOffice required."""
    os.makedirs(out_dir, exist_ok=True)
    out_name = os.path.splitext(os.path.basename(xls_path))[0] + ".xlsx"
    out_path = os.path.join(out_dir, out_name)

    xls = pd.ExcelFile(xls_path, engine="xlrd")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=None)
            df.to_excel(writer, sheet_name=sheet, index=False, header=False)
    return out_path


def build_fee_row_mapping(
    proglist_path=None,
    fee_path=None,
    out_csv="lpu_monitor/config/fee_row_mapping.csv",
):
    proglist_path = proglist_path or find_single_file(INCOMING_PROGLIST)
    fee_path = fee_path or find_single_file(INCOMING_FEE)

    proglist_xlsx = convert_xls_to_xlsx(proglist_path, WORKDIR)
    fee_xlsx = convert_xls_to_xlsx(fee_path, WORKDIR)

    programmes, prog_wb = matcher.load_programme_list(proglist_xlsx)
    fee_rows, fee_wb = matcher.load_fee_rows(fee_xlsx)
    expanded_to_short, _ = matcher.load_abbreviations(prog_wb)
    if not expanded_to_short:
        expanded_to_short, _ = matcher.load_abbreviations(fee_wb)

    mapping = matcher.compute_mapping_rows(programmes, fee_rows, expanded_to_short)

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["official_code", "eligibility_base", "row_label"])
        writer.writeheader()
        for r in mapping:
            writer.writerow({
                "official_code": r["OfficialCode"],
                "eligibility_base": r["EligibilityBase"],
                "row_label": r["FeeRowLabel"],
            })

    from collections import Counter
    status_counts = Counter(r["Status"] for r in mapping)
    needs_review = [r for r in mapping if r["Status"] != "HIGH"]
    return len(mapping), status_counts, needs_review


if __name__ == "__main__":
    count, status_counts, needs_review = build_fee_row_mapping()
    print(f"Wrote {count} rows to lpu_monitor/config/fee_row_mapping.csv")
    print("Status breakdown:", dict(status_counts))
    if needs_review:
        print(f"\n{len(needs_review)} row(s) need manual review:")
        for r in needs_review:
            print(f"  - {r['OfficialCode']}: {r['ProgrammeName']}  [{r['Status']}]")
        sys.exit(1)  # non-zero so the UI clearly flags "needs attention"