"""
run_international_exposure_batch.py

Reuses fetch_fee_detail (already built for fee validation -- same
GetProgramFeeDetail endpoint, no new fetch logic needed) and compares
against the Programme List Excel's two confirmed ground-truth columns.
"""
import asyncio
import csv
import os
import glob

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.fee_detail_client import fetch_fee_detail
from lpu_monitor.config.international_exposure_validation import (
    load_expected_exposure, check_international_exposure,
)

OUT_PATH = "international_exposure_results.csv"


def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) != 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}")
    return files[0]


def get_all_codes(proglist_path, phd_codes_path="lpu_monitor/config/phd_codes.csv"):
    """
    Returns the full code list to check: every non-PhD code from the
    Programme List sheet, PLUS the 135 discipline-specific PhD codes from
    phd_codes.csv (e.g. P1D2) -- but explicitly EXCLUDING the 3 "After
    Ph.D" higher-doctorate codes (P8G/P8H/P8J), per instruction.
    """
    import xlrd
    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)
    header = sheet.row_values(0)
    code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
    elig_col = next(i for i, h in enumerate(header) if "eligibility base" in h.lower())

    codes = []
    for r in range(1, sheet.nrows):
        code = sheet.cell_value(r, code_col)
        elig = str(sheet.cell_value(r, elig_col))
        if not code:
            continue
        if "ph.d" in elig.lower() or "ph.d" in str(code).lower():
            continue  # excludes P8G/P8H/P8J -- not wanted, per instruction
        codes.append(code)

    if os.path.exists(phd_codes_path):
        with open(phd_codes_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                codes.append(row["official_code"])

    return codes


async def run_batch():
    proglist_path = find_single_file("lpu_monitor/incoming/programme_list")
    expected = load_expected_exposure(proglist_path)
    codes = get_all_codes(proglist_path)

    async with BrowserSession(headless=True) as session:
        with open(OUT_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for code in codes:
                try:
                    live = await fetch_fee_detail(session, code)
                    result = check_international_exposure(code, live, expected)
                    writer.writerow([code, result])
                    print(code, "->", result)
                except Exception as e:
                    writer.writerow([code, f"ERROR: {e}"])
                    print(code, "FAILED:", e)


if __name__ == "__main__":
    if os.path.exists(OUT_PATH):
        os.remove(OUT_PATH)
    asyncio.run(run_batch())