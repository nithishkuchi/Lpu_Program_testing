"""run_international_exposure_report.py -- Records: Live International Exposure data."""
import asyncio
import csv
import glob
import json
import os
import re
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.fee_detail_client import fetch_fee_detail

OUT_XLSX = "records_international_exposure.xlsx"

INTL_EXPOSURE_TITLE = "International Exposure Option"
SEMESTER_EXCHANGE_TITLE = "Semester Exchange"


def find_single_file(folder, patterns=("*.xls", "*.xlsx")):
    for pat in patterns:
        files = glob.glob(os.path.join(folder, pat))
        if files:
            return files[0]
    return None


def strip_html(text):
    if not text:
        return ""
    return re.sub(r"<[^>]+>", "", str(text))


def get_all_codes_and_names(proglist_path, phd_codes_path="lpu_monitor/config/phd_codes.csv"):
    codes_names = []
    if proglist_path and os.path.exists(proglist_path):
        import xlrd
        wb = xlrd.open_workbook(proglist_path)
        sheet = wb.sheet_by_index(0)
        header = [str(h).strip().lower() for h in sheet.row_values(0)]
        code_col = next(i for i, h in enumerate(header) if "programme code" in h)
        name_col = next(
            i for i, h in enumerate(header)
            if any(cand in h for cand in ("programme name", "program name", "course name", "name of programme"))
        )
        elig_col = next((i for i, h in enumerate(header) if "eligibility base" in h), None)

        for r in range(1, sheet.nrows):
            code = str(sheet.cell_value(r, code_col)).strip()
            name = str(sheet.cell_value(r, name_col)).strip()
            elig = str(sheet.cell_value(r, elig_col)).strip().lower() if elig_col is not None else ""
            if not code:
                continue
            if "ph.d" in elig or "ph.d" in code.lower() or "ph.d" in name.lower():
                continue
            codes_names.append((code, name))

    if os.path.exists(phd_codes_path):
        with open(phd_codes_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                c = row["official_code"].strip()
                n = row.get("programme_name", "").strip() or c
                codes_names.append((c, n))

    return codes_names


async def run(start_idx=0, end_idx=None):
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes_names = get_all_codes_and_names(proglist)

    if end_idx is not None:
        codes_names = codes_names[start_idx:end_idx]
    else:
        codes_names = codes_names[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live International Exposure"

    headers = [
        "Programme Name",
        "OfficialCode",
        "International Exposure Option",
        "Semester Exchange",
        "Both Options Present?",
    ]
    ws.append(headers)

    print(f"Total programmes to fetch live international exposure for: {len(codes_names)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code, name in codes_names:
                count += 1
                try:
                    live = await fetch_fee_detail(session, code)
                    text = strip_html(json.dumps(live))
                    live_intl = INTL_EXPOSURE_TITLE in text
                    live_exch = SEMESTER_EXCHANGE_TITLE in text

                    str_intl = "Yes" if live_intl else "No"
                    str_exch = "Yes" if live_exch else "No"
                    both = "Yes" if (live_intl and live_exch) else "No"

                    ws.append([name, code, str_intl, str_exch, both])
                    print(f"[{count}/{len(codes_names)}] {code}: Exposure={str_intl}, SemesterExchange={str_exch}")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}", "", ""])
                    print(f"[{count}/{len(codes_names)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live international exposure records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
