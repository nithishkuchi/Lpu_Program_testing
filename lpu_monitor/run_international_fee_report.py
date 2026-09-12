"""run_international_fee_report.py -- Records: Live International Fee data."""
import asyncio
import csv
import glob
import os
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.international_fee_fetch import fetch_international_fee

OUT_XLSX = "records_international_fee.xlsx"


def find_single_file(folder, patterns=("*.xls", "*.xlsx")):
    for pat in patterns:
        files = glob.glob(os.path.join(folder, pat))
        if files:
            return files[0]
    return None


def get_programme_names(proglist_path):
    if not proglist_path or not os.path.exists(proglist_path):
        return {}
    import xlrd
    try:
        wb = xlrd.open_workbook(proglist_path)
        sheet = wb.sheet_by_index(0)
        header = [str(h).strip().lower() for h in sheet.row_values(0)]
        code_col = next(i for i, h in enumerate(header) if "programme code" in h)
        name_col = next(
            i for i, h in enumerate(header)
            if any(cand in h for cand in ("programme name", "program name", "course name", "name of programme"))
        )
        lookup = {}
        for r in range(1, sheet.nrows):
            c = str(sheet.cell_value(r, code_col)).strip()
            n = str(sheet.cell_value(r, name_col)).strip()
            if c:
                lookup[c] = n
        return lookup
    except Exception:
        return {}


def get_codes_to_check():
    seed_path = "lpu_monitor/config/international_seed.csv"
    if os.path.exists(seed_path):
        with open(seed_path, encoding="utf-8") as f:
            return [row["official_code"].strip() for row in csv.DictReader(f) if row.get("official_code")]
    
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    if proglist:
        import xlrd
        wb = xlrd.open_workbook(proglist)
        sheet = wb.sheet_by_index(0)
        header = [str(h).strip().lower() for h in sheet.row_values(0)]
        code_col = next(i for i, h in enumerate(header) if "programme code" in h)
        out = []
        for r in range(1, sheet.nrows):
            c = str(sheet.cell_value(r, code_col)).strip()
            if c and "ph.d" not in c.lower():
                out.append(c)
        return out
    return []


async def run(start_idx=0, end_idx=None):
    codes = get_codes_to_check()
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    names_lookup = get_programme_names(proglist)

    if end_idx is not None:
        codes = codes[start_idx:end_idx]
    else:
        codes = codes[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live International Fees"

    headers = [
        "Programme Name",
        "OfficialCode",
        "SAARC Base Tuition Fee",
        "SAARC Cat A (Above 90%)",
        "SAARC Cat B (Above 75%)",
        "SAARC Cat C (Above 60%)",
        "SAARC Uniform Fee",
        "SAARC Exam Fee",
        "Non-SAARC / Zone 1 Base Tuition Fee ($)",
        "Non-SAARC / Zone 1 Cat A ($)",
        "Non-SAARC / Zone 1 Cat B ($)",
        "Non-SAARC / Zone 1 Cat C ($)",
        "Non-SAARC Uniform Fee ($)",
        "Non-SAARC Exam Fee ($)",
        "Other / Zone 2 Base Tuition Fee ($)",
        "Other / Zone 2 Cat A ($)",
        "Other / Zone 2 Cat B ($)",
        "Other / Zone 2 Cat C ($)",
        "Other Uniform Fee ($)",
        "Other Exam Fee ($)",
    ]
    ws.append(headers)

    print(f"Total programmes to fetch live international fees for: {len(codes)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code in codes:
                count += 1
                name = names_lookup.get(code, "")
                try:
                    live = await fetch_international_fee(session, code)
                    if not isinstance(live, dict):
                        ws.append([name, code, "NO_DATA_RETURNED"])
                        print(f"[{count}/{len(codes)}] {code}: No data returned")
                        continue

                    saarc = live.get("saarc") or {}
                    non_saarc = live.get("nonSAARC") or {}
                    other = live.get("other") or {}

                    row = [
                        name,
                        code,
                        saarc.get("tuitionFee", ""),
                        saarc.get("feeAbove90", ""),
                        saarc.get("feeAbove75", ""),
                        saarc.get("feeAbove60", ""),
                        saarc.get("uniformFee", ""),
                        saarc.get("examFee", ""),
                        non_saarc.get("tuitionFee", ""),
                        non_saarc.get("feeAbove90", ""),
                        non_saarc.get("feeAbove75", ""),
                        non_saarc.get("feeAbove60", ""),
                        non_saarc.get("uniformFee", ""),
                        non_saarc.get("examFee", ""),
                        other.get("tuitionFee", ""),
                        other.get("feeAbove90", ""),
                        other.get("feeAbove75", ""),
                        other.get("feeAbove60", ""),
                        other.get("uniformFee", ""),
                        other.get("examFee", ""),
                    ]
                    ws.append(row)
                    print(f"[{count}/{len(codes)}] {code}: SAARC base:{saarc.get('tuitionFee')} Non-SAARC base:{non_saarc.get('tuitionFee')}")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}"])
                    print(f"[{count}/{len(codes)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live international fee records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
