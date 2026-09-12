"""run_phd_fee_report.py -- Records: Live Fee (PhD) data."""
import asyncio
import csv
import glob
import os
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.fee_validation import unwrap_cuet_response, load_phd_fee_expected

OUT_XLSX = "records_phd_fee.xlsx"


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


def load_phd_entries(csv_path="lpu_monitor/config/phd_codes.csv"):
    entries = []
    if os.path.exists(csv_path):
        with open(csv_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                entries.append({
                    "code": row["official_code"].strip(),
                    "program_id": row["program_id"].strip(),
                    "mode": row.get("mode", "").strip(),
                    "name": row.get("programme_name", "").strip(),
                })
    return entries


async def run(start_idx=0, end_idx=None):
    entries = load_phd_entries()
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    names_lookup = get_programme_names(proglist)

    try:
        phd_master = load_phd_fee_expected()
        base_fee = phd_master.get("base_fee", 50000.0)
    except Exception:
        base_fee = 50000.0

    if end_idx is not None:
        entries = entries[start_idx:end_idx]
    else:
        entries = entries[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live PhD Fee Records"

    headers = [
        "Programme Name",
        "OfficialCode",
        "Mode",
        "Base Fee",
        "Live Scholarship Cat A (p981)",
        "Live Scholarship Cat B (p901)",
        "Live Scholarship Cat C (p801)",
        "Live Fee After Cat A",
        "Live Fee After Cat B",
        "Live Fee After Cat C",
    ]
    ws.append(headers)

    print(f"Total PhD programmes to fetch live fees for: {len(entries)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for item in entries:
                count += 1
                code = item["code"]
                pid = item["program_id"]
                mode = item["mode"]
                name = item["name"] or names_lookup.get(code, "")

                try:
                    cuet_raw = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/ScholarshipCUET?id={pid}",
                        method="GET",
                    )
                    cuet = unwrap_cuet_response(cuet_raw)

                    p981 = cuet.get("p981")
                    p901 = cuet.get("p901")
                    p801 = cuet.get("p801")

                    fee_a = (base_fee - float(p981)) if p981 is not None else ""
                    fee_b = (base_fee - float(p901)) if p901 is not None else ""
                    fee_c = (base_fee - float(p801)) if p801 is not None else ""

                    ws.append([
                        name, code, mode, base_fee,
                        p981 if p981 is not None else "",
                        p901 if p901 is not None else "",
                        p801 if p801 is not None else "",
                        fee_a, fee_b, fee_c,
                    ])
                    print(f"[{count}/{len(entries)}] {code} ({mode}): CatA:{p981} CatB:{p901} CatC:{p801}")
                except Exception as e:
                    ws.append([name, code, mode, base_fee, f"ERROR: {e}", "", "", "", "", ""])
                    print(f"[{count}/{len(entries)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live PhD fee records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
