"""run_annexure_report.py -- Records: Live Scholarship / Annexure data."""
import asyncio
import glob
import os
import re
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.annexure_check import normalize, ANNEXURE_TITLE_MATCH

OUT_XLSX = "records_annexure.xlsx"


def find_single_file(folder, patterns=("*.xls", "*.xlsx")):
    for pat in patterns:
        files = glob.glob(os.path.join(folder, pat))
        if files:
            return files[0]
    return None


def get_programme_codes_and_names(proglist_path):
    if not proglist_path or not os.path.exists(proglist_path):
        return []
    import xlrd
    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)
    header = [str(h).strip().lower() for h in sheet.row_values(0)]
    code_col = next(i for i, h in enumerate(header) if "programme code" in h)
    name_col = next(
        i for i, h in enumerate(header)
        if any(cand in h for cand in ("programme name", "program name", "course name", "name of programme"))
    )
    out = []
    for r in range(1, sheet.nrows):
        c = str(sheet.cell_value(r, code_col)).strip()
        n = str(sheet.cell_value(r, name_col)).strip()
        if c:
            out.append((c, n))
    return out


async def run(start_idx=0, end_idx=None):
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes_names = get_programme_codes_and_names(proglist)

    if end_idx is not None:
        codes_names = codes_names[start_idx:end_idx]
    else:
        codes_names = codes_names[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Scholarship Annexures"

    headers = [
        "Programme Name",
        "OfficialCode",
        "Annexure A (Need Based)",
        "Annexure B (Orphan)",
        "Annexure C (Disability)",
        "Annexure D (Rakshak)",
        "Annexure E (LPU RISE)",
        "Annexure F (Shikshak)",
        "Annexure H/I (Sports)",
        "Annexure J (Start Up)",
        "All Live Scholarship Tabs",
    ]
    ws.append(headers)

    print(f"Total programmes to fetch live scholarship annexures for: {len(codes_names)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code, name in codes_names:
                count += 1
                if code == "P2L2":
                    ws.append([name, code, "SKIPPED", "", "", "", "", "", "", "", "No live data per code_exceptions"])
                    continue

                try:
                    tabs = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramScholarshipTabs?id={code}&filter=1",
                        method="GET",
                    )
                    if not isinstance(tabs, list):
                        ws.append([name, code, "No tabs list returned", "", "", "", "", "", "", "", ""])
                        print(f"[{count}/{len(codes_names)}] {code}: No tabs list returned")
                        continue

                    raw_titles = [str(item.get("title", "")).strip() for item in tabs if item.get("title")]
                    normalized_titles = [normalize(t) for t in raw_titles]

                    def check_annexure(letter):
                        kw = normalize(ANNEXURE_TITLE_MATCH[letter])
                        return "Yes" if any(kw in t for t in normalized_titles) else "No"

                    ann_a = check_annexure("A")
                    ann_b = check_annexure("B")
                    ann_c = check_annexure("C")
                    ann_d = check_annexure("D")
                    ann_e = check_annexure("E")
                    ann_f = check_annexure("F")
                    ann_hi = "Yes" if (check_annexure("H") == "Yes" or check_annexure("I") == "Yes") else "No"
                    ann_j = check_annexure("J")
                    all_tabs = " | ".join(raw_titles)

                    ws.append([
                        name, code,
                        ann_a, ann_b, ann_c, ann_d, ann_e, ann_f, ann_hi, ann_j,
                        all_tabs,
                    ])
                    print(f"[{count}/{len(codes_names)}] {code}: {len(raw_titles)} tabs found live")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}", "", "", "", "", "", "", "", ""])
                    print(f"[{count}/{len(codes_names)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live scholarship/annexure records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
