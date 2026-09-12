"""run_eligibility_report.py -- Records: Live Eligibility data."""
import asyncio
import glob
import os
import re
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.program_api import fetch_program_details

OUT_XLSX = "records_eligibility.xlsx"


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


def clean_html(text):
    if not text:
        return ""
    text = re.sub(r"<br\s*/?>", "\n", str(text), flags=re.IGNORECASE)
    text = re.sub(r"</li>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


async def run(start_idx=0, end_idx=None):
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes_names = get_programme_codes_and_names(proglist)

    if end_idx is not None:
        codes_names = codes_names[start_idx:end_idx]
    else:
        codes_names = codes_names[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Eligibility Records"

    headers = [
        "Programme Name",
        "OfficialCode",
        "Live Indian Eligibility (Clean Text)",
        "Live Eligibility Comments / Relaxation",
        "Live Admission Criteria",
        "Live International Eligibility",
        "Live International Eligibility Comments",
        "Live Indian Eligibility (Raw HTML)",
    ]
    ws.append(headers)

    print(f"Total programmes to fetch live eligibility for: {len(codes_names)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code, name in codes_names:
                count += 1
                try:
                    details = await fetch_program_details(session, code)
                    raw_indian = details.get("eligibilityindian", "") or ""
                    clean_indian = clean_html(raw_indian)
                    comments = clean_html(details.get("indianeligibilitycomments", "") or "")
                    admission = clean_html(details.get("admissionindian", "") or "")
                    intl_elig = clean_html(details.get("eligibilityinternational", "") or "")
                    intl_comments = clean_html(details.get("internationaleligibilitycomments", "") or "")

                    ws.append([
                        name,
                        code,
                        clean_indian,
                        comments,
                        admission,
                        intl_elig,
                        intl_comments,
                        raw_indian,
                    ])
                    preview = (clean_indian[:50] + "...") if len(clean_indian) > 50 else clean_indian
                    print(f"[{count}/{len(codes_names)}] {code}: {preview.replace(chr(10), ' ')}")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}", "", "", "", "", ""])
                    print(f"[{count}/{len(codes_names)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live eligibility records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
