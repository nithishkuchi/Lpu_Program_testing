"""run_lpunest_links_report.py -- Records: LPUNEST Links only."""
import asyncio, re, glob, os
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession

OUT_XLSX = "records_lpunest_links.xlsx"

def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) != 1:
        raise ValueError(f"Expected 1 file in {folder}, found {len(files)}")
    return files[0]

def get_non_phd_codes(proglist_path):
    import xlrd
    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)
    header = sheet.row_values(0)
    code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
    name_col = next(i for i, h in enumerate(header) if "programme name" in h.lower())
    out = []
    for r in range(1, sheet.nrows):
        code, name = sheet.cell_value(r, code_col), sheet.cell_value(r, name_col)
        if code and "ph.d" not in str(code).lower() and "ph.d" not in str(name).lower():
            out.append((code, name))
    return out

def extract_link(details):
    m = re.search(r'href=["\']([^"\']+)["\']', details or "", re.I)
    return m.group(1) if m else None

async def fetch_lpunest_links(session, code):
    data = await session.call_api(
        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgrammePlacementsRankingsDetails/{code}",
        method="GET",
    )
    wanted = {"Question Paper Pattern": None, "Syllabus": None, "Sample Questions": None, "Test Cities": None}
    if not isinstance(data, list):
        return wanted
    for item in data:
        if item.get("headingText") != "LPUNEST Related Information" or item.get("type") != "L":
            continue
        details = item.get("details", "")
        m = re.search(r"<strong>\s*(.*?)\s*</strong>", details, re.I | re.S)
        if not m:
            continue
        name = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if name in wanted:
            wanted[name] = extract_link(details)
    return wanted

async def run():
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes = get_non_phd_codes(proglist)
    wb = Workbook()
    ws = wb.active
    ws.append(["Programme Name", "OfficialCode", "All Links Present?",
               "Question Paper Pattern", "Syllabus", "Sample Questions", "Test Cities"])
    async with BrowserSession(headless=True) as session:
        for code, name in codes:
            try:
                links = await fetch_lpunest_links(session, code)
                vals = list(links.values())
                ws.append([name, code, "Yes" if all(vals) else "No", *vals])
                print(code, "OK" if all(vals) else "MISSING SOME")
            except Exception as e:
                ws.append([name, code, f"ERROR: {e}", "", "", "", ""])
                print(code, "FAILED:", e)
    wb.save(OUT_XLSX)
    print(f"Saved {OUT_XLSX}")

if __name__ == "__main__":
    asyncio.run(run())