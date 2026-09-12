"""run_discipline_links_report.py -- Records: Discipline Links only."""
import asyncio, glob, os
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession

OUT_XLSX = "records_discipline_links.xlsx"

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

async def fetch_discipline_links(session, code):
    data = await session.call_api(
        url="https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramDetails",
        method="POST", body={"officialCode": code},
    )
    if not data:
        return None, None
    p = data[0]
    return p.get("schoolLink"), p.get("brochureLink")

async def run():
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes = get_non_phd_codes(proglist)
    wb = Workbook()
    ws = wb.active
    ws.append(["Programme Name", "OfficialCode", "All Links Present?", "School Link", "Brochure Link"])
    async with BrowserSession(headless=True) as session:
        for code, name in codes:
            try:
                school, brochure = await fetch_discipline_links(session, code)
                present = "Yes" if (school and brochure) else "No"
                ws.append([name, code, present, school, brochure])
                print(code, present)
            except Exception as e:
                ws.append([name, code, f"ERROR: {e}", "", ""])
                print(code, "FAILED:", e)
    wb.save(OUT_XLSX)
    print(f"Saved {OUT_XLSX}")

if __name__ == "__main__":
    asyncio.run(run())