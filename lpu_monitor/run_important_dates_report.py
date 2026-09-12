"""run_important_dates_report.py -- Records: Live Important Dates data."""
import asyncio
import glob
import os
import re
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.important_dates_fetch import fetch_important_dates

OUT_XLSX = "records_important_dates.xlsx"

ALWAYS_SKIP_CODES = {"P8G", "P8H", "P8J", "P132-NNL", "P2L2"}


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
        if c and c not in ALWAYS_SKIP_CODES:
            out.append((c, n))
    return out


def strip_html(text):
    if not text:
        return ""
    return re.sub(r"<[^>]+>", "", str(text)).strip()


def parse_schedule_items(items):
    """Parses live schedule items into active dates and structured key fields."""
    if not isinstance(items, list):
        return None, "", {}

    active_schedule = None
    summary_lines = []
    item_map = {}

    for item in items:
        raw_title = strip_html(item.get("title", ""))
        norm_title = raw_title.lower()

        # Find active schedule value
        val = ""
        sched_no = None
        for n in range(1, 6):
            if item.get(f"isSchedule{n}Active"):
                val = strip_html(item.get(f"schedule{n}", ""))
                sched_no = n
                break
        
        # Fallback to schedule 1 if no active boolean is true
        if not val:
            for n in range(1, 6):
                s_val = strip_html(item.get(f"schedule{n}", ""))
                if s_val:
                    val = s_val
                    sched_no = n
                    break

        if sched_no and active_schedule is None:
            active_schedule = sched_no

        if raw_title and val:
            summary_lines.append(f"{raw_title}: {val}")
            item_map[norm_title] = val

    # Helper to find date for candidate keywords
    def find_date(*keywords):
        for t, v in item_map.items():
            if all(k in t for k in keywords):
                return v
        return ""

    key_dates = {
        "reg_start": find_date("registration"),
        "slot_booking": find_date("slot"),
        "last_date_apply": find_date("last date", "apply") or find_date("submission"),
        "entrance_exam": find_date("entrance") or find_date("dates of examination"),
        "display_result": find_date("result") if "scholarship" not in find_date("result") else "",
        "last_date_adm": find_date("admission"),
        "scholarship_exam": find_date("scholarship exam"),
        "scholarship_result": find_date("scholarship result"),
    }

    return active_schedule, "\n".join(summary_lines), key_dates


async def run(start_idx=0, end_idx=None):
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes_names = get_programme_codes_and_names(proglist)

    if end_idx is not None:
        codes_names = codes_names[start_idx:end_idx]
    else:
        codes_names = codes_names[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Important Dates"

    headers = [
        "Programme Name",
        "OfficialCode",
        "Active Schedule Number",
        "All Active Dates (Summary)",
        "Online Registration Starts",
        "Slot Booking for Online Exam",
        "Last Date to Apply / Book Slot",
        "Dates of Entrance Exam",
        "Display of Result",
        "Last Date of Admission",
        "Dates of Scholarship Exam",
        "Scholarship Result",
    ]
    ws.append(headers)

    print(f"Total programmes to fetch live important dates for: {len(codes_names)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code, name in codes_names:
                count += 1
                try:
                    live = await fetch_important_dates(session, code)
                    active_sched, summary, kd = parse_schedule_items(live)

                    ws.append([
                        name,
                        code,
                        f"Schedule {active_sched}" if active_sched else "None",
                        summary,
                        kd.get("reg_start", ""),
                        kd.get("slot_booking", ""),
                        kd.get("last_date_apply", ""),
                        kd.get("entrance_exam", ""),
                        kd.get("display_result", ""),
                        kd.get("last_date_adm", ""),
                        kd.get("scholarship_exam", ""),
                        kd.get("scholarship_result", ""),
                    ])
                    print(f"[{count}/{len(codes_names)}] {code}: Schedule {active_sched}, {len(summary.splitlines())} dates captured")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}", "", "", "", "", "", "", "", "", ""])
                    print(f"[{count}/{len(codes_names)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live important dates records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
