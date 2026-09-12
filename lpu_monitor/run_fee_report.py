"""run_fee_report.py -- Records: Live Fee (non-PhD) data."""
import asyncio
import csv
import glob
import os
import re
import sys
from openpyxl import Workbook
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.fee_validation import unwrap_cuet_response, determine_active_phase

OUT_XLSX = "records_fee.xlsx"


def find_single_file(folder, patterns=("*.xls", "*.xlsx")):
    for pat in patterns:
        files = glob.glob(os.path.join(folder, pat))
        if files:
            return files[0]
    raise ValueError(f"Expected 1 file matching {patterns} in {folder}")


def get_non_phd_codes(proglist_path):
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
        code = str(sheet.cell_value(r, code_col)).strip()
        name = str(sheet.cell_value(r, name_col)).strip()
        if code and "ph.d" not in code.lower() and "ph.d" not in name.lower():
            out.append((code, name))
    return out


def load_program_id_lookup(seed_path="lpu_monitor/config/programme_seed.csv"):
    lookup = {}
    if os.path.exists(seed_path):
        with open(seed_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                lookup[row["OfficialCode"].strip()] = row["ProgramId"].strip()
    return lookup


async def run(start_idx=0, end_idx=None):
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    codes = get_non_phd_codes(proglist)
    id_lookup = load_program_id_lookup()

    if end_idx is not None:
        codes = codes[start_idx:end_idx]
    else:
        codes = codes[start_idx:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Fee Records"

    headers = [
        "Programme Name",
        "OfficialCode",
        "Active Phase",
        "Active Phase Cat A (98%)",
        "Active Phase Cat B (90%)",
        "Active Phase Cat C (80%)",
        "Active Phase Cat D (70%)",
        "Phase 1 Cat A", "Phase 1 Cat B", "Phase 1 Cat C", "Phase 1 Cat D",
        "Phase 2 Cat A", "Phase 2 Cat B", "Phase 2 Cat C", "Phase 2 Cat D",
        "Phase 3 Cat A", "Phase 3 Cat B", "Phase 3 Cat C", "Phase 3 Cat D",
        "Phase 4 Cat A", "Phase 4 Cat B", "Phase 4 Cat C", "Phase 4 Cat D",
        "Phase 5 Cat A", "Phase 5 Cat B", "Phase 5 Cat C", "Phase 5 Cat D",
        "Open Fee Tab",
        "Qualification",
        "Criteria 98%",
        "Criteria 90%",
        "Criteria 80%",
        "Criteria 70%",
    ]
    ws.append(headers)

    print(f"Total non-PhD programmes to fetch live fees for: {len(codes)}")

    count = 0
    async with BrowserSession(headless=True) as session:
        try:
            for code, name in codes:
                count += 1
                program_id = id_lookup.get(code)
                if not program_id:
                    ws.append([name, code, "NO_PROGRAM_ID_FOUND"])
                    print(f"[{count}/{len(codes)}] {code}: No ProgramId found")
                    continue

                try:
                    cuet_raw = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/ScholarshipCUET?id={program_id}",
                        method="GET",
                    )
                    cuet = unwrap_cuet_response(cuet_raw)
                    active_phase = determine_active_phase(cuet)

                    def get_val(key):
                        v = cuet.get(key)
                        return v if v is not None else ""

                    act_a = get_val(f"p98{active_phase}") if active_phase else ""
                    act_b = get_val(f"p90{active_phase}") if active_phase else ""
                    act_c = get_val(f"p80{active_phase}") if active_phase else ""
                    act_d = get_val(f"p70{active_phase}") if active_phase else ""

                    row = [
                        name,
                        code,
                        f"Phase {active_phase}" if active_phase else "None",
                        act_a, act_b, act_c, act_d,
                        get_val("p981"), get_val("p901"), get_val("p801"), get_val("p701"),
                        get_val("p982"), get_val("p902"), get_val("p802"), get_val("p702"),
                        get_val("p983"), get_val("p903"), get_val("p803"), get_val("p703"),
                        get_val("p984"), get_val("p904"), get_val("p804"), get_val("p704"),
                        get_val("p985"), get_val("p905"), get_val("p805"), get_val("p705"),
                        get_val("openFeeTab"),
                        get_val("qualification"),
                        get_val("criteria98"),
                        get_val("criteria90"),
                        get_val("criteria80"),
                        get_val("criteria70"),
                    ]
                    ws.append(row)
                    print(f"[{count}/{len(codes)}] {code}: Active Phase {active_phase}, A:{act_a} B:{act_b} C:{act_c}")
                except Exception as e:
                    ws.append([name, code, f"ERROR: {e}"])
                    print(f"[{count}/{len(codes)}] {code} FAILED: {e}")

                if count % 20 == 0:
                    wb.save(OUT_XLSX)
        finally:
            wb.save(OUT_XLSX)
            print(f"Saved live fee records to {OUT_XLSX}")


if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    asyncio.run(run(start, end))
