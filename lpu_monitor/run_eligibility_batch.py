"""
run_eligibility_batch.py
Fetches live eligibility from LPU API and compares against expected Excel data.
"""
import asyncio
import csv
import glob
import os

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.program_api import fetch_eligibility_indian
from lpu_monitor.config.eligibility_validation import (
    compare_eligibility,
    load_category_aliases,
)

INCOMING_ELIGIBILITY = "lpu_monitor/incoming/eligibility_excel"
ELIGIBILITY_EXPECTED = "lpu_monitor/config/eligibility_expected.csv"
OUTPUT_CSV = "eligibility_check_results.csv"


def load_expected_codes(path):
    codes = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            code = row["official_code"].strip()
            indian = row.get("eligibility_indian", "").strip()
            if not indian:
                continue
            codes[code] = {
                "indian": indian,
                "relaxation": row.get("relaxation", "").strip(),
                "equivalent": row.get("equivalent_qualification", "").strip(),
            }
    return codes


def find_eligibility_excel():
    files = glob.glob(os.path.join(INCOMING_ELIGIBILITY, "*.xls*"))
    if len(files) == 1:
        return files[0]
    if not files:
        return None
    return files[0]


async def run_batch(code_list, expected_dict, out_path, category_aliases):
    match_count = 0
    diff_count = 0
    error_count = 0

    async with BrowserSession(headless=True) as session:
        with open(out_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for code in code_list:
                try:
                    live_html = await fetch_eligibility_indian(session, code)
                except Exception as e:
                    error_count += 1
                    writer.writerow([code, "ERROR", str(e)])
                    print(f"  {code} ERROR: {e}")
                    continue

                expected = expected_dict[code]
                diff = compare_eligibility(
                    expected["indian"],
                    expected["relaxation"],
                    live_html,
                    category_aliases,
                    expected_equivalent=expected.get("equivalent", ""),
                )

                if not diff:
                    match_count += 1
                    writer.writerow([code, "MATCH", ""])
                    print(f"  {code} -> MATCH")
                else:
                    diff_count += 1
                    writer.writerow([code, "DIFF", diff])
                    print(f"  {code} -> DIFF")
                    for line in diff.split("\n"):
                        if line.strip():
                            print(f"      {line.strip()}")

    return match_count, diff_count, error_count


if __name__ == "__main__":
    expected = load_expected_codes(ELIGIBILITY_EXPECTED)
    code_list = sorted(expected.keys())
    print(f"Total codes with eligibility to test: {len(code_list)}")

    xlsx_path = find_eligibility_excel()
    category_aliases = load_category_aliases(xlsx_path) if xlsx_path else {}
    if category_aliases:
        print(f"Loaded {len(category_aliases)} category aliases")

    # ── Batch slicing: change these two lines each run ──
    BATCH_START = 0
    BATCH_END = len(code_list)
    # Examples: 0-50, 50-100, 100-150, ... or 0 to len() for all at once
    # ────────────────────────────────────────────────────

    batch = code_list[BATCH_START:BATCH_END]
    print(f"Running batch [{BATCH_START}:{BATCH_END}] — {len(batch)} codes\n")

    if BATCH_START == 0 and os.path.exists(OUTPUT_CSV):
        os.remove(OUTPUT_CSV)

    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(["official_code", "status", "diff"])

    matches, diffs, errors = asyncio.run(
        run_batch(batch, expected, OUTPUT_CSV, category_aliases)
    )

    print(f"\n{'='*50}")
    print(f"  MATCH: {matches}  |  DIFF: {diffs}  |  ERROR: {errors}")
    print(f"{'='*50}")