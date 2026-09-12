import asyncio, csv, os
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.fee_validation import load_fee_expected, check_fee
import csv

def load_fee_exceptions(path="lpu_monitor/config/fee_exceptions.csv"):
    exceptions = set()
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            exceptions.add(row["official_code"])
    return exceptions
async def run_batch(codes_with_ids, out_path="fee_check_results.csv"):
    fee_expected = load_fee_expected()
    async with BrowserSession(headless=True) as session:
        with open(out_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            fee_exceptions = load_fee_exceptions()
            for code, program_id in codes_with_ids:
                if code in fee_exceptions:
                    writer.writerow([code, "SKIPPED"])
                    continue
                try:
                    cuet = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/ScholarshipCUET?id={program_id}",
                        method="GET",
                    )
                    result = check_fee(code, cuet, fee_expected)
                    writer.writerow([code, result])
                    print(code, "->", result)
                except Exception as e:
                    writer.writerow([code, f"ERROR: {e}"])
                    print(code, "FAILED:", e)

if __name__ == "__main__":
    codes_with_ids = []
    with open("lpu_monitor/config/programme_seed.csv", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            codes_with_ids.append((row["OfficialCode"], row["ProgramId"]))

    fee_expected_codes = set(load_fee_expected().keys())
    codes_with_ids = [(c, pid) for c, pid in codes_with_ids if c in fee_expected_codes]

    print(f"Total codes to test: {len(codes_with_ids)}")

    BATCH_START, BATCH_END = 0,len(codes_with_ids)-1# change this each run: 0-50, 50-100, 100-150, ...

    if BATCH_START == 0 and os.path.exists("fee_check_results.csv"):
        os.remove("fee_check_results.csv")

    asyncio.run(run_batch(codes_with_ids[BATCH_START:BATCH_END]))