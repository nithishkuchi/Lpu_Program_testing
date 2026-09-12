import asyncio, csv, os
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.fee_validation import check_fee_phd, load_phd_fee_expected


async def run_phd_batch(codes_with_ids_modes, phd_fee_expected, out_path="phd_fee_results.csv"):
    async with BrowserSession(headless=True) as session:
        with open(out_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for code, program_id, mode in codes_with_ids_modes:
                try:
                    cuet = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/ScholarshipCUET?id={program_id}",
                        method="GET",
                    )
                    result = check_fee_phd(code, mode, cuet, phd_fee_expected)
                    writer.writerow([code, result])
                    print(code, "->", result)
                except Exception as e:
                    writer.writerow([code, f"ERROR: {e}"])
                    print(code, "FAILED:", e)


if __name__ == "__main__":
    phd_fee_expected = load_phd_fee_expected()

    entries = []
    with open("lpu_monitor/config/phd_codes.csv", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            entries.append((row["official_code"], row["program_id"], row["mode"]))

    print(f"Total PhD codes to test: {len(entries)}")

    BATCH_START, BATCH_END = 0, 135  # change each run: 0-50, 50-100, 100-135

    if BATCH_START == 0 and os.path.exists("phd_fee_results.csv"):
        os.remove("phd_fee_results.csv")

    asyncio.run(run_phd_batch(entries[BATCH_START:BATCH_END], phd_fee_expected))