import asyncio
import json
import csv

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.annexure_check import (
    load_lookup,
    check_programme,
    ANNEXURE_TITLE_MATCH,
)

OUTPUT_FILE = "annexure_results.csv"


async def run_validation(codes):
    lookup = load_lookup()

    async with BrowserSession(headless=True) as session:

        with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["ProgrammeCode", "Mismatches"])

            for code in codes:
                if code == "P2L2":
                    writer.writerow([code, "SKIPPED: no live data, see code_exceptions.csv"])
                    continue
                print(f"Checking {code}...")

                try:
                    tabs = await session.call_api(
                        url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramScholarshipTabs?id={code}&filter=1",
                        method="GET",
                    )

                    result = check_programme(
                        code,
                        tabs,
                        lookup,
                        ANNEXURE_TITLE_MATCH,
                    )

                    writer.writerow(
                        [
                            code,
                            json.dumps(result["mismatches"]),
                        ]
                    )

                    print(code, result)

                except Exception as e:

                    print(f"{code} FAILED")

                    writer.writerow(
                        [
                            code,
                            f"ERROR : {e}",
                        ]
                    )


if __name__ == "__main__":

    lookup = load_lookup()

    all_codes = list(lookup.keys())

    test_codes = list(lookup.keys())

    asyncio.run(run_validation(test_codes))