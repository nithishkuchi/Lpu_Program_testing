"""
demo_fetch.py

Small standalone script to validate the browser-session approach end-to-end
against a couple of known officialCodes -- run this BEFORE wiring the browser
session into the full monitoring pipeline, to confirm the Cloudflare-clearing
approach actually works from your machine/network.

One-time setup:
    pip install playwright
    playwright install chromium
    # If you want the harder-to-fingerprint real-Chrome path (recommended):
    playwright install chrome

Run with:
    python -m lpu_monitor.demo_fetch
"""

import asyncio
import logging
import json

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.fee_detail_client import fetch_fee_detail
from lpu_monitor.config.seed_loader import load_programme_seed

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")

# One representative code per category, drawn from the verified master programme
# list (programme_seed.csv) -- picked specifically to stress-test different shapes
# of fee structure before the normalizer/field_mapping.yaml is finalized.
TEST_OFFICIAL_CODES = {
    "P102": "Diploma, After 10th (CSE)",
    "P1D2": "Ph.D Computer Science Full Time",
    "P102-NNP": "Diploma specialization variant (AI and ML)",
    "P102-L": "Diploma, Lateral Entry (After ITI)",
    "P132": "B.Tech regular UG (baseline, already confirmed working)",
    "P132-L": "B.Tech, Lateral Entry (After 3yr Diploma)",
    "P132-HNF": "B.Tech Honours variant",
    "P5K0-NNKA": "Pharm.D -- heavy Indian/International fee split, good edge case",
    "P371-NNAF": "MBA -- PG, After Graduation",
    "P164-HN7": "MCA Honours (AI and ML) -- PG Honours",
    "P490-HH": "B.A., LL.B. (Hons.) integrated law -- unusual 'no exit option' programme",
    
}


def check_for_duplicate_entries(data, code: str) -> None:
    """
    The P132 test showed the response as a JSON array. Before trusting entry [0]
    as "the" record for a programme, confirm whether multiple array entries are
    truly identical or subtly differ (see conversation notes) -- this determines
    whether the normalizer can safely ignore entries [1:] or must inspect all of
    them.
    """
    if not isinstance(data, list) or len(data) <= 1:
        print(f"  [{code}] response is not a multi-entry array -- nothing to compare.")
        return

    first = data[0]
    all_identical = all(entry == first for entry in data[1:])
    print(f"  [{code}] response has {len(data)} array entries. All identical to entry[0]? {all_identical}")

    if not all_identical:
        # Show exactly which keys differ across entries so this can be extended
        # into the normalizer's field_mapping.yaml with confidence.
        all_keys = set()
        for entry in data:
            all_keys.update(entry.keys())
        differing_keys = [
            k for k in sorted(all_keys)
            if len({entry.get(k) for entry in data}) > 1
        ]
        print(f"    Differing keys across entries: {differing_keys}")


async def main() -> None:
    async with BrowserSession(headless=True) as session:
        seed = load_programme_seed()
        for code, label in TEST_OFFICIAL_CODES.items():
            program_id = seed[code]["ProgramId"]
            print(f"\n=== {code} ({label}) ===")
            try:
                data = await fetch_fee_detail(session, code)
                scholarship_slab = await session.call_api(url=f"https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramScholarshipSlab?id={program_id}",method="GET",)
                with open(f"lpu_monitor/tests/fixtures/scholarship_slab_{code}.json","w",encoding="utf-8",) as f:json.dump(scholarship_slab, f, indent=2)
                print(f"SUCCESS -- received {len(str(data))} characters of JSON")
                check_for_duplicate_entries(data, code)
            except RuntimeError as exc:
                print(f"FAILED: {exc}")


if __name__ == "__main__":
    asyncio.run(main())
