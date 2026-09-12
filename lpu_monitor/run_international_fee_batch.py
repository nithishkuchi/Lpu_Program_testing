"""Batch runner -- only codes in international_seed.csv (refreshed via
rebuild_international_seed.py)."""
import asyncio, csv, os
from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.international_fee_fetch import fetch_international_fee
from lpu_monitor.config.international_fee_validation import load_expected, check_international_fee

OUT_PATH = "international_fee_results.csv"

def get_seed_codes(path="lpu_monitor/config/international_seed.csv"):
    with open(path, encoding="utf-8") as f:
        return [row["official_code"] for row in csv.DictReader(f)]

async def run_batch():
    expected = load_expected()
    codes = get_seed_codes()
    async with BrowserSession(headless=True) as session:
        with open(OUT_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for code in codes:
                try:
                    live = await fetch_international_fee(session, code)
                    result = check_international_fee(code, live, expected)
                    writer.writerow([code, result])
                    print(code, "->", result)
                except Exception as e:
                    writer.writerow([code, f"ERROR: {e}"])
                    print(code, "FAILED:", e)

if __name__ == "__main__":
    if os.path.exists(OUT_PATH):
        os.remove(OUT_PATH)
    asyncio.run(run_batch())