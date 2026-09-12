"""
check_highlights_v2.py

Fixed extraction: strips ALL HTML tags first, then checks for the plain
text titles directly -- no longer depends on any specific tag structure
(<strong> alone, <strong><em>, or anything else). Confirmed bug: the
previous version's regex required <strong><em> together, missing P371-NNAF
(and likely others) which use <strong> alone.
"""
import asyncio, json, re
from lpu_monitor.api_clients.browser_session import BrowserSession

CODES = ["P132", "P102", "P371-NNAF"]

def strip_html(text):
    return re.sub(r"<[^>]+>", "", text)

async def main():
    async with BrowserSession(headless=True) as session:
        for code in CODES:
            data = await session.call_api(
                url="https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramFeeDetail",
                method="POST", body={"officialCode": code},
            )
            text = strip_html(json.dumps(data))
            has_intl = "International Exposure Option" in text
            has_exchange = "Semester Exchange" in text
            print(f"{code}: International Exposure = {has_intl}, Semester Exchange = {has_exchange}")

asyncio.run(main())