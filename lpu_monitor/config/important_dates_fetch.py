"""
important_dates_fetch.py

Thin wrapper around GetNESTImportantDates. Unlike ScholarshipCUET, this
endpoint returns the list directly (confirmed from 3 real captures) -- no
list-wrapping to undo here.
"""
from lpu_monitor.api_clients.browser_session import BrowserSession

URL = "https://webapi.lpu.in/lpunest/api/LPUNEST/GetNESTImportantDates"


async def fetch_important_dates(session: BrowserSession, official_code: str):
    return await session.call_api(url=URL, method="POST", body={"officialCode": official_code})