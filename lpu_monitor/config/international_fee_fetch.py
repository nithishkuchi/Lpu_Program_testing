"""Wraps InternationalFees/get-by-code -- confirmed working script/endpoint."""
from lpu_monitor.api_clients.browser_session import BrowserSession

URL = "https://webapi.lpu.in/InternationalFee/api/InternationalFees/get-by-code"

async def fetch_international_fee(session: BrowserSession, official_code: str):
    return await session.call_api_via_request_context(
        url=URL, method="POST", body={"OfficialCode": official_code},
    )