"""
fee_detail_client.py

Thin wrapper around GetProgramFeeDetail. Deliberately contains no retry/session
logic of its own -- all of that lives in BrowserSession, so this file stays a
one-purpose module per the "each API is its own file" design.
"""

from .browser_session import BrowserSession

FEE_DETAIL_URL = "https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramFeeDetail"


async def fetch_fee_detail(session: BrowserSession, official_code: str) -> dict:
    """
    Returns the raw JSON payload for a given officialCode, e.g. "P132".
    Raises RuntimeError if the call ultimately fails after retries.
    """
    return await session.call_api(
        url=FEE_DETAIL_URL,
        method="POST",
        body={"officialCode": official_code},
    )


async def fetch_fee_detail_diagnostic(session: BrowserSession, official_code: str) -> dict:
    """
    Same call, but via Playwright's APIRequestContext instead of in-page fetch().
    Use this ONLY to compare against fetch_fee_detail() while diagnosing the
    "Failed to fetch" issue -- see call_api_via_request_context()'s docstring.
    """
    return await session.call_api_via_request_context(
        url=FEE_DETAIL_URL,
        method="POST",
        body={"officialCode": official_code},
    )
