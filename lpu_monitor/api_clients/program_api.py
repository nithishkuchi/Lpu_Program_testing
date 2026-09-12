"""
lpu_monitor/api_clients/program_api.py
Fetches programme details from LPU's public API.
"""

from typing import Any, Dict


GET_PROGRAM_DETAILS_URL = (
    "https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramDetails"
)

GET_PLACEMENTS_RANKINGS_URL = (
    "https://webapi.lpu.in/webProgrammes/api/"
    "ProgramSearch/GetProgrammePlacementsRankingsDetails"
)


async def fetch_program_details(
    session,
    official_code: str,
) -> Dict[str, Any]:
    result = await session.call_api(
        url=GET_PROGRAM_DETAILS_URL,
        method="POST",
        body={
            "officialCode": official_code,
            "curriculumSkipped": "0",
        },
    )

    if isinstance(result, dict):
        return next(iter(result.values()), {})

    if isinstance(result, list) and result:
        return result[0]

    return {}


async def fetch_eligibility_indian(
    session,
    official_code: str,
) -> str:
    """Returns the raw eligibilityindian HTML field from the live page."""

    details = await fetch_program_details(
        session,
        official_code,
    )

    return details.get("eligibilityindian", "") or ""


async def fetch_achievements_records(
    session,
    official_code: str,
):
    """
    Fetch placement/research/ranking information
    for one programme.

    This is the confirmed working LPU endpoint:

    GetProgrammePlacementsRankingsDetails/{official_code}
    """

    result = await session.call_api(
        url=f"{GET_PLACEMENTS_RANKINGS_URL}/{official_code}",
        method="GET",
    )

    if isinstance(result, list):
        return result

    return []