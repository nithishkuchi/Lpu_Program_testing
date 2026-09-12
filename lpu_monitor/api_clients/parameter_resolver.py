"""
parameter_resolver.py

Thin wrapper around GetParameterFromUrl -- converts a programme URL into its
OfficialCode and internal identifiers.

CONFIRMED from a real captured request (network log, 2026-07-24): this endpoint
is GET, with the programme's *relative path* (not the full URL) as a query
parameter, e.g.:

    GET https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetParameterFromUrl
        ?url=/engineering/b-tech-computer-science

Note the relative path starts with "/engineering/..." -- NOT "/programmes/engineering/...".
Strip the "/programmes" prefix (and domain) from a full lpu.in URL before calling this.
"""

from .browser_session import BrowserSession

PARAMETER_RESOLVER_URL = "https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetParameterFromUrl"


def _to_relative_path(programme_url: str) -> str:
    """
    Converts a full programme URL (or a "/programmes/..." path) into the
    relative-path format this endpoint expects, e.g.:
        "https://www.lpu.in/programmes/engineering/b-tech-computer-science"
        -> "/engineering/b-tech-computer-science"
    """
    path = programme_url
    for prefix in ("https://www.lpu.in", "https://lpu.in", "http://www.lpu.in"):
        if path.startswith(prefix):
            path = path[len(prefix):]
            break
    if path.startswith("/programmes"):
        path = path[len("/programmes"):]
    return path


async def resolve_official_code(session: BrowserSession, programme_url: str) -> list[dict]:
    """
    Returns the resolver JSON, e.g.:
    [{"OfficialCode": "P132", "ProgramId": 570, "DisciplineId": 13, ...}]
    """
    relative_path = _to_relative_path(programme_url)
    return await session.call_api(
        url=f"{PARAMETER_RESOLVER_URL}?url={relative_path}",
        method="GET",
        body=None,
    )

