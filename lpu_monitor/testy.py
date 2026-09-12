import asyncio
import json

from lpu_monitor.api_clients.browser_session import BrowserSession


PROGRAM_CODE = "P132"
PROGRAM_ID = 570


async def main():
    url = (
        "https://webapi.lpu.in/webProgrammes/api/ProgramSearch/"
        f"ScholarshipCUET?id={PROGRAM_ID}"
    )

    async with BrowserSession(headless=True) as session:
        response = await session.call_api(
            url=url,
            method="GET",
        )

        # Print the COMPLETE API response
        print("\n" + "=" * 80)
        print(f"PROGRAM: {PROGRAM_CODE}")
        print(f"PROGRAM ID: {PROGRAM_ID}")
        print("CUET API RESPONSE")
        print("=" * 80)

        print(json.dumps(response, indent=2, ensure_ascii=False))

        # Also save the COMPLETE response to a JSON file
        with open(
            "p132_cuet_response.json",
            "w",
            encoding="utf-8",
        ) as f:
            json.dump(response, f, indent=2, ensure_ascii=False)

        print("\n" + "=" * 80)
        print("Complete response saved to: p132_cuet_response.json")
        print("=" * 80)


if __name__ == "__main__":
    asyncio.run(main())