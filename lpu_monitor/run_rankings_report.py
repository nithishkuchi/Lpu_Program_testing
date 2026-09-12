"""
run_rankings_report.py

Records: Ranking information for ALL programmes.
"""

import asyncio
import glob
import os
import re

import pandas as pd
import xlrd

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.program_api import fetch_achievements_records


OUT_XLSX = "records_rankings.xlsx"
PROGRAMME_LIST_FOLDER = "lpu_monitor/incoming/programme_list"


def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))

    if len(files) != 1:
        raise ValueError(
            f"Expected exactly 1 file in {folder}, found {len(files)}"
        )

    return files[0]


def get_all_programmes(proglist_path):

    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)

    header = [
        str(value).strip()
        for value in sheet.row_values(0)
    ]

    code_col = next(
        i
        for i, h in enumerate(header)
        if "programme code" in h.lower()
    )

    name_col = next(
        i
        for i, h in enumerate(header)
        if "programme name" in h.lower()
    )

    programmes = []

    for row in range(1, sheet.nrows):

        code = str(
            sheet.cell_value(row, code_col)
        ).strip()

        name = str(
            sheet.cell_value(row, name_col)
        ).strip()

        if not code:
            continue

        programmes.append(
            {
                "code": code,
                "name": name,
            }
        )

    return programmes


def clean_html(text):

    if not text:
        return ""

    text = re.sub(
        r"<[^>]+>",
        " ",
        str(text),
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def parse_records(
    api_response,
    programme_code,
    programme_name,
):

    if not isinstance(api_response, list):
        return []

    records = []

    for item in api_response:

        if not isinstance(item, dict):
            continue

        category = clean_html(
            item.get(
                "headingText",
                "",
            )
        )

        details = item.get(
            "details",
            "",
        ) or ""

        display_order = item.get(
            "displayOrder",
            999,
        )

        if not category:
            continue

        records.append(
            {
                "Programme Name": programme_name,
                "OfficialCode": programme_code,
                "Category": category,
                "Details": clean_html(details),
                "Display Order": display_order,
            }
        )

    return records


async def main():

    print("=" * 70)
    print("RANKINGS RECORDS REPORT")
    print("=" * 70)

    proglist_path = find_single_file(
        PROGRAMME_LIST_FOLDER
    )

    programmes = get_all_programmes(
        proglist_path
    )

    print(
        f"Programme list : {proglist_path}"
    )

    print(
        f"Programmes     : {len(programmes)}"
    )

    print()

    all_records = []

    async with BrowserSession(headless=True) as session:

        for index, programme in enumerate(
            programmes,
            start=1,
        ):

            code = programme["code"]
            name = programme["name"]

            print(
                f"[{index}/{len(programmes)}] "
                f"{code} - {name}"
            )

            try:

                response = await fetch_achievements_records(
                    session,
                    code,
                )

                records = parse_records(
                    response,
                    code,
                    name,
                )

                ranking_records = [
                    record
                    for record in records
                    if "ranking"
                    in record["Category"].lower()
                ]

                all_records.extend(
                    ranking_records
                )

                print(
                    f"    Ranking records: "
                    f"{len(ranking_records)}"
                )

            except Exception as exc:

                print(
                    f"    FAILED: {exc}"
                )

    columns = [
        "Programme Name",
        "OfficialCode",
        "Category",
        "Details",
        "Display Order",
    ]

    df = pd.DataFrame(
        all_records,
        columns=columns,
    )

    if not df.empty:
        df = df.sort_values(
            by=[
                "Programme Name",
                "Display Order",
            ],
            na_position="last",
        )

    df.to_excel(
        OUT_XLSX,
        index=False,
        engine="openpyxl",
    )

    print()
    print("=" * 70)
    print("RANKINGS REPORT COMPLETE")
    print("=" * 70)
    print(
        f"Total ranking records : {len(df)}"
    )
    print(
        f"Output                : {os.path.abspath(OUT_XLSX)}"
    )


if __name__ == "__main__":
    asyncio.run(main())