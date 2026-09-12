"""
run_student_achievements_report.py

Records: Student achievements for ALL programmes.
"""

import asyncio
import glob
import os
import re

import pandas as pd
import xlrd

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.api_clients.program_api import fetch_achievements_records


OUT_XLSX = "records_student_achievements.xlsx"
PROGRAMME_LIST_FOLDER = "lpu_monitor/incoming/programme_list"


def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))

    if len(files) != 1:
        raise ValueError(
            f"Expected exactly 1 file in {folder}, found {len(files)}"
        )

    return files[0]


def get_all_programmes(proglist_path):
    """
    Read all programme codes and programme names
    from the programme list.
    """

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
    """
    Remove HTML tags and normalize whitespace.
    """

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


def extract_funding(details):
    """
    Extract funding amounts such as:

        ₹2 crore
        ₹5 lakh
        2 crore
        5 lakh

    Returns the numeric amount and unit separately.
    """

    text = clean_html(details)

    funding_match = re.search(
        r"₹?\s*(\d+(?:\.\d+)?)\s*(crore|lakh)",
        text,
        re.IGNORECASE,
    )

    if not funding_match:
        return "", ""

    amount = funding_match.group(1)
    unit = funding_match.group(2).lower()

    return amount, unit


def parse_student_achievements(
    api_response,
    programme_code,
    programme_name,
):
    """
    Convert the raw API response into student-achievement records.

    If the API response is not a list, return an empty list.
    """

    if not isinstance(api_response, list):
        return []

    records = []

    for item in api_response:

        if not isinstance(item, dict):
            continue

        details = item.get(
            "details",
            "",
        ) or ""

        heading = item.get(
            "headingText",
            "Unknown",
        ) or "Unknown"

        heading = clean_html(heading)

        display_order = item.get(
            "displayOrder",
            999,
        )

        amount, unit = extract_funding(
            details
        )

        records.append(
            {
                "Programme Name": programme_name,
                "OfficialCode": programme_code,
                "Student/Team": heading,
                "Details": clean_html(details),
                "Funding Amount": amount if amount else "N/A",
                "Funding Unit": unit if unit else "N/A",
                "Display Order": display_order,
            }
        )

    return records


def build_student_achievements_excel(
    records,
    out_path=OUT_XLSX,
):

    columns = [
        "Programme Name",
        "OfficialCode",
        "Student/Team",
        "Details",
        "Funding Amount",
        "Funding Unit",
        "Display Order",
    ]

    df = pd.DataFrame(
        records,
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
        out_path,
        index=False,
        engine="openpyxl",
    )

    print(
        f"Built student achievements: "
        f"{len(df)} records -> {out_path}"
    )


async def main():

    print("=" * 70)
    print("STUDENT ACHIEVEMENTS RECORDS REPORT")
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

                records = parse_student_achievements(
                    response,
                    code,
                    name,
                )

                all_records.extend(
                    records
                )

                print(
                    f"    Student achievement records: "
                    f"{len(records)}"
                )

            except Exception as exc:

                print(
                    f"    FAILED: {exc}"
                )

    build_student_achievements_excel(
        all_records
    )

    print()
    print("=" * 70)
    print("STUDENT ACHIEVEMENTS REPORT COMPLETE")
    print("=" * 70)
    print(
        f"Total records : {len(all_records)}"
    )
    print(
        f"Output        : "
        f"{os.path.abspath(OUT_XLSX)}"
    )


if __name__ == "__main__":
    asyncio.run(main())