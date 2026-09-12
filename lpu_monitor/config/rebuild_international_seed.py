import xlrd
import csv
import glob
import os


PROGRAMME_LIST_FOLDER = "lpu_monitor/incoming/programme_list"
OUTPUT_CSV = "lpu_monitor/config/international_seed.csv"


def find_single_file(folder, pattern="*.xls"):
    files = glob.glob(os.path.join(folder, pattern))

    if len(files) != 1:
        raise ValueError(
            f"Expected exactly 1 file in {folder}, found {len(files)}"
        )

    return files[0]


def clean(value):
    if value is None:
        return ""

    return " ".join(
        str(value)
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
        .split()
    ).strip()


def build(
    proglist_path=None,
    out_csv=OUTPUT_CSV
):
    if proglist_path is None:
        proglist_path = find_single_file(
            PROGRAMME_LIST_FOLDER
        )

    print("Reading:")
    print(os.path.abspath(proglist_path))

    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)

    header = sheet.row_values(0)

    # ---------------------------------------------------------
    # Find required columns
    # ---------------------------------------------------------

    code_col = None
    name_col = None
    intl_col = None

    for i, value in enumerate(header):

        h = clean(value).lower()

        if "programme code" in h:
            code_col = i

        if "programme name" in h:
            name_col = i

        # IMPORTANT:
        # We specifically want:
        # "To be offered to International Applicants 2026"
        if "to be offered to international applicants" in h:
            intl_col = i

    if code_col is None:
        raise ValueError(
            "Programme Code column not found."
        )

    if name_col is None:
        raise ValueError(
            "Programme Name column not found."
        )

    if intl_col is None:
        raise ValueError(
            "To be offered to International Applicants column not found."
        )

    print()
    print("Columns found:")
    print("Programme Code :", code_col)
    print("Programme Name :", name_col)
    print("International  :", intl_col)

    # ---------------------------------------------------------
    # Build international seed
    # ---------------------------------------------------------

    rows = []

    for r in range(1, sheet.nrows):

        code = clean(
            sheet.cell_value(r, code_col)
        )

        name = clean(
            sheet.cell_value(r, name_col)
        )

        intl = clean(
            sheet.cell_value(r, intl_col)
        )

        if not code:
            continue

        # Non-PhD only
        if "ph.d" in code.lower():
            continue

        if "ph.d" in name.lower():
            continue

        # International applicants = Yes
        if intl.lower().startswith("yes"):

            rows.append(
                {
                    "official_code": code,
                    "programme_name": name
                }
            )

    # ---------------------------------------------------------
    # Write CSV
    # EXACTLY TWO COLUMNS
    # ---------------------------------------------------------

    output_folder = os.path.dirname(out_csv)

    if output_folder:
        os.makedirs(
            output_folder,
            exist_ok=True
        )

    with open(
        out_csv,
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=[
                "official_code",
                "programme_name"
            ]
        )

        writer.writeheader()
        writer.writerows(rows)

    print()
    print(
        f"Wrote {len(rows)} international-eligible codes "
        f"-> {out_csv}"
    )

    return len(rows)


if __name__ == "__main__":
    build()