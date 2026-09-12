# # """
# # run_important_dates_batch.py

# # The actual fetch-and-compare VALIDATION step -- same role as run_fee_batch.py
# # and run_phd_fee_batch.py, just for Important Dates. Reads the mapping CSV
# # already built by important_dates_mapping_builder.py (a separate "Setup"
# # step, same relationship fee_row_mapping_builder.py has to run_fee_batch.py).
# # """
# # import asyncio
# # import csv
# # import os
# # import xlrd

# # from lpu_monitor.api_clients.browser_session import BrowserSession
# # from lpu_monitor.config.important_dates_fetch import fetch_important_dates
# # from lpu_monitor.config.important_dates_validation import (
# #     load_mapped_dates, load_nest_test_codes, check_important_dates,
# # )

# # OUT_PATH = "important_dates_results.csv"


# # def get_all_codes(proglist_path):
# #     wb = xlrd.open_workbook(proglist_path)
# #     sheet = wb.sheet_by_index(0)
# #     header = sheet.row_values(0)
# #     code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
# #     codes = []
# #     for r in range(1, sheet.nrows):
# #         code = sheet.cell_value(r, code_col)
# #         if code:
# #             codes.append(code)
# #     return codes


# # def find_single_file(folder, pattern="*.xls"):
# #     import glob
# #     files = glob.glob(os.path.join(folder, pattern))
# #     if len(files) != 1:
# #         raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}")
# #     return files[0]


# # async def run_batch():
# #     proglist_path = find_single_file("lpu_monitor/incoming/programme_list")
# #     mapped = load_mapped_dates("lpu_monitor/config/Mapped_Programme_Dates_2026.csv")
# #     nest_codes = load_nest_test_codes(proglist_path)
# #     codes = get_all_codes(proglist_path)

# #     async with BrowserSession(headless=True) as session:
# #         with open(OUT_PATH, "a", newline="", encoding="utf-8") as f:
# #             writer = csv.writer(f)
# #             for code in codes:
# #                 try:
# #                     live = await fetch_important_dates(session, code)
# #                     result = check_important_dates(code, live, mapped, nest_codes)
# #                     writer.writerow([code, result])
# #                     print(code, "->", result)
# #                 except Exception as e:
# #                     writer.writerow([code, f"ERROR: {e}"])
# #                     print(code, "FAILED:", e)


# # if __name__ == "__main__":
# #     if os.path.exists(OUT_PATH):
# #         os.remove(OUT_PATH)
# #     asyncio.run(run_batch())

# """
# run_important_dates_batch.py

# The actual fetch-and-compare VALIDATION step -- same role as run_fee_batch.py
# and run_phd_fee_batch.py, just for Important Dates. Reads the mapping CSV
# already built by important_dates_mapping_builder.py (a separate "Setup"
# step, same relationship fee_row_mapping_builder.py has to run_fee_batch.py).

# check_important_dates() no longer returns a status/dict -- it returns a
# plain "heading / expected / live" text report (or an empty string if
# there's nothing worth showing for that code). See important_dates_validation.py
# for why: this tool no longer tries to auto-judge right vs wrong, it just
# surfaces expected vs live side by side for a human to review.
# """
# import asyncio
# import csv
# import os
# import xlrd

# from lpu_monitor.api_clients.browser_session import BrowserSession
# from lpu_monitor.config.important_dates_fetch import fetch_important_dates
# from lpu_monitor.config.important_dates_validation import (
#     load_mapped_dates, load_nest_test_codes, check_important_dates,
# )

# OUT_PATH = "important_dates_results.csv"

# # These 5 codes are always excluded from this check entirely -- nothing is
# # fetched, computed, or written to the results file for them, regardless of
# # where they fall in the programme list. Confirmed list (not derived from
# # any rule -- these are specific codes given directly):
# ALWAYS_SKIP_CODES = {"P8G", "P8H", "P8J", "P132-NNL", "P2L2"}


# def get_all_codes(proglist_path):
#     wb = xlrd.open_workbook(proglist_path)
#     sheet = wb.sheet_by_index(0)
#     header = sheet.row_values(0)
#     code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
#     codes = []
#     for r in range(1, sheet.nrows):
#         code = sheet.cell_value(r, code_col)
#         if code:
#             codes.append(code)
#     return codes


# def find_single_file(folder, pattern="*.xls"):
#     import glob
#     files = glob.glob(os.path.join(folder, pattern))
#     if len(files) != 1:
#         raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}")
#     return files[0]


# async def run_batch():
#     proglist_path = find_single_file("lpu_monitor/incoming/programme_list")
#     mapped = load_mapped_dates("lpu_monitor/config/Mapped_Programme_Dates_2026.csv")
#     nest_codes = load_nest_test_codes(proglist_path)
#     codes = [c for c in get_all_codes(proglist_path) if c not in ALWAYS_SKIP_CODES]

#     async with BrowserSession(headless=True) as session:
#         with open(OUT_PATH, "a", newline="", encoding="utf-8") as f:
#             writer = csv.writer(f)
#             for code in codes:
#                 try:
#                     live = await fetch_important_dates(session, code)
#                     report = check_important_dates(code, live, mapped, nest_codes)
#                     writer.writerow([code, report])
#                     print(code, "->", report if report else "NO Errors")
#                 except Exception as e:
#                     writer.writerow([code, f"ERROR: {e}"])
#                     print(code, "FAILED:", e)


# if __name__ == "__main__":
#     if os.path.exists(OUT_PATH):
#         os.remove(OUT_PATH)
#     asyncio.run(run_batch())
"""
run_important_dates_batch.py

The actual fetch-and-compare VALIDATION step -- same role as run_fee_batch.py
and run_phd_fee_batch.py, just for Important Dates. Reads the mapping CSV
already built by important_dates_mapping_builder.py (a separate "Setup"
step, same relationship fee_row_mapping_builder.py has to run_fee_batch.py).

check_important_dates() no longer returns a status/dict -- it returns a
plain "heading / expected / live" text report (or an empty string if
there's nothing worth showing for that code). See important_dates_validation.py
for why: this tool no longer tries to auto-judge right vs wrong, it just
surfaces expected vs live side by side for a human to review.

NOTE: programme names are NOT written into this CSV. That's handled
centrally in app.py instead, by looking up official_code against the
already-uploaded Programme List Excel at display time -- one shared lookup
that works for every check's results generically, rather than needing each
run_xxxx.py batch script to duplicate that lookup itself.
"""
import asyncio
import csv
import os
import xlrd

from lpu_monitor.api_clients.browser_session import BrowserSession
from lpu_monitor.config.important_dates_fetch import fetch_important_dates
from lpu_monitor.config.important_dates_validation import (
    load_mapped_dates, load_nest_test_codes, check_important_dates,
)

OUT_PATH = "important_dates_results.csv"

# These 5 codes are always excluded from this check entirely -- nothing is
# fetched, computed, or written to the results file for them, regardless of
# where they fall in the programme list. Confirmed list (not derived from
# any rule -- these are specific codes given directly):
ALWAYS_SKIP_CODES = {"P8G", "P8H", "P8J", "P132-NNL", "P2L2"}


def get_all_codes(proglist_path):
    wb = xlrd.open_workbook(proglist_path)
    sheet = wb.sheet_by_index(0)
    header = sheet.row_values(0)
    code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
    codes = []
    for r in range(1, sheet.nrows):
        code = sheet.cell_value(r, code_col)
        if code:
            codes.append(code)
    return codes


def find_single_file(folder, pattern="*.xls"):
    import glob
    files = glob.glob(os.path.join(folder, pattern))
    if len(files) != 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}")
    return files[0]


async def run_batch():
    proglist_path = find_single_file("lpu_monitor/incoming/programme_list")
    mapped = load_mapped_dates("lpu_monitor/config/Mapped_Programme_Dates_2026.csv")
    nest_codes = load_nest_test_codes(proglist_path)
    codes = [c for c in get_all_codes(proglist_path) if c not in ALWAYS_SKIP_CODES]

    async with BrowserSession(headless=True) as session:
        with open(OUT_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for code in codes:
                try:
                    live = await fetch_important_dates(session, code)
                    report = check_important_dates(code, live, mapped, nest_codes)
                    writer.writerow([code, report])
                    print(code, "->", report if report else "(nothing to report)")
                except Exception as e:
                    writer.writerow([code, f"ERROR: {e}"])
                    print(code, "FAILED:", e)


if __name__ == "__main__":
    if os.path.exists(OUT_PATH):
        os.remove(OUT_PATH)
    asyncio.run(run_batch())
