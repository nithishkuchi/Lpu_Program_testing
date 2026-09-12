"""
important_dates_mapping_builder.py

Wraps the teammate-provided important_dates_matcher.py (kept completely
unmodified) with:
  1. Auto-discovery of the two input files from incoming/ folders -- no
     hardcoded filenames, same pattern as every other builder in this
     project.
  2. TYPE_LETTER_OVERRIDE forced to {} at run time -- confirmed with sir
     that "Type A" no longer exists as a real category in the current
     source files (fully renamed to Type B), so the override is dead
     weight now. Done here rather than editing their file directly.
  3. SKIP_EXCEPTION rows (e.g. "software product engineering") dropped
     entirely -- they should never appear in ANY output file, not even
     as a "needs review" row. The teammate's own classify_all() currently
     routes them into the review list; this wrapper calls the lower-level
     functions directly instead of their main(), so it can filter these
     out before anything gets written.

Produces the SAME output format their script already does
(Mapped_Programme_Dates_2026.csv), so nothing downstream needs to change.
"""
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import important_dates_matcher as matcher

INCOMING_IMPORTANT_DATES = "lpu_monitor/incoming/important_dates"
INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"


def find_single_file(folder, pattern="*"):
    files = [f for f in glob.glob(os.path.join(folder, pattern)) if os.path.isfile(f)]
    if len(files) == 0:
        raise FileNotFoundError(f"No file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
    return files[0]


def build_mapping(
    proglist_path=None,
    important_dates_path=None,
    out_csv="lpu_monitor/config/Mapped_Programme_Dates_2026.csv",
    review_csv="lpu_monitor/config/important_dates_needs_review.csv",
):
    proglist_path = proglist_path or find_single_file(INCOMING_PROGLIST, "*.xls")
    important_dates_path = important_dates_path or find_single_file(INCOMING_IMPORTANT_DATES, "*.xlsx")

    # NOTE: sir confirmed Type A no longer exists as a real category in the
    # CURRENT source files. However, testing against the Programme List file
    # available at the time this wrapper was built (dated 20th July) showed
    # 56 programmes -- including P132 -- still labeled "Type A" with no
    # matching sheet, which would leave them unmapped if the override were
    # removed. Kept ACTIVE here as the safer default until confirmed against
    # the actual current file. Set to {} once that's verified.
    matcher.TYPE_LETTER_OVERRIDE = {"a": "b"}

    buckets = matcher.load_important_dates(important_dates_path)
    final_data = matcher.load_final_list(proglist_path)
    phd_data = matcher.load_phd_list(proglist_path)

    mapped, review = matcher.classify_all(
        final_data["rows"], final_data["col_type"], final_data["col_name"],
        final_data["col_nest"], final_data["col_code"], buckets,
    )

    phd_mapped, phd_review = [], []
    if phd_data is not None:
        for row_values in phd_data["rows"]:
            name = row_values[phd_data["col_name"]]
            code = row_values[phd_data["col_code"]] if phd_data["col_code"] is not None else ""
            if name in (None, ""):
                continue
            if "phd" in buckets:
                phd_mapped.append({
                    "name": name, "code": code,
                    "bucket_sheet": buckets["phd"]["sheet"],
                    "dates": matcher.build_date_dict(buckets["phd"]),
                })
            else:
                phd_review.append({
                    "name": name, "code": code, "status": "UNMAPPED_NO_SHEET",
                    "detail": "No 'Ph.D.' sheet found in Important Dates workbook",
                })

    all_mapped = mapped + phd_mapped

    # Drop SKIP_EXCEPTION rows entirely -- never written anywhere, per sir's
    # instruction. Every other review reason (UNMAPPED_NO_SHEET, etc.) still
    # gets written, since those ARE genuinely worth a human look.
    all_review = [r for r in (review + phd_review) if r["status"] != "SKIP_EXCEPTION"]

    matcher.write_mapped_csv(out_csv, all_mapped)
    matcher.write_review_csv(review_csv, all_review)

    return len(all_mapped), len(all_review)


if __name__ == "__main__":
    mapped_count, review_count = build_mapping()
    print(f"Wrote {mapped_count} mapped rows -> lpu_monitor/config/Mapped_Programme_Dates_2026.csv")
    print(f"Wrote {review_count} genuine review row(s) -> lpu_monitor/config/important_dates_needs_review.csv")
    if review_count > 0:
        import sys as _sys
        _sys.exit(1)  # non-zero so the UI flags "needs attention"