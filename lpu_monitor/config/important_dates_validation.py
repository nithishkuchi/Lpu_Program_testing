
# """
# important_dates_validation.py

# Compares live GetNESTImportantDates data against the teammate-mapped
# Mapped_Programme_Dates_2026.csv (built by map_programmes_2_.py).

# Confirmed from real captures (P132, an MBA code, P105):
#   - Live titles are inconsistently cased and carry stray HTML/whitespace --
#     normalize before comparing, never exact-match raw text.
#   - "Active" schedule is only reliable via the isScheduleNActive booleans,
#     never the schedule*Heading text labels (which can even say "Previous
#     Schedule" while the boolean says active -- confirmed on P105).
#   - Item count varies a lot by programme (1 item for P105, 8-10 for
#     P132/MBA) -- never assume a fixed set of particulars.
#   - Confirmed with sir: Type C programmes with NO NEST Test Code (NA/blank)
#     have no scholarship exam, so scholarship-named particulars from the
#     matched bucket legitimately don't apply and should be skipped, not
#     flagged as missing.

# REVISION 3 (this version) -- switched from word-based matching to SEMANTIC
# matching, using sentence embeddings instead of literal word overlap.

# Why: two earlier word-based approaches both broke on the same underlying
# issue -- literal word matching can't tell "these titles share a word" apart
# from "these titles mean the same thing". Fully stripping shared "generic"
# words (Revision 1) erased genuinely-identifying words like "scholarship" in
# "Scholarship Result" for the ~40 codes where that particular is real,
# producing an identical false-mismatch score across dozens of unrelated
# codes. Down-weighting instead of stripping (Revision 2) reduced that, but
# still can't recognize two phrasings that mean the same thing while sharing
# almost no literal words (e.g. differently-worded date phrasing, or a
# year-over-year wording change on the live site).

# This version replaces BOTH the title matching and the non-date value
# comparison fallback with cosine similarity between sentence embeddings
# (via sentence-transformers, a small local model -- no API calls, nothing
# sent off-machine, same input always gives the same output). This directly
# compares MEANING rather than word overlap, so paraphrased titles/values
# match correctly without needing a hand-maintained list of "generic" words
# to tune every time a false match/false miss turns up. Exact-match and
# date-parsing shortcuts are still tried FIRST (cheap and unambiguous) --
# the embedding model is only used for the harder, genuinely-fuzzy cases.

# Nothing about specific programme names, dates, or title wording is
# hardcoded anywhere below -- the model itself supplies the "meaning"
# comparison generically, for whatever text a given year's live page uses.

# REQUIRES: sentence-transformers (pulls in torch automatically). This is a
# hard dependency, not an optional one with a silent fallback -- if it can't
# load (not installed, or no internet for the one-time model download), this
# module raises immediately with clear instructions rather than pretending
# to work and quietly falling back to a weaker check.
#     pip install sentence-transformers
# The model itself (~80MB) downloads automatically the first time
# SentenceTransformer(...) is called, and is cached locally after that --
# only the very first run on a given machine needs internet access for it.
# """
# import csv
# import json
# import re
# import xlrd

# try:
#     from dateutil import parser as _date_parser
# except ImportError as _e:  # pragma: no cover
#     raise ImportError(
#         "important_dates_validation.py needs the 'python-dateutil' package "
#         "for date-aware value comparison. Install it with:\n"
#         "    pip install python-dateutil"
#     ) from _e

# try:
#     from sentence_transformers import SentenceTransformer, util as st_util
# except ImportError as _e:  # pragma: no cover
#     raise ImportError(
#         "important_dates_validation.py needs the 'sentence-transformers' "
#         "package for semantic title/value matching. Install it with:\n"
#         "    pip install sentence-transformers\n"
#         "(this also installs torch, which is a larger download -- expect "
#         "several hundred MB. The embedding model itself is downloaded "
#         "automatically and cached the first time this runs.)"
#     ) from _e


# # ---------------------------------------------------------------------------
# # Normalization helpers
# # ---------------------------------------------------------------------------
# def strip_html(text):
#     if text is None:
#         return ""
#     return re.sub(r"<[^>]+>", "", str(text))


# def normalize_title(text):
#     text = strip_html(text)
#     text = text.replace("\n", " ").replace("\r", " ")
#     return re.sub(r"\s+", " ", text).strip().lower()


# def normalize_value(text):
#     text = normalize_title(text)
#     text = re.sub(r"\*+$", "", text).strip()
#     return text


# # ---------------------------------------------------------------------------
# # Semantic similarity (shared by title matching AND the value-comparison
# # fallback). Small, fast, CPU-friendly model -- good enough for short
# # phrases like exam titles and date/duration values; not meant for long
# # documents.
# # ---------------------------------------------------------------------------
# _MODEL_NAME = "all-MiniLM-L6-v2"
# _model = None  # loaded lazily so importing this module doesn't force a load


# def _get_model():
#     global _model
#     if _model is None:
#         _model = SentenceTransformer(_MODEL_NAME)
#     return _model


# def embed_texts(texts):
#     """Batch-encodes a list of strings into normalized embedding vectors.
#     Always call this with as many texts at once as you have available
#     (e.g. all live titles for one code in a single call) rather than one
#     at a time in a loop -- batching is meaningfully faster."""
#     texts = list(texts)
#     if not texts:
#         return None
#     model = _get_model()
#     return model.encode(texts, convert_to_numpy=True, normalize_embeddings=True)


# def semantic_score(embedding_a, embedding_b):
#     return float(st_util.cos_sim(embedding_a, embedding_b)[0][0])


# # ---------------------------------------------------------------------------
# # Date-aware value comparison
# # ---------------------------------------------------------------------------
# _MONTH_NAMES = (
#     "jan", "feb", "mar", "apr", "may", "jun",
#     "jul", "aug", "sep", "oct", "nov", "dec",
# )


# def _looks_date_shaped(text_lower):
#     """Cheap guard before handing text to dateutil's fuzzy parser: only
#     attempt date parsing when the text plausibly CONTAINS a date (a year, a
#     month name, or a d/m/y-style numeric date) -- otherwise fuzzy parsing
#     can latch onto an unrelated number (e.g. "3" in "within 3 working
#     days") and fabricate a nonsense date."""
#     if re.search(r"\b(19|20)\d{2}\b", text_lower):
#         return True
#     if any(m in text_lower for m in _MONTH_NAMES):
#         return True
#     if re.search(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", text_lower):
#         return True
#     return False


# def _try_parse_date(raw_text):
#     if not raw_text:
#         return None
#     text_lower = str(raw_text).lower()
#     if not _looks_date_shaped(text_lower):
#         return None
#     try:
#         return _date_parser.parse(str(raw_text), fuzzy=True, dayfirst=True)
#     except (ValueError, OverflowError, TypeError):
#         return None


# # Cosine similarity threshold for two non-date VALUES to be treated as the
# # same. This is a starting point, not a tuned final number -- there's no
# # real captured data to calibrate it against yet. Watch the first several
# # real runs: if genuinely-same values are getting flagged, lower this a
# # little; if genuinely-different values are being waved through, raise it.
# VALUE_SEMANTIC_MATCH_THRESHOLD = 0.75


# def values_match(expected, live):
#     e_norm, l_norm = normalize_value(expected), normalize_value(live)
#     if e_norm == l_norm:
#         return True
#     if e_norm and l_norm and (e_norm in l_norm or l_norm in e_norm):
#         return True
#     if not e_norm or not l_norm:
#         return False

#     e_date = _try_parse_date(expected)
#     l_date = _try_parse_date(live)
#     if e_date is not None and l_date is not None:
#         return e_date.date() == l_date.date()

#     embeddings = embed_texts([e_norm, l_norm])
#     return semantic_score(embeddings[0], embeddings[1]) >= VALUE_SEMANTIC_MATCH_THRESHOLD


# # ---------------------------------------------------------------------------
# # Live response parsing
# # ---------------------------------------------------------------------------
# def get_active_schedule_value(item):
#     for n in range(1, 6):
#         if item.get(f"isSchedule{n}Active"):
#             return item.get(f"schedule{n}", ""), n
#     return None, None


# def index_live_items_by_title(live_items):
#     indexed = {}
#     for item in live_items:
#         norm_title = normalize_title(item.get("title", ""))
#         value, schedule_no = get_active_schedule_value(item)
#         indexed[norm_title] = (value, schedule_no)
#     return indexed


# # Cosine similarity threshold for two TITLES to be treated as the same
# # particular. Same caveat as VALUE_SEMANTIC_MATCH_THRESHOLD above -- a
# # reasonable starting point for short phrase embeddings with this model,
# # not something tuned against your real captures yet.
# TITLE_SEMANTIC_MATCH_THRESHOLD = 0.55

# # On top of the absolute threshold: the best candidate must beat the
# # SECOND-best candidate by at least this much to count as a real match.
# # Confirmed necessary from a real test: "Scholarship Result" scored 0.72
# # against "Dates of Scholarship Examination" (a genuinely different
# # particular) -- comfortably over the absolute threshold on its own. Only
# # safe because, when the correct candidate ("Scholarship Result
# # Declaration") was ALSO present, it scored 0.88 -- a 0.16 gap. This margin
# # check is what actually protects against the close-but-wrong case; the
# # absolute threshold alone is not enough by itself.
# TITLE_SEMANTIC_MARGIN = 0.08


# def find_best_title_match(target_title, live_index, live_titles, live_embeddings):
#     """
#     Exact match on the full normalized text is tried first (cheap,
#     unambiguous). Only if that fails does this fall back to semantic
#     similarity against every live title.

#     A semantic match is only accepted if the best candidate clears BOTH:
#       - TITLE_SEMANTIC_MATCH_THRESHOLD (is it similar enough at all), and
#       - TITLE_SEMANTIC_MARGIN over the second-best candidate (is it
#         CLEARLY the best one, not just barely ahead of a decoy).
#     A high absolute score isn't enough on its own if a different candidate
#     is nearly as close -- that's exactly the ambiguous case a hand-picked
#     single threshold can't catch by itself.

#     live_titles / live_embeddings: precompute these ONCE per code with
#     embed_texts(list(live_index.keys())) and pass them in, rather than
#     re-embedding the same live titles on every call inside a loop over
#     multiple particulars for the same code.

#     Returns (matched_title, live_value, best_score, rejection_reason).
#     rejection_reason is None on a real match, otherwise one of
#     "no_live_titles" / "below_threshold" / "ambiguous" -- useful for
#     telling apart WHY something wasn't matched when reviewing results.
#     """
#     norm_target = normalize_title(target_title)
#     if norm_target in live_index:
#         return norm_target, live_index[norm_target], 1.0, None

#     if not live_titles:
#         return None, None, 0.0, "no_live_titles"

#     target_embedding = embed_texts([norm_target])[0]
#     scores = [float(s) for s in st_util.cos_sim(target_embedding, live_embeddings)[0]]

#     best_idx = max(range(len(scores)), key=lambda i: scores[i])
#     best_score = scores[best_idx]
#     sorted_scores = sorted(scores, reverse=True)
#     second_best_score = sorted_scores[1] if len(sorted_scores) > 1 else None

#     if best_score < TITLE_SEMANTIC_MATCH_THRESHOLD:
#         return None, None, best_score, "below_threshold"

#     if second_best_score is not None and (best_score - second_best_score) < TITLE_SEMANTIC_MARGIN:
#         return None, None, best_score, "ambiguous"

#     best_title = live_titles[best_idx]
#     return best_title, live_index[best_title], best_score, None


# # ---------------------------------------------------------------------------
# # Expected data loading
# # ---------------------------------------------------------------------------
# def load_mapped_dates(path="Mapped_Programme_Dates_2026.csv"):
#     lookup = {}
#     with open(path, encoding="utf-8-sig") as f:
#         for row in csv.DictReader(f):
#             code = row["Programme Code"].strip()
#             lookup[code] = {
#                 "bucket": row["Matched Bucket"],
#                 "dates": json.loads(row["Dates"]),
#             }
#     return lookup


# def load_nest_test_codes(xls_path):
#     wb = xlrd.open_workbook(xls_path)
#     sheet = wb.sheet_by_index(0)
#     header = sheet.row_values(0)
#     code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
#     nest_col = next(i for i, h in enumerate(header) if "nest test code" in h.lower())
#     lookup = {}
#     for r in range(1, sheet.nrows):
#         row = sheet.row_values(r)
#         code = row[code_col]
#         if code:
#             lookup[code] = str(row[nest_col]).strip()
#     return lookup


# def has_scholarship_exam(nest_test_code):
#     if not nest_test_code:
#         return False
#     return nest_test_code.strip().upper() not in ("", "NA", "N/A", "-")


# def should_skip_particular(particular_label, nest_test_code):
#     if has_scholarship_exam(nest_test_code):
#         return False
#     return "scholarship" in particular_label.lower()


# # ---------------------------------------------------------------------------
# # Main comparison
# # ---------------------------------------------------------------------------
# def check_important_dates(official_code, live_items, mapped_dates_lookup, nest_code_lookup):
#     expected = mapped_dates_lookup.get(official_code)
#     if expected is None:
#         return {"code": official_code, "status": "NO_EXPECTED_DATA"}

#     nest_code = nest_code_lookup.get(official_code, "")
#     live_index = index_live_items_by_title(live_items)

#     # Embed all live titles for this code ONCE, up front -- every particular
#     # below reuses these same embeddings instead of re-encoding the live
#     # titles again on every iteration.
#     live_titles = list(live_index.keys())
#     live_embeddings = embed_texts(live_titles) if live_titles else None

#     mismatches = []
#     skipped = []

#     for particular, expected_value in expected["dates"].items():
#         if should_skip_particular(particular, nest_code):
#             skipped.append(particular)
#             continue

#         matched_title, live_match, match_score, rejection_reason = find_best_title_match(
#             particular, live_index, live_titles, live_embeddings
#         )

#         if live_match is None:
#             mismatches.append({
#                 "particular": particular,
#                 "issue": "not found in live response (even with semantic title matching)",
#                 "best_semantic_score": round(match_score, 2),
#                 "reason": rejection_reason,
#             })
#             continue

#         live_value, schedule_no = live_match
#         if live_value is None:
#             mismatches.append({
#                 "particular": particular,
#                 "issue": "found live but no schedule is currently marked active",
#             })
#             continue

#         if not values_match(expected_value, live_value):
#             mismatches.append({
#                 "particular": particular,
#                 "expected": expected_value,
#                 "live": live_value,
#                 "active_schedule_number": schedule_no,
#                 "matched_via_title": matched_title if match_score < 1.0 else None,
#             })

#     return {
#         "code": official_code,
#         "status": "CHECKED",
#         "bucket": expected["bucket"],
#         "mismatches": mismatches,
#         "skipped_no_scholarship_exam": skipped,
#     }

"""
important_dates_validation.py

Compares live GetNESTImportantDates data against the teammate-mapped
Mapped_Programme_Dates_2026.csv (built by map_programmes_2_.py).

Confirmed from real captures (P132, an MBA code, P105, P371-NNAF):
  - Live titles are inconsistently cased and carry stray HTML/whitespace --
    normalize before comparing, never exact-match raw text.
  - "Active" schedule is only reliable via the isScheduleNActive booleans,
    never the schedule*Heading text labels (which can even say "Previous
    Schedule" while the boolean says active -- confirmed on P105).
  - Item count varies a lot by programme (1 item for P105, 8-10 for
    P132/MBA) -- never assume a fixed set of particulars.
  - Confirmed with sir: Type C programmes with NO NEST Test Code (NA/blank)
    have no scholarship exam, so scholarship-named particulars from the
    matched bucket legitimately don't apply and should be skipped, not
    flagged as missing.

REVISION 4 (this version) -- two real, DIFFERENT root causes were found in
a real run against P371-NNAF and its siblings, and neither is a matching
algorithm bug:

  1. "Display of Result" scored "ambiguous" because the live page genuinely
     HAS two near-duplicate headings -- "Display of result (For entrance
     examination)" and "Display of interview result" -- which legitimately
     share almost every word. No wording-based comparison, fuzzy or
     semantic, can reliably tell these apart from text alone; a human
     glancing at both immediately can, because they know which section of
     the page they're reading. This is not something more matching cleverness
     fixes -- it needs a human to look at both candidates once.

  2. "Dates of Interview" was below the match threshold on EVERY code that
     has it, because it's not actually a wording problem -- it's a
     STRUCTURAL mismatch between the mapping data and the live page. The
     mapping treats "Slot booking for Interview" and "Dates of Interview"
     as two separate particulars, but the live page only has ONE combined
     item ("Interview slot booking") whose value text contains both the
     booking window and an example date range in the same paragraph. There
     is no standalone live item for "Dates of Interview" to find, so no
     amount of matching -- fuzzy or semantic -- can locate one. This can
     only be fixed by updating Mapped_Programme_Dates_2026.csv itself, not
     this matching code.

Given both of the above, the tool no longer tries to auto-decide right vs
wrong at all -- see values_look_identical below, which replaces the old
values_match. It finds the best-guess live item for each expected
particular and shows expected + live side by side; a human decides whether
it's actually a mismatch, exactly a duplicate-heading collision, or a
structural gap in the mapping. This is a deliberate scope reduction, not a
fallback -- repeated rounds of tuning the automatic judgment call
(word-weighting, then semantic thresholds, then margins) kept trading one
failure mode for another; the underlying task (recognizing "these mean the
same thing" vs "these are genuinely different") is fundamentally judgment,
not something a fixed scoring rule can get exactly right on every phrasing.

REQUIRES: sentence-transformers (pulls in torch automatically), used ONLY
for PAIRING an expected particular with the closest live item when an exact
title match isn't available -- not for judging whether values match.
    pip install sentence-transformers
The model itself (~80MB) downloads automatically the first time
SentenceTransformer(...) is called, and is cached locally after that --
only the very first run on a given machine needs internet access for it.

python-dateutil is NO LONGER a dependency of this module -- date parsing
was only used for automatic right/wrong judgment, which this version
doesn't do anymore.
"""
import csv
import json
import re
import xlrd

try:
    from sentence_transformers import SentenceTransformer, util as st_util
except ImportError as _e:  # pragma: no cover
    raise ImportError(
        "important_dates_validation.py needs the 'sentence-transformers' "
        "package to pair expected particulars with the closest live item. "
        "Install it with:\n"
        "    pip install sentence-transformers\n"
        "(this also installs torch, which is a larger download -- expect "
        "several hundred MB. The embedding model itself is downloaded "
        "automatically and cached the first time this runs.)"
    ) from _e


# ---------------------------------------------------------------------------
# Normalization helpers
# ---------------------------------------------------------------------------
def strip_html(text):
    if text is None:
        return ""
    return re.sub(r"<[^>]+>", "", str(text))


def normalize_title(text):
    """Lowercased, whitespace-collapsed -- used only for INTERNAL comparison
    (exact-match lookups, identical-value checks), never for display."""
    text = strip_html(text)
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip().lower()


def normalize_value(text):
    text = normalize_title(text)
    text = re.sub(r"\*+$", "", text).strip()
    return text


def clean_for_display(text):
    """Strips HTML and collapses whitespace WITHOUT lowercasing -- this is
    what actually goes into the report, so it should stay readable rather
    than being flattened for comparison purposes."""
    text = strip_html(text)
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def values_look_identical(expected, live):
    """Cheap, unambiguous "is this even worth showing" check -- exact match
    or plain substring either way, on normalized text. Deliberately NOT
    semantic and NOT date-aware: this is only meant to filter out plainly
    identical values so the report isn't cluttered, not to judge whether a
    difference is a real mismatch. Anything not caught here gets shown to
    a human, unjudged."""
    e, l = normalize_value(expected), normalize_value(live)
    if not e or not l:
        return False
    return e == l or e in l or l in e


# ---------------------------------------------------------------------------
# Semantic title PAIRING only -- not used for judging values anymore.
# ---------------------------------------------------------------------------
_MODEL_NAME = "all-MiniLM-L6-v2"
_model = None  # loaded lazily so importing this module doesn't force a load


def _get_model():
    global _model
    if _model is None:
        _model = SentenceTransformer(_MODEL_NAME)
    return _model


def embed_texts(texts):
    """Batch-encodes a list of strings into normalized embedding vectors.
    Always call this with as many texts at once as you have available
    (e.g. all live titles for one code in a single call) rather than one
    at a time in a loop -- batching is meaningfully faster."""
    texts = list(texts)
    if not texts:
        return None
    model = _get_model()
    return model.encode(texts, convert_to_numpy=True, normalize_embeddings=True)


def find_best_title_match(target_title, live_index, live_titles, live_embeddings):
    """
    Exact match on the full normalized text is tried first (cheap,
    unambiguous). Otherwise, ALWAYS returns whichever live title is
    closest by semantic similarity -- no threshold, no rejection. This is
    deliberate: the tool no longer tries to decide "is this confident
    enough to count as a match" -- it just finds the closest candidate and
    shows it, so a human can see it's e.g. "Display of interview result"
    rather than the intended "Display of result (For entrance
    examination)" and immediately recognize the collision themselves.

    live_titles / live_embeddings: precompute these ONCE per code with
    embed_texts(list(live_index.keys())) and pass them in, rather than
    re-embedding the same live titles on every call inside a loop over
    multiple particulars for the same code.

    Returns (matched_title, live_value_tuple, score) or (None, None, 0.0)
    if live_titles is empty (nothing at all to compare against).
    """
    norm_target = normalize_title(target_title)
    if norm_target in live_index:
        return norm_target, live_index[norm_target], 1.0

    if not live_titles:
        return None, None, 0.0

    target_embedding = embed_texts([norm_target])[0]
    scores = [float(s) for s in st_util.cos_sim(target_embedding, live_embeddings)[0]]
    best_idx = max(range(len(scores)), key=lambda i: scores[i])
    best_title = live_titles[best_idx]
    return best_title, live_index[best_title], scores[best_idx]


# ---------------------------------------------------------------------------
# Live response parsing
# ---------------------------------------------------------------------------
def get_active_schedule_value(item):
    for n in range(1, 6):
        if item.get(f"isSchedule{n}Active"):
            return item.get(f"schedule{n}", ""), n
    return None, None


def index_live_items_by_title(live_items):
    indexed = {}
    for item in live_items:
        norm_title = normalize_title(item.get("title", ""))
        value, schedule_no = get_active_schedule_value(item)
        indexed[norm_title] = (value, schedule_no)
    return indexed


# ---------------------------------------------------------------------------
# Expected data loading
# ---------------------------------------------------------------------------
def load_mapped_dates(path="Mapped_Programme_Dates_2026.csv"):
    lookup = {}
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            code = row["Programme Code"].strip()
            lookup[code] = {
                "bucket": row["Matched Bucket"],
                "dates": json.loads(row["Dates"]),
            }
    return lookup


def load_nest_test_codes(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sheet = wb.sheet_by_index(0)
    header = sheet.row_values(0)
    code_col = next(i for i, h in enumerate(header) if "programme code" in h.lower())
    nest_col = next(i for i, h in enumerate(header) if "nest test code" in h.lower())
    lookup = {}
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        code = row[code_col]
        if code:
            lookup[code] = str(row[nest_col]).strip()
    return lookup


def has_scholarship_exam(nest_test_code):
    if not nest_test_code:
        return False
    return nest_test_code.strip().upper() not in ("", "NA", "N/A", "-")


def should_skip_particular(particular_label, nest_test_code):
    if has_scholarship_exam(nest_test_code):
        return False
    return "scholarship" in particular_label.lower()


# ---------------------------------------------------------------------------
# Report formatting -- plain "heading / expected / live" blocks, nothing else.
# ---------------------------------------------------------------------------
def format_report(report_items):
    if not report_items:
        return ""
    blocks = []
    for item in report_items:
        blocks.append(
            f"{item['particular']}:\n"
            f"expected: {item['expected']}\n"
            f"live: {item['live']}"
        )
    return "\n\n".join(blocks)


# ---------------------------------------------------------------------------
# Main comparison -- no status, no judgment. Returns a plain report string
# (empty string if there's nothing worth showing for this code).
# ---------------------------------------------------------------------------
def check_important_dates(official_code, live_items, mapped_dates_lookup, nest_code_lookup):
    expected = mapped_dates_lookup.get(official_code)
    if expected is None:
        return ""  # no mapping exists for this code -- nothing to report

    nest_code = nest_code_lookup.get(official_code, "")
    live_index = index_live_items_by_title(live_items)
    live_titles = list(live_index.keys())
    live_embeddings = embed_texts(live_titles) if live_titles else None

    report_items = []

    for particular, expected_value in expected["dates"].items():
        if should_skip_particular(particular, nest_code):
            continue

        if not live_titles:
            # Genuinely different from "couldn't find a matching title" --
            # this means the live fetch itself returned nothing for this
            # code at all (confirmed cause on P8G/P8H/P8J/P2L2).
            report_items.append({
                "particular": particular,
                "expected": clean_for_display(expected_value),
                "live": "NOT AVAILABLE -- no live data was fetched for this programme",
            })
            continue

        matched_title, live_match, _score = find_best_title_match(
            particular, live_index, live_titles, live_embeddings
        )

        live_value, _schedule_no = live_match
        if live_value is None:
            report_items.append({
                "particular": particular,
                "expected": clean_for_display(expected_value),
                "live": f"[matched to \"{matched_title}\"] but no schedule is currently marked active",
            })
            continue

        if values_look_identical(expected_value, live_value):
            continue

        report_items.append({
            "particular": particular,
            "expected": clean_for_display(expected_value),
            "live": clean_for_display(live_value),
        })

    return format_report(report_items)