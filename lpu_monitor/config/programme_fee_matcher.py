"""
generate_programme_fee_mapping.py
==================================================================
FIX APPLIED (confirmed by testing against a real failure case):
P13Q-L ("B.Tech. (Food Technology) [Lateral Entry]") was wrongly
matching "B.Tech. [Lateral Entry] (Construction Technology)" instead
of the correct generic "B.Tech. [Lateral Entry]" row. Root cause: the
shared generic word "Technology" inflated both halves of the
similarity score. Fixed by:
  1. Adding generic academic-category words to STOPWORDS.
  2. Computing the character-level similarity check on the SAME
     stopword-filtered tokens as the word-overlap check.
Verified this does NOT weaken genuine matches (a real Construction
Technology programme still scores a perfect 1.0 against its own row).
==================================================================
"""

import csv
import glob
import os
import re
import subprocess
import sys
import tempfile
from collections import Counter
from difflib import SequenceMatcher

import openpyxl

# Auto-discovered from incoming/ folders -- no hardcoded filenames, same
# pattern as important_dates_mapping_builder.py and every other builder in
# this project. Exact filenames change every cycle (dates in the name), so
# hardcoding them silently breaks the script the next time files are dropped.
INCOMING_FEE_EXCEL = "lpu_monitor/incoming/fee_excel"
INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"
OUTPUT_CSV = "lpu_monitor/config/programme_fee_mapping.csv"

FEE_SHEET_NAMES = ["After 10th", "After 10+2", "Lateral Entry", "After Graduation"]

EXACT_THRESHOLD = 0.90
FUZZY_THRESHOLD = 0.55
AMBIGUOUS_MARGIN = 0.06


def ensure_xlsx(path, workdir):
    if path.lower().endswith(".xlsx"):
        return path
    if not path.lower().endswith(".xls"):
        raise ValueError(f"Unsupported file type: {path}")
    out_name = os.path.splitext(os.path.basename(path))[0] + ".xlsx"
    out_path = os.path.join(workdir, out_name)
    soffice = None
    for candidate in ("soffice", "libreoffice"):
        try:
            subprocess.run([candidate, "--version"], capture_output=True, check=True)
            soffice = candidate
            break
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
    if soffice is None:
        raise RuntimeError(
            "Could not find LibreOffice on PATH to convert .xls to .xlsx. "
            "Either install LibreOffice, or re-save the two files as .xlsx "
            "and update the file path constants above."
        )
    subprocess.run(
        [soffice, "--headless", "--convert-to", "xlsx", "--outdir", workdir, path],
        check=True, capture_output=True,
    )
    if not os.path.exists(out_path):
        raise RuntimeError(f"LibreOffice conversion of {path} did not produce {out_path}")
    return out_path


def normalize(s):
    if s is None:
        return ""
    s = str(s).lower()
    # Strip trailing tie-up/partnership boilerplate (e.g. "in tieup with IBM",
    # "in collaboration with Comptia") before anything else. This text is
    # business metadata attached to some fee-sheet labels, not part of the
    # academic programme identity, and it dilutes token-overlap similarity
    # against the programme name badly enough to sink a correct match below
    # a generic "all branches"-style fallback row. No programme name in the
    # programme list contains this pattern, so stripping it here is safe.
    s = re.sub(r"\bin\s+(?:tie[\s-]?up|collaboration|association|partnership)\s+with\b.*$", " ", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\[.*?\]", " ", s)
    s = re.sub(r"\{.*?\}", " ", s)
    s = re.sub(r"\[.*$", " ", s)
    s = re.sub(r"\bhons\.?\b", "hons", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


STOPWORDS = {
    "b", "m", "of", "and", "the", "in", "for", "a", "an", "with", "hons",
    "bachelor", "bachelors", "master", "masters", "programme", "program",
    "technology", "engineering", "science", "sciences", "studies", "design",
}


def _stem(t):
    if len(t) > 4 and t.endswith("s") and not t.endswith("ss"):
        return t[:-1]
    return t


def token_set(s):
    return {_stem(t) for t in normalize(s).split() if t not in STOPWORDS and len(t) > 1}


def similarity(a_norm, b_norm):
    if not a_norm or not b_norm:
        return 0.0
    ta = {_stem(t) for t in a_norm.split() if t not in STOPWORDS}
    tb = {_stem(t) for t in b_norm.split() if t not in STOPWORDS}
    a_filtered = " ".join(sorted(ta))
    b_filtered = " ".join(sorted(tb))
    seq_ratio = SequenceMatcher(None, a_filtered, b_filtered).ratio()
    jaccard = len(ta & tb) / len(ta | tb) if (ta | tb) else 0.0
    score = 0.35 * seq_ratio + 0.65 * jaccard
    if jaccard == 0.0:
        score = min(score, 0.2)
    return score


def clean_label(raw):
    s = str(raw).replace("\r\n", "\n")
    s = re.sub(r"\nNote:-.*", "", s, flags=re.S)
    s = re.sub(r"\(From \d+.*?onwards\)", "", s, flags=re.I)
    s = s.replace("\n", " / ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    s = re.sub(r"\s*/\s*$", "", s).strip()
    return s


_DEGREE_PREFIX = (
    r"B\.?\s?Tech\.?|M\.?\s?Tech\.?|B\.?\s?Sc\.?|M\.?\s?Sc\.?|BCA|MCA|BBA|MBA|"
    r"B\.?\s?Com\.?|M\.?\s?Com\.?|B\.?\s?A\.?|M\.?\s?A\.?|BPA|MPA|"
    r"B\.?\s?Design|M\.?\s?Design|B\s?Design|Diploma|Bachelor|Master|"
    r"BPT|MPT|B\.?\s?Pharm\.?|M\.?\s?Pharm\.?|B\.?\s?Ed\.?|M\.?\s?Ed\.?|Integrated"
)
_SEGMENT_SPLIT_RE = re.compile(r"[/,]\s*(?=(?:" + _DEGREE_PREFIX + r"))", re.I)


def split_segments(clean):
    parts = [p.strip() for p in _SEGMENT_SPLIT_RE.split(clean) if p.strip()]
    parts = parts or [clean]
    expanded = list(parts)
    for seg in parts:
        for pm in re.finditer(r"\(([^)]*/[^)]*)\)", seg):
            inner = pm.group(1)
            if re.search(r"\bexcept\b", inner, re.I):
                continue
            alts = [a.strip() for a in inner.split("/") if a.strip()]
            if len(alts) < 2:
                continue
            if all(re.sub(r"\s", "", a).isupper() or len(re.sub(r"\s", "", a)) <= 4 for a in alts):
                continue
            for alt in alts:
                expanded.append(seg[:pm.start()] + f"({alt})" + seg[pm.end():])
    return expanded


def load_abbreviations(wb):
    expanded_to_short = {}
    short_to_expanded = {}
    for sheet_name in wb.sheetnames:
        if "abbrev" not in sheet_name.lower():
            continue
        ws = wb[sheet_name]
        header = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]
        try:
            exp_col = next(i for i, h in enumerate(header) if "expand" in h)
            short_col = next(i for i, h in enumerate(header) if "short" in h)
        except StopIteration:
            continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) <= max(exp_col, short_col):
                continue
            expanded, short = row[exp_col], row[short_col]
            if not expanded or not short:
                continue
            exp_norm = normalize(expanded)
            short_norm = normalize(short)
            if exp_norm and short_norm:
                expanded_to_short[exp_norm] = short_norm
                short_to_expanded[short_norm] = exp_norm
    return expanded_to_short, short_to_expanded


def derive_branch_code(name_norm, expanded_to_short):
    for exp_norm, short in expanded_to_short.items():
        if exp_norm and exp_norm in name_norm:
            return short
    return None


PAREN_CODE_RE = re.compile(r"\(([a-zA-Z][a-zA-Z. ]{0,8})\)\s*$")


def trailing_paren_code(raw_name):
    name_stripped = re.sub(r"\[.*$", "", raw_name.strip()).strip()
    m = PAREN_CODE_RE.search(name_stripped)
    if not m:
        return None
    inner = m.group(1).strip()
    bare = inner.replace(" ", "").replace(".", "")
    if not bare or not bare.isupper() or len(bare) > 8:
        return None
    return normalize(inner)


def load_programme_list(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    target_sheet = None
    header_row_idx = None
    col = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for r in range(1, min(ws.max_row, 5) + 1):
            values = [c.value for c in ws[r]]
            headers = [str(v).strip().lower() if v else "" for v in values]
            if any("eligibility" in h for h in headers) and any("programme name" in h or "program name" in h for h in headers):
                target_sheet = sheet_name
                header_row_idx = r
                for i, h in enumerate(headers):
                    if "eligibility" in h and "elig" not in col:
                        col["elig"] = i
                    if ("programme name" in h or "program name" in h) and "name" not in col:
                        col["name"] = i
                    if h.startswith("discipline") and "discipline" not in col:
                        col["discipline"] = i
                    if "programme code" in h or "program code" in h or h == "code":
                        col.setdefault("code", i)
                break
        if target_sheet:
            break
    if target_sheet is None or "elig" not in col or "name" not in col or "code" not in col:
        raise RuntimeError("Could not locate the programme-list sheet/columns automatically.")
    ws = wb[target_sheet]
    programmes = []
    for r in range(header_row_idx + 1, ws.max_row + 1):
        elig = ws.cell(r, col["elig"] + 1).value
        name = ws.cell(r, col["name"] + 1).value
        code = ws.cell(r, col["code"] + 1).value
        if code is None or name is None:
            continue
        elig = str(elig).strip() if elig else ""
        name = str(name).strip()
        code = str(code).strip()
        if not code or not name:
            continue
        if "ph.d" in elig.lower() or "ph.d" in name.lower():
            continue
        programmes.append({"EligibilityBase": elig, "ProgrammeName": name, "OfficialCode": code})
    return programmes, wb


def load_fee_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    fee_rows = []
    sheets_present = [s for s in FEE_SHEET_NAMES if s in wb.sheetnames]
    if not sheets_present:
        raise RuntimeError(f"None of {FEE_SHEET_NAMES} found (sheets present: {wb.sheetnames}).")
    for sheet_name in sheets_present:
        ws = wb[sheet_name]
        group = None
        category = None
        for i in range(1, ws.max_row + 1):
            a = ws.cell(i, 1).value
            b = ws.cell(i, 2).value
            if a is None:
                continue
            a = str(a).strip()
            if not a:
                continue
            if b is None or (isinstance(b, str) and not b.strip()):
                if a.upper().startswith("AFTER"):
                    group = a
                    category = None
                elif re.match(r"^[A-Z0-9 .,/()&\-]+$", a) and len(a) < 60:
                    category = a
                continue
            label = clean_label(a)
            if not label:
                continue
            fee_rows.append({
                "Sheet": sheet_name, "Group": group, "Category": category, "Row": i,
                "CleanLabel": label, "Segments": split_segments(label),
                "IsExclusion": bool(re.search(r"\bexcept\b", label, re.I)),
            })
    return fee_rows, wb


def eligibility_to_sheet_group(elig, fee_rows):
    e = elig.lower()
    sheets = {r["Sheet"] for r in fee_rows}
    def sheet_matching(kw):
        for s in sheets:
            if any(k in s.lower() for k in kw):
                return s
        return None
    if "10th" in e:
        return sheet_matching(["10th"]), None
    if "10+2" in e or "12th" in e:
        return sheet_matching(["10+2", "12th"]), None
    if "graduation" in e and "lateral" not in e:
        return sheet_matching(["graduation"]), None
    if "lateral" in e:
        sheet = sheet_matching(["lateral"])
        groups = {r["Group"] for r in fee_rows if r["Sheet"] == sheet and r["Group"]}
        if "diploma" in e:
            group = next((g for g in groups if "diploma" in g.lower()), None)
        elif "iti" in e:
            group = next((g for g in groups if "iti" in g.lower()), None)
        else:
            group = None
        return sheet, group
    return None, None


GENERIC_MARKERS = ["all branches", "all disciplines", "except"]


def extract_code_sets(label, expanded_to_short):
    known_short = set(expanded_to_short.values())
    exclusion, inclusion = set(), set()
    for m in re.finditer(r"\(([^)]*)\)", label):
        inner = m.group(1)
        is_excl = bool(re.search(r"\bexcept\b", inner, re.I))
        inner_wo_except = re.sub(r"\bexcept\b", " ", inner, flags=re.I)
        parts = re.split(r"[/,&]|\band\b", inner_wo_except, flags=re.I)
        codes = {normalize(part) for part in parts if normalize(part)}
        codes &= known_short
        if codes:
            (exclusion if is_excl else inclusion).update(codes)
    return exclusion, inclusion


def _leading_token(text_norm):
    for t in text_norm.split():
        if t not in STOPWORDS:
            return t
    return None


def _leading_tokens_agree(a, b):
    if not a or not b:
        return False
    if a == b:
        return True
    shorter, longer = sorted([a, b], key=len)
    return len(shorter) >= 3 and longer.startswith(shorter)


def family_prefix_tokens(label_norm):
    earliest = len(label_norm)
    for marker in GENERIC_MARKERS:
        idx = label_norm.find(marker)
        if idx != -1:
            earliest = min(earliest, idx)
    text = label_norm[:earliest]
    return [t for t in text.split() if t not in STOPWORDS]


def compute_mapping_rows(programmes, fee_rows, expanded_to_short):
    out_rows = []
    for p in programmes:
        elig = p["EligibilityBase"]
        code = p["OfficialCode"]
        name = p["ProgrammeName"]
        sheet, group = eligibility_to_sheet_group(elig, fee_rows)
        if sheet is None:
            out_rows.append(_row(code, name, elig, "", "NOT_FOUND", f'Could not map EligibilityBase "{elig}".'))
            continue
        candidates = [f for f in fee_rows if f["Sheet"] == sheet and (group is None or f["Group"] == group)]
        if not candidates:
            out_rows.append(_row(code, name, elig, "", "NOT_FOUND", f'No fee rows for sheet "{sheet}".'))
            continue
        name_norm = normalize(name)
        name_tokens = {_stem(t) for t in name_norm.split() if t not in STOPWORDS}
        name_code = trailing_paren_code(name) or derive_branch_code(name_norm, expanded_to_short)
        TIER_EXACT, TIER_INCLUSION, TIER_FUZZY, TIER_FALLBACK = 4, 3, 2, 1

        # Does a genuine "catch-all" fee row (e.g. "B.Tech - all branches
        # ... except CSE/IT") exist for this programme's sheet/group at all?
        # Some sheets (After Graduation, Lateral Entry) have NO such row --
        # every branch there gets its own dedicated fee row, so a fuzzy match
        # with minor wording drift (e.g. "Engineering" vs "Engg.", "Laboratory"
        # vs "Lab") is genuinely the best and only sensible answer and must
        # not be rejected. Where a real catch-all DOES exist, a plain-named
        # programme should prefer it over a fuzzy match to a more specific
        # named variant (see segment_covered_by_name below) instead of being
        # swept into that variant's row just because they share a branch word.
        has_generic_fallback_option = any(
            set(family_prefix_tokens(normalize(c["CleanLabel"]))) and
            set(family_prefix_tokens(normalize(c["CleanLabel"]))).issubset(name_tokens)
            for c in candidates
        )

        scored = []
        for c in candidates:
            excl, incl = extract_code_sets(c["CleanLabel"], expanded_to_short)
            if name_code and name_code in excl:
                continue
            best_seg_score = 0.0
            best_seg_tokens = set()
            for seg in c["Segments"]:
                seg_norm = normalize(seg)
                seg_score = similarity(name_norm, seg_norm)
                if seg_score > best_seg_score:
                    best_seg_score = seg_score
                    best_seg_tokens = {_stem(t) for t in seg_norm.split() if t not in STOPWORDS}
            # A fee-row segment can name a MORE SPECIFIC variant than the
            # programme (e.g. "Civil - AI and ML" vs a plain "Civil
            # Engineering" programme). Sharing the base branch word ("civil")
            # plus the generic degree word ("tech") is enough to clear the
            # fuzzy-similarity threshold even though the programme never says
            # "AI" or "ML" anywhere. A fuzzy match must not silently add
            # specialization the programme name doesn't state, so it's only
            # trusted when every token in the matched segment is already
            # present in the programme name (segment_covered_by_name). This
            # only gates the coarse fuzzy tier; exact matches and
            # code-based inclusion matches are unaffected.
            segment_covered_by_name = best_seg_tokens.issubset(name_tokens)
            family_agrees = _leading_tokens_agree(_leading_token(name_norm), _leading_token(normalize(c["CleanLabel"])))
            if best_seg_score >= EXACT_THRESHOLD:
                scored.append((TIER_EXACT, best_seg_score, c, "exact"))
            elif incl and name_code and name_code in incl and family_agrees:
                scored.append((TIER_INCLUSION, best_seg_score, c, "inclusion_match"))
            elif (best_seg_score >= FUZZY_THRESHOLD and (family_agrees or best_seg_score >= 0.65)
                  and (segment_covered_by_name or not has_generic_fallback_option)):
                scored.append((TIER_FUZZY, best_seg_score, c, "fuzzy"))
            else:
                fam_tokens = family_prefix_tokens(normalize(c["CleanLabel"]))
                if fam_tokens and set(fam_tokens).issubset(name_tokens):
                    coverage = len(fam_tokens) / max(len(name_tokens), 1)
                    scored.append((TIER_FALLBACK, 0.25 + 0.20 * coverage, c, "generic_fallback"))
        if not scored:
            out_rows.append(_row(code, name, elig, "", "NOT_FOUND", "No fee row matched."))
            continue
        top_tier = max(s[0] for s in scored)
        top_group = sorted((s for s in scored if s[0] == top_tier), key=lambda t: t[1], reverse=True)
        top_score, _, top_c, top_kind = top_group[0]
        second_score = top_group[1][1] if len(top_group) > 1 else 0.0
        if top_kind == "exact":
            out_rows.append(_row(code, name, elig, top_c["CleanLabel"], "HIGH", "Exact normalized match."))
        elif len(top_group) == 1 or (top_score - second_score) >= AMBIGUOUS_MARGIN:
            reason = {
                "fuzzy": "Closest fee row by name similarity, clear margin.",
                "inclusion_match": "Branch code matches a branch-restricted fee row.",
                "generic_fallback": "Category-wide fee row; branch not excluded.",
            }[top_kind]
            out_rows.append(_row(code, name, elig, top_c["CleanLabel"], "HIGH", reason))
        else:
            labels = " | ".join(sorted({t[2]["CleanLabel"] for t in top_group if (top_score - t[1]) < AMBIGUOUS_MARGIN}))
            out_rows.append(_row(code, name, elig, labels, "AMBIGUOUS", "Multiple equally strong fee rows matched."))
    return out_rows


def _row(code, name, elig, label, status, reason):
    return {"OfficialCode": code, "ProgrammeName": name, "EligibilityBase": elig,
            "FeeRowLabel": label, "Status": status, "Reason": reason}


def find_single_file(folder, pattern="*"):
    files = [f for f in glob.glob(os.path.join(folder, pattern)) if os.path.isfile(f)]
    if len(files) == 0:
        raise FileNotFoundError(f"No file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Expected exactly 1 file in {folder}, found {len(files)}: {files}")
    return files[0]


def build_mapping(proglist_path=None, fee_path=None, out_csv=OUTPUT_CSV):
    proglist_path = proglist_path or find_single_file(INCOMING_PROGLIST, "*.xls")
    fee_path = fee_path or find_single_file(INCOMING_FEE_EXCEL, "*.xls")

    with tempfile.TemporaryDirectory() as workdir:
        prog_xlsx = ensure_xlsx(proglist_path, workdir)
        fee_xlsx = ensure_xlsx(fee_path, workdir)
        programmes, prog_wb = load_programme_list(prog_xlsx)
        fee_rows, fee_wb = load_fee_rows(fee_xlsx)
        expanded_to_short, _ = load_abbreviations(prog_wb)
        if not expanded_to_short:
            expanded_to_short, _ = load_abbreviations(fee_wb)

    mapping = compute_mapping_rows(programmes, fee_rows, expanded_to_short)

    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    fieldnames = ["OfficialCode", "ProgrammeName", "EligibilityBase", "FeeRowLabel", "Status", "Reason"]
    with open(out_csv, "w", newline="", encoding="utf-8") as fo:
        w = csv.DictWriter(fo, fieldnames=fieldnames)
        w.writeheader()
        for r in mapping:
            w.writerow(r)

    return mapping, out_csv


def main():
    mapping, out_csv = build_mapping()
    codes = [r["OfficialCode"] for r in mapping]
    dupes = {c for c in codes if codes.count(c) > 1}
    status_counts = Counter(r["Status"] for r in mapping)
    print(f"Wrote {len(mapping)} rows to {out_csv}")
    print("Status breakdown:", dict(status_counts))
    if dupes:
        print(f"WARNING: duplicate OfficialCode(s): {sorted(dupes)}")
    needs_review = [r for r in mapping if r["Status"] != "HIGH"]
    if needs_review:
        print(f"\n{len(needs_review)} row(s) need review:")
        for r in needs_review:
            print(f"  - {r['OfficialCode']}: {r['ProgrammeName']}  [{r['Status']}]  -> {r['FeeRowLabel']}")
    if needs_review:
        sys.exit(1)  # non-zero so the UI flags "needs attention", same convention as important_dates_mapping_builder.py


if __name__ == "__main__":
    main()