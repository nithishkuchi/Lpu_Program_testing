import csv
import json
import re

ANNEXURE_TITLE_MATCH = {
    "A": "need based",
    "B": "orphan",
    "C": "certain disabilit",
    "D": "rakshak",
    "E": "lpu rise",
    "F": "shikshak",
    "H": "sports",
    "I": "sports",
    "J": "start up",
}
def load_lookup(path="lpu_monitor/config/annexure_lookup.csv"):
    lookup = {}

    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            lookup[row["OfficialCode"]] = row

    return lookup
def normalize(text):
    text = text.lower()
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("startup", "start up")
    text = re.sub(r"[-/]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = text.replace("disabilities", "disability")
    return text.strip()
def check_programme(code, tabs_json, lookup, title_match):

    row = lookup.get(code)

    if not row:
        return {
            "code": code,
            "status": "NO_LOOKUP_ROW"
        }

    actual_titles = [normalize(item["title"]) for item in tabs_json]
    mismatches = []

    for letter, keyword in title_match.items():

        if letter in ("H", "I"):
            expected = row["H"] == "Yes" or row["I"] == "Yes"
        else:
            expected = row[letter] == "Yes"
        keyword = normalize(keyword)
        found = any(keyword in t for t in actual_titles)
        
        if expected != found:
            mismatches.append({
                "annexure": letter,
                "expected": expected,
                "found": found
            })

    return {
        "code": code,
        "mismatches": mismatches
    }
if __name__ == "__main__":
    lookup = load_lookup()

    print("Checking P102...")

    with open("live_capture_tabs_filter1.json", encoding="utf-8") as f:
        p102_tabs = json.load(f)

    print(
        check_programme(
            "P102",
            p102_tabs,
            lookup,
            ANNEXURE_TITLE_MATCH,
        )
    )

    print()

    print("Checking P124...")

    with open("tabs_filter1_P124.json", encoding="utf-8") as f:
        p124_tabs = json.load(f)

    print(
        check_programme(
            "P124",
            p124_tabs,
            lookup,
            ANNEXURE_TITLE_MATCH,
        )
    )
    print()

    print("Checking P132...")

    with open("tabs_filter1_P132.json", encoding="utf-8") as f:
        p132_tabs = json.load(f)

    print(
        check_programme(
            "P132",
            p132_tabs,
            lookup,
            ANNEXURE_TITLE_MATCH,
        )
    )