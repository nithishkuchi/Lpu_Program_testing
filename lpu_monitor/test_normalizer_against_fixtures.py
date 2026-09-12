"""
test_normalizer_against_fixtures.py

Runs the normalizer against all real fixtures saved so far, printing a
summary. This is a manual verification script, not a formal pytest suite yet
(that's a later step) -- but it directly exercises every confirmed response
shape: normal duplicate, duplicate-with-missing-field, empty array, and the
PhD case.

Run with: python -m lpu_monitor.test_normalizer_against_fixtures
"""

import json
from pathlib import Path

from lpu_monitor.normalization.normalizer import FieldMappingConfig, normalize_fee_response

FIXTURES_DIR = Path(__file__).parent / "tests" / "fixtures"
CONFIG_PATH = Path(__file__).parent / "config" / "field_mapping.yaml"

TEST_CASES = [
    ("P132", "get_program_fee_detail_P132.json", "normal case, identical duplicates"),
    ("P102-NNP", "get_program_fee_detail_P102-NNP.json", "duplicate with missing UniformFee in entry[1]"),
    ("P8G", "get_program_fee_detail_P8G.json", "empty array -- higher doctorate, no fee data"),
    ("PhD-CS", "get_program_fee_detail_PhD-CS.json", "regular PhD, corrects earlier over-generalization"),
]


def main() -> None:
    field_mapping = FieldMappingConfig(path=str(CONFIG_PATH))

    for official_code, filename, description in TEST_CASES:
        print(f"\n{'=' * 70}")
        print(f"{official_code} -- {description}")
        print("=" * 70)

        with open(FIXTURES_DIR / filename, encoding="utf-8") as f:
            raw = json.load(f)

        result = normalize_fee_response(raw, field_mapping, official_code=official_code)

        print(f"Applicable: {result.applicable}")
        print(f"Fields extracted: {len(result.fields)}")
        print(f"Inconsistencies found: {len(result.inconsistencies)}")
        print(f"Missing keys (schema drift signal): {len(result.missing_keys)}")

        if result.inconsistencies:
            for inc in result.inconsistencies:
                print(f"  INCONSISTENCY: {inc}")

        if result.missing_keys:
            for mk in result.missing_keys[:5]:
                print(f"  MISSING: {mk}")
            if len(result.missing_keys) > 5:
                print(f"  ... and {len(result.missing_keys) - 5} more")

        if result.fields:
            print("\nSample extracted fields (first 8):")
            for field in result.fields[:8]:
                print(f"  {field.canonical_name:40s} = {field.value!r}")

        # Specific check for the P102-NNP case: confirm UniformFee coalesced
        # to "4000 (One time)" (the non-null value), not None.
        if official_code == "P102-NNP":
            uniform_fee_field = next((f for f in result.fields if f.canonical_name == "uniform_fee"), None)
            assert uniform_fee_field is not None, "uniform_fee field missing entirely!"
            assert uniform_fee_field.value == "4000 (One time)", \
                f"Coalesce FAILED: expected '4000 (One time)', got {uniform_fee_field.value!r}"
            print("\n  PASS: uniform_fee correctly coalesced to the non-null value")

        # Specific check for P8G: must be not-applicable, zero fields.
        if official_code == "P8G":
            assert result.applicable is False, "P8G should be not-applicable (empty array)"
            assert len(result.fields) == 0, "P8G should produce zero fields"
            print("\n  PASS: empty array correctly treated as not-applicable")

        # Specific check for PhD-CS: must be applicable with real data
        # (this is the whole point of the earlier correction).
        if official_code == "PhD-CS":
            assert result.applicable is True, "PhD-CS should be applicable -- it has real fee data!"
            full_fee_field = next((f for f in result.fields if f.canonical_name == "full_fee_per_phase_current"), None)
            assert full_fee_field is not None and full_fee_field.value == "60000", \
                f"Expected FullFee=60000 for PhD-CS, got {full_fee_field}"
            print("\n  PASS: PhD-CS correctly shows real fee data (FullFee=60000), "
                  "confirming regular PhDs are NOT empty like higher doctorates")

    print(f"\n{'=' * 70}")
    print("ALL CHECKS PASSED")
    print("=" * 70)


if __name__ == "__main__":
    main()
