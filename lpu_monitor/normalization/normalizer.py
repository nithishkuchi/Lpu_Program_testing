"""
normalizer.py

Converts a raw GetProgramFeeDetail API response into a flat list of canonical
(field_group, canonical_name, value) rows, using config/field_mapping.yaml.

Implements the two confirmed response-shape rules from real testing (see
HANDOFF.md Section 3):

1. EMPTY ARRAY -> "not applicable." Returns an empty list, logs plainly,
   never raises an error, never treated as a diff-worthy change.

2. MULTI-ENTRY ARRAYS -> COALESCE. For every mapped field, take the first
   non-null value found scanning all array entries in order -- never a fixed
   index. If entries disagree (two different non-null values for the same
   field), that's logged distinctly as a real data inconsistency, not
   silently resolved.

Also implements schema-drift detection (SDD Section 20): if a field the
mapping expects is missing from EVERY entry (the key doesn't exist at all,
as opposed to existing with a null value), that's logged separately -- it
likely means LPU changed the API's response shape, not that the programme
simply has no value for that field.
"""

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import yaml

logger = logging.getLogger("lpu_monitor.normalizer")

_SENTINEL_MISSING = object()  # distinguishes "key absent" from "key present but null"


@dataclass
class NormalizedField:
    field_group: str
    canonical_name: str
    value: Any
    source_field_names: str  # the raw json_path(s) this came from, for traceability


@dataclass
class NormalizationResult:
    fields: list[NormalizedField]
    inconsistencies: list[dict]   # entries disagreed on a non-null value
    missing_keys: list[dict]      # mapped field never appeared in ANY entry (schema drift)
    applicable: bool              # False if the response was an empty array


class FieldMappingConfig:
    """Loads and caches field_mapping.yaml."""

    def __init__(self, path: str = "config/field_mapping.yaml"):
        self.path = Path(path)
        with open(self.path, encoding="utf-8") as f:
            self._raw = yaml.safe_load(f) or {}

    def field_group(self, group_name: str) -> list[dict]:
        """Returns the list of {json_path, canonical_name} dicts for a group."""
        return self._raw.get(group_name, [])

    def groups(self) -> list[str]:
        return list(self._raw.keys())


def _coalesce_field(entries: list[dict], key: str) -> tuple[Any, bool]:
    """
    Scans all entries in order, returns the first non-null value found for
    `key`, plus whether the key was present (even as null) in at least one
    entry. If the key is truly absent from every entry, present=False --
    this is the schema-drift signal, distinct from "present but null
    everywhere" (which just means no data, e.g. UniformFee: null for a
    programme with no uniform requirement).
    """
    present_anywhere = False
    non_null_values = []
    for entry in entries:
        if key in entry:
            present_anywhere = True
            if entry[key] is not None and entry[key] != "":
                non_null_values.append(entry[key])

    if not non_null_values:
        return None, present_anywhere

    # Return the first non-null value found; separately flag if any other
    # entry disagreed with a DIFFERENT non-null value (real inconsistency).
    first_value = non_null_values[0]
    return first_value, present_anywhere


def normalize_fee_response(
    raw_response: Any,
    field_mapping: FieldMappingConfig,
    official_code: str = "",
) -> NormalizationResult:
    """
    Main entry point. `raw_response` is whatever GetProgramFeeDetail returned
    (already JSON-decoded) -- expected to be a list, but this function is
    defensive about that (see the shape-handling below), since a bare object
    was also observed in some earlier informal testing.
    """
    inconsistencies: list[dict] = []
    missing_keys: list[dict] = []

    # --- Shape handling ---
    if raw_response is None:
        logger.info("[%s] fee response is None -- treating as not applicable", official_code)
        return NormalizationResult(fields=[], inconsistencies=[], missing_keys=[], applicable=False)

    if isinstance(raw_response, list):
        entries = raw_response
    elif isinstance(raw_response, dict):
        # Defensive: handle a bare object the same as a single-entry array.
        entries = [raw_response]
    else:
        logger.warning("[%s] fee response has unexpected type %s -- treating as not applicable",
                        official_code, type(raw_response))
        return NormalizationResult(fields=[], inconsistencies=[], missing_keys=[], applicable=False)

    if len(entries) == 0:
        # Confirmed rule: empty array = genuinely no fee data (e.g. P8G,
        # a higher doctorate). Not an error.
        logger.info("[%s] fee response is an empty array -- no fee data available "
                    "for this programme (confirmed pattern, e.g. higher doctorates)",
                    official_code)
        return NormalizationResult(fields=[], inconsistencies=[], missing_keys=[], applicable=False)

    # --- Coalesce + schema-drift detection, field by field ---
    fields: list[NormalizedField] = []
    mapping_rows = field_mapping.field_group("fee")

    for row in mapping_rows:
        key = row["json_path"]  # bare field name per the confirmed mapping format
        canonical_name = row["canonical_name"]

        value, present_anywhere = _coalesce_field(entries, key)

        if not present_anywhere:
            # The mapped key doesn't exist in ANY entry -- this is schema
            # drift, not "no data." Worth a distinct log so a real API
            # contract change doesn't get silently swallowed as "unchanged."
            missing_keys.append({"official_code": official_code, "canonical_name": canonical_name, "raw_key": key})
            logger.warning("[%s] mapped field '%s' (raw key '%s') not found in ANY response "
                            "entry -- possible API schema change, not confirmed as 'no data'",
                            official_code, canonical_name, key)
            continue

        # Check for genuine disagreement: multiple distinct non-null values.
        distinct_non_null = {e[key] for e in entries if key in e and e[key] not in (None, "")}
        if len(distinct_non_null) > 1:
            inconsistencies.append({
                "official_code": official_code,
                "canonical_name": canonical_name,
                "raw_key": key,
                "values_found": list(distinct_non_null),
            })
            logger.warning("[%s] field '%s' has GENUINELY DIFFERING non-null values across "
                            "entries: %s -- using first one found, but this should be "
                            "verified against the live page",
                            official_code, canonical_name, distinct_non_null)

        fields.append(NormalizedField(
            field_group="fee",
            canonical_name=canonical_name,
            value=value,
            source_field_names=key,
        ))

    return NormalizationResult(
        fields=fields,
        inconsistencies=inconsistencies,
        missing_keys=missing_keys,
        applicable=True,
    )
