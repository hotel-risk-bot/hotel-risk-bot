"""
coverage_recovery.py — per-file line-of-coverage safety net for the proposal extractor.

Why (Sep 16 2026, Jalabapa / City Express by Marriott Yulee): a seven-file submission came
back with no Umbrella and no Cyber even though the Great Point (StarStone) umbrella quote and
the At-Bay cyber quote were both uploaded and both extract perfectly on their own. One GPT
call asked to return an entire program inside one output budget dropped the two smallest
quotes. Nothing downstream noticed, so the proposal simply omitted two lines of coverage.

This module makes the omission impossible to miss:
  1. classify_file(text, filename) -> set of coverage lines the file is a quote for
     (deterministic keyword scoring — no model involved);
  2. run_coverage_recovery(data, items, extract_fn) compares those lines against the
     coverage keys actually extracted; for every line that has a quote file but no
     coverage, it calls extract_fn(file_text, filename) — a single-file run of the SAME
     full extraction prompt — and merges the returned coverage(s) in;
  3. every recovery, and every file whose line is still missing afterwards, is written
     to data["warnings"] so the review screen shows it.
"""
import logging
import re

logger = logging.getLogger(__name__)

# Line -> (phrases that score for it, phrases that score against it)
_LINE_SIGNALS = {
    "umbrella": (
        ("umbrella", "excess liability", "following form excess", "commercial excess",
         "schedule of underlying", "underlying insurance", "lead umbrella", "retained limit"),
        ("statement of values", "commercial property coverage", "building and personal property"),
    ),
    "cyber": (
        ("cyber insurance quote", "cyber liability", "cyber insurance", "network security",
         "privacy liability", "ransomware", "data breach", "cyber policy", "at-bay", "coalition"),
        (),
    ),
    "general_liability": (
        ("commercial general liability", "general liability coverage form", "cg 00 01",
         "premises/operations", "products/completed operations aggregate", "general aggregate"),
        ("excess liability", "umbrella"),
    ),
    "property": (
        ("statement of values", "total insured value", "building and personal property",
         "business income", "replacement cost", "named storm", "wind/hail", "commercial property",
         "cp 00 10", "coinsurance"),
        ("cyber", "workers compensation"),
    ),
    "workers_compensation": (
        ("workers compensation", "workers' compensation", "employers liability", "employers' liability",
         "class code", "wc 00 00"),
        (),
    ),
    "epli": (
        ("employment practices liability", "employment practices", "epli", "wrongful termination",
         "third party discrimination", "wage and hour"),
        (),
    ),
    "crime": (
        ("crime coverage", "employee theft", "employee dishonesty", "forgery or alteration",
         "funds transfer fraud", "commercial crime"),
        (),
    ),
    "flood": (
        ("flood insurance", "standard flood insurance policy", "nfip", "flood zone", "private flood",
         "flood coverage", "ground up x flood", "flood quote"),
        (),
    ),
    "commercial_auto": (
        ("business auto", "commercial auto", "auto liability", "hired and non-owned auto",
         "garage keepers", "ca 00 01"),
        ("umbrella", "excess liability"),
    ),
    "equipment_breakdown": (
        ("equipment breakdown", "boiler and machinery", "mechanical breakdown"),
        (),
    ),
    "liquor_liability": (
        ("liquor liability coverage", "liquor liability quote", "dram shop"),
        (),
    ),
}

# Which extracted coverage keys satisfy a line.
_LINE_KEY_PREFIXES = {
    "umbrella": ("umbrella", "excess_liability", "excess"),
    "cyber": ("cyber",),
    "general_liability": ("general_liability", "gl"),
    "property": ("property", "excess_property"),
    "workers_compensation": ("workers_comp", "workers_compensation"),
    "epli": ("epli", "employment"),
    "crime": ("crime",),
    "flood": ("flood",),
    "commercial_auto": ("commercial_auto", "auto"),
    "equipment_breakdown": ("equipment_breakdown", "boiler"),
    "liquor_liability": ("liquor",),
}

# GPT sometimes names keys differently; map them onto the canonical ones on merge.
_KEY_ALIASES = {
    "excess_liability": "umbrella", "excess": "umbrella", "umbrella_liability": "umbrella",
    "umbrella_excess": "umbrella", "excess_umbrella": "umbrella",
    "cyber_liability": "cyber", "cyber_insurance": "cyber",
    "gl": "general_liability", "commercial_general_liability": "general_liability",
    "workers_comp": "workers_compensation", "wc": "workers_compensation",
    "employment_practices": "epli", "employment_practices_liability": "epli",
    "auto": "commercial_auto", "business_auto": "commercial_auto",
    "boiler_machinery": "equipment_breakdown",
    "liquor": "liquor_liability",
}

_MIN_SCORE = 3


def classify_file(text: str, filename: str = "") -> dict:
    """Return {line: score} for every line this file looks like a quote for.
    A file can carry more than one line (package quotes). Scores count phrase
    occurrences in the text (capped per phrase) plus a filename bonus, minus
    occurrences of the line's anti-phrases."""
    low = (text or "").lower()
    fn = (filename or "").lower()
    if not low.strip():
        return {}
    scores = {}
    for line, (pos, neg) in _LINE_SIGNALS.items():
        s = 0
        for p in pos:
            s += min(low.count(p), 6)
        for n in neg:
            s -= min(low.count(n), 6)
        # filename hints are strong signals
        for p in pos:
            if p in fn:
                s += 4
        if line == "umbrella" and re.search(r"\b(xs|umb|umbrella|excess)\b", fn):
            s += 4
        if line == "cyber" and re.search(r"\b(cy|cyber|atbay|at-bay|coalition)\b", fn):
            s += 4
        if line == "workers_compensation" and re.search(r"\bwc\b", fn):
            s += 4
        if s >= _MIN_SCORE:
            scores[line] = s
    # Property and GL both mention each other's words; keep only the dominant one
    # unless both are strong (a true package quote).
    if "property" in scores and "general_liability" in scores:
        p, g = scores["property"], scores["general_liability"]
        if p >= 3 * g:
            scores.pop("general_liability")
        elif g >= 3 * p:
            scores.pop("property")
    # A cyber quote mentions "liability" a lot; drop GL when cyber dominates.
    if "cyber" in scores and "general_liability" in scores and scores["cyber"] >= 2 * scores["general_liability"]:
        scores.pop("general_liability")
    # An umbrella quote lists its underlying GL; drop GL when umbrella dominates.
    if "umbrella" in scores and "general_liability" in scores and scores["umbrella"] >= scores["general_liability"]:
        scores.pop("general_liability")
    return scores


def _has_line(coverages: dict, line: str) -> bool:
    prefixes = _LINE_KEY_PREFIXES.get(line, (line,))
    for k, v in (coverages or {}).items():
        if not isinstance(v, dict):
            continue
        kl = str(k).lower()
        if any(kl == p or kl.startswith(p + "_") or kl.startswith(p) for p in prefixes):
            if v.get("carrier") or v.get("premium") or v.get("total_premium") or v.get("limits") or v.get("coverage_limits"):
                return True
    return False


def _canonical_key(key: str) -> str:
    kl = str(key or "").lower().strip()
    return _KEY_ALIASES.get(kl, kl)


def _merge_recovered(data: dict, recovered: dict, wanted_lines: set, filename: str) -> list:
    """Merge coverages from a single-file extraction for the lines that were missing.
    Returns the list of keys added."""
    added = []
    covs = data.setdefault("coverages", {})
    rec_covs = recovered.get("coverages") or {}
    if isinstance(rec_covs, list):
        rec_covs = {str(c.get("coverage_type", f"cov_{i}")).lower().replace(" ", "_"): c
                    for i, c in enumerate(rec_covs) if isinstance(c, dict)}
    for rkey, rcov in rec_covs.items():
        if not isinstance(rcov, dict):
            continue
        if not (rcov.get("carrier") or rcov.get("premium") or rcov.get("total_premium")):
            continue
        ckey = _canonical_key(rkey)
        line = None
        for ln, prefixes in _LINE_KEY_PREFIXES.items():
            if any(ckey == p or ckey.startswith(p) for p in prefixes):
                line = ln
                break
        if line not in wanted_lines:
            continue
        if _has_line(covs, line):
            continue
        target = ckey
        if target in covs and isinstance(covs[target], dict) and covs[target]:
            # occupied by an empty/unknown shell? only overwrite empties
            if covs[target].get("carrier") or covs[target].get("premium"):
                continue
        rcov["_recovered_from"] = filename
        covs[target] = rcov
        added.append(target)
    # Named insureds / subjectivities from the recovered file are usually already
    # present; only fill client_info fields that are empty.
    ci = data.setdefault("client_info", {})
    for k, v in (recovered.get("client_info") or {}).items():
        if v and not ci.get(k):
            ci[k] = v
    return added


def run_coverage_recovery(data: dict, items: list, extract_fn) -> list:
    """items: [{filename, text}] (full, untruncated per-file text).
    extract_fn(text, filename) -> dict in the main extraction schema, or None.
    Returns warning strings (also appended to data['warnings'])."""
    warnings = []
    if not isinstance(data, dict):
        return warnings
    covs = data.setdefault("coverages", {})
    if not isinstance(covs, dict):
        return warnings
    for it in items or []:
        fname = it.get("filename", "") or ""
        text = it.get("text", "") or ""
        if fname.lower().endswith((".xlsx", ".xls", ".xlsb", ".csv")):
            continue  # SOVs are not quotes
        lines = classify_file(text, fname)
        if not lines:
            continue
        missing = {ln for ln in lines if not _has_line(covs, ln)}
        if not missing:
            logger.info(f"Coverage recovery: '{fname}' -> {sorted(lines)} all present")
            continue
        logger.warning(f"Coverage recovery: '{fname}' looks like {sorted(missing)} but no such coverage was extracted - re-extracting this file alone")
        recovered = None
        try:
            recovered = extract_fn(text, fname)
        except Exception as e:
            logger.error(f"Coverage recovery extraction failed for '{fname}': {e}")
        added = _merge_recovered(data, recovered or {}, missing, fname) if isinstance(recovered, dict) else []
        still_missing = {ln for ln in missing if not _has_line(covs, ln)}
        if added:
            w = (f"Coverage recovery — '{fname}' is a {', '.join(sorted(missing))} quote that the main extraction "
                 f"skipped; recovered {', '.join(added)} from a single-file re-extraction. Review it.")
            warnings.append(w)
            logger.warning(w)
        if still_missing:
            w = (f"REVIEW REQUIRED — '{fname}' looks like a {', '.join(sorted(still_missing))} quote but no "
                 f"{', '.join(sorted(still_missing))} coverage could be extracted. Add it manually before generating.")
            warnings.append(w)
            logger.warning(w)
    if warnings:
        data.setdefault("warnings", [])
        if isinstance(data["warnings"], list):
            data["warnings"].extend(warnings)
    return warnings
