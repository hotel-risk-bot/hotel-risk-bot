"""
forms_schedule_audit.py — deterministic Forms & Endorsements completeness check.

Why this exists (Sep 15 2026): the Neptune Lodging Partners proposal listed 11 of the
39 forms on the Houston Specialty excess quote. GPT had extracted them, and then two
prefix/description cleaners silently threw away every HSIC CX ES exclusion (Assault or
Battery, Total Firearms, Sublimits or Reduced Limits in Underlying Insurance, Sexual
Misconduct, ...). A proposal that omits an exclusion is an E&O exposure, so the forms
list can no longer depend on an LLM being exhaustive or on a prefix table being right.

This module reads the carrier quote text directly, finds every row of the forms
schedule that carries a real form number, and reconciles it against the extracted
forms_endorsements list:

  * rows on the quote's schedule that the extractor missed are APPENDED
    (flagged schedule_verified=True so downstream filters leave them alone);
  * the coverage gets a `forms_audit` dict the review UI and the docx use to show
    "quote lists N numbered forms / extracted M / added K" and any residual gap;
  * a warning string is pushed onto data["warnings"].

It is deliberately conservative: it only ever ADDS rows that literally appear with a
form number in the coverage's own quote file, and it never removes anything.
"""
import logging
import re

logger = logging.getLogger(__name__)

# A complete form number, matched against the LAST column of a schedule row.
# ISO ("CX 21 13 04 13", "CG 20 29 12 19", "IL 09 85 12 20"), carrier-prefixed
# ("HSIC CX ES 01 44 04 21", "LBGL 21 53 12 24", "TSIC 004 0424", "FORMS - SCHED 08 12"),
# dashed ("LIA-19182-0724", "FUT-SS 01 22") and edition-in-parentheses variants
# ("LB TRIA (12/2022)", "DS PN Annual (02-2022)").
_FORM_NUMBER_FULL_RE = re.compile(
    r"^(?:"
    r"(?:[A-Z]{1,6}(?:\s*-\s*[A-Z]{1,6})?\s+){0,4}"          # alpha prefix tokens
    r"(?:[A-Z]{1,3}\s?)?\d{2,5}(?:[\s\-/]\d{2,4}){0,4}"      # numeric groups (+ optional letter)
    r"|[A-Z]{2,6}(?:-[A-Z0-9]{2,8}){1,3}"                     # dashed carrier codes
    r"|[A-Z]{2,6}(?:\s+[A-Za-z]{2,8}){0,2}\s*\(\d{2}[-/]\d{2,4}\)"  # edition in parentheses
    r")$"
)
# Fallback for single-spaced text (no column gap): a short, tightly-formed ISO-style
# number at the end of the line — at most two alpha tokens so it cannot swallow the
# tail of an upper-case description.
_FORM_NUMBER_TAIL_RE = re.compile(
    r"\s(?P<num>[A-Z]{2,6}(?:\s[A-Z]{2,4})?\s\d{2}(?:\s\d{2}){2,3})\s*$"
)

_SCHEDULE_HEADER_RE = re.compile(
    r"^\s*(FORMS?\s*:?|FORMS?\s+(?:AND|&)\s+ENDORSEMENTS?|SCHEDULE\s+OF\s+FORMS?(?:\s+AND\s+ENDORSEMENTS?)?|"
    r"FORMS?\s+SCHEDULE|ENDORSEMENTS?\s+SCHEDULE|APPLICABLE\s+FORMS?|POLICY\s+FORMS?(?:\s+AND\s+ENDORSEMENTS?)?)\s*:?\s*$",
    re.I,
)
_SCHEDULE_END_RE = re.compile(
    r"^\s*(CONDITIONS?\s*:|SUBJECTIVITIES|SUBJECT\s+TO\s*:|PREMIUM\s*(?:&|AND)?\s*FEES|TERMS\s+AND\s+CONDITIONS|"
    r"THIS\s+ENDORSEMENT\s+CHANGES|BINDING\s+REQUIREMENTS|NOTES?\s*:)",
    re.I,
)
_COLUMN_HEADER_RE = re.compile(r"^\s*FORM\s+(NAME|TITLE|NUMBER|#|DESCRIPTION)\b", re.I)
_COLUMN_SPLIT_RE = re.compile(r"\s{2,}|\t")


def _compact(num: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(num or "").upper())


def _is_form_number(tok: str) -> bool:
    t = (tok or "").strip()
    if not t or len(t) > 40 or not re.search(r"\d", t):
        return False
    return bool(_FORM_NUMBER_FULL_RE.match(t))


def _split_row(line: str):
    """Return (description, form_number) when `line` is a schedule row, else None."""
    t = line.rstrip()
    if not t.strip():
        return None
    cols = [c for c in _COLUMN_SPLIT_RE.split(t.strip()) if c.strip()]
    if len(cols) >= 2:
        num = cols[-1].strip()
        if _is_form_number(num):
            desc = " ".join(c.strip() for c in cols[:-1])
            return desc, num
        # "FORM NUMBER   DESCRIPTION" layouts (number first)
        num = cols[0].strip()
        if _is_form_number(num) and not _is_form_number(cols[-1].strip()):
            return " ".join(c.strip() for c in cols[1:]), num
        return None
    m = _FORM_NUMBER_TAIL_RE.search(t)
    if m:
        return t[: m.start("num")].strip(), m.group("num").strip()
    return None


def _looks_like_continuation(line: str) -> bool:
    """A wrapped second line of a form description: short, no form number, no colon,
    single column (Houston Specialty wraps long titles onto a new line)."""
    t = line.strip()
    if not t or len(t) > 80 or ":" in t or "$" in t:
        return False
    if _split_row(t):
        return False
    if _SCHEDULE_HEADER_RE.match(t) or _SCHEDULE_END_RE.match(t) or _COLUMN_HEADER_RE.match(t):
        return False
    letters = [c for c in t if c.isalpha()]
    if len(letters) < 3:
        return False
    if len(_COLUMN_SPLIT_RE.split(t)) > 1:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return upper_ratio > 0.85 or t[0].islower() or t.startswith("(") or t.endswith(")")


def parse_forms_schedule(text: str) -> list[dict]:
    """Return [{form_number, description}] for every schedule row in `text`.
    Wrapped titles are joined with the following line. Rows found outside an
    explicit schedule block are accepted only when they sit in a run of at least
    3 numbered rows within a few lines of each other (what a schedule table looks
    like once the PDF is flattened to text)."""
    if not text:
        return []
    lines = text.splitlines()
    rows = []
    in_block = False
    for i, raw in enumerate(lines):
        line = raw.rstrip()
        if _SCHEDULE_HEADER_RE.match(line):
            in_block = True
            continue
        if in_block and _SCHEDULE_END_RE.match(line):
            in_block = False
        if _COLUMN_HEADER_RE.match(line):
            continue
        split = _split_row(line)
        if not split:
            continue
        desc, num = split
        desc = desc.strip(" \t-–—|")
        if not desc or "$" in desc or len(desc) > 140 or desc.count(".") > 2:
            continue
        if re.search(r"\b(POLICY|QUOTE|REFERENCE|ACCOUNT)\s*(NO|NUMBER|#)", desc, re.I):
            continue
        j = i + 1
        if j < len(lines) and _looks_like_continuation(lines[j]):
            desc = f"{desc} {lines[j].strip()}"
        rows.append({"form_number": num, "description": " ".join(desc.split()), "line_no": i, "in_block": in_block})

    kept, run = [], []
    for r in rows:
        if r["in_block"]:
            if run:
                if len(run) >= 3:
                    kept.extend(run)
                run = []
            kept.append(r)
        else:
            if run and r["line_no"] - run[-1]["line_no"] > 4:
                if len(run) >= 3:
                    kept.extend(run)
                run = []
            run.append(r)
    if len(run) >= 3:
        kept.extend(run)

    seen, out = set(), []
    for r in kept:
        c = _compact(r["form_number"])
        if c in seen:
            continue
        seen.add(c)
        out.append({"form_number": r["form_number"], "description": r["description"]})
    return out


def pick_source_text(items: list, carrier: str, coverage_key: str) -> tuple[str, str]:
    """Choose the uploaded file that most likely IS this coverage's quote.
    Scores each file on carrier-name mentions and coverage keywords."""
    if not items:
        return "", ""
    key = (coverage_key or "").lower()
    if key.startswith("umbrella") or key.startswith("excess_liab"):
        kws = ("umbrella", "excess liability", "commercial excess", "following form", "schedule of underlying")
        anti = ("commercial property", "statement of values", "workers compensation")
    elif key.startswith("general_liability"):
        kws = ("general liability", "commercial general liability", "cgl")
        anti = ("excess liability", "commercial excess", "umbrella", "statement of values")
    elif key.startswith("liquor"):
        kws = ("liquor liability",)
        anti = ("umbrella", "commercial property")
    elif key.startswith("commercial_auto"):
        kws = ("business auto", "commercial auto", "auto liability")
        anti = ("umbrella", "commercial property")
    else:
        return "", ""
    carrier_tokens = [t.lower().strip(",.") for t in (carrier or "").split() if len(t) >= 4
                      and t.lower() not in ("insurance", "company", "specialty", "group", "limited", "underwriters")]
    best, best_score, best_name = "", 0, ""
    for it in items:
        txt = it.get("text") or ""
        if not txt:
            continue
        low = txt.lower()
        score = sum(low.count(k) for k in kws) * 3
        score += sum(min(low.count(t), 25) for t in carrier_tokens) * 2
        score -= sum(low.count(a) for a in anti) * 2
        if score > best_score:
            best, best_score, best_name = txt, score, it.get("filename", "")
    if best_score < 6:
        return "", ""
    return best, best_name


def reconcile_coverage_forms(cov: dict, source_text: str, source_name: str, coverage_key: str,
                             reject_fn=None) -> dict:
    """Merge schedule rows the extractor missed into cov['forms_endorsements'].
    `reject_fn(form_dict) -> bool` may veto a candidate (used to keep sibling-coverage
    rows out when one PDF holds two quotes). Returns the audit dict."""
    schedule = parse_forms_schedule(source_text)
    existing = cov.get("forms_endorsements") or []
    if not isinstance(existing, list):
        existing = []
    have = set()
    for f in existing:
        if isinstance(f, dict):
            c = _compact(f.get("form_number"))
            if c:
                have.add(c)
    added, vetoed = [], []
    for row in schedule:
        c = _compact(row["form_number"])
        if not c:
            continue
        # Edition-date variants: "CX 21 13" vs "CX 21 13 04 13" count as present.
        if c in have or any(h.startswith(c) or c.startswith(h) for h in have if len(h) >= 6 and len(c) >= 6):
            continue
        cand = {"form_number": row["form_number"], "description": row["description"], "schedule_verified": True}
        if reject_fn and reject_fn(cand):
            vetoed.append(cand)
            continue
        added.append(cand)
        have.add(c)
    if added:
        merged = existing + added
        # Put the list back in the quote's own schedule order when the schedule
        # accounts for most of it; anything not on the schedule keeps its place at the end.
        order = {_compact(r["form_number"]): i for i, r in enumerate(schedule)}

        def _pos(f):
            if not isinstance(f, dict):
                return len(order) + 1
            c = _compact(f.get("form_number"))
            if c in order:
                return order[c]
            for k, i in order.items():
                if len(k) >= 6 and len(c) >= 6 and (k.startswith(c) or c.startswith(k)):
                    return i
            return len(order) + 1
        if len([f for f in merged if _pos(f) <= len(order)]) >= 0.8 * len(merged):
            merged = sorted(merged, key=_pos)
        cov["forms_endorsements"] = merged
    audit = {
        "source_file": source_name,
        "schedule_rows": len(schedule),
        "extracted_before": len(existing),
        "added": [{"form_number": a["form_number"], "description": a["description"]} for a in added],
        "vetoed": [{"form_number": v["form_number"], "description": v["description"]} for v in vetoed],
        "final_count": len(cov.get("forms_endorsements") or []),
    }
    cov["forms_audit"] = audit
    return audit


def run_forms_schedule_audit(data: dict, items: list, reject_fn_factory=None) -> list[str]:
    """Entry point called at the end of extraction. Returns warning strings (also
    appended to data['warnings'])."""
    warnings = []
    covs = data.get("coverages", {}) if isinstance(data, dict) else {}
    if not isinstance(covs, dict):
        return warnings
    targets = [k for k in covs if k.startswith("umbrella") or k.startswith("general_liability")
               or k.startswith("liquor") or k.startswith("commercial_auto")]
    for key in targets:
        cov = covs.get(key)
        if not isinstance(cov, dict):
            continue
        carrier = str(cov.get("carrier") or cov.get("carrier_name") or "")
        text, fname = pick_source_text(items, carrier, key)
        if not text:
            logger.info(f"Forms audit: no distinct source file identified for {key} ({carrier}) - skipped")
            continue
        reject_fn = reject_fn_factory(key) if reject_fn_factory else None
        audit = reconcile_coverage_forms(cov, text, fname, key, reject_fn)
        msg = (f"{key}: quote '{fname}' lists {audit['schedule_rows']} numbered forms; "
               f"extractor had {audit['extracted_before']}")
        if audit["added"]:
            names = "; ".join(f"{a['form_number']} {a['description']}" for a in audit["added"][:12])
            more = "" if len(audit["added"]) <= 12 else f" (+{len(audit['added']) - 12} more)"
            w = f"Forms schedule audit — {msg}. ADDED {len(audit['added'])} missing rows: {names}{more}"
            warnings.append(w)
            logger.warning(w)
        else:
            logger.info(f"Forms schedule audit — {msg}. Nothing missing.")
        if audit["vetoed"]:
            names = "; ".join(f"{v['form_number']} {v['description']}" for v in audit["vetoed"][:8])
            w = f"Forms schedule audit — {key}: {len(audit['vetoed'])} schedule rows NOT added because they look like another coverage's forms — verify: {names}"
            warnings.append(w)
            logger.warning(w)
    if warnings:
        data.setdefault("warnings", [])
        if isinstance(data["warnings"], list):
            data["warnings"].extend(warnings)
    return warnings
