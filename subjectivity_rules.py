"""
subjectivity_rules.py - carrier-specific clean-up of binding subjectivities (Oct 28 2026).

Stefan (Oct 28 2026): on two of his programs the carrier quote prints items that are the
broker's business, not the client's. They never go into the client proposal, and the client
gets 5 days, not 30, for the umbrella paperwork.

  Umbrella / excess with PartnerRe (Alta Risk, LLC is the program manager) - always remove:
    "Alta Risk, LLC does not accept, review, reject or approve certificates of insurance. ..."
    "When applicable, completed copy of Underlying Policy(ies) including Primary, are required
     within 60 days. Failure to provide Alta with requested Underlying Policy information ..."
    "Please return a copy of this completed form to us for our files. Thank you."
    "Net Premiums are due 20 days from Effective Date of this Coverage."
    "Your office is responsible for all Surplus Lines Filings and Taxes."
    "Within 60 Days: - Complete copies of ALL Underlying Policies not placed through Alta Risk, LLC"
  and change 30 days to 5 days for:
    "Within 30 Days: - Signed and selected TRIA"
    "Within 30 Days: - Completed SLF confirmation"
    "Within 30 Days: - Acord 125 and 131 completed and signed by the insured"

  Property with Starr Surplus Lines - always remove the broker's surplus lines obligations:
    "4. Receipt of the completed and signed Surplus Lines Tax Filing Confirmation form warranting
     that the broker/agent will accept full responsibility ... within 10 days of binding ..."
    "1) The broker/agent holds the surplus lines license in the state where coverage will be
     provided to the Insured ..."
    "2) The broker/agent accepts full responsibility for compliance, including ... the filing of
     the surplus lines affidavit ... Surplus Lines Filings Confirmation form."

The rules match on key phrases, not exact text, so numbering ("4.", "1)"), checkboxes and the
model's small rewordings do not matter. They run after extraction (the review screen shows the
cleaned list) and again in the generator, so an item never prints however it got into the data
(a re-added line, an older session, a program template). Idempotent.
"""

import logging
import re

logger = logging.getLogger(__name__)

_PARTNERRE_CARRIER_RE = re.compile(r"partner\s*re\b|partnerre", re.I)
_ALTA_RE = re.compile(r"\balta\s+risk\b", re.I)
_STARR_CARRIER_RE = re.compile(r"\bstarr\b", re.I)

# (name, test) pairs; a test takes the normalised lower-case text of one item.
_PARTNERRE_REMOVE = (
    ("certificates", lambda t: re.search(r"does not (?:accept,?\s*)?(?:review|approve)\b[^.]*?certificates? of insurance", t)),
    ("underlying policies within 60 days",
     lambda t: re.search(r"cop(?:y|ies) of (?:the |all )?underlying polic", t)
     and (re.search(r"\b60\b|sixty", t) or "not placed through alta" in t)),
    ("return a copy of this form", lambda t: "please return a copy of this completed form" in t),
    ("net premiums due 20 days", lambda t: re.search(r"net premiums? (?:are|is) due (?:within )?(?:20|twenty)", t)),
    ("surplus lines filings and taxes", lambda t: re.search(r"your office is responsible for (?:all )?surplus lines", t)),
)
_PARTNERRE_FIVE_DAYS = (
    ("TRIA", lambda t: re.search(r"\btria\b|terrorism", t)),
    ("SLF confirmation", lambda t: re.search(r"\bslf\b|surplus lines? filing confirmation", t)),
    ("ACORD 125 and 131", lambda t: re.search(r"acord\s*125", t)),
)
_STARR_REMOVE = (
    ("surplus lines filing confirmation",
     lambda t: re.search(r"surplus lines (?:tax )?filings? confirmation", t)
     and re.search(r"broker|agent|within 10 days|within ten", t)),
    ("broker holds the surplus lines license", lambda t: re.search(r"broker\s*/?\s*agent holds the surplus lines license", t)),
    ("broker accepts responsibility for surplus lines compliance",
     lambda t: re.search(r"broker\s*/?\s*agent accepts full responsibility", t)),
)
_THIRTY_RE = re.compile(r"\b(within\s+)30(\s+days?\b)", re.I)
_THIRTY_WORDS_RE = re.compile(r"\bthirty\s*\(30\)\s*(days?)\b", re.I)


def _text(item):
    if isinstance(item, dict):
        for k in ("description", "text", "subjectivity", "condition", "item", "name"):
            if isinstance(item.get(k), str) and item[k].strip():
                return item[k]
        return ""
    return item if isinstance(item, str) else ""


def _norm(s):
    t = str(s or "").replace("’", "'").replace("–", "-").replace("—", "-").lower()
    t = re.sub(r"^[\s☐☑☒\-\*•]*(?:\(?\d{1,2}[.)]\s*)?", "", t)
    return re.sub(r"\s+", " ", t).strip()


def _is_umbrella_key(key):
    k = str(key or "").lower()
    return k.startswith("umbrella") or k.startswith("excess_liability") or k in ("excess", "excess_layer_2",
                                                                               "excess_layer_3")


def _is_property_key(key):
    k = str(key or "").lower()
    return k.startswith("property") or k.startswith("excess_property")


def is_partnerre(key, cov):
    if not isinstance(cov, dict):
        return False
    if cov.get("program_template") == "partnerre_umbrella":
        return True
    if not (_is_umbrella_key(key) or _PARTNERRE_CARRIER_RE.search(str(cov.get("carrier") or ""))):
        return False
    if _PARTNERRE_CARRIER_RE.search(str(cov.get("carrier") or "")):
        return True
    blob = " ".join(_text(x) for x in (cov.get("subjectivities") or []) + (cov.get("terms_conditions") or [])
                    if x is not None)
    return bool(_ALTA_RE.search(blob))


def is_starr_property(key, cov):
    if not isinstance(cov, dict):
        return False
    if cov.get("program_template") == "starr_property":
        return True
    return _is_property_key(key) and bool(_STARR_CARRIER_RE.search(str(cov.get("carrier") or "")))


def _set_text(item, new):
    if isinstance(item, dict):
        out = dict(item)
        for k in ("description", "text", "subjectivity", "condition", "item", "name"):
            if isinstance(out.get(k), str) and out[k].strip():
                out[k] = new
                return out
        out["description"] = new
        return out
    return new


def clean_items(key, cov, items, rewrite=True):
    """Return (kept_items, removed_texts, changed_pairs) for one list on one coverage."""
    if not isinstance(items, list):
        return items, [], []
    partnerre = is_partnerre(key, cov)
    starr = is_starr_property(key, cov)
    if not (partnerre or starr):
        return items, [], []
    removes = (_PARTNERRE_REMOVE if partnerre else ()) + (_STARR_REMOVE if starr else ())
    kept, removed, changed = [], [], []
    for it in items:
        raw = _text(it)
        t = _norm(raw)
        if t and any(test(t) for _n, test in removes):
            removed.append(raw.strip())
            continue
        if rewrite and partnerre and t and any(test(t) for _n, test in _PARTNERRE_FIVE_DAYS):
            new = _THIRTY_RE.sub(lambda m: m.group(1) + "5" + m.group(2), raw)
            new = _THIRTY_WORDS_RE.sub(lambda m: "five (5) " + m.group(1), new)
            if new != raw:
                changed.append((raw.strip(), new.strip()))
                it = _set_text(it, new)
        kept.append(it)
    return kept, removed, changed


def clean_coverage(key, cov):
    """Clean one coverage in place: subjectivities (remove + 5-day rewrite) and, for the program
    templates' wording, terms_conditions (remove only). Returns (removed, changed)."""
    removed, changed = [], []
    if not isinstance(cov, dict):
        return removed, changed
    subs, r1, c1 = clean_items(key, cov, cov.get("subjectivities"))
    if r1 or c1:
        cov["subjectivities"] = subs
    terms, r2, _c2 = clean_items(key, cov, cov.get("terms_conditions"), rewrite=False)
    if r2:
        cov["terms_conditions"] = terms
    return r1 + r2, c1


def run_subjectivity_rules(data, items=None):
    """Extraction pass: apply the rules to every coverage and note what changed for the review
    screen. `items` is unused (same signature as the other quote detail passes)."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict):
        return notes
    for key, cov in covs.items():
        removed, changed = clean_coverage(key, cov)
        if not (removed or changed):
            continue
        who = "PartnerRe umbrella" if is_partnerre(key, cov) else "Starr property"
        parts = []
        if removed:
            parts.append(f"left off {len(removed)} broker item{'s' if len(removed) != 1 else ''}")
        if changed:
            parts.append(f"changed {len(changed)} deadline{'s' if len(changed) != 1 else ''} from 30 to 5 days")
        notes.append(f"Subjectivities ({who}): " + " and ".join(parts) + " per the standing rule.")
        logger.info(f"Subjectivity rules [{key}]: removed {removed}; changed {changed}")
    if notes:
        ws = data.get("warnings")
        if not isinstance(ws, list):
            ws = [ws] if isinstance(ws, str) and ws.strip() else []
        data["warnings"] = ws + notes
    return notes
