"""
excess_liquor.py - an excess policy over the liquor liability alone is its own line (Nov 24 2026).

Stefan, Nov 24 2026 (WG Hotel LLC): two excess quotes were uploaded.
  * Alta / PartnerRe: $3,000,000 excess of the Southlake General Liability and Hired Non Owned
    Auto. It builds the umbrella tower.
  * General Star ("Monoline Excess Liquor Liability"): $3,000,000 excess of the Southlake Liquor
    Liability only ($1,000,000 each common cause), limited to the two hotels that serve alcohol.

The model returns every excess quote in the umbrella slots, so General Star printed as
"2nd Excess Layer" under PartnerRe: a second layer of a tower it is not part of. Had the model
listed it first, it would have printed as "Umbrella / Excess 1" and PartnerRe as the second layer,
and an "Each Occurrence" limit on it would have been added to the Total Umbrella / Excess Limit.

This pass moves such a policy to its own line, "excess_liquor_liability" (Excess Liquor
Liability), and closes the gap it leaves in the umbrella layers. It runs after extraction (the
review screen shows the line under its own name) and again in the generator, so a session
extracted before this change prints correctly too. Idempotent.

A policy counts as excess liquor only when the quote says so:
  * its schedule of underlying insurance lists liquor liability and nothing else, or
  * it has no underlying schedule and its layer description or form names it as excess of the
    liquor liability, with no general liability, auto or employers liability named.
An umbrella over the GL, the auto AND the liquor liability stays an umbrella.
"""

import logging
import re

logger = logging.getLogger(__name__)

EXCESS_LIQUOR_KEY = "excess_liquor_liability"
EXCESS_LIQUOR_LABEL = "Excess Liquor Liability"

_LAYER_CHAIN = ("umbrella", "umbrella_layer_2", "umbrella_layer_3", "umbrella_layer_4")
_ALIAS_SLOTS = ("excess_liability", "excess", "excess_layer_2", "excess_layer_3",
                "umbrella_alt_1", "umbrella_alt_2", "umbrella_alt_3")
_LIQUOR_RE = re.compile(r"\bliquor\b|\bdram\s*shop\b", re.I)
_OTHER_LINE_RE = re.compile(
    r"general\s+liab|\bcgl\b|\bgl\b|\bauto|employer|workers|umbrella|products|professional|garage|foreign", re.I)


def _underlying_names(cov):
    """Coverage names on the policy's schedule of underlying insurance."""
    rows = cov.get("underlying_insurance")
    if isinstance(rows, (str, dict)):
        rows = [rows]
    out = []
    for r in rows if isinstance(rows, list) else []:
        if isinstance(r, dict):
            name = str(r.get("coverage") or r.get("type") or r.get("line") or r.get("description") or "").strip()
        else:
            name = str(r or "").strip()
        if name:
            out.append(name)
    return out


def is_excess_liquor(cov):
    """True when the policy is excess of the liquor liability and of nothing else."""
    if not isinstance(cov, dict):
        return False
    names = _underlying_names(cov)
    if names:
        return all(_LIQUOR_RE.search(n) and not _OTHER_LINE_RE.search(n) for n in names)
    text = " ".join(str(cov.get(k) or "") for k in ("layer_description", "policy_form", "coverage_form", "form_type"))
    return bool(re.search(r"excess\s+(?:of\s+)?(?:the\s+)?(?:underlying\s+)?liquor|liquor\s+liability\s+excess|"
                          r"excess\s+liquor", text, re.I)) and not _OTHER_LINE_RE.search(text)


def _carrier_words(name):
    return {w for w in re.findall(r"[a-z]+", str(name or "").lower())
            if w not in ("insurance", "company", "co", "inc", "the", "of", "corporation", "corp", "ltd", "llc")}


def _describe(cov):
    lim = ""
    for l in cov.get("limits") or []:
        if isinstance(l, dict):
            m = re.search(r"\$\s?[\d,]+", str(l.get("limit") or ""))
            if m:
                lim = m.group(0).replace(" ", "")
                break
    carrier = str(cov.get("carrier") or "").strip() or "The policy"
    return f"{carrier} ({lim} excess of the liquor liability)" if lim else f"{carrier} (excess of the liquor liability)"


def run_excess_liquor(data, items=None):
    """Move an excess policy that sits over the liquor liability alone out of the umbrella
    slots, into its own line. Returns review notes. `items` is unused (same signature as the
    other quote detail passes)."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict):
        return notes
    moved = None
    for key in _LAYER_CHAIN + _ALIAS_SLOTS:
        cov = covs.get(key)
        if not is_excess_liquor(cov):
            continue
        if isinstance(covs.get(EXCESS_LIQUOR_KEY), dict) and covs[EXCESS_LIQUOR_KEY].get("carrier"):
            # a second one has no line of its own to go to; it stays where it is and is named
            notes.append(f"Excess check — {_describe(cov)} is a second excess liquor policy and is still shown "
                         f"with the umbrella layers. Check how it should be presented.")
            continue
        covs[EXCESS_LIQUOR_KEY] = covs.pop(key)
        moved = covs[EXCESS_LIQUOR_KEY]
        notes.append(f"Excess check — {_describe(moved)} is shown as {EXCESS_LIQUOR_LABEL}, a line of its own, "
                     f"not as a layer of the umbrella tower.")
        logger.info(f"Excess liquor: moved {key} ({moved.get('carrier')}) to {EXCESS_LIQUOR_KEY}")
    if moved is not None:
        # close the gap: the remaining layers move up, everything else keeps its place
        present = [k for k in _LAYER_CHAIN if k in covs]
        rename = {old: new for old, new in zip(present, _LAYER_CHAIN) if old != new}
        if rename:
            reordered = {rename.get(k, k): v for k, v in covs.items()}
            covs.clear()
            covs.update(reordered)
        # payment rows the model wrote for it as an umbrella or excess layer
        words = _carrier_words(moved.get("carrier"))
        for p in data.get("payment_options") or []:
            if (isinstance(p, dict) and words and words <= _carrier_words(p.get("carrier"))
                    and re.search(r"umbrella|excess", str(p.get("coverage_type") or ""), re.I)):
                p["coverage_type"] = EXCESS_LIQUOR_LABEL
    if notes:
        ws = data.get("warnings")
        if not isinstance(ws, list):
            ws = [ws] if isinstance(ws, str) and ws.strip() else []
        data["warnings"] = [w for w in ws if w not in notes] + notes
    return notes
