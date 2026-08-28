"""
Client-facing renewal email builder for the proposal generator.

Takes the same extracted/edited proposal data used to build the DOCX and produces
a copy-paste email to the insured: a short coverage summary (line, carrier, key
limit) plus the premium comparison (total and per-line increase/decrease) pulled
from the same expiring-vs-proposed logic the Premium Summary page uses.

Two entry points:
  build_email_context(data) -> dict   deterministic facts, no AI
  draft_email(context, instruction, previous) -> {"subject","body"}   AI-written
"""

import os
import re
import logging

from proposal_generator import (
    _clean_carrier_name,
    _parse_currency,
    fmt_currency_cents,
)

logger = logging.getLogger(__name__)

# Mirrors the Premium Summary display names so the email and the document agree.
COVERAGE_LABELS = {
    "property": "Property",
    "property_alt_1": "Property (Option 2)",
    "property_alt_2": "Property (Option 3)",
    "excess_property": "Excess Property (Layer 1)",
    "excess_property_2": "Excess Property (Layer 2)",
    "general_liability": "General Liability",
    "general_liability_alt_1": "General Liability (Option 2)",
    "general_liability_alt_2": "General Liability (Option 3)",
    "umbrella": "Umbrella / Excess 1",
    "umbrella_alt_1": "Umbrella / Excess 2",
    "umbrella_alt_2": "Umbrella / Excess 3",
    "umbrella_alt_3": "Umbrella / Excess 4",
    "umbrella_layer_2": "2nd Excess Layer",
    "umbrella_layer_3": "3rd Excess Layer",
    "umbrella_layer_4": "4th Excess Layer",
    "excess_liability": "Excess Liability",
    "excess": "Excess Liability",
    "workers_comp": "Workers Compensation",
    "workers_compensation": "Workers Compensation",
    "workers_compensation_alt_1": "Workers Comp (Option 2)",
    "commercial_auto": "Commercial Auto",
    "flood": "Flood",
    "wind": "Wind / Named Storm",
    "earthquake": "Earthquake",
    "epli": "EPLI",
    "cyber": "Cyber",
    "cyber_alt_1": "Cyber (Option 2)",
    "terrorism": "Terrorism / TRIA",
    "crime": "Crime",
    "inland_marine": "Inland Marine",
    "equipment_breakdown": "Equipment Breakdown",
    "liquor_liability": "Liquor Liability",
    "innkeepers": "Innkeepers Legal Liability",
    "environmental": "Environmental",
}

_EXPIRING_KEY_MAP = {
    "general_liability": ["general_liability", "gl"],
    "workers_comp": ["workers_comp", "workers_compensation"],
    "workers_compensation": ["workers_comp", "workers_compensation"],
    "commercial_auto": ["commercial_auto", "auto"],
    "equipment_breakdown": ["equipment_breakdown", "boiler_machinery"],
}

# Limit descriptions worth quoting in a client email, most representative first.
_KEY_LIMIT_HINTS = [
    "total insured value", "tiv", "blanket limit", "limit of liability",
    "each occurrence", "occurrence limit", "per occurrence",
    "general aggregate", "aggregate limit", "each accident",
    "employers liability", "policy limit", "limit",
]


def _num(val):
    if val in (None, "", 0):
        return 0.0
    if isinstance(val, str):
        return _parse_currency(val) or 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _fmt_delta(amount):
    """Sign-aware currency so a credit reads -$2,750.00, not $-2,750.00."""
    if amount < 0:
        return "-" + fmt_currency_cents(abs(amount))
    return fmt_currency_cents(amount)


def _expiring_for(key, expiring, expiring_details):
    """Same lookup order the Premium Summary page uses."""
    if expiring_details:
        for k in [key] + _EXPIRING_KEY_MAP.get(key, []):
            detail = expiring_details.get(k)
            if isinstance(detail, dict):
                amt = _num(detail.get("premium") or detail.get("total_premium"))
                if amt:
                    return amt
    if expiring:
        for k in [key] + _EXPIRING_KEY_MAP.get(key, []):
            amt = _num(expiring.get(k))
            if amt:
                return amt
    return 0.0


def _deductibles(cov):
    """Flatten a coverage's deductible schedule into 'label amount' strings."""
    out = []
    for d in (cov.get("deductibles") or []):
        if isinstance(d, dict):
            desc = str(d.get("description") or d.get("peril") or "").strip()
            amt = str(d.get("amount") or d.get("deductible") or "").strip()
            if desc and amt:
                out.append(f"{desc}: {amt}")
            elif amt:
                out.append(amt)
        elif isinstance(d, str) and d.strip():
            out.append(d.strip())
    return out


def _key_limit(cov):
    """Pick the one limit worth putting in front of an owner."""
    if not isinstance(cov, dict):
        return ""
    tiv = _num(cov.get("tiv"))
    if tiv:
        return fmt_currency_cents(tiv).replace(".00", "") + " TIV"

    limits = cov.get("limits") or []
    if isinstance(limits, list) and limits:
        pairs = []
        for item in limits:
            if isinstance(item, dict):
                desc = str(item.get("description") or item.get("coverage") or "").strip()
                amt = str(item.get("limit") or item.get("amount") or "").strip()
                if desc and amt:
                    pairs.append((desc, amt))
        for hint in _KEY_LIMIT_HINTS:
            for desc, amt in pairs:
                if hint in desc.lower():
                    return amt
        if pairs:
            return pairs[0][1]
    return ""



# Stefan's presentation order for coverage lines. Anything unlisted sorts after,
# keeping its relative order. Property first, casualty next, ancillary last.
_LINE_ORDER = [
    "property", "property_alt_1", "property_alt_2",
    "excess_property", "excess_property_2",
    "general_liability", "general_liability_alt_1", "general_liability_alt_2",
    "liquor_liability", "innkeepers",
    "umbrella", "umbrella_alt_1", "umbrella_alt_2", "umbrella_alt_3",
    "umbrella_layer_2", "umbrella_layer_3", "umbrella_layer_4",
    "excess_liability", "excess",
    "commercial_auto",
    "workers_comp", "workers_compensation", "workers_compensation_alt_1",
    "epli",
    "cyber", "cyber_alt_1",
    "crime",
    "equipment_breakdown",
    "inland_marine",
    "flood", "wind", "earthquake",
    "terrorism",
    "environmental",
]


# Trailing corporate boilerplate an owner does not need to read. Stripped
# iteratively so "Starr Surplus Lines Insurance Company" -> "Starr Surplus Lines"
# while "Berkley Specialty" keeps the word that actually identifies the paper.
_CARRIER_TAIL = {
    "company", "companies", "co", "co.", "corporation", "corp", "corp.",
    "incorporated", "inc", "inc.", "insurance", "ins", "ins.", "insco",
    "group", "ltd", "ltd.", "llc", "plc", "syndicate", "syndicates",
}


def _short_carrier(name):
    """Conversational carrier name for a client email."""
    n = (name or "").strip().rstrip(".,")
    if not n:
        return ""
    if "lloyd" in n.lower():
        return "Lloyd's"
    words = n.split()
    while len(words) > 1 and words[-1].lower().strip(".,") in _CARRIER_TAIL:
        words.pop()
    if words and words[0].lower() == "the" and len(words) > 1:
        words.pop(0)
    out = " ".join(words).rstrip(".,")
    return out if len(out) > 1 else n


def _email_label(key, label):
    """Prose-friendly line name. COVERAGE_LABELS stays a mirror of the Premium
    Summary page for parity; an email should not read "Umbrella / Excess 1"."""
    out = re.sub(r"\s+1$", "", label)          # only layer, drop the ordinal
    out = out.replace(" / ", "/")              # Terrorism/TRIA, Umbrella/Excess
    return out


def _line_rank(key):
    try:
        return _LINE_ORDER.index(key)
    except ValueError:
        return len(_LINE_ORDER)

def _norm_addr(loc):
    """Normalized address key so a hotel and its signs are not counted as separate sites."""
    parts = [
        str(loc.get("address") or "").strip().lower(),
        str(loc.get("city") or "").strip().lower(),
        str(loc.get("state") or "").strip().lower(),
    ]
    key = " ".join(p for p in parts if p)
    return " ".join(key.split())


def _location_summary(data):
    """Distinct physical locations and total insured value.

    Counts unique addresses, not schedule rows: a property schedule routinely
    carries separate rows for a building and its signs at the same address, and
    telling an owner they have three locations when they have one hotel is wrong.
    """
    locs = data.get("locations") or []
    keys = set()
    tiv_total = 0.0
    for loc in locs:
        if not isinstance(loc, dict):
            continue
        k = _norm_addr(loc)
        if k:
            keys.add(k)
        tiv_total += _num(loc.get("tiv"))

    count = len(keys) if keys else len(locs)

    if not tiv_total:
        for cov in (data.get("coverages") or {}).values():
            if isinstance(cov, dict):
                tiv_total = max(tiv_total, _num(cov.get("tiv")))
    return count, tiv_total


def build_email_context(data, answers=None):
    """Deterministic facts for the email. No AI, no rounding surprises."""
    data = data or {}
    answers = answers or {}
    expiring_carriers = answers.get("expiring_carriers") or {}
    client = data.get("client_info") or {}
    coverages = data.get("coverages") or {}
    expiring = data.get("expiring_premiums") or {}
    expiring_details = data.get("expiring_details") or {}
    has_expiring = bool(expiring or expiring_details)

    lines, optional_lines = [], []
    total_proposed = 0.0
    total_expiring = 0.0

    for key, cov in coverages.items():
        if not isinstance(cov, dict):
            continue
        label = COVERAGE_LABELS.get(key, key.replace("_", " ").title())
        proposed = _num(cov.get("total_premium")) or _num(cov.get("premium"))
        exp = _expiring_for(key, expiring, expiring_details) if has_expiring else 0.0

        entry = {
            "_rank": _line_rank(key),
            "coverage": _email_label(key, label),
            "carrier": _short_carrier(_clean_carrier_name(cov.get("carrier", ""))) or "TBD",
            "key_limit": _key_limit(cov),
            "deductibles": _deductibles(cov),
            "expiring_carrier": _short_carrier(expiring_carriers.get(key) or ""),
            "proposed_premium": fmt_currency_cents(proposed) if proposed else "",
            "expiring_premium": fmt_currency_cents(exp) if exp else "",
            "dollar_change": "",
            "pct_change": "",
            "direction": "n/a",
        }
        if proposed and exp:
            delta = proposed - exp
            entry["dollar_change"] = _fmt_delta(delta)
            entry["pct_change"] = f"{(delta / exp * 100):+.1f}%"
            entry["direction"] = "increase" if delta > 0 else ("decrease" if delta < 0 else "flat")

        if cov.get("optional"):
            optional_lines.append(entry)
        else:
            lines.append(entry)
            total_proposed += proposed
            total_expiring += exp

    totals = {
        "proposed": fmt_currency_cents(total_proposed) if total_proposed else "",
        "expiring": fmt_currency_cents(total_expiring) if total_expiring else "",
        "dollar_change": "",
        "pct_change": "",
        "direction": "n/a",
    }
    if total_proposed and total_expiring:
        delta = total_proposed - total_expiring
        totals["dollar_change"] = _fmt_delta(delta)
        totals["pct_change"] = f"{(delta / total_expiring * 100):+.1f}%"
        totals["direction"] = "increase" if delta > 0 else ("decrease" if delta < 0 else "flat")

    lines.sort(key=lambda e: e["_rank"])
    optional_lines.sort(key=lambda e: e["_rank"])

    not_quoted = [c["coverage"] for c in optional_lines if not c["proposed_premium"]]
    optional_lines = [c for c in optional_lines if c["proposed_premium"]]

    _loc_count, _tiv = _location_summary(data)

    # The intro names the hotel, not the legal entity: "AC Hotel Orlando", not
    # "AC Orlando Hospitality LLC". DBA wins; otherwise strip the entity suffix.
    _hotel = (client.get("dba") or "").strip()
    if not _hotel:
        _hotel = re.sub(
            r"[,]?\s+\b(LLC|L\.L\.C\.|LLLP|LLP|LP|INC|INC\.|CORP|CORPORATION|CO|COMPANY|LTD)\b\.?$",
            "", (client.get("named_insured") or "").strip(), flags=re.I).strip()

    team = data.get("service_team") or {}
    return {
        "contact_first_name": (answers.get("contact_first_name") or "").strip(),
        "highlights": (answers.get("highlights") or "").strip(),
        "signoff": (answers.get("signoff") or "").strip(),
        "not_quoted": not_quoted,
        "hotel_name": _hotel or client.get("named_insured") or "the account",
        "named_insured": client.get("named_insured") or client.get("dba") or "the insured",
        "dba": client.get("dba") or "",
        "effective_date": client.get("effective_date") or "",
        "location_count": _loc_count,
        "total_tiv": fmt_currency_cents(_tiv).replace(".00", "") if _tiv else "",
        "has_expiring": has_expiring,
        "coverages": lines,
        "optional_coverages": optional_lines,
        "totals": totals,
        "service_team": team,
    }


_SYSTEM_PROMPT = """You write short renewal emails for Stefan Burkey, Hotel Franchise Practice Leader at HUB International, to the hotel owner or operator.

House style. Match this shape closely:

Hi Ajit,

Attached, is the finalized proposal for AC Hotel Orlando with a $87,053.27 decrease, or 30.7%, with $34,605,417 in total insured value.

Proposed carriers are Starr Surplus Lines Insurance Company for Property, Berkley Specialty Insurance Company for General Liability, StarStone National Insurance Company for Umbrella/Excess, Travelers for Equipment Breakdown, and Evanston Insurance Company for Terrorism/TRIA.

The key points are:
- Property includes a $10,000 property damage deductible, $10,000 time element deductible, $25,000 water damage deductible, $50,000 wind deductible, and 3% named windstorm subject to a $100,000 minimum per occurrence.
- Equipment Breakdown carries a $25,000 combined deductible with $4,000 property damage and 24 hours business income.
- Crime and Cyber are recommended but not included in the total premium; Cyber was quoted separately at $3,476.42. Supplemental applications may be needed to finalize quotes for the optional coverages.

Let me know if you would like to walk through the proposal next week or if you have any questions.

Enjoy your weekend.

Stefan

Structure, in this order:
1. Greeting by contact first name. If no name is given, skip the greeting line entirely.
2. One sentence: "Attached, is the finalized proposal for <hotel name>" plus the dollar and percent change and the total insured value. Use the HOTEL NAME you are given, never the legal entity. Mention the number of locations only when it is more than one.
3. One sentence listing proposed carriers: "Proposed carriers are <carrier> for <line>, <carrier> for <line>, and <carrier> for <line>." Keep the lines in EXACTLY the order the facts give them - they are already sorted Property, General Liability, Umbrella, Auto, Workers Compensation, EPLI, Cyber, then ancillary. Do not alphabetize and do not regroup.
4. "The key points are:" followed by short bullets. One bullet per line of coverage that actually has deductibles or terms worth stating, in the same line order. A final bullet listing the coverages that are recommended but not included in the total premium (name each one; include a separately quoted premium where one exists), closing with the sentence "Supplemental applications may be needed to finalize quotes for the optional coverages." Skip the whole section if there is nothing substantive.
5. Offer to walk through it.
6. Sign-off flavor if one was given, then "Stefan" on its own line.

Rules:
- Bullets are ONLY for the key points section. Everything else is flowing sentences.
- Never restate every coverage line with its own premium. The proposal document does that.
- Name the proposed carrier. If NO expiring carrier is given for a line, say nothing about moving, switching, or renewing carriers. Only when an expiring carrier is actually given may you say the account is moving from that carrier to the proposed one, or that the incumbent held or improved the program if they match.
- Use the carrier names exactly as given. They are already shortened for a client email; never expand them back to full legal names.
- Never invent numbers, carriers, limits, deductibles, or dates. Use only what you are given. Omit rather than guess.
- No em dashes. No emoji. No markdown bold or headers. Plain text for Outlook.
- Warm and direct. No "I hope this email finds you well." No corporate padding. Keep the whole email under about 200 words.

Return exactly two parts:
SUBJECT: <one line>
BODY:
<the email>"""


def _context_to_prompt(ctx):
    out = []
    if ctx.get("contact_first_name"):
        out.append(f"Contact first name: {ctx['contact_first_name']}")
    out.append(f"HOTEL NAME to use in the intro: {ctx.get('hotel_name')}")
    out.append(f"Legal named insured (do NOT use in the intro): {ctx['named_insured']}")
    if ctx["effective_date"]:
        out.append(f"Effective date: {ctx['effective_date']}")
    size = []
    n = ctx.get("location_count") or 0
    if n > 1:
        size.append(f"{n} locations")
    if ctx.get("total_tiv"):
        size.append(f"total insured value {ctx['total_tiv']}")
    if size:
        out.append("STATE THIS IN THE INTRO: " + ", ".join(size))

    t = ctx["totals"]
    if ctx["has_expiring"] and t["dollar_change"]:
        out.append(
            f"HEADLINE: total premium {t['proposed']} versus expiring {t['expiring']}, "
            f"a {t['direction']} of {t['dollar_change']} ({t['pct_change']})."
        )
    elif t["proposed"]:
        out.append(f"Total proposed premium {t['proposed']}. No expiring premium on file, so do not reference any change.")
    else:
        out.append("No premium totals available. Do not state any premium figures.")

    out.append("")
    out.append("Coverages, ALREADY IN THE ORDER TO PRESENT THEM:")
    for c in ctx["coverages"]:
        bits = [f"- {c['coverage']}: proposed carrier {c['carrier']}"]
        if c.get("expiring_carrier"):
            same = c["expiring_carrier"].lower()[:6] in (c["carrier"] or "").lower()
            bits.append(
                f"expiring carrier {c['expiring_carrier']}"
                + (" (same carrier, renewing)" if same else " (moving carriers)")
            )
        if c["key_limit"]:
            bits.append(f"limit {c['key_limit']}")
        if c["dollar_change"]:
            bits.append(f"change {c['dollar_change']} ({c['pct_change']})")
        if c.get("deductibles"):
            bits.append("deductibles " + "; ".join(c["deductibles"]))
        out.append(", ".join(bits))

    rec = [f"{c['coverage']} (quoted separately at {c['proposed_premium']})"
           for c in ctx["optional_coverages"]]
    rec += [f"{name} (not yet quoted)" for name in ctx.get("not_quoted") or []]
    if rec:
        out.append("")
        out.append("RECOMMENDED OPTIONAL COVERAGES, not included in the total premium: " + ", ".join(rec)
                   + ". The email's final key-points bullet must list these as recommended but not included in the "
                     'total premium and close with: "Supplemental applications may be needed to finalize quotes '
                     'for the optional coverages."')

    if ctx.get("highlights"):
        out.append("")
        out.append("Highlights Stefan wants worked in (use these, they matter): " + ctx["highlights"])

    if ctx.get("signoff"):
        out.append("")
        out.append(f"Sign-off flavor to use before his name: {ctx['signoff']}")

    return "\n".join(out)


def draft_email(ctx, instruction=None, previous=None, model=None):
    """Draft (or refine) the client email. Raises on API failure."""
    from openai import OpenAI

    if not os.environ.get("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY is not set in the environment")

    messages = [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {"role": "user", "content": "Proposal facts:\n\n" + _context_to_prompt(ctx)},
    ]
    if previous:
        messages.append({"role": "assistant", "content": previous})
    if instruction:
        messages.append({
            "role": "user",
            "content": (
                "Revise the email above per this instruction. Keep every factual figure "
                "unchanged unless the instruction explicitly says otherwise. "
                "Instruction: " + instruction
            ),
        })

    client = OpenAI()
    resp = client.chat.completions.create(
        model=model or os.environ.get("EMAIL_MODEL", "gpt-5.6-terra"),
        messages=messages,
        max_completion_tokens=2000,
    )
    text = (resp.choices[0].message.content or "").strip()
    return _split_subject_body(text)


def _split_subject_body(text):
    subject, body = "", text
    if "SUBJECT:" in text:
        after = text.split("SUBJECT:", 1)[1]
        if "BODY:" in after:
            subject, body = after.split("BODY:", 1)
        else:
            parts = after.split("\n", 1)
            subject = parts[0]
            body = parts[1] if len(parts) > 1 else ""
    return {"subject": subject.strip(), "body": body.strip()}
