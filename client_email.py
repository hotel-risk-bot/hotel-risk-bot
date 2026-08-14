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
            "coverage": label,
            "carrier": _clean_carrier_name(cov.get("carrier", "")) or "TBD",
            "key_limit": _key_limit(cov),
            "deductibles": _deductibles(cov),
            "expiring_carrier": (expiring_carriers.get(key) or "").strip(),
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

    not_quoted = [c["coverage"] for c in optional_lines if not c["proposed_premium"]]
    optional_lines = [c for c in optional_lines if c["proposed_premium"]]

    team = data.get("service_team") or {}
    return {
        "contact_first_name": (answers.get("contact_first_name") or "").strip(),
        "highlights": (answers.get("highlights") or "").strip(),
        "signoff": (answers.get("signoff") or "").strip(),
        "not_quoted": not_quoted,
        "named_insured": client.get("named_insured") or client.get("dba") or "the insured",
        "dba": client.get("dba") or "",
        "effective_date": client.get("effective_date") or "",
        "location_count": len(data.get("locations") or []),
        "has_expiring": has_expiring,
        "coverages": lines,
        "optional_coverages": optional_lines,
        "totals": totals,
        "service_team": team,
    }


_SYSTEM_PROMPT = """You write short renewal emails for Stefan Burkey, Hotel Franchise Practice Leader at HUB International, to the hotel owner or operator.

House style, follow it closely:

Hi Debbie,

Attached, happy to provide the renewal proposal for TownePlace Suites delivering a $98,507.12 decrease, or 49.2%, from expiring. Our program carrier Starr is quoting a much better premium and AOP remains $10,000, 5% named windstorm, $25,000 all other wind, and $25,000 water damage. Flood and earthquake are still excluded. Let me know if you would like to walk through the proposal next week or if you have any questions.

Enjoy your weekend.

Stefan

Rules:
- Two short paragraphs at most, then the sign-off. This is a note, not a summary document.
- Greet by the contact first name if one is given. If none is given, open with "Attached" and no greeting line.
- Lead with the headline: total change in dollars and percent, stated in the first sentence.
- Name the carrier. If an expiring carrier is given and it differs from the proposed one, say the account is moving from the expiring carrier to the proposed one. If it is the same carrier, say the incumbent held or improved the program.
- Work the deductible structure into prose, not a list: AOP, named windstorm, wind/hail, water damage. Only ones you were given.
- If coverages were presented but not quoted, say those perils are excluded or not included, in one short clause.
- NEVER use bullet points, dashes as list markers, tables, headers, or markdown. Flowing sentences only.
- Do not restate every coverage line with its own premium. The proposal document does that.
- Never invent numbers, carriers, limits, deductibles, or dates. Use only what you are given. Omit rather than guess.
- No em dashes. No emoji. Plain text for Outlook.
- Close with an offer to walk through it, then the sign-off flavor if one was given, then "Stefan" on its own line. No title block.
- Warm and direct. No "I hope this email finds you well." No corporate padding.

Return exactly two parts:
SUBJECT: <one line>
BODY:
<the email>"""


def _context_to_prompt(ctx):
    out = []
    if ctx.get("contact_first_name"):
        out.append(f"Contact first name: {ctx['contact_first_name']}")
    out.append(f"Named insured: {ctx['named_insured']}")
    if ctx["dba"]:
        out.append(f"Property / DBA: {ctx['dba']}")
    if ctx["effective_date"]:
        out.append(f"Effective date: {ctx['effective_date']}")

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
    out.append("Coverages:")
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

    if ctx.get("not_quoted"):
        out.append("")
        out.append("Presented but NOT quoted, so not included in the program: " + ", ".join(ctx["not_quoted"]))

    if ctx["optional_coverages"]:
        out.append("")
        out.append("Optional coverages quoted for consideration: " + ", ".join(
            f"{c['coverage']} {c['proposed_premium']}".strip() for c in ctx["optional_coverages"]))

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
        model=model or os.environ.get("EMAIL_MODEL", "gpt-5.4-mini"),
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
