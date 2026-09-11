"""
HotelBound (RT Specialty / Lloyd's) property quote parser.

A HotelBound quote is a fixed-format document: an RT "Insurance Proposal" cover (Cost Summary,
Subjectivities, Remarks), a Lloyd's "Quotation Memorandum" (Risk Details, Deductibles A-H,
Schedule of Sublimits, Warranties, Subjectivities, Premium and Fees, Terms/Conditions), the
Schedule of Values table (LID / street / construction / values), the Cost Break Out By Location
table, and the Shared Capacity and Dedicated Limits Disclosure.

Because the layout is stable, everything the proposal needs is parsed deterministically from the
`pdftotext -layout` text — no GPT involved — and the result supersedes whatever GPT or an
uploaded SOV produced for the PROPERTY section of the proposal (Stefan, Sep 2026):

  1. premium = non-admitted premium; total = RT "Total Policy Cost" (all fees + taxes)
  2. subjectivities (RT cover + memo "Due at binding" list, de-duplicated)
  3. terrorism included in premium
  4. deductibles A-H reduced to the clauses that actually apply to the scheduled locations
     (Tier One / Tier Two / Hail Belt lookups by state + county)
  5. fully earned fees + minimum earned premium disclosure (incl. the RT program brokerage fee)
  6. per-location / per-building schedule (flood zone, year built, ISO construction, units,
     sprinklered, sq ft, Building / Contents / BI / TIV)
  plus the Shared Capacity and Dedicated Limits figures for the disclosure page.

Public entry points:
    parse_hotelbound_quote(text) -> dict          (hb["detected"] False when not a HotelBound quote)
    applicable_deductibles(hb) -> (rows, notes)
    apply_hotelbound_to_data(data, hb) -> None    (mutates the extracted data dict in place)
"""

import logging
import re

from hotelbound_territories import TIER1_TX_VA, TIER1_MD_ME, TIER2, HAIL_BELT_FALLBACK

logger = logging.getLogger(__name__)

ISO_CONSTRUCTION = {
    1: "Frame",
    2: "Joisted Masonry",
    3: "Non-Combustible",
    4: "Masonry Non-Combustible",
    5: "Modified Fire Resistive",
    6: "Fire Resistive",
}

ATC_OCCUPANCY = {
    2: "Apartments",
    3: "Hotel or Motel",
    5: "Retail",
    8: "Office",
    37: "General Commercial",
    42: "Condominium",
}

STATE_NAMES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "DC": "District of Columbia",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois",
    "IN": "Indiana", "IA": "Iowa", "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana",
    "ME": "Maine", "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota",
    "MS": "Mississippi", "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon",
    "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota",
    "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont", "VA": "Virginia",
    "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
}
_NAME_TO_ABBR = {v.lower(): k for k, v in STATE_NAMES.items()}

# Deductible clause C: "Louisiana through Virginia on the Eastern Seaboard, excluding Florida" —
# the Tier One states in that stretch (Texas and Florida are clause B; Hawaii is clause E).
_GULF_ATLANTIC_TIER1_STATES = {"LA", "MS", "AL", "GA", "SC", "NC", "VA"}

_FOOTER_RE = re.compile(
    r"RT Specialty is a division|subsidiary of Ryan Specialty|In California:\s*RSG|"
    r"150 South US Highway|Phone - \(561\)|^\s*UQT \d|^\s*QUOTATION MEMORANDUM\s*$|Page \|\s*\d+|"
    r"^\s*B \d\.\d\.\d\s+APP\d+|^\s*\d+ of \d+\s*$"
)


# ---------------------------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------------------------
def _num(s):
    try:
        return float(str(s).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError, AttributeError):
        return 0.0


def _fmt_money(v, cents=False):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return f"${v:,.2f}" if cents else f"${v:,.0f}"


def _clean_lines(block):
    """Drop RT footer/header noise lines from a text block, keep everything else."""
    return [l for l in block.split("\n") if not _FOOTER_RE.search(l)]


def _norm_county(c):
    c = (c or "").lower().strip()
    c = re.sub(r"[’']", "", c)
    c = re.sub(r"\b(county|parish|borough)\b", "", c)
    c = re.sub(r"\bst\.?\s", "st ", c)
    c = re.sub(r"[^a-z0-9 ]", " ", c)
    return re.sub(r"\s+", " ", c).strip()


def _county_in(state, county, table):
    entry = table.get((state or "").upper())
    if not entry:
        return False
    if entry == "ALL":
        return True
    target = _norm_county(county)
    if not target:
        return False
    # Exact match on the normalized name. "Baltimore City" and "Baltimore" (county) are listed
    # separately in the tables, so no suffix tolerance — the quote's County column uses the
    # same spelling as the program definitions.
    return any(_norm_county(c) == target for c in entry)


def _bullets(lines, bullet_chars="•·-"):
    """Join wrapped bullet items: a line starting with a bullet char starts a new item;
    following indented lines without a bullet continue it."""
    items = []
    for raw in lines:
        s = raw.strip()
        # pdftotext emits stray single-letter lines ("S") between RT cover bullets
        if not s or (len(s) == 1 and s.isalpha()):
            continue
        if s[0] in bullet_chars:
            items.append(s.lstrip(bullet_chars).strip())
        elif items:
            items[-1] = (items[-1] + " " + s).strip()
    return [re.sub(r"\s+", " ", i) for i in items if i]


# ---------------------------------------------------------------------------------------------
# main parser
# ---------------------------------------------------------------------------------------------
def parse_hotelbound_quote(text: str) -> dict:
    hb = {
        "detected": False,
        # premium / fees
        "cost_summary": [],           # [{"label","amount"}] RT cover Cost Summary lines (excl. total)
        "total_policy_cost": 0.0,
        "base_premium": 0.0,
        "memo_fees": [],              # [{"label","amount"}] Quotation Memorandum PREMIUM AND FEES
        "memo_total": 0.0,
        "program_brokerage_fee": 0.0,
        "association_fee": 0.0,
        "inspection_fee": 0.0,
        "commission": "",
        # narrative pieces
        "subjectivities": [],
        "warranties": [],
        "terrorism": "",
        "terrorism_included": False,
        "valuation": [],
        "payment_terms": "",
        "policy_period": "",
        "lead_insurer": "",
        "insured": "",
        "home_state": "",
        "coverage_type_text": "",
        "tiv": 0.0,
        "program_limit": 0.0,
        "monthly_limit_of_indemnity": "",
        "territorial_limits": "",
        # deductibles
        "deductible_clauses": [],     # parsed A-H
        "waiting_period_hours": "",
        "hail_belt": {},
        # earned premium
        "mep_text": "",
        "mep_amount": 0.0,
        "mep_percent": "",
        "mep_tier_condition": "",
        "fully_earned_fees": [],
        "program_brokerage_fee_text": "",
        # schedule
        "schedule": [],               # per LID rows (SOV table merged with Cost Break Out)
        "account_name": "",
        "control_number": "",
        # shared limits disclosure
        "shared_limits": {},
    }
    if not text:
        return hb
    up = text.upper()
    markers = ("HOTELBOUND", "RT SPECIALTY", "RYAN TURNER SPECIALTY", "RYAN SPECIALTY")
    if not (any(m in up for m in markers) and ("QUOTATION MEMORANDUM" in up or "COST SUMMARY" in up)):
        return hb
    if "HOTELBOUND" not in up:
        # RT Specialty but not the HotelBound program — leave to the generic path
        return hb
    hb["detected"] = True

    lines = text.split("\n")

    # ---- RT cover: Cost Summary ------------------------------------------------------------
    m = re.search(r"Cost Summary\s*\n(.*?)Total Policy Cost\s+\$?\s*([\d,]+(?:\.\d{2})?)", text, re.S | re.I)
    if m:
        for l in m.group(1).split("\n"):
            mm = re.match(r"^\s*([A-Za-z][A-Za-z /&()\-]+?)\s{2,}\$?\s*([\d,]+(?:\.\d{2})?)\s*$", l)
            if mm:
                hb["cost_summary"].append({"label": mm.group(1).strip(), "amount": _num(mm.group(2))})
        hb["total_policy_cost"] = _num(m.group(2))
    for item in hb["cost_summary"]:
        if re.search(r"premium", item["label"], re.I) and not hb["base_premium"]:
            hb["base_premium"] = item["amount"]

    # ---- Quotation Memorandum: PREMIUM AND FEES ----------------------------------------------
    m = re.search(r"PREMIUM AND FEES:?(.*?)(?:\n\s*\n\s*\n|Commission:)", text, re.S)
    if m:
        for l in m.group(1).split("\n"):
            mm = re.match(r"^\s*\*?\s*([A-Za-z][A-Za-z \-]+?):\s*\$?\s*([\d,]+(?:\.\d{2})?)\s*$", l)
            if mm:
                label, amt = mm.group(1).strip(), _num(mm.group(2))
                if label.lower() == "total":
                    hb["memo_total"] = amt
                else:
                    hb["memo_fees"].append({"label": label, "amount": amt})
                    ll = label.lower()
                    if "brokerage" in ll:
                        hb["program_brokerage_fee"] = amt
                    elif "association" in ll or "membership" in ll:
                        hb["association_fee"] = amt
                    elif "inspection" in ll:
                        hb["inspection_fee"] = amt
                    elif "premium" in ll and not hb["base_premium"]:
                        hb["base_premium"] = amt
    m = re.search(r"Commission:\s*([\d.]+%)", text)
    if m:
        hb["commission"] = m.group(1)

    # ---- Subjectivities (RT cover bullets + memo "Due at binding" list) ------------------------
    subj = []
    m = re.search(r"\nSubjectivities\s*\n(.*?)The Subjectivities outlined above", text, re.S)
    if m:
        subj += _bullets(_clean_lines(m.group(1)))
    m = re.search(r"SUBJECTIVITIES:\s*(?:Due at binding:)?\s*\n(.*?)\n\s*PREMIUM AND FEES", text, re.S)
    if m:
        subj += _bullets(_clean_lines(m.group(1)))
    seen, dedup = set(), []
    for s in subj:
        k = re.sub(r"[^a-z0-9]", "", s.lower())
        if k and k not in seen:
            seen.add(k)
            dedup.append(s)
    hb["subjectivities"] = dedup

    # ---- Warranties -------------------------------------------------------------------------
    m = re.search(r"WARRANTIES:\s*(.*?)\n\s*SUBJECTIVITIES:", text, re.S)
    if m:
        for l in _clean_lines(m.group(1)):
            s = re.sub(r"\s+", " ", l.strip())
            if s.lower().startswith("warrant"):
                hb["warranties"].append(s)
            elif s and hb["warranties"]:
                hb["warranties"][-1] += " " + s

    # ---- Terrorism / valuation / payment / period / insurer ----------------------------------
    m = re.search(r"TERRORISM:\s*(.+)", text)
    if m:
        hb["terrorism"] = re.sub(r"\s+", " ", m.group(1)).strip()
        hb["terrorism_included"] = hb["terrorism"].lower().startswith("included")
    m = re.search(r"VALUATION:\s*(.+?)\n\s*\n", text, re.S)
    if m:
        hb["valuation"] = [re.sub(r"\s+", " ", l).strip() for l in m.group(1).split("\n") if l.strip()]
    m = re.search(r"PAYMENT TERMS:\s*(.+)", text)
    if m:
        # "Payment is due in our office 15 days from inception" — "our office" is RT's, not the insured's
        hb["payment_terms"] = re.sub(r"\s+in our office", "", re.sub(r"\s+", " ", m.group(1))).strip()
    m = re.search(r"PERIOD:\s*From\s+(.+?)\s+to\s+(.+?)\s+both days", text, re.S)
    if m:
        hb["policy_period"] = f"{m.group(1).strip()} to {m.group(2).strip()}"
    m = re.search(r"LEAD INSURER:\s*(.+)", text)
    if m:
        hb["lead_insurer"] = re.sub(r"\s+", " ", m.group(1)).strip()
    m = re.search(r"\bTYPE:\s*(.+?)\n\s*\n", text, re.S)
    if m:
        hb["coverage_type_text"] = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(".")
    m = re.search(r"Insured Name:\s*(.+)", text)
    if m:
        hb["insured"] = m.group(1).strip()
    m = re.search(r"Home State:\s*([A-Z]{2})", text)
    if m:
        hb["home_state"] = m.group(1)
    m = re.search(r"\bTIV:\s*\$?\s*([\d,]+)", text)
    if m:
        hb["tiv"] = _num(m.group(1))
    m = re.search(r"\$\s*([\d,]+)\s+Overall Program Limit of Liability", text)
    if m:
        hb["program_limit"] = _num(m.group(1))
    m = re.search(r"Monthly Limit of Indemnity fraction:\s*([0-9]+/[0-9]+)", text)
    if m:
        hb["monthly_limit_of_indemnity"] = m.group(1)
    m = re.search(r"TERRITORIAL LIMITS:\s*(.+?)\n\s*\n", text, re.S)
    if m:
        hb["territorial_limits"] = re.sub(r"\s+", " ", m.group(1)).strip()
    m = re.search(r"ACCOUNT NAME:\s*(.+?)\s{2,}", text)
    if m:
        hb["account_name"] = m.group(1).strip()
    m = re.search(r"CONTROL\s*#:\s*\n?\s*(\d+)", text)
    if m:
        hb["control_number"] = m.group(1)

    # ---- Minimum earned premium / fully earned fees -----------------------------------------
    m = re.search(r"(The Minimum Premium is .*?)(?:\n\s*\n|·\s*For properties)", text, re.S)
    if m:
        hb["mep_text"] = re.sub(r"\s+", " ", " ".join(_clean_lines(m.group(1)))).strip()
        mm = re.search(r"calculated to be \$?\s*([\d,]+)", hb["mep_text"])
        if mm:
            hb["mep_amount"] = _num(mm.group(1))
        mm = re.search(r"Minimum Premium is (\d+%)", hb["mep_text"])
        if mm:
            hb["mep_percent"] = mm.group(1)
    m = re.search(r"For properties located in .Tier One. or .Tier Two. Counties(.*?)(?:\n\s*\d+\.\s|\n\s*\n)", text, re.S)
    if m:
        hb["mep_tier_condition"] = re.sub(r"\s+", " ", "For properties located in Tier One or Tier Two Counties" + m.group(1)).strip()
    m = re.search(r"Fully Earned Fees:\s*\n(.*?)\n\s*\d+\.\s", text, re.S)
    if m:
        hb["fully_earned_fees"] = _bullets(_clean_lines(m.group(1)))
    m = re.search(r"\d+\.\s+(RT Specialty has charged a program brokerage fee.*?fully earned at (?:inception|binding)\.)", text, re.S)
    if m:
        hb["program_brokerage_fee_text"] = re.sub(r"\s+", " ", m.group(1)).strip()

    # ---- Schedule of Sublimits A..GG ---------------------------------------------------------
    hb["sublimits"] = _parse_sublimits(text)
    m = re.search(r"ATTACHMENTS:[ \t]*(.*?)\n[ \t]*\n", text, re.S)
    if m:
        hb["attachments"] = [re.sub(r"\s+", " ", l).strip() for l in m.group(1).split("\n") if l.strip()]
    else:
        hb["attachments"] = []
    m = re.search(r"These limitations may include, but are not limited to, the following:\s*\n(.*?)\n\s*\n", text, re.S)
    if m:
        hb["limitations"] = _bullets(_clean_lines(m.group(1)))
    else:
        hb["limitations"] = []

    # ---- Hail Belt (live definition) --------------------------------------------------------
    hb["hail_belt"] = _parse_hail_belt(text) or dict(HAIL_BELT_FALLBACK)

    # ---- Deductible clauses A-H -------------------------------------------------------------
    hb["deductible_clauses"] = _parse_deductible_clauses(text)
    for cl in hb["deductible_clauses"]:
        if cl["kind"] == "waiting_period":
            hb["waiting_period_hours"] = cl.get("hours", "")

    # ---- Schedule of Values + Cost Break Out --------------------------------------------------
    hb["schedule"] = _parse_schedule(text)
    if not hb["tiv"] and hb["schedule"]:
        hb["tiv"] = sum(r["tiv"] for r in hb["schedule"])

    # ---- Shared Capacity / Aggregate Layer / Margin Clause / Equipment Breakdown --------------
    sl = {}
    m = re.search(r"program limit of insurance of\s*\$\s*([\d,]+)\s+per Occurrence", text, re.I)
    if m:
        sl["program_limit"] = _num(m.group(1))
    m = re.search(r"dedicated limits equal to each\s+EOCs?\s+Limit of Liability,\s+not to exceed\s*\$\s*([\d,]+)", text, re.I)
    if m:
        sl["dedicated_cap"] = _num(m.group(1))
    m = re.search(r"Account is used to pay Members up to\s*\$\s*([\d,]+)\s+per\s+Occurrence", text, re.I)
    if m:
        sl["aggregate_layer_per_occurrence"] = _num(m.group(1))
    m = re.search(r"Aggregate Layer is insured by\s+(.+?)\s+and funded", text, re.I | re.S)
    if m:
        sl["aggregate_layer_insurer"] = re.sub(r"\s+", " ", m.group(1)).strip()
    m = re.search(r"third-party administrator \(.TPA.\), which at\s+the time of the issuance of this Endorsement is\s+([A-Za-z&. ]+?)\.", text, re.S)
    if m:
        sl["aggregate_layer_tpa"] = m.group(1).strip()
    m = re.search(r"Margin Clause percentage:\s*(\d+%)", text)
    if m:
        sl["margin_clause"] = m.group(1)
    m = re.search(r"added to the master\s+([A-Z][A-Za-z&. ]{2,60}?)\s+Equipment\s+Breakdown\s+Policy\s+([A-Z0-9]+)", text)
    if m:
        sl["eb_carrier"] = re.sub(r"\s+", " ", m.group(1)).strip()
        sl["eb_policy"] = m.group(2).strip()
    m = re.search(r"total limit payable in any one occurrence is\s*\$\s*([\d,]+)", text, re.I)
    if m:
        sl["eb_limit"] = _num(m.group(1))
    m = re.search(r"MINIMUM AM BEST FINANCIAL RATING OF\s+([A-Z][-+]?\s*\d*)", text)
    if m:
        sl["min_am_best"] = m.group(1).strip()
    hb["shared_limits"] = sl

    logger.info(
        f"HotelBound quote parsed: premium={hb['base_premium']}, total={hb['total_policy_cost']}, "
        f"locations={len(hb['schedule'])}, clauses={len(hb['deductible_clauses'])}, "
        f"subjectivities={len(hb['subjectivities'])}, mep={hb['mep_amount']}"
    )
    return hb


def _parse_sublimits(text):
    """SCHEDULE OF SUBLIMITS: 'A. $500,000 per Occurrence per Insured Member as respects Extra Expense;'
    -> [{"letter","description","limit"}]. Items wrap across lines and page footers."""
    m = re.search(r"SCHEDULE OF\s+(?:SUBLIMITS:)?\s*(.*?)The above may not include all sublimits", text, re.S)
    if not m:
        return []
    body = "\n".join(_clean_lines(m.group(1)))
    body = re.sub(r"\n\s*SUBLIMITS:\s*", "\n", body)
    pieces = re.split(r"\n\s*([A-Z]{1,2})\.\s+", "\n" + body)
    out = []
    for i in range(1, len(pieces) - 1, 2):
        letter = pieces[i]
        txt = re.sub(r"\s+", " ", pieces[i + 1]).strip().rstrip(";.").strip()
        if not txt:
            continue
        mm = re.search(r"^(.*?)\s+as respects\s+(.+)$", txt, re.I)
        if mm:
            limit, desc = mm.group(1).strip(), mm.group(2).strip()
            # trailing qualifiers ("; limit 5 mile radius") stay with the description
            desc = re.sub(r"\s*;\s*", " — ", desc)
        else:
            mm = re.match(r"^(.*?)\s+(Included\.?)$", txt)
            mn = re.match(r"^In respect of (.+?) to be reported within (\d+ days).*?not to exceed a \$([\d,]+) limit per location"
                          r"(?: and not to exceed \$([\d,]+) in cumulative new Total Insured Values)?", txt, re.I)
            if mm:
                desc, limit = mm.group(1).strip(), "Included"
            elif mn:
                desc = f"{mn.group(1)} — report within {mn.group(2)} of acquisition"
                limit = f"${mn.group(3)} per location"
                if mn.group(4):
                    limit += f"; ${mn.group(4)} cumulative new TIV"
            else:
                desc, limit = txt, ""
        # "per Occurrence per Insured Member" is implied program-wide — shorten the limit column
        limit = re.sub(r"\s*per Occurrence per Insured Member", " per Occurrence", limit)
        limit = re.sub(r"\s*per Insured Member", "", limit)
        out.append({"letter": letter, "description": desc, "limit": limit})
    return out


def _parse_hail_belt(text):
    m = re.search(
        r"Hail Belt Counties means all locations situated within the counties specified as below:(.*?)(?:\n\s*\n\s*\n|\Z)",
        text, re.S)
    if not m:
        return {}
    flat = " ".join(" ".join(_clean_lines(m.group(1))).split())
    names = sorted(STATE_NAMES.values(), key=len, reverse=True)
    parts = re.split(r"\b(" + "|".join(re.escape(n) for n in names) + r")\s*-\s*", flat)
    out = {}
    for i in range(1, len(parts) - 1, 2):
        abbr = _NAME_TO_ABBR[parts[i].lower()]
        counties = [c.strip() for c in parts[i + 1].strip().rstrip(",").split(",") if c.strip()]
        if any("all counties" in c.lower() or "entire state" in c.lower() for c in counties):
            out[abbr] = "ALL"
        else:
            out[abbr] = counties
    return out


def _parse_deductible_clauses(text):
    m = re.search(r"DEDUCTIBLES:(.*?)(?:For the purposes of applying percentage deductibles|SCHEDULE OF\s+SUBLIMITS)", text, re.S)
    if not m:
        return []
    body = "\n".join(_clean_lines(m.group(1)))
    # split on "A. " ... "H. " at line start
    pieces = re.split(r"\n\s*([A-Z])\.\s+", "\n" + body)
    clauses = []
    for i in range(1, len(pieces) - 1, 2):
        letter = pieces[i]
        txt = re.sub(r"\s+", " ", pieces[i + 1]).strip()
        cl = {"letter": letter, "text": txt, "kind": "other", "percent": "", "minimum": 0.0,
              "amount": 0.0, "zone": "", "hours": ""}
        low = txt.lower()
        mp = re.match(r"(\d+(?:\.\d+)?)%", txt)
        mmin = re.search(r"minimum deductible of \$\s*([\d,]+)", txt, re.I)
        mamt = re.match(r"\$\s*([\d,]+)", txt)
        if mp:
            cl["percent"] = mp.group(1) + "%"
        if mmin:
            cl["minimum"] = _num(mmin.group(1))
        if mamt:
            cl["amount"] = _num(mamt.group(1))
        if "waiting period" in low:
            cl["kind"] = "waiting_period"
            mh = re.search(r"exceeds\s+(\d+)\s+consecutive hours", txt, re.I)
            if mh:
                cl["hours"] = mh.group(1)
        elif "water damage" in low:
            cl["kind"] = "water_damage"
        elif "hail belt" in low:
            cl["kind"] = "hail_belt"
            cl["zone"] = "HAIL"
        elif "named windstorm" in low and mp:
            cl["kind"] = "named_windstorm"
            if "hawaii" in low:
                cl["zone"] = "HI"
            elif "maryland through maine" in low:
                cl["zone"] = "MD_ME"
            elif "tier two" in low:
                cl["zone"] = "GULF_ATL"
            elif "florida" in low:
                cl["zone"] = "FL_TX1"
        elif "other than as described below" in low or ("direct physical loss" in low and mamt and not mp):
            cl["kind"] = "aop"
        clauses.append(cl)
    return clauses


def _parse_schedule(text):
    """Merge the Schedule of Values table (attributes) with Cost Break Out (clean address + $)."""
    rows = {}
    order = []

    # --- SOV table: anchor on LID at line start; continuation lines (wrapped street) are
    # short indented lines with no digits that follow a LID row.
    m = re.search(r"\n\s*LID\s+Street\s+City\s+State(.*?)(?:For Reference in the above table|\Z)", text, re.S)
    sov_block = m.group(1) if m else ""
    sov_re = re.compile(
        r"^\s*(?P<lid>\d{9,15})\s+(?P<pre>.+?)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})\s+"
        r"(?P<county>[A-Za-z .'\-]+?)\s{2,}(?P<flood>Shaded X|[A-Z0-9]{1,6})\s+(?P<yr>\d{4})\s+"
        r"(?P<ints>(?:\d+\s+){3,5})(?P<sprk>[YNP])\s+(?P<sqft>[\d,]+)\s+"
        r"\$?(?P<bldg>[\d,]+)\s+\$?(?P<bpp>[\d,]+)\s+\$?(?P<bi>[\d,]+)\s+\$?(?P<tiv>[\d,]+)\s*$"
    )
    last_lid = None
    for line in sov_block.split("\n"):
        mm = sov_re.match(line)
        if mm:
            ints = [int(x) for x in mm.group("ints").split()]
            atc, nbldg, stories, units, iso = None, None, None, None, None
            if len(ints) >= 5:
                atc, nbldg, stories, units, iso = ints[0], ints[1], ints[2], ints[3], ints[4]
            elif len(ints) == 4:
                atc, nbldg, units, iso = ints
            elif len(ints) == 3:
                atc, units, iso = ints
            pre = re.split(r"\s{2,}", mm.group("pre").strip())
            street = pre[0].strip() if pre else ""
            city = pre[-1].strip() if len(pre) > 1 else ""
            r = {
                "lid": mm.group("lid"), "street": street, "city": city,
                "state": mm.group("state"), "zip": mm.group("zip"), "county": mm.group("county").strip(),
                "flood_zone": mm.group("flood"), "year_built": mm.group("yr"),
                "atc_code": atc, "occupancy": ATC_OCCUPANCY.get(atc, ""),
                "num_buildings": nbldg, "num_stories": stories, "num_units": units,
                "iso_code": iso, "construction": ISO_CONSTRUCTION.get(iso, ""),
                "sprinklered": mm.group("sprk"), "sqft": _num(mm.group("sqft")),
                "bldg_value": _num(mm.group("bldg")), "bpp_value": _num(mm.group("bpp")),
                "bi_value": _num(mm.group("bi")), "tiv": _num(mm.group("tiv")),
                "premium": 0.0, "fee": 0.0, "agg_premium": 0.0,
            }
            rows[r["lid"]] = r
            order.append(r["lid"])
            last_lid = r["lid"]
        elif last_lid and line.strip() and not re.search(r"\d", line) and len(line.strip()) < 25 \
                and not line.strip().lower().startswith(("for reference", "atc", "iso")):
            rows[last_lid]["street"] = (rows[last_lid]["street"] + " " + line.strip()).strip()
        elif not line.strip():
            last_lid = None

    # --- Cost Break Out By Location ---
    m = re.search(r"Cost\s+Break\s+Out\s+By\s+Location(.*?)(?:SHARED CAPACITY|\Z)", text, re.S | re.I)
    if m:
        cb_re = re.compile(
            r"^\s*(?P<lid>\d{9,15})\s{2,}(?P<street>.+?)\s{2,}(?P<city>.+?)\s{2,}(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})\s{2,}"
            r"(?P<county>[A-Za-z .'\-]+?)\s{2,}\$?(?P<agg>[\d,]+(?:\.\d{2})?)\s{2,}\$?(?P<prem>[\d,]+(?:\.\d{2})?)\s{2,}"
            r"\$?(?P<fee>[\d,]+(?:\.\d{2})?)\s{2,}\$?(?P<bldg>[\d,]+)\s{2,}\$?(?P<bpp>[\d,]+)\s{2,}\$?(?P<bi>[\d,]+)\s{2,}\$?(?P<tiv>[\d,]+)"
        )
        for line in m.group(1).split("\n"):
            mm = cb_re.match(line)
            if not mm:
                continue
            lid = mm.group("lid")
            r = rows.get(lid)
            if r is None:
                r = {"lid": lid, "flood_zone": "", "year_built": "", "atc_code": None, "occupancy": "",
                     "num_buildings": None, "num_stories": None, "num_units": None, "iso_code": None,
                     "construction": "", "sprinklered": "", "sqft": 0.0}
                rows[lid] = r
                order.append(lid)
            r.update({
                "street": mm.group("street").strip(), "city": mm.group("city").strip(),
                "state": mm.group("state"), "zip": mm.group("zip"), "county": mm.group("county").strip(),
                "agg_premium": _num(mm.group("agg")), "premium": _num(mm.group("prem")), "fee": _num(mm.group("fee")),
                "bldg_value": _num(mm.group("bldg")), "bpp_value": _num(mm.group("bpp")),
                "bi_value": _num(mm.group("bi")), "tiv": _num(mm.group("tiv")),
            })

    out = [rows[l] for l in order]
    # building numbers: same street+city -> 1-1, 1-2 ...
    groups = {}
    for r in out:
        key = (re.sub(r"\W", "", r["street"].lower()), re.sub(r"\W", "", r["city"].lower()), r["state"])
        groups.setdefault(key, []).append(r)
    loc_no = 0
    for key, grp in groups.items():
        loc_no += 1
        for i, r in enumerate(grp, 1):
            r["location_number"] = loc_no
            r["building_number"] = i
            r["label"] = f"{loc_no}-{i}" if len(grp) > 1 else str(loc_no)
    return out


# ---------------------------------------------------------------------------------------------
# deductible applicability
# ---------------------------------------------------------------------------------------------
def location_zones(state, county, hail_belt):
    """Return the set of territory tags a location falls in."""
    st = (state or "").upper()
    z = set()
    t1a = _county_in(st, county, TIER1_TX_VA)
    t1b = _county_in(st, county, TIER1_MD_ME)
    t2 = _county_in(st, county, TIER2)
    if t1a or t1b:
        z.add("TIER1")
    if t2:
        z.add("TIER2")
    if st == "FL" or (st == "TX" and t1a):
        z.add("FL_TX1")
    if (st in ("TX", "LA") and t2) or (st in _GULF_ATLANTIC_TIER1_STATES and t1a):
        z.add("GULF_ATL")
    if t1b:
        z.add("MD_ME")
    if st == "HI":
        z.add("HI")
    if _county_in(st, county, hail_belt or {}):
        z.add("HAIL")
    return z


def _loc_label(r):
    return f"{r.get('street', '')}, {r.get('city', '')}, {r.get('state', '')}".strip(", ")


_ZONE_TITLES = {
    "FL_TX1": "Florida / Texas Tier One counties",
    "GULF_ATL": "Texas & Louisiana Tier Two, and Tier One counties Louisiana through Virginia",
    "MD_ME": "Tier One counties Maryland through Maine",
    "HI": "Hawaii",
    "HAIL": "Hail Belt counties",
}


def applicable_deductibles(hb):
    """
    Reduce the A-H deductible schedule to the clauses that apply to the scheduled locations.
    Returns (rows, notes): rows = [{"letter","description","amount","applies_to":[labels]}],
    notes = list of sentences for the paragraph under the table.
    """
    sched = hb.get("schedule") or []
    hail = hb.get("hail_belt") or {}
    # unique locations (buildings at one address collapse)
    locs = []
    seen = set()
    for r in sched:
        lab = _loc_label(r)
        if lab in seen:
            continue
        seen.add(lab)
        locs.append({"label": lab, "zones": location_zones(r.get("state"), r.get("county"), hail),
                     "state": r.get("state", ""), "county": r.get("county", "")})
    all_labels = [l["label"] for l in locs]

    rows, notes = [], []
    skipped = []
    for cl in hb.get("deductible_clauses") or []:
        kind = cl["kind"]
        if kind == "aop":
            rows.append({"letter": cl["letter"], "description": "All Other Perils (per Occurrence)",
                         "amount": _fmt_money(cl["amount"]), "applies_to": list(all_labels)})
        elif kind == "water_damage":
            rows.append({"letter": cl["letter"], "description": "Water Damage",
                         "amount": f"{_fmt_money(cl['amount'])} each Building", "applies_to": list(all_labels)})
        elif kind == "waiting_period":
            hrs = cl.get("hours") or "24"
            rows.append({"letter": cl["letter"],
                         "description": "Waiting Period — Business Income, Ingress/Egress, Civil or Military Authority, Service Interruption",
                         "amount": f"{hrs} consecutive hours per location, then excess of the applicable deductible",
                         "applies_to": list(all_labels)})
        elif kind in ("named_windstorm", "hail_belt"):
            zone = cl.get("zone")
            hit = [l["label"] for l in locs if zone in l["zones"]]
            peril = "Named Windstorm" if kind == "named_windstorm" else "Wind, Tornado & Hail (not in conjunction with a Named Windstorm)"
            desc = f"{peril} — {_ZONE_TITLES.get(zone, zone)}"
            amt = (f"{cl['percent']} of the total replacement cost value of each Unit of Insurance "
                   f"(each building, its contents, and its annual Business Interruption value), per Occurrence")
            if cl.get("minimum"):
                amt += f"; minimum {_fmt_money(cl['minimum'])} per Occurrence"
            if hit:
                rows.append({"letter": cl["letter"], "description": desc, "amount": amt, "applies_to": hit})
            else:
                skipped.append(desc)
        else:
            rows.append({"letter": cl["letter"], "description": f"Deductible {cl['letter']}",
                         "amount": cl["text"], "applies_to": list(all_labels)})

    if skipped:
        if len(skipped) == len([c for c in hb.get("deductible_clauses") or [] if c["kind"] in ("named_windstorm", "hail_belt")]):
            notes.append("The HotelBound program's percentage deductibles for Named Windstorm (Florida, Texas and "
                         "Louisiana Tier One / Tier Two counties, Tier One counties Louisiana through Maine, and Hawaii) "
                         "and for Wind, Tornado and Hail in the Hail Belt do not apply to the scheduled location(s), "
                         "which sit outside those territories. Only the deductibles shown above apply.")
        else:
            notes.append("Not applicable to the scheduled locations: " + "; ".join(skipped) + ".")
    hail_named = any("HAIL" in l["zones"] and "TIER1" not in l["zones"] for l in locs)
    if hail_named and any(r["letter"] and r["description"].startswith("Wind, Tornado") for r in rows):
        notes.append("If a loss at a Hail Belt location outside a Tier One county is caused by a Named Windstorm, "
                     "the deductible is the greater of the All Other Perils deductible and the Wind, Tornado & Hail deductible.")
    notes.append("Deductibles apply per Occurrence to each Insured Member. Where two or more deductibles apply to one "
                 "Occurrence, the total deducted will not exceed the single largest applicable deductible; waiting periods "
                 "apply in addition to the property damage deductible. Percentage deductibles apply separately to each "
                 "Unit of Insurance (each separate building or structure, its contents, property in the open, and its "
                 "annual Business Interruption value).")
    tier_locs = [l["label"] for l in locs if l["zones"] & {"TIER1", "TIER2"}]
    hb["_tier_locations"] = tier_locs
    return rows, notes


# ---------------------------------------------------------------------------------------------
# apply to extracted data
# ---------------------------------------------------------------------------------------------
def apply_hotelbound_to_data(data: dict, hb: dict) -> None:
    """Write the parsed HotelBound quote into the extracted-data dict (mutates in place)."""
    if not hb or not hb.get("detected"):
        return
    coverages = data.setdefault("coverages", {})
    prop = coverages.get("property")
    if not isinstance(prop, dict):
        prop = {}
        coverages["property"] = prop

    # -- carrier / program identity
    if not (prop.get("carrier") or "").strip() or "lloyd" not in (prop.get("carrier") or "").lower():
        prop["carrier"] = hb.get("lead_insurer") or prop.get("carrier") or "Certain Underwriters at Lloyd's, London"
    prop["carrier_admitted"] = False
    prop["program"] = "HotelBound Insurance Program"
    prop["policy_form"] = prop.get("policy_form") or "HotelBound Manuscript Wording"
    if hb.get("policy_period"):
        prop["policy_period"] = hb["policy_period"]

    # -- 1. premium: base premium + all fees/taxes = Total Policy Cost
    if hb.get("base_premium"):
        prop["premium"] = hb["base_premium"]
    if hb.get("total_policy_cost"):
        prop["total_premium"] = hb["total_policy_cost"]
        if hb.get("base_premium"):
            prop["taxes_fees"] = round(hb["total_policy_cost"] - hb["base_premium"], 2)
    # itemized cost lines for the Premium & Fees table: memo fees (brokerage / inspection /
    # association) replace the RT cover's single "Carrier Policy Fee" line when they reconcile.
    items = []
    memo_fee_items = [f for f in hb.get("memo_fees", []) if "premium" not in f["label"].lower()]
    memo_fee_total = round(sum(f["amount"] for f in memo_fee_items), 2)
    for c in hb.get("cost_summary", []):
        ll = c["label"].lower()
        if "carrier policy fee" in ll and memo_fee_items and abs(memo_fee_total - c["amount"]) < 1.0:
            for f in memo_fee_items:
                if f["amount"] > 0:
                    items.append({"label": f["label"], "amount": f["amount"]})
        else:
            items.append({"label": c["label"], "amount": c["amount"]})
    if not items and hb.get("base_premium"):
        items.append({"label": "Property Premium", "amount": hb["base_premium"]})
    prop["premium_breakdown"] = items
    prop["taxes_fees_breakdown"] = {i["label"]: i["amount"] for i in items if "premium" not in i["label"].lower()}

    # -- 2. subjectivities
    if hb.get("subjectivities"):
        prop["subjectivities"] = list(hb["subjectivities"])
    if hb.get("warranties"):
        prop["warrants"] = [{"condition": w,
                             "consequence": "Warranty — a breach may allow insurers to deny a related loss"}
                            for w in hb["warranties"]]

    # -- sublimits A..GG replace whatever GPT found for the property extensions table
    if hb.get("sublimits"):
        prop["additional_coverages"] = [
            {"description": s["description"], "limit": s["limit"]} for s in hb["sublimits"]
        ]
    # Manuscript limitations (Late Notice, Roof ACV 15 yrs+, Cosmetic Damage, Wind Driven Elements...)
    # go into forms_endorsements so the generator's roof / high-risk highlighting catches them.
    if hb.get("limitations"):
        forms = [f for f in (prop.get("forms_endorsements") or []) if isinstance(f, dict)]
        have = {re.sub(r"[^a-z0-9]", "", (f.get("description") or "").lower()) for f in forms}
        for lim in hb["limitations"]:
            k = re.sub(r"[^a-z0-9]", "", lim.lower())
            if k and k not in have:
                forms.append({"form_number": "HB Manuscript", "description": lim})
                have.add(k)
        prop["forms_endorsements"] = forms
    if hb.get("attachments"):
        prop["policy_attachments"] = list(hb["attachments"])

    # -- 3. terrorism
    prop["terrorism_included"] = bool(hb.get("terrorism_included"))
    prop["terrorism_text"] = hb.get("terrorism", "")

    # -- 4. deductibles (only what applies to the scheduled locations)
    rows, notes = applicable_deductibles(hb)
    if rows:
        prop["deductibles"] = [
            {"description": r["description"], "type": r["description"], "amount": r["amount"],
             "applies_to": r["applies_to"], "letter": r["letter"]}
            for r in rows
        ]
    prop["deductible_notes"] = notes

    # -- 5. earned premium / fees disclosure
    disclosure = []
    if hb.get("program_brokerage_fee"):
        disclosure.append(
            f"RT Specialty has charged a program brokerage fee of {_fmt_money(hb['program_brokerage_fee'])}. "
            "This fee is collected and used for costs associated with underwriting, administration, servicing, "
            "marketing, claim reporting and loss control of the HotelBound Master Policy and is 100% fully earned "
            "at binding; it is not refundable under any circumstances, including cancellation."
        )
    elif hb.get("program_brokerage_fee_text"):
        disclosure.append(hb["program_brokerage_fee_text"])
    if hb.get("association_fee"):
        disclosure.append(f"The HotelBound Association fee of {_fmt_money(hb['association_fee'])} is fully earned at inception.")
    for f in hb.get("fully_earned_fees", []):
        if "brokerage" in f.lower() and hb.get("program_brokerage_fee"):
            continue
        disclosure.append(f"Fully earned fee: {f}.")
    if hb.get("mep_text"):
        disclosure.append("Minimum Earned Premium: " + hb["mep_text"])
    elif hb.get("mep_amount"):
        disclosure.append(f"Minimum Earned Premium: {hb.get('mep_percent') or '35%'} of the non-admitted premium plus "
                          f"100% of the aggregate contribution, calculated at {_fmt_money(hb['mep_amount'])} for this insured.")
    if hb.get("mep_tier_condition"):
        disclosure.append(hb["mep_tier_condition"] + ".")
    disclosure.append("Hurricane / Named Windstorm minimum earned premium provisions may apply per the policy form.")
    prop["earned_premium_disclosure"] = disclosure

    # -- 6. schedule of values by location & building
    sched = hb.get("schedule") or []
    if sched:
        cbl = []
        for r in sched:
            cbl.append({
                "premise": r.get("location_number"), "building": r.get("building_number"),
                "label": r.get("label"), "lid": r["lid"],
                "address": f"{r['street']}, {r['city']}, {r['state']} {r['zip']}",
                "street": r["street"], "city": r["city"], "state": r["state"], "zip": r["zip"],
                "county": r.get("county", ""), "flood_zone": r.get("flood_zone", ""),
                "year_built": r.get("year_built", ""), "occupancy": r.get("occupancy", ""),
                "construction": r.get("construction", ""), "iso_code": r.get("iso_code"),
                "num_stories": r.get("num_stories"), "num_units": r.get("num_units"),
                "sprinklered": r.get("sprinklered", ""), "sqft": r.get("sqft", 0),
                "building_value": _fmt_money(r["bldg_value"]), "bpp_value": _fmt_money(r["bpp_value"]),
                "business_income": _fmt_money(r["bi_value"]), "tiv": _fmt_money(r["tiv"]),
                "building_value_num": r["bldg_value"], "bpp_value_num": r["bpp_value"],
                "bi_value_num": r["bi_value"], "tiv_num": r["tiv"],
                "premium": r.get("premium", 0.0), "fee": r.get("fee", 0.0), "agg_premium": r.get("agg_premium", 0.0),
                "valuation": "RC", "cause_of_loss": "All Risk excl. Flood & Earth Movement",
            })
        prop["coverage_by_location"] = cbl
        prop["schedule_of_values"] = [
            {"location": f"{c['label']}: {c['address']}", "address": c["address"],
             "building": c["building_value_num"], "contents": c["bpp_value_num"],
             "business_income": c["bi_value_num"], "tiv": c["tiv_num"]}
            for c in cbl
        ]

    tiv = hb.get("tiv") or sum(r["tiv"] for r in sched)
    if tiv:
        prop["tiv"] = _fmt_money(tiv)
    limit = hb.get("program_limit") or tiv
    if limit:
        prop["limits"] = [
            {"description": "Overall Program Limit of Liability any one Occurrence (except Flood and Earth Movement)",
             "limit": _fmt_money(limit)},
            {"description": "Real Property (Building)", "limit": "Scheduled per location — see Schedule of Insured Values"},
            {"description": "Business Personal Property (Contents)", "limit": "Scheduled per location — see Schedule of Insured Values"},
            {"description": "Business Income", "limit": "Actual Loss Sustained, scheduled per location"
                + (f"; Monthly Limit of Indemnity {hb['monthly_limit_of_indemnity']}" if hb.get("monthly_limit_of_indemnity") else "")},
            {"description": "Terrorism", "limit": "Included — premium included in Total Policy Cost" if hb.get("terrorism_included") else (hb.get("terrorism") or "See quote")},
        ]
    if hb.get("valuation"):
        prop["valuation"] = "; ".join(hb["valuation"])

    # -- top-level locations[] so Schedule of Locations / Info Summary reflect the quote when no SOV
    # One entry per LOCATION (buildings at the same address collapse, TIV summed) — the Schedule
    # of Locations page de-duplicates by address and would otherwise show only the first building.
    if sched:
        acct = hb.get("account_name") or hb.get("insured") or ""
        rebuilt = []
        by_loc = {}
        for r in sched:
            key = r.get("location_number") or r["lid"]
            if key not in by_loc:
                by_loc[key] = {
                    "number": str(r.get("location_number") or len(by_loc) + 1),
                    "name": acct, "corporate_entity": (data.get("client_info") or {}).get("named_insured", "") or acct,
                    "dba": "", "address": r["street"], "city": r["city"], "state": r["state"], "zip": r["zip"],
                    "county": r.get("county", ""), "description": r.get("occupancy") or "Hotel/Motel",
                    "tiv": 0.0, "year_built": r.get("year_built", ""), "construction": r.get("construction", ""),
                    "sprinklered": r.get("sprinklered", ""), "num_units": 0, "num_buildings": 0,
                }
                rebuilt.append(by_loc[key])
            loc = by_loc[key]
            loc["tiv"] += r["tiv"]
            loc["num_units"] = (loc["num_units"] or 0) + (r.get("num_units") or 0)
            loc["num_buildings"] += 1
        data["locations"] = rebuilt

    # -- payment options / MEP row for the Payment Options page
    po = data.setdefault("payment_options", [])
    if isinstance(po, list):
        po[:] = [p for p in po if not (isinstance(p, dict) and "property" in (p.get("coverage_type") or "").lower())]
        mep = ""
        if hb.get("mep_amount"):
            mep = (f"{hb.get('mep_percent') or '35%'} of non-admitted premium + 100% aggregate contribution "
                   f"= {_fmt_money(hb['mep_amount'])}")
        po.append({
            "carrier": (prop.get("carrier") or "Lloyd's of London") + " (HotelBound)",
            "coverage_type": "Property",
            "terms": hb.get("payment_terms") or "Payment due within 15 days of inception",
            "mep": mep,
        })

    # -- everything else the generator needs, in one place
    prop["hotelbound"] = {
        "detected": True,
        "account_name": hb.get("account_name", ""),
        "control_number": hb.get("control_number", ""),
        "home_state": hb.get("home_state", ""),
        "lead_insurer": hb.get("lead_insurer", ""),
        "policy_period": hb.get("policy_period", ""),
        "coverage_type_text": hb.get("coverage_type_text", ""),
        "tiv": tiv,
        "program_limit": limit,
        "monthly_limit_of_indemnity": hb.get("monthly_limit_of_indemnity", ""),
        "territorial_limits": hb.get("territorial_limits", ""),
        "terrorism": hb.get("terrorism", ""),
        "terrorism_included": bool(hb.get("terrorism_included")),
        "valuation": hb.get("valuation", []),
        "payment_terms": hb.get("payment_terms", ""),
        "commission": hb.get("commission", ""),
        "base_premium": hb.get("base_premium", 0.0),
        "total_policy_cost": hb.get("total_policy_cost", 0.0),
        "memo_total": hb.get("memo_total", 0.0),
        "program_brokerage_fee": hb.get("program_brokerage_fee", 0.0),
        "association_fee": hb.get("association_fee", 0.0),
        "inspection_fee": hb.get("inspection_fee", 0.0),
        "mep_amount": hb.get("mep_amount", 0.0),
        "mep_percent": hb.get("mep_percent", ""),
        "tier_locations": hb.get("_tier_locations", []),
        "waiting_period_hours": hb.get("waiting_period_hours", ""),
        "deductible_clauses": hb.get("deductible_clauses", []),
        "shared_limits": hb.get("shared_limits", {}),
    }
    logger.info(
        f"HotelBound applied: premium={prop.get('premium')}, total={prop.get('total_premium')}, "
        f"deductibles={len(prop.get('deductibles') or [])}, buildings={len(sched)}, "
        f"subjectivities={len(prop.get('subjectivities') or [])}"
    )


def find_hotelbound_text(items):
    """Given [{'filename','text'}] pick the file that is a HotelBound quote (full, untruncated text)."""
    best = None
    for it in items or []:
        t = it.get("text") or ""
        u = t.upper()
        if "HOTELBOUND" in u and "QUOTATION MEMORANDUM" in u:
            if best is None or len(t) > len(best):
                best = t
    return best
