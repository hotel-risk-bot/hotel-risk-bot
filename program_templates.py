"""
Program templates: build the Starr property and PartnerRe umbrella coverages for a proposal
from HUB's exclusive program terms instead of a carrier quote (Oct 15 2026).

Why: the Starr Surplus Lines property program (placed through Amwins) and the PartnerRe
umbrella (Heartland Hospitality Franchise Program, Alta Risk) are exclusive to HUB, so the
forms, sublimits, terms and layout never change. Only the client details, the deductibles
(property), the limit (umbrella) and the premium change. Stefan, Oct 15 2026:
  * premium is typed in each time, already including taxes and fees (no tax calculation);
  * every property deductible is selectable: AOP, Water Damage, Earth Movement, Wind/Hail
    (flat or % of location TIV with a minimum) and Named Storm % for Tier 1 counties;
  * Boiler & Machinery is a per-proposal checkbox (unticked when a separate EB quote, such as
    Travelers, is part of the proposal);
  * umbrella limits $1M / $2M / $3M / $5M;
  * the property section is built off the Amwins SOV.

Templates are the Amwins "Starr Ground Up x Flood Quote" (Segen Group LLC, 10/31/2026) and the
Alta "5xP Excess Quote" (Segen Group LLC, 10/31/2026). Nothing here calls an AI model: the
coverages are deterministic and render through the normal generator path. Wholesaler / MGA
names (Amwins, Alta Risk, Heartland) never appear in what this module builds.

Entry point: apply_program_templates(data, inputs, sov_data) -> list of notes.
"""

from __future__ import annotations

import copy
import re
from datetime import datetime

try:  # the Tier 1 county tables live with the HotelBound program
    from hotelbound_territories import TIER1_TX_VA, TIER1_MD_ME
except Exception:  # pragma: no cover - module always ships with the repo
    TIER1_TX_VA, TIER1_MD_ME = {}, {}


STARR_KEY = "starr_property"
PARTNERRE_KEY = "partnerre_umbrella"

STARR_CARRIER = "Starr Surplus Lines Insurance Company"
STARR_AM_BEST = "A (XV)"
PARTNERRE_CARRIER = "PartnerRe Insurance Solutions Bermuda Ltd."
PARTNERRE_AM_BEST = "A+ (XV)"

STARR_TRIA_RATE = 0.05      # $1,100 on $22,000 in the template
PARTNERRE_TRIA_RATE = 0.02  # $160 on $8,000 in the template

UMBRELLA_LIMITS = (1_000_000, 2_000_000, 3_000_000, 5_000_000)

# Defaults straight from the Starr template
STARR_DEFAULT_DEDUCTIBLES = {
    "aop": 10_000,
    "water": 25_000,
    "earth_movement": 25_000,
    "wind": {"mode": "flat", "amount": 25_000, "pct": 2, "minimum": 25_000},
    "named_storm": {"apply": "auto", "pct": 5, "minimum": 100_000},
}


# ─── Starr property template content ─────────────────────────────────────────

_STARR_SUBLIMITS = [
    ("Earth Movement", "$250,000 per occurrence and in the annual aggregate"),
    ("Earth Movement in the States of Alaska, California or Hawaii", "Not Covered"),
    ("Earth Movement in the New Madrid or Pacific Northwest Seismic Zones (Appendices A and B)", "Not Covered"),
    ("Maximum payable for all Earth Movement losses in any one policy term", "$250,000"),
    ("Flood (including Storm Surge)", "Not Covered"),
    ("Flood (including Storm Surge) at any location wholly or partially within FEMA Flood Zones A, AE, AH, AO, AR, A99, V or VE", "Not Covered"),
    ("Accounts Receivable", "$1,000,000"),
    ("Backup of Sewers and Drains", "$100,000"),
    ("Course of Construction", "$1,500,000"),
    ("Data Processing", "$500,000"),
    ("Debris Removal", "The greater of 25% of the adjusted direct property loss or $2,500,000"),
    ("Demolition and Increased Cost of Construction (undamaged portion)", "Included"),
    ("Demolition and Increased Cost of Construction", "$1,000,000"),
    ("Extra Expense", "$500,000"),
    ("Fine Arts", "$100,000 ($10,000 any one item)"),
    ("Fire and Police Department Service Charges", "$50,000"),
    ("Leased or Rented Equipment", "$100,000"),
    ("Leasehold Interest", "$250,000"),
    ("Mobile Equipment", "$100,000 ($25,000 maximum per item)"),
    ("Newly Acquired Locations", "$1,000,000"),
    ("Pollution and Contamination Clean Up", "$50,000 per occurrence and in the annual aggregate"),
    ("Protection and Preservation of Property", "$50,000"),
    ("Signs", "$100,000"),
    ("Spoilage", "$100,000"),
    ("Temporary Removal of Property", "$50,000"),
    ("Transit", "$50,000"),
    ("Trees and Shrubs", "$50,000 (not to exceed $1,000 per tree or shrub)"),
    ("Unnamed Locations", "$250,000"),
    ("Valuable Papers and Records", "$500,000"),
    ("Vehicles", "Not Covered"),
    ("Vacant Locations at Time of Loss: Theft", "$250,000"),
    ("Vacant Locations at Time of Loss: Vandalism and Malicious Mischief", "$250,000"),
    ("Vacant Locations at Time of Loss: Water Damage", "$250,000"),
]

_STARR_BM_SUBLIMITS = [
    ("Boiler and Machinery", "Included in the Policy Limit of Liability, any one accident"),
    ("Boiler and Machinery: Ammonia Contamination", "$100,000 any one accident"),
    ("Boiler and Machinery: Consequential Damage", "$100,000 any one accident"),
    ("Boiler and Machinery: Expediting Expenses", "$100,000 any one accident"),
    ("Boiler and Machinery: Hazardous Substances", "$100,000 any one accident"),
    ("Boiler and Machinery: Water Damage", "$500,000 any one accident"),
]

_STARR_TIME_LIMITS = [
    ("Civil and Military Authority (time limit)", "30 consecutive days"),
    ("Ingress / Egress (time limit)", "30 consecutive days"),
    ("Newly Acquired Locations (time limit)", "90 consecutive days"),
    ("Extended Period of Indemnity", "180 consecutive days"),
    ("Off Premises Power", "$50,000 (48 hour waiting period)"),
]

# (form number, description) in the template's order; "bm" marks the B&M endorsement
_STARR_FORMS = [
    ("SSPN-018 07 23", "Notice to Texas Policy Holder"),
    ("PR 911", "Starr Surplus Lines Insurance Company Declarations"),
    ("PR 907", "Policy Security Page"),
    ("IL 00 17 11 98", "Common Policy Conditions"),
    ("CP 00 90 07 88", "Commercial Property Conditions"),
    ("IL P 001 01 04", "Advisory Notice to Policyholders (OFAC)"),
    ("PR 908", "Policyholder State Notices"),
    ("PR 001 D (05-12)", "Property Coverage Declarations Form"),
    ("PR 002 (03/23)", "Property Coverage Form General Conditions"),
    ("PR 003 (02/12)", "Property Coverage Form - Property Section"),
    ("PR 004 (02/12)", "Property Coverage Form Business Interruption Section"),
    ("PR 006 (02/12)", "Accounts Receivable Endorsement"),
    ("PR 902", "Additional Insureds and Loss Payees Endorsement"),
    ("PR 007 (02/12)", "Agreed Amount Endorsement (Business Interruption)"),
    ("PR 008 (02/12)", "Agreed Amount Endorsement (Property)"),
    ("PR 903", "Application of Sublimits Endorsement"),
    ("PR 913", "Backup of Sewers and Drains"),
    ("PR 926", "Biological, Chemical or Nuclear Exclusion"),
    ("PR 012 (07/13)", "Boiler and Machinery Endorsement", "bm"),
    ("PR 013 (02/12)", "Business Interruption Extended Period of Indemnity Endorsement"),
    ("PR 018 (02/12)", "Course of Construction Endorsement"),
    ("Starr 04/20", "Communicable Disease Exclusion"),
    ("PR 034 (04/21)", "Demolition & Increased Cost of Construction Endorsement"),
    ("PR 022 (02/12)", "Earth Movement Endorsement"),
    ("PR 023 (03/23)", "Data Processing Endorsement"),
    ("PR 024 (03/23)", "Electronic Date Recognition Clause Endorsement (Combined)"),
    ("PR 028 (02/12)", "Extra Expense Endorsement"),
    ("PR 030 (11/16)", "Fine Arts Endorsement"),
    ("PR 029 (02/12)", "Fire and Police Department Service Charges Endorsement"),
    ("PR 035 (02/12)", "Ingress/Egress Endorsement"),
    ("PR 037 (02/12)", "Leasehold Interest Endorsement"),
    ("PR 039 (07/13)", "Lender's Loss Payable Endorsement"),
    ("PR 925", "Minimum Earned Premium Endorsement"),
    ("PR 041 (11/16)", "Mobile Equipment Endorsement"),
    ("PR 043 (11/16)", "Newly Acquired Locations Endorsement"),
    ("PR 044 (02/12)", "Occurrence Limit of Liability Endorsement"),
    ("PR 045 (02/12)", "Off-Premises Power Endorsement"),
    ("PR 921", "Policy Amendment Endorsement"),
    ("PR 918", "Policy Amendment Endorsement - Additional Named Insured"),
    ("PR 906", "Policy Amendment Endorsement - Territory"),
    ("PR 049 (02/12)", "Pollution and Contamination Clean-Up Endorsement"),
    ("PR 909", "Pre-Existing Damages Exclusion"),
    ("LMA 5401 (Starr amended 10/23)", "Property Cyber and Data Exclusion"),
    ("NMA1191", "Radioactive Contamination Exclusion"),
    ("PR 053 (02/12)", "Rental Value Insurance Endorsement"),
    ("PR 054 (09/14)", "Replacement Cost Endorsement"),
    ("PR 056 (09/14)", "Schedule of Locations Endorsement"),
    ("SSIL-0005 (07/20)", "Service of Process Clause Endorsement"),
    ("PR 912", "Spoilage Endorsement"),
    ("PR 057 (02/12)", "Sprinkler Leakage - Earth Movement Extension"),
    ("PR 058 (07/13)", "Standard Mortgage Endorsement"),
    ("PR 059 (02/12)", "Temporary Removal of Property Endorsement"),
    ("61330 (01/15)", "Terrorism Exclusion (for Certified Acts of Terrorism under the Terrorism Risk Insurance Act, as amended)"),
    ("61333 (01/15)", "Terrorism Risk Insurance Act, as amended, Cap on Losses Endorsement"),
    ("61331 (01/15)", "Total Terrorism Exclusion"),
    ("PR 067 (02/12)", "Trade or Economic Sanctions Endorsement"),
    ("PR 064 (07/13)", "Transit Endorsement"),
    ("PR 065 (02/12)", "Unnamed Location Coverage Endorsement (Real and Personal Property)"),
    ("PR 099 (04/21)", "Vacant Property Endorsement"),
    ("PR 066 (07/13)", "Valuable Papers and Records Endorsement"),
    ("NMA2918", "War and Terrorism Exclusion (as respects Transit)"),
    ("PR 073 (02/12)", "Appendix A - New Madrid Seismic Zone"),
    ("PR 074 (02/12)", "Appendix B - Pacific Northwest Seismic Zone"),
    ("PR 075 (09/15)", "Appendix C - Definition of Tier 1 Wind Counties"),
    ("PR 904", "Claims Notice"),
    ("PR 928", "Policy Amendment Endorsement - Water"),
    ("PR 200 (01/26)", "Roof Limitation Endorsement"),
]

_STARR_TERMS = [
    "Coverages and/or extensions of coverage not specifically mentioned, even though they may be "
    "outlined in the submission, are not included.",
    "30% Minimum Earned Premium. Premium is to be paid in full within 30 days of inception.",
    "Business Interruption coverage applies only to locations where Business Interruption values "
    "have been declared.",
    "A 72 hour occurrence definition applies to Wind, Flood, Earth Movement, riot, riot attending a "
    "strike, civil commotion and vandalism and malicious mischief.",
    "Owned electrical transmission and distribution lines and their supporting structures located "
    "beyond 1,000 feet from any insured premises are excluded.",
    "All deductibles apply per occurrence, except coverage under the Boiler and Machinery "
    "endorsement, which applies any one accident. When two or more deductibles apply to an "
    "occurrence, the total deducted will not exceed the largest applicable deductible; a deductible "
    "that applies per location applies separately to each location where physical damage occurred.",
    "Percentage deductibles are calculated on 100% of the values reported for the location where the "
    "loss occurred (property damage), and on 100% of the Time Element values that would have been "
    "earned in the 12 months following the occurrence at that location (time element).",
    "TRIA election forms are due within 15 days of binding or the insured will be billed for Terrorism.",
    "All requested binding documents must be received no later than 30 days after the effective date; "
    "otherwise notice of cancellation will be issued.",
    "The company may withdraw this quotation at any time prior to the effective date, and coverage may "
    "not be bound without prior written authorization from the company.",
]

_STARR_SUBJECTIVITIES = [
    "Written request to bind coverage",
    "Signed TRIA election / rejection form (due within 15 days of binding)",
    "Signed state surplus lines packet",
    "Signed Statement of Property Values (within 30 days of the effective date)",
    "Completed Business Income worksheet (within 30 days of the effective date)",
    "Currently valued loss runs",
]


# ─── PartnerRe umbrella template content ─────────────────────────────────────

_PARTNERRE_FORMS = [
    ("PRE HRT 0101 08 26", "Commercial Excess Liability Declarations"),
    ("PRE ALT 0102 01 26", "Schedule of Forms and Endorsements"),
    ("PRE ALT 0104 01 26", "Schedule of Controlling Underlying Insurance"),
    ("PRE ALT 0104M 01 26", "Schedule of Controlling Underlying Insurance - Manuscript"),
    ("PRE ALT CN 1302 01 26", "Claims Notice"),
    ("PRE ALT 0107 01 26", "Common Policy Conditions"),
    ("PRE ALT 0108 01 26", "Minimum Earned Premium Endorsement"),
    ("PRE ALT 0109 01 26", "Service of Suit Clause"),
    ("PRE ALT SIG 01 01 26", "Signature Page"),
    ("PRE ALT 0113 01 26", "OFAC Endorsement"),
    ("PRE ALT 0114 01 26", "Cap on Losses from Certified Acts of Terrorism"),
    ("CX 0001 04 13", "Commercial Excess Liability Coverage Form"),
    ("PRE ALT 0202 01 26", "Access or Disclosure of Confidential or Personal Information Exclusion"),
    ("PRE ALT 0207 01 26", "Aircraft Liability Exclusion"),
    ("PRE ALT 0209 01 26", "Aircraft Products and Grounding Exclusion"),
    ("PRE ALT 0212 01 26", "Asbestos Exclusion"),
    ("PRE ALT 0221 01 26", "Automobile Liability Exclusion - Owned Auto"),
    ("PRE ALT 0239 01 26", "Communicable Disease Exclusion"),
    ("PRE ALT 0247 01 26", "Cross Suits Exclusion"),
    ("PRE ALT 0273 01 26", "ERISA Exclusion"),
    ("PRE ALT 0280 01 26", "Employee Benefits Liability Exclusion"),
    ("PRE ALT 0281 01 26", "Employers Liability Exclusion"),
    ("PRE ALT 0284 01 26", "Employment Related Practices Exclusion"),
    ("PRE ALT 0287 01 26", "Exterior Insulation and Finish Systems Exclusion"),
    ("PRE ALT 0300 01 26", "Fungi or Bacteria Exclusion"),
    ("PRE ALT 0326 01 26", "Lead Exclusion"),
    ("PRE ALT 0329 01 26", "Limitation of Coverage to Designated Premises or Project", "premises"),
    ("PRE ALT 0330 01 26", "Limited Coverage Territory"),
    ("PRE ALT 0332 01 26", "Liquor Liability Exclusion"),
    ("PRE ALT 0349 01 26", "Nuclear Energy Liability Exclusion"),
    ("PRE ALT 0351 01 26", "NY - Employee Injury Exclusion - Absolute"),
    ("PRE ALT 0374 01 26", "Pollution Exclusion - Total"),
    ("PRE ALT 0375 01 26", "Pre-Existing Injury or Damage Exclusion"),
    ("PRE ALT 0381 01 26", "Professional Services Exclusion"),
    ("PRE ALT 0403 01 26", "Recording and Distribution of Material or Information in Violation of Law Exclusion"),
    ("PRE ALT 0421 01 26", "Silica or Silica-Related Dust Exclusion"),
    ("PRE ALT 0424 07 26", "Sublimited Coverages Endorsement"),
    ("PRE ALT 0443 01 26", "Unmanned Aircraft Exclusion"),
    ("PRE ALT 0450 01 26", "War Exclusion"),
    ("PRE ALT 0452 01 26", "Electronic Data (Cyber Liability Exclusion)"),
    ("PRE ALT 0456 01 26", "PFAS Exclusion"),
    ("PRE ALT 0457 01 26", "Biometric Information Exclusion"),
    ("PRE ALT SLN {ST} 0126", "{ST} Surplus Lines Notice", "state"),
]

_PARTNERRE_TERMS = [
    "Regardless of the number of coverage parts quoted or bound, the limits of insurance shown apply "
    "once for the entire policy and not separately for each coverage part.",
    "Primary carrier(s) must be rated A- VII or better by A.M. Best.",
    "Minimum primary limits (when applicable): General Liability $1,000,000 / $2,000,000 / $1,000,000; "
    "Auto Liability $1,000,000 CSL; Employers Liability $500,000 / $500,000 / $500,000.",
    "Uninsured / Underinsured Motorists coverage, when applicable, is excluded unless otherwise noted.",
    # Certificates, underlying-policy copies and the 20-day net premium are the broker's terms with
    # the program manager, not the client's (Stefan, Oct 28 2026 - see subjectivity_rules.py).
    "25% Minimum Earned Premium. Premium is flat and non-auditable.",
    "Quotation is valid for 30 days from the quote date.",
]

# The client gets 5 days, not the program's 30; underlying-policy copies are the broker's item
# (Stefan, Oct 28 2026 - see subjectivity_rules.py).
_PARTNERRE_SUBJECTIVITIES = [
    "Signed and selected TRIA election (within 5 days of binding)",
    "Completed surplus lines filing confirmation (within 5 days of binding)",
    "ACORD 125 and 131 completed and signed by the insured (within 5 days of binding)",
]


# ─── helpers ─────────────────────────────────────────────────────────────────

def _money(v):
    """Parse 25000 / "25000" / "$25,000" / "25k" / "$1M" -> int dollars (0 when blank)."""
    if v is None or v == "":
        return 0
    if isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).strip().lower().replace(",", "").replace("$", "")
    m = re.match(r"^(\d+(?:\.\d+)?)\s*([km]|mm)?", s)
    if not m:
        return 0
    n = float(m.group(1))
    if m.group(2) == "k":
        n *= 1_000
    elif m.group(2) in ("m", "mm"):
        n *= 1_000_000
    return int(round(n))


def _amount(v):
    """Premium-style amount with cents kept (float)."""
    if v is None or v == "" or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = str(v).strip().replace(",", "").replace("$", "")
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0


def _pct(v, default):
    if v is None or v == "":
        return default
    try:
        n = float(str(v).replace("%", "").strip())
    except ValueError:
        return default
    return n if n > 0 else default


def _fmt(n):
    return f"${int(round(n)):,}"


def _fmt_pct(p):
    return (f"{p:g}%")


def _mdy(value):
    """Normalise a date to MM/DD/YYYY ('' when unparseable)."""
    if not value:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    s = str(value).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%Y", "%B %d, %Y"):
        try:
            return datetime.strptime(s[:19] if fmt.startswith("%Y-%m-%d %H") else s, fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
    if m:
        y = int(m.group(3))
        y = y + 2000 if y < 100 else y
        try:
            return datetime(y, int(m.group(1)), int(m.group(2))).strftime("%m/%d/%Y")
        except ValueError:
            return ""
    return ""


def _plus_one_year(mdy):
    try:
        d = datetime.strptime(mdy, "%m/%d/%Y")
    except (TypeError, ValueError):
        return ""
    try:
        return d.replace(year=d.year + 1).strftime("%m/%d/%Y")
    except ValueError:  # Feb 29
        return d.replace(year=d.year + 1, day=28).strftime("%m/%d/%Y")


def _norm_county(c):
    c = (c or "").lower().strip()
    c = re.sub(r"[’']", "", c)
    c = re.sub(r"\b(county|parish|borough)\b", "", c)
    c = re.sub(r"\bst\.?\s", "st ", c)
    c = re.sub(r"[^a-z0-9 ]", " ", c)
    return re.sub(r"\s+", " ", c).strip()


def is_tier1_county(state, county):
    """True when the location sits in a Tier 1 wind county (Appendix C style list: Gulf and
    Atlantic coastal counties, Florida and Hawaii whole-state)."""
    st = (state or "").strip().upper()
    for table in (TIER1_TX_VA, TIER1_MD_ME):
        entry = table.get(st)
        if entry == "ALL":
            return True
        if entry and _norm_county(county) in {_norm_county(c) for c in entry}:
            return True
    return False


def _state_from_address(addr):
    m = re.search(r",\s*([A-Z]{2})\s*\d{5}", str(addr or ""))
    if m:
        return m.group(1)
    m = re.search(r"\b([A-Z]{2})\s+\d{5}(?:-\d{4})?\s*$", str(addr or ""))
    return m.group(1) if m else ""


def _property_sov(sov_data):
    """The property (Amwins) SOV: the merged session SOV keeps it under _property_sov."""
    if not isinstance(sov_data, dict):
        return {}
    ps = sov_data.get("_property_sov")
    if isinstance(ps, dict) and ps.get("locations"):
        return ps
    return sov_data


def _sov_locations(sov):
    locs = [l for l in (sov.get("locations") or []) if isinstance(l, dict)]
    return [l for l in locs if (l.get("tiv") or 0) > 0] or locs


def _loc_label(loc):
    name = (loc.get("dba") or loc.get("corporate_name") or "").strip()
    city = (loc.get("city") or "").strip()
    if name and city:
        return f"{name} - {city}"
    return name or (loc.get("address") or "").strip() or "Location"


def _loc_address(loc):
    parts = [str(loc.get("address") or "").strip(), str(loc.get("city") or "").strip()]
    st_zip = " ".join(p for p in (str(loc.get("state") or "").strip(),
                                  str(loc.get("zip_code") or loc.get("zip") or "").strip()) if p)
    return ", ".join(p for p in parts + [st_zip] if p)


def _sov_aop(sov):
    vals = []
    for src in ("locations_raw", "locations"):
        for l in (sov.get(src) or []):
            if isinstance(l, dict):
                v = _money(l.get("aop_deductible"))
                if v:
                    vals.append(v)
    return max(vals) if vals else 0


def _premium_block(cov, inp, tria_rate):
    premium = _amount(inp.get("premium"))
    taxes = _amount(inp.get("taxes_fees"))
    total = _amount(inp.get("total_premium"))
    if not total and (premium or taxes):
        total = round(premium + taxes, 2)
    if total and premium and not taxes and total > premium:
        taxes = round(total - premium, 2)
    tria = _amount(inp.get("tria_premium"))
    if not tria and premium:
        tria = float(int(round(premium * tria_rate)))
    cov["premium"] = premium
    cov["taxes_fees"] = taxes
    cov["total_premium"] = total
    cov["tria_premium"] = tria
    return premium, taxes, total, tria


def _context(data, inputs, sov):
    ci = data.get("client_info") if isinstance(data.get("client_info"), dict) else {}
    summary = sov.get("summary") if isinstance(sov.get("summary"), dict) else {}
    eff = (_mdy(inputs.get("effective_date")) or _mdy(ci.get("effective_date"))
           or _mdy(summary.get("effective_date")))
    if not eff:
        gl = (data.get("coverages") or {}).get("general_liability") or {}
        eff = _mdy(gl.get("effective_date")) if isinstance(gl, dict) else ""
    exp = _plus_one_year(eff) if eff else ""
    mailing = (ci.get("address") or ci.get("mailing_address") or summary.get("mailing_address") or "")
    locs = _sov_locations(sov)
    home_state = (_state_from_address(mailing)
                  or next((str(l.get("state") or "").strip().upper() for l in locs if l.get("state")), ""))
    return {
        "named_insured": (ci.get("named_insured") or summary.get("named_insured") or "").strip(),
        "mailing": str(mailing).strip(),
        "effective": eff,
        "expiration": exp,
        "home_state": home_state,
        "locations": locs,
    }


def _policy_period(ctx):
    if ctx["effective"] and ctx["expiration"]:
        return f"{ctx['effective']} to {ctx['expiration']} (12:01 a.m. standard time at the mailing address)"
    return ""


# ─── Starr property ──────────────────────────────────────────────────────────

def _wind_text(wind, locs, named_storm_applies):
    label = ("Wind / Hail, other than Named Storm (property damage and time element combined)"
             if named_storm_applies else "Wind / Hail (property damage and time element combined)")
    mode = str(wind.get("mode") or "flat").lower()
    if mode in ("pct", "percent", "%"):
        pct = _pct(wind.get("pct"), 2)
        minimum = _money(wind.get("minimum")) or 25_000
        txt = (f"{_fmt_pct(pct)} of the total insured values at each location involved in the loss, "
               f"subject to a {_fmt(minimum)} minimum per occurrence")
        txt += _per_location_examples(locs, pct, minimum)
        return label, txt
    amount = _money(wind.get("amount")) or 25_000
    return label, f"{_fmt(amount)} Per Occurrence"


def _per_location_examples(locs, pct, minimum, only=None):
    rows = [l for l in locs if (only is None or id(l) in only) and (l.get("tiv") or 0) > 0]
    if not rows or len(rows) > 6:
        return ""
    parts = []
    for l in rows:
        amt = max(minimum, (l.get("tiv") or 0) * pct / 100.0)
        parts.append(f"{_loc_label(l)}: {_fmt(amt)}")
    return " (" + "; ".join(parts) + ")"


def build_starr_property(data, inputs, sov_data):
    """Return (coverage dict, notes)."""
    notes = []
    sov = _property_sov(sov_data)
    ctx = _context(data, inputs, sov)
    locs = ctx["locations"]
    totals = sov.get("totals") if isinstance(sov.get("totals"), dict) else {}
    tiv = (totals.get("tiv") or 0) or sum((l.get("tiv") or 0) for l in locs)
    bldg = (totals.get("building_value") or 0) or sum((l.get("building_value") or 0) for l in locs)
    cont = (totals.get("contents_value") or 0) or sum((l.get("contents_value") or 0) for l in locs)
    bi = (totals.get("bi_value") or 0) or sum((l.get("bi_value") or 0) for l in locs)
    other = sum(((l.get("other_value") or 0) + (l.get("sign_value") or 0) + (l.get("pool_value") or 0))
                for l in locs)

    ded_in = inputs.get("deductibles") if isinstance(inputs.get("deductibles"), dict) else {}
    include_bm = inputs.get("include_bm")
    has_eb_quote = bool((data.get("coverages") or {}).get("equipment_breakdown"))
    if include_bm is None or include_bm == "auto":
        include_bm = not has_eb_quote
    include_bm = bool(include_bm)

    cov = {
        "carrier": STARR_CARRIER,
        "carrier_admitted": False,
        "am_best_rating": STARR_AM_BEST,
        "program_template": STARR_KEY,
        "policy_form": "PR 001, PR 002, PR 003, PR 004 and accompanying endorsements",
        "policy_period": _policy_period(ctx),
        "effective_date": ctx["effective"],
        "expiration_date": ctx["expiration"],
        "tiv": _fmt(tiv) if tiv else "",
        "valuation": ("Replacement Cost (Actual Cash Value if the property is not repaired or replaced "
                      "within 2 years); Time Element: Actual Loss Sustained"),
        "coinsurance": [{"coverage": "All Coverages",
                         "percentage": "100% (Waived by Agreed Amount Endorsement)"}],
    }
    premium, taxes, total, tria = _premium_block(cov, inputs, STARR_TRIA_RATE)

    perils = ("Property Damage, Business Interruption, Extra Expense and Boiler & Machinery"
              if include_bm else "Property Damage, Business Interruption and Extra Expense")
    summary = [
        {"label": "Coverage", "value": f"All risks of direct physical loss or damage, covering {perils}"},
        {"label": "Coverage Territory", "value": "United States of America"},
        {"label": "Description and Location of Property Insured", "value": "Per the Schedule of Values on file"},
        {"label": "Carrier Participation", "value": "100%"},
        {"label": "Minimum Earned Premium", "value": "30%"},
        {"label": "Cancellation", "value": "Thirty (30) days, except ten (10) days for non-payment of premium"},
        {"label": "Terrorism (TRIA)",
         "value": (f"Optional, {_fmt(tria)} additional premium plus applicable taxes and fees; "
                   "signed election or rejection required at binding") if tria else
                  "Optional, available for an additional premium; signed election or rejection required at binding"},
        {"label": "Flood", "value": "Excluded"},
    ]
    if not include_bm:
        summary.append({"label": "Boiler & Machinery",
                        "value": ("Not included; Equipment Breakdown is provided under a separate policy"
                                  if has_eb_quote else "Not included")})
    cov["summary_rows"] = summary

    limits = [{"description": "Policy Limit of Liability",
               "limit": f"{_fmt(tiv)} any one occurrence, excess of policy deductibles" if tiv else "Per SOV"}]
    if bldg:
        limits.append({"description": "Building", "limit": _fmt(bldg)})
    if cont:
        limits.append({"description": "Business Personal Property / Contents", "limit": _fmt(cont)})
    if bi:
        limits.append({"description": "Business Income / Rents (Actual Loss Sustained)", "limit": _fmt(bi)})
    if other:
        limits.append({"description": "Other Property (signs, pools and other values)", "limit": _fmt(other)})
    cov["limits"] = limits

    # deductibles
    sov_aop = _sov_aop(sov)
    aop = _money(ded_in.get("aop")) or sov_aop or STARR_DEFAULT_DEDUCTIBLES["aop"]
    water = _money(ded_in.get("water")) or STARR_DEFAULT_DEDUCTIBLES["water"]
    em = _money(ded_in.get("earth_movement")) or STARR_DEFAULT_DEDUCTIBLES["earth_movement"]
    wind = dict(STARR_DEFAULT_DEDUCTIBLES["wind"])
    if isinstance(ded_in.get("wind"), dict):
        wind.update({k: v for k, v in ded_in["wind"].items() if v not in (None, "")})
    ns = dict(STARR_DEFAULT_DEDUCTIBLES["named_storm"])
    if isinstance(ded_in.get("named_storm"), dict):
        ns.update({k: v for k, v in ded_in["named_storm"].items() if v not in (None, "")})

    tier1 = [l for l in locs if is_tier1_county(l.get("state"), l.get("county"))]
    ns_mode = str(ns.get("apply") or "auto").lower()
    if ns_mode == "always":
        ns_locs = list(locs)
    elif ns_mode == "never":
        ns_locs = []
    else:
        ns_locs = tier1
    if tier1 and ns_mode == "never":
        notes.append("Starr template: " + ", ".join(_loc_label(l) for l in tier1) +
                     " sits in a Tier 1 wind county but the Named Storm deductible was set to 'Never'. "
                     "Confirm the Named Storm terms with Starr before presenting.")

    deds = [{"description": "All Other Perils (property damage and time element combined)",
             "amount": f"{_fmt(aop)} Per Occurrence"}]
    deds.append({"description": "Water Damage", "amount": f"{_fmt(water)} Per Occurrence"})
    deds.append({"description": "Earth Movement", "amount": f"{_fmt(em)} Per Occurrence"})
    w_label, w_txt = _wind_text(wind, locs, bool(ns_locs))
    deds.append({"description": w_label, "amount": w_txt})
    if ns_locs:
        pct = _pct(ns.get("pct"), 5)
        minimum = _money(ns.get("minimum")) or 100_000
        txt = (f"{_fmt_pct(pct)} of the total insured values at each location involved in the loss, "
               f"subject to a {_fmt(minimum)} minimum per occurrence")
        txt += _per_location_examples(locs, pct, minimum, only={id(l) for l in ns_locs})
        label = "Named Storm, Tier 1 wind counties (Appendix C)"
        if len(locs) > 1 and len(ns_locs) < len(locs):
            label += ": applies to " + ", ".join(_loc_label(l) for l in ns_locs)
        deds.append({"description": label, "amount": txt})
    if include_bm:
        deds.append({"description": "Boiler and Machinery",
                     "amount": f"{_fmt(aop)} Any One Accident (property damage and time element combined)"})
    for d in deds:
        d["type"] = d["description"]
    cov["deductibles"] = deds

    sub = [{"description": d, "limit": v} for d, v in _STARR_SUBLIMITS]
    if include_bm:
        sub += [{"description": d, "limit": v} for d, v in _STARR_BM_SUBLIMITS]
    sub += [{"description": d, "limit": v} for d, v in _STARR_TIME_LIMITS]
    cov["additional_coverages"] = sub

    forms = []
    for row in _STARR_FORMS:
        if len(row) > 2 and row[2] == "bm" and not include_bm:
            continue
        forms.append({"form_number": row[0], "description": row[1], "template_form": True})
    cov["forms_endorsements"] = forms
    cov["terms_conditions"] = list(_STARR_TERMS)
    cov["subjectivities"] = list(_STARR_SUBJECTIVITIES)

    if not locs:
        notes.append("REVIEW REQUIRED - Starr property template: no SOV locations were found, so the "
                     "Schedule of Values and TIV are empty. Upload the Amwins SOV and apply the template again.")
    if not premium and not total:
        notes.append("REVIEW REQUIRED - Starr property template: no premium entered.")
    if include_bm and has_eb_quote:
        notes.append("Starr template: Boiler & Machinery is included in the Starr property AND a separate "
                     "Equipment Breakdown quote is in this proposal. Untick 'Include Boiler & Machinery' if the "
                     "EB quote replaces it.")
    if tier1:
        notes.append("Starr template: Tier 1 wind county location(s): " +
                     ", ".join(f"{_loc_label(l)} ({l.get('county') or '?'} County, {l.get('state')})" for l in tier1) +
                     (". Named Storm deductible applied." if ns_locs else "."))
    cov["template_summary"] = {
        "tiv": tiv, "locations": len(locs), "aop": aop, "aop_source": ("input" if _money(ded_in.get("aop"))
                                                                       else "sov" if sov_aop else "default"),
        "include_bm": include_bm, "named_storm_locations": len(ns_locs), "tier1_locations": len(tier1),
    }
    return cov, notes


# ─── PartnerRe umbrella ──────────────────────────────────────────────────────

def _limit_value(limits, *keys, exclude=()):
    for lim in limits or []:
        if not isinstance(lim, dict):
            continue
        d = str(lim.get("description") or lim.get("type") or "").lower()
        if any(k in d for k in keys) and not any(x in d for x in exclude):
            v = _money(lim.get("limit") or lim.get("amount"))
            if v:
                return v
    return 0


def _gl_hnoa(gl):
    v = _limit_value(gl.get("limits"), "hired", "non-owned", "non owned", "hnoa")
    if v:
        return v
    for ac in (gl.get("additional_coverages") or []):
        if isinstance(ac, dict):
            d = str(ac.get("description") or "").lower()
            if "hired" in d or "non-owned" in d or "non owned" in d:
                return _money(ac.get("limit")) or 1_000_000
    for f in (gl.get("forms_endorsements") or []):
        d = str(f.get("description") if isinstance(f, dict) else f).lower()
        if "hired" in d and ("non-owned" in d or "non owned" in d or "nonowned" in d):
            return 1_000_000
    return 0


def _underlying_from_gl(data, inputs):
    covs = data.get("coverages") or {}
    gl = covs.get("general_liability")
    if not isinstance(gl, dict) or not gl.get("carrier"):
        gl = next((v for k, v in covs.items() if k.startswith("general_liability")
                   and isinstance(v, dict) and v.get("carrier")), None)
    carrier_override = str(inputs.get("underlying_carrier") or "").strip()
    notes = []
    if isinstance(gl, dict) and gl.get("carrier"):
        carrier = carrier_override or str(gl.get("carrier")).strip()
        lims = gl.get("limits") or []
        occ = _limit_value(lims, "each occurrence", "per occurrence") or 1_000_000
        agg = _limit_value(lims, "general aggregate", "aggregate", exclude=("product", "completed")) or 2_000_000
        prod = _limit_value(lims, "product", "completed")
        pai = _limit_value(lims, "personal", "advertising") or occ
        hnoa = _gl_hnoa(gl)
    else:
        carrier = carrier_override or "To be confirmed"
        occ, agg, prod, pai, hnoa = 1_000_000, 2_000_000, 0, 1_000_000, 0
        notes.append("REVIEW REQUIRED - PartnerRe umbrella template: no General Liability quote was found in "
                     "this proposal, so the underlying schedule uses minimum limits ($1,000,000 / $2,000,000 / "
                     "$1,000,000) with the carrier 'To be confirmed'. Edit it in the JSON editor or add the GL quote.")
    gl_limits = (f"{_fmt(occ)} Each Occurrence / {_fmt(agg)} General Aggregate / "
                 f"{(_fmt(prod) if prod else 'Included')} Products-Completed Operations Aggregate / "
                 f"{_fmt(pai)} Personal and Advertising Injury")
    rows = [{"carrier": carrier, "coverage": "General Liability", "limits": gl_limits}]
    if hnoa:
        rows.append({"carrier": carrier, "coverage": "Hired and Non-Owned Auto", "limits": _fmt(hnoa)})
    return rows, carrier, notes


def build_partnerre_umbrella(data, inputs, sov_data):
    notes = []
    sov = _property_sov(sov_data) if sov_data else {}
    ctx = _context(data, inputs, sov or {})
    limit = _money(inputs.get("limit")) or 5_000_000
    if limit not in UMBRELLA_LIMITS:
        notes.append(f"PartnerRe umbrella template: limit {_fmt(limit)} is outside the program options "
                     "($1M / $2M / $3M / $5M). Confirm with the program before presenting.")
    lim_m = f"${limit // 1_000_000}M" if limit % 1_000_000 == 0 else _fmt(limit)

    underlying, ul_carrier, ul_notes = _underlying_from_gl(data, inputs)
    notes += ul_notes

    # designated premises: SOV locations first, then the GL schedule / designated premises
    addresses = []
    for l in (sov_data or {}).get("locations") or []:
        if isinstance(l, dict):
            a = _loc_address(l)
            if a and a not in addresses:
                addresses.append(a)
    if not addresses:
        for l in data.get("locations") or []:
            if isinstance(l, dict):
                a = ", ".join(p for p in (str(l.get("address") or "").strip(), str(l.get("city") or "").strip(),
                                          " ".join(x for x in (str(l.get("state") or "").strip(),
                                                               str(l.get("zip") or "").strip()) if x)) if p)
                if a and a not in addresses:
                    addresses.append(a)
    state = ctx["home_state"] or "TX"
    if not ctx["home_state"]:
        notes.append("PartnerRe umbrella template: the insured's state could not be read from the mailing address "
                     "or SOV; the surplus lines notice form defaults to TX. Correct it in Forms & Endorsements.")

    cov = {
        "carrier": PARTNERRE_CARRIER,
        "carrier_admitted": False,
        "am_best_rating": PARTNERRE_AM_BEST,
        "program_template": PARTNERRE_KEY,
        "policy_form": "Commercial Excess Liability Coverage Form CX 0001 04 13, following form over General Liability",
        "policy_period": _policy_period(ctx),
        "effective_date": ctx["effective"],
        "expiration_date": ctx["expiration"],
        "layer_description": f"{lim_m} xs Primary",
        "limits": [
            {"description": "Each Occurrence", "limit": _fmt(limit)},
            {"description": "Aggregate", "limit": _fmt(limit)},
            {"description": "Attachment", "limit": "Excess of the scheduled underlying insurance"},
        ],
        "underlying_insurance": underlying,
        "first_dollar_defense": False,
    }
    premium, taxes, total, tria = _premium_block(cov, inputs, PARTNERRE_TRIA_RATE)
    cov["summary_rows"] = [
        {"label": "Limits of Liability",
         "value": f"{_fmt(limit)} Each Occurrence / {_fmt(limit)} Aggregate, excess of scheduled underlying insurance"},
        {"label": "Follow Form", "value": f"General Liability ({lim_m} xs Primary)"},
        {"label": "Rating Basis", "value": "Flat / non-auditable"},
        {"label": "Minimum Earned Premium", "value": "25%"},
        {"label": "Terrorism (TRIA)",
         "value": (f"Optional, {_fmt(tria)} additional premium; must also be purchased on the underlying "
                   "policy to be eligible") if tria else
                  "Optional; must also be purchased on the underlying policy to be eligible"},
    ]
    forms = []
    for row in _PARTNERRE_FORMS:
        num, desc = row[0], row[1]
        tag = row[2] if len(row) > 2 else ""
        if tag == "state":
            num, desc = num.replace("{ST}", state), desc.replace("{ST}", state)
        if tag == "premises" and addresses:
            desc = f"{desc}: " + "; ".join(addresses)
        forms.append({"form_number": num, "description": desc, "template_form": True})
    cov["forms_endorsements"] = forms
    cov["terms_conditions"] = list(_PARTNERRE_TERMS)
    cov["subjectivities"] = list(_PARTNERRE_SUBJECTIVITIES)
    if not premium and not total:
        notes.append("REVIEW REQUIRED - PartnerRe umbrella template: no premium entered.")
    if not addresses:
        notes.append("REVIEW REQUIRED - PartnerRe umbrella template: no location addresses found for the "
                     "Designated Premises form (PRE ALT 0329). Add them in Forms & Endorsements.")
    cov["template_summary"] = {"limit": limit, "underlying_carrier": ul_carrier,
                               "underlying_rows": len(underlying), "premises": len(addresses), "state": state}
    return cov, notes


# ─── entry point ─────────────────────────────────────────────────────────────

def _fill_client_info(data, sov_data):
    ci = data.get("client_info")
    if not isinstance(ci, dict):
        ci = {}
        data["client_info"] = ci
    sov = _property_sov(sov_data) if sov_data else {}
    summary = (sov or {}).get("summary") if isinstance((sov or {}).get("summary"), dict) else {}
    if not summary and isinstance(sov_data, dict) and isinstance(sov_data.get("summary"), dict):
        summary = sov_data["summary"]
    if not ci.get("named_insured") and summary.get("named_insured"):
        ci["named_insured"] = summary["named_insured"]
    if not ci.get("address") and summary.get("mailing_address"):
        ci["address"] = summary["mailing_address"]
    eff = _mdy(ci.get("effective_date")) or _mdy(summary.get("effective_date"))
    if eff and not ci.get("effective_date"):
        ci["effective_date"] = eff
    if eff and not ci.get("expiration_date"):
        ci["expiration_date"] = _plus_one_year(eff)


def apply_program_templates(data, inputs, sov_data=None):
    """Build the selected template coverages into data["coverages"] in place.

    inputs = {
      "effective_date": "10/25/2026",                     # optional
      "property": {"enabled": true, "premium": 22000, "taxes_fees": 2192.5,
                   "total_premium": 24192.5, "tria_premium": "",  # blank -> 5% of premium
                   "include_bm": true | false | "auto",
                   "deductibles": {"aop": 10000, "water": 25000, "earth_movement": 25000,
                                   "wind": {"mode": "flat"|"pct", "amount": 25000, "pct": 2, "minimum": 25000},
                                   "named_storm": {"apply": "auto"|"always"|"never", "pct": 5, "minimum": 100000}}},
      "umbrella": {"enabled": true, "limit": 5000000, "premium": 8000, "taxes_fees": 660,
                   "total_premium": 8660, "tria_premium": "", "underlying_carrier": ""}
    }
    Returns a list of notes (also appended to data["warnings"])."""
    if not isinstance(data, dict):
        return []
    inputs = inputs if isinstance(inputs, dict) else {}
    covs = data.get("coverages")
    if not isinstance(covs, dict):
        covs = {}
        data["coverages"] = covs
    _fill_client_info(data, sov_data)
    notes = []
    eff = inputs.get("effective_date")

    p_in = inputs.get("property") if isinstance(inputs.get("property"), dict) else {}
    u_in = inputs.get("umbrella") if isinstance(inputs.get("umbrella"), dict) else {}

    if p_in.get("enabled"):
        prev = covs.get("property")
        if isinstance(prev, dict) and prev.get("carrier") and prev.get("program_template") != STARR_KEY:
            notes.append(f"Starr template replaced the extracted property quote ({prev.get('carrier')}).")
        cov, n = build_starr_property(data, dict(p_in, effective_date=eff or p_in.get("effective_date")), sov_data)
        covs["property"] = cov
        notes += n
    elif isinstance(covs.get("property"), dict) and covs["property"].get("program_template") == STARR_KEY:
        del covs["property"]  # template unticked on a re-apply

    if u_in.get("enabled"):
        prev = covs.get("umbrella")
        if isinstance(prev, dict) and prev.get("carrier") and prev.get("program_template") != PARTNERRE_KEY:
            notes.append(f"PartnerRe template replaced the extracted umbrella quote ({prev.get('carrier')}).")
        cov, n = build_partnerre_umbrella(data, dict(u_in, effective_date=eff or u_in.get("effective_date")), sov_data)
        covs["umbrella"] = cov
        notes += n
    elif isinstance(covs.get("umbrella"), dict) and covs["umbrella"].get("program_template") == PARTNERRE_KEY:
        del covs["umbrella"]

    data["program_templates"] = copy.deepcopy(inputs)
    # replace any earlier template notes, keep everything else
    ws = data.get("warnings")
    if not isinstance(ws, list):
        ws = [ws] if isinstance(ws, str) and ws.strip() else []
    ws = [w for w in ws if not (isinstance(w, str) and ("Starr template" in w or "PartnerRe template" in w
                                                         or "Starr property template" in w
                                                         or "PartnerRe umbrella template" in w))]
    data["warnings"] = ws + notes
    return notes
