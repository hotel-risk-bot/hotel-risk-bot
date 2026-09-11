"""
ACORD 125 FL (2016/03) generator from a HUB hospitality SOV (.xlsx / .csv).

    from acord125_from_sov import generate
    info = generate(sov_path, out_path, first_named_insured=None, blank_path=None)

`info` is a dict: out_name, pages, locations, rooms, tiv, named_insureds,
corrections (list of str), blanks (list of str), warnings (list of str).

Rules (per Stefan / acord-125-from-sov skill):
  * First Named Insured = corporate/parent entity from "Corporate Name (LLC)"
    (or the override passed in); the SOV "Named Insured" entity goes second.
  * > 4 locations: the form's own page 2 is duplicated for each further group
    of 4 and inserted directly after page 2 (premises block only).
  * A premises continuation schedule (all locations) goes at the end - no
    "ACORD 125" in the title, no footer.
  * Yellow Highlight annotations on the template are stripped.
  * Producer block is constant HUB International Florida / Stefan Burkey.
    No wholesale broker information anywhere.
  * Output name: ACORD 125 - <First NI> (<Second NI>) <MM-DD-YYYY eff>.pdf

CLI:  python3 acord125_from_sov.py <SOV.xlsx|csv> <out.pdf> [first named insured]
"""
import os
import re
import csv
import sys
import datetime
import collections

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_BLANK = os.path.join(HERE, "assets", "acord125_blank.pdf")

PRODUCER = {
    "Producer_FullName_A": "HUB International Florida",
    "Producer_MailingAddress_LineOne_A": "1560 Orange Ave, Suite 750",
    "Producer_MailingAddress_LineTwo_A": "Winter Park FL 32789",
    "Producer_ContactPerson_FullName_A": "Stefan Burkey",
    "Producer_ContactPerson_PhoneNumber_A": "407-636-8133",
    "Producer_ContactPerson_EmailAddress_A": "Stefan.Burkey@HUBInternational.com",
    "Producer_AuthorizedRepresentative_FullName_A": "Stefan Burkey",
}
SHORT_FLAG = {
    "InterContinental Hotels & Resorts": "IHG",
    "Hilton Worldwide": "Hilton",
    "Marriott International": "Marriott",
    "Wyndham Hotels & Resorts": "Wyndham",
    "Choice Hotels": "Choice",
    "Best Western Hotels & Resorts": "Best Western",
}
# Header aliases: canonical name -> alternates seen in HUB pipeline exports
ALIASES = {
    "Corporate Name (LLC)": ["Corporate Name", "Corporate Name (LLC)", "Corporate Entity", "Parent Entity"],
    "Franchise": ["Franchise", "Franchise(s)", "Franchisor", "Brand"],
    "County.": ["County.", "County"],
    "Full Address": ["Full Address", "Address", "Location Address", "Property Address"],
    "Mailing Address": ["Mailing Address", "Mailing Addr"],
    "Loc #": ["Loc #", "Loc", "Location #", "Location Number"],
    "Building #": ["Building #", "Bldg #", "Building"],
    "# of Rooms": ["# of Rooms", "Rooms", "Room Count"],
    "# of Floors": ["# of Floors", "Floors", "Stories"],
    "% Sprinklered": ["% Sprinklered", "Sprinklered %", "Sprinkler %"],
    "Application Contact Name": ["Application Contact Name", "Contact Name", "Application Contact"],
    "Application - Phone": ["Application - Phone", "Contact Phone", "Application Phone"],
    "Inspection Email": ["Inspection Email", "Contact Email", "Application Email"],
    "Property Owned / Managed": ["Property Owned / Managed", "Owned / Managed", "Owned/Managed"],
}
TEMPLATE_STALE = ["Garland Woods", "Sabeena", "33-1946451", "Country Clubs", "Smithfield"]


# ─── SOV loading ────────────────────────────────────────────────────────────

def _num(v):
    """Coerce a cell to a number when it looks like one (handles '$1,234', '12%', '1,234.5')."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s == "":
        return None
    t = s.replace("$", "").replace(",", "").replace(" ", "")
    pct = t.endswith("%")
    if pct:
        t = t[:-1]
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        f = float(t)
        if pct:
            f = f / 100.0 if f > 1 else f
        return int(f) if f.is_integer() else f
    return s


SOV_DATE = {"value": ""}


def _read_rows(path):
    """Return (headers, rows) from .xlsx/.xlsm/.csv. Header row is auto-detected
    (first row containing 'DBA' or 'Full Address' within the first 10 rows)."""
    ext = os.path.splitext(path)[1].lower()
    raw = []
    SOV_DATE["value"] = ""
    if ext in (".xlsx", ".xlsm", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        m = re.search(r"(\d{1,2})[_/-](\d{1,2})[_/-](\d{4})", ws.title or "")
        if m:
            SOV_DATE["value"] = "%s/%s/%s" % (m.group(1), m.group(2), m.group(3))
        for r in ws.iter_rows(values_only=True):
            raw.append(list(r))
        wb.close()
    elif ext in (".csv", ".txt"):
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
            sample = fh.read(4096)
            fh.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            except csv.Error:
                dialect = csv.excel
            for r in csv.reader(fh, dialect):
                raw.append([(_num(c) if c not in (None, "") else None) for c in r])
    else:
        raise ValueError("Unsupported file type %s - upload an .xlsx or .csv SOV" % ext)

    if not SOV_DATE["value"]:
        m = re.search(r"(\d{1,2})[_/-](\d{1,2})[_/-](\d{4})", os.path.basename(path))
        if m:
            SOV_DATE["value"] = "%s/%s/%s" % (m.group(1), m.group(2), m.group(3))
    hdr_i = 0
    for i, r in enumerate(raw[:10]):
        cells = [str(c).strip() for c in r if c is not None]
        if any(c in ("DBA", "Full Address", "Loc #") for c in cells):
            hdr_i = i
            break
    hdr = [str(c).strip() if c is not None else "" for c in raw[hdr_i]]
    rows = []
    for r in raw[hdr_i + 1:]:
        if not any(c not in (None, "") for c in r):
            continue
        d = {}
        for h, v in zip(hdr, r):
            if not h:
                continue
            if isinstance(v, str):
                v = v.strip()
                if v == "":
                    v = None
            d[h] = v
        rows.append(d)
    return hdr, rows


def _canon(d):
    """Apply header aliases so the rest of the code can use canonical keys."""
    out = dict(d)
    for canon, alts in ALIASES.items():
        if out.get(canon) not in (None, ""):
            continue
        for a in alts:
            if out.get(a) not in (None, ""):
                out[canon] = out[a]
                break
    return out


# ─── helpers ────────────────────────────────────────────────────────────────

def g(d, k, default=None):
    v = d.get(k)
    return default if v in (None, "") else v


def money(v):
    v = _num(v)
    return "" if v in (None, "", 0) or not isinstance(v, (int, float)) else "{:,}".format(int(round(v)))


def num(v):
    v = _num(v)
    return "" if v in (None, "") or not isinstance(v, (int, float)) else "{:,}".format(int(round(v)))


def nval(v, default=0):
    v = _num(v)
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else default


def parse_addr(full, fallback_city="", fallback_state="", fallback_zip=""):
    """'street, city, state, zip, country' -> (street, city, zip). Tolerant of
    fewer parts and of 'City, ST 12345' forms."""
    if not full:
        return "", fallback_city, fallback_zip
    p = [x.strip() for x in str(full).split(",") if x.strip()]
    street = p[0] if p else ""
    city = p[1] if len(p) > 1 else fallback_city
    zc = fallback_zip
    m = re.search(r"\b(\d{5})(?:-\d{4})?\b", str(full))
    if m:
        zc = m.group(1)
    # 'City, ST 12345' style -> city may be fine already; 'Tallahassee, Florida, 32303' handled by regex above
    return street, city, zc


def yes(v):
    return str(v or "").strip().lower().startswith("yes")


def entity_type(name):
    n = name.upper().replace(".", "")
    if "LLC" in n or "L L C" in n:
        return "LimitedLiabilityCorporation"
    if re.search(r"\bINC\b", n) or "CORP" in n:
        return "Corporation"
    if re.search(r"\bLP\b|\bLLP\b|PARTNERS", n):
        return "Partnership"
    return "Other"


def is_vacant(d):
    return "vacant" in str(g(d, "DBA", "")).lower() or "vacant" in str(g(d, "Building Name", "")).lower()


def is_office(d):
    return str(g(d, "Location Type", "")).lower() == "office" or "office" in str(g(d, "DBA", "")).lower()


def is_hotel(d):
    if is_vacant(d) or is_office(d):
        return False
    lt = str(g(d, "Location Type", "")).lower()
    return lt == "hotel" or (lt == "" and nval(g(d, "# of Rooms")) > 0)


def parse_eff(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.month, v.day, v.year
    s = str(v).split(" ")[0].strip()
    if "/" in s:
        parts = [int(x) for x in s.split("/")[:3]]
        m, dd, y = parts
    elif "-" in s:
        y, m, dd = [int(x) for x in s.split("-")[:3]]
    else:
        raise ValueError("Unrecognised Effective Date %r" % v)
    if y < 100:
        y += 2000
    return m, dd, y


def fmt_flag(d):
    return g(d, "Franchise") or g(d, "Franchise(s)") or ""


# ─── main ───────────────────────────────────────────────────────────────────

def generate(sov_path, out_path, first_named_insured=None, blank_path=None, today=None):
    blank = blank_path or DEFAULT_BLANK
    if not os.path.exists(blank):
        raise FileNotFoundError("ACORD 125 blank template not found at %s" % blank)
    TODAY = (today or datetime.date.today()).strftime("%m/%d/%Y")
    warnings, corrections, blanks = [], [], []

    hdr, rows = _read_rows(sov_path)
    locs = [_canon(r) for r in rows]
    locs = [d for d in locs if g(d, "Full Address") or g(d, "DBA")]
    if not locs:
        raise ValueError("No location rows found - the SOV needs a header row with 'DBA' / 'Full Address' and one row per location.")
    for i, d in enumerate(locs):
        if g(d, "Loc #") is None:
            d["Loc #"] = i + 1
            warnings.append("Loc # missing on row %d - assigned %d" % (i + 2, i + 1))
        d["Loc #"] = int(nval(d["Loc #"], i + 1))
        d["Building #"] = int(nval(g(d, "Building #", 1), 1))
    locs.sort(key=lambda d: (d["Loc #"], d["Building #"]))

    # --- county sanity: same city, different county inside the SOV -> majority wins
    by_city = collections.defaultdict(list)
    for d in locs:
        city = str(g(d, "City") or parse_addr(g(d, "Full Address"))[1]).strip().lower()
        if city and g(d, "County."):
            by_city[city].append(str(d["County."]).strip())
    for d in locs:
        city = str(g(d, "City") or parse_addr(g(d, "Full Address"))[1]).strip().lower()
        cs = by_city.get(city, [])
        if len(set(cs)) > 1:
            maj = collections.Counter(cs).most_common(1)[0][0]
            cur = str(g(d, "County.", "")).strip()
            if cur != maj and cs.count(maj) > cs.count(cur):
                corrections.append("Loc %s: county changed from %s to %s (other %s locations in the SOV are in %s County)"
                                   % (d["Loc #"], cur, maj, city.title(), maj))
                d["County."] = maj

    first = locs[0]
    # --- named insureds
    corp_raw = next((g(d, "Corporate Name (LLC)") for d in locs if g(d, "Corporate Name (LLC)")), None)
    if first_named_insured and str(first_named_insured).strip():
        corp_raw = str(first_named_insured).strip()
    corp_entities = [x.strip() for x in str(corp_raw or "").replace("\n", ";").split(";") if x.strip()]
    ni2 = str(g(first, "Named Insured", "") or "").replace("\n", " ").strip()
    if corp_entities:
        ni1 = next((e for e in corp_entities if e != ni2), corp_entities[0])
        named = [ni1] + ([ni2] if ni2 and ni2 != ni1 else [])
    else:
        if not ni2:
            raise ValueError("No named insured found - fill 'Corporate Name (LLC)' or 'Named Insured' in the SOV, or enter the First Named Insured on the form.")
        named = [ni2]
        blanks.append("First Named Insured: SOV 'Corporate Name (LLC)' is blank, so %s is shown as the only Named Insured - add the parent entity if there is one" % ni2)

    hotels = [d for d in locs if is_hotel(d)]
    total_rooms = int(sum(nval(g(d, "# of Rooms")) for d in hotels))
    tot = lambda k: sum(nval(g(d, k)) for d in locs)
    rest_locs = [str(d["Loc #"]) for d in locs if yes(g(d, "Restaurant"))]
    liq = ["Loc %s $%s (%s%%)" % (d["Loc #"], money(d["Liquor Sales"]), g(d, "Liquor %", ""))
           for d in locs if nval(g(d, "Liquor Sales")) > 0]
    gas_locs = [str(d["Loc #"]) for d in locs if yes(g(d, "Gas or Tanks Present"))]

    def locs_str(ids):
        ids = list(ids)
        if len(ids) == 1:
            return "Loc %s" % ids[0]
        return "Locs %s and %s" % (", ".join(ids[:-1]), ids[-1]) if len(ids) > 1 else ""

    def uniform(col, *answers):
        vals = [str(g(d, col, "")).strip().lower() for d in locs if not is_vacant(d)]
        vals = [v for v in vals if v]
        return bool(vals) and all(any(v.startswith(a) for a in answers) for v in vals)

    def by(k):
        groups = collections.OrderedDict()
        for d in locs:
            v = g(d, k)
            if v:
                groups.setdefault(str(v), []).append(str(d["Loc #"]))
        return ", ".join("%s (%s)" % (v, "Loc %s" % l[0] if len(l) == 1 else "Locs %s" % ", ".join(l)) for v, l in groups.items())

    gl_code = "45191" if any(nval(g(d, "# of Floors")) >= 4 for d in hotels) else "45190"
    try:
        m, dd, y = parse_eff(g(first, "Effective Date"))
    except Exception:
        raise ValueError("Effective Date is missing or unreadable in the SOV (expected M/D/YYYY).")
    eff_s, exp_s, prior_eff = ("%02d/%02d/%d" % (m, dd, y), "%02d/%02d/%d" % (m, dd, y + 1), "%02d/%02d/%d" % (m, dd, y - 1))
    eff_file = "%02d-%02d-%d" % (m, dd, y)
    mail_state = g(first, "State", "")
    mail1, mailcity, mailzip = parse_addr(g(first, "Mailing Address"), "", mail_state, "")
    if not g(first, "Mailing Address"):
        blanks.append("Mailing address (SOV 'Mailing Address' is blank)")
    policies = str(g(first, "Policies", ""))
    _pol = []
    for d in locs:
        for p in str(g(d, "Policies", "")).split(","):
            if p.strip() and p.strip() not in _pol:
                _pol.append(p.strip())
    all_policies = ", ".join(_pol)
    umb = nval(g(first, "Umbrella Limit"))
    state_name = {"FL": "Florida", "GA": "Georgia", "AL": "Alabama", "SC": "South Carolina", "NC": "North Carolina",
                  "TN": "Tennessee", "TX": "Texas", "LA": "Louisiana", "MS": "Mississippi", "VA": "Virginia"}.get(mail_state, mail_state)

    def display_name(d, short):
        fl = fmt_flag(d)
        flag = SHORT_FLAG.get(fl, fl) if short else fl
        dba = str(g(d, "DBA", ""))
        if not flag or flag.lower() in dba.lower() or dba.lower() in flag.lower():
            return dba
        return "%s (%s)" % (dba, flag)

    def ops_short(d):
        if is_vacant(d):
            acres = g(d, "Vacant Land (Acres)")
            return "Vacant building%s - no operations" % (" and %s acres vacant land" % acres if acres else "")
        if is_office(d):
            return "Corporate office for hotel operations"
        rest = "restaurant/lounge, " if yes(g(d, "Restaurant")) else ""
        pool = "pool, " if nval(g(d, "# of Pools")) > 0 else ""
        fit = "fitness" if g(d, "Fitness Room") and str(g(d, "Fitness Room")).lower() not in ("no", "none") else ""
        s = "%s - %s, %s rooms, %s sty, %s%s%s" % (display_name(d, True), str(g(d, "Hotel Class", "hotel")).lower(),
                                                    num(g(d, "# of Rooms")) or "?", num(g(d, "# of Floors")) or "?", rest, pool, fit)
        return s.rstrip(", ")

    def ops_long(d):
        if is_vacant(d):
            acres = g(d, "Vacant Land (Acres)")
            sq = g(d, "LRO (Sqft)") or g(d, "SqFt")
            return "Vacant building%s%s - no operations" % (" (%s sq ft)" % num(sq) if sq else "",
                                                             " and %s acres vacant land" % acres if acres else "")
        if is_office(d):
            return "Corporate office for hotel operations (%s sq ft, %s story)" % (num(g(d, "SqFt") or g(d, "LRO (Sqft)")), num(g(d, "# of Floors")) or "1")
        feats = [x for x, ok in [("owned restaurant/lounge", yes(g(d, "Restaurant"))),
                                 ("outdoor pool", nval(g(d, "# of Pools")) > 0),
                                 ("fitness room", g(d, "Fitness Room") and str(g(d, "Fitness Room")).lower() not in ("no", "none"))] if ok]
        return "%s - %s hotel, %s rooms, %s stories%s" % (display_name(d, False), str(g(d, "Hotel Class", "")).lower(),
                                                         num(g(d, "# of Rooms")) or "?", num(g(d, "# of Floors")) or "?",
                                                         ("; " + ", ".join(feats)) if feats else "")

    def premises_values(group):
        V = {}
        for d, s in zip(group, "ABCD"):
            a1, city, zc = parse_addr(g(d, "Full Address"), g(d, "City", ""), g(d, "State", ""), "")
            city = g(d, "City") or city
            sq = g(d, "SqFt") or g(d, "LRO (Sqft)")
            owned = str(g(d, "Property Owned / Managed", "Owned")).lower().startswith("own")
            V.update({
                "CommercialStructure_Location_ProducerIdentifier_%s" % s: str(d["Loc #"]),
                "CommercialStructure_Building_ProducerIdentifier_%s" % s: str(d["Building #"]),
                "CommercialStructure_PhysicalAddress_LineOne_%s" % s: a1,
                "CommercialStructure_PhysicalAddress_CityName_%s" % s: city,
                "CommercialStructure_PhysicalAddress_CountyName_%s" % s: str(g(d, "County.", "")),
                "CommercialStructure_PhysicalAddress_StateOrProvinceCode_%s" % s: str(g(d, "State", mail_state)),
                "CommercialStructure_PhysicalAddress_PostalCode_%s" % s: zc,
                "CommercialStructure_InsuredInterest_OwnerIndicator_%s" % s: owned,
                "CommercialStructure_InsuredInterest_TenantIndicator_%s" % s: not owned,
                "CommercialStructure_AnnualRevenueAmount_%s" % s: money(g(d, "Total Sales")),
                "BuildingOccupancy_OccupiedArea_%s" % s: num(sq),
                "Construction_BuildingArea_%s" % s: num(sq),
                "BuildingOccupancy_OperationsDescription_%s" % s: ops_short(d),
                "CommercialStructure_Question_ABBCode_%s" % s: "N",
            })
        return V

    V = dict(PRODUCER)
    V["Form_CompletionDate_A"] = TODAY
    V["Policy_Status_QuoteIndicator_A"] = True
    V["Policy_LineOfBusiness_CommercialProperty_A"] = "Property" in all_policies
    V["Policy_LineOfBusiness_CommercialGeneralLiability_A"] = "Liability" in all_policies
    V["Policy_LineOfBusiness_UmbrellaIndicator_A"] = "Umbrella" in all_policies
    V["Policy_LineOfBusiness_BusinessAutoIndicator_A"] = "Auto" in all_policies
    V["CommercialPolicy_Attachment_HotelMotelSupplementIndicator_A"] = True
    V["CommercialPolicy_Attachment_RestaurantTavernSupplementIndicator_A"] = bool(rest_locs)
    V["CommercialPolicy_Attachment_StatementOfValuesIndicator_A"] = True
    V["CommercialPolicy_Attachment_AdditionalPremisesScheduleIndicator_A"] = len(locs) > 4
    V["CommercialPolicy_Attachment_VacantBuildingSupplementIndicator_A"] = any(is_vacant(d) for d in locs)
    V["CommercialPolicy_Attachment_LossSummaryIndicator_A"] = True
    V["Policy_EffectiveDate_A"], V["Policy_ExpirationDate_A"] = eff_s, exp_s
    V["Policy_Payment_DirectBillIndicator_A"] = True
    for name, s in zip(named, "ABC"):
        V.update({"NamedInsured_FullName_%s" % s: name,
                  "NamedInsured_MailingAddress_LineOne_%s" % s: mail1,
                  "NamedInsured_MailingAddress_CityName_%s" % s: mailcity,
                  "NamedInsured_MailingAddress_StateOrProvinceCode_%s" % s: mail_state,
                  "NamedInsured_MailingAddress_PostalCode_%s" % s: mailzip,
                  "NamedInsured_GeneralLiabilityCode_%s" % s: gl_code,
                  "NamedInsured_SICCode_%s" % s: "7011",
                  "NamedInsured_NAICSCode_%s" % s: "721110",
                  "NamedInsured_LegalEntity_%sIndicator_%s" % (entity_type(name), s): True})
    fein = next((g(d, "FEIN") for d in locs if g(d, "FEIN")), None)
    if fein:
        V["NamedInsured_TaxIdentifier_A"] = str(fein)
    else:
        blanks.append("FEIN")
    contact_name = g(first, "Application Contact Name", "")
    contact_phone = g(first, "Application - Phone", "")
    contact_email = g(first, "Inspection Email", "")
    V["NamedInsured_Contact_ContactDescription_A"] = "Inspection"
    V["NamedInsured_Contact_FullName_A"] = str(contact_name or "")
    V["NamedInsured_Contact_PrimaryPhoneNumber_A"] = str(contact_phone or "")
    V["NamedInsured_Contact_PrimaryEmailAddress_A"] = str(contact_email or "")
    V["NamedInsured_Contact_PrimaryBusinessPhoneIndicator_A"] = True
    if not (contact_name and contact_phone and contact_email):
        blanks.append("Inspection contact name / phone / email")
    V.update(premises_values(locs[:4]))
    V["BusinessInformation_BusinessType_RestaurantIndicator_A"] = bool(rest_locs)
    V["BusinessInformation_BusinessType_OtherIndicator_A"] = True
    V["BusinessInformation_BusinessType_OtherDescription_A"] = "Hotels / Hospitality"

    flags = sorted({fmt_flag(d) for d in hotels if fmt_flag(d)})
    others = []
    office_locs = [str(d["Loc #"]) for d in locs if is_office(d) and not is_vacant(d)]
    vacant_locs = [str(d["Loc #"]) for d in locs if is_vacant(d)]
    if office_locs:
        others.append("a corporate office (Loc %s)" % ", ".join(office_locs))
    if vacant_locs:
        others.append("%s vacant building/land parcel%s (Loc %s)" % ("one" if len(vacant_locs) == 1 else len(vacant_locs), "" if len(vacant_locs) == 1 else "s", ", ".join(vacant_locs)))
    all_sprk = bool(hotels) and all(nval(g(d, "% Sprinklered")) >= 1 for d in hotels)
    _lb = collections.OrderedDict()
    for d in locs:
        if yes(g(d, "Restaurant")) and nval(g(d, "Liquor %")) > 0:
            pct = nval(g(d, "Liquor %"))
            b = "under 30%" if pct < 30 else ("30-75%" if pct <= 75 else "over 75%")
            _lb.setdefault(b, []).append(str(d["Loc #"]))
    liq_note = " (liquor %s)" % "; ".join("%s of receipts at %s" % (b, locs_str(l)) for b, l in _lb.items()) if _lb else ""
    any_pool = any(nval(g(d, "# of Pools")) > 0 for d in hotels)
    V["CommercialPolicy_OperationsDescription_A"] = (
        "Owner/operator of %d franchised hotel%s in %s (%s guest rooms total)%s%s. "
        % (len(hotels), "" if len(hotels) == 1 else "s", state_name, num(total_rooms),
           " under %s flags" % ", ".join(flags) if flags else "",
           ", plus " + " and ".join(others) if others else "")
        + ("Full service hotels at Locs %s operate owned restaurant/lounge with table service%s. " % (", ".join(rest_locs), liq_note) if rest_locs else "")
        + "All hotels: " + ("outdoor pool (guest only, fenced/self-closing gate, no diving board/slide, VGB compliant), " if any_pool else "")
        + "fitness room, guest laundry, " + ("100% sprinklered, " if all_sprk else "")
        + "central station monitored fire alarm, 24 hr employee on duty, daily room rentals only (no hourly, no >30 day stays). "
        "No owned autos, shuttle, valet or guest transportation. All hotel operations >3 years. "
        + ("Locations 5-%d on the following premises page; " % len(locs) if len(locs) > 4 else "")
        + "see premises continuation schedule and SOV for all %d locations." % len(locs)
    )
    V["CommercialPolicy_Question_KAACode_A"] = "Y"
    V["CommercialPolicy_FormalSafetyProgram_OtherIndicator_B"] = True
    V["CommercialPolicy_FormalSafetyProgram_OtherDescription_B"] = "HT awareness, TIPS, background checks, pest control"
    if gas_locs:
        V["CommercialPolicy_Question_ABCCode_A"] = "Y"
        V["CommercialPolicy_AnyExposureToFlammableExplosivesChemicalsExplanation_A"] = (
            "Natural gas/propane for commercial cooking at Locs %s (NFPA 96 / UL 300 compliant suppression); "
            "pool chemicals stored and managed on a regular schedule at all hotel locations." % ", ".join(gas_locs))
    else:
        V["CommercialPolicy_Question_ABCCode_A"] = "N"
    V["CommercialPolicy_Question_AADCode_A"] = "N"
    V["CommercialPolicy_Question_ABBCode_A"] = "N"

    liab_only = "".join("Loc %s (%s) - %s only. " % (d["Loc #"], "vacant building/land" if is_vacant(d) else g(d, "DBA", ""),
                                                     " and ".join(x.strip().lower() for x in str(g(d, "Policies", "")).split(",") if x.strip()))
                        for d in locs if "Property" not in str(g(d, "Policies", "")))
    cov = all_policies.replace("Liability", "General Liability").replace("Auto", "Hired & Non-Owned Auto (no owned autos)")
    if umb:
        cov = cov.replace("Umbrella", "Umbrella $%s" % money(umb))
    fl_req, eq_req = yes(g(first, "Flood Coverage Required?")), yes(g(first, "Earthquake Required?"))
    fl_lim, eq_lim = money(g(first, "Flood Limit Requested")), money(g(first, "Earthquake Limit Requested"))
    if fl_req and eq_req and fl_lim == eq_lim and fl_lim:
        cat_note = "Flood and Earthquake requested at $%s each. " % fl_lim
    else:
        cat_note = ("Flood requested%s. " % (" at $" + fl_lim if fl_lim else "") if fl_req else "") + \
                   ("Earthquake requested%s. " % (" at $" + eq_lim if eq_lim else "") if eq_req else "")
    attrs = []
    if all_sprk:
        attrs.append("100% sprinklered")
    if uniform("Wiring Type", "copper"):
        attrs.append("copper wiring")
    if all(g(d, "No Federal Pacific, Zinsco, Stablok, or Challenger") in (True, "TRUE", "True", "Yes", "yes", 1) for d in locs if not is_vacant(d)):
        attrs.append("no FPE/Zinsco/Stab-Lok/Challenger panels")
    if uniform("Exterior Insulating Finishing System (EIFS)", "no", "none", "0"):
        attrs.append("no EIFS")
    if uniform("Solar Panels (Roof)", "no"):
        attrs.append("no solar")
    if uniform("Basement", "no"):
        attrs.append("no basements")
    fp, bal = uniform("Fireplace in Rooms?", "no"), uniform("Balconies - Guestroom", "no")
    if fp and bal:
        attrs.append("no guestroom fireplaces or balconies")
    elif fp:
        attrs.append("no guestroom fireplaces")
    elif bal:
        attrs.append("no guestroom balconies")
    attr_note = ("All buildings " + ", ".join(attrs) + ". ") if attrs else ""
    bb_note = " (no bed bug claims in 5 yrs)" if uniform("Any past bed bug claims (5 years)?", "no") else ""
    tail = []
    if uniform("Any known construction for the upcoming year", "no"):
        tail.append("construction")
    if uniform("Any known sponsored events for upcoming year", "no"):
        tail.append("sponsored events")
    tail_note = (" No known %s for the upcoming year." % " or ".join(tail)) if tail else ""
    sov_dated = (" dated %s" % SOV_DATE["value"]) if SOV_DATE["value"] else ""
    V["CommercialPolicy_RemarkText_A"] = (
        "PORTFOLIO SUMMARY (per SOV%s): %d locations, %s rooms. Total TIV $%s (Building $%s; Contents $%s; Business Income $%s). "
        "Total annual sales $%s. " % (sov_dated, len(locs), num(total_rooms), money(tot("TIV")), money(tot("Building Limit")),
                                       money(tot("Contents Limit")), money(tot("Business Income Limit")), money(tot("Total Sales")))
        + ("Liquor sales: %s. " % "; ".join(liq) if liq else "")
        + "Coverage requested: %s. " % cov
        + cat_note
        + liab_only
        + attr_note
        + "Construction: %s. Coastal tiers: %s. " % (by("Construction") or "see SOV", by("Coastal Exposure") or "see SOV")
        + "Underwriting subjectivities confirmed: >3 yrs operational experience; no prior A&M, A&B or human trafficking losses; "
        + "no liability losses over $100,000; pest control contract incl. bed bug prevention%s; background checks on all employees; " % bb_note
        + "no homeless shelter, hourly or long-term (>30 day) rentals; annual human trafficking awareness training; TIPS-trained "
        "servers with written intoxicated-patron guidelines; HNOA controls (<5 employees, <20 trips/month, no one under 21, "
        "employee auto insurance confirmed)." + tail_note
    )
    V["PriorCoverage_PolicyYear_A"] = str(y - 1)
    for key, col in [("GeneralLiability", "Expiring Liability Carrier"), ("Property", "Expiring Property Carrier"),
                     ("OtherLine", "Expiring Umbrella Carrier")]:
        if g(first, col):
            V["PriorCoverage_%s_InsurerFullName_A" % key] = str(first[col])
            V["PriorCoverage_%s_EffectiveDate_A" % key] = prior_eff
            V["PriorCoverage_%s_ExpirationDate_A" % key] = eff_s
    V["PriorCoverage_OtherLine_LineOfBusinessCode_A"] = "Umbrella" if g(first, "Expiring Umbrella Carrier") else ""
    V["LossHistory_InformationYearCount_A"] = "5"
    exp_p, exp_l, exp_u = g(first, "Expiring Property Carrier", "n/a"), g(first, "Expiring Liability Carrier", "n/a"), g(first, "Expiring Umbrella Carrier", "n/a")
    if exp_p == exp_l:
        expiring = "%s (Property and General Liability), %s (Umbrella%s)" % (exp_p, exp_u, " $%s" % money(umb) if umb else "")
    else:
        expiring = "%s (Property), %s (General Liability), %s (Umbrella%s)" % (exp_p, exp_l, exp_u, " $%s" % money(umb) if umb else "")
    V["CommercialPolicy_RemarkText_B"] = (
        "New business to HUB International, effective %s. Expiring program: %s. "
        "Five-year currently valued loss runs for all lines to follow; loss history section to be completed from carrier loss runs. "
        "Attachments: Statement of Values (SOV), premises schedule (all %d locations), Hotel/Motel Supplement%s%s, ACORD 140 (Property), "
        "ACORD 126 (GL), ACORD 131 (Umbrella)." % (eff_s, expiring, len(locs),
                                                   ", Restaurant/Tavern Supplement (Locs %s)" % ", ".join(rest_locs) if rest_locs else "",
                                                   ", Vacant Building Supplement (%s)" % locs_str(vacant_locs) if vacant_locs else "")
    )

    # ─── fill ───────────────────────────────────────────────────────────
    def fill(values, rename=None):
        doc = pymupdf.open(blank)
        seen = set()
        for page in doc:
            for a in list(page.annots(types=[pymupdf.PDF_ANNOT_HIGHLIGHT])):
                page.delete_annot(a)
            for w in page.widgets():
                n = w.field_name
                seen.add(n)
                if w.field_type == pymupdf.PDF_WIDGET_TYPE_CHECKBOX:
                    w.field_value = bool(values.get(n, False))
                elif w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT:
                    if n == "Form_EditionIdentifier_A":
                        continue
                    val = str(values.get(n, "") or "")
                    if val == "":
                        doc.xref_set_key(w.xref, "V", "()")
                        doc.xref_set_key(w.xref, "AP", "null")
                    if n.startswith("BuildingOccupancy_OperationsDescription_"):
                        w.text_fontsize = 7
                    w.field_value = val
                else:
                    continue
                w.update()
                if rename:
                    doc.xref_set_key(w.xref, "T", "(%s%s)" % (n, rename))
        return doc, seen

    doc, seen = fill(V)
    missing = [k for k in V if k not in seen]
    if missing:
        warnings.append("Template is missing %d expected field(s): %s" % (len(missing), ", ".join(missing[:8])))

    insert_at = 2
    for i in range(4, len(locs), 4):
        grp = locs[i:i + 4]
        doc2, _ = fill(premises_values(grp), rename="_L%s" % grp[0]["Loc #"])
        doc.insert_pdf(doc2, from_page=1, to_page=1, start_at=insert_at)
        insert_at += 1
        doc2.close()

    # ─── continuation schedule (all locations) - no ACORD title, no footer
    page = doc.new_page(width=612, height=792)
    y0 = 40
    page.insert_text((22, y0), "PREMISES INFORMATION CONTINUATION (ADDITIONAL PREMISES SCHEDULE)", fontsize=9, fontname="helv")
    page.insert_text((22, y0 + 13), "Applicant: %s     Proposed effective date: %s     Date: %s" % (" / ".join(named), eff_s, TODAY), fontsize=8, fontname="helv")
    page.insert_text((22, y0 + 24), "Source: client Statement of Values. All premises owned by the applicant unless noted.", fontsize=7.5, fontname="helv")
    cols = [("LOC #", 22), ("BLDG #", 46), ("STREET, CITY, STATE, ZIP, COUNTY", 76), ("INTEREST", 296), ("ROOMS", 338),
            ("ANNUAL REVENUES", 368), ("SQ FT OCC. / TOTAL", 436), ("TIV", 520)]
    yy = y0 + 44

    def wrap(text, size, width=510):
        words, lines, cur = str(text).split(" "), [], ""
        for w in words:
            t = (cur + " " + w).strip()
            if pymupdf.get_text_length(t, fontname="helv", fontsize=size) > width and cur:
                lines.append(cur)
                cur = w
            else:
                cur = t
        if cur:
            lines.append(cur)
        return lines or [""]

    def header_row(pg, yv):
        pg.draw_rect(pymupdf.Rect(20, yv - 10, 592, yv + 3), color=(0, 0, 0), fill=(0.88, 0.88, 0.88), width=0.5)
        for t, x in cols:
            pg.insert_text((x, yv), t, fontsize=6.5, fontname="helv")

    header_row(page, yy)
    yy += 12
    for d in locs:
        if yy > 670:
            page = doc.new_page(width=612, height=792)
            yy = 40
            header_row(page, yy)
            yy += 12
        a1, city, zc = parse_addr(g(d, "Full Address"), g(d, "City", ""), g(d, "State", ""), "")
        city = g(d, "City") or city
        sq = g(d, "SqFt") or g(d, "LRO (Sqft)")
        owned = str(g(d, "Property Owned / Managed", "Owned")).lower().startswith("own")
        has_prop = "Property" in str(g(d, "Policies", "")) and nval(g(d, "TIV")) > 0
        for x, t in [(22, str(d["Loc #"])), (46, str(d["Building #"])),
                     (76, "%s, %s, %s %s (%s Co.)" % (a1, city, g(d, "State", mail_state), zc, g(d, "County.", "?"))),
                     (296, "Owner" if owned else "Tenant"), (338, num(g(d, "# of Rooms")) or "-"),
                     (368, "$" + money(g(d, "Total Sales")) if nval(g(d, "Total Sales")) > 0 else "-"),
                     (436, "%s / %s" % (num(sq), num(sq)) if sq else "-"),
                     (520, "$" + money(d["TIV"]) if has_prop else "Liability only")]:
            page.insert_text((x, yy), t, fontsize=7, fontname="helv")
        yy += 10
        for ln in wrap("Description of operations: " + ops_long(d), 6.5):
            page.insert_text((76, yy), ln, fontsize=6.5, fontname="helv")
            yy += 8.5
        yy += 0.5
        extra = []
        if has_prop and nval(g(d, "Building Limit")) > 0:
            extra.append("Bldg $%s; Contents $%s; BI $%s%s" % (money(d["Building Limit"]), money(g(d, "Contents Limit")),
                                                              money(g(d, "Business Income Limit")),
                                                              "; Pool $%s" % money(d["Pool Limit"]) if nval(g(d, "Pool Limit")) > 0 else ""))
            extra.append("Yr built %s; %s; ISO %s; %d%% sprinklered; roof %s" % (
                g(d, "Yr Built", "?"), g(d, "Construction", "?"), g(d, "ISO #", "?"),
                int(round(nval(g(d, "% Sprinklered")) * (100 if nval(g(d, "% Sprinklered")) <= 1 else 1))),
                g(d, "Year Roof Fully Replaced", "?")))
        extra += ["GL class: %s" % g(d, "Liability Class Codes", "see SOV"), "Any area leased to others: N"]
        for e in extra:
            for ln in wrap(e, 6.5):
                page.insert_text((76, yy), ln, fontsize=6.5, fontname="helv", color=(0.25, 0.25, 0.25))
                yy += 8.5
        page.draw_line((20, yy - 4), (592, yy - 4), color=(0.6, 0.6, 0.6), width=0.3)
        yy += 6
    page.insert_text((22, yy + 6), "TOTALS: %d locations | %s rooms | Annual revenues $%s | TIV $%s" % (
        len(locs), num(total_rooms), money(tot("Total Sales")), money(tot("TIV"))), fontsize=7.5, fontname="helv")

    doc.save(out_path, garbage=3, deflate=True)
    n_pages = len(doc)
    doc.close()

    # ─── stale-text check
    chk = pymupdf.open(out_path)
    text = "\n".join(p.get_text() for p in chk)
    for p in chk:
        for w in p.widgets():
            if w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT and w.field_value:
                text += "\n" + str(w.field_value)
    chk.close()
    stale = [s for s in TEMPLATE_STALE if s.lower() in text.lower()]
    if stale:
        warnings.append("Stale template strings still present: %s" % ", ".join(stale))

    blanks += ["Employee counts (full/part time) and business start date",
               "Prior policy numbers and premiums (carriers/terms are filled)",
               "Loss history (5 yrs - to be completed from loss runs)",
               "Mortgagees / additional interests",
               "General information questions 1, 4, 5, 7-10, 12-15 (not in the SOV)",
               "Producer license / NPN"]
    safe = lambda s: re.sub(r"[\\/:*?\"<>|]+", "-", s).strip()
    out_name = "ACORD 125 - %s%s %s.pdf" % (safe(named[0]), " (%s)" % safe(named[1]) if len(named) > 1 else "", eff_file)
    return {
        "out_name": out_name, "pages": n_pages, "locations": len(locs), "hotels": len(hotels),
        "rooms": total_rooms, "tiv": int(tot("TIV")), "sales": int(tot("Total Sales")),
        "named_insureds": named, "effective": eff_s, "expiration": exp_s,
        "lines": all_policies, "restaurant_locs": rest_locs, "vacant_locs": vacant_locs, "office_locs": office_locs,
        "corrections": corrections, "blanks": blanks, "warnings": warnings,
    }


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    info = generate(sys.argv[1], sys.argv[2], first_named_insured=(sys.argv[3] if len(sys.argv) > 3 else None))
    import json
    print(json.dumps(info, indent=2))
