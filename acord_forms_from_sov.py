"""
ACORD 126 (General Liability Section), ACORD 130 (Workers Compensation
Application) and ACORD 140 (Property Section) generators from a HUB
hospitality SOV / Pipeline export (.xlsx / .csv).

    from acord_forms_from_sov import generate_form, FORMS
    info = generate_form("140", sov_path, out_path, first_named_insured=None)

`info` is a dict like the ACORD 125 report: out_name, pages, locations,
named_insureds, effective, expiration, corrections, blanks, warnings, plus
form-specific counts.

Shared rules (same as acord125_from_sov):
  * Producer block is constant HUB International Midwest Limited, 203 N LaSalle,
    Suite 2000, Chicago, IL 60601 / Stefan Burkey. No wholesale broker names.
  * First Named Insured: the override typed on the page; else the SOV "Client
    Name" when the locations belong to several corporate entities; else the
    corporate/parent entity of the first row (the ACORD 125 rule). Every other
    entity is listed on the form's remarks / continuation schedule.
  * Blanks stay blank (the PDF stays editable) and are reported back.
  * Yellow Highlight annotations on the templates are stripped.
  * Templates: assets/acord126_blank.pdf, assets/acord140_blank.pdf and
    assets/acord130_blank.pdf (standard edition). Until the standard ACORD 130
    blank is dropped into assets/, the Florida edition assets/acord130fl_blank.pdf
    is used and the report says so.

ACORD 140: one premises page per building row with values; page 1 carries the
first, page 2 the second, and page 2 is duplicated for every further building
(fields renamed _L<loc>B<bldg>). Portfolio remarks on page 2.
ACORD 126: limits per the HUB hotel program, schedule of hazards (first three
rows on the form, every location on a continuation schedule), hotel answers to
the general information questions.
ACORD 130: locations and class codes (first three / five on the form, the rest
on continuation schedules), employer's liability 1M/1M/1M, hotel answers.

CLI:  python3 acord_forms_from_sov.py <126|130|140> <SOV.xlsx|csv> <out.pdf> [first named insured]
"""
import os
import re
import sys
import json
import datetime
import collections

import pymupdf

import acord125_from_sov as A
from acord125_from_sov import g, money, num, nval, parse_addr, yes, entity_type, is_vacant, is_office, is_hotel, parse_eff, fmt_flag

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
BLANKS = {
    "126": [os.path.join(ASSETS, "acord126_blank.pdf")],
    "140": [os.path.join(ASSETS, "acord140_blank.pdf")],
    "130": [os.path.join(ASSETS, "acord130_blank.pdf"), os.path.join(ASSETS, "acord130fl_blank.pdf")],
}
FORMS = collections.OrderedDict([
    ("125", "ACORD 125 - Commercial Insurance Application"),
    ("126", "ACORD 126 - Commercial General Liability Section"),
    ("140", "ACORD 140 - Property Section"),
    ("130", "ACORD 130 - Workers Compensation Application"),
])
PRODUCER_NAME = A.PRODUCER["Producer_FullName_A"]
GL_LIMITS = {  # HUB hotel program standard limits (per Stefan's filled ACORD 126)
    "GeneralLiability_GeneralAggregate_LimitAmount_A": "2,000,000",
    "GeneralLiability_ProductsAndCompletedOperations_AggregateLimitAmount_A": "1,000,000",
    "GeneralLiability_PersonalAndAdvertisingInjury_LimitAmount_A": "1,000,000",
    "GeneralLiability_EachOccurrence_LimitAmount_A": "1,000,000",
    "GeneralLiability_FireDamageRentedPremises_EachOccurrenceLimitAmount_A": "100,000",
    "GeneralLiability_MedicalExpense_EachPersonLimitAmount_A": "5,000",
    "GeneralLiability_OtherCoverageLimitDescription_A": "Hired & Non-Owned Auto",
    "GeneralLiability_OtherCoverageLimitAmount_A": "1,000,000",
}
EL_LIMIT = "1,000,000"
# compact output: object streams and deflated fonts keep a 17-page ACORD 140 under 500 KB
SAVE_KW = dict(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=1)
WC_CLASSES = [  # (SOV payroll column, NCCI class, description)
    ("Hotel Payroll", "9052", "Hotel NOC - all other employees & salespersons, drivers"),
    ("Restaurant Payroll", "9058", "Hotel - restaurant employees"),
    ("Clerical Payroll", "8810", "Clerical office employees NOC"),
    ("Sales Payroll", "8742", "Salespersons or collectors - outside"),
    ("Other Payroll", "", "Other (see remarks)"),
]
STALE = A.TEMPLATE_STALE + ["Triton Hospitality", "Sima Patel", "Ritesh Patel", "Air2 Homes", "Sai Shankar",
                            "HUB International Florida", "Mercy Dr", "Gateway Corporate", "Forklift", "11539 International"]


# ─── shared SOV context ─────────────────────────────────────────────────────

def _entity_like(s):
    return bool(re.search(r"\b(LLC|L\.L\.C|INC|CORP|LP|LLP|LLLP|LTD|COMPANY|CO\.|PARTNERS|TRUST|HOLDINGS|MANAGEMENT|GROUP)\b",
                          str(s or "").upper().replace(".", ". ")))


def load(sov_path, first_named_insured=None, today=None):
    """Read the SOV once and derive everything the three forms share."""
    TODAY = (today or datetime.date.today()).strftime("%m/%d/%Y")
    warnings, corrections, blanks = [], [], []
    hdr, rows = A._read_rows(sov_path)
    locs = [A._canon(r) for r in rows]
    locs = [d for d in locs if g(d, "Full Address") or g(d, "DBA")]
    if not locs:
        raise ValueError("No location rows found - the SOV needs a header row with 'DBA' / 'Full Address' and one row per location.")
    for i, d in enumerate(locs):
        if g(d, "Loc #") is None:
            d["Loc #"] = i + 1
            warnings.append("Loc # missing on row %d - assigned %d" % (i + 2, i + 1))
        d["Loc #"] = int(nval(d["Loc #"], i + 1))
        d["Building #"] = int(nval(g(d, "Building #", 1), 1))
        if not g(d, "City"):
            d["City"] = parse_addr(g(d, "Full Address"))[1]
    locs.sort(key=lambda d: (d["Loc #"], d["Building #"]))

    # county sanity (same rule as the ACORD 125)
    by_city = collections.defaultdict(list)
    for d in locs:
        city = str(g(d, "City", "")).strip().lower()
        if city and g(d, "County."):
            by_city[city].append(str(d["County."]).strip())
    for d in locs:
        city = str(g(d, "City", "")).strip().lower()
        cs = by_city.get(city, [])
        if len(set(cs)) > 1:
            maj = collections.Counter(cs).most_common(1)[0][0]
            cur = str(g(d, "County.", "")).strip()
            if cur != maj and cs.count(maj) > cs.count(cur):
                corrections.append("Loc %s: county changed from %s to %s (other %s locations in the SOV are in %s County)"
                                   % (d["Loc #"], cur, maj, city.title(), maj))
                d["County."] = maj

    first = locs[0]
    # one row per location (building 1 carries the location-level data)
    seen, loc_rows = set(), []
    for d in locs:
        if d["Loc #"] not in seen:
            seen.add(d["Loc #"])
            loc_rows.append(d)

    # named insureds
    corp_names = []
    for d in locs:
        c = str(g(d, "Corporate Name (LLC)", "") or "").replace("\n", ";")
        for x in c.split(";"):
            x = x.strip()
            if x and x not in corp_names:
                corp_names.append(x)
    client = str(g(first, "Client Name", "") or "").strip()
    ni2 = str(g(first, "Named Insured", "") or "").replace("\n", " ").strip()
    fni_source = ""
    if first_named_insured and str(first_named_insured).strip():
        ni1, fni_source = str(first_named_insured).strip(), "entered on the page"
    elif len(corp_names) > 1 and client and _entity_like(client) and client not in corp_names:
        ni1, fni_source = client, "SOV Client Name (the locations belong to %d corporate entities)" % len(corp_names)
    elif corp_names:
        ni1 = next((e for e in corp_names if e != ni2), corp_names[0])
        fni_source = "SOV Corporate Name (LLC)"
    elif ni2:
        ni1, fni_source = ni2, "SOV Named Insured"
        blanks.append("First Named Insured: SOV 'Corporate Name (LLC)' is blank, so %s is the only Named Insured - add the parent entity if there is one" % ni2)
    elif client:
        ni1, fni_source = client, "SOV Client Name"
    else:
        raise ValueError("No named insured found - fill 'Corporate Name (LLC)', 'Client Name' or 'Named Insured' in the SOV, or enter the First Named Insured on the page.")
    named = [ni1]
    if ni2 and ni2 != ni1:
        named.append(ni2)
    others = [c for c in corp_names if c not in named]  # entities listed on schedules / remarks

    try:
        m, dd, y = parse_eff(g(first, "Effective Date"))
    except Exception:
        raise ValueError("Effective Date is missing or unreadable in the SOV (expected M/D/YYYY).")
    eff_s, exp_s, prior_eff = "%02d/%02d/%d" % (m, dd, y), "%02d/%02d/%d" % (m, dd, y + 1), "%02d/%02d/%d" % (m, dd, y - 1)
    eff_file = "%02d-%02d-%d" % (m, dd, y)
    mail_state = str(g(first, "State", "") or "")
    mail1, mailcity, mailzip = parse_addr(g(first, "Mailing Address"), "", mail_state, "")
    mail_st = mail_state
    mraw = str(g(first, "Mailing Address", "") or "")
    mm = re.search(r",\s*([A-Z]{2})\s*,?\s*\d{5}", mraw)
    if mm:
        mail_st = mm.group(1)
    else:
        for full, ab in A_STATES.items():
            if (", %s," % full) in mraw or (", %s " % full) in mraw:
                mail_st = ab
                break
    if not mraw:
        blanks.append("Mailing address (SOV 'Mailing Address' is blank)")
    fein = next((g(d, "FEIN") for d in locs if g(d, "FEIN")), None)
    contact_name = str(g(first, "Application Contact Name", "") or "")
    contact_phone = str(g(first, "Application - Phone", "") or "")
    contact_email = str(g(first, "Inspection Email", "") or "")
    pol = []
    for d in locs:
        for p in str(g(d, "Policies", "") or "").split(","):
            if p.strip() and p.strip() not in pol:
                pol.append(p.strip())
    return {
        "today": TODAY, "locs": locs, "loc_rows": loc_rows, "first": first, "named": named, "others": others,
        "fni_source": fni_source, "corp_names": corp_names, "client": client,
        "eff": eff_s, "exp": exp_s, "prior_eff": prior_eff, "eff_file": eff_file, "year": y,
        "mail1": mail1, "mailcity": mailcity, "mailzip": mailzip, "mail_state": mail_st, "fein": fein,
        "contact": (contact_name, contact_phone, contact_email), "policies": ", ".join(pol),
        "warnings": warnings, "corrections": corrections, "blanks": blanks,
    }


A_STATES = {"Florida": "FL", "Georgia": "GA", "Alabama": "AL", "South Carolina": "SC", "North Carolina": "NC", "Tennessee": "TN",
            "Texas": "TX", "Louisiana": "LA", "Mississippi": "MS", "Virginia": "VA", "Kentucky": "KY", "Arkansas": "AR",
            "Oklahoma": "OK", "Missouri": "MO", "Illinois": "IL", "Indiana": "IN", "Ohio": "OH", "Michigan": "MI",
            "Pennsylvania": "PA", "New York": "NY", "New Jersey": "NJ", "Maryland": "MD", "Delaware": "DE", "Colorado": "CO",
            "Arizona": "AZ", "Nevada": "NV", "California": "CA", "Utah": "UT", "Washington": "WA", "Oregon": "OR",
            "Wisconsin": "WI", "Minnesota": "MN", "Iowa": "IA", "Kansas": "KS", "Nebraska": "NE", "West Virginia": "WV",
            "Massachusetts": "MA", "Connecticut": "CT", "New Mexico": "NM", "Idaho": "ID", "Montana": "MT", "Wyoming": "WY",
            "North Dakota": "ND", "South Dakota": "SD", "Vermont": "VT", "New Hampshire": "NH", "Maine": "ME", "Hawaii": "HI",
            "Alaska": "AK", "Rhode Island": "RI", "District of Columbia": "DC"}


# ─── helpers ────────────────────────────────────────────────────────────────

def pct100(v):
    v = nval(v, None)
    if v is None:
        return None
    return int(round(v * 100)) if v <= 1 else int(round(v))


def loc_state(d, ctx):
    return str(g(d, "State", "") or ctx["mail_state"] or "")


def addr_line(d, ctx):
    a1, city, zc = parse_addr(g(d, "Full Address"), g(d, "City", ""), loc_state(d, ctx), "")
    city = g(d, "City") or city
    return "%s, %s, %s %s" % (a1, city, loc_state(d, ctx), zc), a1, city, zc


def flag_short(d):
    fl = fmt_flag(d)
    return A.SHORT_FLAG.get(fl, fl)


def display_name(d):
    dba = str(g(d, "DBA", "") or "").strip()
    flag = flag_short(d)
    if not dba:
        return str(g(d, "Corporate Name (LLC)", "") or "Location %s" % d["Loc #"])
    if not flag or flag.lower() in dba.lower() or dba.lower() in flag.lower():
        return dba
    return "%s (%s)" % (dba, flag)


def loc_type(d):
    lt = str(g(d, "Location Type", "") or "").strip()
    if is_vacant(d):
        return "Vacant Building"
    if is_office(d):
        return "Office"
    if lt:
        return lt
    return "Hotel" if nval(g(d, "# of Rooms")) > 0 else "Other"


def describe(d, with_entity=False):
    """One-line description of a premises / building for the ACORD forms."""
    t = loc_type(d)
    dba = str(g(d, "DBA", "") or "").strip()
    if t == "Hotel":
        parts = ["%s - %s hotel" % (display_name(d), str(g(d, "Hotel Class", "")).lower()) if g(d, "Hotel Class") else "%s - hotel" % display_name(d)]
        if nval(g(d, "# of Rooms")) > 0:
            parts.append("%s rooms" % num(g(d, "# of Rooms")))
        if nval(g(d, "# of Floors")) > 0:
            parts.append("%s sty" % num(g(d, "# of Floors")))
        if yes(g(d, "Restaurant")):
            parts.append("restaurant/lounge")
        if nval(g(d, "# of Pools")) > 0:
            parts.append("pool")
        s = ", ".join(parts)
    elif t == "Vacant Building":
        s = "Vacant building - no operations"
    elif t == "Office":
        s = "Corporate office for hotel operations"
    else:
        s = "%s - %s" % (dba or t, t.lower()) if dba and t.lower() not in dba.lower() else (dba or t)
        if nval(g(d, "# of Floors")) > 0:
            s += ", %s sty" % num(g(d, "# of Floors"))
        if g(d, "SqFt") or g(d, "LRO (Sqft)"):
            s += ", %s sq ft" % num(g(d, "SqFt") or g(d, "LRO (Sqft)"))
    if with_entity and g(d, "Corporate Name (LLC)"):
        s = "%s (%s)" % (s, g(d, "Corporate Name (LLC)"))
    return s


def clean_code(v):
    """'103-Built-up roof NO gravel WITH presence of gutters' -> 'Built-up roof NO gravel WITH presence of gutters'."""
    s = str(v or "").strip()
    return re.sub(r"^\d{2,3}\s*-\s*", "", s)


def roof_short(v):
    """Short roof material for the 69 pt ROOF TYPE box: 'Built-up roof', 'Normal Shingle',
    'Concrete/Clay Tiles', 'Metal Sheathing'."""
    s = clean_code(v)
    s = re.sub(r"\s*\(.*?\)", "", s)
    s = re.split(r"\s+(NO|WITH|w/)\s+", s)[0]
    return s.strip()


def fire_protection(d):
    """Short fire-protection line: sprinklers, smoke alarms, fire alarm."""
    parts = []
    spr = pct100(g(d, "% Sprinklered"))
    if spr is not None:
        parts.append("Sprinklered %d%%" % spr if spr > 0 else "No sprinklers")
    sm = str(g(d, "Smoke Alarm", "") or "").lower()
    if sm:
        parts.append("hard-wired smoke alarms" if "hard" in sm else "smoke alarms: " + str(g(d, "Smoke Alarm")))
    fa = str(g(d, "Fire Alarm", "") or "").lower()
    if fa:
        parts.append("central station fire alarm" if "central" in fa else ("local fire alarm" if "local" in fa else "fire alarm: " + str(g(d, "Fire Alarm"))))
    return "; ".join(parts)


def class_split(v):
    """'45191 - Hotels --w Pools or Beaches--Four Stories or More' -> ('45191', 'Hotels - w Pools or Beaches - Four Stories or More')."""
    s = str(v or "").strip()
    m = re.match(r"^(\d{4,5})\s*-\s*(.*)$", s)
    code, desc = (m.group(1), m.group(2)) if m else ("", s)
    desc = re.sub(r"\s*-{2,}\s*", " - ", desc).strip()
    return code, desc


def wrap(text, size, width=510, font="helv"):
    words, lines, cur = str(text).split(" "), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if pymupdf.get_text_length(t, fontname=font, fontsize=size) > width and cur:
            lines.append(cur)
            cur = w
        else:
            cur = t
    if cur:
        lines.append(cur)
    return lines or [""]


def fit_size(text, width, base=8.0, floor=5.5, font="helv"):
    """Largest font size (<= base) at which `text` fits in `width` points."""
    size = base
    while size > floor and pymupdf.get_text_length(str(text), fontname=font, fontsize=size) > width - 3:
        size -= 0.5
    return size


def fit_box(text, width, height, base=8.0, floor=5.5):
    """(size, text) so that the wrapped text fits a multi-line widget; the text is
    cut with ' ...' when it does not fit even at the smallest size."""
    size = base
    while size >= floor:
        lines = wrap(text, size, width - 6)
        if len(lines) * size * 1.16 <= height - 1:
            return size, text
        size -= 0.5
    size = floor
    max_lines = max(1, int((height - 1) / (size * 1.16)))
    words = str(text).split(" ")
    while words:
        words = words[:-1]
        cand = " ".join(words) + " ..."
        if len(wrap(cand, size, width - 6)) <= max_lines:
            return size, cand
    return size, ""


def field_head(doc, xref):
    """The dictionary that owns the field name: the widget itself when it carries /T,
    else the nearest ancestor with /T (ACORD templates keep many values on a parent
    field object with the widget as its only kid)."""
    seen = 0
    while xref and seen < 8:
        if doc.xref_get_key(xref, "T")[0] != "null":
            return xref
        par = doc.xref_get_key(xref, "Parent")
        if par[0] != "xref":
            return xref
        xref = int(par[1].split()[0])
        seen += 1
    return xref


def wipe_text(doc, w):
    """Blank a text widget for real: clear the value on the field head, drop any
    shadowing /V on the widget and its appearance stream."""
    head = field_head(doc, w.xref)
    doc.xref_set_key(head, "V", "null")
    if head != w.xref:
        doc.xref_set_key(w.xref, "V", "null")
    doc.xref_set_key(w.xref, "AP", "null")


def fill(blank, values, font_rules=None, rename=None, shrink=True, pages=None):
    """Fill a copy of `blank`. values: field name -> str / bool. font_rules: list of
    (regex, size). Text that does not fit its widget is shrunk down to 5.5 pt.
    rename: suffix appended to every field name (unique names for duplicated pages).
    pages: page indices to process (default all; a duplicated page only needs its own)."""
    doc = pymupdf.open(blank)
    seen = set()
    font_rules = font_rules or []
    for pno, page in enumerate(doc):
        if pages is not None and pno not in pages:
            continue
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
                wipe_text(doc, w)
                for rx, size in font_rules:
                    if re.search(rx, n):
                        w.text_fontsize = size
                        break
                if val and shrink:
                    base = w.text_fontsize or 8.0
                    if base == 0:
                        base = 8.0
                    if w.rect.height < 20:  # single-line widget
                        sz = fit_size(val, w.rect.width, base=base)
                    else:                   # multi-line box: wrap, shrink, then cut
                        sz, val = fit_box(val, w.rect.width, w.rect.height, base=base)
                    if sz < base:
                        w.text_fontsize = sz
                w.field_value = val
            else:
                continue
            w.update()
            if rename:
                head = field_head(doc, w.xref)
                doc.xref_set_key(head, "T", "(%s%s)" % (n, rename))
    return doc, seen


_NAMES = {}


def template_fields(blank):
    """Set of field names on a template (cached per path)."""
    if blank not in _NAMES:
        d = pymupdf.open(blank)
        _NAMES[blank] = {w.field_name for p in d for w in p.widgets()}
        d.close()
    return _NAMES[blank]


def stale_check(out_path, warnings):
    chk = pymupdf.open(out_path)
    text = "\n".join(p.get_text() for p in chk)
    for p in chk:
        for w in p.widgets():
            if w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT and w.field_value:
                text += "\n" + str(w.field_value)
    n = len(chk)
    chk.close()
    stale = [s for s in STALE if s.lower() in text.lower()]
    if stale:
        warnings.append("Stale template strings still present: %s" % ", ".join(stale))
    return n


def safe_name(s):
    return re.sub(r"[\\/:*?\"<>|]+", "-", str(s)).strip()


def out_name(form, ctx):
    named = ctx["named"]
    return "ACORD %s - %s%s %s.pdf" % (form, safe_name(named[0]), " (%s)" % safe_name(named[1]) if len(named) > 1 else "", ctx["eff_file"])


class Sheet:
    """Simple continuation-schedule writer (plain pages, no ACORD title, no footer)."""

    def __init__(self, doc, title, ctx, note=""):
        self.doc, self.title, self.ctx, self.note = doc, title, ctx, note
        self.page, self.y = None, 0
        self.cols = []
        self.new_page(first=True)

    def new_page(self, first=False):
        self.page = self.doc.new_page(width=612, height=792)
        self.y = 40
        if first:
            self.page.insert_text((22, self.y), self.title, fontsize=9, fontname="helv")
            self.page.insert_text((22, self.y + 13), "Applicant: %s     Proposed effective date: %s     Date: %s"
                                  % (" / ".join(self.ctx["named"]), self.ctx["eff"], self.ctx["today"]), fontsize=8, fontname="helv")
            if self.note:
                for ln in wrap(self.note, 7.5, 560):
                    self.page.insert_text((22, self.y + 24), ln, fontsize=7.5, fontname="helv")
                    self.y += 9
            self.y += 44
        else:
            self.page.insert_text((22, self.y), self.title + " (continued)", fontsize=8, fontname="helv")
            self.y += 20
        if self.cols:
            self.header()

    def set_cols(self, cols):
        self.cols = cols
        self.header()

    def header(self):
        self.page.draw_rect(pymupdf.Rect(20, self.y - 10, 592, self.y + 3), color=(0, 0, 0), fill=(0.88, 0.88, 0.88), width=0.5)
        for t, x in self.cols:
            self.page.insert_text((x, self.y), t, fontsize=6.5, fontname="helv")
        self.y += 12

    def ensure(self, h):
        if self.y + h > 740:
            self.new_page()

    def row(self, cells, size=7, color=(0, 0, 0)):
        self.ensure(12)
        for x, t in cells:
            self.page.insert_text((x, self.y), str(t), fontsize=size, fontname="helv", color=color)
        self.y += 10

    def text(self, x, txt, size=6.5, color=(0.25, 0.25, 0.25), width=510):
        for ln in wrap(txt, size, width):
            self.ensure(9)
            self.page.insert_text((x, self.y), ln, fontsize=size, fontname="helv", color=color)
            self.y += 8.5

    def rule(self):
        self.page.draw_line((20, self.y - 4), (592, self.y - 4), color=(0.6, 0.6, 0.6), width=0.3)
        self.y += 6


def entity_note(ctx, rows=None):
    """'Named insureds (location entities): X (Loc 1); Y (Loc 2, 3).' for the entities
    behind the rows on this form (all SOV entities when rows is None)."""
    rows = ctx["locs"] if rows is None else rows
    ents = collections.OrderedDict()
    for d in rows:
        c = str(g(d, "Corporate Name (LLC)", "") or "").strip()
        if c and c not in ctx["named"]:
            ents.setdefault(c, []).append(str(d["Loc #"]))
    if not ents:
        return ""
    return "Named insureds (location entities): " + "; ".join(
        "%s (Loc %s)" % (c, ", ".join(sorted(set(l), key=int))) for c, l in ents.items()) + "."


# ─── ACORD 140 ──────────────────────────────────────────────────────────────

def _bi_forms(d):
    return "72 hrs"


def _subjects(d, blanks):
    """Rows for the premises block: (subject, amount, coins, valuation, cause, ded, forms)."""
    rows = []
    def add(subject, amt, coins="AA", val="RC", forms=""):
        if nval(amt) > 0:
            rows.append((subject, money(amt), coins, val, "Special", "", forms))
    add("Building", g(d, "Building Limit"))
    add("Business Personal Property", g(d, "Contents Limit"))
    if nval(g(d, "Business Income Limit")) > 0:
        rows.append(("Business Income w EE", money(g(d, "Business Income Limit")), "", "", "Special", "", _bi_forms(d)))
    if nval(g(d, "Other Limit")) > 0:
        add(str(g(d, "Other Limit - Description", "") or "Other"), g(d, "Other Limit"))
    sign, pool = nval(g(d, "Sign Limit")), nval(g(d, "Pool Limit"))
    if len(rows) >= 4 and sign > 0 and pool > 0:
        rows.append(("Outdoor Signs / Pool", money(sign + pool), "", "RC", "Special", "", ""))
    else:
        if sign > 0:
            rows.append(("Outdoor Signs", money(sign), "", "RC", "Special", "", ""))
        if pool > 0:
            rows.append(("Swimming Pool", money(pool), "", "RC", "Special", "", ""))
    return rows[:5]


def _premises_values(d, ctx, s, rowsfx, blanks, remark_key):
    """Values for one premises page. s = 'A' (page 1) or 'B' (page 2 / duplicates);
    rowsfx = 'ABCDE' or 'GHIJK'."""
    V = {}
    full, a1, city, zc = addr_line(d, ctx)
    V["CommercialStructure_Location_ProducerIdentifier_%s" % s] = str(d["Loc #"])
    V["CommercialStructure_PhysicalAddress_LineOne_%s" % s] = full
    V["CommercialStructure_Building_ProducerIdentifier_%s" % s] = str(d["Building #"])
    V["CommercialStructure_Building_SublocationDescription_%s" % s] = describe(d)
    for (subj, amt, coins, val, col, ded, forms), r in zip(_subjects(d, blanks), rowsfx):
        V["CommercialProperty_Premises_SubjectOfInsuranceCode_%s" % r] = subj
        V["CommercialProperty_Premises_LimitAmount_%s" % r] = amt
        V["CommercialProperty_Premises_CoinsurancePercent_%s" % r] = coins
        V["CommercialProperty_Premises_ValuationCode_%s" % r] = val
        V["CommercialProperty_Premises_CauseOfLossCode_%s" % r] = col
        V["CommercialProperty_Premises_DeductibleAmount_%s" % r] = ded
        V["CommercialProperty_Premises_FormsAndConditions_%s" % r] = forms
    V["CommercialProperty_Attachment_BusinessIncomeExtraExpenseIndicator_%s" % s] = nval(g(d, "Business Income Limit")) > 0
    # premises remark: owner entity, franchise, flood zone, coastal tier
    bits = []
    if g(d, "Corporate Name (LLC)"):
        bits.append("Owner / named insured: %s." % str(g(d, "Corporate Name (LLC)")).rstrip("."))
    brand, franchisor = str(g(d, "Franchise(s)", "") or ""), str(g(d, "Franchise", "") or "")
    if brand or franchisor:
        bits.append("Franchise: %s%s." % (brand or franchisor, " (%s)" % franchisor if brand and franchisor and franchisor.lower() not in brand.lower() else ""))
    if g(d, "Property Owned / Managed"):
        bits.append("Interest: %s." % g(d, "Property Owned / Managed"))
    if g(d, "Flood Zone"):
        bits.append("Flood zone %s (SOV)." % g(d, "Flood Zone"))
    if g(d, "Coastal Exposure"):
        bits.append("Coastal tier: %s." % g(d, "Coastal Exposure"))
    if yes(g(d, "Flood Coverage Required?")):
        bits.append("Flood coverage requested%s." % (" $%s" % money(g(d, "Flood Limit Requested")) if nval(g(d, "Flood Limit Requested")) > 0 else ""))
    if g(d, "Exterior Insulating Finishing System (EIFS)"):
        bits.append("EIFS: %s." % g(d, "Exterior Insulating Finishing System (EIFS)"))
    if yes(g(d, "Solar Panels (Roof)")):
        bits.append("Roof solar panels: Yes.")
    V[remark_key] = " ".join(bits)
    # construction block
    V["Construction_ConstructionCode_%s" % s] = str(g(d, "Construction", "") or "")
    V["Construction_StoreyCount_%s" % s] = num(g(d, "# of Floors"))
    bsm = str(g(d, "Basement", "") or "").strip().lower()
    V["Construction_BasementCount_%s" % s] = "0" if bsm.startswith("no") else ("1" if bsm.startswith("yes") else "")
    V["CommercialStructure_BuiltYear_%s" % s] = num(g(d, "Yr Built")).replace(",", "")
    V["Construction_BuildingArea_%s" % s] = num(g(d, "SqFt") or g(d, "LRO (Sqft)"))
    for col, key in [("Year Electrical Updated", "Wiring"), ("Year Roof Fully Replaced", "Roofing"),
                     ("Year Plumbing Updated", "Plumbing"), ("Year HVAC Updated", "Heating")]:
        yr = num(g(d, col)).replace(",", "")
        if yr and len(yr) == 4:
            V["BuildingImprovement_%sYear_%s" % (key, s)] = yr
            V["BuildingImprovement_%sIndicator_%s" % (key, s)] = True
    V["Construction_RoofMaterialCode_%s" % s] = roof_short(g(d, "Roof Type"))
    # fire protection
    spr = pct100(g(d, "% Sprinklered"))
    if spr is not None:
        V["BuildingFireProtection_Alarm_SprinklerPercent_%s" % s] = str(spr)
    fa = str(g(d, "Fire Alarm", "") or "")
    V["BuildingFireProtection_Alarm_ProtectionDescription_%s" % s] = fire_protection(d)
    V["BuildingFireProtection_Alarm_CentralStationIndicator_%s" % s] = "central" in fa.lower()
    V["BuildingFireProtection_Alarm_LocalGongIndicator_%s" % s] = "local" in fa.lower()
    # mortgagee
    if g(d, "Mortgage"):
        V["AdditionalInterest_Interest_MortgageeIndicator_%s" % s] = True
        V["AdditionalInterest_FullName_%s" % s] = str(g(d, "Mortgage"))
        V["AdditionalInterest_Item_LocationProducerIdentifier_%s" % s] = str(d["Loc #"])
        V["AdditionalInterest_Item_BuildingProducerIdentifier_%s" % s] = str(d["Building #"])
    return V


def generate_140(ctx, out_path, blank):
    warnings, corrections, blanks = ctx["warnings"], ctx["corrections"], list(ctx["blanks"])
    locs = ctx["locs"]
    prem = [d for d in locs if nval(g(d, "TIV")) > 0 or nval(g(d, "Building Limit")) > 0 or nval(g(d, "Contents Limit")) > 0]
    skipped = [d for d in locs if d not in prem and "Property" in str(g(d, "Policies", "") or "")]
    for d in skipped:
        warnings.append("Loc %s Bldg %s (%s): Property requested but no values on the SOV - not placed on the ACORD 140"
                        % (d["Loc #"], d["Building #"], display_name(d)))
    if not prem:
        raise ValueError("No building rows with property values - nothing to put on the ACORD 140.")
    tot = lambda k: sum(nval(g(d, k)) for d in prem)

    V = {
        "Form_CompletionDate_A": ctx["today"],
        "Producer_FullName_A": PRODUCER_NAME,
        "Policy_EffectiveDate_A": ctx["eff"],
        "NamedInsured_FullName_A": "; ".join(ctx["named"]) + (" (see remarks for location entities)" if ctx["others"] else ""),
        "Producer_AuthorizedRepresentative_FullName_A": "Stefan Burkey",
    }
    V.update(_premises_values(prem[0], ctx, "A", "ABCDE", blanks, "CommercialProperty_Premises_RemarkText_A"))
    if len(prem) > 1:
        V.update(_premises_values(prem[1], ctx, "B", "GHIJK", blanks, "CommercialProperty_Premises_RemarkText_C"))
    # portfolio remarks (page 2)
    by = lambda k: ", ".join("%s (%s)" % (v, "Loc %s" % l[0] if len(l) == 1 else "Locs %s" % ", ".join(l))
                              for v, l in collections.OrderedDict((str(g(d, k)), [str(x["Loc #"]) for x in prem if str(g(x, k)) == str(g(d, k))])
                                                                   for d in prem if g(d, k)).items())
    fl_locs = [str(d["Loc #"]) for d in prem if str(g(d, "Flood Zone", "")).upper().startswith(("A", "V"))]
    hotels = [d for d in prem if loc_type(d) == "Hotel"]
    all_spr = bool(hotels) and all((pct100(g(d, "% Sprinklered")) or 0) >= 100 for d in hotels)
    remarks = ("PORTFOLIO SUMMARY (per SOV): %d premises, %s rooms. Total TIV $%s (Building $%s; Contents $%s; Business Income $%s). "
               % (len(prem), num(sum(nval(g(d, "# of Rooms")) for d in hotels)), money(tot("TIV")), money(tot("Building Limit")),
                  money(tot("Contents Limit")), money(tot("Business Income Limit")))
               + "Valuation: replacement cost, agreed amount, special form; deductibles per quote. "
               + ("All hotels 100%% sprinklered with central station monitored fire alarm. " if all_spr else "")
               + "Construction: %s. " % (by("Construction") or "see SOV")
               + ("Coastal tiers: %s. " % by("Coastal Exposure") if by("Coastal Exposure") else "")
               + ("FEMA special flood hazard area (zone A/AE/V) at Locs %s; flood coverage requested per SOV. " % ", ".join(fl_locs) if fl_locs else "")
               + ("Premises 3 to %d on the additional premises pages that follow. " % len(prem) if len(prem) > 2 else "")
               + entity_note(ctx, prem) + " See the Statement of Values for every building, update year and construction detail.")
    V["CommercialPropertyLineOfBusiness_RemarkText_A"] = remarks

    rules = [(r"SublocationDescription_", 7), (r"RoofMaterialCode_", 6), (r"ProtectionDescription_", 7), (r"PhysicalAddress_LineOne_", 7.5)]
    doc, seen = fill(blank, V, rules)
    missing = [k for k in V if k not in seen]
    if missing:
        warnings.append("Template is missing %d expected field(s): %s" % (len(missing), ", ".join(missing[:8])))
    insert_at = 2
    for i, d in enumerate(prem[2:], 3):
        pv = _premises_values(d, ctx, "B", "GHIJK", blanks, "CommercialProperty_Premises_RemarkText_C")
        pv["Producer_FullName_A"] = PRODUCER_NAME
        pv["CommercialPropertyLineOfBusiness_RemarkText_A"] = "Additional premises %d of %d. Portfolio summary and named insureds on page 2." % (i, len(prem))
        doc2, _ = fill(blank, pv, rules, rename="_L%sB%s" % (d["Loc #"], d["Building #"]), pages=[1])
        doc.insert_pdf(doc2, from_page=1, to_page=1, start_at=insert_at)
        insert_at += 1
        doc2.close()
    doc.save(out_path, **SAVE_KW)
    doc.close()
    n_pages = stale_check(out_path, warnings)

    # blanks report
    def missing_on(col, label):
        ids = [str(d["Loc #"]) for d in prem if g(d, col) in (None, "")]
        if ids:
            blanks.append("%s (Loc %s)" % (label, ", ".join(ids)))
    missing_on("Yr Built", "Year built")
    missing_on("Construction", "Construction type")
    missing_on("# of Floors", "Number of stories")
    missing_on("% Sprinklered", "Sprinkler percentage")
    missing_on("Year Roof Fully Replaced", "Roof replacement year")
    missing_on("Year Electrical Updated", "Electrical update year")
    missing_on("Basement", "Basement (Y/N)")
    blanks += ["Property deductibles (per quote)", "Protection class, distance to hydrant / fire station, wind class",
               "Heating, burglar alarm and exposures blocks (not in the SOV)",
               "Sinkhole / mine subsidence elections", "Mortgagees (SOV 'Mortgage' column) where blank"]
    return {"out_name": out_name("140", ctx), "pages": n_pages, "locations": len({d["Loc #"] for d in prem}), "premises": len(prem),
            "tiv": int(tot("TIV")), "rooms": int(sum(nval(g(d, "# of Rooms")) for d in hotels)),
            "named_insureds": ctx["named"], "fni_source": ctx["fni_source"], "effective": ctx["eff"], "expiration": ctx["exp"],
            "lines": ctx["policies"], "corrections": corrections, "blanks": blanks, "warnings": warnings}


# ─── ACORD 126 ──────────────────────────────────────────────────────────────

def _hazards(d, ctx):
    """Hazard rows for one location: (loc, class code, basis, exposure, classification, note)."""
    rows = []
    t = loc_type(d)
    code, desc = class_split(g(d, "Liability Class Codes"))
    if t == "Hotel":
        sales = g(d, "Hotel Sales") or g(d, "Total Sales")
        rows.append((code, "Sales", money(sales), desc or "Hotels", ""))
        if nval(g(d, "Restaurant Sales")) > 0:
            pct = nval(g(d, "Liquor %"))
            rows.append(("", "Sales", money(g(d, "Restaurant Sales")), "Restaurant / lounge (owned)%s" % (" - liquor %s%% of receipts" % num(pct) if pct else ""), "class code to confirm"))
        if nval(g(d, "Liquor Sales")) > 0:
            rows.append(("", "Sales", money(g(d, "Liquor Sales")), "Liquor liability", "class code to confirm"))
        if nval(g(d, "Sundry Sales (Marketplace)")) > 0:
            rows.append(("", "Sales", money(g(d, "Sundry Sales (Marketplace)")), "Sundry / marketplace", ""))
    elif t.upper().startswith("LRO"):
        rows.append((code or "61212", "Area", num(g(d, "LRO (Sqft)") or g(d, "SqFt")), desc or "Lessor's risk only - buildings or premises", ""))
    elif t == "Vacant Building":
        rows.append((code or "68604", "Area", num(g(d, "LRO (Sqft)") or g(d, "SqFt")), desc or "Vacant buildings - not factories", ""))
    elif t == "Office":
        rows.append((code or "61223", "Area", num(g(d, "SqFt")), desc or "Buildings or premises - office", ""))
    else:
        exposure = g(d, "Total Sales") or g(d, "Other Sales")
        rows.append((code, "Sales" if nval(exposure) > 0 else "", money(exposure), desc or (str(g(d, "DBA", "")) or t), "" if code else "class code to confirm"))
    return rows


def generate_126(ctx, out_path, blank):
    warnings, corrections, blanks = ctx["warnings"], ctx["corrections"], list(ctx["blanks"])
    rows = [d for d in ctx["loc_rows"] if "Liability" in str(g(d, "Policies", "") or "")]
    if not rows:
        rows = list(ctx["loc_rows"])
        warnings.append("No location has 'Liability' in the SOV Policies column - every location was placed on the ACORD 126")
    hotels = [d for d in rows if loc_type(d) == "Hotel"]
    haz = []
    for d in rows:
        for i, (code, basis, exp, desc, note) in enumerate(_hazards(d, ctx), 1):
            haz.append((d, i, code, basis, exp, desc, note))
    primary = [h for h in haz if h[1] == 1]   # the form's three rows carry one hazard per location
    no_code = sorted({str(h[0]["Loc #"]) for h in haz if not h[2] and h[6]}, key=int)
    no_exp = sorted({str(h[0]["Loc #"]) for h in haz if not h[4]}, key=int)

    V = {
        "Form_CompletionDate_A": ctx["today"],
        "Producer_FullName_A": PRODUCER_NAME,
        "Policy_EffectiveDate_A": ctx["eff"],
        "NamedInsured_FullName_A": "; ".join(ctx["named"]) + (" (location entities on attached schedule)" if ctx["others"] else ""),
        "GeneralLiability_CoverageIndicator_A": True,
        "GeneralLiability_OccurrenceIndicator_A": True,
        "Producer_AuthorizedRepresentative_FullName_A": "Stefan Burkey",
    }
    V.update(GL_LIMITS)
    if len(rows) > 1:
        V["GeneralLiability_GeneralAggregate_LimitAppliesPerLocationIndicator_A"] = True
    else:
        V["GeneralLiability_GeneralAggregate_LimitAppliesPerPolicyIndicator_A"] = True
    for (d, i, code, basis, exp, desc, note), s in zip(primary, "ABC"):
        V["GeneralLiability_Hazard_LocationProducerIdentifier_%s" % s] = str(d["Loc #"])
        V["GeneralLiability_Hazard_HazardProducerIdentifier_%s" % s] = str(i)
        V["GeneralLiability_Hazard_ClassCode_%s" % s] = code
        V["GeneralLiability_Hazard_PremiumBasisCode_%s" % s] = basis
        V["GeneralLiability_Hazard_Exposure_%s" % s] = exp
        V["GeneralLiability_Hazard_Classification_%s" % s] = desc
    # remarks (page 1: 3 lines; page 4: 2 lines) - the long version goes on the schedule page
    total_sales = sum(nval(g(d, "Total Sales")) for d in rows)
    rest_locs = [str(d["Loc #"]) for d in rows if yes(g(d, "Restaurant"))]
    V["GeneralLiabilityLineOfBusiness_RemarkText_A"] = (
        "All %d locations (%d hazard rows) on the attached schedule of hazards; S = gross sales. Total annual sales $%s. "
        "Occurrence form, per-location aggregate, hired & non-owned auto; no A&B, A&M or human trafficking exclusions (franchise compliant)."
        % (len(rows), len(haz), money(total_sales)))
    V["GeneralLiabilityLineOfBusiness_RemarkText_B"] = (
        ("Owned restaurant / lounge at Locs %s (TIPS-trained servers). " % ", ".join(rest_locs) if rest_locs else "")
        + "Daily rentals only, background checks, pest control contract, annual HT awareness training. Loss runs (5 yrs) to follow.")
    # contractors (not applicable) and products questions
    for q in ["Contractors_Question_AABCode_A", "Contractors_Question_AAICode_A", "Contractors_Question_AACCode_A",
              "Contractors_Question_ABBCode_A", "Contractors_Question_AADCode_A", "Contractors_Question_AAECode_A",
              "GeneralLiabilityLineOfBusiness_Question_AAJCode_A", "GeneralLiabilityLineOfBusiness_Question_ABACode_A",
              "GeneralLiabilityLineOfBusiness_Question_ABBCode_A", "GeneralLiabilityLineOfBusiness_Question_ABCCode_A",
              "GeneralLiabilityLineOfBusiness_Question_ABDCode_A", "GeneralLiabilityLineOfBusiness_Question_ABECode_A",
              "GeneralLiabilityLineOfBusiness_Question_ABFCode_A", "GeneralLiabilityLineOfBusiness_Question_ABGCode_A",
              "GeneralLiabilityLineOfBusiness_Question_ABHCode_A", "GeneralLiabilityLineOfBusiness_Question_ABICode_A"]:
        V[q] = "N"
    # general information (page 3 / 4)
    Q = "GeneralLiabilityLineOfBusiness_Question_%sCode_A"
    V[Q % "ABJ"] = "N"   # 1 medical facilities
    V[Q % "ACA"] = "N"   # 2 radioactive
    V[Q % "ACJ"] = "N"   # 3 hazardous material
    acq = [d for d in rows if g(d, "Purchase Date") and _recent(g(d, "Purchase Date"), ctx["year"])]
    V[Q % "ACB"] = "Y" if acq else "N"   # 4 operations sold / acquired in 5 yrs
    if acq:
        V["GeneralLiabilityLineOfBusiness_AnyListedOperationsSoldAcquiredLastFiveYearsExplanation_A"] = \
            "Acquired: " + "; ".join("Loc %s %s (%s)" % (d["Loc #"], display_name(d), str(g(d, "Purchase Date")).split(" ")[0]) for d in acq)
    V[Q % "ACC"] = "N"   # 5 rent/loan equipment
    V["Contractors_Question_AAHCode_A"] = "N"   # 6 watercraft, docks
    garages = [d for d in rows if "parking" in (str(g(d, "Location Type", "")) + " " + str(g(d, "Amenities", "")) + " " + str(g(d, "DBA", ""))).lower()]
    V[Q % "ACD"] = "Y"   # 7 parking facilities
    V["GeneralLiabilityLineOfBusiness_AnyParkingFacilitiesOwnedRentedExplanation_A"] = \
        "Owned surface guest parking at all hotels" + ("; parking garage at Loc %s" % ", ".join(str(d["Loc #"]) for d in garages) if garages else "")
    fee = [d for d in garages if nval(g(d, "Other Sales")) > 0 or nval(g(d, "Total Sales")) > 0 and loc_type(d) != "Hotel"]
    valet = [d for d in rows if yes(g(d, "Valet Parking"))]
    V[Q % "KAG"] = "Y" if (fee or valet) else "N"   # 8 fee charged for parking
    if fee or valet:
        V["GeneralLiabilityLineOfBusiness_FeeChargedForParkingExplanation_A"] = \
            ("Public parking garage at Loc %s; " % ", ".join(str(d["Loc #"]) for d in fee) if fee else "") + \
            ("valet at Loc %s" % ", ".join(str(d["Loc #"]) for d in valet) if valet else "")
    rec = []
    am = " ".join(str(g(d, "Amenities", "")) + " " + str(g(d, "Fitness Room", "")) + " " + str(g(d, "Pool Amenities", "")) for d in hotels).lower()
    if "fitness" in am:
        rec.append("fitness room")
    pools = [d for d in hotels if nval(g(d, "# of Pools")) > 0]
    if pools:
        rec.append("swimming pool (guest only)")
    if "hot tub" in am or "spa" in am:
        rec.append("spa / hot tub")
    if "tennis" in am:
        rec.append("tennis court")
    V[Q % "ACE"] = "Y" if rec else "N"   # 9 recreation facilities
    V["GeneralLiabilityLineOfBusiness_RecreationalFacilitiesProvidedExplanation_A"] = ", ".join(rec)
    rooms = int(sum(nval(g(d, "# of Rooms")) for d in hotels))
    V[Q % "KAA"] = "Y" if hotels else "N"   # 10 lodging operations incl. apartments
    V["GeneralLiability_Otherlodgingoperationsdescription_A"] = "Hotel guest rooms only: %s rooms at %d hotel%s; no apartments" % (num(rooms), len(hotels), "" if len(hotels) == 1 else "s") if hotels else ""
    V[Q % "KAH"] = "Y" if pools else "N"   # 11 swimming pool
    if pools:
        pa = " ".join(str(g(d, "Pool Amenities", "")) for d in pools).lower()
        V["SwimmingPool_ApprovedFenceIndicator_A"] = "fenced" in pa or "restricted access" in pa
        V["SwimmingPool_LimitedAccessIndicator_A"] = "guest only" in pa or "restricted access" in pa
        V["SwimmingPool_DivingBoardIndicator_A"] = "diving board" in pa and "no diving board" not in pa
        V["SwimmingPool_SlideIndicator_A"] = ("slide" in pa) and ("no slide" not in pa)
        V["SwimmingPool_InGroundIndicator_A"] = True
    ev = [d for d in rows if yes(g(d, "Any known sponsored events for upcoming year"))]
    V[Q % "ACF"] = "Y" if ev else "N"   # 12 social events sponsored
    if ev:
        V["GeneralLiabilityLineOfBusiness_SportingOrSocialEventsSponsoredExplanation_A"] = "Loc %s - see SOV" % ", ".join(str(d["Loc #"]) for d in ev)
    V[Q % "KAB"] = "N"   # 13 athletic teams
    cons = [d for d in rows if yes(g(d, "Any known construction for the upcoming year"))]
    V[Q % "ACG"] = "Y" if cons else "N"   # 14 structural alterations
    if cons:
        V["GeneralLiabilityLineOfBusiness_StructuralAlterationsContemplatedExplanation_A"] = "Loc %s - see SOV" % ", ".join(str(d["Loc #"]) for d in cons)
    V[Q % "ACH"] = "N"   # 15 demolition
    V[Q % "AAB"] = "N"   # 16 joint ventures
    V[Q % "AAC"] = "N"   # 17 lease employees
    V[Q % "AAD"] = "N"   # 18 labor interchange
    V[Q % "AAE"] = "N"   # 19 day care
    V[Q % "AAF"] = "N"   # 20 crimes on premises (3 yrs)
    V["Contractors_Question_AAGCode_A"] = "Y"   # 21 formal written safety and security policy
    V[Q % "AAG"] = "N"   # 22 promotional literature safety representations

    rules = [(r"Hazard_Classification_", 7), (r"RemarkText_", 7.5), (r"Explanation_", 7.5)]
    doc, seen = fill(blank, V, rules)
    missing = [k for k in V if k not in seen]
    if missing:
        warnings.append("Template is missing %d expected field(s): %s" % (len(missing), ", ".join(missing[:8])))
    # continuation schedule of hazards
    note = ("Premium basis: Sales = (S) gross sales, Area = (A) square feet. Exposures per the Statement of Values; class codes marked "
            "'to confirm' are left for the producer. Underwriting subjectivities: over 3 years of operational experience; no prior A&M, "
            "A&B or human trafficking losses; no liability losses over $100,000; pest control contract incl. bed bug prevention; background "
            "checks on all employees; no homeless shelter, hourly or >30 day rentals; annual human trafficking awareness training; TIPS-trained "
            "servers with written intoxicated-patron guidelines; HNOA controls (<5 employees, <20 trips/month, employee auto insurance confirmed). "
            + entity_note(ctx, rows))
    sh = Sheet(doc, "SCHEDULE OF HAZARDS CONTINUATION (ADDITIONAL HAZARDS SCHEDULE)", ctx, note)
    sh.set_cols([("LOC #", 22), ("HAZ #", 50), ("CLASS CODE", 78), ("PREMIUM BASIS", 130), ("EXPOSURE", 196), ("CLASSIFICATION / PREMISES", 262)])
    cur = None
    for (d, i, code, basis, exp, desc, note) in haz:
        if d is not cur:
            if cur is not None:
                sh.rule()
            full = addr_line(d, ctx)[0]
            sh.text(22, "Loc %s - %s - %s%s" % (d["Loc #"], display_name(d), full,
                                               " (%s)" % g(d, "Corporate Name (LLC)") if g(d, "Corporate Name (LLC)") else ""), 7, (0, 0, 0), 560)
            cur = d
        sh.row([(22, d["Loc #"]), (50, i), (78, code or ("to confirm" if note else "-")), (130, basis), (196, exp or "-"), (262, desc[:95])], 6.8)
    sh.rule()
    sh.row([(22, "TOTALS: %d locations | %d hazard rows | annual sales $%s" % (len(rows), len(haz), money(total_sales)))], 7.5)
    doc.save(out_path, **SAVE_KW)
    doc.close()
    n_pages = stale_check(out_path, warnings)
    if no_code:
        blanks.append("GL class codes for restaurant / liquor / other hazards (Loc %s)" % ", ".join(no_code))
    if no_exp:
        blanks.append("Sales exposure (Loc %s)" % ", ".join(no_exp))
    blanks += ["Territory codes and rates", "Additional interests / additional insureds (franchisors, mortgagees)",
               "Questions 4, 17, 18, 20 confirmed as No - adjust if any acquisition, leased employees, labor interchange or crime applies",
               "Claims-made / tail questions (left blank, occurrence form)"]
    return {"out_name": out_name("126", ctx), "pages": n_pages, "locations": len(rows), "hazards": len(haz),
            "sales": int(total_sales), "rooms": rooms, "named_insureds": ctx["named"], "fni_source": ctx["fni_source"],
            "effective": ctx["eff"], "expiration": ctx["exp"], "lines": ctx["policies"],
            "corrections": corrections, "blanks": blanks, "warnings": warnings}


def _recent(v, year):
    try:
        m, dd, y = parse_eff(v)
        return year - y <= 5
    except Exception:
        return False


# ─── ACORD 130 ──────────────────────────────────────────────────────────────

def _wc_rows(d):
    """Class rows for one location: (class, description, employees, remuneration)."""
    rows = []
    emp = nval(g(d, "# Full Time Employees")) + nval(g(d, "# Part Time Employees"))
    t = loc_type(d)
    any_pay = any(nval(g(d, col)) > 0 for col, _, _ in WC_CLASSES)
    for col, code, desc in WC_CLASSES:
        pay = nval(g(d, col))
        if code == "9052":
            if t != "Hotel":
                continue
            rows.append((code, desc, num(emp) if emp else "", money(pay) if pay > 0 else (money(g(d, "Total Payroll")) if not any_pay else "")))
        elif pay > 0:
            rows.append((code, desc, "", money(pay)))
        elif t == "Hotel" and not any_pay and (code == "8810" or (code == "9058" and yes(g(d, "Restaurant")))):
            rows.append((code, desc, "", ""))   # blank row for the client to complete
    if t != "Hotel":
        rows.insert(0, ("", "%s - %s (class to confirm)" % (display_name(d), t.lower()), num(emp) if emp else "",
                        money(g(d, "Total Payroll")) if nval(g(d, "Total Payroll")) > 0 else ""))
    return rows


def generate_130(ctx, out_path, blank):
    warnings, corrections, blanks = ctx["warnings"], ctx["corrections"], list(ctx["blanks"])
    rows = [d for d in ctx["loc_rows"] if re.search(r"\bWC\b|Workers", str(g(d, "Policies", "") or ""), re.I)]
    if not rows:
        rows = [d for d in ctx["loc_rows"] if loc_type(d) == "Hotel"] or list(ctx["loc_rows"])
        warnings.append("No location has 'WC' in the SOV Policies column - the hotels were placed on the ACORD 130; confirm which entities need workers compensation")
    if os.path.basename(blank).startswith("acord130fl"):
        warnings.append("ACORD 130 FL edition used: drop the standard ACORD 130 blank at assets/acord130_blank.pdf to switch editions")
    hotels = [d for d in rows if loc_type(d) == "Hotel"]
    first = ctx["first"]
    names = template_fields(blank)
    V = {
        "Form_CompletionDate_A": ctx["today"],
        "Producer_FullName_A": PRODUCER_NAME,
        "Producer_MailingAddress_LineOne_A": A.PRODUCER["Producer_MailingAddress_LineOne_A"],
        "Producer_MailingAddress_CityName_A": "Chicago",
        "Producer_MailingAddress_StateOrProvinceCode_A": "IL",
        "Producer_MailingAddress_PostalCode_A": "60601",
        "Producer_ContactPerson_PhoneNumber_A": A.PRODUCER["Producer_ContactPerson_PhoneNumber_A"],
        "NamedInsured_FullName_A": "; ".join(ctx["named"]) + (" (entities on attached schedule)" if ctx["others"] else ""),
        "NamedInsured_MailingAddress_LineOne_A": ctx["mail1"],
        "NamedInsured_MailingAddress_CityName_A": ctx["mailcity"],
        "NamedInsured_MailingAddress_StateOrProvinceCode_A": ctx["mail_state"],
        "NamedInsured_MailingAddress_PostalCode_A": ctx["mailzip"],
        "NamedInsured_SICCode_A": "7011",
        "Policy_Status_QuoteIndicator_A": True,
        "Policy_Payment_DirectBillIndicator_A": True,
        "Policy_Payment_AnnualIndicator_A": True,
        "Policy_Audit_AtExpirationIndicator_A": True,
        "Policy_EffectiveDate_A": ctx["eff"],
        "Policy_ExpirationDate_A": ctx["exp"],
        "Policy_ParticipatingIndicator_A": True,
        "WorkersCompensationEmployersLiability_EmployersLiability_EachAccidentLimitAmount_A": EL_LIMIT,
        "WorkersCompensationEmployersLiability_EmployersLiability_DiseasePolicyLimitAmount_A": EL_LIMIT,
        "WorkersCompensationEmployersLiability_EmployersLiability_DiseaseEachEmployeeLimitAmount_A": EL_LIMIT,
        "Client_Attachment_AdditionalLocationsIndicator_A": len(rows) > 3,
    }
    # optional fields that only the standard edition carries
    for k, v in [("Producer_ContactPerson_FullName_A", "Stefan Burkey"),
                 ("Producer_ContactPerson_EmailAddress_A", A.PRODUCER["Producer_ContactPerson_EmailAddress_A"]),
                 ("NamedInsured_NAICSCode_A", "721110")]:
        if k in names:
            V[k] = v
    et = entity_type(ctx["named"][0])
    if "NamedInsured_LegalEntity_%sIndicator_A" % et in names:
        V["NamedInsured_LegalEntity_%sIndicator_A" % et] = True
    elif et == "LimitedLiabilityCorporation" and "NamedInsured_LegalEntity_OtherIndicator_A" in names:
        V["NamedInsured_LegalEntity_OtherIndicator_A"] = True
        V["NamedInsured_LegalEntity_OtherDescription_A"] = "LLC"
    if ctx["fein"]:
        V["NamedInsured_TaxIdentifier_A"] = str(ctx["fein"])
    else:
        blanks.append("FEIN (per named insured)")
    states = []
    for d in rows:
        st = loc_state(d, ctx)
        if st and st not in states:
            states.append(st)
    for st, s in zip(states, "ABCDEFGHIJ"):
        V["WorkersCompensation_PartOne_StateOrProvinceCode_%s" % s] = st
    for d, s in zip(rows, "ABC"):
        full, a1, city, zc = addr_line(d, ctx)
        V["Location_ProducerIdentifier_%s" % s] = str(d["Loc #"])
        V["Location_PhysicalAddress_LineOne_%s" % s] = a1
        V["Location_PhysicalAddress_CityName_%s" % s] = city
        V["Location_PhysicalAddress_CountyName_%s" % s] = str(g(d, "County.", "") or "")
        V["Location_PhysicalAddress_StateOrProvinceCode_%s" % s] = loc_state(d, ctx)
        V["Location_PhysicalAddress_PostalCode_%s" % s] = zc
    # rating information
    rating = []
    for d in rows:
        for code, desc, emp, pay in _wc_rows(d):
            rating.append((d, code, desc, emp, pay))
    for (d, code, desc, emp, pay), s in zip(rating, "ABCDE"):
        V["WorkersCompensation_RateClass_LocationProducerIdentifier_%s" % s] = str(d["Loc #"])
        V["WorkersCompensation_RateClass_ClassificationCode_%s" % s] = code
        V["WorkersCompensation_RateClass_DutiesDescription_%s" % s] = desc
        V["WorkersCompensation_RateClass_EmployeeCount_%s" % s] = emp
        V["WorkersCompensation_RateClass_EstimatedRemuneration_%s" % s] = pay
    V["WorkersCompensation_AdditionalClassCodesIndicator_A"] = len(rating) > 5
    # nature of business
    flags = sorted({flag_short(d) for d in hotels if flag_short(d)})
    rooms = int(sum(nval(g(d, "# of Rooms")) for d in hotels))
    non_hotel = [d for d in rows if loc_type(d) != "Hotel"]
    rest_locs = [str(d["Loc #"]) for d in rows if yes(g(d, "Restaurant"))]
    V["CommercialPolicy_OperationsDescription_A"] = (
        "Owner/operator of %d franchised hotel%s (%s guest rooms) in %s%s%s. Front desk, housekeeping, maintenance, management and "
        "clerical staff%s. No owned autos or shuttle; background checks on all employees; annual HT awareness training."
        % (len(hotels), "" if len(hotels) == 1 else "s", num(rooms), ", ".join(states) if states else "the state of operation",
           " under %s flags" % ", ".join(flags) if flags else "",
           " plus %d non-hotel location%s (see remarks)" % (len(non_hotel), "" if len(non_hotel) == 1 else "s") if non_hotel else "",
           "; restaurant / lounge staff at Locs %s" % ", ".join(rest_locs) if rest_locs else ""))
    # general information questions (24)
    Qy = "WorkersCompensationLineOfBusiness_Question_%sYesIndicator_A"
    Qn = "WorkersCompensationLineOfBusiness_Question_%sNoIndicator_A"
    part_time = any(nval(g(d, "# Part Time Employees")) > 0 for d in rows)
    answers = {"ABE": False, "ACD": False, "AAD": False, "KAR": False, "KAS": bool(non_hotel), "KAT": True, "KAU": False, "ABC": True,
               "ABI": False, "AAE": False, "KAV": part_time, "AAF": False, "ABJ": False, "ABH": False, "AAB": False,
               "ACB": False, "ABA": False, "AAI": False, "ABF": False, "KAW": False, "AAG": False, "ABG": False}
    for code, ans in answers.items():
        V[Qy % code] = ans
        V[Qn % code] = not ans
    V["WorkersCompensationNoticeOfAssignment_Question_KAKYesIndicator_A"] = False
    V["WorkersCompensationNoticeOfAssignment_Question_KAKNoIndicator_A"] = True
    total_sales = sum(nval(g(d, "Total Sales")) for d in rows)
    V["WorkersCompensation_EstimatedAnnualRevenues_A"] = money(total_sales)
    name, phone, email = ctx["contact"]
    V["NamedInsured_InspectionContact_FullName_A"] = name
    V["NamedInsured_InspectionContact_PhoneNumber_A"] = phone
    if not (name and phone):
        blanks.append("Inspection / accounting / claims contacts")
    # remarks
    pay_tot = collections.OrderedDict()
    for d in rows:
        for col, code, desc in WC_CLASSES:
            if nval(g(d, col)) > 0:
                pay_tot[code or "Other"] = pay_tot.get(code or "Other", 0) + nval(g(d, col))
    emp_tot = int(sum(nval(g(d, "# Full Time Employees")) + nval(g(d, "# Part Time Employees")) for d in rows))
    rem = ["%d location%s and %d class row%s on the attached schedules." % (len(rows), "" if len(rows) == 1 else "s", len(rating), "" if len(rating) == 1 else "s")]
    if pay_tot:
        rem.append("Est. annual remuneration: " + "; ".join("%s $%s" % (c, money(v)) for c, v in pay_tot.items()) + ".")
    else:
        rem.append("Payroll by class and employee counts to be provided by the client (not on the SOV).")
    if emp_tot:
        rem.append("%s employees in total." % num(emp_tot))
    rem.append("Q6: outside contractors (landscaping, pool, pest control, elevator) with certificates. Q8: written safety program, HT awareness training, TIPS where alcohol is served.")
    if non_hotel:
        rem.append("Q5: " + "; ".join("Loc %s %s" % (d["Loc #"], describe(d)) for d in non_hotel) + ".")
    rem.append("Loss runs (5 yrs), prior carrier and experience mod to follow.")
    V["WorkersCompensationLineOfBusiness_RemarkText_A"] = " ".join(x for x in rem if x)
    # ownership / combinability
    ents = entity_note(ctx, rows)
    if ents:
        V["WorkersCompensationLineOfBusiness_Question_KBIYesIndicator_A"] = True
        n_ent = len({str(g(d, "Corporate Name (LLC)", "") or "").strip() for d in rows if g(d, "Corporate Name (LLC)")})
        V["WorkersCompensationLineOfBusiness_BusinessesRelatedByCommonOwnershipExplanation_A"] = \
            "%d entities under common ownership of %s - names, addresses and FEINs on the attached locations schedule" % (n_ent, ctx["named"][0])
        V["WorkersCompensation_FormerNamesAndOwners_A"] = ents.replace("Named insureds (location entities): ", "Entities to be covered: ")
    if g(first, "Exp Mod"):
        V["PriorCoverage_ModificationFactor_A"] = str(g(first, "Exp Mod"))
    wc_carrier = g(first, "Expiring WC Carrier") or g(first, "Expiring Workers Comp Carrier")
    if wc_carrier:
        V["PriorCoverage_EffectiveYear_A"] = str(ctx["year"] - 1)
        V["PriorCoverage_InsurerFullName_A"] = str(wc_carrier)

    rules = [(r"DutiesDescription_", 7), (r"OperationsDescription_", 7), (r"RemarkText_", 7), (r"Explanation_", 7)]
    doc, seen = fill(blank, V, rules)
    missing = [k for k in V if k not in seen]
    if missing:
        warnings.append("Template is missing %d expected field(s): %s" % (len(missing), ", ".join(missing[:8])))
    # continuation schedules
    if len(rows) > 3 or ents:
        sh = Sheet(doc, "LOCATIONS CONTINUATION (ADDITIONAL LOCATIONS SCHEDULE)", ctx,
                   "All physical locations where employees work, with the entity that employs them, per the Statement of Values. FEIN per entity to be provided where blank.")
        sh.set_cols([("#", 22), ("STREET, CITY, STATE, ZIP", 44), ("COUNTY", 300), ("ENTITY (NAMED INSURED) / FEIN", 370)])
        for d in rows:
            full = addr_line(d, ctx)[0]
            ent = str(g(d, "Corporate Name (LLC)", "") or "")
            if g(d, "FEIN"):
                ent += " / %s" % g(d, "FEIN")
            sh.row([(22, d["Loc #"]), (44, full[:60]), (300, g(d, "County.", "") or ""), (370, ent[:50])], 6.8)
            sh.text(44, "%s" % describe(d), 6.3)
            sh.y += 3
        sh.rule()
    if len(rating) > 5:
        sh = Sheet(doc, "RATING INFORMATION CONTINUATION (ADDITIONAL CLASS CODES)", ctx,
                   "Class codes per NCCI; employee counts and estimated annual remuneration per the Statement of Values (blank where not provided).")
        sh.set_cols([("LOC", 22), ("CLASS", 48), ("CATEGORIES, DUTIES, CLASSIFICATIONS", 88), ("# EMPLOYEES", 380), ("EST. ANNUAL REMUNERATION", 450)])
        cur = None
        for (d, code, desc, emp, pay) in rating:
            if d is not cur:
                if cur is not None:
                    sh.rule()
                sh.text(22, "Loc %s - %s - %s" % (d["Loc #"], display_name(d), addr_line(d, ctx)[0]), 7, (0, 0, 0), 560)
                cur = d
            sh.row([(22, d["Loc #"]), (48, code or "TBD"), (88, desc[:70]), (380, emp or "-"), (450, ("$" + pay) if pay else "-")], 6.8)
        sh.rule()
        tot_pay = sum(pay_tot.values())
        sh.row([(22, "TOTALS: %d locations | %d class rows | %s employees | estimated remuneration %s"
                 % (len(rows), len(rating), num(emp_tot) or "-", ("$" + money(tot_pay)) if tot_pay else "to be provided"))], 7.5)
    doc.save(out_path, **SAVE_KW)
    doc.close()
    n_pages = stale_check(out_path, warnings)
    if not pay_tot:
        blanks.append("Payroll by class (Hotel / Clerical / Restaurant / Sales payroll columns are blank)")
    if not emp_tot:
        blanks.append("Employee counts (# Full Time / # Part Time Employees)")
    blanks += ["Individuals included / excluded (officers, owners)", "Prior carrier, policy numbers, premiums and experience mod",
               "Loss history (5 yrs - to be completed from loss runs)", "Years in business (not on the SOV)",
               "Questions 10, 14, 17-22 answered No by default - adjust where the client's answer differs",
               "Other states (Part 3), deductibles and dividend plans"]
    return {"out_name": out_name("130", ctx), "pages": n_pages, "locations": len(rows), "class_rows": len(rating), "employees": emp_tot,
            "payroll": int(sum(pay_tot.values())), "states": states, "rooms": rooms, "named_insureds": ctx["named"], "fni_source": ctx["fni_source"],
            "effective": ctx["eff"], "expiration": ctx["exp"], "lines": ctx["policies"],
            "corrections": corrections, "blanks": blanks, "warnings": warnings,
            "edition": "FL" if os.path.basename(blank).startswith("acord130fl") else "standard"}


# ─── entry points ───────────────────────────────────────────────────────────

def blank_for(form):
    for p in BLANKS[form]:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("No ACORD %s blank template found (looked for %s)" % (form, ", ".join(BLANKS[form])))


GENERATORS = {"126": generate_126, "140": generate_140, "130": generate_130}


def generate_form(form, sov_path, out_path, first_named_insured=None, blank_path=None, today=None, ctx=None):
    form = str(form)
    if form == "125":
        return A.generate(sov_path, out_path, first_named_insured=first_named_insured, blank_path=blank_path, today=today)
    if form not in GENERATORS:
        raise ValueError("Unknown ACORD form %r (use 125, 126, 130 or 140)" % form)
    ctx = ctx or load(sov_path, first_named_insured=first_named_insured, today=today)
    ctx = dict(ctx, warnings=list(ctx["warnings"]), corrections=list(ctx["corrections"]), blanks=list(ctx["blanks"]))
    return GENERATORS[form](ctx, out_path, blank_path or blank_for(form))


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    info = generate_form(sys.argv[1], sys.argv[2], sys.argv[3], first_named_insured=(sys.argv[4] if len(sys.argv) > 4 else None))
    print(json.dumps(info, indent=2))
