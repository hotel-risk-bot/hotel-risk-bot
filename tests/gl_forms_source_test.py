"""Regression checks for the liability forms audit and scrambled-text policies.

Run from the repo root:  python3 tests/gl_forms_source_test.py
Self-contained: synthetic text in the layout pdftotext -layout / OCR produces, no OpenAI calls.

Why (Nov 17 2026, WS Management Inc): the submission had two GL policies that both bind, a
Southlake (Futuristic) package quote for twelve hotels and the in-force Burlington policy for one
hotel. The proposal's second GL section listed Southlake's 49 forms, and both GL sections gained
three "forms" that were hotel addresses:
  * the Burlington PDF (a Vertafore export) draws its text with fonts that carry no character
    map, so every page extracted as digits and symbols; its carrier name was nowhere in its text,
    and the audit's keyword score picked the Southlake file for the Burlington coverage;
  * Southlake's Schedule of Locations rows ("10   7216 55th Street, Sacramento, CA, 95823") read
    as form "10": a bare two-digit number passed as a form number (rows 1-9 were one digit).

  * forms_schedule_audit.pick_source_text - source tag, then the one file printing the premium,
    then only files that carry the carrier's name (text or file name); never a keyword guess
  * forms_schedule_audit._is_form_number / _row_desc_ok - no bare numbers, no street addresses
  * forms_schedule_audit - OCR-style single-spaced rows, "LISTING OF FORMS AND ENDORSEMENTS"
    header across pages (the Human Trafficking exclusion sat alone on page 3), "D-2 01 20"
  * proposal_extractor._page_scrambled + extract_text_from_pdf_smart - a mostly scrambled file is
    read by OCR (first 30 pages), the best-scored 30K characters kept, a review note added
"""
import os
import re
import subprocess
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import forms_schedule_audit as fsa  # noqa: E402
import proposal_extractor as pe  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


PACKAGE_FILE = "1700000001.pdf"
POLICY_FILE = "Tulare GL Harborline 363HX10001-02 2025-26.pdf"

# Futuristic-style package quote: a schedule of locations right before the schedule of forms.
PACKAGE_TEXT = """
--- Page 15 ---
                         GENERAL LIABILITY COVERAGE PART
TOTAL PACKAGE PREMIUM                             $229,962.00
 Stamping Fee                                        $413.93
 Surplus Lines Tax                                  $6,898.86
TOTAL PACKAGE COST                                $237,274.79
Commercial general liability, general liability limits, general liability premium.
Lakeview Specialty Insurance Company - general liability package.

--- Page 16 ---
                                               Schedule of Locations
1    100 HARBOR PLACE, San Bernardino, CA, 92408

2    200 HARNEY STREET, San Diego, CA, 92110

9    900 Milliken Ave, Rancho Cucamonga, CA, 91730

10   7216 55th Street, Sacramento, CA, 95823

11   1085 East El Camino Real, Sunnyvale, CA, 94087

12   1350 Veterans Blvd, San Bernardino, CA, 92408
                                                      FUT 1005 04/24
                        SCHEDULE OF FORMS

FORM NAME     EDITION   FORM DESCRIPTION                                      COVERAGE LINE

IL DS 00      09/08     Common Policy Declarations                            Interline
CG DS 01      10/10     General Liability Declarations                        General Liability
FUT 1005      04/24     Schedule of Locations                                 Interline
CG 00 01      04/13     Commercial General Liability Coverage Form            General Liability
CG 21 47      12/07     Employment Related Practices Exclusion                General Liability
GLF-105       04/24     Premium Audit Conditions Amended                      General Liability
GLF-101       04/24     Exclusion - Asbestos                                  General Liability
IL 00 17      11/01     Common Policy Conditions                              Interline
"""

# The in-force policy, as pdftotext returns it: junk on every page, one readable charges box.
_JUNK = ("8295 30 5 3 012 542 21 295 30043 1504 5 2310 0 501 45502 292 21431 033450 530 5 8 "
         "K7^H^T^DW=Y=K09OHK92^K7^ 0?AJ C?AFG;?;EAJ:@V E:;FP:?:8G89B8:G8 ")
POLICY_SCRAMBLED = "\n".join(f"--- Page {p} ---\n" + _JUNK * 12 for p in range(1, 5)) + """
--- Page 5 ---
""" + _JUNK * 6 + """
                                                             Summary of Charges
                                                             Premium                $42,559.00
                                                             Broker Fee                $400.00
                                                             Total                  $44,312.38
"""

# The same policy read by OCR: single-spaced number-first rows, a three-page listing.
POLICY_OCR = """
--- Page 5 ---
COMMON POLICY DECLARATIONS THE HARBORLINE INSURANCE COMPANY
Item 1. Named Insured and Mailing Address Cherry Hospitality, LLC dba Hampton Inn and Suites
Commercial General Liability $40,532 Terrorism Risk Insurance Act $2,027
Total Policy Premium or Deposit Premium $42,559
Summary of Charges Premium $42,559.00 Broker Fee $400.00 Total $44,312.38

--- Page 6 ---
POLICY NUMBER: 363HX10001-02
THIS ENDORSEMENT CHANGES THE POLICY. PLEASE READ IT CAREFULLY.
LISTING OF FORMS AND ENDORSEMENTS
This listing forms a part of the following:
Commercial General Liability Policy
NUMBER TITLE
Cover Page
HLX-I-0002 05 24 Policy Cover Page
Interline
HLX-I-0101 10 25 Common Policy Declarations
D-2 01 20 California Surplus Lines Notice
HLX-I-0978 12 20 Policyholder Disclosure Notice of Terrorism Insurance Coverage
IL 00 17 11 98 Common Policy Conditions
IL P 001 01 04 U.S. Treasury Department's "OFAC" Advisory Notice to Policyholders
Commercial General Liability
CG 00 01 04 13 Commercial General Liability Coverage Form
CG 21 44 07 98 Limitation Of Coverage To Designated Premises or Project
HLX-I-0150 03 03 Page 1 of 3

--- Page 7 ---
POLICY NUMBER: 363HX10001-02
LISTING OF FORMS AND ENDORSEMENTS
NUMBER TITLE
CG 21 47 12 07 Employment Related Practices Exclusion
HLX-G-0052 03 21 Deductible Liability Insurance
HLX-G-0104 11 21 Exclusion - Lead Bearing Substance
HLX-I-0150 03 03 Page 2 of 3

--- Page 8 ---
POLICY NUMBER: 363HX10001-02
LISTING OF FORMS AND ENDORSEMENTS
NUMBER TITLE
HLX-G-0302 12 24 Exclusion - Human Trafficking
HLX-G-0322 08 23 Excl - Mechanical Bucking Ride(s), Device(s) or Equipment
HLX-I-0150 03 03 Page 3 of 3

--- Page 11 ---
SUPPLEMENTAL SCHEDULE OF LOCATIONS
Loc. No. Bldg. No. Address County, Borough or Parish Rating Territory
1 1 1100 North Cherry Street Tulare, CA 93274 012
"""

# ---------------------------------------------------------------- location rows are not forms
rows = fsa.parse_forms_schedule(PACKAGE_TEXT)
nums = [r["form_number"] for r in rows]
check(not any(n.strip().isdigit() for n in nums), f"bare numbers read as forms: {nums}")
check(not any(re.search(r"\b(Street|Blvd|Real)\b", r["description"]) for r in rows), "an address read as a form")
check(len(rows) == 8 and nums[0] == "IL DS 00 09/08", nums)
for n in ("10", "12", "2024", "95823"):
    check(not fsa._is_form_number(n), f"bare number {n} is not a form number")
for n in ("CG 00 01 04 13", "IFG-G-0002-DL 10 25", "HLX-I-0101 10 25", "D-2 01 20", "GLF-105", "IL P 001 01 04"):
    check(fsa._is_form_number(n), f"{n} is a form number")
check(not fsa._row_desc_ok("7216 55th Street, Sacramento, CA, 95823"), "address title")
check(not fsa._row_desc_ok("1100 North Cherry Street Tulare, CA 93274"), "address title, no comma before state")
check(fsa._row_desc_ok("Exclusion - Designated Ongoing Operations"), "a real title passes")

# ---------------------------------------------------------------- OCR-style listing over three pages
rows = fsa.parse_forms_schedule(POLICY_OCR)
nums = [r["form_number"] for r in rows]
check("HLX-G-0302 12 24" in nums and "HLX-G-0322 08 23" in nums,
      f"the two rows alone on the listing's last page are kept: {nums}")
check(any(r["description"] == "Exclusion - Human Trafficking" for r in rows), rows)
check({"HLX-I-0002 05 24", "HLX-I-0101 10 25", "D-2 01 20", "CG 00 01 04 13", "IL P 001 01 04"} <= set(nums), nums)
check(not any(re.search(r"Page \d+ of \d+", r["description"]) for r in rows), "page footers are not forms")
check(len(rows) == 13, f"13 listed forms, got {len(rows)}: {nums}")

# ---------------------------------------------------------------- which file is this coverage's quote
items_scrambled = [{"filename": PACKAGE_FILE, "text": PACKAGE_TEXT * 4},
                   {"filename": POLICY_FILE, "text": POLICY_SCRAMBLED}]
burl = {"carrier": "Harborline Insurance Company", "premium": 42559, "total_premium": 44312.38}
lake = {"carrier": "LAKEVIEW SPECIALTY INSURANCE COMPANY", "premium": 229962, "total_premium": 240774.79}
txt, name = fsa.pick_source_text(items_scrambled, burl["carrier"], "general_liability_alt_1", burl)
check(name == POLICY_FILE, f"the policy that prints the premium, not the keyword-heavy package: {name!r}")
txt, name = fsa.pick_source_text(items_scrambled, burl["carrier"], "general_liability_alt_1",
                                 {"carrier": burl["carrier"]})
check(name != PACKAGE_FILE, f"no figures: a file without the carrier's name is never picked ({name!r})")
txt, name = fsa.pick_source_text(items_scrambled, "Granite Ridge Insurance Company", "general_liability_alt_1",
                                 {"carrier": "Granite Ridge Insurance Company"})
check(name == "", f"carrier named nowhere: no audit ({name!r})")
txt, name = fsa.pick_source_text(items_scrambled, lake["carrier"], "general_liability", lake)
check(name == PACKAGE_FILE, name)
txt, name = fsa.pick_source_text(items_scrambled, lake["carrier"], "general_liability",
                                 dict(lake, _source_file=POLICY_FILE))
check(name == POLICY_FILE, "a recorded source file wins")
items_ocr = [{"filename": PACKAGE_FILE, "text": PACKAGE_TEXT * 4}, {"filename": POLICY_FILE, "text": POLICY_OCR}]
txt, name = fsa.pick_source_text(items_ocr, burl["carrier"], "general_liability_alt_1", {"carrier": burl["carrier"]})
check(name == POLICY_FILE, f"read by OCR, the carrier's name picks the policy ({name!r})")

# ---------------------------------------------------------------- the audit end to end
for items, label in ((items_scrambled, "scrambled"), (items_ocr, "OCR")):
    data = {"coverages": {
        "general_liability": dict(lake, forms_endorsements=[]),
        "general_liability_alt_1": dict(burl, forms_endorsements=[]),
    }, "warnings": []}
    fsa.run_forms_schedule_audit(data, items, pe._forms_audit_reject_factory)
    gl = [f["form_number"] for f in data["coverages"]["general_liability"]["forms_endorsements"]]
    alt = [f["form_number"] for f in data["coverages"]["general_liability_alt_1"]["forms_endorsements"]]
    check(not any(n.strip().isdigit() for n in gl + alt), f"{label}: no location rows as forms")
    check("GLF-105 04/24" in gl and not any(n.startswith("HLX") for n in gl), f"{label}: package forms on the package GL")
    check(not any(n.startswith(("GLF", "FUT", "CG DS")) for n in alt), f"{label}: no package forms on the policy: {alt}")
    if label == "OCR":
        check("HLX-G-0302 12 24" in alt and len(alt) == 13, f"OCR: the policy own 13 forms: {alt}")
    else:
        check(alt == [], f"scrambled: nothing readable, nothing added: {alt}")

# ---------------------------------------------------------------- scrambled-page detector
check(pe._page_scrambled(_JUNK * 10), "junk text is scrambled")
check(not pe._page_scrambled(POLICY_OCR), "OCR text is readable")
county_list = ("Vista, Butler, Calhoun, Carroll, Cass, Cedar, Cerro Gordo, Cherokee, Chickasaw, Clarke, Clay, "
               "Clayton, Clinton, Crawford, Dallas, Davis, Decatur, Delaware, Des Moines, Dickinson, Dubuque, ") * 6
check(not pe._page_scrambled(county_list), "a list of county names is not scrambled")
sov_rows = "LID Street City State Zip County Agg Premium Fee Bldg. Value BPP Value BI Value TIV\n" + (
    "11111111101   500 Ocean View Road,   Del Mar   CA   92014   San Diego   $9,000   $54,000   $8,100   "
    "$14,000,000   $1,000,000   $3,000,000   $18,000,000\n") * 3
check(not pe._page_scrambled(sov_rows), "a schedule of figures with its labels is not scrambled")
check(not pe._page_scrambled("8295 30 5 3 012"), "a nearly empty page is never called scrambled")

# ---------------------------------------------------------------- a scrambled PDF is read by OCR
PAGES = 40
_ocr_text = {
    1: "NOTICE: THE INSURANCE POLICY THAT YOU HAVE PURCHASED IS BEING ISSUED BY AN INSURER THAT IS NOT "
       "LICENSED BY THE STATE OF CALIFORNIA. " * 14,
    5: POLICY_OCR.split("--- Page 6 ---")[0].split("--- Page 5 ---")[1],
    6: POLICY_OCR.split("--- Page 6 ---")[1].split("--- Page 7 ---")[0],
    24: "COMMERCIAL GENERAL LIABILITY DECLARATIONS Item 1. LIMITS OF INSURANCE\n"
        "$2,000,000 General Aggregate Limit\n$1,000,000 Each Occurrence Limit\n$100,000 Damage To Premises Rented To You\n"
        "$5,000 Medical Expense Limit Any One Person\nSCHEDULE OF CLASSES premium basis gross sales\n",
}
for p in range(2, 31):
    _ocr_text.setdefault(p, "SECTION I - COVERAGES. We will pay those sums that the insured becomes legally obligated "
                            "to pay as damages because of bodily injury or property damage to which this insurance "
                            "applies. " * 18)
calls = []


def fake_run(cmd, capture_output=True, text=True, timeout=30):
    if cmd[0] == "pdfinfo":
        return types.SimpleNamespace(returncode=0, stdout=f"Pages:          {PAGES}\n")
    if cmd[0] == "pdftotext":
        return types.SimpleNamespace(returncode=0, stdout=_JUNK * 15)
    raise AssertionError(cmd)


def fake_ocr(pdf_path, pages):
    calls.append(list(pages))
    return {p: (_ocr_text[p], "", "") for p in pages if p in _ocr_text}


real_run, real_ocr, real_has = pe.subprocess.run, pe._ocr_pages, pe.HAS_PDF2IMAGE
pe.subprocess.run, pe._ocr_pages, pe.HAS_PDF2IMAGE = fake_run, fake_ocr, True
pe.os.path.getsize = lambda p: 1000
try:
    pe._OCR_REVIEW_NOTES.clear()
    out = pe.extract_text_from_pdf_smart("/tmp/abcdef12_Tulare GL policy.pdf")
finally:
    pe.subprocess.run, pe._ocr_pages, pe.HAS_PDF2IMAGE = real_run, real_ocr, real_has
check(calls == [list(range(1, 31))], f"pages 1-30 read by OCR in one batch: {calls}")
kept = [int(m) for m in re.findall(r"--- Page (\d+) ---", out)]
check({5, 6, 24} <= set(kept), f"declarations, forms listing and limits kept: {kept}")
check(len(out) <= pe.SCRAMBLED_OCR_MAX_CHARS + 40 * len(kept), f"kept within the scrambled-file budget: {len(out)}")
check("8295 30 5 3" not in out, "no junk text left")
check(kept == sorted(kept), "reading order")
notes = pe._OCR_REVIEW_NOTES.get("Tulare GL policy.pdf") or []
check(notes == [("scrambled", "pages 1-30 of 40 were read by OCR; pages 31-40 were not read")], notes)

print(f"gl_forms_source_test: {PASS} checks passed")
