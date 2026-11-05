"""Regression checks: a property quote's own ISO "IL" interline forms stay on the proposal (Nov 5 2026).

Run from the repo root:  python3 tests/property_interline_forms_test.py
No OpenAI calls.

Why: Harp Hotels Inc's Tower Hill / Lloyd's property quote (QT0000152418) prints 35 forms on its
Form Schedule. The property forms cleaner drops every IL form (GL quotes carry IL forms that GPT
copies into property) and anything titled "Common Policy Conditions", so the proposal listed 25:
War and Civil War, Microorganism, Asbestos, Radioactive Contamination, Computer-Related Losses
and Electromagnetic Storm exclusions were among the 10 missing.
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import proposal_extractor as pe  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


SCHEDULE = [
    ("CP 00 10 (06/07)", "Building and Personal Property Coverage Form"),
    ("CP 10 30 (06/07)", "Causes of Loss - Special Form"),
    ("CP 10 32 (08/08)", "Water Exclusion Endorsement"),
    ("CP 12 11 (10/00)", "Burglary and Robbery Protective Safeguards"),
    ("CP 14 10 (06/95)", "Additional Covered Property"),
    ("EC 00 01 (01/17)", "Table of Contents - Building and Personal Property"),
    ("EC 00 17 (01/17)", "Common Policy Conditions - Cancellation and Nonrenewal"),
    ("EC 01 25 (05/23)", "Florida Changes"),
    ("EC 04 01 (01/17)", "Florida - Sinkhole Loss Coverage"),
    ("IL 09 35 (07/02)", "Exclusion of Certain Computer-Related Losses"),
    ("TC 00 10 (11/13)", "Existing Damage Exclusion Endorsement"),
    ("TC 00 13 (11/13)", "Fire Code Compliance Exclusion"),
    ("TC 00 30 (10/18)", "Business Income (and Extra Expense) Coverage Form"),
    ("TC 00 90 (06/21)", "Commercial Property Conditions"),
    ("TC 01 00 (05/23)", "Property Enhancement Endorsement"),
    ("TC 01 12 (05/23)", "Florida Changes - Mediation or Appraisal and Neutral Evaluation"),
    ("TC 14 20 (11/16)", "Additional Property Not Covered"),
    ("TC 15 12 (08/16)", "Exclusion Earth Movement"),
    ("CP 01 40 (07/06)", "Exclusion of Loss Due to Virus or Bacteria"),
    ("EC 03 25 (05/24)", "Named Storm Deductible"),
    ("IL 03 70 (11/03)", "Institute Radioactive Contamination, Chemical, Biological, Bio-Chemical and "
                         "Electromagnetic Weapons Exclusion Clause"),
    ("IL 04 15 (04/98)", "Protective Safeguards"),
    ("IL 04 64 (01/38)", "War and Civil War Exclusion Clause"),
    ("IL 50 18 (09/05)", "Microorganism Exclusion (Absolute)"),
    ("IL 50 19 (09/05)", "Asbestos Endorsement"),
    ("IL 55 10 (05/23)", "Electromagnetic / Geomagnetic Storm Exclusion"),
    ("TC 15 36 (01/16)", "Hotel/Motel Establishment - Amendatory Endorsement"),
    ("IL 09 52 (01/15)", "Cap on Losses from Certified Acts of Terrorism"),
    ("EB-0006 (07/21)", "Front Page Form"),
    ("EC 00 12 (01/17)", "Standard Policy Coverage - Declarations"),
    ("IL 09 85 (12/20)", "Disclosure Pursuant to Terrorism Risk Insurance Act"),
    ("IL P 001 (01/04)", "OFAC Advisory Notice"),
    ("IL-FEES (01/17)", "Important Notice About Fees"),
    ("LLO-0001-HO (04/25)", "Lloyd's Certificate"),
    ("Privacy Notice (07/25)", "Privacy Notice"),
]
WAS_MISSING = {"EC 00 17 (01/17)", "IL 09 35 (07/02)", "IL 03 70 (11/03)", "IL 04 15 (04/98)", "IL 04 64 (01/38)",
               "IL 50 18 (09/05)", "IL 50 19 (09/05)", "IL 55 10 (05/23)", "IL 09 52 (01/15)", "IL 09 85 (12/20)"}

PROPERTY_TEXT = (
    "--- Page 2 ---\nThis proposal is a summary of coverage offered by Underwriters at Lloyds, London\n"
    "--- Page 3 ---\nCoverage Summary\nInsured Name: Harp Hotels Inc\nCoverage Form - Commercial Property\n"
    "Total Premium $37,243.15\n"
    "--- Page 4 ---\nPREMISES AND BUILDINGS\nNamed Storm: 5%\nValuation Basis: Replacement Cost\nCoinsurance: 80%\n"
    "1 7900 SW LOST RIVER RD STUART FL 34997 $7,250,000\nBuilding Personal Property $300,000\n"
    "Business Income with Extra Expense Coverage - 1/4 $400,000\n"
    "--- Page 9 ---\nForm Schedule\nForm Description\n"
    + "\n".join(f"{n} {d}" for n, d in SCHEDULE)
)
GL_TEXT = (
    "Berkley Specialty Insurance Company\nCommercial General Liability Coverage Part\n"
    "Commercial General Liability Coverage Form CG 00 01 04 13\n"
    "Nuclear Energy Liability Exclusion Endorsement IL 00 21 09 08\n"
    "Florida Changes - Cancellation and Nonrenewal IL 02 55 03 16\n"
    "Disclosure Pursuant to Terrorism Risk Insurance Act IL 09 85 12 20\n"
    "Employment-Related Practices Exclusion CG 21 47 12 07\n" * 3
)
ITEMS = [{"filename": "Quote_Letter_for_Harp_Hotels_Inc.pdf", "text": PROPERTY_TEXT},
         {"filename": "Berkley GL Quote.pdf", "text": GL_TEXT}]


def fresh_forms(extra=()):
    return [{"form_number": n, "description": d} for n, d in list(SCHEDULE) + list(extra)]


# 1. the cleaner alone loses the 10 forms (the bug)
orig = fresh_forms()
cleaned = pe._clean_property_forms_endorsements(orig)
lost = {f["form_number"] for f in orig} - {f["form_number"] for f in cleaned}
check(lost == WAS_MISSING, sorted(lost))

# 2. the property quote file is the one picked
from forms_schedule_audit import pick_source_text  # noqa: E402
_t, _name = pick_source_text(ITEMS, "Underwriters at Lloyds, London", "property")
check(_name == "Quote_Letter_for_Harp_Hotels_Inc.pdf", _name)
check(pick_source_text(ITEMS, "Berkley Specialty Insurance Company", "general_liability")[1] == "Berkley GL Quote.pdf",
      "GL pick unchanged")

# 3. restore: all 35 back, in the quote's order, restored rows flagged
cov = {"carrier": "Underwriters at Lloyds, London", "forms_endorsements": orig}
forms, back, src = pe._restore_property_quote_forms("property", cov, orig, cleaned, ITEMS)
check([f["form_number"] for f in forms] == [n for n, _d in SCHEDULE], [f["form_number"] for f in forms])
check({b["form_number"] for b in back} == WAS_MISSING and all(b.get("schedule_verified") for b in back), back)
check(src == "Quote_Letter_for_Harp_Hotels_Inc.pdf", src)
check(not any(f.get("schedule_verified") for f in forms if f["form_number"] not in WAS_MISSING), "others untouched")

# 4. a second cleaner pass keeps them (HotelBound override re-cleans property)
check(len(pe._clean_property_forms_endorsements(forms)) == 35, "re-clean keeps verified rows")

# 5. GL forms GPT copied into property stay out: liability prefix, "liability" in the title,
#    or an IL form the property quote never prints
bleed = [("CG 21 47 12 07", "Employment-Related Practices Exclusion"),
         ("IL 00 21 09 08", "Nuclear Energy Liability Exclusion Endorsement"),
         ("IL 02 55 03 16", "Florida Changes - Cancellation and Nonrenewal")]
orig2 = fresh_forms(bleed)
cleaned2 = pe._clean_property_forms_endorsements(orig2)
forms2, back2, _s = pe._restore_property_quote_forms("property", cov, orig2, cleaned2, ITEMS)
nums2 = {f["form_number"] for f in forms2}
check(len(forms2) == 35 and not ({n for n, _d in bleed} & nums2), sorted(nums2 - {n for n, _d in SCHEDULE}))

# 6. no property quote among the files (SOV only) -> nothing restored, no error
f3, b3, s3 = pe._restore_property_quote_forms("property", cov, orig, cleaned,
                                                [{"filename": "SOV.xlsx", "text": "Statement of Values TIV"}])
check(b3 == [] and len(f3) == 25, (len(f3), b3))

# 7. the extraction pass is wired: Patch M restores and reports
src_code = open(pe.__file__, encoding="utf-8").read()
check("_restore_property_quote_forms(_pkey, _pcov, _f, _c, all_items)" in src_code, "Patch M calls the restore")
check("kept {len(_back)} interline forms printed on" in src_code, "review screen note")

# 8. the generator prints them in the property forms table
import docx  # noqa: E402
import proposal_generator as pg  # noqa: E402

gen = {"client_info": {"named_insured": "Harp Hotels Inc"},
       "coverages": {"property": {"carrier": "Underwriters at Lloyds, London", "premium": 35176,
                                  "total_premium": 37243.15, "forms_endorsements": forms}}}
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "p.docx")
    pg.generate_proposal(gen, out)
    d = docx.Document(out)
    cells = " | ".join(c.text for t in d.tables for r in t.rows for c in r.cells)
for n in ("IL 04 64", "IL 50 18", "IL 50 19", "IL 55 10", "IL 03 70", "IL 09 35", "EC 00 17"):
    check(n in cells, f"{n} printed in the proposal")

print(f"property_interline_forms_test: {PASS} checks passed")
