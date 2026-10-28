"""Regression checks for subjectivity_rules.py (Stefan, Oct 28 2026).

Run from the repo root:  python3 tests/subjectivity_rules_test.py
Self-contained, no OpenAI calls. The quoted items are the carriers' own wording as Stefan listed it.

  * PartnerRe umbrella (Alta Risk program): six broker items left off, three deadlines 30 -> 5 days
  * Starr property: the broker's three surplus lines items left off
  * other umbrellas / other Starr lines untouched; idempotent; the generator never prints them;
    the PartnerRe program template carries the same wording
"""
import copy
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import subjectivity_rules as sr  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


ALTA_REMOVE = [
    "Alta Risk, LLC does not accept, review, reject or approve certificates of insurance. Any requests for "
    "amendments to coverage must be made in writing and approved by Alta Risk, LLC.",
    "When applicable, completed copy of Underlying Policy(ies) including Primary, are required within 60 days. "
    "Failure to provide Alta with requested Underlying Policy information on or before the date specified in "
    "the binder will result in policy cancellation.",
    "Please return a copy of this completed form to us for our files. Thank you.",
    "Net Premiums are due 20 days from Effective Date of this Coverage.",
    "Your office is responsible for all Surplus Lines Filings and Taxes.",
    "Within 60 Days: - Complete copies of ALL Underlying Policies not placed through Alta Risk, LLC",
]
ALTA_FIVE = [
    ("Within 30 Days: - Signed and selected TRIA", "Within 5 Days: - Signed and selected TRIA"),
    ("Within 30 Days: - Completed SLF confirmation", "Within 5 Days: - Completed SLF confirmation"),
    ("Within 30 Days: - Acord 125 and 131 completed and signed by the insured",
     "Within 5 Days: - Acord 125 and 131 completed and signed by the insured"),
]
ALTA_KEEP = ["Prior to Binding: - Signed request to bind"]

STARR_REMOVE = [
    "4. Receipt of the completed and signed Surplus Lines Tax Filing Confirmation form warranting that the "
    "broker/agent will accept full responsibility for compliance of the Surplus Lines laws and the collection "
    "and remittance of the applicable surplus lines tax and/or stamping fees on 100% of the premium, must be "
    "received within 10 days of binding. If the Surplus Lines Filing Confirmation form is not received within "
    "10 days, we reserve the right to cancel the binder and/or any policy issued in connection with the binder.",
    "1) The broker/agent holds the surplus lines license in the state where coverage will be provided to the "
    "Insured and accepts the full obligation to comply with each state's surplus lines laws and regulations in "
    "conjunction with this transaction.",
    "2) The broker/agent accepts full responsibility for compliance, including but not limited to, the filing "
    "of the surplus lines affidavit (as per the state's requirements), and the collection and remittance of "
    "the surplus lines tax and any applicable stamping fee on 100% of the premium of this policy. The "
    "broker/agent's acceptance of these requirements is to be confirmed by completing, signing and dating the "
    "attached Surplus Lines Filings Confirmation form.",
]
STARR_KEEP = [
    "Written Request to Bind Coverage",
    "Signed TRIA Election form",
    "Broker will provide licensing information which will be verified prior to binding.",
    "TRIA Forms due within 15 days of binding or Insured will be billed for Terrorism",
]
OTHER_UMB = ["Premium is due no later than twenty (20) days from inception.",
             "Within 30 Days: - Signed and selected TRIA"]


def fresh():
    return {"coverages": {
        "umbrella": {"carrier": "PartnerRe Insurance Solutions Bermuda Ltd.",
                     "subjectivities": ALTA_REMOVE[:3] + [a for a, _ in ALTA_FIVE] + ALTA_REMOVE[3:] + ALTA_KEEP},
        "property": {"carrier": "Starr Surplus Lines Insurance Company",
                     "subjectivities": ["☐  " + STARR_KEEP[0]] + STARR_REMOVE + STARR_KEEP[1:]},
        "umbrella_layer_2": {"carrier": "StarStone National Insurance Company", "subjectivities": list(OTHER_UMB)},
        "general_liability": {"carrier": "Starr Indemnity & Liability Company", "subjectivities": list(STARR_REMOVE)},
    }}


# 1. PartnerRe umbrella
d = fresh()
notes = sr.run_subjectivity_rules(d)
u = d["coverages"]["umbrella"]["subjectivities"]
check(u == [b for _, b in ALTA_FIVE] + ALTA_KEEP, f"PartnerRe list: {u}")
check(any("PartnerRe umbrella" in n and "6 broker items" in n and "3 deadlines" in n for n in notes), notes)

# 2. Starr property
p = d["coverages"]["property"]["subjectivities"]
check(p == ["☐  " + STARR_KEEP[0]] + STARR_KEEP[1:], f"Starr list: {p}")
check(any("Starr property" in n and "3 broker items" in n for n in notes), notes)

# 3. Untouched: another umbrella carrier, Starr on a liability line
check(d["coverages"]["umbrella_layer_2"]["subjectivities"] == OTHER_UMB, "Great Point / StarStone untouched")
check(d["coverages"]["general_liability"]["subjectivities"] == STARR_REMOVE, "Starr GL untouched")

# 4. Alta quote extracted without the carrier name still matches; dict rows keep their shape
d2 = {"coverages": {"umbrella_layer_3": {"carrier": "",
      "subjectivities": [{"description": ALTA_REMOVE[0]}, {"description": ALTA_FIVE[0][0]}, ALTA_REMOVE[4]]}}}
sr.run_subjectivity_rules(d2)
check(d2["coverages"]["umbrella_layer_3"]["subjectivities"] == [{"description": ALTA_FIVE[0][1]}],
      d2["coverages"]["umbrella_layer_3"]["subjectivities"])

# 5. Idempotent
d3 = fresh()
sr.run_subjectivity_rules(d3)
snap = copy.deepcopy(d3["coverages"])
check(sr.run_subjectivity_rules(d3) == [] and d3["coverages"] == snap, "second run changes nothing")

# 6. Odd shapes never raise
for odd in ({"coverages": None}, {"coverages": {"umbrella": None}}, {"coverages": {"umbrella": {"subjectivities": None}}},
            {"coverages": {"umbrella": {"carrier": "PartnerRe", "subjectivities": "Within 30 Days: - Signed and selected TRIA"}}},
            {"coverages": {"property": {"carrier": "Starr", "subjectivities": [None, 5, ""]}}}):
    sr.run_subjectivity_rules(odd)
check(True, "odd shapes")

# 7. Generator: an un-cleaned session never prints the items
import docx  # noqa: E402
import proposal_generator as pg  # noqa: E402

gen = fresh()
gen["client_info"] = {"named_insured": "Sample Hotel LLC"}
for k, c in gen["coverages"].items():
    c.update(premium=1000, total_premium=1000)
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "s.docx")
    pg.generate_proposal(gen, out)
    text = "\n".join(p.text for p in docx.Document(out).paragraphs)
subj = text.split("Binding Subjectivities", 1)[1]
prop_sec = subj.split("Property —", 1)[1].split("General Liability —", 1)[0]
umb_sec = subj.split("Umbrella / Excess 1 —", 1)[1].split("2nd Excess Layer —", 1)[0]
for item in ALTA_REMOVE:
    check(item[:60] not in text, f"PartnerRe item not printed anywhere: {item[:60]}")
for item in STARR_REMOVE:
    check(item[3:60] not in prop_sec, f"Starr item not printed under Property: {item[:60]}")
check("Within 5 Days: - Signed and selected TRIA" in umb_sec and "Within 30 Days" not in umb_sec,
      "PartnerRe section shows 5 days")
check("Prior to Binding: - Signed request to bind" in umb_sec, "PartnerRe item Stefan did not list is kept")
check(STARR_REMOVE[1][3:60] in subj.split("General Liability —", 1)[1], "Starr GL keeps its items")
check("Within 5 Days: - Signed and selected TRIA" in text, "5-day TRIA printed for PartnerRe")
check("Within 30 Days: - Signed and selected TRIA" in text, "other umbrella keeps its own 30 days")
check("Written Request to Bind Coverage" in text and "Broker will provide licensing information" in text,
      "Starr items Stefan did not list are kept")

# 8. PartnerRe program template wording follows the same rules
import program_templates as pt  # noqa: E402

check(all("5 days" in s for s in pt._PARTNERRE_SUBJECTIVITIES), pt._PARTNERRE_SUBJECTIVITIES)
check(not any("underlying polic" in s.lower() for s in pt._PARTNERRE_SUBJECTIVITIES), "no 60-day underlying item")
check(not any("certificates" in t.lower() or "underlying polic" in t.lower() or "20 days" in t
              for t in pt._PARTNERRE_TERMS), pt._PARTNERRE_TERMS)
tpl = {"program_template": "partnerre_umbrella", "carrier": "PartnerRe Insurance Solutions Bermuda Ltd.",
       "subjectivities": list(pt._PARTNERRE_SUBJECTIVITIES), "terms_conditions": list(pt._PARTNERRE_TERMS)}
check(sr.clean_coverage("umbrella", copy.deepcopy(tpl)) == ([], []), "template already clean")

print(f"subjectivity_rules_test: {PASS} checks passed")
