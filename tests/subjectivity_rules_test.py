"""Regression checks for subjectivity_rules.py (Stefan, Oct 28 2026).

Run from the repo root:  python3 tests/subjectivity_rules_test.py
Self-contained, no OpenAI calls. The quoted items are the carriers' own wording as Stefan listed it.

  * PartnerRe umbrella (Alta Risk program): six broker items left off, three deadlines 30 -> 5 days
  * Starr property: the broker's three surplus lines items left off
  * other umbrellas / other Starr lines untouched; idempotent; the generator never prints them;
    the PartnerRe program template carries the same wording
  * every carrier (Nov 24 2026): the surplus lines licensee warranty and "quotes are valid for
    30 days" left off Binding Subjectivities; Terms & Conditions keep their validity line
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

# 9. Every carrier (Stefan, Nov 24 2026): the surplus lines licensee warranty and the 30-day quote
#    validity line are left off; a validity sentence on a real item drops and the rest stays.
SL_LICENSEE = ("With your signature, you hereby warrant and represent that the surplus lines licensee indicated "
               "above is responsible for the collection and remittance of the surplus lines taxes, stamping fees "
               "and/or other charges in connection with the surplus lines placement of this policy.")
VALID_30 = ["Quotes are valid for 30 days.",
            "This quote is valid for thirty (30) days from the date of issue.",
            "Quotation valid for 30 days",
            "Valid for 30 days from the quote date",
            "Terms quoted are good for 30 days or until the effective date, whichever is earlier."]
GEN_KEEP = ["Signed and dated TRIA election form within 30 days of binding",
            "Surplus lines licensee name and license number prior to binding",
            "Premium is due no later than twenty (20) days from inception.",
            "Coverage cannot be bound without a signed application.",
            "Signed TRIA election required; the TRIA offer is valid for 30 days from binding."]
MIXED = "No known or reported losses since the quote date. Quotes are valid for 30 days."


def gen_fresh():
    return {"coverages": {
        "property": {"carrier": "Certain Underwriters at Lloyd's, London",
                     "subjectivities": ["☐  " + SL_LICENSEE, VALID_30[0], GEN_KEEP[0], MIXED, GEN_KEEP[1]]},
        "general_liability": {"carrier": "Southlake Specialty Insurance Company",
                              "subjectivities": [{"description": "4. " + SL_LICENSEE}] + VALID_30[1:] + GEN_KEEP[2:]
                              + [{"description": "Signed application; quotes are valid for 30 days."}]},
        "umbrella": {"carrier": "PartnerRe Insurance Solutions Bermuda Ltd.",
                     "subjectivities": ALTA_REMOVE + [a for a, _ in ALTA_FIVE] + [VALID_30[0]] + ALTA_KEEP},
    }}


g = gen_fresh()
gnotes = sr.run_subjectivity_rules(g)
gp = g["coverages"]["property"]["subjectivities"]
check(gp == [GEN_KEEP[0], "No known or reported losses since the quote date.", GEN_KEEP[1]], f"Lloyd's property: {gp}")
gl = g["coverages"]["general_liability"]["subjectivities"]
check(gl == GEN_KEEP[2:] + [{"description": "Signed application"}], f"Southlake GL: {gl}")
gu = g["coverages"]["umbrella"]["subjectivities"]
check(gu == [b for _, b in ALTA_FIVE] + ALTA_KEEP, f"PartnerRe also loses the validity line: {gu}")
check(any("PartnerRe umbrella" in n and "6 broker items" in n and "3 deadlines" in n for n in gnotes), gnotes)
check(any(n.startswith("Subjectivities: left off 10 items (surplus lines licensee warranty, quotes valid for 30 days)")
          for n in gnotes), gnotes)
check(sr.general_rule(SL_LICENSEE) == "surplus lines licensee warranty", "Stefan's exact wording")
check(all(sr.general_rule(v) == "quotes valid for 30 days" for v in VALID_30), "validity rewordings")
check(not any(sr.general_rule(k) for k in GEN_KEEP + OTHER_UMB + STARR_KEEP + ALTA_KEEP), "real items kept")
snap = copy.deepcopy(g["coverages"])
check(sr.run_subjectivity_rules(g) == [] and g["coverages"] == snap, "general rules idempotent")

# 10. Terms & Conditions keep the template's validity line (Binding Subjectivities only)
tpl2 = copy.deepcopy(tpl)
sr.clean_coverage("umbrella", tpl2)
check("Quotation is valid for 30 days from the quote date." in tpl2["terms_conditions"], tpl2["terms_conditions"])

# 11. Generator: an un-cleaned session never prints them; single-string / single-dict lists print whole
gen2 = gen_fresh()
gen2["client_info"] = {"named_insured": "Sample Hotel LLC"}
gen2["coverages"]["crime"] = {"carrier": "Travelers", "subjectivities": "Signed crime application prior to binding"}
gen2["coverages"]["cyber"] = {"carrier": "Coalition",
                              "subjectivities": {"description": "Signed cyber application. Quotes are valid for 30 days."}}
gen2["coverages"]["umbrella_layer_2"] = dict(tpl, carrier="PartnerRe Insurance Solutions Bermuda Ltd.")
for k, c in gen2["coverages"].items():
    c.update(premium=1000, total_premium=1000)
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "s2.docx")
    pg.generate_proposal(gen2, out)
    d2doc = docx.Document(out)
paras = [p.text for p in d2doc.paragraphs]
start = max(i for i, t in enumerate(paras) if t.strip() == "Binding Subjectivities")
end = next((i for i in range(start + 1, len(paras)) if paras[i].strip() == "Named Insureds"), len(paras))
subj2 = "\n".join(paras[start:end])
check("warrant and represent that the surplus lines licensee" not in subj2.lower(), "licensee warranty not printed")
for v in VALID_30 + ["Quotes are valid for 30 days", "quotes are valid for 30 days"]:
    check(v.rstrip(".") not in subj2, f"validity line not printed: {v}")
for k in GEN_KEEP:
    check("☐  " + k in subj2, f"kept item printed: {k}")
check("☐  No known or reported losses since the quote date." in subj2, "rest of a mixed item printed")
check("☐  Signed application" in subj2, "rest of a dict item printed")
check("☐  Signed crime application prior to binding" in subj2, "a single-string list prints as one line")
check("☐  Signed cyber application." in subj2, "a single-dict list prints as one line")
cells = "\n".join(c.text for t in d2doc.tables for r in t.rows for c in r.cells)
check("Quotation is valid for 30 days from the quote date." in cells, "template Terms & Conditions keep it")

print(f"subjectivity_rules_test: {PASS} checks passed")
