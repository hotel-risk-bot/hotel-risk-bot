"""Regression checks for an excess policy that sits over the liquor liability alone.

Run from the repo root:  python3 tests/excess_liquor_test.py
No OpenAI calls.

Why (Stefan, Nov 24 2026, WG Hotel LLC): "here are the excess, one is liquor liability excess,
other is to build the tower". Two excess quotes: one $3,000,000 excess of the General Liability
(the umbrella tower) and one $3,000,000 "Monoline Excess Liquor Liability" excess of the liquor
liability only, for the two hotels that serve alcohol. The model returned both in the umbrella
slots, so the liquor policy printed as "2nd Excess Layer" of a tower it is not part of.

  * excess_liquor.is_excess_liquor / run_excess_liquor - the policy moves to its own line,
    "excess_liquor_liability", and the umbrella layers close up
  * proposal_extractor - the pass runs with the other quote detail passes
  * proposal_generator - its own Premium Summary row and section, the hotels it applies to, and
    it adds nothing to the Total Umbrella / Excess Limit
  * client_email, the review screen - the line's name
"""
import copy
import logging
import os
import re
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ.setdefault("OPENAI_API_KEY", "test")
logging.disable(logging.CRITICAL)

import docx  # noqa: E402
from docx.table import Table  # noqa: E402
from docx.text.paragraph import Paragraph  # noqa: E402
import client_email  # noqa: E402
import excess_liquor as xl  # noqa: E402
import proposal_generator as pg  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


TOWER = {"carrier": "Summit Re Insurance Solutions Ltd.", "premium": 79588, "total_premium": 80088,
         "layer_description": "$3,000,000 Excess of Scheduled Policy",
         "limits": [{"description": "Each Occurrence", "limit": "$3,000,000"},
                    {"description": "Aggregate", "limit": "$3,000,000"}],
         "underlying_insurance": [{"carrier": "Lakeshore Specialty Insurance Company", "coverage": "General Liability",
                                   "limits": "$1,000,000 Each Occurrence / $2,000,000 Aggregate"},
                                  {"carrier": "Lakeshore Specialty Insurance Company",
                                   "coverage": "Hired and Non-Owned Auto", "limits": "$1,000,000 Combined Single Limit"}]}
LIQUOR_XS = {"carrier": "Northfield Star Indemnity Company", "premium": 8925, "total_premium": 8925,
             "layer_description": "$3,000,000 in excess of underlying Liquor Liability",
             "limits": [{"description": "Each Occurrence", "limit": "$3,000,000"},
                        {"description": "Aggregate", "limit": "$3,000,000"}],
             "designated_premises": ["615 North Lancaster Pike, Exton, PA 19341", "5426 Lower York Road, New Hope, PA 18938"],
             "underlying_insurance": [{"carrier": "Lakeshore Specialty Insurance Company", "coverage": "Liquor Liability",
                                       "limits": "$1,000,000 Each Common Cause / $1,000,000 Aggregate"}]}
GL = {"carrier": "Lakeshore Specialty Insurance Company", "premium": 200000, "total_premium": 209000,
      "limits": [{"description": "Each Occurrence", "limit": "$1,000,000"},
                 {"description": "General Aggregate", "limit": "$2,000,000"}]}


def make(**slots):
    covs = {"general_liability": copy.deepcopy(GL)}
    for k, v in slots.items():
        covs[k] = copy.deepcopy(v)
    return {"client_info": {"named_insured": "Keystone Lodging Group LLC", "effective_date": "12/01/2026"},
            "coverages": covs,
            "payment_options": [
                {"carrier": "Summit Re Insurance Solutions Ltd.", "coverage_type": "Umbrella", "terms": "Annual"},
                {"carrier": "Northfield Star Indemnity Company", "coverage_type": "Excess Liability", "terms": "Annual"}]}


# ------------------------------------------------------------------ 1. which policies count
check(xl.is_excess_liquor(LIQUOR_XS), "underlying schedule lists liquor liability only")
check(not xl.is_excess_liquor(TOWER), "excess of the GL and auto is the umbrella tower")
check(not xl.is_excess_liquor(dict(TOWER, underlying_insurance=TOWER["underlying_insurance"] + LIQUOR_XS["underlying_insurance"])),
      "an umbrella over GL, auto and liquor liability stays an umbrella")
check(not xl.is_excess_liquor({"carrier": "X", "underlying_insurance": [{"coverage": "General Liability / Liquor Liability"}]}),
      "one row naming both lines")
check(xl.is_excess_liquor({"carrier": "X", "layer_description": "Monoline Excess Liquor Liability"}), "named in the layer")
check(xl.is_excess_liquor({"carrier": "X", "policy_form": "Excess Liquor Liability Coverage Form"}), "named in the form")
check(not xl.is_excess_liquor({"carrier": "X", "layer_description": "$5,000,000 Excess of General Liability, Auto and Liquor Liability"}),
      "no underlying schedule, several lines named")
check(not xl.is_excess_liquor({"carrier": "X", "layer_description": "$5,000,000 Excess of Underlying"}), "says nothing of liquor")
check(not xl.is_excess_liquor({"carrier": "X"}) and not xl.is_excess_liquor(None) and not xl.is_excess_liquor("x"), "empty")
check(xl.is_excess_liquor({"carrier": "X", "underlying_insurance": ["Liquor Liability - $1,000,000"]}), "string rows")
check(xl.is_excess_liquor({"carrier": "X", "underlying_insurance": [{"type": "Dram Shop Liability"}]}), "dram shop")

# ------------------------------------------------------------------ 2. second layer -> its own line
d = make(umbrella=TOWER, umbrella_layer_2=LIQUOR_XS)
notes = xl.run_excess_liquor(d)
covs = d["coverages"]
check("umbrella_layer_2" not in covs, list(covs))
check(covs["excess_liquor_liability"]["carrier"] == "Northfield Star Indemnity Company", covs.get("excess_liquor_liability"))
check(covs["umbrella"]["carrier"] == "Summit Re Insurance Solutions Ltd.", "the tower keeps its place")
check(len(notes) == 1 and notes[0].startswith("Excess check — Northfield Star Indemnity Company ($3,000,000 excess of the "
                                                "liquor liability) is shown as Excess Liquor Liability"), notes)
check(d["warnings"] == notes and "--" not in notes[0], d["warnings"])
check([p["coverage_type"] for p in d["payment_options"]] == ["Umbrella", "Excess Liquor Liability"], d["payment_options"])
# again: nothing more to move, no second note
check(xl.run_excess_liquor(d) == [] and d["warnings"] == notes, "idempotent")

# ------------------------------------------------------------------ 3. listed first -> the tower moves up
d = make(umbrella=LIQUOR_XS, umbrella_layer_2=TOWER, umbrella_layer_3=dict(TOWER, carrier="Third Layer Insurance Company"))
order_before = list(d["coverages"])
xl.run_excess_liquor(d)
covs = d["coverages"]
check(covs["umbrella"]["carrier"] == "Summit Re Insurance Solutions Ltd.", "tower is the first layer")
check(covs["umbrella_layer_2"]["carrier"] == "Third Layer Insurance Company", "third layer moves up to second")
check("umbrella_layer_3" not in covs, list(covs))
check(covs["excess_liquor_liability"]["carrier"] == "Northfield Star Indemnity Company", list(covs))
check(list(covs)[0] == order_before[0] == "general_liability", "other lines keep their order")

# ------------------------------------------------------------------ 4. alias slots, and what stays put
d = make(umbrella=TOWER, excess_liability=LIQUOR_XS)
xl.run_excess_liquor(d)
check("excess_liability" not in d["coverages"] and "excess_liquor_liability" in d["coverages"], list(d["coverages"]))
d = make(umbrella=TOWER, umbrella_layer_2=dict(TOWER, carrier="Second Layer Insurance Company"))
before = copy.deepcopy(d)
check(xl.run_excess_liquor(d) == [] and d == before, "a real two-layer tower is untouched")
d = make(liquor_liability={"carrier": "Lakeshore Specialty Insurance Company", "premium": 5000,
                           "limits": [{"description": "Each Common Cause", "limit": "$1,000,000"}]})
before = copy.deepcopy(d)
check(xl.run_excess_liquor(d) == [] and d == before, "the primary liquor liability is not an excess policy")
check(xl.run_excess_liquor({}) == [] and xl.run_excess_liquor({"coverages": None}) == [], "no coverages")
# two excess liquor policies: the first takes the line, the second is named for review
d = make(umbrella=TOWER, umbrella_layer_2=LIQUOR_XS, umbrella_layer_3=dict(LIQUOR_XS, carrier="Other Liquor Excess Company"))
n = xl.run_excess_liquor(d)
check(d["coverages"]["excess_liquor_liability"]["carrier"] == "Northfield Star Indemnity Company", list(d["coverages"]))
check(any("second excess liquor policy" in x and "Other Liquor Excess Company" in x for x in n), n)

# ------------------------------------------------------------------ 5. the extraction runs it
src = open(os.path.join(ROOT, "proposal_extractor.py"), encoding="utf-8").read()
check('("excess_liquor", "run_excess_liquor")' in src, "pass registered with the quote detail passes")
check(src.index('("excess_liquor", "run_excess_liquor")') < src.index('("subjectivity_rules"'), "before the subjectivity rules")

# ------------------------------------------------------------------ 6. the proposal (a session extracted before the change)
d = make(umbrella=TOWER, umbrella_layer_2=LIQUOR_XS)
d["coverages"]["property"] = {"carrier": "Starr Surplus Lines Insurance Company", "premium": 100000, "total_premium": 104000}
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "p.docx")
    pg.generate_proposal(copy.deepcopy(d), out)
    doc = docx.Document(out)
    lines = []
    for el in doc.element.body.iterchildren():
        if el.tag.endswith("}p"):
            t = Paragraph(el, doc).text.strip()
            if t:
                lines.append(("P", [t]))
        elif el.tag.endswith("}tbl"):
            for r in Table(el, doc).rows:
                lines.append(("T", [c.text.strip() for c in r.cells]))
rows = [c for k, c in lines if k == "T"]
paras = [c[0] for k, c in lines if k == "P"]
text = "\n".join(" | ".join(c) for _k, c in lines)

check(["Excess Liquor Liability", "Northfield Star Indemnity Company", "$8,925.00"] in rows, "Premium Summary row")
check(["Umbrella / Excess 1", "Summit Re Insurance Solutions Ltd.", "$80,088.00"] in rows, "tower row")
check("2nd Excess Layer" not in text and "2nd Excess Liability Layer" not in text, "not a layer of the tower")
total = next(r for r in rows if r[0] == "Total" and len(r) == 3)
check(total[-1] == "$402,013.00", total)                       # 104,000 + 209,000 + 80,088 + 8,925
check(["Total Umbrella / Excess Limit", "$3,000,000"] in rows, [r for r in rows if "Total Umbrella" in r[0]])
check("Excess Liquor Liability Coverage" in paras, "its own section")
i = next(n for n, (k, c) in enumerate(lines) if c == ["Excess Liquor Liability Coverage"])
nxt = next((n for n, (k, c) in enumerate(lines) if n > i and k == "P" and c[0] in
            ("Coverage Recommendations", "Confirmation to Bind Agreement")), len(lines))
section = lines[i:nxt]
sec_rows = [c for k, c in section if k == "T"]
check(["Carrier", "Northfield Star Indemnity Company"] in sec_rows, sec_rows[:4])
check(["Lakeshore Specialty Insurance Company", "Liquor Liability", "$1,000,000 Each Common Cause / $1,000,000 Aggregate"]
      in sec_rows, "underlying liquor liability")
check(["1", "615 North Lancaster Pike, Exton, PA 19341"] in sec_rows and ["2", "5426 Lower York Road, New Hope, PA 18938"] in sec_rows,
      "the two hotels it applies to")
check(any(k == "P" and c[0].startswith("This Excess Liquor Liability policy applies only to the following locations")
          for k, c in section), "says it is limited to those hotels")
check(paras.index("Excess Liquor Liability Coverage") > paras.index("Umbrella / Excess Liability Coverage"), "after the umbrella")
check(any(r[0] == "Northfield Star Indemnity Company" and r[-1] == "Excess Liquor Liability" for r in rows), "carrier ratings row")
check("--" not in "\n".join(paras[paras.index("Excess Liquor Liability Coverage"):nxt]), "no double dashes")

# ------------------------------------------------------------------ 7. the line's name elsewhere
check(client_email.COVERAGE_LABELS.get("excess_liquor_liability") == "Excess Liquor Liability", "client email label")
order = client_email._LINE_ORDER
check(order.index("liquor_liability") < order.index("excess_liquor_liability") < order.index("umbrella"), order)
html = open(os.path.join(ROOT, "templates", "proposal_web.html"), encoding="utf-8").read()
check("excess_liquor_liability: 'Excess Liquor Liability'" in html, "review screen name")
check('<option value="excess_liquor_liability">Excess Liquor Liability</option>' in html, "Add Coverage choice")
web = open(os.path.join(ROOT, "web_app.py"), encoding="utf-8").read()
check('"excess_liquor_liability": "Excess Liquor Liability"' in web, "extraction summary name")
gen = open(os.path.join(ROOT, "proposal_generator.py"), encoding="utf-8").read()
check(len(re.findall(r'"excess_liquor_liability": "Excess Liquor Liability"', gen)) == 3, "three label tables in the generator")

print(f"excess_liquor_test: {PASS} checks passed")
