"""Regression checks for the client email subject and exclusion bullets (Stefan, Nov 4 2026).

Run from the repo root:  python3 tests/client_email_bullets_test.py
No OpenAI calls: the model is replaced by a fake that returns the Harp Hotels draft Stefan saw
(subject without the corporate name, GL and umbrella exclusions merged in one bullet,
"Terrorism/Tria").
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("OPENAI_API_KEY", "test")

import client_email as ce  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


MODEL_DRAFT = """SUBJECT: Holiday Inn Express Renewal Proposal
BODY:
Hi Hiren,

Attached, is the finalized proposal for Holiday Inn Express with a $12,366.32 decrease, or 17.9%, with $7,950,000 in total insured value.

Proposed carriers are Lloyd's for Property, Berkley Specialty for General Liability, and StarStone National for Umbrella/Excess.

The key points are:
- Property deductibles are $25,000 for All Other Perils, 5% for Named Storm, and 5% for Sinkhole.
- General Liability deductibles are $2,500 per claim.
- Please note the General Liability policy excludes Employment-Related Practices, Total Pollution, and Total Liquor Liability, and the Umbrella/Excess excludes Abuse or Molestation, Communicable Disease, Human Trafficking, and No Coverage for Sublimits.
- Equipment Breakdown, Terrorism/Tria, and Pollution are recommended but not included in the total premium and have not yet been quoted. Supplemental applications may be needed to finalize quotes for the optional coverages.

Let me know if you would like to walk through the proposal or if you have any questions.

Stefan"""


class _FakeClient:
    def __init__(self, *a, **k):
        self.chat = types.SimpleNamespace(completions=types.SimpleNamespace(create=self._create))
        self.calls = 0

    def _create(self, **kwargs):
        _FakeClient.last_messages = kwargs["messages"]
        msg = types.SimpleNamespace(content=_FakeClient.reply)
        return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)])


_FakeClient.reply = MODEL_DRAFT
import openai  # noqa: E402
openai.OpenAI = _FakeClient

GL_FORMS = [{"form_number": "CG 21 47 12 07", "description": "Employment-Related Practices Exclusion"},
            {"form_number": "CG 21 49 09 99", "description": "Total Pollution Exclusion Endorsement"},
            {"form_number": "BR 5121", "description": "Total Liquor Liability Exclusion"}]
UMB_FORMS = [{"form_number": "", "description": "Abuse or Molestation Exclusion"},
             {"form_number": "", "description": "Communicable Disease Exclusion"},
             {"form_number": "", "description": "Human Trafficking Exclusion"},
             {"form_number": "", "description": "Total pollution exclusion with a building heating, cooling and "
                                                 "dehumidifying equipment exception and a hostile fire exception"},
             {"form_number": "", "description": "No Coverage for Sublimits of Insurance Endorsement"}]
DATA = {
    "client_info": {"named_insured": "HARP HOTELS INC", "dba": "Holiday Inn Express"},
    "locations": [{"address": "7900 SW Lost River Rd", "city": "Stuart", "state": "FL", "tiv": 7950000}],
    "coverages": {
        "property": {"carrier": "Underwriters at Lloyds, London", "total_premium": 37243.15,
                     "deductibles": [{"description": "All Other Perils", "amount": "$25,000"}]},
        "general_liability": {"carrier": "Berkley Specialty Insurance Company", "total_premium": 13520.13,
                              "forms_endorsements": GL_FORMS,
                              "limits": [{"description": "Each Occurrence", "limit": "$1,000,000"}]},
        "umbrella": {"carrier": "StarStone National Insurance Company", "total_premium": 6061,
                     "forms_endorsements": UMB_FORMS},
        "terrorism": {"carrier": "", "premium": 0, "optional": True},
    },
    "expiring_premiums": {"property": 37766, "general_liability": 24343, "umbrella": 7081},
}

# 1. subject pieces
check(ce._corporate_short("Harp Hotels Inc") == "Harp Hotels", "Inc dropped")
check(ce._corporate_short("Vista Lodging, LLC DBA Courtyard") == "Vista Lodging", "LLC and DBA dropped")
check(ce._corporate_short("Smith Holdings Co. Inc.") == "Smith Holdings", "stacked suffixes")
check(ce.email_subject("Harp Hotels", "Holiday Inn Express", True) == "Harp Hotels - Holiday Inn Express Renewal Proposal",
      "renewal subject")
check(ce.email_subject("Harp Hotels", "", False) == "Harp Hotels Insurance Proposal", "no DBA, new business")
check(ce.email_subject("Holiday Inn Express", "Holiday Inn Express", True) == "Holiday Inn Express Renewal Proposal",
      "no repeat when the names match")

ctx = ce.build_email_context(DATA, {"contact_first_name": "Hiren"})
check(ctx["subject"] == "Harp Hotels - Holiday Inn Express Renewal Proposal", ctx["subject"])

# 2. exclusion bullets: one per line, umbrella separate, topics as printed
b = ctx["exclusion_bullets"]
check(len(b) == 2, b)
check(b[0].startswith("Please note the General Liability policy excludes Employment-Related Practices, Total "
                      "Pollution, and Total Liquor Liability."), b[0])
check(b[1].startswith("Please note the Umbrella/Excess excludes Abuse or Molestation, Communicable Disease, "
                      "Human Trafficking"), b[1])
check("Total Pollution (with a Building Heating, Cooling and Dehumidifying Equipment exception and a Hostile "
      "Fire exception)" in b[1], b[1])
check(b[1].endswith(", and Sublimits of Insurance.") and "No Coverage" not in b[1], b[1])

# 2b. topic wording
check(ce._exclusion_topic("EXCLUSION - ASSAULT OR BATTERY") == "Assault or Battery", "A&B topic")
check(ce._exclusion_topic("Cannabis Exclusion With Hemp and Lessor Risk Exceptions")
      == "Cannabis (with Hemp and Lessor Risk exceptions)", ce._exclusion_topic(
          "Cannabis Exclusion With Hemp and Lessor Risk Exceptions"))
check(ce._exclusion_topic("PFAS EXCLUSION") == "PFAS", ce._exclusion_topic("PFAS EXCLUSION"))
check(ce._exclusion_topic("TOTAL FIREARMS EXCLUSION ENDORSEMENT") == "Total Firearms", "endorsement tail")
check(ce._fix_acronyms("Terrorism/Tria and Epli, Hi Ada") == "Terrorism/TRIA and EPLI, Hi Ada", "acronyms only")

# 3. draft: subject forced, the merged bullet replaced by the two bullets, TRIA capitals
r = ce.draft_email(ctx)
body = r["body"]
check(r["subject"] == "Harp Hotels - Holiday Inn Express Renewal Proposal", r["subject"])
check(body.count("- Please note") == 2 and "and the Umbrella/Excess excludes" not in body, body)
lines = body.split("\n")
gl_i = next(i for i, ln in enumerate(lines) if ln.startswith("- Please note the General Liability"))
check(lines[gl_i + 1].startswith("- Please note the Umbrella/Excess"), "umbrella bullet right after GL")
check(lines[gl_i - 1].startswith("- General Liability deductibles"), "bullets sit where the model put them")
check("Terrorism/TRIA" in body and "Tria" not in body, "TRIA keeps its capitals")
check(r["missing_exclusions"] == [], r["missing_exclusions"])
facts = _FakeClient.last_messages[1]["content"]
check("SUBJECT to use exactly: Harp Hotels - Holiday Inn Express Renewal Proposal" in facts, "subject in the facts")
check("KEY EXCLUSION BULLETS" in facts and b[1] in facts, "bullets in the facts")

# 4. model left the exclusions out entirely -> bullets added before the optional coverages bullet
_FakeClient.reply = "\n".join(ln for ln in MODEL_DRAFT.split("\n") if "Please note" not in ln)
r = ce.draft_email(ctx)
lines = r["body"].split("\n")
opt_i = next(i for i, ln in enumerate(lines) if "recommended but not included" in ln)
check(lines[opt_i - 1].startswith("- Please note the Umbrella/Excess") and
      lines[opt_i - 2].startswith("- Please note the General Liability"), lines[opt_i - 2:opt_i + 1])

# 5. instruction about the subject keeps the model's subject
_FakeClient.reply = MODEL_DRAFT.replace("SUBJECT: Holiday Inn Express Renewal Proposal", "SUBJECT: Harp renewal")
r = ce.draft_email(ctx, instruction="make the subject shorter", previous=MODEL_DRAFT)
check(r["subject"] == "Harp renewal", r["subject"])

# 6. no key-points section at all -> section added before the closing line
_FakeClient.reply = "SUBJECT: x\nBODY:\nHi Hiren,\n\nAttached is the proposal.\n\nLet me know if you have questions.\n\nStefan"
r = ce.draft_email(ctx)
check("The key points are:\n- Please note the General Liability" in r["body"] and
      r["body"].index("Please note") < r["body"].index("Let me know"), r["body"])

# 6b. a deductible bullet that says "excluding" stays; a "Please note" paragraph is replaced
_FakeClient.reply = MODEL_DRAFT.replace(
    "- Property deductibles are $25,000 for All Other Perils, 5% for Named Storm, and 5% for Sinkhole.",
    "- Property deductibles are $25,000 for All Other Perils, excluding Named Storm and Sinkhole at 5%.").replace(
    "- Please note the General Liability policy excludes", "Please note the General Liability policy excludes")
r = ce.draft_email(ctx)
check("excluding Named Storm and Sinkhole at 5%" in r["body"], "deductible bullet kept")
check(r["body"].count("Please note") == 2 and r["body"].count("- Please note") == 2, r["body"])

# 7. sublimit sentence rides on the umbrella bullet only
ctx2 = dict(ctx, sublimit_note="Because the excess policy excludes coverage that is sublimited in the underlying policy, "
                               "the excess limit does not apply above the $300,000 Assault and Battery limit in the "
                               "General Liability policy.")
b2 = ce.exclusion_bullets(ctx2)
check("sublimited" in b2[1] and "sublimited" not in b2[0], b2)

print(f"client_email_bullets_test: {PASS} checks passed")
