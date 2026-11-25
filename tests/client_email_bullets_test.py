"""Regression checks for the client email subject, exclusion bullets and model (Stefan, Nov 4 and Nov 24 2026).

Run from the repo root:  python3 tests/client_email_bullets_test.py
No OpenAI calls: the model is replaced by a fake that returns the Harp Hotels draft Stefan saw
(subject without the corporate name, GL and umbrella exclusions merged in one bullet,
"Terrorism/Tria").

Nov 24 2026 (WG Hotel): "on proposal draft email update with better model: need to reduce all
the please". Five lines of coverage carried exclusions, so five bullets in a row opened "Please
note the ...". Stefan sent the email with each bullet opening on the line itself ("Property
policy excludes ..."). The bullets are now written that way, a "Please note" lead-in the model
writes is removed, and the email is drafted by OpenAI's flagship with a fallback chain.
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("OPENAI_API_KEY", "test")

import logging  # noqa: E402
logging.disable(logging.CRITICAL)
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
        _FakeClient.requests.append(kwargs)
        if kwargs["model"] in _FakeClient.fail:
            raise RuntimeError(f"The model `{kwargs['model']}` does not exist or you do not have access to it.")
        text = "" if kwargs["model"] in _FakeClient.empty else _FakeClient.reply
        msg = types.SimpleNamespace(content=text)
        return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)])


_FakeClient.reply = MODEL_DRAFT
_FakeClient.requests, _FakeClient.fail, _FakeClient.empty = [], set(), set()
for _v in ("EMAIL_MODEL", "EMAIL_FALLBACK_MODELS", "EMAIL_MAX_TOKENS", "EMAIL_REASONING_EFFORT"):
    os.environ.pop(_v, None)
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
check(b[0] == "General Liability policy excludes Employment-Related Practices, Total Pollution, and Total "
              "Liquor Liability.", b[0])
check(b[1].startswith("Umbrella/Excess excludes Abuse or Molestation, Communicable Disease, Human Trafficking"), b[1])
check(not any("please" in x.lower() for x in b), b)
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
check("please" not in body.lower(), body)                       # the model's "Please note" bullet is gone
check(body.count(" excludes ") == 2 and "and the Umbrella/Excess excludes" not in body, body)
lines = body.split("\n")
gl_i = next(i for i, ln in enumerate(lines) if ln.startswith("- General Liability policy excludes"))
check(lines[gl_i + 1].startswith("- Umbrella/Excess excludes"), "umbrella bullet right after GL")
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
check(lines[opt_i - 1].startswith("- Umbrella/Excess excludes") and
      lines[opt_i - 2].startswith("- General Liability policy excludes"), lines[opt_i - 2:opt_i + 1])

# 5. instruction about the subject keeps the model's subject
_FakeClient.reply = MODEL_DRAFT.replace("SUBJECT: Holiday Inn Express Renewal Proposal", "SUBJECT: Harp renewal")
r = ce.draft_email(ctx, instruction="make the subject shorter", previous=MODEL_DRAFT)
check(r["subject"] == "Harp renewal", r["subject"])

# 6. no key-points section at all -> section added before the closing line
_FakeClient.reply = "SUBJECT: x\nBODY:\nHi Hiren,\n\nAttached is the proposal.\n\nLet me know if you have questions.\n\nStefan"
r = ce.draft_email(ctx)
check("The key points are:\n- General Liability policy excludes" in r["body"] and
      r["body"].index("General Liability policy excludes") < r["body"].index("Let me know"), r["body"])

# 6b. a deductible bullet that says "excluding" stays; a "Please note" paragraph is replaced
_FakeClient.reply = MODEL_DRAFT.replace(
    "- Property deductibles are $25,000 for All Other Perils, 5% for Named Storm, and 5% for Sinkhole.",
    "- Property deductibles are $25,000 for All Other Perils, excluding Named Storm and Sinkhole at 5%.").replace(
    "- Please note the General Liability policy excludes", "Please note the General Liability policy excludes")
r = ce.draft_email(ctx)
check("excluding Named Storm and Sinkhole at 5%" in r["body"], "deductible bullet kept")
check("Please note" not in r["body"] and r["body"].count("- General Liability policy excludes") == 1
      and r["body"].count("- Umbrella/Excess excludes") == 1, r["body"])

# 7. sublimit sentence rides on the umbrella bullet only
ctx2 = dict(ctx, sublimit_note="Because the excess policy excludes coverage that is sublimited in the underlying policy, "
                               "the excess limit does not apply above the $300,000 Assault and Battery limit in the "
                               "General Liability policy.")
b2 = ce.exclusion_bullets(ctx2)
check("sublimited" in b2[1] and "sublimited" not in b2[0], b2)

# ------------------------------------------------------------------ Nov 24 2026
# 8. the WG Hotel email: five lines with exclusions, each bullet as Stefan sent it
WG = {
    "client_info": {"named_insured": "WG HOTEL, LLC"},
    "locations": [{"address": "815 North Pottstown Pike", "city": "Exton", "state": "PA", "tiv": 22500000}],
    "coverages": {
        "property": {"carrier": "Certain Underwriters at Lloyd's, London", "total_premium": 170359.41,
                     "forms_endorsements": [{"form_number": "", "description": "Cosmetic Damage Exclusion"},
                                            {"form_number": "", "description": "Marijuana Exclusion"}]},
        "general_liability": {"carrier": "Southlake Specialty Insurance Company", "total_premium": 222916.18,
                              "forms_endorsements": [
                                  {"form_number": "CG 21 32 05 09", "description": "Communicable Disease Exclusion"},
                                  {"form_number": "CG 21 47 12 07", "description": "Employment Related Practices Exclusion"},
                                  # the title exactly as the Southlake forms list cuts it
                                  {"form_number": "CG 21 65 12 04",
                                   "description": "Total Pollution Exclusion With a Building Heating, Cooling and Dehumidifying"}]},
        "excess_liquor_liability": {"carrier": "General Star Indemnity Company", "total_premium": 8925,
                                    "forms_endorsements": [
                                        {"form_number": "", "description": "Exclusion - Human Trafficking"},
                                        {"form_number": "", "description": "Sexual Misconduct Exclusion"},
                                        {"form_number": "", "description": "Assault or Battery Exclusion"},
                                        {"form_number": "", "description": "Exclusion of Coverage Subject to Sublimits"}]},
        "umbrella": {"carrier": "PartnerRe Insurance Solutions Bermuda Ltd.", "total_premium": 80088,
                     "forms_endorsements": [{"form_number": "", "description": "Communicable Disease Exclusion"},
                                            {"form_number": "", "description": "Liquor Liability Exclusion"},
                                            {"form_number": "", "description": "Sublimited Coverages Exclusion"}]},
        "umbrella_layer_2": {"carrier": "Ategrity Specialty Insurance Company", "total_premium": 20000,
                             "forms_endorsements": [{"form_number": "", "description": "Human Trafficking Exclusion"}]},
    },
}
wctx = ce.build_email_context(WG, {"contact_first_name": "Harry"})
wb = wctx["exclusion_bullets"]
check([x.split(" excludes ")[0] for x in wb] == ["Property policy", "General Liability policy", "Excess Liquor Liability",
                                                 "Umbrella/Excess", "2nd Excess Layer"], wb)
check(wb[0] == "Property policy excludes Cosmetic Damage and Marijuana.", wb[0])
check(wb[2].startswith("Excess Liquor Liability excludes Human Trafficking, Sexual Misconduct, Assault or Battery"), wb[2])
check(not any("please" in x.lower() for x in wb), wb)
# the title the quote cuts short reads whole, in the same form as an uncut one
check(wb[1] == "General Liability policy excludes Communicable Disease, Employment Related Practices, and Total "
               "Pollution (with a Building Heating, Cooling and Dehumidifying Equipment exception and a Hostile Fire "
               "exception).", wb[1])
check(ce._exclusion_topic("Total Pollution Exclusion Endorsement") == "Total Pollution", "other pollution titles untouched")

# 9. the sublimit sentence belongs to the umbrella, never to the excess liquor policy listed before it
note = "Because the excess policy excludes coverage that is sublimited in the underlying policy, the excess limit " \
       "does not apply above the $300,000 Assault and Battery limit in the General Liability policy."
wb2 = ce.exclusion_bullets(dict(wctx, sublimit_note=note))
check(wb2[3].endswith(note) and all(note not in x for i, x in enumerate(wb2) if i != 3), wb2)
# only the 2nd layer names the sublimit exclusion -> the sentence goes there
alt = {"coverages": [{"coverage": "Umbrella/Excess", "key_exclusions": ["Communicable Disease"]},
                     {"coverage": "2nd Excess Layer", "key_exclusions": ["Coverage Subject to Sublimits"]}],
       "sublimit_note": note}
ab = ce.exclusion_bullets(alt)
check(ab[1].endswith(note) and note not in ab[0], ab)
# an excess liquor policy's own sublimit exclusion does not produce a General Liability sentence
only_liq = {"client_info": {"named_insured": "X LLC"}, "coverages": {
    "general_liability": {"carrier": "A", "total_premium": 1,
                          "limits": [{"description": "Assault and Battery Sublimit", "limit": "$300,000"}]},
    "excess_liquor_liability": dict(WG["coverages"]["excess_liquor_liability"])}}
check(ce.build_email_context(only_liq, {})["sublimit_note"] == "", "no GL sublimit sentence from the liquor excess")

# 10. a lead-in the model writes anywhere is removed; the point stays
check(ce.drop_please_lead("- Please note the roof was replaced in 2019.") == "- The roof was replaced in 2019.", "bullet")
check(ce.drop_please_lead("Please be advised that flood is excluded.") == "Flood is excluded.", "sentence")
check(ce.drop_please_lead("please note: wind is 5%") == "Wind is 5%", "colon")
same = "Let me know if you have questions, and please send the signed application."
check(ce.drop_please_lead(same) == same, "a request in Stefan's own words is not touched")
_FakeClient.reply = MODEL_DRAFT.replace("- General Liability deductibles are $2,500 per claim.",
                                        "- Please note General Liability deductibles are $2,500 per claim.")
r = ce.draft_email(ctx)
check("- General Liability deductibles are $2,500 per claim." in r["body"] and "lease note" not in r["body"], r["body"])
r = ce.draft_email(ctx, instruction="open the deductible bullet with please note", previous=MODEL_DRAFT)
check("- Please note General Liability deductibles" in r["body"], "an instruction that asks for it keeps it")
_FakeClient.reply = MODEL_DRAFT

# 11. the prompt no longer teaches the lead-in
check("Please note the" not in ce._SYSTEM_PROMPT.split("Structure, in this order:")[0], "sample email")
check("- General Liability policy excludes Human Trafficking and Total Pollution." in ce._SYSTEM_PROMPT, "sample bullet")
check('Never write "Please note"' in ce._SYSTEM_PROMPT, "rule")

# 12. the model: OpenAI's flagship by default, the Railway variable still decides, a failure falls through
check(ce.email_models() == ["gpt-6.2-astra", "gpt-6.2-sol", "gpt-6.1-sol"], ce.email_models())
_FakeClient.requests.clear()
r = ce.draft_email(ctx)
req = _FakeClient.requests[-1]
check(r["model"] == "gpt-6.2-astra" and r["models_failed"] == [] and len(_FakeClient.requests) == 1, r["model"])
check(req["max_completion_tokens"] == 8000, req["max_completion_tokens"])     # room for the model to think
check("temperature" not in req and "reasoning_effort" not in req, sorted(req))
os.environ["EMAIL_MODEL"] = "gpt-6.2-sol"
check(ce.email_models() == ["gpt-6.2-sol", "gpt-6.1-sol"], ce.email_models())
check(ce.draft_email(ctx)["model"] == "gpt-6.2-sol", "Railway variable decides")
check(ce.email_models("gpt-6-luna")[0] == "gpt-6-luna", "explicit argument first")
os.environ["EMAIL_FALLBACK_MODELS"] = "gpt-6.1-terra, gpt-6.2-sol"
check(ce.email_models() == ["gpt-6.2-sol", "gpt-6.1-terra"], ce.email_models())
os.environ.pop("EMAIL_MODEL"); os.environ.pop("EMAIL_FALLBACK_MODELS")
# the flagship is not available to the account -> the next model writes the same email
_FakeClient.fail = {"gpt-6.2-astra"}
_FakeClient.requests.clear()
r = ce.draft_email(ctx)
check(r["model"] == "gpt-6.2-sol" and r["models_failed"] == ["gpt-6.2-astra"], (r["model"], r["models_failed"]))
check([q["model"] for q in _FakeClient.requests] == ["gpt-6.2-astra", "gpt-6.2-sol"], "tried in order")
check("- Umbrella/Excess excludes" in r["body"], "same deterministic finish")
# it answers with nothing (all of its allowance spent thinking) -> the next model
_FakeClient.fail, _FakeClient.empty = set(), {"gpt-6.2-astra", "gpt-6.2-sol"}
r = ce.draft_email(ctx)
check(r["model"] == "gpt-6.1-sol" and r["models_failed"] == ["gpt-6.2-astra", "gpt-6.2-sol"], r["models_failed"])
# every model fails -> the error reaches the page, as before
_FakeClient.fail, _FakeClient.empty = {"gpt-6.2-astra", "gpt-6.2-sol", "gpt-6.1-sol"}, set()
try:
    ce.draft_email(ctx)
    check(False, "should raise")
except RuntimeError as e:
    check("does not exist" in str(e), str(e))
_FakeClient.fail = set()
os.environ["EMAIL_REASONING_EFFORT"] = "medium"
_FakeClient.requests.clear()
ce.draft_email(ctx)
check(_FakeClient.requests[-1].get("reasoning_effort") == "medium", "effort only when the variable is set")
os.environ.pop("EMAIL_REASONING_EFFORT")

# 13. the page and the endpoint say which model wrote it
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
web = open(os.path.join(ROOT, "web_app.py"), encoding="utf-8").read()
check('"model": result.get("model") or ""' in web, "endpoint returns the model")
html = open(os.path.join(ROOT, "templates", "proposal_web.html"), encoding="utf-8").read()
check('id="email-model"' in html and "'Drafted with ' + d.model" in html, "page shows the model")

print(f"client_email_bullets_test: {PASS} checks passed")
