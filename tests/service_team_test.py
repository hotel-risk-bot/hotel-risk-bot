"""Regression checks for the proposal's Service Team choices (Nov 20 2026).

Run from the repo root:  python3 tests/service_team_test.py
No OpenAI calls.

Stefan, Nov 20 2026: add Monique Miller-Stewart, Hotel Franchise Team Lead and Account
Executive, to the Account Executive dropdown. The dropdown lives in two places that must agree:
the <select> on the review screen (templates/proposal_web.html) and ACCOUNT_EXECUTIVES in
proposal_generator.py. A value the generator does not know silently prints Maureen Harvey.
"""
import os
import re
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import docx  # noqa: E402
import proposal_generator as pg  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


# 1. Monique's entry, as Stefan gave it
m = pg.ACCOUNT_EXECUTIVES.get("monique_miller_stewart")
check(m == {"role": "Hotel Franchise Team Lead and Account Executive", "name": "Monique Miller-Stewart",
            "phone": "O: 312-279-4865", "email": "monique.millerstewar@hubinternational.com"}, m)

# 2. the page's dropdowns and the generator's lists carry the same choices
html = open(os.path.join(ROOT, "templates", "proposal_web.html"), encoding="utf-8").read()


def options(select_id):
    block = re.search(rf'<select id="{select_id}".*?</select>', html, re.S).group(0)
    return re.findall(r'<option value="([^"]*)">([^<]*)</option>', block)


ae_opts = options("edit-account-executive")
check([v for v, _t in ae_opts] == list(pg.ACCOUNT_EXECUTIVES), [v for v, _t in ae_opts])
check(dict(ae_opts)["monique_miller_stewart"]
      == "Monique Miller-Stewart - Hotel Franchise Team Lead and Account Executive", dict(ae_opts))
check(ae_opts[0][0] == "maureen_harvey", "Maureen Harvey stays first, the default")
for v, text in ae_opts:
    check(text.startswith(pg.ACCOUNT_EXECUTIVES[v]["name"]), f"{v}: option text names the same person")
am_opts = [v for v, _t in options("edit-associate-am") if v]
check(am_opts == list(pg.ASSOCIATE_ACCOUNT_MANAGERS), am_opts)

# 3. what the generator builds from a selection
team = pg._build_service_team({"service_team": {"account_executive": "monique_miller_stewart"}})
check([p["name"] for p in team] == ["Stefan Burkey", "Monique Miller-Stewart", "Sheena Callazo, RPLU"], team)
check(pg._build_service_team({})[1]["name"].startswith("Maureen Harvey"), "no selection -> Maureen Harvey")
check(pg._build_service_team({"service_team": {"account_executive": "someone_new"}})[1]["name"]
      .startswith("Maureen Harvey"), "unknown value -> Maureen Harvey")

# 4. the proposal prints her row, and the table is still 7.5 in wide
data = {"client_info": {"named_insured": "Sample Hotel LLC"},
        "service_team": {"account_executive": "monique_miller_stewart", "associate_account_manager": "liam_omahoney"},
        "coverages": {"property": {"carrier": "Starr Surplus Lines Insurance Company", "premium": 1000,
                                   "total_premium": 1000}}}
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "p.docx")
    pg.generate_proposal(data, out)
    d = docx.Document(out)
    tbl = next(t for t in d.tables if [c.text for c in t.rows[0].cells][:2] == ["Role", "Name"])
    rows = [[c.text for c in r.cells] for r in tbl.rows]
    widths = [c.width.inches for c in tbl.rows[1].cells]
check(rows[2] == ["Hotel Franchise Team Lead and Account Executive", "Monique Miller-Stewart", "O: 312-279-4865",
                  "monique.millerstewar@hubinternational.com"], rows[2])
check([r[1] for r in rows[1:]] == ["Stefan Burkey", "Monique Miller-Stewart", "Liam O'Mahoney", "Sheena Callazo, RPLU"],
      rows)
check(abs(sum(widths) - 7.5) < 0.01, widths)

# 5. every address fits its column on one line (Carlito has Calibri's metrics; skipped where the
#    font is not installed). Cell padding is 0.08 in a side.
font_path = "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"
if os.path.exists(font_path):
    from PIL import ImageFont
    f = ImageFont.truetype(font_path, 1000)
    people = ([pg.SERVICE_TEAM_LEADER, pg.SERVICE_TEAM_CLAIMS_ADVOCATE] + list(pg.ACCOUNT_EXECUTIVES.values())
              + list(pg.ASSOCIATE_ACCOUNT_MANAGERS.values()))
    for col, key in ((1, "name"), (2, "phone"), (3, "email")):
        for p in people:
            for line in p[key].split("\n"):
                need = f.getlength(line) / 1000 * 10 / 72
                check(need <= widths[col] - 0.16, f"{line!r} needs {need:.2f} in, column gives {widths[col] - 0.16:.2f}")

print(f"service_team_test: {PASS} checks passed")
