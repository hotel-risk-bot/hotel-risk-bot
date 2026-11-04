"""Regression checks for the verified OCR of image-only quote pages (Nov 4 2026).

Run from the repo root:  python3 tests/ocr_verify_test.py
Needs poppler (pdftoppm / pdfinfo) and reportlab, both already used by the app. No OpenAI calls:
the model is mocked, so the test exercises rendering, the two-read check, the 300 DPI re-read,
the review notes and the hand-off of those notes to the review screen's warnings.

Why: Harp Hotels Inc's Tower Hill quote is a 13-page "Microsoft: Print To PDF" file with no text
layer. The old OCR sent every page at "low" detail (512 px), so the Premises and Buildings page
(Building $7,250,000 / BPP $300,000 / BI $400,000) came back as $1,250,000 / $50,000 / $412,000
on one run and a $2,170,000 TIV on another.
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


def make_image_pdf(path, pages):
    """A PDF whose pages are pictures of text (no text layer), like a Print To PDF quote."""
    from io import BytesIO
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from PIL import Image, ImageDraw
    c = canvas.Canvas(path, pagesize=letter)
    for lines in pages:
        img = Image.new("L", (850, 1100), 255)
        d = ImageDraw.Draw(img)
        for i, line in enumerate(lines):
            d.text((60, 80 + 22 * i), line, fill=0)
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), 0, 0, width=letter[0], height=letter[1])
        c.showPage()
    c.save()


TRUE = ("Location 1\n7900 SW LOST RIVER RD STUART FL 34997\nBldg# 1 7900 SW LOST RIVER RD STUART FL 34997 $7,250,000\n"
        "Building Personal Property $300,000\nBusiness Income with Extra Expense Coverage - 1/4 $400,000\nNamed Storm: 5%")
BAD = TRUE.replace("$7,250,000", "$1,250,000")

# 1. figures: normalised multiset, display form
check(pe._ocr_figures("$7,250,000 $300,000.00 5% $37,243.15") == ["300000", "37243.15", "5%", "7250000"],
      pe._ocr_figures("$7,250,000 $300,000.00 5% $37,243.15"))
check(pe._ocr_figures(TRUE) != pe._ocr_figures(BAD), "a misread digit changes the figures")
check(pe._ocr_show("7250000") == "$7,250,000" and pe._ocr_show("5%") == "5%", "display form")
check(pe._ocr_note_key("/x/3f9a01bc_Quote Letter.pdf") == "Quote Letter.pdf", "upload prefix stripped")

td = tempfile.mkdtemp()
pdf = os.path.join(td, "3f9a01bc_Quote_Letter_for_Harp_Hotels_Inc.pdf")
make_image_pdf(pdf, [["Coverage Summary", "Total Premium $37,243.15"], TRUE.split("\n")])
check(len(pe._render_page_b64(pdf, 2, 200)) > 1000, "page renders to PNG")

pe.OCR_MODEL, pe.OCR_CHECK_MODEL = "main", "check"
calls = []


def plan(table):
    """table: model -> list of texts returned in order for that model."""
    seq = {k: list(v) for k, v in table.items()}

    def _m(img_b64, model):
        calls.append(model)
        return seq[model].pop(0) if len(seq[model]) > 1 else seq[model][0]
    return _m


# 2. both reads agree -> main read, two calls
calls.clear()
pe._ocr_image = plan({"main": [TRUE], "check": [TRUE]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check("$7,250,000" in t and n == "" and lvl == "" and calls == ["main", "check"], (t, n, calls))

# 3. check read wrong -> 300 DPI re-read agrees with main -> main read
calls.clear()
pe._ocr_image = plan({"main": [TRUE, TRUE], "check": [BAD]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check("$7,250,000" in t and lvl == "info" and "confirmed $7,250,000" in n and calls.count("main") == 2, (n, calls))

# 4. main read wrong -> re-read agrees with check -> re-read text (the right figure)
calls.clear()
pe._ocr_image = plan({"main": [BAD, TRUE], "check": [TRUE]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check("$7,250,000" in t and "$1,250,000" not in t and lvl == "info" and "confirmed $7,250,000" in n, (t, n))

# 5. nothing agrees -> note names both figures
calls.clear()
pe._ocr_image = plan({"main": [BAD, BAD.replace("$300,000", "$350,000")], "check": [TRUE]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check(lvl == "review" and "could not be confirmed" in n and "$1,250,000" in n and "$7,250,000" in n, n)

# 5b. the check read only skipped a row -> re-read has it -> main read, no note
SKIP = TRUE.replace("Building Personal Property $300,000\n", "")
calls.clear()
pe._ocr_image = plan({"main": [TRUE, TRUE], "check": [SKIP]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check("$300,000" in t and n == "" and lvl == "" and calls.count("main") == 2, (n, calls))

# 5c. a repeated figure listed once more is not a disagreement -> no re-read
calls.clear()
pe._ocr_image = plan({"main": [TRUE], "check": [TRUE + "\nNamed Storm: 5%"]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check(n == "" and calls == ["main", "check"], (n, calls))

# 5d. main misread + check skipped that row -> the re-read's own figure is shown, page flagged
calls.clear()
pe._ocr_image = plan({"main": [BAD, TRUE], "check": [TRUE.replace("$7,250,000", "")]})
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check(lvl == "review" and "one read: $1,250,000" in n and "third read: $7,250,000" in n and "$7,250,000" in t, n)

# 6. main read fails -> check read used and flagged
def _fail_main(img_b64, model):
    if model == "main":
        raise RuntimeError("timeout")
    return TRUE
pe._ocr_image = _fail_main
t, n, lvl = pe._ocr_page_verified(pdf, 2)
check("$7,250,000" in t and lvl == "review" and "read only once" in n, n)

# 7. whole file: page markers, notes registered under the name the pipeline uses
pe._OCR_REVIEW_NOTES.clear()
pe._ocr_image = lambda b, m: TRUE if m == "main" else BAD     # check model misreads every page
text = pe._extract_with_ocr(pdf, 0)          # page count looked up with pdfinfo
check(text.count("--- Page") == 2 and "$7,250,000" in text and "$1,250,000" not in text, text[:200])
notes = pe._OCR_REVIEW_NOTES.get("Quote_Letter_for_Harp_Hotels_Inc.pdf") or []
check(len(notes) == 2 and all(lvl == "info" for lvl, _n in notes), notes)
pe._OCR_REVIEW_NOTES.clear()
_n3 = {"c": 0}
def _three_ways(b, m):
    if m == "check":
        return TRUE
    _n3["c"] += 1
    return BAD if _n3["c"] % 2 else BAD.replace("$300,000", "$350,000")
pe._ocr_image = _three_ways
pe._extract_with_ocr(pdf, 1)        # one page, so the mocked reads come in a fixed order
notes = pe._OCR_REVIEW_NOTES.get("Quote_Letter_for_Harp_Hotels_Inc.pdf") or []
check(notes and all(lvl == "review" for lvl, _n in notes), notes)

# 8. the notes reach the review screen as REVIEW REQUIRED warnings
pe._OCR_REVIEW_NOTES.clear()
src = open(pe.__file__, encoding="utf-8").read()
check('for _lvl, _on in _OCR_REVIEW_NOTES.pop(it.get("filename", ""), [])' in src,
      "structure_insurance_data attaches OCR notes")
check('f"REVIEW REQUIRED — \'{it.get(\'filename\', \'\')}\' is a scanned / image-only "' in src, "review wording")
check('"detail": "low"' not in src, "no low-detail OCR left")

print(f"ocr_verify_test: {PASS} checks passed")
