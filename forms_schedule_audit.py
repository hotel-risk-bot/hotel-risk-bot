"""
forms_schedule_audit.py — deterministic Forms & Endorsements completeness check.

Why this exists (Sep 15 2026): the Neptune Lodging Partners proposal listed 11 of the
39 forms on the Houston Specialty excess quote. GPT had extracted them, and then two
prefix/description cleaners silently threw away every HSIC CX ES exclusion (Assault or
Battery, Total Firearms, Sublimits or Reduced Limits in Underlying Insurance, Sexual
Misconduct, ...). A proposal that omits an exclusion is an E&O exposure, so the forms
list can no longer depend on an LLM being exhaustive or on a prefix table being right.

This module reads the carrier quote text directly, finds every row of the forms
schedule that carries a real form number, and reconciles it against the extracted
forms_endorsements list:

  * rows on the quote's schedule that the extractor missed are APPENDED
    (flagged schedule_verified=True so downstream filters leave them alone);
  * the coverage gets a `forms_audit` dict the review UI and the docx use to show
    "quote lists N numbered forms / extracted M / added K" and any residual gap;
  * a warning string is pushed onto data["warnings"].

It is deliberately conservative: it only ever ADDS rows that literally appear with a
form number in the coverage's own quote file, and it never removes anything.

Package quotes (Oct 6 2026, Bloom Ventures / Futuristic GL + Liquor + Business Auto in
one PDF): such a schedule carries a Line column ("General Liability", "Business Auto",
"Liquor Liability", "Interline") after the title. Before this the whole 53-row schedule
was reconciled against EVERY liability coverage from that file, so the Business Auto
section gained 42 General Liability forms. The parser now reads that column: a row
tagged for another line is never offered to this coverage, Interline rows go to every
coverage on the policy, a row the carrier tags for this line is accepted even when a
prefix rule would have vetoed it, and a leading edition date ("01/25") moves from the
title into the form number.

Berkley Aspire layout (Oct 13 2026, Triumph Investments Group GL quote 2642731876): the
audit reported "26 numbered forms" on a 50-form schedule and ADDED two page footers as
forms ("ENP9014 1024 Page 3 of 4", "ENP5101 1124 Page 1 of 1"), rows the generator then
prints because schedule_verified rows are never dropped. Four gaps, all fixed here:
  * letter-suffixed numbers ("BR5269G 1024", "BR902I 0816", 24 of the 50) and an edition
    glued onto the number ("EMO30011023") did not read as form numbers;
  * a title that wraps onto two lines puts the number alone on the middle line
    ("CG2107 0514" between "Exclusion - Access Or ... - Limited" and "Bodily Injury ...");
    those rows are now rebuilt from the indented title halves above and below;
  * a form number beside a page counter or a copyright line is page furniture, never a
    row ("Page 3 of 4", "Includes copyrighted material ...", "(c) ISO Properties");
  * the schedule block now ends at a page break unless the next page carries on with it
    (a "(Cont'd)" header or numbered rows near the top), so the parser no longer reads
    on into the TRIA disclosure notice that follows the schedule.
Matching against the extracted list also treats O/0 and I/L/1 as the same character, so
"BR902I" read by GPT as "BR9021" is not appended a second time.
"""
import logging
import re

logger = logging.getLogger(__name__)

# A complete form number, matched against the LAST (or first) column of a schedule row.
# ISO ("CX 21 13 04 13", "CG 20 29 12 19", "IL 09 85 12 20"), carrier-prefixed
# ("HSIC CX ES 01 44 04 21", "LBGL 21 53 12 24", "TSIC 004 0424", "FORMS - SCHED 08 12"),
# letter-suffixed Berkley numbers ("BR5269G 1024", "BR902I 0816"), an edition glued onto
# the number ("EMO30011023" = EMO3001 10/23; the trailing MMYY must be a real month),
# dashed ("LIA-19182-0724", "FUT-SS 01 22") and edition-in-parentheses variants
# ("LB TRIA (12/2022)", "DS PN Annual (02-2022)").
_FORM_NUMBER_FULL_RE = re.compile(
    r"^(?:"
    r"(?:[A-Z]{1,6}(?:\s*-\s*[A-Z]{1,6})?\s+){0,4}"          # alpha prefix tokens
    r"(?:(?:[A-Z]{1,3}\s?)?\d{2,5}"                           # numeric group (+ optional letter prefix)
    r"|[A-Z]{1,4}\d{2,5}[A-Z]{1,2})"                          # letter prefix AND suffix ("BR5269G")
    r"(?:[\s\-/]\d{2,4}){0,4}"                                # edition / further numeric groups
    r"|[A-Z]{2,5}\d{3,5}(?:0[1-9]|1[0-2])\d{2}"               # edition glued on ("EMO30011023")
    r"|[A-Z]{2,6}(?:-[A-Z0-9]{1,8}){1,3}(?:\s\d{2}){0,2}"     # dashed carrier codes ("IFG-G-0002-DL 10 25")
    r"|[A-Z]-\d{1,4}\s\d{2}\s\d{2}"                          # one-letter state notices with edition ("D-2 01 20")
    r"|[A-Z]{2,6}(?:\s+[A-Za-z]{2,8}){0,2}\s*\(\d{2}[-/]\d{2,4}\)"  # edition in parentheses
    r")$"
)
# Fallback for single-spaced text (no column gap): a short, tightly-formed ISO-style
# number at the end of the line — at most two alpha tokens so it cannot swallow the
# tail of an upper-case description.
_FORM_NUMBER_TAIL_RE = re.compile(
    r"\s(?P<num>[A-Z]{2,6}(?:\s[A-Z]{2,4})?\s\d{2}(?:\s\d{2}){2,3})\s*$"
)

_SCHEDULE_HEADER_RE = re.compile(
    r"^\s*(FORMS?\s*:?|FORMS?\s+(?:AND|&)\s+ENDORSEMENTS?|SCHEDULE\s+OF\s+FORMS?(?:\s+AND\s+ENDORSEMENTS?)?|"
    r"(?:LISTING|LIST)\s+OF\s+FORMS?(?:\s+(?:AND|&)\s+ENDORSEMENTS?)?|"   # Burlington "LISTING OF FORMS AND ENDORSEMENTS"
    r"FORMS?\s+SCHEDULE|ENDORSEMENTS?\s+SCHEDULE|APPLICABLE\s+FORMS?|POLICY\s+FORMS?(?:\s+AND\s+ENDORSEMENTS?)?)"
    r"(?:\s*[\(\[]?\s*-?\s*CONT(?:INUED|['’]?D)?\.?\s*[\)\]]?)?\s*:?\s*$",   # "(Cont'd)", "Continued"
    re.I,
)
# Page breaks as the audit receives them: the smart extractor's "--- Page N ---" markers
# (form feeds from a whole-document pdftotext run are converted to the same marker).
_PAGE_MARK_RE = re.compile(r"^\s*-{2,}\s*Page(?:\s+\d+)?\s*-{2,}\s*$", re.I)
# Page furniture that can share a line with a form number but is never a schedule row:
# the quote's own form number beside its page counter ("ENP9014 1024   Page 3 of 4"),
# copyright lines printed under ISO / NAIC forms.
_FOOTER_DESC_RE = re.compile(
    r"^\s*page\s+\d+\b"
    r"|\bpage\s+\d+\s*(?:of|/)\s*\d+\b"
    r"|©|\(c\)\s*(?:19|20)\d{2}"
    r"|\bcopyrighted\s+material\b"
    r"|\binsurance\s+services\s+office\b|\biso\s+properties\b",
    re.I,
)
_SCHEDULE_END_RE = re.compile(
    r"^\s*(CONDITIONS?\s*:|SUBJECTIVITIES|SUBJECT\s+TO\s*:|PREMIUM\s*(?:&|AND)?\s*FEES|TERMS\s+AND\s+CONDITIONS|"
    r"THIS\s+ENDORSEMENT\s+CHANGES|BINDING\s+REQUIREMENTS|NOTES?\s*:)",
    re.I,
)
_COLUMN_HEADER_RE = re.compile(r"^\s*FORM\s+(NAME|TITLE|NUMBER|#|DESCRIPTION)\b", re.I)
_COLUMN_SPLIT_RE = re.compile(r"\s{2,}|\t")
_EDITION_RE = re.compile(r"^(\d{2}[/-]\d{2,4})\s+(?=\S)(.*)$")

# Line-of-business labels a package schedule prints as its last column (longest first so
# "commercial general liability" wins over "general liability" and "liability" alone never
# matches). Short, ambiguous labels (property, auto, crime, ...) only count once the schedule
# has shown it uses a Line column at all — see _tag_schedule_lobs.
_LOB_LABELS = (
    ("commercial general liability", "gl"), ("general liability", "gl"), ("cgl", "gl"), ("gl", "gl"),
    ("liquor liability", "liquor"), ("liquor", "liquor"),
    ("business auto", "auto"), ("commercial auto", "auto"), ("business automobile", "auto"),
    ("automobile", "auto"), ("auto", "auto"), ("garage", "auto"),
    ("interline", "interline"), ("common policy", "interline"), ("common", "interline"),
    ("all lines", "interline"), ("package", "interline"), ("il", "interline"),
    ("employment practices liability", "epli"), ("employment practices", "epli"), ("epli", "epli"), ("epl", "epli"),
    ("cyber liability", "cyber"), ("cyber", "cyber"),
    ("commercial property", "property"), ("property", "property"),
    ("umbrella", "umbrella"), ("excess liability", "umbrella"), ("commercial excess", "umbrella"), ("excess", "umbrella"),
    ("workers compensation", "wc"), ("workers' compensation", "wc"), ("workers comp", "wc"),
    ("crime", "crime"), ("inland marine", "inland_marine"),
    ("equipment breakdown", "eb"), ("boiler and machinery", "eb"), ("boiler & machinery", "eb"),
)
_LOB_TAIL_RE = re.compile(
    r"^(?P<title>.*?\S)\s+(?P<lob>" + "|".join(re.escape(l) for l, _ in _LOB_LABELS) + r")\s*$", re.I)
_LOB_MID_RE = re.compile(
    r"^(?P<title>.*\S)\s+(?P<lob>" + "|".join(re.escape(l) for l, _ in _LOB_LABELS) + r")\s+(?P<tail>[A-Za-z][^\n]*?)\s*$", re.I)
_LOB_KEY = {l: k for l, k in _LOB_LABELS}
_AMBIGUOUS_LOBS = {"property", "auto", "automobile", "crime", "excess", "common", "package", "il", "gl", "epl", "cyber", "liquor", "garage"}


def _compact(num: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(num or "").upper())


_FUZZY_MAP = str.maketrans("OIL", "011")


def _fuzzy(compact: str) -> str:
    """A compacted number with the characters PDF text and GPT swap most often folded
    together (O/0, I/L/1). Used only to decide whether a row is already extracted."""
    return (compact or "").translate(_FUZZY_MAP)


def _is_form_number(tok: str) -> bool:
    t = (tok or "").strip()
    if not t or len(t) > 40 or not re.search(r"\d", t):
        return False
    # A bare number is a row, page or location number, never a form number (Nov 17 2026, WS
    # Management: Southlake's Schedule of Locations rows "10   7216 55th Street, Sacramento, CA,
    # 95823" were read as form "10" and added to both GL sections; rows 1-9 escaped only because
    # one digit was too short to match).
    if re.fullmatch(r"\d{1,6}", t):
        return False
    return bool(_FORM_NUMBER_FULL_RE.match(t))


def _split_row(line: str):
    """Return (description, form_number) when `line` is a schedule row, else None."""
    t = line.rstrip()
    if not t.strip():
        return None
    cols = [c for c in _COLUMN_SPLIT_RE.split(t.strip()) if c.strip()]
    if len(cols) >= 2:
        num = cols[-1].strip()
        if _is_form_number(num):
            desc = " ".join(c.strip() for c in cols[:-1])
            return desc, num
        # "FORM NUMBER   DESCRIPTION" layouts (number first)
        num = cols[0].strip()
        if _is_form_number(num) and not _is_form_number(cols[-1].strip()):
            return " ".join(c.strip() for c in cols[1:]), num
        return None
    m = _FORM_NUMBER_TAIL_RE.search(t)
    if m:
        return t[: m.start("num")].strip(), m.group("num").strip()
    return _split_single_spaced_number_first(t)


def _split_single_spaced_number_first(t: str):
    """'IFG-I-0101 10 25 Common Policy Declarations' / 'CG 00 01 04 13 Commercial General ...':
    a number-first row with single spaces, as OCR writes a forms listing (Nov 17 2026, the
    Burlington GL policy read by OCR). The longest leading run of tokens that is a form number
    with a letter in it, followed by a title that starts with a capitalised word."""
    toks = t.strip().split()
    for k in range(min(6, len(toks) - 1), 0, -1):
        num = " ".join(toks[:k])
        rest = " ".join(toks[k:])
        if (re.search(r"[A-Z]", num) and _is_form_number(num)
                and re.match(r"[A-Z][A-Za-z'’.]", rest) and sum(1 for c in rest if c.isalpha()) >= 6):
            return rest, num
    return None


def _looks_like_continuation(line: str) -> bool:
    """A wrapped second line of a form description: short, no form number, no colon,
    single column (Houston Specialty wraps long titles onto a new line)."""
    t = line.strip()
    if not t or len(t) > 80 or ":" in t or "$" in t:
        return False
    if _split_row(t):
        return False
    if _SCHEDULE_HEADER_RE.match(t) or _SCHEDULE_END_RE.match(t) or _COLUMN_HEADER_RE.match(t):
        return False
    letters = [c for c in t if c.isalpha()]
    if len(letters) < 3:
        return False
    if len(_COLUMN_SPLIT_RE.split(t)) > 1:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return upper_ratio > 0.85 or t[0].islower() or t.startswith("(") or t.endswith(")")


# A street address in the title column: a schedule of locations, not of forms
# ("7216 55th Street, Sacramento, CA, 95823", "1100 North Cherry Street Tulare, CA 93274").
_ADDRESS_DESC_RE = re.compile(r"^\d{1,6}\s+\S.*\b[A-Z]{2}\s*,?\s*\d{5}(?:-\d{4})?\s*$")


def _row_desc_ok(desc: str) -> bool:
    """A schedule row's title passes: not empty, not money or a reference number, not
    page furniture (page counter, copyright line), not a street address."""
    if not desc or "$" in desc or len(desc) > 140 or desc.count(".") > 2:
        return False
    if re.search(r"\b(POLICY|QUOTE|REFERENCE|ACCOUNT)\s*(NO|NUMBER|#)", desc, re.I):
        return False
    if _ADDRESS_DESC_RE.match(desc.strip()):
        return False
    return not _FOOTER_DESC_RE.search(desc)


def _schedule_row(line: str):
    """(description, form_number) for a one-line schedule row whose title passes the row
    checks, else None."""
    split = _split_row(line)
    if not split:
        return None
    desc, num = split
    desc = desc.strip(" \t-–—|")
    return (desc, num) if _row_desc_ok(desc) else None


def _lone_form_number(line: str):
    """The form number when a line holds nothing else: the middle line of a title the PDF
    wrapped around its number. Needs a letter, so a bare page or class number never counts."""
    t = (line or "").strip()
    if not t or len(t) > 30 or _COLUMN_SPLIT_RE.search(t) or not re.search(r"[A-Z]", t):
        return None
    return t if _is_form_number(t) else None


def _is_title_fragment(line: str) -> bool:
    """An indented line of title text with no form number of its own: one half of a title
    wrapped around a lone number (Berkley prints the number on the middle line)."""
    if not line or not line.strip():
        return False
    t = line.strip()
    if len(line) - len(line.lstrip(" ")) < 4 or len(t) > 160 or "$" in t or ":" in t:
        return False
    if (_PAGE_MARK_RE.match(t) or _SCHEDULE_HEADER_RE.match(t) or _SCHEDULE_END_RE.match(t)
            or _COLUMN_HEADER_RE.match(t) or _FOOTER_DESC_RE.search(t)):
        return False
    if _COLUMN_SPLIT_RE.search(t) or _split_row(t) or _lone_form_number(t):
        return False
    return sum(1 for c in t if c.isalpha()) >= 3


def _is_prose(text: str) -> bool:
    """A sentence of running text (a notice or policy wording), not a title or a label:
    ten or more words, most of the small ones lower-case."""
    words = (text or "").split()
    if len(words) < 10:
        return False
    low = sum(1 for w in words if w[0].islower())
    return low >= 6 and low >= 0.4 * len(words)


def _schedule_continues(lines: list, start: int, look: int = 15) -> bool:
    """At a page break inside the schedule: does the next page carry on with it? True on
    a schedule header (including "(Cont'd)") or a numbered row within the first `look`
    non-blank lines; False on an end marker, the next page break, or running text that
    comes first (Berkley's TRIA notice after the forms schedule)."""
    seen = 0
    for k in range(start, len(lines)):
        t = lines[k].rstrip()
        if not t.strip():
            continue
        if _PAGE_MARK_RE.match(t) or _SCHEDULE_END_RE.match(t):
            return False
        if _SCHEDULE_HEADER_RE.match(t) or _COLUMN_HEADER_RE.match(t):
            return True
        if _schedule_row(t) or _lone_form_number(t):
            return True
        if _is_prose(t.strip()):
            return False
        seen += 1
        if seen >= look:
            return False
    return False


def parse_forms_schedule(text: str) -> list[dict]:
    """Return [{form_number, description}] for every schedule row in `text`.
    Wrapped titles are joined with the following line, or with the indented halves
    above and below when the number sits alone on the middle line. Page footers are
    never rows, and an explicit schedule block ends at a page break unless the next
    page continues it. Rows found outside an explicit schedule block are accepted only
    when they sit in a run of at least 3 numbered rows within a few lines of each other
    (what a schedule table looks like once the PDF is flattened to text)."""
    if not text:
        return []
    lines = text.replace("\f", "\n--- Page ---\n").splitlines()
    rows = []
    in_block = False
    used = set()   # lines already joined into a row as part of its title
    for i, raw in enumerate(lines):
        line = raw.rstrip()
        if _PAGE_MARK_RE.match(line):
            # The schedule ends at a page break unless the next page carries on with it.
            if in_block and not _schedule_continues(lines, i + 1):
                in_block = False
            continue
        if _SCHEDULE_HEADER_RE.match(line):
            in_block = True
            continue
        if in_block and _SCHEDULE_END_RE.match(line):
            in_block = False
        if _COLUMN_HEADER_RE.match(line) or i in used:
            continue
        row = _schedule_row(line)
        if row:
            desc, num = row
            j = i + 1
            if j < len(lines) and _looks_like_continuation(lines[j]):
                desc = f"{desc} {lines[j].strip()}"
                used.add(j)
        else:
            # Berkley layout: a two-line title wraps AROUND its number, which sits alone
            # on the middle line, with the title halves indented above and below it.
            num = _lone_form_number(line)
            if not num:
                continue
            take = [k for k in (i - 1, i + 1)
                    if 0 <= k < len(lines) and k not in used and _is_title_fragment(lines[k])]
            parts = [lines[k].strip() for k in take]
            if not parts or not all(_row_desc_ok(p) for p in parts):
                continue
            used.update(take)
            desc = " ".join(parts)
        desc = " ".join(desc.split())
        # "FUT 1025 | 01/25 | Vehicle Schedule | Business Auto": the edition belongs with the number
        m_ed = _EDITION_RE.match(desc)
        if m_ed and not re.search(r"\d{2}[/-]\d{2,4}$", num):
            num, desc = f"{num} {m_ed.group(1)}", m_ed.group(2).strip()
        rows.append({"form_number": num, "description": desc, "line_no": i, "in_block": in_block})

    kept, run = [], []
    for r in rows:
        if r["in_block"]:
            if run:
                if len(run) >= 3:
                    kept.extend(run)
                run = []
            kept.append(r)
        else:
            if run and r["line_no"] - run[-1]["line_no"] > 4:
                if len(run) >= 3:
                    kept.extend(run)
                run = []
            run.append(r)
    if len(run) >= 3:
        kept.extend(run)

    seen, out = set(), []
    for r in kept:
        c = _compact(r["form_number"])
        if c in seen:
            continue
        seen.add(c)
        out.append({"form_number": r["form_number"], "description": r["description"], "lob": None})
    _tag_schedule_lobs(out)
    return out


def _tag_schedule_lobs(rows: list) -> bool:
    """Read the Line column off the end of each title when the schedule has one.
    A label counts only when the schedule uses it on at least two rows (or it is an
    unambiguous multi-word label), and the column is accepted only when at least three
    rows and half the schedule carry a label. Sets row['lob'] and trims the title.
    Returns True when the schedule turned out to be line-tagged."""
    hits = []
    for r in rows:
        m = _LOB_TAIL_RE.match(r["description"])
        hits.append((m.group("title"), _LOB_KEY[m.group("lob").lower()], m.group("lob").lower()) if m else None)
    counts = {}
    for h in hits:
        if h:
            counts[h[2]] = counts.get(h[2], 0) + 1
    usable = {lab for lab, n in counts.items() if n >= 2 or lab not in _AMBIGUOUS_LOBS}
    tagged = [h for h in hits if h and h[2] in usable]
    if len(tagged) < 3 or len(tagged) < 0.5 * len(rows):
        return False
    for r, h in zip(rows, hits):
        if h and h[2] in usable and h[0].strip():
            r["description"], r["lob"] = h[0].strip(" -–—|"), h[1]
    # A wrapped title ("Notice Of Cancellation Additional Insureds Per | General Liability" +
    # "Written Contract" on the next line) puts the label mid-string: take the rightmost
    # usable label when only a short, number-free tail follows it.
    for r in rows:
        if r["lob"]:
            continue
        m = _LOB_MID_RE.match(r["description"])
        if m and m.group("lob").lower() in usable and len(m.group("tail")) <= 40 and not re.search(r"[\d:$]", m.group("tail")):
            r["description"] = (m.group("title").strip(" -–—|") + " " + m.group("tail").strip()).strip()
            r["lob"] = _LOB_KEY[m.group("lob").lower()]
    return True


def schedule_is_tagged(rows: list) -> bool:
    n = sum(1 for r in rows if r.get("lob"))
    return n >= 3 and n >= 0.5 * len(rows)


# Forms that can never belong to an auto or liquor section (GL / umbrella / excess forms),
# used when the extractor hands no veto function for those coverages: a prefix test on
# the number plus a token test so carrier-prefixed excess numbers ("HSIC CX ES 01 44")
# are caught too.
_EXCESS_TOKENS = {"CX", "XS", "EX", "CU", "CSXC", "EXL", "UMB", "EXCESS", "UMBRELLA"}
_DEFAULT_REJECT_PREFIXES = {
    "commercial_auto": ("CG ", "CG-", "GLF", "LL ", "LL-", "CYB", "EPL"),
    "liquor": ("CA ", "CA-", "CYB", "EPL"),
}


def _default_reject(coverage_key: str):
    for k, prefixes in _DEFAULT_REJECT_PREFIXES.items():
        if (coverage_key or "").lower().startswith(k):
            def _veto(f, _p=prefixes):
                fn = str(f.get("form_number") or "").upper().strip()
                if not fn:
                    return False
                if fn.startswith(_p):
                    return True
                return any(t in _EXCESS_TOKENS for t in re.split(r"[\s\-/]+", fn))
            return _veto
    return None


_GENERIC_CARRIER_WORDS = {"insurance", "company", "specialty", "group", "limited", "underwriters",
                          "corporation", "national", "american", "mutual", "casualty", "indemnity",
                          "assurance", "surplus", "lines", "certain", "lloyds", "london", "syndicate"}


def _carrier_words(carrier: str) -> list:
    """Distinctive words of a carrier name ('Burlington Insurance Company' -> ['burlington'])."""
    return [t for t in re.findall(r"[a-z0-9]+", (carrier or "").lower().replace("'", ""))
            if len(t) >= 4 and t not in _GENERIC_CARRIER_WORDS and not t.isdigit()]


def _amount(v) -> float:
    try:
        return float(re.sub(r"[^\d.]", "", str(v)) or 0) if not isinstance(v, (int, float)) else float(v)
    except ValueError:
        return 0.0


def _printed_in(value, text: str) -> bool:
    """True when the amount is printed in the text ("$42,559.00", "42,559", "44312.38")."""
    v = _amount(value)
    if v < 100 or not text:
        return False
    forms = {f"{v:,.2f}", f"{v:.2f}"}
    if abs(v - round(v)) < 0.005 and v >= 1000:
        forms.add(f"{v:,.0f}")
    return any(re.search(r"(?<![\d.,])" + re.escape(f) + r"(?!\d|[.,]\d)", text) for f in forms)


def pick_source_text(items: list, carrier: str, coverage_key: str, cov: dict = None) -> tuple[str, str]:
    """Choose the uploaded file that most likely IS this coverage's quote.

    Nov 17 2026 (WS Management): the Burlington GL policy for one hotel had a scrambled text
    layer, so its carrier name was nowhere in its text, and the old keyword score picked the
    Southlake package (117 mentions of "general liability") for it; the Burlington section got
    Southlake's 49 forms. The file is now chosen by what ties it to this coverage, strongest
    first: the file a pass recorded (_source_file / _recovered_from); the one file that prints
    the coverage's premium or total; then, when the carrier has a distinctive name, only files
    whose text or file name carry it (none = no audit, never a guess). Coverage keywords and
    carrier mentions only rank the files that are left."""
    if not items:
        return "", ""
    files = [it for it in items if it.get("text")]
    cov = cov if isinstance(cov, dict) else {}
    src = str(cov.get("_source_file") or cov.get("_recovered_from") or "").strip()
    if src:
        for it in files:
            if (it.get("filename") or "").strip() == src:
                return it["text"], it.get("filename", "")
    words = _carrier_words(carrier)

    def _names_carrier(it):
        low, fname = it["text"].lower(), (it.get("filename") or "").lower()
        return any(w in low or w in fname for w in words)

    figs = [cov.get("total_premium"), cov.get("premium")]
    fig_hits = [it for it in files if any(_printed_in(f, it["text"]) for f in figs)]
    if len(fig_hits) == 1 and (not words or _names_carrier(fig_hits[0])):
        return fig_hits[0]["text"], fig_hits[0].get("filename", "")
    if words:
        named = [it for it in files if _names_carrier(it)]
        if not named:
            return "", ""
        items = [it for it in named if it in fig_hits] or named
    key = (coverage_key or "").lower()
    if key.startswith("umbrella") or key.startswith("excess_liab"):
        kws = ("umbrella", "excess liability", "commercial excess", "following form", "schedule of underlying")
        anti = ("commercial property", "statement of values", "workers compensation")
    elif key.startswith("general_liability"):
        kws = ("general liability", "commercial general liability", "cgl")
        anti = ("excess liability", "commercial excess", "umbrella", "statement of values")
    elif key.startswith("liquor"):
        kws = ("liquor liability",)
        anti = ("umbrella", "commercial property")
    elif key.startswith("commercial_auto"):
        kws = ("business auto", "commercial auto", "auto liability")
        anti = ("umbrella", "commercial property")
    elif key.startswith("property") or key.startswith("excess_property"):
        # Used by the extractor to tell which file is the property quote (Nov 5 2026); the
        # schedule audit itself still covers the liability lines only.
        kws = ("commercial property", "building personal property", "business personal property",
               "building and personal property", "business income", "causes of loss", "coinsurance",
               "named storm", "replacement cost")
        anti = ("commercial general liability", "general liability coverage", "excess liability", "umbrella",
                "workers compensation", "business auto", "statement of values")
    else:
        return "", ""
    carrier_tokens = words
    best, best_score, best_name = "", 0, ""
    for it in items:
        txt = it.get("text") or ""
        if not txt:
            continue
        low = txt.lower()
        score = sum(low.count(k) for k in kws) * 3
        score += sum(min(low.count(t), 25) for t in carrier_tokens) * 2
        score -= sum(low.count(a) for a in anti) * 2
        if score > best_score:
            best, best_score, best_name = txt, score, it.get("filename", "")
    if best_score < 6:
        return "", ""
    return best, best_name


def reconcile_coverage_forms(cov: dict, source_text: str, source_name: str, coverage_key: str,
                             reject_fn=None, lob_allow=None) -> dict:
    """Merge schedule rows the extractor missed into cov['forms_endorsements'].
    `reject_fn(form_dict) -> bool` may veto a candidate (used to keep sibling-coverage
    rows out when one PDF holds two quotes). `lob_allow` is the set of Line-column keys
    (see _LOB_LABELS) that belong to this coverage; on a line-tagged schedule a row tagged
    for another line is skipped outright and a row tagged for this line is accepted even
    when reject_fn would veto it (the carrier's own schedule outranks a prefix rule).
    Returns the audit dict."""
    schedule = parse_forms_schedule(source_text)
    tagged = schedule_is_tagged(schedule) and lob_allow is not None
    existing = cov.get("forms_endorsements") or []
    if not isinstance(existing, list):
        existing = []
    have, have_fuzzy = set(), set()
    for f in existing:
        if isinstance(f, dict):
            c = _compact(f.get("form_number"))
            if c:
                have.add(c)
                have_fuzzy.add(_fuzzy(c))
    added, vetoed, other_line = [], [], 0
    line_rows = 0
    for row in schedule:
        c = _compact(row["form_number"])
        if not c:
            continue
        lob = row.get("lob")
        if tagged and lob and lob not in lob_allow:
            other_line += 1
            continue
        line_rows += 1
        # Edition-date variants: "CX 21 13" vs "CX 21 13 04 13" count as present, and so
        # does a number GPT read with O/0 or I/L/1 swapped ("BR9021 0816" for "BR902I 0816").
        if c in have or _fuzzy(c) in have_fuzzy or any(
                h.startswith(c) or c.startswith(h) for h in have if len(h) >= 6 and len(c) >= 6):
            continue
        cand = {"form_number": row["form_number"], "description": row["description"], "schedule_verified": True}
        carrier_says_ours = bool(tagged and lob and lob in lob_allow)
        if not carrier_says_ours and reject_fn and reject_fn(cand):
            vetoed.append(cand)
            continue
        added.append(cand)
        have.add(c)
        have_fuzzy.add(_fuzzy(c))
    if added:
        merged = existing + added
        # Put the list back in the quote's own schedule order when the schedule
        # accounts for most of it; anything not on the schedule keeps its place at the end.
        order = {_compact(r["form_number"]): i for i, r in enumerate(schedule)}

        def _pos(f):
            if not isinstance(f, dict):
                return len(order) + 1
            c = _compact(f.get("form_number"))
            if c in order:
                return order[c]
            for k, i in order.items():
                if len(k) >= 6 and len(c) >= 6 and (k.startswith(c) or c.startswith(k)):
                    return i
            return len(order) + 1
        if len([f for f in merged if _pos(f) <= len(order)]) >= 0.8 * len(merged):
            merged = sorted(merged, key=_pos)
        cov["forms_endorsements"] = merged
    audit = {
        "source_file": source_name,
        "schedule_rows": len(schedule),
        "line_rows": line_rows if tagged else len(schedule),
        "line_tagged": tagged,
        "other_line_rows": other_line,
        "extracted_before": len(existing),
        "added": [{"form_number": a["form_number"], "description": a["description"]} for a in added],
        "vetoed": [{"form_number": v["form_number"], "description": v["description"]} for v in vetoed],
        "final_count": len(cov.get("forms_endorsements") or []),
    }
    cov["forms_audit"] = audit
    return audit


def run_forms_schedule_audit(data: dict, items: list, reject_fn_factory=None) -> list[str]:
    """Entry point called at the end of extraction. Returns warning strings (also
    appended to data['warnings'])."""
    warnings = []
    covs = data.get("coverages", {}) if isinstance(data, dict) else {}
    if not isinstance(covs, dict):
        return warnings
    targets = [k for k in covs if k.startswith("umbrella") or k.startswith("general_liability")
               or k.startswith("liquor") or k.startswith("commercial_auto")]
    present = {k.lower() for k in covs if isinstance(covs.get(k), dict)}

    def _has(prefixes):
        return any(k.startswith(p) for k in present for p in prefixes)

    def _allow(key):
        """Line-column keys that belong to this coverage. A line with no coverage of its
        own (liquor, EPLI, cyber, crime, ... written as endorsements to the GL policy)
        rides with the general liability section; Interline forms go to every section."""
        k = key.lower()
        if k.startswith("general_liability"):
            allow = {"gl", "interline"}
            for lob, prefixes in (("liquor", ("liquor",)), ("epli", ("epli", "employment")), ("cyber", ("cyber",)),
                                  ("crime", ("crime",)), ("eb", ("equipment_breakdown", "boiler")),
                                  ("inland_marine", ("inland",))):
                if not _has(prefixes):
                    allow.add(lob)
            return allow
        if k.startswith("liquor"):
            return {"liquor", "interline"}
        if k.startswith("commercial_auto"):
            return {"auto", "interline"}
        if k.startswith("umbrella") or k.startswith("excess_liab"):
            return {"umbrella", "interline"}
        return None

    for key in targets:
        cov = covs.get(key)
        if not isinstance(cov, dict):
            continue
        carrier = str(cov.get("carrier") or cov.get("carrier_name") or "")
        text, fname = pick_source_text(items, carrier, key, cov)
        if not text:
            logger.info(f"Forms audit: no distinct source file identified for {key} ({carrier}) - skipped")
            continue
        reject_fn = (reject_fn_factory(key) if reject_fn_factory else None) or _default_reject(key)
        audit = reconcile_coverage_forms(cov, text, fname, key, reject_fn, lob_allow=_allow(key))
        if audit.get("line_tagged"):
            msg = (f"{key}: quote '{fname}' lists {audit['schedule_rows']} numbered forms, "
                   f"{audit['line_rows']} for this line (its Line column, interline included); "
                   f"extractor had {audit['extracted_before']}")
        else:
            msg = (f"{key}: quote '{fname}' lists {audit['schedule_rows']} numbered forms; "
                   f"extractor had {audit['extracted_before']}")
        if audit["added"]:
            names = "; ".join(f"{a['form_number']} {a['description']}" for a in audit["added"][:12])
            more = "" if len(audit["added"]) <= 12 else f" (+{len(audit['added']) - 12} more)"
            w = f"Forms schedule audit — {msg}. ADDED {len(audit['added'])} missing rows: {names}{more}"
            warnings.append(w)
            logger.warning(w)
        else:
            logger.info(f"Forms schedule audit — {msg}. Nothing missing.")
        if audit["vetoed"]:
            names = "; ".join(f"{v['form_number']} {v['description']}" for v in audit["vetoed"][:8])
            w = f"Forms schedule audit — {key}: {len(audit['vetoed'])} schedule rows NOT added because they look like another coverage's forms — verify: {names}"
            warnings.append(w)
            logger.warning(w)
    if warnings:
        data.setdefault("warnings", [])
        if isinstance(data["warnings"], list):
            data["warnings"].extend(warnings)
    return warnings
