"""
terrorism_liability.py - the liability extension of a standalone terrorism policy (Oct 20 2026).

Why this exists (Vista Lodging, AEGIS Terror Shield quote from Lloyd's):
  The Lloyd's / AEGIS quote is 54 pages. Printed page [40] (PDF page 41) is the schedule of the
  "General Liability Extension (Act of Terrorism, Act of Sabotage) (Claims Made and Reported)":
      3. Limit of Liability:  USD 2,000,000 each and every Claim and in the aggregate
                              USD 2,000,000 during the Policy Period.
                              ... part of and not in addition to the Sum Insured ...
                              Defense ... payable within the Limit of Liability ...
      4. Deductible:          USD 5,000 each and every Claim.
      6. Retroactive Date:    30 October 2026
  The proposal showed only the property terrorism limits. In a multi-file run the quote's text
  is cut to its fair share of the 150K-character budget by position, so the model never saw
  page 41, and the terrorism schema had no slot for a liability part anyway.

What it does (deterministic, no AI call):
  find_liability_extension(text) reads the extension schedule from a quote's text (title +
  "Limit of Liability: USD/$ ..." line, deductible, retroactive date, trigger, defense, whether
  the limit sits inside the property sum insured, the reporting window).
  run_terrorism_liability_pass(data, items) attaches it to the extracted terrorism coverage as
  cov["liability_extension"], adds one Coverage Limits row, and writes an info note for the
  review screen. A terrorism file that talks about a liability extension whose schedule could not
  be read gets a REVIEW REQUIRED note instead of silence.
"""

import logging
import re
from datetime import datetime

logger = logging.getLogger(__name__)

_MONEY_RE = re.compile(r"(?:USD|US\$|\$)\s*(\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?", re.I)
_TITLE_RE = re.compile(
    r"(General|Terrorism|Third[\s\-]+Party|Public)\s+Liability\s+Extension"
    r"|Terrorism\s+(?:and\s+Sabotage\s+)?Liability(?!\s+Extension\s+to)",
    re.I,
)
_LIMIT_LINE_RE = re.compile(r"Limits?\s+of\s+Liability\s*:\s*(?=(?:USD|US\$|\$)\s*\d)", re.I)
# a numbered schedule item ("4.   Deductible:") or a new policy section ends a value
_ITEM_END = r"(?=\n\s*\d{1,2}\.\s+[A-Za-z]|\n\s*SECTION\b|\n---\s*Page|\Z)"
_DEDUCTIBLE_RE = re.compile(r"(?:Deductible|Self[\s\-]*Insured\s+Retention|Retention)\s*:\s*(.+?)" + _ITEM_END,
                            re.I | re.S)
_RETRO_RE = re.compile(r"Retro(?:active)?\s+Date\s*:\s*([^\n]+)", re.I)
_PAGE_RE = re.compile(r"---\s*Page\s+(\d+)\s*---")
_TERROR_FILE_RE = re.compile(r"\b(act of terrorism|terror shield|terrorism and sabotage|act of sabotage)\b", re.I)
_NUM_WORDS = {"thirty": 30, "sixty": 60, "ninety": 90, "one hundred twenty": 120, "one hundred and twenty": 120}


def _squash(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _money(n):
    return f"${n:,.0f}"


def _amount(s):
    m = _MONEY_RE.search(s or "")
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def _date(s):
    s = _squash(s).rstrip(".")
    for fmt in ("%d %B %Y", "%B %d, %Y", "%B %d %Y", "%m/%d/%Y", "%d %b %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    m = re.search(r"\d{1,2}/\d{1,2}/\d{4}", s)
    return m.group(0) if m else s


def _clean_title(raw):
    t = re.sub(r"\bENDORSEMENT\b", " ", raw)
    t = _squash(t)
    # stop at the schedule heading / first numbered item
    t = re.split(r"\bSCHEDULE\b|\s\d{1,2}\.\s", t)[0].strip()
    # balance parentheses ("(Act of Terrorism, Act of Sabotage) (Claims Made and Reported)")
    while t.count("(") > t.count(")") and "(" in t:
        t = t[: t.rfind("(")].strip()
    return t


def _page_of(text, pos):
    pages = [(m.start(), int(m.group(1))) for m in _PAGE_RE.finditer(text)]
    page = None
    for start, num in pages:
        if start <= pos:
            page = num
    return page


def find_liability_extension(text):
    """Return the parsed liability extension of a terrorism quote, or None."""
    if not text:
        return None
    for lm in _LIMIT_LINE_RE.finditer(text):
        window_start = max(0, lm.start() - 2500)
        titles = list(_TITLE_RE.finditer(text, window_start, lm.start()))
        if not titles:
            continue
        tm = titles[-1]
        title = _clean_title(text[tm.start(): tm.start() + 300])
        block = text[tm.start(): lm.start() + 6000]
        rel = lm.start() - tm.start()
        value_m = re.compile(r"(.+?)" + _ITEM_END, re.S).match(block, rel + (lm.end() - lm.start()))
        limit_text = _squash(value_m.group(1)) if value_m else ""
        amounts = [float(m.group(1).replace(",", "")) for m in _MONEY_RE.finditer(limit_text)]
        if not amounts:
            continue
        each = amounts[0]
        agg = amounts[1] if len(amounts) > 1 else (each if re.search(r"in the aggregate", limit_text, re.I) else None)
        between = limit_text
        m2 = list(_MONEY_RE.finditer(limit_text))
        if len(m2) > 1:
            between = limit_text[m2[0].end(): m2[1].start()]
        basis_m = re.search(r"(each and every claim|each claim|any one claim|each occurrence|any one occurrence|"
                            r"each loss occurrence|each act)", between, re.I)
        basis = basis_m.group(1).lower() if basis_m else "each claim"

        ded_text = ""
        dm = _DEDUCTIBLE_RE.search(block, rel)
        if dm:
            ded_text = _squash(dm.group(1)).rstrip(".")
        ded_amt = _amount(ded_text)
        ded_basis_m = re.search(r"(each and every claim|each claim|any one claim|each occurrence|any one occurrence)",
                                ded_text, re.I)
        retro = ""
        rm = _RETRO_RE.search(block, rel)
        if rm:
            retro = _date(rm.group(1))

        # The extension's own wording (insuring agreements) follows the schedule.
        body = text[tm.start(): tm.start() + 25000]
        low = _squash(body).lower()
        title_low = title.lower()
        if "claims made and reported" in title_low or "claims made and reported" in low:
            trigger = "Claims made and reported"
        elif "claims made" in title_low or "claims made" in low:
            trigger = "Claims made"
        elif "occurrence" in title_low:
            trigger = "Occurrence"
        else:
            trigger = ""
        report_days = None
        rd = re.search(r"no\s+event\s+later\s+than\s+([a-z\s]+?)\s*\((\d{1,3})\)\s*days\s+after\s+the\s+expiration", low)
        if rd:
            report_days = int(rd.group(2))
        else:
            rd = re.search(r"within\s+(\d{1,3})\s+days\s+(?:after|of)\s+the\s+(?:expiration|end of the policy period)", low)
            if rd:
                report_days = int(rd.group(1))

        within = bool(re.search(r"(payable|included)\s+within\s+the\s+limit\s+of\s+liability|"
                                r"within\s+and\s+not\s+in\s+addition\s+to\s+(?:the|such)\s+limit|"
                                r"included\s+within\s+and\s+not\s+in\s+addition", limit_text + " " + low, re.I))
        no_duty = bool(re.search(r"do(?:es)?\s+not\s+have\s+a\s+duty\s+to\s+defend|no\s+duty\s+to\s+defend|"
                                 r"shall\s+not\s+be\s+called\s+upon\s+to\s+assume\s+the\s+handling", low))
        part_of = bool(re.search(r"part\s+of\s+and\s+not\s+in\s+addition\s+to\s+the\s+sum\s+insured|"
                                 r"part\s+of,?\s+and\s+not\s+in\s+addition\s+to,?\s+the\s+sum\s+insured",
                                 limit_text + " " + low, re.I))
        eroded = bool(re.search(r"impaired\s+by\s+the\s+reduction\s+or\s+exhaustion", low))
        printed = None
        page_no = _page_of(text, lm.start())
        if page_no is not None:
            pm = re.search(r"---\s*Page\s+%d\s*---(.*?)(?=---\s*Page\s+\d+\s*---|\Z)" % page_no, text, re.S)
            if pm:
                bm = re.findall(r"\[\s*(\d{1,3})\s*\]", pm.group(1))
                if bm:
                    printed = int(bm[-1])

        limit_str = f"{_money(each)} {basis}"
        if agg:
            limit_str += f" / {_money(agg)} aggregate"
        ext = {
            "title": title or "Terrorism Liability Extension",
            "limit_each": each,
            "limit_aggregate": agg or 0,
            "limit": limit_str,
            "deductible": (f"{_money(ded_amt)} {ded_basis_m.group(1).lower()}" if ded_amt is not None and ded_basis_m
                           else (_money(ded_amt) if ded_amt is not None else "")),
            # amount + sentence-case basis survive the generator's re-casing of all-lowercase strings
            "deductible_amount": ded_amt or 0,
            "deductible_basis": (ded_basis_m.group(1).capitalize() if ded_basis_m else ""),
            "trigger": trigger,
            "retroactive_date": retro,
            "reporting_days": report_days or 0,
            "defense_within_limits": within,
            "no_duty_to_defend": no_duty,
            "part_of_sum_insured": part_of,
            "eroded_by_property_losses": eroded,
            "source_page": page_no,
            "printed_page": printed,
        }
        return ext
    return None


def describe_terms(ext, sum_insured=""):
    """Rows for the proposal's Terrorism Liability Extension table (label, value)."""
    rows = [["Coverage", ext.get("title") or "Terrorism Liability Extension"]]
    each, agg = ext.get("limit_each") or 0, ext.get("limit_aggregate") or 0
    if each:
        lim = f"{_money(each)} each and every claim"
        if agg and agg == each:
            lim += " and in the aggregate"
        elif agg:
            lim += f"; {_money(agg)} aggregate"
        rows.append(["Limit of Liability", lim])
    elif ext.get("limit"):
        rows.append(["Limit of Liability", ext["limit"]])
    if ext.get("deductible_amount"):
        basis = str(ext.get("deductible_basis") or "").strip().lower()
        rows.append(["Deductible", f"{_money(float(ext['deductible_amount']))}" + (f" {basis}" if basis else "")])
    elif ext.get("deductible"):
        rows.append(["Deductible", ext["deductible"]])
    trig = ext.get("trigger") or ""
    if trig:
        if ext.get("reporting_days") and "claims made" in trig.lower():
            trig += (f"; the claim must be first made during the policy period and reported no later than "
                     f"{ext['reporting_days']} days after expiration")
        rows.append(["Coverage Trigger", trig])
    if ext.get("retroactive_date"):
        rows.append(["Retroactive Date", ext["retroactive_date"]])
    d = []
    if ext.get("defense_within_limits"):
        d.append("Defense, settlement and supplementary payments are paid within, and reduce, the limit")
    if ext.get("no_duty_to_defend"):
        d.append("underwriters have no duty to defend")
    if d:
        rows.append(["Defense", "; ".join(d)])
    if ext.get("part_of_sum_insured"):
        rows.append(["Limit Relationship",
                     "Part of, not in addition to, the policy sum insured"
                     + (f" ({sum_insured})" if sum_insured else "")
                     + ("; property losses paid under the policy reduce the liability limit available"
                        if ext.get("eroded_by_property_losses") else "")])
    return rows


def _terrorism_keys(covs):
    return [k for k, v in covs.items()
            if isinstance(v, dict) and (k == "terrorism" or k.startswith("terrorism_"))]


def _carrier_tokens(carrier):
    stop = {"certain", "underwriters", "at", "lloyd's", "lloyds", "london", "insurance", "company", "the", "of"}
    return [t for t in re.findall(r"[a-z']+", (carrier or "").lower()) if len(t) > 3 and t not in stop]


def run_terrorism_liability_pass(data, items):
    """Attach the terrorism liability extension to the extracted terrorism coverage.
    items: [{"filename", "text"}]. Returns the notes written to data["warnings"]."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict):
        return notes
    keys = _terrorism_keys(covs)
    if not keys:
        return notes
    found = []
    mentions = []
    for it in items or []:
        text = it.get("text") or ""
        if not _TERROR_FILE_RE.search(text):
            continue
        ext = find_liability_extension(text)
        if ext:
            ext["source_file"] = it.get("filename", "")
            found.append((it, ext))
        elif re.search(r"liability\s+extension", text, re.I):
            mentions.append(it.get("filename", ""))
    used = set()
    for key in keys:
        cov = covs[key]
        pick = None
        toks = _carrier_tokens(cov.get("carrier", ""))
        for it, ext in found:
            if id(ext) in used:
                continue
            low = (it.get("text") or "").lower()
            if not toks or any(t in low for t in toks) or len(found) == 1:
                pick = ext
                break
        if not pick:
            continue
        used.add(id(pick))
        prev = cov.get("liability_extension") if isinstance(cov.get("liability_extension"), dict) else {}
        merged = dict(prev)
        merged.update({k: v for k, v in pick.items() if v not in (None, "", 0, False)})
        for flag in ("defense_within_limits", "no_duty_to_defend", "part_of_sum_insured",
                     "eroded_by_property_losses"):
            merged[flag] = bool(pick.get(flag) or prev.get(flag))
        cov["liability_extension"] = merged

        limits = cov.get("limits")
        if not isinstance(limits, list):
            limits = []
        kept = []
        for row in limits:
            d = (row.get("description", "") if isinstance(row, dict) else str(row)).lower()
            if "liability" in d and ("extension" in d or "terrorism liability" in d or "general liability" in d):
                continue   # the model's own version of the same row
            kept.append(row)
        each, agg = pick.get("limit_each") or 0, pick.get("limit_aggregate") or 0
        # Mixed case on purpose (Claim / Sum Insured are the policy's defined terms): the
        # generator re-cases strings that are entirely lower case.
        val = _money(each) + (" each and every Claim and in the aggregate" if agg == each else
                              (f" each and every Claim / {_money(agg)} aggregate" if agg else " each and every Claim"))
        if pick.get("part_of_sum_insured"):
            val += ", part of the Sum Insured"
        kept.append({"description": "Terrorism Liability (General Liability Extension)", "limit": val})
        cov["limits"] = kept

        where = ""
        if pick.get("source_page"):
            where = f"PDF page {pick['source_page']}"
            if pick.get("printed_page") and pick["printed_page"] != pick["source_page"]:
                where += f" (printed page {pick['printed_page']})"
            where += " of "
        note = (f"Terrorism liability extension read from {where}'{pick.get('source_file', '')}': "
                f"{describe_terms(pick)[1][1] if len(describe_terms(pick)) > 1 else pick.get('limit', '')}"
                + (f", deductible {pick['deductible']}" if pick.get("deductible") else "")
                + (f", {pick['trigger'].lower()}" if pick.get("trigger") else "")
                + (f", retroactive date {pick['retroactive_date']}" if pick.get("retroactive_date") else "")
                + ".")
        notes.append(note)
        logger.info(f"Terrorism liability pass [{key}]: {note}")
    if not found and mentions:
        for fn in mentions:
            if not any(isinstance(covs[k].get("liability_extension"), dict) and covs[k]["liability_extension"].get("limit")
                       for k in keys):
                notes.append(f"REVIEW REQUIRED — '{fn}' refers to a terrorism liability extension but its schedule "
                             f"(limit, deductible, retroactive date) was not found in the text read. Check the quote "
                             f"and add the liability limit to the Terrorism section.")
    if notes:
        ws = data.get("warnings")
        if not isinstance(ws, list):
            ws = [ws] if isinstance(ws, str) and ws.strip() else []
        data["warnings"] = ws + notes
    return notes
