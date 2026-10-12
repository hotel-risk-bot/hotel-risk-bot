"""
premium_reconcile.py - premiums that agree with the figures printed on each quote, and the
EPL coverage part of Berkley / Southlake GL quotes shown as its own EPLI line (Oct 12 2026).

Why this exists (Jasmin Hospitality Inc, Berkley GL quote 6812287783):
  The quote prints   Total Premium                                        $113,463.00
                   + Total Surcharges, Taxes and Fees                       $6,636.32
                   = Total including Premium, Surcharges, Taxes and Fees  $120,099.32
  The extraction returned premium $113,338 / total $119,974.32. The model followed the old
  prompt rule "Exclude TRIA premiums from totals" and took Berkley's $125 terrorism charge out
  of both numbers, although Berkley includes it in the advance premium. It also showed the
  $2,255 Employment-Related Practices coverage part as its own EPLI line while leaving it inside
  the GL number, so the program total counted it twice. The web extraction path had no check
  against the printed figures at all.

Stefan's rules (Oct 12 2026):
  * a line's total is the all-in figure printed on its quote;
  * for Berkley and Southlake GL quotes the EPL coverage part is shown as its own EPLI line
    (Southlake prints it as "EPLI (Optional Coverage)" but it sits inside the GL Total Premium);
  * when a line printed INSIDE another quote's premium (EPLI, liquor, terrorism, hired /
    non-owned auto, workplace violence, equipment breakdown inside property ...) is shown as
    its own coverage, the host gives up exactly that premium, so the lines from that quote still
    add up to the printed total (GL $117,844.32 + EPLI $2,255 = $120,099.32).

How it works (deterministic, no AI call):
  1. split_embedded_epli: on a Berkley / Southlake GL quote, find the EPL premium line
     (Berkley "Employment-Related Practices Liability $2,255.00", Southlake "EPLI (Optional
     Coverage) $1,302.00") and its limits / retention; create the EPLI coverage when the model
     did not, and move EPL rows out of the GL limits.
  2. run_premium_reconciliation: read every labelled amount on each quote file; find the printed
     premium + taxes/fees = total (strongest), else premium + total, else a total alone. A line
     counts as inside a printed premium when the premium items printed above it add up to that
     premium (the "summing block"). For every coverage named in the file (carrier match) pick the
     printed figures closest to its extracted premium, then
       - set premium / taxes / total to the printed figures when it is the only main line on
         that quote, and
       - take out of it every carved line (EPLI ...) printed inside its premium.
     Package quotes with several main lines (Southlake GL + liquor + property) keep the model's
     package allocation and only get the carve-out.
  3. cov["premium_check"] on the host, cov["carved_from"] / cov["carved_note"] on carved lines,
     and an info note in data["warnings"] (yellow notes box in the review screen).

Nothing changes when a quote prints no total or no coverage is close enough to its figures.
HotelBound property quotes are skipped (hotelbound_quote.py owns those numbers).
"""

import logging
import re

logger = logging.getLogger(__name__)

# Lines that commonly ride inside another line's quote and premium.
_CARVABLE_BASES = {
    "epli", "liquor_liability", "terrorism", "workplace_violence", "cyber", "crime",
    "commercial_auto", "hired_non_owned_auto", "hnoa", "abuse_molestation", "active_assailant",
    "enviro_pack", "environmental", "pollution", "innkeepers", "innkeepers_liability",
    "garage_keepers", "employee_benefits", "equipment_breakdown", "boiler", "flood",
    "earthquake", "wind", "wind_deductible_buydown", "deductible_buydown",
}

# Which host a carved line can come out of (terrorism: any line).
_GL_RIDERS = {"epli", "liquor_liability", "workplace_violence", "cyber", "crime", "commercial_auto",
              "hired_non_owned_auto", "hnoa", "abuse_molestation", "active_assailant", "enviro_pack",
              "environmental", "pollution", "innkeepers", "innkeepers_liability", "garage_keepers",
              "employee_benefits"}
_PROPERTY_RIDERS = {"equipment_breakdown", "boiler", "flood", "earthquake", "wind",
                    "wind_deductible_buydown", "deductible_buydown"}


def _can_carry(host_key, carved_key):
    hb, cb = _base_key(host_key), _base_key(carved_key)
    if cb in _GL_RIDERS:
        return hb == "general_liability"
    if cb in _PROPERTY_RIDERS:
        return hb in ("property", "excess_property", "excess_property_2")
    return cb == "terrorism"


# Carriers whose GL quotes carry an EPL coverage part that Stefan wants as its own line.
_EPL_SPLIT_CARRIERS = ("berkley", "southlake")

# Words that say a printed amount is NOT part of the quote's premium.
_NOT_INCLUDED_RE = re.compile(
    r"\b(optional|not\s+included|if\s+elected|if\s+accepted|available\s+for|"
    r"additional\s+premium\s+of|may\s+be\s+purchased|excluded\s+from\s+total)\b",
    re.IGNORECASE,
)
_CONDITIONAL_RE = re.compile(r"\b(if|accepted|elected|optional|when)\b", re.IGNORECASE)
_LIMIT_TAIL_RE = re.compile(
    r"\b(aggregate|each|limit|occurrence|claim|accident|person|deductible|retention|exposure|"
    r"basis|per)\b", re.IGNORECASE)
_EPL_RE = re.compile(r"(employment[\s\-]*(related[\s\-]*)?practices?|\bepli\b|\bepl\b)", re.IGNORECASE)

_AMOUNT_RE = re.compile(
    r"(?<![\w.,$])(\$\s?)?(\d{1,3}(?:,\d{3})+|\d+)(\.\d{2})?(?!\d|,\d|\.\d)"
)

_FEE_WORDS = ("tax", "fee", "surcharge", "stamping", "figa", "empa", "fslso", "assessment",
              "inspection", "service charge", "slt")

_GENERIC_CARRIER_WORDS = {
    "insurance", "ins", "company", "co", "corp", "corporation", "inc", "incorporated", "llc",
    "ltd", "limited", "the", "of", "and", "specialty", "surplus", "lines", "national", "group",
    "assurance", "indemnity", "casualty", "underwriters", "syndicate", "es", "excess", "us",
    "usa", "america", "american", "mutual", "fire", "general", "property", "risk", "a", "an",
}

_TOL = {"triple": 0.10, "pair": 0.05, "total": 0.03}
_RANK = {"triple": 0, "pair": 1, "total": 2}

_LABELS = {
    "general_liability": "General Liability", "property": "Property", "umbrella": "Umbrella",
    "epli": "EPLI", "liquor_liability": "Liquor Liability", "terrorism": "Terrorism",
    "commercial_auto": "Auto", "cyber": "Cyber", "crime": "Crime",
    "equipment_breakdown": "Equipment Breakdown", "flood": "Flood",
    "workplace_violence": "Workplace Violence",
}


# --------------------------------------------------------------------------- helpers

def _num(v):
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"-?[\d,]+(?:\.\d+)?", str(v))
    if not m:
        return 0.0
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return 0.0


def _money(v):
    return f"${v:,.2f}"


def _money0(v):
    return f"${v:,.0f}" if abs(v - round(v)) < 0.005 else f"${v:,.2f}"


def _base_key(key):
    return re.sub(r"_alt_\d+$", "", str(key or ""))


def _label(key):
    b = _base_key(key)
    name = _LABELS.get(b, b.replace("_", " ").title())
    if key != b:
        try:
            name += f" (option {int(str(key).rsplit('_', 1)[1]) + 1})"
        except ValueError:
            pass
    return name


def _words(text):
    return set(re.findall(r"[a-z0-9]+", (text or "").lower()))


def _carrier_tokens(name):
    toks = re.findall(r"[a-z0-9]+", (name or "").lower().replace("'", ""))
    distinctive = [t for t in toks if t not in _GENERIC_CARRIER_WORDS and len(t) > 1]
    return distinctive or toks


def _carrier_in_text(carrier, text_words, text_lower):
    toks = _carrier_tokens(carrier)
    if not toks:
        return False
    if all(t in text_words for t in toks):
        return True
    joined = "".join(toks)  # run-together brand names ("StarStone" vs "Star Stone")
    return len(joined) >= 6 and joined in re.sub(r"[^a-z0-9]", "", text_lower)


def _add_warning(data, note):
    ws = data.get("warnings")
    if not isinstance(ws, list):
        ws = [ws] if isinstance(ws, str) and ws.strip() else []
    ws.append(note)
    data["warnings"] = ws


def _label_kind(label):
    """Classify the text in front of an amount.
    'T'  printed all-in total ("Total including Premium, Surcharges, Taxes and Fees",
         "Total Cost of Policy", "Amount Due", "Grand Total" ...)
    'Tw' bare "Total" (used only when premium + fees add up to it exactly)
    'F'  total of taxes / fees ("Total Surcharges, Taxes and Fees")
    'P'  premium ("Total Premium", "Advance Premium", "Policy Premium" ...)
    'X'  terrorism / TRIA charge, or a total that only applies if TRIA is accepted
    'f'  a single tax / fee line
    ''   anything else"""
    l = " " + re.sub(r"\s+", " ", (label or "").lower()).strip() + " "
    has_total = re.search(r"\btotal\b", l) is not None
    fee_word = any(w in l for w in _FEE_WORDS)
    terror = re.search(r"\b(terrorism|tria)\b", l) is not None
    if terror and has_total and _CONDITIONAL_RE.search(l):
        return "X"   # "Total Cost if TRIA is accepted"
    if "grand total" in l or any(w in l for w in ("amount due", "balance due", "amount payable")):
        return "T"
    if has_total and any(w in l for w in (" cost", "total due", "payable", "to be paid")):
        return "T"
    if has_total and fee_word and any(w in l for w in ("premium", " including", " incl", " with ")):
        return "T"
    if has_total and fee_word:
        return "F"
    if re.fullmatch(r" (total|totals|total amount|total policy|total annual|policy total)\s?:? ", l):
        return "Tw"
    if terror and "exclu" not in l and not has_total:
        return "X"
    if fee_word:
        return "f"
    if "premium" in l and not any(w in l for w in (
            "minimum", "earned", "retained", " rate", " per ", "audit", "commission", "broker",
            "credit", "return", "deposit premium of")):
        return "P"
    return ""


def _parse_amounts(text):
    """Every dollar amount in the text with the label in front of it (same line, or the
    previous line when a long label wrapped) and the words right after it."""
    out = []
    lines = (text or "").splitlines()
    prev_plain = ""
    for li, line in enumerate(lines):
        kept = []
        for m in _AMOUNT_RE.finditer(line):
            dollar, whole, cents = m.group(1), m.group(2), m.group(3)
            if not (dollar or cents or "," in whole):
                continue  # bare integers are years, class codes, zips
            try:
                val = float(whole.replace(",", "") + (cents or ""))
            except ValueError:
                continue
            if val <= 0:
                continue
            kept.append((m, val, bool(cents)))
        cursor = 0
        for i, (m, val, has_cents) in enumerate(kept):
            own = re.sub(r"[\s:.\-_$=+]+$", "", line[cursor:m.start()]).strip()
            nxt = kept[i + 1][0].start() if i + 1 < len(kept) else len(line)
            tail = line[m.end():nxt].strip()
            cursor = m.end()
            label = own
            if (len(re.sub(r"[^A-Za-z]", "", own)) < 3 or
                    ("total" not in own.lower() and "total" in prev_plain.lower())):
                if prev_plain:
                    label = (prev_plain + " " + own).strip()
            out.append({"value": round(val, 2), "label": label, "own": own, "tail": tail,
                        "order": (li, m.start()), "line": li, "col": m.start(),
                        "cents": has_cents, "context": line, "kind": _label_kind(label)})
        stripped = line.strip()
        if stripped and not kept:
            prev_plain = stripped if len(stripped) <= 100 else ""
        elif stripped:
            prev_plain = ""
    return out


def _printed_entries(value, amounts):
    return [a for a in amounts if abs(a["value"] - value) < 0.005]


def _figure_sets(amounts):
    """Candidate figure sets printed in one file:
    triple  premium + taxes/fees = total, every figure printed (a "Total ... Taxes and Fees"
            line, or the single tax/fee lines between the premium and the total add up); a
            terrorism charge listed between them is folded into the premium when the numbers
            only add up with it
    pair    a printed premium and an all-in total, fees implied
    total   an all-in total on its own"""
    T = [a for a in amounts if a["kind"] in ("T", "Tw")]
    P = [a for a in amounts if a["kind"] == "P"]
    F = [a for a in amounts if a["kind"] == "F"]
    sets, seen = [], set()

    def _add(mode, pv, fv, tv):
        key = (pv, tv) if mode != "total" else ("t", tv)
        if key in seen:
            return
        seen.add(key)
        sets.append({"mode": mode, "P": pv, "F": fv, "T": tv})

    for t in T:
        for p in P:
            if not (p["value"] < t["value"] <= p["value"] * 1.6):
                continue
            hit = None
            for f in F:
                if abs(p["value"] + f["value"] - t["value"]) <= 0.02:
                    hit = (p["value"], f["value"])
                    break
            if hit is None:
                lo, hi = sorted((p["order"], t["order"]))
                between = [a for a in amounts if lo < a["order"] < hi]
                fees = round(sum(a["value"] for a in between if a["kind"] == "f"), 2)
                tria = round(sum(a["value"] for a in between if a["kind"] == "X"
                                 and not _NOT_INCLUDED_RE.search(a["context"] or "")), 2)
                if fees and abs(p["value"] + fees - t["value"]) <= 0.02:
                    hit = (p["value"], fees)
                elif fees and tria and abs(p["value"] + tria + fees - t["value"]) <= 0.02:
                    hit = (round(p["value"] + tria, 2), fees)
            if hit is not None:
                _add("triple", hit[0], hit[1], t["value"])
    for t in T:
        if t["kind"] != "T":
            continue
        for p in P:
            if p["value"] < t["value"] <= p["value"] * 1.35 and not any(
                    s["T"] == t["value"] and s["mode"] == "triple" for s in sets):
                _add("pair", p["value"], round(t["value"] - p["value"], 2), t["value"])
    for t in T:
        if t["kind"] == "T":
            _add("total", None, None, t["value"])
    return sets


def _summing_block(amounts, p_entry):
    """The premium items printed above a premium figure that add up to it (Berkley: class
    premiums + additional coverages above "Total Advance Premium"; Southlake: GL Premium +
    optional coverages above "TOTAL PREMIUM"). Items carry cents and have no limit wording
    after them; limits, exposures and rating bases are skipped. None when nothing adds up."""
    P = p_entry["value"]
    tol = max(1.0, P * 0.0005)
    idx = amounts.index(p_entry)
    total, block = 0.0, []
    for a in reversed(amounts[:idx]):
        if p_entry["line"] - a["line"] > 80:
            break
        if a["kind"] in ("T", "Tw", "F", "f"):
            break  # crossed into another premium / fee block
        if not a["cents"] or a["value"] >= P - tol or _LIMIT_TAIL_RE.search(a["tail"] or ""):
            continue  # limits, exposures, rating bases, a repeat of the premium itself
        total += a["value"]
        block.append(a)
        if abs(total - P) <= tol:
            return block if len(block) >= 2 else None
        if total > P + tol:
            return None
    return None


class _FileFigures:
    def __init__(self, fname, text):
        self.fname = fname
        self.text = text
        self.lower = text.lower()
        self.words = _words(self.lower)
        self.amounts = _parse_amounts(text)
        self.sets = _figure_sets(self.amounts)
        self._blocks = {}

    def blocks(self, P):
        if P not in self._blocks:
            out = []
            for a in self.amounts:
                if a["kind"] == "P" and abs(a["value"] - P) < 0.005:
                    b = _summing_block(self.amounts, a)
                    if b:
                        out.append(b)
            self._blocks[P] = out
        return self._blocks[P]

    def inside(self, value, P):
        """'block' when the amount is one of the items adding up to premium P, 'no' when P has
        such items and the amount is not among them, 'unverified' when P has none."""
        bl = self.blocks(P)
        if not bl:
            return "unverified"
        for b in bl:
            if any(abs(e["value"] - value) < 0.005 for e in b):
                return "block"
        return "no"

    def has_own_set(self, value):
        return any(s["P"] is not None and abs(s["P"] - value) < 0.005 for s in self.sets)

    def printed_optional(self, value):
        ents = _printed_entries(value, self.amounts)
        return bool(ents) and all(_NOT_INCLUDED_RE.search(e["context"] or "") for e in ents)


def _items_to_files(items):
    files = []
    for it in items or []:
        fname = (it.get("filename") or "").strip()
        text = it.get("text") or ""
        if not isinstance(text, str) or len(text) < 200:
            continue
        if fname.lower().endswith((".xlsx", ".xls", ".xlsb", ".csv")):
            continue
        files.append(_FileFigures(fname, text))
    return files


def _attributed(covs, ff):
    out = []
    for key, cov in covs.items():
        if not isinstance(cov, dict) or not str(cov.get("carrier") or "").strip():
            continue
        hb = cov.get("hotelbound")
        if isinstance(hb, dict) and hb.get("detected"):
            continue
        src = cov.get("_source_file") or cov.get("_recovered_from")
        if src and str(src).strip() and str(src).strip() != ff.fname:
            continue
        if _carrier_in_text(cov.get("carrier"), ff.words, ff.lower):
            out.append(key)
    return out


# --------------------------------------------------------------------------- EPLI split

def _epl_premium_entry(ff, gl_premium):
    """The EPL premium line on a Berkley / Southlake GL quote, or None."""
    best = None
    for a in ff.amounts:
        if not a["cents"] or a["value"] <= 0:
            continue
        if not _EPL_RE.search(a["own"] or a["label"] or ""):
            continue
        if _LIMIT_TAIL_RE.search(a["tail"] or ""):
            continue
        if gl_premium and a["value"] >= gl_premium * 0.5:
            continue
        if re.search(r"\b(exclu|form)\b", (a["own"] or "").lower()):
            continue
        # Prefer an entry that sits inside a premium that adds up (summing block)
        score = 0
        for s in ff.sets:
            if s["P"] and s["P"] > a["value"] and ff.inside(a["value"], s["P"]) == "block":
                score = 2
                break
        if best is None or score > best[0]:
            best = (score, a)
    return best[1] if best else None


def _epl_terms(ff, prem):
    """EPL limits / retention printed around the premium line (Berkley schedule rows) or on the
    EPL row of the optional-coverages table (Southlake: limit, retention, premium)."""
    limits, retention = [], None
    lines = ff.text.splitlines()
    for a in ff.amounts:
        if a is prem or not (0 <= a["line"] - prem["line"] <= 8):
            continue
        tail = (a["tail"] or "").lower()
        own = (a["own"] or "").lower()
        if "aggregate" in tail and not any(d == "Aggregate Limit" for d, _ in limits):
            limits.append(("Aggregate Limit", a["value"]))
        elif "each claim" in tail and not any(d == "Each Claim Limit" for d, _ in limits):
            limits.append(("Each Claim Limit", a["value"]))
        elif ("deductible" in own or "retention" in own) and retention is None:
            retention = a["value"]
    if not limits:
        for li, line in enumerate(lines):
            m = _EPL_RE.search(line)
            if not m or "third party" in line.lower():
                continue
            row = [a for a in ff.amounts if a["line"] == li and a["col"] > m.start()]
            if len(row) >= 3 and abs(row[-1]["value"] - prem["value"]) < 0.005:
                limits.append(("EPL Limit of Liability", row[0]["value"]))
                retention = retention or row[1]["value"]
                break
        for a in ff.amounts:
            lab = (a["label"] or "").lower()
            if "third party" in lab and _EPL_RE.search(lab):
                limits.append(("Third Party Discrimination / Harassment", a["value"]))
                break
    rows = [{"description": d, "limit": _money0(v)} for d, v in limits]
    if retention:
        rows.append({"description": "Retention Per Claim", "limit": _money0(retention)})
    return rows


def split_embedded_epli(data, items):
    """Berkley and Southlake GL quotes carry an Employment(-Related) Practices Liability coverage
    part inside the GL premium. Stefan (Oct 12 2026): show it as its own EPLI line and take it
    out of GL. Creates coverages["epli"] when the model did not; run_premium_reconciliation then
    carves it out of the GL premium and total. Returns the notes it added."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict):
        return notes
    files = _items_to_files(items)
    for gl_key in [k for k in list(covs) if _base_key(k) == "general_liability"]:
        gl = covs.get(gl_key)
        if not isinstance(gl, dict):
            continue
        carrier = str(gl.get("carrier") or "")
        if not any(c in carrier.lower() for c in _EPL_SPLIT_CARRIERS):
            continue
        for ff in files:
            if gl_key not in _attributed(covs, ff):
                continue
            prem = _epl_premium_entry(ff, _num(gl.get("premium")))
            if prem is None:
                continue
            value = prem["value"]
            existing = None
            for k, c in covs.items():
                if _base_key(k) == "epli" and isinstance(c, dict):
                    existing = k
                    break
            terms = _epl_terms(ff, prem)
            if existing:
                ep = covs[existing]
                same_carrier = _carrier_tokens(ep.get("carrier")) == _carrier_tokens(carrier)
                if not same_carrier and str(ep.get("carrier") or "").strip():
                    logger.info(f"EPLI split: {existing} is a separate {ep.get('carrier')} quote; "
                                f"{carrier} EPL part (${value:,.2f}) left inside {gl_key}")
                    break
                if abs(_num(ep.get("premium")) - value) > 0.5:
                    logger.info(f"EPLI split: {existing} premium {ep.get('premium')} -> printed {value}")
                    ep["premium"] = value
                    ep["total_premium"] = value
                    ep["taxes_fees"] = 0
                ep.setdefault("carrier", carrier)
                if not ep.get("limits") and terms:
                    ep["limits"] = terms
                ep["_source_file"] = ff.fname
                ep["embedded_in"] = gl_key
                ep["optional"] = False
                epli_key = existing
            else:
                epl_forms = [f for f in (gl.get("forms_endorsements") or [])
                             if isinstance(f, dict) and (_EPL_RE.search(str(f.get("description") or ""))
                                                         or re.match(r"\s*(EPL|PN0009|BR5027|BR5100EP)",
                                                                     str(f.get("form_number") or ""), re.I))]
                covs["epli"] = {
                    "carrier": carrier,
                    "carrier_admitted": gl.get("carrier_admitted", False),
                    "am_best_rating": gl.get("am_best_rating", ""),
                    "premium": value,
                    "taxes_fees": 0,
                    "total_premium": value,
                    "limits": terms,
                    "forms_endorsements": [dict(f) for f in epl_forms],
                    "subjectivities": [],
                    "policy_period": gl.get("policy_period", ""),
                    "_source_file": ff.fname,
                    "embedded_in": gl_key,
                }
                epli_key = "epli"
                logger.info(f"EPLI split: created epli from {gl_key} ({carrier}) premium {value} "
                            f"terms {terms} forms {len(epl_forms)}")
            # EPL rows now live in the EPLI section, not in the GL limits table
            for lk in ("limits", "additional_coverages", "optional_coverages"):
                rows = gl.get(lk)
                if isinstance(rows, list):
                    kept = [r for r in rows if not (isinstance(r, dict) and _EPL_RE.search(
                        str(r.get("description") or r.get("coverage") or r.get("name") or "")))]
                    if len(kept) != len(rows):
                        gl[lk] = kept
            # The coverage-recovery note "epli is carried inside ... general liability" is now stale
            ws = data.get("warnings")
            if isinstance(ws, list):
                data["warnings"] = [w for w in ws if not (
                    isinstance(w, str) and w.startswith("Coverage check") and " epli is carried inside" in w)]
            notes.append(epli_key)
            break
    return notes


# --------------------------------------------------------------------------- reconciliation

def run_premium_reconciliation(data, items):
    """Reconcile coverage premiums against each quote file's printed figures and carve
    embedded lines out of their host. Returns the notes added to data["warnings"]."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict) or not covs or not items:
        return notes

    candidates = []  # (score tuple, file, host key, set, carved keys, mode flags)
    for ff in _items_to_files(items):
        if not ff.sets:
            continue
        attributed = _attributed(covs, ff)
        if not attributed:
            continue
        for h in attributed:
            host = covs[h]
            g = _num(host.get("premium"))
            gt = _num(host.get("total_premium"))
            for s in ff.sets:
                carved = []
                if s["P"] is not None:
                    for c in attributed:
                        if c == h or _base_key(c) not in _CARVABLE_BASES or not _can_carry(h, c):
                            continue
                        cp = _num(covs[c].get("premium"))
                        if cp <= 0 or cp >= s["P"] or not _printed_entries(cp, ff.amounts):
                            continue
                        where = ff.inside(cp, s["P"])
                        if where == "block" or (
                                where == "unverified" and not covs[c].get("optional")
                                and not ff.has_own_set(cp) and not ff.printed_optional(cp)):
                            carved.append(c)
                c_prem = sum(_num(covs[c].get("premium")) for c in carved)
                if s["mode"] in ("triple", "pair"):
                    if g <= 0:
                        continue
                    d_incl = abs(g - s["P"]) / s["P"]
                    d_excl = abs(g + c_prem - s["P"]) / s["P"]
                else:
                    if gt <= 0:
                        continue
                    c_tot = sum(max(_num(covs[c].get("total_premium")), _num(covs[c].get("premium")))
                                for c in carved)
                    d_incl = abs(gt - s["T"]) / s["T"]
                    d_excl = abs(gt + c_tot - s["T"]) / s["T"]
                d = min(d_incl, d_excl)
                if d > _TOL[s["mode"]]:
                    continue
                score = (_RANK[s["mode"]], round(d, 3), -s["T"])
                candidates.append((score, ff, h, s, carved, d_incl <= d_excl))

    candidates.sort(key=lambda x: x[0])
    used_hosts, carved_done = set(), set()
    for score, ff, h, s, carved, host_had_carved in candidates:
        if h in used_hosts or h in carved_done:
            continue
        host = covs.get(h)
        if not isinstance(host, dict):
            continue
        carved = [c for c in carved if c not in carved_done and c not in used_hosts]
        attributed = _attributed(covs, ff)
        # Another main line priced in this same quote -> package: keep the model's allocation,
        # only carve.
        others = [m for m in attributed if m != h and m not in carved
                  and (_printed_entries(_num(covs[m].get("premium")), ff.amounts)
                       or ff.has_own_set(_num(covs[m].get("premium"))))
                  and _num(covs[m].get("premium")) > 0]
        package = bool(others)
        used_hosts.add(h)
        old_p, old_t = _num(host.get("premium")), _num(host.get("total_premium"))

        carved_rows = []
        c_prem = c_tot = 0.0
        for ck in carved:
            c = covs[ck]
            cp = round(_num(c.get("premium")), 2)
            ct = round(_num(c.get("total_premium")), 2)
            if not (ct >= cp and _printed_entries(ct, ff.amounts) and ct < s["T"]):
                ct = cp
            c["premium"], c["total_premium"], c["taxes_fees"] = cp, ct, round(ct - cp, 2)
            c["optional"] = False
            c["carved_from"] = h
            c["carved_note"] = (f"Part of the {host.get('carrier') or ''} {_label(h)} quote "
                                f"({ff.fname}); its premium is included in that quote's printed "
                                f"premium, so it is taken out of {_label(h)}.").replace("  ", " ")
            carved_rows.append({"key": ck, "premium": cp, "total_premium": ct})
            c_prem += cp
            c_tot += ct
            carved_done.add(ck)

        if package:
            if not carved_rows:
                logger.info(f"Premium check: '{ff.fname}' prices {[h] + others} together (package) - "
                            f"{h} left as extracted")
                continue
            if host_had_carved:
                new_p, new_t = round(old_p - c_prem, 2), round(old_t - c_tot, 2)
            else:
                new_p, new_t = round(old_p, 2), round(old_t, 2)
        elif s["mode"] in ("triple", "pair"):
            new_p = round(s["P"] - c_prem, 2)
            new_t = round(s["T"] - c_tot, 2)
        else:
            new_t = round(s["T"] - c_tot, 2)
            new_p = round(old_p - c_prem, 2) if (host_had_carved and c_prem) else round(old_p, 2)
            if new_p <= 0 or new_p > new_t:
                new_p = round(old_p, 2) if 0 < old_p <= new_t else new_t
        if new_p <= 0 or new_t <= 0 or new_t < new_p:
            logger.warning(f"Premium check: skipped {h} - figures did not reconcile "
                           f"(premium {new_p}, total {new_t})")
            continue

        changed = abs(new_p - old_p) > 0.009 or abs(new_t - old_t) > 0.009
        host["premium"], host["total_premium"] = new_p, new_t
        host["taxes_fees"] = round(new_t - new_p, 2)
        host["premium_check"] = {
            "source_file": ff.fname, "mode": "package carve-out" if package else s["mode"],
            "printed_premium": s["P"], "printed_taxes_fees": s["F"], "printed_total": s["T"],
            "extracted_premium": round(old_p, 2), "extracted_total": round(old_t, 2),
            "carved": carved_rows,
        }
        logger.info(f"Premium check [{h}] '{ff.fname}' ({host['premium_check']['mode']}): premium "
                    f"{old_p} -> {new_p}, total {old_t} -> {new_t}, carved {carved_rows}")
        if not changed and not carved_rows:
            continue

        if s["mode"] == "triple":
            printed = (f"prints premium {_money(s['P'])} + taxes/fees {_money(s['F'])} = "
                       f"{_money(s['T'])}")
        elif s["mode"] == "pair":
            printed = f"prints premium {_money(s['P'])} and total {_money(s['T'])}"
        else:
            printed = f"prints a total of {_money(s['T'])}"
        if carved_rows:
            parts = ", ".join(f"{_label(r['key'])} {_money(r['total_premium'])}" for r in carved_rows)
            together = round(new_t + c_tot, 2)
            note = (f"Premium check: '{ff.fname}' {printed}. {parts} is part of that quote and "
                    f"shows as its own line, so {_label(h)} carries {_money(new_t)} "
                    f"({_money(new_p)} premium); together they equal {_money(together)}"
                    f"{'' if package else ', the printed total'}. Removing the separate line gives "
                    f"its premium back to {_label(h)}.")
        else:
            note = (f"Premium check: {_label(h)} set to {_money(new_t)} total "
                    f"({_money(new_p)} premium); '{ff.fname}' {printed}. "
                    f"The extraction had {_money(old_t)} total ({_money(old_p)} premium).")
        notes.append(note)
        _add_warning(data, note)

    # An EPLI line split out of a GL quote that could not be carved (no figures matched) is
    # still inside the GL number - say so instead of silently double counting.
    for k, c in covs.items():
        if (isinstance(c, dict) and c.get("embedded_in") and not c.get("carved_from")
                and isinstance(covs.get(c["embedded_in"]), dict)):
            note = (f"REVIEW REQUIRED - {_label(k)} {_money(_num(c.get('premium')))} comes from the "
                    f"{_label(c['embedded_in'])} quote but its premium could not be matched to that "
                    f"quote's printed figures. Check that {_label(c['embedded_in'])} does not also "
                    f"include it.")
            notes.append(note)
            _add_warning(data, note)
    return notes


def run_premium_pass(data, items):
    """EPLI split for Berkley / Southlake GL quotes, then printed-figure reconciliation."""
    try:
        split_embedded_epli(data, items)
    except Exception as e:  # never block an extraction
        logger.warning(f"EPLI split failed (non-fatal): {e}")
    return run_premium_reconciliation(data, items)
