"""
workers_comp_quote.py - workers compensation quotes read from the printed figures (Oct 20 2026).

Why this exists (Vista Lodging, Markel Insurance Company workers comp quote from mPortal):
  The quote prints one class row and a premium build-up:
      State Code Description                      Payroll     Net Rate  PR Net Rate  Rate   Premium
      TX    9052 Hotel - All Other Employees ...  $350,000.00 0.430*    0.504        0.679  $2,377
      Liability Limit: 1000/1000/1000 $150 · TX - Premium Incentives for Small Employers -0.150 -$379
      TX - Schedule Rating -0.300 -$644 · TX E-MOD 0.000 $0 · Premium Discount $0 · Expense Constant $140
      Terrorism Charge $84 · Catastrophe $35 · Final Premium $1,763.00
  plus a payment plan table and two notices (TRIA disclosure, a California fraud warning).
  The proposal showed the carrier as Non-Admitted (workers comp is written admitted), no
  statutory Part One row, none of the rating steps, six separate payment rows that each read
  "Down Payment 100%; Down Payment $1,763", and the California fraud notice as the only "form"
  on a Texas policy.

What it does (deterministic, no AI call):
  parse_wc_quote(text) reads the class rows (the manual rate is the column that reproduces the
  row premium: payroll / 100 x rate), the premium build-up to the final premium, the employers
  liability limits, the payment plans and the policy period.
  run_workers_comp_pass(data, items) applies it to every extracted workers comp coverage:
  admitted unless the quote says surplus lines / non-admitted, a "Workers Compensation -
  Statutory (ST)" limit row, rating_basis with the row premium, premium_development (the
  build-up, rendered by the generator), one consolidated payment row per WC quote, and no
  out-of-state notices in forms_endorsements.
"""

import logging
import re

logger = logging.getLogger(__name__)

_STATES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California", "CO": "Colorado",
    "CT": "Connecticut", "DE": "Delaware", "DC": "District of Columbia", "FL": "Florida", "GA": "Georgia",
    "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland", "MA": "Massachusetts",
    "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri", "MT": "Montana",
    "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico",
    "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota",
    "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont", "VA": "Virginia", "WA": "Washington",
    "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
}

_WC_FILE_RE = re.compile(r"workers['’]?\s*compensation", re.I)
_MONEY = r"-?\$\s?-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?|-?\$\s?\d+(?:\.\d{2})?"
_CLASS_ROW_RE = re.compile(
    r"^\s*(?P<state>[A-Z]{2})\s+(?P<code>\d{4})(?:\s*-\s*\d{2})?\b(?P<desc>.*?)"
    r"\$\s?(?P<payroll>\d{1,3}(?:,\d{3})+(?:\.\d{2})?|\d+(?:\.\d{2})?)\s+"
    r"(?P<nums>(?:-?\d*\.\d+\*?\s+){0,4}-?\d*\.\d+\*?)\s+"
    r"\$\s?(?P<premium>\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*$"
)
_STEP_RE = re.compile(
    r"^\s*(?P<label>[A-Za-z][A-Za-z0-9 ,&/()'.\-]{2,90}?)\s*:?\s{2,}"
    r"(?:(?P<factor>[\d/.\-]+%?)\s{2,})?(?P<amount>" + _MONEY + r")\s*$"
)
_FINAL_RE = re.compile(r"^\s*(final premium|total (?:estimated )?premium|estimated annual premium|"
                       r"total estimated annual premium|total policy premium)\b", re.I)
_PLAN_RE = re.compile(r"^\s*(?P<plan>[A-Za-z0-9][A-Za-z0-9 ,\-]{1,60}?)\s+(?P<pct>\d{1,3})%\s+\$\s?(?P<amt>[\d,]+\.\d{2})\s*$")
_EL_RE = re.compile(r"Liability\s+Limits?\s*:?\s+(\d{3,4})\s*/\s*(\d{3,4})\s*/\s*(\d{3,4})", re.I)
_PERIOD_RE = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})\s*(?:-|to|through)\s*(\d{1,2}/\d{1,2}/\d{4})", re.I)
_EFF_RE = re.compile(r"Effective\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})", re.I)
_CARRIER_RE = re.compile(r"quote\s+for\s+the\s+captioned\s+insured\s+with\s+([A-Z][A-Za-z&.,'\s]+?(?:Company|Corporation|Inc\.?))",
                         re.I)
_NONADMITTED_RE = re.compile(r"\bnon[\s\-]?admitted\b|\bsurplus\s+lines?\b|\bexcess\s+and\s+surplus\b", re.I)


def _num(s):
    try:
        return float(str(s).replace("$", "").replace(",", "").replace(" ", ""))
    except ValueError:
        return None


def _money(v, cents=False):
    if v is None:
        return ""
    neg = v < 0
    s = f"${abs(v):,.2f}" if cents else f"${abs(v):,.0f}"
    return ("-" if neg else "") + s


def _is_text_only(line):
    s = line.strip()
    if not s or "$" in s or re.search(r"\d{3,}", s):
        return False
    if re.match(r"^[A-Z]{2}\s+\d{4}\b", s):
        return False
    if re.search(r"\b(rate|payroll|premium|description|state code)\b", s, re.I):
        return False
    return len(s) <= 90


def _factor_display(label, factor, amount):
    if not factor:
        return ""
    f = factor.strip()
    if re.fullmatch(r"\d{3,4}/\d{3,4}/\d{3,4}", f):
        def _short(x):
            v = int(x) * 1000
            return f"${v // 1000000}M" if v % 1000000 == 0 else f"${v // 1000}K"
        return " / ".join(_short(x) for x in f.split("/"))
    if re.search(r"e-?mod|experience", label, re.I):
        v = _num(f)
        return "None" if (v is not None and v == 0) else f
    if f.endswith("%"):
        return f
    v = _num(f)
    if v is not None and -1 < v < 1 and v != 0 and "." in f:
        return f"{v * 100:+.0f}%".replace("+", "") if v < 0 else f"+{v * 100:.0f}%"
    return f


def _label_display(label):
    l = re.sub(r"\s+", " ", label).strip().rstrip(":")
    st = None
    m = re.match(r"^([A-Z]{2})\s*-\s*(.+)$", l)
    if m and m.group(1) in _STATES:
        st, l = m.group(1), m.group(2)
    m = re.match(r"^([A-Z]{2})\s+(E-?MOD.*)$", l, re.I)
    if m and m.group(1).upper() in _STATES:
        st, l = m.group(1).upper(), m.group(2)
    low = l.lower()
    if re.search(r"e-?mod\b|experience mod", low):
        l = "Experience Modification"
    elif low.startswith("liability limit"):
        l = "Employers Liability Increased Limits"
    elif low.startswith("terrorism"):
        l = "Terrorism (TRIA)"
    elif low.startswith("catastrophe"):
        l = "Catastrophe (Other Than Certified Acts of Terrorism)"
    elif low.startswith("premium incentives for small employers"):
        l = "Premium Incentive for Small Employers"
    return f"{l} ({st})" if st else l


def parse_wc_quote(text):
    """Read a workers comp quote's class rows, premium build-up, EL limits, payment plans and term."""
    if not text or not _WC_FILE_RE.search(text):
        return None
    lines = text.splitlines()
    rows, steps, plans = [], [], []
    final = None
    in_steps = False
    in_plans = False
    for i, line in enumerate(lines):
        m = _CLASS_ROW_RE.match(line)
        if m:
            payroll = _num(m.group("payroll"))
            prem = _num(m.group("premium"))
            nums = [n for n in m.group("nums").split()]
            rate = None
            net = []
            for n in nums:
                v = _num(n.rstrip("*"))
                if v is None:
                    continue
                if rate is None and payroll and prem and abs(payroll / 100.0 * v - prem) <= max(1.5, prem * 0.01):
                    rate = n.rstrip("*")
                else:
                    net.append(n)
            if rate is None and nums:
                rate = nums[-1].rstrip("*")
            desc = m.group("desc").strip()
            if i > 0 and _is_text_only(lines[i - 1]) and not lines[i - 1].strip().endswith("."):
                desc = (lines[i - 1].strip() + " " + desc).strip()
            if i + 1 < len(lines) and _is_text_only(lines[i + 1]) and not _STEP_RE.match(lines[i + 1]):
                nxt = lines[i + 1].strip()
                if not re.match(r"^(Liability Limit|Premium|Expense|Terrorism|Catastrophe|Final|Total)\b", nxt, re.I):
                    desc = (desc + " " + nxt).strip()
            rows.append({
                "state": m.group("state"), "class_code": m.group("code"),
                "classification": re.sub(r"\s+", " ", desc).strip(" ,"),
                "payroll": _money(payroll), "rate": rate, "premium": _money(prem),
                "_payroll": payroll, "_premium": prem, "_net_rates": net,
            })
            in_steps = True
            continue
        if in_steps:
            if _FINAL_RE.match(line):
                am = re.search(_MONEY, line)
                if am:
                    final = _num(am.group(0))
                in_steps = False
                continue
            sm = _STEP_RE.match(line)
            if sm and not line.strip().startswith("$"):
                amt = _num(sm.group("amount"))
                steps.append({"label": sm.group("label").strip(), "factor": sm.group("factor") or "", "amount": amt})
        if re.search(r"payment\s+plan", line, re.I):
            in_plans = True
            continue
        if in_plans:
            pm = _PLAN_RE.match(line)
            if pm:
                plans.append({"plan": re.sub(r"\s+", " ", pm.group("plan")).strip(), "pct": int(pm.group("pct")),
                              "amount": _num(pm.group("amt"))})
            elif plans and line.strip() and not re.search(r"payment|down", line, re.I):
                in_plans = False
    if not rows and final is None:
        return None
    el = _EL_RE.search(text)
    period = None
    pm = _PERIOD_RE.search(text)
    if pm:
        period = f"{pm.group(1)} - {pm.group(2)}"
    eff = _EFF_RE.search(text)
    cm = _CARRIER_RE.search(text)
    return {
        "carrier": re.sub(r"\s+", " ", cm.group(1)).strip() if cm else "",
        "class_rows": rows,
        "steps": steps,
        "final_premium": final,
        "el_limits": [int(x) * 1000 for x in el.groups()] if el else [],
        "payment_plans": plans,
        "policy_period": period,
        "effective_date": eff.group(1) if eff else "",
        "states": sorted({r["state"] for r in rows}),
        "non_admitted_text": bool(_NONADMITTED_RE.search(text)),
    }


def premium_development(parsed):
    """Rows for the proposal's Premium Development table: [item, factor, amount]."""
    out = []
    rows = parsed.get("class_rows") or []
    manual = sum(r.get("_premium") or 0 for r in rows)
    if rows:
        if len(rows) == 1:
            r = rows[0]
            item = f"Manual Premium ({r['payroll']} payroll x {r['rate']})"
        else:
            item = "Manual Premium (all classes)"
        out.append({"item": item, "factor": "", "amount": _money(manual, cents=True)})
    for s in parsed.get("steps") or []:
        out.append({"item": _label_display(s["label"]),
                    "factor": _factor_display(s["label"], s.get("factor"), s.get("amount")),
                    "amount": _money(s.get("amount"), cents=True)})
    if parsed.get("final_premium") is not None:
        out.append({"item": "Final Premium", "factor": "", "amount": _money(parsed["final_premium"], cents=True)})
    return out


def payment_terms(parsed):
    lines = []
    for p in parsed.get("payment_plans") or []:
        plan, pct, amt = p["plan"], p["pct"], p["amount"]
        if pct == 100:
            lines.append(f"{plan}: {_money(amt, cents=True)} in full")
        elif not pct or not amt:
            lines.append(f"{plan}: no down payment")
        else:
            lines.append(f"{plan}: {_money(amt, cents=True)} down ({pct}%)")
    return "\n".join(lines)


def _wc_keys(covs):
    return [k for k, v in covs.items() if isinstance(v, dict) and (k.startswith("workers_comp"))]


def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _state_names_in(text):
    low = (text or "").lower()
    return {ab for ab, nm in _STATES.items() if re.search(r"\b" + re.escape(nm.lower()) + r"\b", low)}


def run_workers_comp_pass(data, items):
    """Apply the parsed WC quote(s) to the extracted workers comp coverage(s). Returns notes."""
    notes = []
    covs = data.get("coverages") if isinstance(data, dict) else None
    if not isinstance(covs, dict):
        return notes
    keys = _wc_keys(covs)
    if not keys:
        return notes
    parsed_files = []
    for it in items or []:
        p = parse_wc_quote(it.get("text") or "")
        if p and (p["class_rows"] or p["final_premium"] is not None):
            p["source_file"] = it.get("filename", "")
            parsed_files.append(p)
    used = set()
    for key in keys:
        cov = covs[key]
        carrier_n = _norm(cov.get("carrier"))
        pick = None
        for p in parsed_files:
            if id(p) in used:
                continue
            pc = _norm(p.get("carrier"))
            if (carrier_n and pc and (carrier_n in pc or pc in carrier_n)) or len(parsed_files) == 1:
                pick = p
                break
        changes = []
        # 1. Admitted: workers compensation is written by admitted (licensed) carriers.
        if pick is None or not pick.get("non_admitted_text"):
            if cov.get("carrier_admitted") is not True:
                cov["carrier_admitted"] = True
                changes.append("admitted")
        if pick is None:
            continue
        used.add(id(pick))
        states = pick.get("states") or []
        # 2. Limits: statutory Part One + the EL limits printed on the quote.
        limits = cov.get("limits") if isinstance(cov.get("limits"), list) else []
        has_stat = any("statutory" in str((r.get("limit") if isinstance(r, dict) else r) or "").lower()
                       or "workers comp" in str((r.get("description") if isinstance(r, dict) else r) or "").lower()
                       for r in limits)
        if not has_stat:
            st = ", ".join(states)
            limits = [{"description": "Workers Compensation (Part One)",
                       "limit": "Statutory" + (f" ({st})" if st else "")}] + limits
            changes.append("statutory row")
        if pick.get("el_limits") and len(pick["el_limits"]) == 3:
            labels = ("EL - Each Accident", "EL - Disease Policy Limit", "EL - Disease Each Employee")
            have = {_norm(r.get("description")) for r in limits if isinstance(r, dict)}
            for lab, amt in zip(labels, pick["el_limits"]):
                if _norm(lab) not in have:
                    limits.append({"description": lab, "limit": f"${amt:,}"})
        cov["limits"] = limits
        # 3. Rating basis with the row premium and the manual rate.
        if pick.get("class_rows"):
            cov["rating_basis"] = [{"state": r["state"], "location": "", "class_code": r["class_code"],
                                    "classification": r["classification"], "payroll": r["payroll"],
                                    "rate": r["rate"], "premium": r["premium"]} for r in pick["class_rows"]]
            changes.append("rating basis")
        # 4. Premium build-up.
        dev = premium_development(pick)
        if dev:
            cov["premium_development"] = dev
            changes.append("premium development")
        fp = pick.get("final_premium")
        if fp:
            if abs((float(cov.get("total_premium") or 0)) - fp) > 0.5 or abs(float(cov.get("premium") or 0) - fp) > 0.5:
                changes.append(f"premium {cov.get('total_premium') or cov.get('premium')} -> {fp:,.2f}")
            cov["premium"] = fp
            cov["total_premium"] = fp
            cov["taxes_fees"] = 0
        if pick.get("policy_period") and not cov.get("policy_period"):
            cov["policy_period"] = pick["policy_period"]
        # 5. Forms: no notices for states this policy does not cover.
        forms = cov.get("forms_endorsements") if isinstance(cov.get("forms_endorsements"), list) else []
        kept = []
        for f in forms:
            desc = f"{f.get('form_number', '')} {f.get('description', '')}" if isinstance(f, dict) else str(f)
            named = _state_names_in(desc)
            if named and states and not (named & set(states)) and re.search(
                    r"notice|warning|advisory|disclosure", desc, re.I):
                changes.append(f"dropped out-of-state notice '{desc.strip()}'")
                continue
            kept.append(f)
        cov["forms_endorsements"] = kept
        cov["rating_states"] = states
        # 6. One payment row for this quote.
        terms = payment_terms(pick)
        if terms:
            pos = data.get("payment_options")
            if not isinstance(pos, list):
                pos = [pos] if pos else []
            carrier = cov.get("carrier") or pick.get("carrier") or ""
            cn = _norm(carrier)
            rest = []
            for po in pos:
                if isinstance(po, dict):
                    pcn = _norm(po.get("carrier"))
                    ctype = str(po.get("coverage_type") or "").lower()
                    same_carrier = bool(cn and pcn and (cn in pcn or pcn in cn))
                    if same_carrier and ("work" in ctype or "comp" in ctype or not ctype):
                        continue
                rest.append(po)
            rest.append({"carrier": carrier, "coverage_type": "Workers Compensation", "terms": terms, "mep": ""})
            data["payment_options"] = rest
            changes.append("payment plans")
        if changes:
            notes.append(f"Workers comp ('{pick.get('source_file', '')}'): read from the quote's printed figures — "
                         + ", ".join(changes) + ".")
            logger.info(f"Workers comp pass [{key}]: {', '.join(changes)}")
    if notes:
        ws = data.get("warnings")
        if not isinstance(ws, list):
            ws = [ws] if isinstance(ws, str) and ws.strip() else []
        data["warnings"] = ws + notes
    return notes
