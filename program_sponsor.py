"""
program_sponsor.py - keep a program sponsor's name off the client's named insureds (Oct 20 2026).

Why this exists (Vista Lodging, Great Point Insurance member umbrella quote):
  Member umbrella programs print the sponsor that holds the master policy as the "Named Insured"
  ("Named Insured: Paramount Real Estate Group, Inc." ... "Insured: Vista Lodging LLC", "offered
  through the Paramount Real Estate Group, Inc.", "Member Coverage Period"). The extraction took
  the sponsor as the umbrella's named insured and the proposal listed "Paramount Real Estate
  Group, Inc." as Named Insured #2. Stefan (Oct 20 2026): list only the client's names, never the
  umbrella corporation.

What it does (deterministic, no AI call):
  detect_program_sponsors(items) reads every quote for sponsor wording; is_program_sponsor(name)
  matches a name against them (and a short list of known sponsors); scrub_program_sponsors(data)
  removes the sponsor from named_insureds / additional_named_insureds / additional_insureds,
  from every coverage's named-insured fields (a sponsor named_insured becomes the member insured)
  and from client_info, and records data["program_sponsors"] so the generator can filter again.
"""

import logging
import re

logger = logging.getLogger(__name__)

# Sponsors seen on HUB placements. Names only - never the client's own entities.
KNOWN_SPONSORS = (
    "Paramount Real Estate Group, Inc.",
    "Paramount Programs, Inc.",
)

_ENTITY = (r"(?P<name>[A-Z][A-Za-z0-9&.,'\- ]{2,90}?(?:,?\s+(?:Inc\.?|Incorporated|LLC|L\.L\.C\.|Corp\.?|Corporation|"
           r"Company|Ltd\.?|Trust|Association|RRG|Risk Retention Group|Purchasing Group)))")
_PHRASES = (
    re.compile(r"offered\s+through\s+(?:the\s+)?" + _ENTITY, re.S),
    re.compile(r"Members?\s+of\s+(?:the\s+)?" + _ENTITY + r"\s+may\s+not\s+be\s+protected", re.S),
    re.compile(r"captioned\s+Insured\s+under\s+(?:the\s+)?" + _ENTITY, re.S),
    re.compile(r"(?:member|participant)s?\s+of\s+(?:the\s+)?" + _ENTITY + r"\s+(?:program|purchasing group|master policy)",
               re.S | re.I),
)
_NAMED_INSURED_RE = re.compile(r"Named\s+Insured\s*:\s*([^\n]+)", re.I)
_MEMBER_INSURED_RE = re.compile(r"(?:^|\n)\s*(?:Insured|Member|Participant)\s*:\s*([^\n]+)", re.I)
_MEMBER_MARKERS = re.compile(r"member\s+coverage\s+period|participant\s+coverage\s+period|members?\s+of\s+the\s+",
                             re.I)
_SUFFIX_RE = re.compile(r"\b(inc|incorporated|llc|l l c|corp|corporation|company|co|ltd|limited|the)\b")


def _norm(s):
    t = re.sub(r"[^a-z0-9 ]+", " ", str(s or "").lower())
    t = _SUFFIX_RE.sub(" ", t)
    return re.sub(r"\s+", " ", t).strip()


def _clean(s):
    s = re.sub(r"\s+", " ", str(s or "")).strip(" ,;:")
    if s.endswith(".") and not re.search(r"\b(Inc|Corp|Co|Ltd)\.$", s):
        s = s.rstrip(".").strip()
    return s


def detect_program_sponsors(items):
    """Sponsor names found in the quote texts, with the member insured each one sponsors."""
    found = {}
    for it in items or []:
        text = it.get("text") or ""
        if not text:
            continue
        names = set()
        for rx in _PHRASES:
            for m in rx.finditer(text):
                names.add(_clean(m.group("name")))
        # Structure: "Named Insured: <sponsor>" plus a separate "Insured: <member>" and member wording.
        if _MEMBER_MARKERS.search(text):
            ni = [_clean(m.group(1)) for m in _NAMED_INSURED_RE.finditer(text)]
            mem = [_clean(m.group(1)) for m in _MEMBER_INSURED_RE.finditer(text)]
            mem_n = {_norm(x) for x in mem if x}
            for n in ni:
                if n and _norm(n) not in mem_n and mem_n:
                    names.add(n)
        member = ""
        mm = _MEMBER_INSURED_RE.search(text)
        if mm:
            member = _clean(mm.group(1))
        for n in names:
            if len(_norm(n)) < 6:
                continue
            if member and _norm(n) == _norm(member):
                continue
            found.setdefault(n, {"member": member, "file": it.get("filename", "")})
    return found


def is_program_sponsor(name, sponsors=None):
    n = _norm(name if not isinstance(name, dict) else (name.get("name") or ""))
    if len(n) < 6:
        return False
    pool = list(KNOWN_SPONSORS) + list(sponsors or [])
    for sp in pool:
        s = _norm(sp)
        if not s:
            continue
        # exact, or the candidate starts with the sponsor's name ("... Group, Inc. (Program Sponsor)").
        # Never a bare substring: "Sunrise Hotel Group LLC" must not match a sponsor "Hotel Group, Inc."
        if n == s or (len(s) >= 10 and n.startswith(s + " ")):
            return True
    return False


def _filter_list(lst, sponsors, where, removed):
    if not isinstance(lst, list):
        return lst
    out = []
    for x in lst:
        nm = x.get("name") if isinstance(x, dict) else x
        if nm and is_program_sponsor(nm, sponsors):
            removed.append((where, _clean(nm)))
            continue
        out.append(x)
    return out


def scrub_program_sponsors(data, sponsors=None, member=""):
    """Remove program sponsors from the client's named-insured data. Returns [(where, name)]."""
    if not isinstance(data, dict):
        return []
    sponsors = list(dict.fromkeys(list(sponsors or []) + list(data.get("program_sponsors") or [])))
    removed = []
    for key in ("named_insureds", "additional_named_insureds", "additional_insureds"):
        if key in data:
            data[key] = _filter_list(data.get(key), sponsors, key, removed)
    ci = data.get("client_info")
    client = ""
    if isinstance(ci, dict):
        client = ci.get("named_insured") or ""
        if client and is_program_sponsor(client, sponsors):
            removed.append(("client_info", _clean(client)))
            ci["named_insured"] = member or data.get("client_name") or ""
            client = ci["named_insured"]
    covs = data.get("coverages")
    if isinstance(covs, dict):
        for ck, cov in covs.items():
            if not isinstance(cov, dict):
                continue
            ni = cov.get("named_insured")
            if isinstance(ni, str) and ni and is_program_sponsor(ni, sponsors):
                removed.append((f"{ck}.named_insured", _clean(ni)))
                cov["program_sponsor"] = _clean(ni)
                cov["named_insured"] = member or client
            for lk in ("named_insureds", "additional_named_insureds", "additional_insureds"):
                if lk in cov:
                    cov[lk] = _filter_list(cov.get(lk), sponsors, f"{ck}.{lk}", removed)
    if sponsors or removed:
        known = list(data.get("program_sponsors") or [])
        for s in sponsors + [n for _, n in removed]:
            if s and not any(_norm(s) == _norm(k) for k in known):
                known.append(s)
        data["program_sponsors"] = known
    return removed


def run_program_sponsor_pass(data, items):
    """Detect sponsors in the quote files and scrub them. Returns notes for data['warnings']."""
    found = detect_program_sponsors(items)
    member = ""
    for info in found.values():
        if info.get("member"):
            member = info["member"]
            break
    removed = scrub_program_sponsors(data, list(found.keys()), member)
    notes = []
    names = sorted({n for _, n in removed})
    if names:
        files = sorted({info["file"] for info in found.values() if info.get("file")})
        notes.append("Program sponsor left off the Named Insureds: " + "; ".join(names)
                     + (f" (the master-policy holder on '{', '.join(files)}'; the client is the member insured)"
                        if files else "") + ".")
        logger.info(f"Program sponsor pass: removed {removed}")
        ws = data.get("warnings")
        if not isinstance(ws, list):
            ws = [ws] if isinstance(ws, str) and ws.strip() else []
        data["warnings"] = ws + notes
    return notes
