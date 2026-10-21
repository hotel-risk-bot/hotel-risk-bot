"""
Appearance streams for filled ACORD form fields.

MuPDF can only draw widget text in regular Helvetica (any other /DA font name is
mapped back to Helv), and 8 pt regular Helvetica reads faint next to the bold
ACORD labels in most viewers. This module writes its own appearance streams:

  * text fields in Helvetica-Bold (base-14, no embedding), sized to fit the box
    (single line: shrink to 5.5 pt; multi-line: wrap, shrink, then cut with " ..."),
    left / centre / right per the field's /Q, and the /DA set to the same bold font
    so a viewer that regenerates the appearance after an edit keeps the look;
  * checked boxes as a drawn check mark (no ZapfDingbats dependency) with /MK /CA
    set to the heavy check so regenerated appearances match.

    from acord_appearance import write_text_appearance, write_check_appearance
    size, text = write_text_appearance(doc, widget, text, base_size)   # after widget.update()
    write_check_appearance(doc, widget)                                 # after widget.update()
"""
import re

import pymupdf

BOLD_NAME = "HeBo"                      # resource / DA name
BOLD_METRICS = "hebo"                   # pymupdf.get_text_length font name
BOLD_FONT_OBJ = "<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>"
MIN_SIZE = 5.5
LINE_HEIGHT = 1.16                      # multi-line leading as a multiple of the font size
MULTILINE_FLAG = 1 << 12
COMB_FLAG = 1 << 24


# ─── metrics ────────────────────────────────────────────────────────────────

def text_width(text, size):
    return pymupdf.get_text_length(str(text), fontname=BOLD_METRICS, fontsize=size)


def wrap(text, size, width):
    """Greedy word wrap on the bold metrics; a word longer than the width is cut."""
    words, lines, cur = str(text).split(" "), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if text_width(t, size) > width and cur:
            lines.append(cur)
            cur = w
        else:
            cur = t
    if cur:
        lines.append(cur)
    out = []
    for ln in lines:
        while text_width(ln, size) > width and len(ln) > 1:
            cut = ln
            while text_width(cut, size) > width and len(cut) > 1:
                cut = cut[:-1]
            out.append(cut)
            ln = ln[len(cut):]
        out.append(ln)
    return out or [""]


def fit_single(text, width, base, floor=MIN_SIZE):
    """Largest size (<= base) at which the text fits on one line."""
    size = base
    while size > floor and text_width(text, size) > width - 4:
        size -= 0.5
    return size


def fit_multi(text, width, height, base, floor=MIN_SIZE):
    """(size, text) for a multi-line box: shrink until the wrapped text fits, then cut
    with ' ...' at the smallest size."""
    size = base
    while size >= floor:
        if len(wrap(text, size, width - 4)) * size * LINE_HEIGHT <= height - 2:
            return size, text
        size -= 0.5
    size = floor
    max_lines = max(1, int((height - 2) / (size * LINE_HEIGHT)))
    words = str(text).split(" ")
    while words:
        words = words[:-1]
        cand = " ".join(words) + " ..."
        if len(wrap(cand, size, width - 4)) <= max_lines:
            return size, cand
    return size, ""


# ─── PDF plumbing ───────────────────────────────────────────────────────────

def pdf_escape(text):
    """WinAnsi-encoded PDF literal string body (parentheses, backslashes and
    non-ASCII bytes escaped)."""
    out = []
    for b in str(text).encode("cp1252", errors="replace"):
        if b in (0x28, 0x29, 0x5C):
            out.append("\\" + chr(b))
        elif 32 <= b < 127:
            out.append(chr(b))
        else:
            out.append("\\%03o" % b)
    return "".join(out)


def _acroform(doc):
    """(xref, key prefix) addressing the AcroForm dictionary, which may be an
    indirect object or a direct dictionary inside the catalog."""
    cat = doc.pdf_catalog()
    af = doc.xref_get_key(cat, "AcroForm")
    if af[0] == "xref":
        return int(af[1].split()[0]), ""
    return cat, "AcroForm/"


def bold_font_xref(doc):
    """The Helvetica-Bold font object registered in the AcroForm /DR (created once)."""
    cached = getattr(doc, "_acord_bold_xref", None)
    if cached:
        return cached
    ax, pre = _acroform(doc)
    cur = doc.xref_get_key(ax, pre + "DR/Font/" + BOLD_NAME)
    if cur[0] == "xref":
        fx = int(cur[1].split()[0])
    else:
        fx = doc.get_new_xref()
        doc.update_object(fx, BOLD_FONT_OBJ)
        dr = doc.xref_get_key(ax, pre + "DR")
        if dr[0] != "dict":
            doc.xref_set_key(ax, pre + "DR", "<</Font<</%s %d 0 R>>>>" % (BOLD_NAME, fx))
        elif doc.xref_get_key(ax, pre + "DR/Font")[0] != "dict":
            doc.xref_set_key(ax, pre + "DR/Font", "<</%s %d 0 R>>" % (BOLD_NAME, fx))
        else:
            doc.xref_set_key(ax, pre + "DR/Font/" + BOLD_NAME, "%d 0 R" % fx)
    doc._acord_bold_xref = fx
    return fx


def alignment(doc, widget):
    """The field's /Q (0 left, 1 centre, 2 right), inherited from the parent field
    when the widget itself carries none (PyMuPDF's text_format is not reliable)."""
    xref = widget.xref
    for _ in range(4):
        q = doc.xref_get_key(xref, "Q")
        if q[0] == "int":
            return int(q[1])
        par = doc.xref_get_key(xref, "Parent")
        if par[0] != "xref":
            break
        xref = int(par[1].split()[0])
    return 0


def _ap_xref(doc, widget, state=None):
    """xref of the widget's normal appearance stream (or of one state of it)."""
    key = "AP/N" + ("/" + state if state else "")
    ap = doc.xref_get_key(widget.xref, key)
    if ap[0] == "xref":
        return int(ap[1].split()[0])
    return None


def write_text_appearance(doc, widget, text, base_size, floor=MIN_SIZE):
    """Replace the widget's appearance with Helvetica-Bold text that fits the box.
    Call after widget.update() (which stores /V). Returns (size used, text used)."""
    text = str(text or "")
    if not text:
        return base_size, text
    rect = widget.rect
    w, h = rect.width, rect.height
    flags = widget.field_flags or 0
    if flags & COMB_FLAG:
        return base_size, text        # comb fields keep MuPDF's appearance
    multiline = bool(flags & MULTILINE_FLAG) and h >= 20
    q = alignment(doc, widget)
    base = base_size or 8.0
    if multiline:
        size, text = fit_multi(text, w, h, base, floor)
        lines = wrap(text, size, w - 4)
        lh = size * LINE_HEIGHT
        y0 = h - 2 - size * 0.95
    else:
        size = fit_single(text, w, base, floor)
        lines = [text]
        lh = 0
        y0 = (h - size) / 2.0 + size * 0.22
    parts = ["/Tx BMC", "q", "1 1 %g %g re W n" % (round(w - 2, 3), round(h - 2, 3)), "BT", "0 g", "/%s %g Tf" % (BOLD_NAME, size)]
    px = py = 0.0
    for i, ln in enumerate(lines):
        tw = text_width(ln, size)
        x = 2.0 if q == 0 else ((w - tw) / 2.0 if q == 1 else w - 2.0 - tw)
        y = y0 - i * lh
        parts.append("%g %g Td" % (round(x - px, 3), round(y - py, 3)))
        parts.append("(%s) Tj" % pdf_escape(ln))
        px, py = x, y
    parts += ["ET", "Q", "EMC"]
    data = "\n".join(parts).encode("latin-1")
    fx = bold_font_xref(doc)
    x = _ap_xref(doc, widget)
    if x is None:
        x = doc.get_new_xref()
        doc.update_object(x, "<</Type/XObject/Subtype/Form/BBox[0 0 %g %g]/Resources<</Font<</%s %d 0 R>>>>>>" % (w, h, BOLD_NAME, fx))
        doc.xref_set_key(widget.xref, "AP", "<</N %d 0 R>>" % x)
    else:
        doc.xref_set_key(x, "BBox", "[0 0 %g %g]" % (w, h))
        doc.xref_set_key(x, "Matrix", "[1 0 0 1 0 0]")
        doc.xref_set_key(x, "Resources", "<</Font<</%s %d 0 R>>>>" % (BOLD_NAME, fx))
    doc.update_stream(x, data)
    doc.xref_set_key(widget.xref, "DA", "(0 g /%s %g Tf)" % (BOLD_NAME, size))
    if text != str(widget.field_value or ""):
        head = widget.xref
        doc.xref_set_key(head, "V", "(%s)" % pdf_escape(text))
    return size, text


def write_check_appearance(doc, widget):
    """Draw the checked state as a check mark path (black, 1.4 pt) and point /MK /CA
    at the heavy check so regenerated appearances match. Call after widget.update()."""
    try:
        on = widget.on_state()
    except Exception:
        on = None
    if not on or on == "Off":
        return False
    x = _ap_xref(doc, widget, str(on))
    rect = widget.rect
    w, h = rect.width, rect.height
    s = min(w, h)
    cx, cy = w / 2.0, h / 2.0
    pts = [(cx - 0.30 * s, cy + 0.02 * s), (cx - 0.08 * s, cy - 0.26 * s), (cx + 0.34 * s, cy + 0.30 * s)]
    data = ("q 0 g 0 G %g w 1 J 1 j %g %g m %g %g l %g %g l S Q" % (
        max(1.0, round(s * 0.12, 2)), *[round(v, 3) for p in pts for v in p])).encode("latin-1")
    if x is None:
        x = doc.get_new_xref()
        doc.update_object(x, "<</Type/XObject/Subtype/Form/BBox[0 0 %g %g]>>" % (w, h))
        n = doc.xref_get_key(widget.xref, "AP/N")
        if n[0] == "dict":
            doc.xref_set_key(widget.xref, "AP/N/" + str(on), "%d 0 R" % x)
        else:
            doc.xref_set_key(widget.xref, "AP", "<</N<</%s %d 0 R>>>>" % (on, x))
    else:
        doc.xref_set_key(x, "BBox", "[0 0 %g %g]" % (w, h))
        doc.xref_set_key(x, "Resources", "<<>>")
    doc.update_stream(x, data)
    doc.xref_set_key(widget.xref, "MK/CA", "(4)")
    return True


def embolden_file(src, dst, floor=MIN_SIZE):
    """Re-draw every filled text field of an existing ACORD PDF in bold and every
    checked box as a drawn check (used to refresh files built before this module)."""
    doc = pymupdf.open(src)
    n = 0
    for page in doc:
        for w in page.widgets():
            if w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT:
                if w.field_name == "Form_EditionIdentifier_A" or not w.field_value:
                    continue
                da = doc.xref_get_key(w.xref, "DA")[1]
                m = re.search(r"/\w+ ([\d.]+) Tf", da or "")
                size = float(m.group(1)) if m else (w.text_fontsize or 8.0)
                if not size:
                    size = 8.0
                write_text_appearance(doc, w, w.field_value, size, floor)
                n += 1
            elif w.field_type == pymupdf.PDF_WIDGET_TYPE_CHECKBOX and w.field_value not in (None, "", False, "Off"):
                write_check_appearance(doc, w)
    doc.save(dst, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=1)
    doc.close()
    return n


if __name__ == "__main__":
    import sys
    print(embolden_file(sys.argv[1], sys.argv[2]), "fields redrawn")
