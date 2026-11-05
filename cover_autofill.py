"""
cover_autofill.py  -  fill a cover page WITHOUT {placeholders}
================================================================
Reads the cover-page template the way a person would: finds the field labels
("BOQ Item No.:-", "BOQ Item Description:-", "Proposed Make:-", "Model:-" ...),
works out what each one means, and writes the matching value into the empty
cell next to it (to the right, else below, else right after the label in the
same cell). Anything it doesn't recognise (Customer Name, Address, Approval
Status ...) is left exactly as you typed it.

Fields understood:  BOQ reference / S.No, Description, Make, Model
plus any other BOQ column whose header equals the label (e.g. "Qty", "Unit").
Works alongside {placeholders}: cells that already contain text are never touched.
"""

import re

import docx
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph

FIELD_RULES = [  # order matters: "BOQ Item Description" must be a description, not an item no.
    ("description", re.compile(r"descri|particular|scope of|item detail")),
    ("model", re.compile(r"\bmodel\b|part (no|number)|catalog|product code")),
    ("make", re.compile(r"\bmake\b|manufactur|\bbrand\b|\boem\b")),
    ("sno", re.compile(r"^(boq|item|sr|s|sl|serial|ref)\s*(item)?\s*(no|number|#|ref)\b|^boq\b|^reference\b|^ref\b")),
]
FIELD_TITLES = {"sno": "BOQ ref / S.No", "description": "Description", "make": "Make", "model": "Model no."}


def norm_label(text):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9#/ ]+", " ", (text or "").lower())).strip()


def is_empty(text):
    return re.sub(r"[\s:\-\u2013\u2014_.]", "", text or "") == ""


def classify(text, extra_keys=()):
    """Return a field name for a label-like text, else None."""
    t = (text or "").strip()
    if not t or len(t) > 50:
        return None
    n = norm_label(t)
    if not n or len(n.split()) > 6:
        return None
    for field, rx in FIELD_RULES:
        if rx.search(n):
            return field
    if n in extra_keys:
        return "extra:" + n
    return None


def _unique_cells(row):
    seen, out = set(), []
    for c in row.cells:
        if id(c._tc) not in seen:
            seen.add(id(c._tc))
            out.append(c)
    return out


def _iter_tables(container):
    for t in container.tables:
        yield t
        for row in t.rows:
            for c in _unique_cells(row):
                yield from _iter_tables(c)


def _set_text(paragraph, text, style_run=None):
    """Write text (with real line breaks) into a paragraph as a new run copying the label's font."""
    run = paragraph.add_run()
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i:
            run.add_break()
        run.add_text(line)
    if style_run is not None:
        if style_run.font.name:
            run.font.name = style_run.font.name
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is not None:
                rfonts.set(qn("w:eastAsia"), style_run.font.name)
        if style_run.font.size:
            run.font.size = style_run.font.size
    return run


def _label_style_run(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            if r.text.strip():
                return r
    return None


def autofill(document, values, extra=None):
    """values: {'sno','description','make','model'} -> str.  extra: {normalised BOQ header: value}.
    Returns a list of (label_text, field, how) for everything that was filled."""
    extra = extra or {}
    filled = []
    for table in _iter_tables(document):
        rows = list(table.rows)
        grid = [_unique_cells(r) for r in rows]
        for ri, cells in enumerate(grid):
            for ci, cell in enumerate(cells):
                field = classify(cell.text, extra.keys())
                if not field:
                    continue
                value = extra.get(field[6:], "") if field.startswith("extra:") else values.get(field, "")
                if value in (None, ""):
                    continue
                target, how = None, ""
                right = cells[ci + 1] if ci + 1 < len(cells) else None
                if right is not None and is_empty(right.text) and not classify(right.text, extra.keys()):
                    target, how = right, "right of label"
                elif ri + 1 < len(grid) and ci < len(grid[ri + 1]):
                    below = grid[ri + 1][ci]
                    if is_empty(below.text) and not classify(below.text, extra.keys()):
                        target, how = below, "below label"
                style = _label_style_run(cell)
                if target is not None:
                    para = target.paragraphs[0]
                    _set_text(para, value, style)
                    filled.append((cell.text.strip(), field, how))
                elif re.search(r"[:\-\u2013]\s*$", cell.text.strip()) and len(cell.paragraphs) == 1:
                    _set_text(cell.paragraphs[0], " " + str(value), style)
                    filled.append((cell.text.strip(), field, "after label"))
    # labels written as plain paragraphs ("Model : ")
    for p in document.paragraphs:
        t = p.text.strip()
        field = classify(t.rstrip(":-\u2013 "), extra.keys()) if re.search(r"[:\-\u2013]\s*$", t) else None
        if field:
            value = extra.get(field[6:], "") if field.startswith("extra:") else values.get(field, "")
            if value:
                style = p.runs[0] if p.runs else None
                _set_text(p, " " + str(value), style)
                filled.append((t, field, "after label"))
    return filled


def analyze_template(template_path, extra_keys=()):
    """What does the cover page contain?  -> dict for display in the UI."""
    d = docx.Document(template_path)
    recognised, other, has_placeholders = [], [], False
    for table in _iter_tables(d):
        for row in table.rows:
            cells = _unique_cells(row)
            for ci, cell in enumerate(cells):
                t = cell.text.strip()
                if re.search(r"\{[^{}]+\}", t):
                    has_placeholders = True
                field = classify(t, extra_keys)
                if field:
                    recognised.append((t, FIELD_TITLES.get(field, field.replace("extra:", "BOQ column: "))))
                elif t and ci + 1 < len(cells) and len(t) <= 40 and re.search(r"[:\-\u2013]\s*$", t):
                    other.append(t)
    return {"recognised": recognised, "other": other, "has_placeholders": has_placeholders}
