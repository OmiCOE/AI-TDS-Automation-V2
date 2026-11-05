#!/usr/bin/env python3
"""
TDS Automation Tool  (100% Python - no LibreOffice / no MS Word required)
===========================================================================
Reads a BOQ (Excel), fills a project-specific cover-page template (Word
.docx, with {Column Name} placeholders) for every line item, locates the
matching TDS PDF in a (recursively searched) database folder by exact
model-number filename match, merges cover page + TDS into one PDF, and
writes everything into per-line-item output folders named after the BOQ's
S.No.

The cover page is rendered to PDF using a built-in pure-Python renderer
(python-docx to read + reportlab to draw) - no LibreOffice, no MS Word,
no external programs required. Keyword highlighting in the TDS uses
PyMuPDF, also installed via plain pip.

REQUIREMENTS (install locally - everything via pip, no external programs)
-------------------------------------------------------------------
    pip install openpyxl python-docx pypdf reportlab Pillow pymupdf

RUN
---
    python tds_automation.py
    (a few file/folder picker windows will pop up: BOQ, cover page,
    database folder, output folder)

OUTPUT LAYOUT
-------------
    Output/
      1/
        Cover_Page_1.docx
        TDS_1.pdf              <- cover page + matched TDS (keywords highlighted) merged
      2/
        Cover_Page_2.docx
        TDS_2.pdf
      ...
      missing_tds_log.csv       <- items with no model / no TDS match found
      highlight_summary.csv     <- per-item keyword highlighting results
"""

import csv
import io
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency: pip install openpyxl")

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
except ImportError:
    sys.exit("Missing dependency: pip install python-docx")

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    sys.exit("Missing dependency: pip install pypdf")

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (
        BaseDocTemplate, PageTemplate, Frame, Paragraph, Table, TableStyle,
        Spacer, Image,
    )
except ImportError:
    sys.exit("Missing dependency: pip install reportlab")

try:
    from PIL import Image as PILImage
except ImportError:
    sys.exit("Missing dependency: pip install Pillow")

try:
    import pymupdf
except ImportError:
    sys.exit("Missing dependency: pip install pymupdf")


PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")


# ============================================================================
# SECTION 1: BOQ reading
# ============================================================================

def find_header_row(ws, max_scan_rows=10):
    """Locate the row that holds the column headers (looks for an 'S.No'-like cell)."""
    candidates = {"s.no", "sr.no", "sr no", "s no", "sr.no.", "sno"}
    for row_idx in range(1, max_scan_rows + 1):
        for cell in ws[row_idx]:
            if cell.value and str(cell.value).strip().lower().replace(" ", "") in {
                c.replace(" ", "") for c in candidates
            }:
                return row_idx
    raise ValueError(
        "Could not find a header row containing an 'S.No' column in the first "
        f"{max_scan_rows} rows of the BOQ. Please check the file."
    )


def find_sno_column(headers):
    for idx, h in enumerate(headers):
        if h and str(h).strip().lower().replace(" ", "") in {
            "s.no", "sr.no", "srno", "sno", "s.no."
        }:
            return idx
    raise ValueError("Could not find the S.No column among BOQ headers.")


def format_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_boq(boq_path):
    """Returns (headers list, list of (sno_str, row_dict)). Only real line-item rows returned."""
    wb = openpyxl.load_workbook(boq_path, data_only=True)
    ws = wb.active  # first/active sheet
    header_row_idx = find_header_row(ws)
    raw_headers = [c.value for c in ws[header_row_idx]]
    headers = [str(h).strip() if h else None for h in raw_headers]
    sno_col = find_sno_column(headers)

    items = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        sno_val = row[sno_col] if sno_col < len(row) else None
        try:
            if sno_val is None or str(sno_val).strip() == "":
                continue
            sno_num = float(sno_val)
        except (TypeError, ValueError):
            continue  # section heading / sub-header / blank row -> skip

        sno_str = str(int(sno_num)) if sno_num.is_integer() else str(sno_num)
        row_dict = {}
        for col_idx, header in enumerate(headers):
            if not header:
                continue
            val = row[col_idx] if col_idx < len(row) else None
            row_dict[header] = format_value(val)
        items.append((sno_str, row_dict))

    return headers, items


# ============================================================================
# SECTION 2: Cover page placeholder filling
#            (paragraphs, tables incl. nested, headers/footers)
# ============================================================================

def set_run_text(run, text):
    """Set a run's text, converting '\\n' into real line breaks (<w:br/>)."""
    for t in list(run._element.findall(qn("w:t"))):
        run._element.remove(t)
    for br in list(run._element.findall(qn("w:br"))):
        run._element.remove(br)

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            run._element.append(run._element.makeelement(qn("w:br"), {}))
        t_elem = run._element.makeelement(qn("w:t"), {})
        t_elem.text = line
        t_elem.set(qn("xml:space"), "preserve")
        run._element.append(t_elem)


def replace_in_paragraph(paragraph, mapping):
    runs = paragraph.runs
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if "{" not in full_text:
        return

    def _sub(m):
        key = m.group(1)
        return mapping.get(key, m.group(0))

    new_text = PLACEHOLDER_RE.sub(_sub, full_text)
    if new_text == full_text:
        return

    set_run_text(runs[0], new_text)
    for r in runs[1:]:
        set_run_text(r, "")


def collect_table_paragraphs(table):
    paragraphs = []
    for row in table.rows:
        for cell in row.cells:
            paragraphs.extend(cell.paragraphs)
            for nested_table in cell.tables:
                paragraphs.extend(collect_table_paragraphs(nested_table))
    return paragraphs


def collect_all_paragraphs(document):
    paragraphs = list(document.paragraphs)
    for table in document.tables:
        paragraphs.extend(collect_table_paragraphs(table))
    for section in document.sections:
        for hf in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            paragraphs.extend(hf.paragraphs)
            for table in hf.tables:
                paragraphs.extend(collect_table_paragraphs(table))
    return paragraphs


def fill_cover_page(template_path, mapping, out_path, auto_values=None, extra=None):
    """Fill {placeholders}; then, if auto_values is given, ALSO read the labels on the page
    ("BOQ Item No.:-", "Model:-" ...) and fill the empty cells next to them (cover_autofill)."""
    document = docx.Document(template_path)
    for paragraph in collect_all_paragraphs(document):
        replace_in_paragraph(paragraph, mapping)
    if auto_values:
        import cover_autofill
        cover_autofill.autofill(document, auto_values, extra)
    document.save(out_path)


# ============================================================================
# SECTION 3: Pure-Python DOCX -> PDF renderer (python-docx + reportlab)
#            No LibreOffice / MS Word / subprocess calls.
# ============================================================================

EMU_PER_PT = 12700.0

# Map common Word font names to reportlab's built-in standard-14 families
# (regular, bold, italic, bold-italic) - no font embedding needed.
FONT_FAMILIES = {
    "times new roman": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "times": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "georgia": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "cambria": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
    "courier new": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique"),
    "consolas": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique"),
}
DEFAULT_FAMILY = ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique")


def resolve_font_face(font_name, bold, italic):
    family = FONT_FAMILIES.get((font_name or "").strip().lower(), DEFAULT_FAMILY)
    if bold and italic:
        return family[3]
    if bold:
        return family[1]
    if italic:
        return family[2]
    return family[0]


def emu_to_pt(value):
    return float(value) / EMU_PER_PT


def dxa_to_pt(value):
    """Word table/cell widths are stored in twentieths of a point (dxa)."""
    return float(value) / 20.0


def xml_escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run_font_size_pt(run, default=10.5):
    sz = run.font.size
    return sz.pt if sz is not None else default


def _run_color_hex(run):
    try:
        color = run.font.color
        if color is not None and color.type is not None and color.rgb is not None:
            return "#%s" % str(color.rgb)
    except Exception:
        pass
    return None


def paragraph_alignment(paragraph):
    align = paragraph.alignment
    mapping = {0: TA_LEFT, 1: TA_CENTER, 2: TA_RIGHT, 3: TA_JUSTIFY}
    return mapping.get(int(align), TA_LEFT) if align is not None else TA_LEFT


def _run_segments(run):
    """Yield (text, is_break) segments for a run, honoring <w:br/>, <w:cr/>, <w:tab/>."""
    segments = []
    buf = ""
    for child in run._element:
        tag = child.tag
        if tag == qn("w:t"):
            buf += child.text or ""
        elif tag == qn("w:tab"):
            buf += "    "
        elif tag in (qn("w:br"), qn("w:cr")):
            segments.append((buf, False))
            buf = ""
            segments.append(("", True))
    segments.append((buf, False))
    return segments


def paragraph_to_markup(paragraph):
    """Build a reportlab-markup string from a paragraph's runs."""
    parts = []
    has_text = False
    for run in paragraph.runs:
        for seg_text, is_break in _run_segments(run):
            if is_break:
                parts.append("<br/>")
                continue
            if not seg_text:
                continue
            has_text = True
            seg = xml_escape(seg_text)
            size = _run_font_size_pt(run)
            color = _run_color_hex(run)
            face = resolve_font_face(run.font.name, bool(run.bold), bool(run.italic))
            font_attrs = f'size="{size:.1f}" face="{face}"'
            if color:
                font_attrs += f' color="{color}"'
            seg = f"<font {font_attrs}>{seg}</font>"
            if run.underline:
                seg = f"<u>{seg}</u>"
            parts.append(seg)
    return "".join(parts), has_text


def get_paragraph_style(paragraph, base_name="body"):
    align = paragraph_alignment(paragraph)
    size = 10.5
    face = DEFAULT_FAMILY[0]
    if paragraph.runs:
        size = _run_font_size_pt(paragraph.runs[0])
        face = resolve_font_face(paragraph.runs[0].font.name,
                                  bool(paragraph.runs[0].bold),
                                  bool(paragraph.runs[0].italic))
    return ParagraphStyle(name=base_name, fontName=face, fontSize=size,
                           leading=size * 1.25, alignment=align, spaceAfter=2)


def make_image_flowable(img_bytes, w_pt, h_pt, max_width):
    try:
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        pil_img.load()
    except Exception:
        return None  # unsupported format (e.g. EMF/WMF) - skip gracefully
    if pil_img.mode not in ("RGB", "RGBA", "L"):
        pil_img = pil_img.convert("RGB")
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)

    if not w_pt or not h_pt:
        w_pt, h_pt = pil_img.size

    if w_pt > max_width:
        scale = max_width / w_pt
        w_pt *= scale
        h_pt *= scale

    return Image(buf, width=w_pt, height=h_pt)


def extract_inline_images(paragraph, part, max_width):
    flowables = []
    for run in paragraph.runs:
        for drawing in run._element.findall(".//" + qn("w:drawing")):
            for blip in drawing.findall(".//" + qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                if not rid or rid not in part.rels:
                    continue
                try:
                    img_bytes = part.rels[rid].target_part.blob
                except Exception:
                    continue
                node = drawing.find(qn("wp:inline"))
                if node is None:
                    node = drawing.find(qn("wp:anchor"))
                w_pt = h_pt = None
                if node is not None:
                    extent = node.find(qn("wp:extent"))
                    if extent is not None:
                        w_pt = emu_to_pt(extent.get("cx"))
                        h_pt = emu_to_pt(extent.get("cy"))
                flow = make_image_flowable(img_bytes, w_pt, h_pt, max_width)
                if flow is not None:
                    flowables.append(flow)
    return flowables


def build_paragraph_flowables(paragraph, part, max_width):
    flowables = []
    text_markup, has_text = paragraph_to_markup(paragraph)
    images = extract_inline_images(paragraph, part, max_width)
    if has_text and text_markup.strip():
        flowables.append(Paragraph(text_markup, get_paragraph_style(paragraph)))
    flowables.extend(images)
    if not flowables:
        flowables.append(Spacer(1, 4))
    return flowables


def get_cell_shading_hex(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    if not fill or fill.upper() in ("AUTO", "FFFFFF"):
        return None
    return "#" + fill


def get_table_col_widths_pt(table):
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    widths = []
    if grid is not None:
        for gc in grid.findall(qn("w:gridCol")):
            w = gc.get(qn("w:w"))
            widths.append(dxa_to_pt(w) if w else 72.0)
    if not widths:
        n = len(table.rows[0]._tr.findall(qn("w:tc"))) if table.rows else 1
        widths = [468.0 / max(n, 1)] * max(n, 1)
    return widths


def build_cell_flowables(tc, part, cell_widths_pt):
    width = sum(cell_widths_pt) if cell_widths_pt else 100.0
    flowables = []
    for child in tc:
        if child.tag == qn("w:p"):
            paragraph = DocxParagraph(child, None)
            flowables.extend(build_paragraph_flowables(paragraph, part, width - 8))
        elif child.tag == qn("w:tbl"):
            nested = DocxTable(child, None)
            flowables.append(build_table_flowable(nested, part, width - 8))
    if not flowables:
        flowables.append(Spacer(1, 1))
    return flowables


def build_table_flowable(docx_table, part, available_width):
    tbl = docx_table._tbl
    col_widths_pt = get_table_col_widths_pt(docx_table)
    total_dxa = sum(col_widths_pt) or 1.0
    scale = available_width / total_dxa
    col_widths_pt = [w * scale for w in col_widths_pt]
    n_cols = len(col_widths_pt)

    trs = tbl.findall(qn("w:tr"))
    n_rows = len(trs)

    grid = [[None for _ in range(n_cols)] for _ in range(n_rows)]
    vmerge_anchor = {}

    for r, tr in enumerate(trs):
        c = 0
        for tc in tr.findall(qn("w:tc")):
            while c < n_cols and grid[r][c] == "SKIP":
                c += 1
            if c >= n_cols:
                break

            tcPr = tc.find(qn("w:tcPr"))
            grid_span = 1
            vmerge_val = None
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    grid_span = int(gs.get(qn("w:val")) or "1")
                vm = tcPr.find(qn("w:vMerge"))
                if vm is not None:
                    vmerge_val = vm.get(qn("w:val")) or "continue"

            if vmerge_val == "continue" and c in vmerge_anchor:
                anchor_r = vmerge_anchor[c]
                grid[anchor_r][c]["rowspan"] += 1
                for cc in range(c, min(c + grid_span, n_cols)):
                    grid[r][cc] = "SKIP"
                c += grid_span
                continue

            content_flowables = build_cell_flowables(tc, part, col_widths_pt[c:c + grid_span])
            shading = get_cell_shading_hex(tc)
            cell_info = {"flowables": content_flowables, "rowspan": 1,
                         "colspan": grid_span, "shading": shading}
            grid[r][c] = cell_info

            if vmerge_val == "restart":
                vmerge_anchor[c] = r
            elif vmerge_val is None:
                vmerge_anchor.pop(c, None)

            for cc in range(c + 1, min(c + grid_span, n_cols)):
                grid[r][cc] = "SKIP"
            c += grid_span

    data = [[Spacer(1, 1) for _ in range(n_cols)] for _ in range(n_rows)]
    style_cmds = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    for r in range(n_rows):
        for c in range(n_cols):
            cell = grid[r][c]
            if cell is None or cell == "SKIP":
                continue
            data[r][c] = cell["flowables"] if cell["flowables"] else Spacer(1, 1)
            r2 = r + cell["rowspan"] - 1
            c2 = c + cell["colspan"] - 1
            if cell["rowspan"] > 1 or cell["colspan"] > 1:
                style_cmds.append(("SPAN", (c, r), (c2, r2)))
            if cell["shading"]:
                try:
                    style_cmds.append(("BACKGROUND", (c, r), (c2, r2), colors.HexColor(cell["shading"])))
                except Exception:
                    pass

    table = Table(data, colWidths=col_widths_pt, repeatRows=0)
    table.setStyle(TableStyle(style_cmds))
    return table


def build_body_flowables(document, available_width):
    flowables = []
    for child in document.element.body:
        if child.tag == qn("w:p"):
            paragraph = DocxParagraph(child, document)
            flowables.extend(build_paragraph_flowables(paragraph, document.part, available_width))
        elif child.tag == qn("w:tbl"):
            table = DocxTable(child, document)
            flowables.append(build_table_flowable(table, document.part, available_width))
            flowables.append(Spacer(1, 8))
    return flowables


def collect_header_footer_images(hf_part_container, max_width):
    images = []
    for paragraph in hf_part_container.paragraphs:
        images.extend(extract_inline_images(paragraph, hf_part_container.part, max_width))
    for table in hf_part_container.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    images.extend(extract_inline_images(paragraph, hf_part_container.part, max_width))
    return images


def make_hf_drawer(document, page_w_pt, page_h_pt, margin_l, margin_r, margin_t, margin_b):
    section = document.sections[0]
    usable_w = page_w_pt - margin_l - margin_r
    header_images = collect_header_footer_images(section.header, usable_w)
    footer_images = collect_header_footer_images(section.footer, usable_w)

    def draw(canvas_obj, doc_template):
        canvas_obj.saveState()
        y = page_h_pt - margin_t + 4
        for img in header_images:
            iw, ih = img.drawWidth, img.drawHeight
            x = (page_w_pt - iw) / 2.0
            y -= ih
            try:
                img.drawOn(canvas_obj, x, y)
            except Exception:
                pass
            y -= 4

        y = margin_b - 4
        for img in reversed(footer_images):
            iw, ih = img.drawWidth, img.drawHeight
            x = (page_w_pt - iw) / 2.0
            try:
                img.drawOn(canvas_obj, x, max(y - ih, 2))
            except Exception:
                pass
            y -= ih + 2
        canvas_obj.restoreState()

    return draw


def convert_docx_to_pdf(docx_path, pdf_path):
    """Pure-Python .docx -> .pdf conversion. No LibreOffice / Word needed."""
    document = docx.Document(str(docx_path))
    section = document.sections[0]

    page_w = emu_to_pt(section.page_width)
    page_h = emu_to_pt(section.page_height)
    margin_l = emu_to_pt(section.left_margin)
    margin_r = emu_to_pt(section.right_margin)
    margin_t = emu_to_pt(section.top_margin)
    margin_b = emu_to_pt(section.bottom_margin)

    header_reserve = 70.0
    footer_reserve = 60.0

    frame = Frame(
        margin_l, margin_b + footer_reserve,
        page_w - margin_l - margin_r,
        page_h - margin_t - margin_b - header_reserve - footer_reserve,
        id="body", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )

    draw_hf = make_hf_drawer(document, page_w, page_h, margin_l, margin_r, margin_t, margin_b)

    doc_template = BaseDocTemplate(
        str(pdf_path), pagesize=(page_w, page_h),
        leftMargin=margin_l, rightMargin=margin_r,
        topMargin=margin_t, bottomMargin=margin_b,
    )
    doc_template.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=draw_hf)])

    available_width = page_w - margin_l - margin_r
    story = build_body_flowables(document, available_width) or [Spacer(1, 1)]
    doc_template.build(story)


# ============================================================================
# SECTION 4: TDS lookup (recursive, exact filename match)
# ============================================================================

def build_tds_index(database_dir):
    index = {}
    duplicates = []
    for pdf_path in Path(database_dir).rglob("*"):
        if pdf_path.is_file() and pdf_path.suffix.lower() == ".pdf":
            stem = pdf_path.stem.strip()
            if stem in index and index[stem] != pdf_path:
                duplicates.append((stem, index[stem], pdf_path))
            else:
                index[stem] = pdf_path
    return index, duplicates


def find_tds(model_value, tds_index):
    model_value = (model_value or "").strip()
    if not model_value:
        return None
    return tds_index.get(model_value)


# ============================================================================
# SECTION 5: PDF merging
# ============================================================================

def merge_pdfs(pdf_paths, out_path):
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(str(p))
        for page in reader.pages:
            writer.add_page(page)
    with open(out_path, "wb") as f:
        writer.write(f)


# ============================================================================
# SECTION 5b: Keyword extraction + TDS highlighting (PyMuPDF)
# ============================================================================

STOPWORDS = {
    "a", "an", "the", "and", "or", "of", "in", "on", "at", "to", "for", "with", "by", "as",
    "is", "are", "be", "shall", "should", "will", "this", "that", "these", "those", "from",
    "per", "including", "include", "includes", "minimum", "maximum", "latest", "version",
    "no", "not", "have", "has", "had", "etc", "any", "all", "each", "every", "such", "if",
    "than", "then", "into", "over", "under", "up", "down", "out", "off", "above", "below",
    "between", "more", "most", "other", "some", "only", "own", "same", "so", "too", "very",
    "can", "just", "also", "must", "may", "might", "being", "been", "it", "its", "their",
    "which", "who", "whom", "supply", "installation", "testing", "commissioning", "system",
}

KEYWORD_TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9\-\./]*|\d[A-Za-z0-9\-\./]*")


def extract_keywords(*texts, max_keywords=40):
    """Pull meaningful keywords out of one or more free-text fields (e.g. Description),
    filtering common stopwords and very short/noisy tokens. Order-preserving, deduped."""
    keywords = []
    seen_lower = set()
    for text in texts:
        if not text:
            continue
        for tok in KEYWORD_TOKEN_RE.findall(str(text)):
            tok_clean = tok.strip(" .-/")
            if not tok_clean:
                continue
            low = tok_clean.lower()
            if low in STOPWORDS or low in seen_lower:
                continue
            has_digit = any(ch.isdigit() for ch in tok_clean)
            min_len = 2 if has_digit else 3
            if len(tok_clean) < min_len:
                continue
            seen_lower.add(low)
            keywords.append(tok_clean)
            if len(keywords) >= max_keywords:
                return keywords
    return keywords


def build_keyword_list(description, model_value, max_keywords=40):
    """Full model number goes first (as an exact phrase), then keywords pulled from
    the description and the model string itself."""
    keywords = []
    model_value = (model_value or "").strip()
    if model_value:
        keywords.append(model_value)
    keywords.extend(extract_keywords(description, model_value, max_keywords=max_keywords))
    return keywords


def highlight_keywords_in_pdf(src_pdf_path, keywords, out_pdf_path):
    """Search a TDS PDF for each keyword (case-insensitive) and add a highlight
    annotation over every match. Single-word keywords are matched as WHOLE WORDS
    (so "hard" won't light up inside "hardware"); multi-word keywords (e.g. a
    full model number) are matched as an exact phrase. Writes the result to
    out_pdf_path; the original file in the database folder is never modified.
    Returns (total_highlight_count, list_of_keywords_that_had_at_least_one_match).
    """
    doc = pymupdf.open(str(src_pdf_path))
    total_hits = 0
    matched_keywords = []
    punctuation = ".,;:!?()[]{}\"'`" + "\u2018\u2019\u201c\u201d"
    try:
        for kw in keywords:
            kw = kw.strip()
            if not kw:
                continue
            kw_hits = 0
            is_phrase = " " in kw
            kw_low = kw.lower()
            for page in doc:
                if is_phrase:
                    try:
                        rects = page.search_for(kw)
                    except Exception:
                        rects = []
                    for rect in rects:
                        try:
                            annot = page.add_highlight_annot(rect)
                            annot.update()
                            kw_hits += 1
                        except Exception:
                            pass
                else:
                    for w in page.get_text("words"):
                        word_text = w[4].strip(punctuation)
                        if word_text.lower() == kw_low:
                            rect = pymupdf.Rect(w[0], w[1], w[2], w[3])
                            try:
                                annot = page.add_highlight_annot(rect)
                                annot.update()
                                kw_hits += 1
                            except Exception:
                                pass
            if kw_hits:
                matched_keywords.append(kw)
                total_hits += kw_hits
        doc.save(str(out_pdf_path), garbage=4, deflate=True)
    finally:
        doc.close()
    return total_hits, matched_keywords


# ============================================================================
# SECTION 6: Main pipeline
# ============================================================================

class LocalTDSProvider:
    """TDS source backed by a local folder (searched recursively, exact filename match)."""

    def __init__(self, database_dir):
        self.index, self.duplicates = build_tds_index(database_dir)

    @property
    def count(self):
        return len(self.index)

    def find(self, model_value):
        return find_tds(model_value, self.index)


def run(boq_path, cover_path, database_dir, output_dir,
        tds_provider=None, progress_cb=None, highlight=True, plan=None):
    """Run the pipeline.
    plan:         optional list from TDSLibrary.plan() (possibly edited by the user): decides the TDS,
                  model and make for each line, instead of tds_provider.
    tds_provider: any object with .find(model_value) -> local Path or None and .count.
                  Defaults to a LocalTDSProvider over database_dir.
    progress_cb:  optional callable(done, total, message) for UIs.
    highlight:    set False to skip keyword highlighting (faster)."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    def report(done, total, msg):
        print(f"   [{done}/{total}] {msg}")
        if progress_cb:
            progress_cb(done, total, msg)

    print(f"\nReading BOQ: {boq_path}")
    headers, items = read_boq(boq_path)
    print(f"   Found {len(items)} line item(s).")

    if tds_provider is None and plan is not None:
        class _PlanOnly:                      # the plan already holds every decision
            count = sum(1 for e in plan if e.get("path"))
        tds_provider = _PlanOnly()
    if tds_provider is None:
        print(f"Indexing database: {database_dir}")
        tds_provider = LocalTDSProvider(database_dir)
        if tds_provider.duplicates:
            print(f"   WARNING: {len(tds_provider.duplicates)} duplicate filename(s) found "
                  "(first occurrence used for each).")
    print(f"   TDS source ready ({tds_provider.count} PDF(s) available).")

    work_dir = output_dir / "_work"
    work_dir.mkdir(exist_ok=True)

    missing_log = []
    highlight_log = []
    match_log = []
    plan_by_sno = {e["sno"]: e for e in (plan or [])}
    total = len(items)
    for i, (sno, row) in enumerate(items, start=1):
        report(i - 1, total, f"S.No {sno}: building cover page")
        item_dir = output_dir / sno
        item_dir.mkdir(parents=True, exist_ok=True)

        description = row.get("Description", "") or ""
        model_value = row.get("Model no.") or row.get("Model") or ""
        make_value = row.get("Make", "") or ""
        match_info = None
        entry = plan_by_sno.get(sno)
        if entry is not None:                           # decided by the AI plan (after user review)
            tds_path = entry["path"] if entry.get("use") else None
            model_value = (entry.get("model") or model_value) if entry.get("path") else model_value
            make_value = (entry.get("make") or make_value) if entry.get("path") else make_value
            match_info = {"Matched TDS": entry.get("match", ""), "Confidence": entry.get("confidence", ""),
                          "Score": "", "Used": "Yes" if tds_path else "No", "Why": entry.get("reason", ""),
                          "Decided by": entry.get("by", ""),
                          "Other Candidates": " | ".join(c[0] for c in entry.get("candidates", [])[1:4]),
                          "TDS has text layer": ""}
        elif hasattr(tds_provider, "match_row"):        # automatic ranking (no review step)
            tds_path, match_info = tds_provider.match_row(row)
        else:                                           # exact filename mode
            tds_path = tds_provider.find(model_value)
        if match_info is not None:
            match_log.append({"S.No": sno, "Make": make_value, "Model no.": model_value,
                              "Description": description.replace("\n", " ")[:100], **match_info})

        # cover page: {placeholders} if present + label-based fill (no placeholders needed)
        mapping = {**row, "Model no.": model_value, "Make": make_value, "S.No": sno}
        auto_values = {"sno": sno, "description": description, "make": make_value, "model": model_value}
        canon = {"s.no", "description", "make", "model no.", "model"}
        import cover_autofill
        extra = {cover_autofill.norm_label(h): v for h, v in row.items()
                 if h and h.strip().lower() not in canon and v not in ("", None)}
        cover_docx = item_dir / f"Cover_Page_{sno}.docx"
        fill_cover_page(cover_path, mapping, cover_docx, auto_values, extra)

        cover_pdf = work_dir / f"Cover_Page_{sno}.pdf"
        try:
            convert_docx_to_pdf(cover_docx, cover_pdf)
        except Exception as e:
            print(f"   ! S.No {sno}: cover page rendering failed ({e}), skipping.")
            continue

        final_pdf = item_dir / f"TDS_{sno}.pdf"

        if tds_path:
            if highlight:
                report(i - 1, total, f"S.No {sno}: highlighting keywords")
                keywords = build_keyword_list(description, model_value)
                highlighted_tds = work_dir / f"Highlighted_TDS_{sno}.pdf"
                try:
                    total_hits, matched_keywords = highlight_keywords_in_pdf(
                        tds_path, keywords, highlighted_tds
                    )
                except Exception as e:
                    print(f"   ! S.No {sno}: keyword highlighting failed ({e}), using TDS as-is.")
                    total_hits, matched_keywords = 0, []
                    highlighted_tds = tds_path
                highlight_log.append({
                    "S.No": sno, "Model no.": model_value,
                    "Keywords Searched": len(keywords),
                    "Keywords Matched": len(matched_keywords),
                    "Total Highlights": total_hits,
                    "Matched Keywords": "; ".join(matched_keywords),
                })
            else:
                highlighted_tds = tds_path
            merge_pdfs([cover_pdf, highlighted_tds], final_pdf)
        else:
            shutil.copyfile(cover_pdf, final_pdf)
            if match_info and match_info.get("Matched TDS"):
                reason = (f"Best candidate '{match_info['Matched TDS']}' has only "
                          f"{match_info['Confidence']} confidence - review match_report.csv")
            elif not model_value.strip():
                reason = "No model number in BOQ"
            else:
                reason = "No matching TDS filename found"
            missing_log.append({
                "S.No": sno, "Description": description[:120],
                "Make": row.get("Make", ""), "Model no.": model_value, "Reason": reason,
            })
        report(i, total, f"S.No {sno}: done")

    log_path = output_dir / "missing_tds_log.csv"
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["S.No", "Description", "Make", "Model no.", "Reason"])
        writer.writeheader()
        writer.writerows(missing_log)

    if highlight:
        highlight_log_path = output_dir / "highlight_summary.csv"
        with open(highlight_log_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=[
                "S.No", "Model no.", "Keywords Searched", "Keywords Matched",
                "Total Highlights", "Matched Keywords",
            ])
            writer.writeheader()
            writer.writerows(highlight_log)

    if match_log:
        with open(output_dir / "match_report.csv", "w", newline="", encoding="utf-8") as f:
            fields = ["S.No", "Make", "Model no.", "Description", "Matched TDS", "Confidence", "Score",
                      "Used", "Decided by", "Why", "Other Candidates", "TDS has text layer"]
            writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(match_log)

    shutil.rmtree(work_dir, ignore_errors=True)

    print(f"\nDone. {len(items)} line item(s) processed; {len(missing_log)} missing a TDS.")
    print(f"Output saved in: {output_dir}")
    return {"items": len(items), "missing": len(missing_log), "output_dir": str(output_dir)}


def main():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        sys.exit(
            "tkinter is not available in this Python installation.\n"
            "Either install tkinter (on Linux: sudo apt install python3-tk),\n"
            "or call run(boq_path, cover_path, database_dir, output_dir) directly "
            "from your own script instead of using the file-picker UI."
        )

    print("========================================")
    print("     TDS Automation Tool (File UI)      ")
    print("     (pure Python - no LibreOffice)      ")
    print("========================================\n")
    print("Opening file/folder selection windows...\n")

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    boq = filedialog.askopenfilename(
        title="Select BOQ Excel File (.xlsx)",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not boq:
        print("Operation cancelled: No BOQ file selected.")
        return

    cover = filedialog.askopenfilename(
        title="Select Cover Page Template File (.docx)",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )
    if not cover:
        print("Operation cancelled: No cover page template selected.")
        return

    database = filedialog.askdirectory(title="Select TDS Database Directory")
    if not database:
        print("Operation cancelled: No database directory selected.")
        return

    output = filedialog.askdirectory(title="Select Output Directory")
    if not output:
        print("Operation cancelled: No output directory selected.")
        return

    run(boq, cover, database, output)


if __name__ == "__main__":
    main()
