# -*- coding: utf-8 -*-
"""Generate the Word (.docx) deliverable for Psychotherapy Unit IV notes.

This consumes the SAME single-source content module as the HTML build
(notes_src/content_unit4.py in the repo, kept byte-identical to the build
copy .work/build/content_unit4.py) so the two deliverables cannot drift.

Inline markup tokens are converted to readable Word text (never left raw):
  {{LAW}}/{{CODE}}/{{GUIDE}}/{{THEORY}}/{{EG}}/{{SEC}}  -> bold label prefix, e.g. "LAW: "
  {{c:TEXT}}                                            -> citation "[TEXT]"
  [[term|gloss]]                                        -> "term (gloss)"
  {{fig:NAME}}                                          -> dropped (HTML-only inline SVG);
                                                           {t:"fig"} blocks render a caption line.

Usage:
  PYTHONPATH=/projects/sandbox/.pylibs python3 build_docx_unit4.py <out.docx>
"""
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# --- import the single-source content module ------------------------------
# Prefer the tracked repo copy (notes_src/content_unit4.py); fall back to the
# build copy. Both are kept byte-identical.
_REPO_SRC = "/projects/sandbox/Dango-kiro/notes_src"
_BUILD_SRC = "/projects/sandbox/.work/build"
if os.path.exists(os.path.join(_REPO_SRC, "content_unit4.py")):
    sys.path.insert(0, _REPO_SRC)
    import content_unit4 as C  # noqa: E402
else:
    sys.path.insert(0, _BUILD_SRC)
    import content_unit4 as C  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else \
    "/projects/sandbox/Dango-kiro/Psychotherapy_Unit_IV_Comprehensive_Notes.docx"

# Short citation key -> full "Author, Full Title" form (verified allowlist).
# This MUST stay identical to the CITES dict in .work/build/build_unit4.py so
# the HTML and DOCX deliverables carry byte-identical citation text. Body keys
# and the descriptive-title reader key are already full names and pass through
# unchanged (after any 'PDF p.' suffix is stripped).
CITES = {
    # Unit IV book keys (the affective psychotherapies).
    'Corey': 'Gerald Corey, Theory and Practice of Counseling and Psychotherapy',
    'Sharf': 'Richard S. Sharf, Theories of Psychotherapy and Counseling',
    'Corsini': 'Raymond J. Corsini & Danny Wedding, Current Psychotherapies',
    'Gabbard': 'Glen O. Gabbard, Long-Term Psychodynamic Psychotherapy: A Basic Text',
    'McWilliams': 'Nancy McWilliams, Psychoanalytic Diagnosis',
    'Freud': 'Sigmund Freud, The Interpretation of Dreams',
    'Rogers': 'Carl R. Rogers, Client-Centered Therapy',
    'Yalom': 'Irvin D. Yalom, Existential Psychotherapy',
    'Frankl': 'Viktor E. Frankl, Man\u2019s Search for Meaning',
    'Perls': 'Frederick S. Perls, Ralph Hefferline & Paul Goodman, Gestalt Therapy',
    'Adler': 'Alfred Adler, The Science of Living',
    'Berne': 'Eric Berne, Transactional Analysis in Psychotherapy',
    'Glasser': 'William Glasser, Reality Therapy: A New Approach to Psychiatry',
    'Erickson': 'Milton H. Erickson & Ernest L. Rossi, Hypnotherapy: An Exploratory Casebook',
    'Axline': 'Virginia M. Axline, Play Therapy',
    'Moreno': 'Jacob L. Moreno, Psychodrama',
    # Primary compiled reader for this unit: descriptive title, no invented author.
    'Affective Therapies Reader': 'Affective Therapies (Unit IV compiled reader)',
    # Body / primary-source keys (full body + document forms).
    'APA Dictionary of Psychology': 'American Psychological Association, APA Dictionary of Psychology',
    'APA Division 30': 'American Psychological Association (Division 30, Society of Psychological Hypnosis)',
    'WHO mhGAP': 'World Health Organization, mhGAP Intervention Guide',
    'IACP': 'Indian Association of Clinical Psychologists (secondary summary only)',
}


def _cite(payload):
    """Expand a {{c:...}} payload into the display text for the citation.

    The payload is a short key optionally followed by ', PDF p.NN...' and/or a
    ', Std/s./section' suffix. Book keys map to the full Author, Title form;
    body/reader names pass through. Any 'PDF p.' page token is dropped, while
    a genuine Std/section/s. suffix is preserved. This mirrors _cite() in
    .work/build/build_unit4.py so HTML and DOCX citations match."""
    payload = payload.strip()
    parts = [p.strip() for p in payload.split(',')]
    key = parts[0]
    suffixes = [p for p in parts[1:] if not re.match(r'(?i)^PDF\s*p', p)]
    base = CITES.get(key, key)
    return ', '.join([base] + suffixes) if suffixes else base


# Evidence-label tags -> readable prefix text.
TAG_LABEL = {
    "LAW": "Binding law",
    "CODE": "Professional code",
    "GUIDE": "Guideline",
    "THEORY": "Theoretical view",
    "EG": "Illustrative example",
    "SEC": "Secondary source",
}

_TOKEN_RE = re.compile(
    r"\{\{(LAW|CODE|GUIDE|THEORY|EG|SEC)\}\}"       # 1: evidence tag
    r"|\{\{c:([^}]+)\}\}"                            # 2: citation
    r"|\{\{fig:\w+\}\}"                              # (figure ref, dropped)
    r"|\[\[([^\]|]+)\|([^\]]+)\]\]"                  # 3,4: glossary term|gloss
    r"|\*\*([^*]+)\*\*"                              # 5: bold key-word span
)


def _add_runs(paragraph, text):
    """Add text to a paragraph, converting inline tokens into styled runs.

    Returns the paragraph for chaining.
    """
    pos = 0
    for m in _TOKEN_RE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        tag, cite, gterm, ggloss, bold = (
            m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
        )
        if tag:
            r = paragraph.add_run(TAG_LABEL[tag] + ": ")
            r.bold = True
        elif cite:
            r = paragraph.add_run("[" + _cite(cite) + "]")
            r.italic = True
            r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        elif gterm is not None:
            paragraph.add_run(gterm + " ")
            r = paragraph.add_run("(" + ggloss + ")")
            r.italic = True
        elif bold is not None:
            # A **bold** span may itself contain a nested [[term|gloss]] (the
            # HTML build expands the gloss before the bold pass, so the gloss
            # ends up inside <strong>). Mirror that here: expand any nested
            # gloss into "term (gloss)" while keeping every run bold.
            g = re.match(r"^(.*?)\[\[([^\]|]+)\|([^\]]+)\]\](.*)$", bold, re.S)
            if g:
                pre, gt, gg, post = g.group(1), g.group(2), g.group(3), g.group(4)
                if pre:
                    paragraph.add_run(pre).bold = True
                paragraph.add_run(gt + " ").bold = True
                rg = paragraph.add_run("(" + gg + ")")
                rg.bold = True
                rg.italic = True
                if post:
                    paragraph.add_run(post).bold = True
            else:
                r = paragraph.add_run(bold)
                r.bold = True
        # figure refs ({{fig:...}}) are intentionally dropped (HTML-only).
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])
    return paragraph


def _clean_leading_space(paragraph):
    """If the first run starts with a space (e.g. after a dropped fig ref),
    trim it so bullets look tidy."""
    for run in paragraph.runs:
        if run.text:
            run.text = run.text.lstrip()
            break


def _bullet(doc, item, level=0):
    """Render one list item as a Word list paragraph. item may be a plain
    string or a (lead, [sub, ...]) tuple carrying one nested level."""
    style = "List Bullet" if level == 0 else "List Bullet 2"
    if isinstance(item, (list, tuple)):
        lead, subs = item
        p = doc.add_paragraph(style=style)
        _add_runs(p, lead)
        _clean_leading_space(p)
        for s in subs:
            _bullet(doc, s, level + 1)
    else:
        p = doc.add_paragraph(style=style)
        _add_runs(p, item)
        _clean_leading_space(p)


def _numbered(doc, item, level=0):
    style = "List Number" if level == 0 else "List Number 2"
    if isinstance(item, (list, tuple)):
        lead, subs = item
        p = doc.add_paragraph(style=style)
        _add_runs(p, lead)
        _clean_leading_space(p)
        for s in subs:
            _numbered(doc, s, level + 1)
    else:
        p = doc.add_paragraph(style=style)
        _add_runs(p, item)
        _clean_leading_space(p)


def _shade_cell(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {
        qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): hex_fill,
    })
    tcPr.append(shd)


def _table(doc, b):
    head = b["head"]
    rows = b["rows"]
    if b.get("caption"):
        cap = doc.add_paragraph(style="Caption")
        _add_runs(cap, b["caption"])
    table = doc.add_table(rows=1, cols=len(head))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, h in enumerate(head):
        hdr[i].paragraphs[0].text = ""
        r = hdr[i].paragraphs[0].add_run(h)
        r.bold = True
        _shade_cell(hdr[i], "DDE6F0")
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            # a cell value may itself contain inline tokens
            para = cells[i].paragraphs[0]
            para.text = ""
            _add_runs(para, val)
    doc.add_paragraph()  # spacer after table


_BOX_HEADING = {
    "def": "Definition", "key": "Key points", "law": "Law",
    "ex": "Example", "warn": "Caution", "info": "Note",
}
_BOX_FILL = {
    "def": "EAF2FB", "key": "EAF7EE", "law": "FDECEC",
    "ex": "FFF6E5", "warn": "FFF0F0", "info": "F1F3F5",
}


def _box(doc, b):
    """Render a box/callout as a single-cell shaded table so it reads as a
    distinct block in Word (title in bold, then its text and/or bullets)."""
    kind = b.get("kind", "info")
    title = b.get("title") or _BOX_HEADING.get(kind, "Note")
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    _shade_cell(cell, _BOX_FILL.get(kind, "F1F3F5"))
    # title line
    tp = cell.paragraphs[0]
    tp.text = ""
    tr = tp.add_run(title)
    tr.bold = True
    if b.get("text"):
        bp = cell.add_paragraph()
        _add_runs(bp, b["text"])
    for it in b.get("items", []):
        _box_bullet(cell, it, 0)
    doc.add_paragraph()  # spacer


def _box_bullet(cell, item, level):
    """Bullets inside a box cell. Word list styles don't nest well inside a
    table cell, so render with a leading glyph and indentation."""
    prefix = ("    " * level) + ("\u2022 " if level == 0 else "\u25E6 ")
    if isinstance(item, (list, tuple)):
        lead, subs = item
        p = cell.add_paragraph()
        p.add_run(prefix)
        _add_runs(p, lead)
        for s in subs:
            _box_bullet(cell, s, level + 1)
    else:
        p = cell.add_paragraph()
        p.add_run(prefix)
        _add_runs(p, item)


def _render_block(doc, b):
    t = b["t"]
    if t == "lead":
        p = doc.add_paragraph()
        r = p.add_run("")
        _add_runs(p, b["text"])
        for run in p.runs:
            run.italic = True
    elif t == "h3":
        doc.add_heading(b["text"], level=2)
    elif t == "h4":
        doc.add_heading(b["text"], level=3)
    elif t == "ul":
        for it in b["items"]:
            _bullet(doc, it, 0)
    elif t == "ol":
        for it in b["items"]:
            _numbered(doc, it, 0)
    elif t == "table":
        _table(doc, b)
    elif t == "box":
        _box(doc, b)
    elif t == "fig":
        # Figures are HTML-only inline SVG; render a short caption line only
        # (do not fabricate an image).
        p = doc.add_paragraph()
        r = p.add_run("Figure (diagram in HTML edition): ")
        r.bold = True
        r.italic = True
        _add_runs(p, b.get("caption", ""))
        for run in p.runs:
            run.italic = True
    elif t == "mn":
        p = doc.add_paragraph()
        r = p.add_run("Mnemonic \u2014 " + b["word"] + ": ")
        r.bold = True
        _add_runs(p, b["text"])
    elif t == "note":
        p = doc.add_paragraph()
        _add_runs(p, b["text"])
        for run in p.runs:
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)


def _blocks(doc, blocks):
    for b in blocks:
        _render_block(doc, b)


def _set_footer_credit(doc, credit):
    """Put the exact credit string in the section footer (centered) so it
    repeats on every page."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(credit)
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def build():
    doc = Document()

    # Base font tweaks for readability.
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)

    # ---- Cover / title page ----
    doc.add_paragraph()
    doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title_p.add_run(C.TITLE)
    tr.bold = True
    tr.font.size = Pt(26)
    tr.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub_p.add_run(C.SUBTITLE)
    sr.font.size = Pt(13)
    sr.italic = True
    sr.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph()
    credit_p = doc.add_paragraph()
    credit_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = credit_p.add_run(C.CREDIT)
    cr.bold = True
    cr.font.size = Pt(13)

    # footer credit (repeats on every page)
    _set_footer_credit(doc, C.CREDIT)

    doc.add_page_break()

    # ---- Front matter: how to read (evidence legend + citation convention)
    doc.add_heading("How to read these notes", level=1)
    _blocks(doc, C.INTRO)

    # ---- Front matter: a simple introduction to the affective psychotherapies
    doc.add_heading("Introduction to Affective Psychotherapies", level=1)
    _blocks(doc, C.AFFECTIVE_INTRO)

    # ---- Topic sections (the Unit IV affective-therapy topics) ----
    for sec in C.SECTIONS:
        doc.add_heading(f'{sec["num"]}. {sec["title"]}', level=1)
        _blocks(doc, sec["blocks"])

    # ---- Source note (honesty / limits disclosure) ----
    doc.add_heading("Source note and honest limits", level=1)
    _blocks(doc, C.SOURCE_NOTE)

    # ---- References ----
    doc.add_heading("Key sources & references", level=1)
    _blocks(doc, C.REFERENCES)

    doc.save(OUT)
    print("written", OUT)


if __name__ == "__main__":
    build()
