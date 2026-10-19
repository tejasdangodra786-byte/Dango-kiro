# -*- coding: utf-8 -*-
"""
Generator for the Word (.docx) study-notes deliverable:
"Biological Foundations of Behavior - Unit II: Structure and Functions of Cells".

The content is NOT re-authored here. It is imported from generate_unit2_notes.py,
which holds the single source of truth (CONTENT plus TITLE, CREDIT, SUBTITLE,
HOW_TO_READ and REFERENCES), so the Word file mirrors the HTML exactly.

Inline markup inside point / intro / flow / mnemonic text is parsed here into
python-docx runs:
  **word**                 -> a bold run (memory-hook keyword only)
  [[plain-language text]]  -> a visually DISTINCT run (italic, coloured, shaded)
                              so the plain-language explanation stands out.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 generate_unit2_docx.py
"""

import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor

from generate_unit2_notes import (
    CONTENT,
    CREDIT,
    HOW_TO_READ,
    REFERENCES,
    SUBTITLE,
    TITLE,
)

OUTPUT_DOCX = "Biological_Foundations_Unit2_Structure_and_Functions_of_Cells.docx"

# Colour palette (mirrors the HTML house style)
COLOR_TITLE = RGBColor(0x0F, 0x27, 0x44)      # deep navy
COLOR_ACCENT = RGBColor(0x1D, 0x5F, 0xA6)     # blue
COLOR_INTRO = RGBColor(0x40, 0x50, 0x6A)      # muted slate
COLOR_PLAIN = RGBColor(0x6B, 0x4E, 0x00)      # brown text for plain-language runs
COLOR_MNEMO = RGBColor(0x0E, 0x7C, 0x6B)      # teal
COLOR_LABEL = RGBColor(0x5A, 0x6A, 0x7E)      # grey labels

PLAIN_SHADE = "FFF3D6"   # light-amber shading behind plain-language runs
DEF_SHADE = "F0EAFA"     # light-purple shading behind sourced definitions

# Tokeniser for the two inline markers (same markers as the HTML renderer).
_TOKEN_RE = re.compile(r"\*\*(.+?)\*\*|\[\[(.+?)\]\]")


def _shade_run(run, fill_hex):
    """Apply a solid background shade to a single run (w:shd on the run props)."""
    rpr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    rpr.append(shd)


def add_markup_runs(paragraph, text, base_size=None):
    """Parse **bold** and [[plain]] markers in `text` into runs on `paragraph`.

    Bold markers become bold runs; plain-language markers become visually
    distinct runs (italic + brown text + light-amber shading), wrapped in
    parentheses so they read naturally in the flow of the sentence.
    """
    i = 0
    for m in _TOKEN_RE.finditer(text):
        if m.start() > i:
            run = paragraph.add_run(text[i:m.start()])
            if base_size:
                run.font.size = base_size
        if m.group(1) is not None:
            run = paragraph.add_run(m.group(1))
            run.bold = True
            run.font.color.rgb = COLOR_TITLE
            if base_size:
                run.font.size = base_size
        else:
            run = paragraph.add_run("(" + m.group(2) + ")")
            run.italic = True
            run.font.color.rgb = COLOR_PLAIN
            if base_size:
                run.font.size = base_size
            _shade_run(run, PLAIN_SHADE)
        i = m.end()
    if i < len(text):
        run = paragraph.add_run(text[i:])
        if base_size:
            run.font.size = base_size


def set_cell_background(cell, fill_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def add_footer_credit(document):
    """Put the exact credit in the section footer so it repeats on every page."""
    section = document.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.text = ""
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(CREDIT)
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_LABEL


def add_label(document, text):
    para = document.add_paragraph()
    run = para.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_LABEL
    para.paragraph_format.space_after = Pt(4)
    return para


def add_cover(document):
    # vertical spacing to push the title toward the middle of the first page
    for _ in range(4):
        document.add_paragraph()

    label = document.add_paragraph()
    label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lrun = label.add_run("M.Phil. Clinical Psychology \u00b7 Study Notes")
    lrun.font.size = Pt(11)
    lrun.font.color.rgb = COLOR_LABEL
    lrun.bold = True

    heading = document.add_heading(level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hrun = heading.add_run(TITLE)
    hrun.font.size = Pt(26)
    hrun.font.color.rgb = COLOR_TITLE

    sub = document.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    srun = sub.add_run(SUBTITLE)
    srun.italic = True
    srun.font.size = Pt(12)
    srun.font.color.rgb = COLOR_INTRO

    document.add_paragraph()

    credit = document.add_paragraph()
    credit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    crun = credit.add_run(CREDIT)
    crun.bold = True
    crun.font.size = Pt(13)
    crun.font.color.rgb = COLOR_ACCENT

    document.add_page_break()


def add_how_to_read(document):
    document.add_heading("How to read these notes", level=1)
    for g in HOW_TO_READ:
        para = document.add_paragraph(style="List Bullet")
        add_markup_runs(para, g)
    document.add_page_break()


def add_flow(document, flow):
    add_label(document, "Flowchart \u00b7 " + flow["title"])
    para = document.add_paragraph()
    steps = flow["steps"]
    n = len(steps)
    for idx, (label, sub, is_end) in enumerate(steps):
        box = para.add_run("[ " + label + " \u2014 " + sub + " ]")
        box.bold = True
        box.font.size = Pt(10)
        box.font.color.rgb = COLOR_TITLE if is_end else COLOR_ACCENT
        if idx < n - 1:
            arrow = para.add_run("  \u2192  ")
            arrow.bold = True
            arrow.font.color.rgb = COLOR_ACCENT


def add_mnemonic(document, mn):
    add_label(document, "Mnemonic")
    word = document.add_paragraph()
    wrun = word.add_run(mn["word"])
    wrun.bold = True
    wrun.font.size = Pt(15)
    wrun.font.color.rgb = COLOR_MNEMO

    line = document.add_paragraph()
    add_markup_runs(line, mn["line"])
    for p in line.runs:
        p.italic = True

    for letter, text in mn["items"]:
        para = document.add_paragraph(style="List Bullet")
        lrun = para.add_run(letter + ": ")
        lrun.bold = True
        lrun.font.color.rgb = COLOR_MNEMO
        add_markup_runs(para, text)


def add_definition(document, d):
    table = document.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    set_cell_background(cell, DEF_SHADE)

    label_p = cell.paragraphs[0]
    lrun = label_p.add_run(d["label"].upper())
    lrun.bold = True
    lrun.font.size = Pt(10)
    lrun.font.color.rgb = RGBColor(0x4A, 0x20, 0x80)

    body_p = cell.add_paragraph()
    body_p.add_run(d["text"])

    cite_p = cell.add_paragraph()
    crun = cite_p.add_run("Source: " + d["cite"])
    crun.italic = True
    crun.font.size = Pt(9)
    crun.font.color.rgb = COLOR_LABEL
    document.add_paragraph()


def add_topic(document, t):
    heading = document.add_heading(level=1)
    hrun = heading.add_run("%d. %s" % (t["number"], t["title"]))
    hrun.font.color.rgb = COLOR_TITLE

    intro = document.add_paragraph()
    add_markup_runs(intro, t["intro"])
    for r in intro.runs:
        r.italic = True

    add_label(document, "Key points (exam-ready, one full sentence each)")
    for idx, p in enumerate(t["points"], start=1):
        para = document.add_paragraph(style="List Number")
        add_markup_runs(para, p)

    if t.get("definition"):
        add_definition(document, t["definition"])

    add_flow(document, t["flow"])
    add_mnemonic(document, t["mnemonic"])
    document.add_page_break()


def add_references(document):
    document.add_heading("References", level=1)

    document.add_heading("Books used", level=2)
    for b in REFERENCES["books"]:
        document.add_paragraph(b, style="List Bullet")

    document.add_heading("Named online sources used", level=2)
    for o in REFERENCES["online"]:
        document.add_paragraph(o, style="List Bullet")

    document.add_heading("Secondary sources flagged (not accessed directly)", level=2)
    for f in REFERENCES["flagged"]:
        document.add_paragraph(f, style="List Bullet")


def build_document():
    document = Document()

    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    add_footer_credit(document)
    add_cover(document)
    add_how_to_read(document)
    for t in CONTENT:
        add_topic(document, t)
    add_references(document)
    return document


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(here, OUTPUT_DOCX)
    document = build_document()
    document.save(out_path)
    print("Wrote", out_path)
    print("Topics:", len(CONTENT))
    for t in CONTENT:
        print("  - %s: %d points" % (t["title"], len(t["points"])))


if __name__ == "__main__":
    main()
