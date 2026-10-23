# -*- coding: utf-8 -*-
"""
Extract labelled figures FROM the Kolb & Whishaw textbook
(Fundamentals of Human Neuropsychology) and save them as cropped PNGs under
assets/unit8_figures/from_books/ for the Unit VIII notes generator to embed.

The Kolb PDF has a real text layer and vector/embedded figures, so each figure
is produced by rendering its page at 3x and clipping a rectangle. Crop
rectangles are expressed as fractions of the page width/height
(fx0, fy0, fx1, fy1) and were tuned by rendering each page at 2x, viewing it,
and reading the text layer back; they frame the full labelled figure while
excluding surrounding body text and the page running header/footer.

Each page index below was verified against the page text layer AND by eye:
pages that only CROSS-REFERENCE a figure (e.g. "see Figure 3.3" on idx 236,
"see Figure 11.7" on idx 264) do NOT actually contain that plate and are NOT
cropped here.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 extract_unit8_figures.py
"""

import os
import warnings

warnings.filterwarnings("ignore")  # silence the fitz deprecation FutureWarning
import fitz  # noqa: E402  (pymupdf)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "assets", "unit8_figures", "from_books")

KOLB = os.path.join(HERE, "Fundametals Of Human NeuroPsychology - Bryan Kolb-compressed.pdf")

ZOOM = 3.0  # render scale; a 3x render of these crops is ~600-1400px wide

# (pdf, page_index, (fx0, fy0, fx1, fy1), out_filename)
# Page indices are 0-based as fitz indexes them; the true book Figure number
# for each crop is recorded in the comment and in PROVENANCE.txt.
JOBS = [
    # idx 14: Figure 1.6 Wernicke's 1874 language model (A chimp brain, B human
    # brain with Broca's/Wernicke's areas + arcuate fasciculus). Left column.
    (KOLB, 14, (0.055, 0.075, 0.495, 0.585), "kolb_fig1_6_wernicke_model.png"),
    # idx 64: Figure 3.13 Gyri and sulci (A/B lateral+medial gyri, C/D sulci).
    # Full-width figure band across the top of the page.
    (KOLB, 64, (0.035, 0.065, 0.99, 0.515), "kolb_fig3_13_gyri_sulci.png"),
    # idx 66: Figure 3.16 Brodmann's areas of the cortex (A lateral, B medial)
    # plus the function/map-code/Brodmann-area table. Lower portion of page.
    (KOLB, 66, (0.30, 0.605, 0.975, 0.95), "kolb_fig3_16_brodmann_areas.png"),
    # idx 229: Figure 10.7 Brodmann's map (clean standalone lateral + medial
    # views). Right column, lower half.
    (KOLB, 229, (0.69, 0.585, 0.985, 0.925), "kolb_fig10_7_brodmann_map.png"),
    # idx 258 (top): Figure 11.5 Visual pathways are crossed - each visual
    # field represented in the opposite hemisphere (optic chiasm crossing).
    (KOLB, 258, (0.29, 0.085, 0.95, 0.445), "kolb_fig11_5_visual_pathways.png"),
    # idx 258 (bottom): Figure 11.6 Connection between the two sides of the
    # world via the corpus callosum ("key" split-brain demonstration).
    (KOLB, 258, (0.29, 0.585, 0.99, 0.945), "kolb_fig11_6_corpus_callosum.png"),
    # idx 268: Figure 11.13 Motor asymmetry - successive video frames showing
    # the right side of the mouth opening wider during speech. Top-left column.
    (KOLB, 268, (0.06, 0.075, 0.305, 0.435), "kolb_fig11_13_motor_asymmetry.png"),
]


def extract(pdf_path, page_index, fracs, out_name):
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    r = page.rect
    fx0, fy0, fx1, fy1 = fracs
    clip = fitz.Rect(
        r.x0 + fx0 * r.width,
        r.y0 + fy0 * r.height,
        r.x0 + fx1 * r.width,
        r.y0 + fy1 * r.height,
    )
    pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), clip=clip)
    out_path = os.path.join(OUT_DIR, out_name)
    pix.save(out_path)
    doc.close()
    print("wrote %s  (%dx%d)" % (out_name, pix.width, pix.height))


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for pdf_path, page_index, fracs, out_name in JOBS:
        extract(pdf_path, page_index, fracs, out_name)
    print("done:", len(JOBS), "figures ->", OUT_DIR)


if __name__ == "__main__":
    main()
