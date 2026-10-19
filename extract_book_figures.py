# -*- coding: utf-8 -*-
"""
One-off helper: extract labelled figures FROM the two uploaded textbooks
(Carlson, Physiology of Behavior Ch.2; Kolb & Whishaw, Fundamentals of Human
Neuropsychology) and save them as cropped PNGs under
assets/unit2_figures/from_books/ for use by the Unit II notes generators.

The Carlson PDF is a SCANNED chapter (every page is one full-page raster image,
no text layer), so each figure is produced by rendering the page at 3x and
clipping a rectangle. The Kolb PDF has a real text layer and vector/embedded
figures, so its pages are rendered at 3x and clipped the same way.

Crop rectangles are expressed as fractions of the page width/height (fx0, fy0,
fx1, fy1) and were tuned by rendering each page and reading it back; they frame
the full labelled figure while excluding surrounding body text and the page
running header/footer.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 extract_book_figures.py
"""

import os
import warnings

warnings.filterwarnings("ignore")  # silence the fitz deprecation FutureWarning
import fitz  # noqa: E402  (pymupdf)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "assets", "unit2_figures", "from_books")

CARLSON = os.path.join(HERE, "Structure and function of cells carlson chapter 2.pdf")
KOLB = os.path.join(HERE, "Fundametals Of Human NeuroPsychology - Bryan Kolb-compressed.pdf")

ZOOM = 3.0  # render scale; a 3x render of these crops is ~800-1300px wide

# (pdf, page_index, (fx0, fy0, fx1, fy1), out_filename)
JOBS = [
    # ---- Carlson (scanned chapter 2) --------------------------------------
    # p3 (book p.29) Figure 2.1 principal parts of a multipolar neuron (top-left)
    (CARLSON, 3, (0.085, 0.065, 0.715, 0.425), "carlson_fig2_1_neuron.png"),
    # p7 (book p.33) Figure 2.6 astrocyte structure/location (bottom-right)
    (CARLSON, 7, (0.515, 0.565, 1.0, 0.985), "carlson_fig2_6_astrocyte.png"),
    # p8 (book p.34) Figure 2.7 oligodendrocyte forming myelin (bottom-left)
    (CARLSON, 8, (0.05, 0.565, 0.495, 0.985), "carlson_fig2_7_oligodendrocyte.png"),
    # p9 (book p.35) Figure 2.9 blood-brain barrier capillaries (top-right)
    (CARLSON, 9, (0.555, 0.07, 1.0, 0.415), "carlson_fig2_9_bbb.png"),
    # ---- Kolb (text layer + embedded figures) -----------------------------
    # p48 (book p.49) Figure 3.2 neuron types A/B/C (top band)
    (KOLB, 48, (0.03, 0.035, 0.925, 0.305), "kolb_fig3_2_neuron_types.png"),
    # p48 Table 3.1 Types of glial cells (lower-left)
    (KOLB, 48, (0.03, 0.615, 0.925, 0.945), "kolb_table3_1_glial_types.png"),
    # p89 (book p.90) Figure 4.14 basis/method of recording a resting potential (right column)
    (KOLB, 89, (0.535, 0.06, 1.0, 0.95), "kolb_fig4_14_resting_potential.png"),
    # p93 (book p.94) Figure 4.17 voltage-sensitive Na+/K+ channels across AP phases (right column)
    (KOLB, 93, (0.415, 0.205, 1.0, 0.66), "kolb_fig4_17_action_potential.png"),
    # p101 (book p.102) Figure 5.2 major parts of a synapse (panel A, left+center)
    (KOLB, 101, (0.075, 0.08, 0.655, 0.455), "kolb_fig5_2_synapse.png"),
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
