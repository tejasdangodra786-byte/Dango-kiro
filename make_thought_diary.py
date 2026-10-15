#!/usr/bin/env python3
"""
Generates a CBT Thought Diary (Thought Record) PDF with:
- Explanation of every column + its meaning
- A fully worked ANGER example
- A blank 7-day worksheet

Pure-Python PDF writer (no external dependencies). A4 pages.
Grounded in standard CBT sources:
 - J. S. Beck, Cognitive Behavior Therapy: Basics and Beyond (2nd ed., Guilford)
 - Greenberger & Padesky, Mind Over Mood (Guilford)
 - Hawton et al., Cognitive Behaviour Therapy for Psychiatric Problems (Oxford)
"""

import zlib, struct

# ---------- Low-level PDF builder ----------
class PDF:
    def __init__(self):
        self.objs = []          # list of raw object byte strings
        self.pages = []         # page object ids
        self.W, self.H = 595.28, 841.89  # A4 in points

    def _add(self, data: bytes) -> int:
        self.objs.append(data)
        return len(self.objs)   # object number (1-based)

    def add_page(self, content: bytes):
        stream = zlib.compress(content)
        c_id = self._add(b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(stream)
                         + stream + b"\nendstream")
        # placeholder page dict; parent + fonts filled in build()
        self.pages.append((c_id,))
        return c_id

    def build(self, path):
        # Font objects
        f_reg = self._add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        f_bold = self._add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        f_ital = self._add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Oblique /Encoding /WinAnsiEncoding >>")

        # Reserve pages parent id
        pages_id = len(self.objs) + 1 + len(self.pages)  # will compute after
        # Build page objects
        page_ids = []
        for (c_id,) in self.pages:
            res = (b"<< /Font << /F1 %d 0 R /F2 %d 0 R /F3 %d 0 R >> >>" % (f_reg, f_bold, f_ital))
            page_obj = (b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %.2f %.2f] "
                        b"/Resources %s /Contents %d 0 R >>" % (pages_id, self.W, self.H, res, c_id))
            page_ids.append(self._add(page_obj))

        kids = b" ".join(b"%d 0 R" % pid for pid in page_ids)
        pages_obj = (b"<< /Type /Pages /Count %d /Kids [%s] >>" % (len(page_ids), kids))
        real_pages_id = self._add(pages_obj)
        # fix parent references (we assumed pages_id; ensure it matches)
        assert real_pages_id == pages_id, (real_pages_id, pages_id)
        catalog_id = self._add(b"<< /Type /Catalog /Pages %d 0 R >>" % real_pages_id)

        # Write file
        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for i, obj in enumerate(self.objs, start=1):
            offsets.append(len(out))
            out += b"%d 0 obj\n" % i + obj + b"\nendobj\n"
        xref_pos = len(out)
        n = len(self.objs) + 1
        out += b"xref\n0 %d\n" % n
        out += b"0000000000 65535 f \n"
        for off in offsets[1:]:
            out += b"%010d 00000 n \n" % off
        out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF" % (n, catalog_id, xref_pos)
        with open(path, "wb") as f:
            f.write(out)


# ---------- Text layout helpers ----------
# Approx Helvetica char widths (per 1000 units) for wrapping.
_W = {}
def _char_w(ch):
    # coarse average; good enough for wrapping
    if ch in "iIljtf.,;:!'| ":
        return 0.28
    if ch in "mwMW":
        return 0.88
    if ch.isupper():
        return 0.68
    return 0.52

def text_width(s, size):
    return sum(_char_w(c) for c in s) * size

def esc(s):
    return s.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")

def wrap(text, size, max_w):
    words = text.split(" ")
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if text_width(trial, size) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


class Canvas:
    """Accumulates drawing ops for one page's content stream."""
    def __init__(self, W, H):
        self.W, self.H = W, H
        self.ops = []

    def text(self, x, y, s, size=10, font="F1", color=(0, 0, 0)):
        r, g, b = color
        self.ops.append(b"BT /%s %.2f Tf %.3f %.3f %.3f rg %.2f %.2f Td (%s) Tj ET" %
                         (font.encode(), size, r, g, b, x, self.H - y, esc(s).encode("latin-1", "replace")))

    def line(self, x1, y1, x2, y2, w=0.6, color=(0, 0, 0)):
        r, g, b = color
        self.ops.append(b"%.3f %.3f %.3f RG %.2f w %.2f %.2f m %.2f %.2f l S" %
                         (r, g, b, w, x1, self.H - y1, x2, self.H - y2))

    def rect_fill(self, x, y, w, h, color):
        r, g, b = color
        self.ops.append(b"%.3f %.3f %.3f rg %.2f %.2f %.2f %.2f re f" %
                         (r, g, b, x, self.H - y - h, w, h))

    def rect_stroke(self, x, y, w, h, lw=0.6, color=(0, 0, 0)):
        r, g, b = color
        self.ops.append(b"%.3f %.3f %.3f RG %.2f w %.2f %.2f %.2f %.2f re S" %
                         (r, g, b, lw, x, self.H - y - h, w, h))

    def bytes(self):
        return b"\n".join(self.ops)


# ---------- Colors ----------
NAVY = (0.11, 0.20, 0.34)
BLUE = (0.16, 0.35, 0.58)
LBLUE = (0.90, 0.94, 0.98)
GREY = (0.40, 0.40, 0.40)
LGREY = (0.96, 0.96, 0.96)
LINE = (0.75, 0.78, 0.82)
WHITE = (1, 1, 1)
ACCENT = (0.72, 0.15, 0.15)

MARGIN = 42
CW = 595.28 - 2 * MARGIN  # content width


def header(c, title, subtitle=None):
    c.rect_fill(0, 0, c.W, 70, NAVY)
    c.text(MARGIN, 34, title, size=17, font="F2", color=WHITE)
    if subtitle:
        c.text(MARGIN, 54, subtitle, size=9.5, font="F3", color=(0.85, 0.88, 0.93))

def footer(c, page_no, total):
    c.line(MARGIN, c.H - 34, c.W - MARGIN, c.H - 34, w=0.5, color=LINE)
    c.text(MARGIN, c.H - 22, "CBT Thought Diary  -  for educational & self-help use", size=8, font="F3", color=GREY)
    c.text(c.W - MARGIN - 60, c.H - 22, "Page %d of %d" % (page_no, total), size=8, font="F3", color=GREY)


# =====================================================================
# Build document
# =====================================================================
pdf = PDF()
PAGES = []  # list of Canvas


# ---- helper: flowing text with y-cursor ----
class Flow:
    def __init__(self, c, y):
        self.c = c
        self.y = y
    def h1(self, s):
        self.y += 6
        self.c.text(MARGIN, self.y, s, size=13.5, font="F2", color=NAVY)
        self.y += 6
        self.c.line(MARGIN, self.y, c_width(), self.y, w=1.0, color=BLUE)
        self.y += 16
    def h2(self, s, color=BLUE):
        self.y += 4
        self.c.text(MARGIN, self.y, s, size=11, font="F2", color=color)
        self.y += 15
    def para(self, s, size=9.7, font="F1", color=(0.12,0.12,0.12), gap=3, indent=0):
        for ln in wrap(s, size, CW - indent):
            self.c.text(MARGIN + indent, self.y, ln, size=size, font=font, color=color)
            self.y += size + 3.4
        self.y += gap
    def bullet(self, s, size=9.7):
        self.c.text(MARGIN + 6, self.y, "-", size=size, font="F2", color=BLUE)
        lines = wrap(s, size, CW - 22)
        for i, ln in enumerate(lines):
            self.c.text(MARGIN + 18, self.y, ln, size=size, font="F1", color=(0.12,0.12,0.12))
            self.y += size + 3.4
        self.y += 1.5
    def space(self, h=8):
        self.y += h

def c_width():
    return 595.28 - MARGIN


# ---------------------------------------------------------------------
# PAGE 1 - Cover / Introduction
# ---------------------------------------------------------------------
c = Canvas(pdf.W, pdf.H)
c.rect_fill(0, 0, c.W, 210, NAVY)
c.text(MARGIN, 70, "CBT THOUGHT DIARY", size=30, font="F2", color=WHITE)
c.text(MARGIN, 100, "A 7-Day Thought Record Worksheet", size=14, font="F3", color=(0.85,0.88,0.93))
c.rect_fill(MARGIN, 130, 90, 4, ACCENT)
c.text(MARGIN, 165, "Identify -> Examine -> Reframe your thoughts", size=12, font="F1", color=(0.88,0.90,0.94))
c.text(MARGIN, 186, "Includes a fully worked ANGER example", size=11, font="F3", color=(0.80,0.84,0.90))

f = Flow(c, 250)
f.h2("What is a Thought Diary?")
f.para("A thought diary (also called a Thought Record) is the core self-monitoring tool of Cognitive "
       "Behavioural Therapy (CBT). CBT is built on the idea that it is not events themselves that "
       "directly cause our feelings and behaviour, but the THOUGHTS and INTERPRETATIONS we have about "
       "those events. The same situation can lead to very different emotions depending on what we tell "
       "ourselves about it.")
f.para("When we are distressed, fast, automatic thoughts pop into the mind. These 'automatic thoughts' "
       "often feel true, but they are frequently biased, exaggerated, or unhelpful. The thought diary "
       "slows this process down so you can catch a thought, look at the evidence for and against it, and "
       "replace it with a more balanced, realistic view. Over time this changes how you feel and act.")

f.h2("Why use it?")
f.bullet("It makes invisible, automatic thinking visible and easier to question.")
f.bullet("It breaks the loop where Situation -> Thought -> strong Emotion -> unhelpful Behaviour.")
f.bullet("It builds the skill of responding rather than reacting - especially useful for anger, anxiety and low mood.")
f.bullet("It gives you a written record so you can spot recurring thinking patterns over the week.")

f.h2("The CBT model in one line")
f.para("Situation  ->  Automatic Thought  ->  Emotion  ->  Behaviour / Body reaction", font="F2", size=11, color=NAVY)
f.para("The thought diary lets you step into the arrow between Thought and Emotion - the one place where "
       "change is possible.", font="F3", size=9.5, color=GREY)

f.space(6)
c.rect_fill(MARGIN, f.y, CW, 58, LBLUE)
c.rect_stroke(MARGIN, f.y, CW, 58, lw=0.6, color=LINE)
c.text(MARGIN + 12, f.y + 20, "Based on established CBT sources", size=10, font="F2", color=NAVY)
c.text(MARGIN + 12, f.y + 36, "J. S. Beck - Cognitive Behavior Therapy: Basics & Beyond  |  Greenberger & Padesky - Mind Over Mood", size=8.2, font="F1", color=(0.2,0.2,0.2))
c.text(MARGIN + 12, f.y + 49, "Hawton et al. - Cognitive Behaviour Therapy for Psychiatric Problems (Oxford Medical Publications)", size=8.2, font="F1", color=(0.2,0.2,0.2))

PAGES.append(c)


# ---------------------------------------------------------------------
# PAGE 2 - The 7 columns explained (meaning of each)
# ---------------------------------------------------------------------
c = Canvas(pdf.W, pdf.H)
header(c, "The 7 Columns - and what each one means", "Fill them left to right. This is the heart of the thought diary.")
f = Flow(c, 96)

cols = [
    ("1. Situation / Trigger",
     "What happened? Where, when, who with, what were you doing? Write the plain facts only - like a "
     "camera would record them. No opinions or interpretations here.",
     "Example (anger): 'At 6 pm my colleague took credit for my report in the team meeting.'"),
    ("2. Emotions (rate 0-100%)",
     "Name the feeling in one word (angry, hurt, anxious, ashamed) and rate its intensity from 0 to 100%. "
     "Rating helps you notice change later. You can list more than one emotion.",
     "Example: Angry 90%, Humiliated 60%."),
    ("3. Automatic Thoughts",
     "The exact words, images or beliefs that flashed through your mind in that moment. Ask: 'What was "
     "going through my mind just then?' Circle the 'hot thought' - the one most linked to the strongest "
     "emotion.",
     "Example: 'He deliberately disrespected me. He always steals my work. I can't let people walk over me.'"),
    ("4. Evidence FOR the hot thought",
     "Facts (not opinions) that support the hot thought being true. Being honest here keeps the exercise "
     "fair and believable.",
     "Example: 'He did say the idea was his. He didn't mention my name.'"),
    ("5. Evidence AGAINST the hot thought",
     "Facts that do not fit the hot thought, or that another person might point out. This is where you "
     "widen the lens and challenge the bias.",
     "Example: 'The manager knows I wrote it. He may have been nervous and rushed. He has credited me before.'"),
    ("6. Balanced / Alternative Thought",
     "A fairer, more realistic thought that takes BOTH columns of evidence into account. It is not forced "
     "positive thinking - it is accurate thinking. Rate how much you believe it (0-100%).",
     "Example: 'It was frustrating and I can raise it calmly with him. One meeting doesn't mean he always "
     "steals my work.' (Believe 75%)"),
    ("7. Re-rate Emotion & Outcome",
     "Re-rate the original emotion now (0-100%) and note what you will do differently. A drop in intensity "
     "shows the reframe is working.",
     "Example: Angry now 40%. Action: 'Speak to him privately tomorrow and email the manager the original draft.'"),
]

for title, meaning, ex in cols:
    # keep block together; move to page 3 if not enough room
    if f.y > 720:
        footer(c, 2, 5); PAGES.append(c)
        c = Canvas(pdf.W, pdf.H)
        header(c, "The 7 Columns (continued)", "Meaning of each column with the anger example.")
        f = Flow(c, 96)
    c.rect_fill(MARGIN, f.y - 2, 4, 12, ACCENT)
    c.text(MARGIN + 12, f.y + 8, title, size=11, font="F2", color=NAVY)
    f.y += 16
    for ln in wrap(meaning, 9.6, CW - 12):
        c.text(MARGIN + 12, f.y, ln, size=9.6, font="F1", color=(0.12,0.12,0.12)); f.y += 13
    for ln in wrap(ex, 9.2, CW - 12):
        c.text(MARGIN + 12, f.y, ln, size=9.2, font="F3", color=BLUE); f.y += 12.5
    f.y += 8

footer(c, 2, 5)
PAGES.append(c)


# ---------------------------------------------------------------------
# PAGE 3 - Worked ANGER example as a filled table
# ---------------------------------------------------------------------
def draw_table(c, y0, headers, rows, col_w, row_h, head_h=34, fs_head=8.2, fs_body=8.0, head_color=BLUE):
    x0 = MARGIN
    total_w = sum(col_w)
    # header
    c.rect_fill(x0, y0, total_w, head_h, head_color)
    cx = x0
    for i, h in enumerate(headers):
        for j, ln in enumerate(wrap(h, fs_head, col_w[i] - 8)):
            c.text(cx + 4, y0 + 12 + j * (fs_head + 2), ln, size=fs_head, font="F2", color=WHITE)
        cx += col_w[i]
    # rows
    y = y0 + head_h
    for r_idx, row in enumerate(rows):
        rh = row_h[r_idx] if isinstance(row_h, list) else row_h
        if r_idx % 2 == 1:
            c.rect_fill(x0, y, total_w, rh, LGREY)
        cx = x0
        for i, cell in enumerate(row):
            for j, ln in enumerate(wrap(cell, fs_body, col_w[i] - 8)):
                c.text(cx + 4, y + 11 + j * (fs_body + 2.2), ln, size=fs_body, font="F1", color=(0.1,0.1,0.1))
            cx += col_w[i]
        y += rh
    # grid
    total_h = head_h + (sum(row_h) if isinstance(row_h, list) else row_h * len(rows))
    c.rect_stroke(x0, y0, total_w, total_h, lw=0.7, color=LINE)
    cx = x0
    for w in col_w[:-1]:
        cx += w
        c.line(cx, y0, cx, y0 + total_h, w=0.5, color=LINE)
    yy = y0 + head_h
    c.line(x0, yy, x0 + total_w, yy, w=0.7, color=LINE)
    for k in range(len(rows) - 1):
        yy += (row_h[k] if isinstance(row_h, list) else row_h)
        c.line(x0, yy, x0 + total_w, yy, w=0.4, color=LINE)
    return y0 + total_h

c = Canvas(pdf.W, pdf.H)
header(c, "Worked Example: An Anger Situation", "How the diary transforms a hot, angry reaction into a balanced response.")
f = Flow(c, 92)
f.para("Read this completed row before filling your own. Notice how the emotion drops from 90% to 40% once "
       "the hot thought is examined against real evidence.", size=9.6)

headers = ["1. Situation", "2. Emotion (0-100%)", "3. Automatic thoughts (hot thought *)",
           "4. Evidence FOR", "5. Evidence AGAINST", "6. Balanced thought (believe %)", "7. Re-rate + action"]
row = [
    "At 6 pm, in the team meeting, my colleague presented my report as his own idea.",
    "Angry 90%\nHumiliated 60%",
    "* He deliberately disrespected me. He always steals my work. I can't let anyone walk over me.",
    "He said the idea was his. He did not mention my name.",
    "The manager knows I wrote it. He looked nervous and rushed. He has credited me before. 'Always' is an exaggeration.",
    "It was frustrating and unfair, and I can calmly raise it with him. One meeting does not mean he always steals my work. (Believe 75%)",
    "Angry now 40%. Action: speak to him privately tomorrow; email manager the original draft.",
]
col_w = [72, 60, 90, 68, 82, 90, 49.28]  # sums ~ 511 = CW
row_h = [150]
y_end = draw_table(c, f.y, headers, [row], col_w, row_h, head_h=42, fs_head=7.4, fs_body=7.6)

f.y = y_end + 18
f.h2("What this example teaches", color=NAVY)
f.bullet("The 'hot thought' contained thinking traps: mind-reading ('deliberately'), over-generalising ('always'), and an all-or-nothing rule ('walk over me').")
f.bullet("Looking for evidence AGAINST did not deny the problem - it made the view accurate and less explosive.")
f.bullet("The balanced thought still allows action (a calm conversation) but removes the fuel for aggression.")
f.bullet("The measurable drop in anger (90% -> 40%) is the proof the technique worked.")

footer(c, 3, 5)
PAGES.append(c)


# ---------------------------------------------------------------------
# PAGE 4 & 5 - 7-Day blank worksheet
# ---------------------------------------------------------------------
def build_worksheet_page(c, day_start, day_end, page_no):
    header(c, "7-Day Thought Diary - Worksheet", "One row per day. Fill it in as close to the event as you can.")
    f = Flow(c, 92)
    f.para("Tip: rate emotions 0-100%% in column 2, circle the hot thought in column 3, then re-rate in column 7 "
           "to see your progress.", size=9.2, color=GREY)
    headers = ["Day / Date", "Situation (facts)", "Emotion 0-100%", "Automatic thought (*hot)",
               "Evidence FOR", "Evidence AGAINST", "Balanced thought (believe %)", "Re-rate + action"]
    col_w = [46, 78, 44, 80, 62, 66, 78, 57.28]
    rows = []
    for d in range(day_start, day_end + 1):
        rows.append(["Day %d\n____/____" % d, "", "", "", "", "", "", ""])
    rh = [128] * len(rows)
    draw_table(c, f.y, headers, rows, col_w, rh, head_h=40, fs_head=6.9, fs_body=7.5)
    footer(c, page_no, 5)

c = Canvas(pdf.W, pdf.H)
build_worksheet_page(c, 1, 4, 4)
PAGES.append(c)

c = Canvas(pdf.W, pdf.H)
build_worksheet_page(c, 5, 7, 5)
# add a weekly reflection box under the day-7 table
f = Flow(c, 96)
# find space: after 3 rows of 128 + header 40 starting at ~ y=138 -> ~ 138+40+384 = 562
f.y = 600
f.h2("End-of-week reflection", color=NAVY)
c.rect_stroke(MARGIN, f.y, CW, 150, lw=0.7, color=LINE)
prompts = [
    "Which thinking traps showed up most this week (e.g. mind-reading, all-or-nothing, over-generalising)?",
    "Which situations triggered the strongest anger, and what did they have in common?",
    "Which balanced thought helped the most - and how much did my emotion drop?",
    "One thing I will try differently next week:",
]
yy = f.y + 16
for p in prompts:
    for ln in wrap(p, 9, CW - 20):
        c.text(MARGIN + 10, yy, ln, size=9, font="F1", color=(0.12,0.12,0.12)); yy += 12.5
    c.line(MARGIN + 10, yy + 2, c.W - MARGIN - 10, yy + 2, w=0.4, color=LINE)
    yy += 16
PAGES.append(c)


# ---------------------------------------------------------------------
# Assemble
# ---------------------------------------------------------------------
for c in PAGES:
    pdf.add_page(c.bytes())
pdf.build("/projects/sandbox/Dango-kiro/CBT_Thought_Diary_7Day.pdf")
print("PDF written:", len(PAGES), "pages")
