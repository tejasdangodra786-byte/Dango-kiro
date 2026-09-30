#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate a postgraduate journal-club PowerPoint (.pptx) by writing OpenXML directly.
Paper: Efficacy of Mentalization-Based Therapy in Treating Self-Harm:
       A Systematic Review and Meta-Analysis (SLTB, 2024; DOI 10.1111/sltb.13044)

No external libraries are used (python-pptx is unavailable in this sandbox).
A .pptx is a ZIP archive of XML parts; we assemble a valid one.
"""
import os, zipfile, html, datetime

OUT = "Journal_Club_MBT_SelfHarm.pptx"

# ---- Slide canvas (16:9 EMU) ----
EMU = 914400
W = 12192000
H = 6858000

# ---- Pastel palette (hex, no #) ----
LAVENDER = "E7E0F7"
LILAC    = "D6C7EE"
POWDER   = "D7E8F4"
SAGE     = "D9E8D2"
MINT     = "D5EFE6"
PEACH    = "FBE3D2"
BLUSH    = "F7D9E0"
CREAM    = "FBF6EC"
LGREY    = "EEF0F2"
CHARCOAL = "2E2E38"
ACCENT   = "6E5AA6"   # deep lavender for titles/accents
ACCENT2  = "3E7CB1"   # blue accent
WHITE    = "FFFFFF"

def esc(t): return html.escape(str(t), quote=True)

# ------------- low-level shape builders (DrawingML) -------------
_sid = [1]
def nid():
    _sid[0]+=1
    return _sid[0]

def _runs(text, sz, color=CHARCOAL, bold=False, italic=False):
    b = ' b="1"' if bold else ''
    i = ' i="1"' if italic else ''
    return (f'<a:r><a:rPr lang="en-US" sz="{sz}"{b}{i} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r>')

def para(text, sz=1600, color=CHARCOAL, bold=False, italic=False, bullet=False,
         align="l", lvl=0, space_after=400):
    algn = f' algn="{align}"'
    if bullet:
        buXml = (f'<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>')
        marL = 342900 + lvl*342900
        indent = -228600
        pPr = (f'<a:pPr marL="{marL}" indent="{indent}" lvl="{lvl}"{algn}>'
               f'<a:spcAft><a:spcPts val="{space_after}"/></a:spcAft>{buXml}</a:pPr>')
    else:
        pPr = (f'<a:pPr lvl="{lvl}"{algn}><a:spcAft><a:spcPts val="{space_after}"/></a:spcAft>'
               f'<a:buNone/></a:pPr>')
    if text == "":
        return f'<a:p>{pPr}</a:p>'
    return f'<a:p>{pPr}{_runs(text, sz, color, bold, italic)}</a:p>'

def multi_run_para(runs, sz=1600, bullet=False, lvl=0, align="l", space_after=400):
    """runs: list of (text, {color,bold,italic,sz})"""
    if bullet:
        marL = 342900 + lvl*342900
        pPr = (f'<a:pPr marL="{marL}" indent="-228600" lvl="{lvl}" algn="{align}">'
               f'<a:spcAft><a:spcPts val="{space_after}"/></a:spcAft>'
               f'<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>')
    else:
        pPr = (f'<a:pPr lvl="{lvl}" algn="{align}"><a:spcAft><a:spcPts val="{space_after}"/></a:spcAft><a:buNone/></a:pPr>')
    rs = ""
    for r in runs:
        rs += _runs(r[0], r[1].get("sz", sz), r[1].get("color", CHARCOAL),
                    r[1].get("bold", False), r[1].get("italic", False))
    return f'<a:p>{pPr}{rs}</a:p>'

def textbox(x, y, cx, cy, paras, anchor="t", wrap=True):
    body = "".join(paras) if paras else '<a:p><a:pPr><a:buNone/></a:pPr></a:p>'
    an = f' anchor="{anchor}"'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="tb"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="{'square' if wrap else 'none'}"{an} lIns="45720" tIns="27432" rIns="45720" bIns="27432"><a:normAutofit/></a:bodyPr>{body}</p:txBody></p:sp>'''

def roundrect(x, y, cx, cy, fill, paras, line=None, anchor="ctr", shadow=True, rad=None):
    ln = ""
    if line:
        ln = f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    sh = ''
    if shadow:
        sh = ('<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">'
              '<a:srgbClr val="9AA0A6"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst>')
    av = f'<a:gd name="adj" fmla="val {rad}"/>' if rad is not None else ''
    body = "".join(paras) if paras else '<a:p><a:pPr><a:buNone/></a:pPr></a:p>'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst>{av}</a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}{sh}</p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="{anchor}" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:normAutofit/></a:bodyPr>{body}</p:txBody></p:sp>'''

def rect(x, y, cx, cy, fill, paras=None, line=None, anchor="ctr"):
    ln = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    body = "".join(paras) if paras else '<a:p><a:pPr><a:buNone/></a:pPr></a:p>'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="r"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="{anchor}" lIns="45720" tIns="18288" rIns="45720" bIns="18288"><a:normAutofit/></a:bodyPr>{body}</p:txBody></p:sp>'''

def oval(x, y, cx, cy, fill, paras=None, line=None):
    ln = f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else ''
    body="".join(paras) if paras else '<a:p><a:pPr><a:buNone/></a:pPr></a:p>'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="o"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="45720" tIns="18288" rIns="45720" bIns="18288"><a:normAutofit/></a:bodyPr>{body}</p:txBody></p:sp>'''

def arrow_down(x, y, cx, cy, fill=ACCENT):
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="ar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="downArrow"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></p:spPr>
<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>'''

def arrow_right(x, y, cx, cy, fill=ACCENT):
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="ar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></p:spPr>
<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>'''

def line_shape(x, y, cx, cy, color=ACCENT, w=19050):
    return f'''<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="ln"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>'''

# ---- standard slide chrome ----
def banner(title, kicker=None):
    parts = []
    # top accent bar
    parts.append(rect(0, 0, W, 168000, ACCENT))
    # kicker + title band
    ky = 300000
    if kicker:
        parts.append(textbox(560000, 250000, W-1120000, 360000,
            [para(kicker, sz=1300, color=ACCENT2, bold=True, space_after=0)]))
        ky = 560000
    parts.append(textbox(560000, ky, W-1120000, 900000,
        [para(title, sz=2600, color=ACCENT, bold=True, space_after=0)]))
    parts.append(line_shape(560000, ky+ (760000 if kicker else 820000), W-1120000, 0, LILAC, 28575))
    return parts, (ky + (900000 if kicker else 940000))

def footer(page):
    return [
        textbox(560000, H-360000, 6000000, 300000,
                [para("MBT for Self-Harm  \u2022  Systematic Review & Meta-Analysis (SLTB, 2024)",
                      sz=1000, color="9AA0A6", space_after=0)]),
        textbox(W-1400000, H-360000, 900000, 300000,
                [para(f"{page}", sz=1000, color="9AA0A6", align="r", space_after=0)]),
    ]

print("helpers loaded")

# ============================================================
# SLIDE CONTENT
# Each slide = (list_of_shape_xml, speaker_notes_string)
# ============================================================
SLIDES = []
def add(shapes, notes):
    SLIDES.append(("".join(shapes), notes))

CX_L = 560000              # left margin
CONTENT_W = W - 2*CX_L     # usable width

def content_slide(title, kicker, body_paras, page, cardfill=None):
    s, ytop = banner(title, kicker)
    if cardfill:
        s.append(roundrect(CX_L, ytop+120000, CONTENT_W, H-ytop-620000, cardfill, body_paras, anchor="t"))
    else:
        s.append(textbox(CX_L, ytop+120000, CONTENT_W, H-ytop-560000, body_paras, anchor="t"))
    s += footer(page)
    return s

def two_col(title, kicker, left_head, left_items, left_fill, right_head, right_items, right_fill, page):
    s, ytop = banner(title, kicker)
    colw = (CONTENT_W - 300000)//2
    y = ytop+120000
    ch = H - y - 620000
    lp = [para(left_head, sz=1700, color=ACCENT, bold=True, space_after=500)]
    lp += [para(t, sz=1350, bullet=True, space_after=340) for t in left_items]
    rp = [para(right_head, sz=1700, color=ACCENT, bold=True, space_after=500)]
    rp += [para(t, sz=1350, bullet=True, space_after=340) for t in right_items]
    s.append(roundrect(CX_L, y, colw, ch, left_fill, lp, anchor="t"))
    s.append(roundrect(CX_L+colw+300000, y, colw, ch, right_fill, rp, anchor="t"))
    s += footer(page)
    return s

# ---------- SLIDE 1 : TITLE ----------
def slide_title():
    s = []
    s.append(rect(0,0,W,H,CREAM))
    s.append(rect(0,0,W,120000,ACCENT))
    s.append(rect(0,H-120000,W,120000,ACCENT))
    # decorative pastel circles
    s.append(oval(-500000,-500000,2200000,2200000,LAVENDER))
    s.append(oval(W-1600000,H-1600000,2200000,2200000,POWDER))
    s.append(textbox(900000, 900000, W-1800000, 500000,
        [para("M.Phil. CLINICAL PSYCHOLOGY  \u2022  JOURNAL CLUB", sz=1500, color=ACCENT2, bold=True, align="ctr", space_after=0)]))
    s.append(roundrect(900000, 1550000, W-1800000, 1650000, WHITE,
        [para("Efficacy of Mentalization-Based Therapy in Treating Self-Harm",
              sz=3200, color=ACCENT, bold=True, align="ctr", space_after=200),
         para("A Systematic Review and Meta-Analysis", sz=2200, color=CHARCOAL, italic=True, align="ctr", space_after=0)],
        line=LILAC, anchor="ctr"))
    s.append(textbox(900000, 3400000, W-1800000, 900000, [
        multi_run_para([("Journal: ",{"bold":True,"color":ACCENT2,"sz":1500}),
                        ("Suicide and Life-Threatening Behavior (2024) \u2022 DOI: 10.1111/sltb.13044",{"sz":1500})], align="ctr", space_after=200),
        multi_run_para([("Design: ",{"bold":True,"color":ACCENT2,"sz":1400}),
                        ("Systematic review & random-effects meta-analysis of MBT / MBT-A trials",{"sz":1400})], align="ctr", space_after=0),
    ]))
    s.append(roundrect(2400000, 4550000, W-4800000, 1150000, LAVENDER, [
        multi_run_para([("Presented by: ",{"bold":True,"sz":1500,"color":ACCENT}),("Sonal ______________",{"sz":1500})], align="ctr", space_after=160),
        multi_run_para([("M.Phil. Clinical Psychology  \u2022  Institution: ______________",{"sz":1300})], align="ctr", space_after=160),
        multi_run_para([("Journal Club Date: ______________",{"sz":1300})], align="ctr", space_after=0),
    ], line=LILAC))
    notes = ("Open by stating the full title, that it is a 2024 systematic review and meta-analysis in Suicide and "
             "Life-Threatening Behavior (a leading suicidology journal), and that it pools trials of MBT (adults) and "
             "MBT-A (adolescents) for self-harm. Fill in your name, institution and the journal-club date. "
             "Frame the talk: I will summarise the clinical problem, the rationale for MBT, the review method, the "
             "pooled findings, and then critically appraise the evidence and its clinical implications.")
    return s, notes
s,n = slide_title(); add(s,n)

# ---------- SLIDE 2 : WHY THIS PAPER ----------
def slide_why():
    s, ytop = banner("Why This Paper Matters", "CLINICAL & ACADEMIC RELEVANCE")
    y = ytop+120000
    left_w = 5600000
    lp = [
        para("Self-harm is a major clinical and public-health concern with high recurrence.", sz=1350, bullet=True, space_after=320),
        para("It is strongly linked to emotional dysregulation, trauma, personality pathology and elevated suicide risk.", sz=1350, bullet=True, space_after=320),
        para("Repetition is common, so effective psychological interventions are urgently needed.", sz=1350, bullet=True, space_after=320),
        para("Therapies that target underlying processes (not just the behaviour) are of particular interest.", sz=1350, bullet=True, space_after=320),
        para("Meta-analysis synthesises scattered trials into a pooled estimate to guide practice.", sz=1350, bullet=True, space_after=0),
    ]
    s.append(roundrect(CX_L, y, left_w, H-y-620000, CREAM, lp, anchor="t", line=LILAC))
    # right vertical pathway
    rx = CX_L + left_w + 300000
    rw = CONTENT_W - left_w - 300000
    steps = [("Self-harm", BLUSH), ("Repetition", PEACH), ("Suicide risk", LILAC),
             ("Treatment need", POWDER), ("Evidence-based intervention", SAGE)]
    sy = y+40000
    bh = 620000; gap = 200000
    for i,(t,c) in enumerate(steps):
        s.append(roundrect(rx, sy, rw, bh, c, [para(t, sz=1350, bold=True, align="ctr", space_after=0)]))
        if i < len(steps)-1:
            s.append(arrow_down(rx+rw/2-120000, sy+bh+10000, 240000, gap-20000, ACCENT))
        sy += bh+gap
    s += footer(2)
    return s
add(slide_why(),
    "Motivate the paper. Self-harm is common, highly recurrent, and a leading predictor of later suicide. "
    "Individual trials of MBT exist but vary in size and quality, so clinicians lack a clear pooled answer. "
    "Walk down the right-hand pathway: self-harm tends to repeat, repetition raises suicide risk, which creates "
    "a treatment need, which this meta-analysis addresses by evaluating MBT. Emphasise that MBT is theory-driven, "
    "targeting the mentalizing failures thought to underlie impulsive self-harm.")

# ---------- SLIDE 3 : CLINICAL PROBLEM ----------
def slide_problem():
    s, ytop = banner("The Clinical Problem: Defining Self-Harm", "BACKGROUND")
    y = ytop+120000
    colw = (CONTENT_W-300000)//2
    ch = H-y-980000
    lp = [para("What is self-harm?", sz=1600, color=ACCENT, bold=True, space_after=360),
          para("Intentional self-poisoning or self-injury, irrespective of motive or suicidal intent.", sz=1300, bullet=True, space_after=300),
          para("Overlaps with, but is not identical to, suicidal behaviour \u2014 non-suicidal self-injury and suicidal self-injury can co-occur.", sz=1300, bullet=True, space_after=300),
          para("Serves emotional and interpersonal functions (affect relief, communication, self-punishment).", sz=1300, bullet=True, space_after=0)]
    rp = [para("Common forms", sz=1600, color=ACCENT, bold=True, space_after=360)]
    for f in ["Cutting","Burning","Hitting oneself","Scratching","Self-poisoning (overdose)","Interfering with wound healing"]:
        rp.append(para(f, sz=1300, bullet=True, space_after=240))
    s.append(roundrect(CX_L, y, colw, ch, POWDER, lp, anchor="t"))
    s.append(roundrect(CX_L+colw+300000, y, colw, ch, MINT, rp, anchor="t"))
    # caution bar
    s.append(roundrect(CX_L, y+ch+120000, CONTENT_W, 620000, PEACH, [
        multi_run_para([("\u26A0  Clinical caution:  ",{"bold":True,"color":"B5651D","sz":1350}),
                        ("Every episode must be assessed for suicidal intent, lethality, frequency, triggers, functions and current safety.",
                         {"sz":1300})], space_after=0)], line="E0A96D", anchor="ctr"))
    s += footer(3)
    return s
add(slide_problem(),
    "Define self-harm operationally: intentional self-poisoning or self-injury regardless of intent \u2014 the broad "
    "definition used across UK/NICE literature and this review. Stress the self-harm vs suicide distinction: they "
    "overlap but are not the same; NSSI and suicidal self-injury can coexist in the same person. Note the functions "
    "(emotion regulation, communication, self-punishment) because these are exactly what MBT targets. Read the "
    "caution box aloud: assessment of intent, lethality, frequency, triggers, function and safety is non-negotiable.")

print("slides 1-3 done:", len(SLIDES))

# ---------- SLIDE 4 : SELF-HARM & RECURRENCE (cycle) ----------
def slide_cycle():
    s, ytop = banner("Self-Harm & the Recurrence Cycle", "WHY IT REPEATS")
    y = ytop+120000
    # left triggers/maintainers
    lw = 4400000
    lp = [para("Common triggers", sz=1500, color=ACCENT, bold=True, space_after=260)]
    for t in ["Interpersonal conflict / rejection","Abandonment fears","Shame & emotional overwhelm","Dissociation, trauma reminders"]:
        lp.append(para(t, sz=1200, bullet=True, space_after=180))
    lp.append(para("Maintaining factors", sz=1500, color=ACCENT, bold=True, space_after=260))
    for t in ["Temporary emotional relief","Difficulty identifying emotions","Poor impulse control","Interpersonal reinforcement"]:
        lp.append(para(t, sz=1200, bullet=True, space_after=180))
    s.append(roundrect(CX_L, y, lw, H-y-620000, CREAM, lp, anchor="t", line=LILAC))
    # right cycle diagram (vertical chain)
    rx = CX_L+lw+300000
    rw = CONTENT_W-lw-300000
    steps=[("Trigger",BLUSH),("Emotional arousal",PEACH),("Reduced mentalization",LILAC),
           ("Self-harm urge \u2192 act",POWDER),("Temporary relief",MINT),("Shame / distress \u2192 repeat",BLUSH)]
    sy=y; bh=460000; gap=120000
    for i,(t,c) in enumerate(steps):
        s.append(roundrect(rx, sy, rw, bh, c, [para(t, sz=1250, bold=True, align="ctr", space_after=0)]))
        if i<len(steps)-1:
            s.append(arrow_down(rx+rw/2-100000, sy+bh+2000, 200000, gap-8000, ACCENT))
        sy+=bh+gap
    s += footer(4)
    return s
add(slide_cycle(),
    "Explain recurrence as the central treatment challenge. Interpersonal triggers (conflict, rejection, "
    "abandonment, shame) raise arousal; under arousal, mentalizing collapses, an urge appears, self-harm gives "
    "short-term relief, and the subsequent shame feeds the next episode. Point out the maintaining factors on the "
    "left \u2014 relief acts as negative reinforcement. This cycle is why relapse-prevention planning and a therapy "
    "that targets the mentalizing failure (the pivot point) are both needed.")

# ---------- SLIDE 5 : TREATMENT GAP ----------
def slide_gap():
    body = [
        para("Access & delivery", sz=1600, color=ACCENT, bold=True, space_after=300),
        para("Specialist psychological therapies are scarce; treatment response is variable and dropout is high.", sz=1350, bullet=True, space_after=280),
        para("Individuals with severe emotional dysregulation are hard to engage and retain.", sz=1350, bullet=True, space_after=280),
        para("Evidence base", sz=1600, color=ACCENT, bold=True, space_after=300),
        para("Inconsistent findings across studies, heterogeneous protocols, and limited long-term follow-up.", sz=1350, bullet=True, space_after=280),
        para("Unclear which patients benefit most, and limited pooled evidence on self-harm frequency, suicidal behaviour, BPD symptoms and general functioning.", sz=1350, bullet=True, space_after=0),
    ]
    return content_slide("The Existing Treatment Gap", "RATIONALE", body, 5, cardfill=LGREY)
add(slide_gap(),
    "Frame the gap the review addresses. Effective, specialist psychotherapies for self-harm are limited in "
    "availability and show variable response and high dropout, especially in people with marked emotional "
    "dysregulation. Individual MBT trials disagree, use different protocols, and rarely report long follow-up. So "
    "clinicians cannot yet say confidently how well MBT works or for whom \u2014 which is exactly the pooled question "
    "this meta-analysis sets out to answer.")

# ---------- SLIDE 6 : WHY MBT (conceptual model) ----------
def slide_whymbt():
    s, ytop = banner("Why Mentalization-Based Therapy?", "THEORETICAL RATIONALE")
    y = ytop+120000
    s.append(roundrect(CX_L, y, CONTENT_W, 780000, LAVENDER, [
        multi_run_para([("Mentalization = ",{"bold":True,"color":ACCENT,"sz":1400}),
                        ("the capacity to understand oneself and others in terms of intentional mental states \u2014 thoughts, feelings, wishes, beliefs and intentions.",{"sz":1350})], space_after=0)], line=LILAC, anchor="ctr"))
    # horizontal model
    my = y+920000
    steps=[("Attachment threat",BLUSH),("Mentalization failure",PEACH),("Misreading self/others",LILAC),
           ("Emotional dysregulation",POWDER),("Self-harm",BLUSH)]
    bw=2000000; gap=340000; bh=780000
    total = len(steps)*bw + (len(steps)-1)*gap
    sx = (W-total)//2
    for i,(t,c) in enumerate(steps):
        s.append(roundrect(sx, my, bw, bh, c, [para(t, sz=1250, bold=True, align="ctr", space_after=0)]))
        if i<len(steps)-1:
            s.append(arrow_right(sx+bw+40000, my+bh/2-90000, gap-80000, 180000, ACCENT))
        sx+=bw+gap
    # MBT aims
    ay = my+bh+220000
    s.append(roundrect(CX_L, ay, CONTENT_W, H-ay-560000, MINT, [
        para("MBT aims to:", sz=1450, color=ACCENT, bold=True, space_after=240),
        multi_run_para([("\u2022 improve understanding of mental states   \u2022 increase emotional awareness   \u2022 reduce certainty about negative interpretations",{"sz":1250})], space_after=180),
        multi_run_para([("\u2022 strengthen affect regulation   \u2022 improve interpersonal functioning   \u2022 reduce impulsive, self-destructive behaviour",{"sz":1250})], space_after=0),
    ], line=SAGE, anchor="t"))
    s += footer(6)
    return s
add(slide_whymbt(),
    "Define mentalization plainly, then give the causal model: an attachment threat or interpersonal stressor "
    "triggers a temporary failure of mentalizing; the person misreads their own and others' intentions, becomes "
    "dysregulated, and self-harms. MBT intervenes at the mentalizing-failure link. List its aims: rebuild reflective "
    "understanding, increase emotional awareness, loosen rigid negative interpretations, and improve affect "
    "regulation and relationships \u2014 which should, in turn, reduce impulsive self-harm.")

print("slides through 6:", len(SLIDES))

# ---------- SLIDE 7 : INTRODUCTION ----------
def slide_intro():
    s, ytop = banner("Introduction: Key Concepts", "BACKGROUND")
    y = ytop+120000
    cw=(CONTENT_W-2*260000)//3
    ch=H-y-620000
    c1=[para("Self-harm", sz=1500, color=ACCENT, bold=True, space_after=240),
        para("Intentional self-injury/poisoning; multiple forms & functions; a strong predictor of later suicide.", sz=1200, bullet=True, space_after=0)]
    c2=[para("Mentalization", sz=1500, color=ACCENT, bold=True, space_after=240),
        para("Understanding behaviour via mental states; develops in secure attachment; fails under emotional arousal.", sz=1200, bullet=True, space_after=0)]
    c3=[para("MBT", sz=1500, color=ACCENT, bold=True, space_after=240),
        para("Structured psychotherapy (Bateman & Fonagy). Curious, not-knowing stance; focus on current affect & interpersonal experience.", sz=1200, bullet=True, space_after=0)]
    s.append(roundrect(CX_L, y, cw, ch, POWDER, c1, anchor="t"))
    s.append(roundrect(CX_L+cw+260000, y, cw, ch, LAVENDER, c2, anchor="t"))
    s.append(roundrect(CX_L+2*(cw+260000), y, cw, ch, MINT, c3, anchor="t"))
    s += footer(7)
    return s
add(slide_intro(),
    "Give three tight definitions. Self-harm: intentional injury/poisoning, a strong suicide predictor. "
    "Mentalization: reading behaviour through mental states; it is attachment-based and collapses under stress. "
    "MBT: the structured psychotherapy developed by Bateman and Fonagy, using a curious, not-knowing stance that "
    "keeps attention on present affect and interpersonal experience. Also remind the audience why we meta-analyse: "
    "to pool small trials into a more precise effect estimate and formally examine heterogeneity and bias.")

# ---------- SLIDE 8 : CONCEPTUAL FRAMEWORK ----------
def slide_framework():
    s, ytop = banner("Conceptual Framework", "MECHANISTIC PATHWAY")
    y = ytop+80000
    steps=[("Attachment insecurity / interpersonal stress",LILAC),
           ("Reduced mentalization under emotional arousal",PEACH),
           ("Misinterpretation of self and others",BLUSH),
           ("Affect dysregulation & impulsivity",POWDER),
           ("Self-harm behaviour \u2192 short-term relief",MINT),
           ("MBT targets mentalizing, affect regulation & interpersonal understanding",SAGE),
           ("Reduced self-harm & improved clinical functioning",LAVENDER)]
    bh=520000; gap=90000
    sx=CX_L+1500000; bw=CONTENT_W-3000000
    sy=y
    for i,(t,c) in enumerate(steps):
        bold = i>=5
        s.append(roundrect(sx, sy, bw, bh, c, [para(t, sz=1250, bold=bold, align="ctr", space_after=0)], line=(ACCENT if bold else None)))
        if i<len(steps)-1:
            s.append(arrow_down(sx+bw/2-90000, sy+bh+2000, 180000, gap-8000, ACCENT))
        sy+=bh+gap
    s += footer(8)
    return s
add(slide_framework(),
    "This is the integrating model for the whole talk. Read top to bottom: attachment/interpersonal stress reduces "
    "mentalizing; the person misreads self and others; dysregulation and impulsivity follow; self-harm brings brief "
    "relief but long-term distress. The last two green/lavender boxes are the therapeutic hypothesis: MBT works "
    "upstream on mentalizing and affect regulation, which should reduce self-harm and improve functioning. Keep this "
    "picture in mind when we reach the results.")

# ---------- SLIDE 9 : REVIEW OF LITERATURE (table) ----------
def table(x, y, col_w, row_h, headers, rows, head_fill=ACCENT, head_color=WHITE,
          zebra=(WHITE, LGREY), font=1100, head_font=1150):
    shapes=[]
    cx=x
    # header
    for j,htxt in enumerate(headers):
        shapes.append(rect(cx, y, col_w[j], row_h, head_fill,
            [para(htxt, sz=head_font, color=head_color, bold=True, align="l", space_after=0)]))
        cx+=col_w[j]
    ry=y+row_h
    for i,row in enumerate(rows):
        cx=x
        fill=zebra[i%2]
        for j,cell in enumerate(row):
            shapes.append(rect(cx, ry, col_w[j], row_h, fill,
                [para(cell, sz=font, align="l", space_after=0)], line="D8DCE0"))
            cx+=col_w[j]
        ry+=row_h
    return shapes

def slide_litreview():
    s, ytop = banner("Review of Literature", "KEY CONSTITUENT / RELATED STUDIES")
    y = ytop+100000
    headers=["Study (year)","Setting / N","Intervention vs control","Main finding"]
    cw=[2500000, 2200000, 3200000, CONTENT_W-2500000-2200000-3200000]
    rows=[
        ["Rossouw & Fonagy (2012)","UK; N=80 adolescents","MBT-A vs TAU (12 mo)","MBT-A superior for self-harm & depression; mediated by \u2191 mentalizing"],
        ["Bateman & Fonagy (2009)","UK; adults, BPD","MBT vs structured clinical mgmt","Fewer suicide attempts & self-harm; symptom reduction"],
        ["Laurenssen et al. (2018)","NL; adolescents","MBT-A vs TAU","No clear MBT-A advantage \u2014 mixed / null result"],
        ["Beck et al. (2020)","DK; adolescents, BPD","MBT-A vs TAU","No significant group difference on self-harm"],
        ["Griffiths et al. (2019)","UK; MBT-Ai (group)","Feasibility RCT","Acceptable & feasible; underpowered for efficacy"],
    ]
    s += table(CX_L, y, cw, 620000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1050, head_font=1150)
    s.append(textbox(CX_L, y+620000*6+40000, CONTENT_W, 300000,
        [para("Note: findings are mixed \u2014 early positive trials (Rossouw & Fonagy) contrast with later null/feasibility trials. Sources cross-checked with PubMed.",
              sz=1000, color="9AA0A6", italic=True, space_after=0)]))
    s += footer(9)
    return s
add(slide_litreview(),
    "Summarise the landscape of trials that feed this meta-analysis. The landmark is Rossouw & Fonagy (2012): 80 "
    "adolescents, MBT-A beat TAU on self-harm and depression, and the effect was mediated by improved mentalizing "
    "and reduced attachment avoidance. Bateman & Fonagy's adult BPD work is the parent evidence base. Crucially, "
    "later trials (Laurenssen 2018, Beck 2020) were null or mixed, and some (Griffiths) were only feasibility "
    "studies. Flag this heterogeneity now \u2014 it is central to interpreting the pooled effect and its wide "
    "confidence intervals. Verify exact citations against the paper's reference list.")

print("slides through 9:", len(SLIDES))

# ---------- SLIDE 10 : GRAPHICAL COMPARISON (bar of study sizes) ----------
def hbar(x, y, max_w, h, val, maxval, fill, label, valtxt):
    shapes=[]
    bw=int(max_w*val/maxval)
    shapes.append(textbox(x-1900000, y-30000, 1850000, h+60000, [para(label, sz=1050, align="r", space_after=0)], anchor="ctr"))
    shapes.append(rect(x, y, max_w, h, WHITE, line="E0E4E8"))
    shapes.append(roundrect(x, y, max(bw,60000), h, fill, [], rad=8000, shadow=False))
    shapes.append(textbox(x+max(bw,60000)+40000, y-30000, 900000, h+60000, [para(valtxt, sz=1050, bold=True, space_after=0)], anchor="ctr"))
    return shapes

def slide_compare():
    s, ytop = banner("Graphical Comparison of Constituent Trials", "SAMPLE SIZE (illustrative)")
    y=ytop+240000
    data=[("Rossouw & Fonagy 2012",80,BLUSH),("Bateman & Fonagy 2009",134,PEACH),
          ("Laurenssen 2018",109,LILAC),("Beck 2020",111,POWDER),("Griffiths 2019 (pilot)",53,SAGE)]
    maxv=140; bx=CX_L+2000000; bmax=CONTENT_W-2000000-1000000; bh=520000; gap=260000
    sy=y
    for lbl,v,c in data:
        s += hbar(bx, sy, bmax, bh, v, maxv, c, lbl, f"N={v}")
        sy+=bh+gap
    s.append(textbox(CX_L, sy+40000, CONTENT_W, 400000,
        [para("Small samples (all N<150) \u2192 limited power and wide confidence intervals in the pooled analysis. "
              "Values illustrative from published trial reports; confirm against the paper.",
              sz=1050, color="9AA0A6", italic=True, space_after=0)]))
    s += footer(10)
    return s
add(slide_compare(),
    "This bar chart makes the small-sample problem visible: every constituent trial enrolled fewer than ~150 "
    "participants. Small samples mean low statistical power in the individual studies and wide confidence intervals "
    "even after pooling. Use this to set up the appraisal: a large pooled effect size from small, heterogeneous "
    "trials should be read cautiously. State clearly these Ns are taken from the published trial reports and should "
    "be checked against the review's data-extraction table.")

# ---------- SLIDE 11 : RESEARCH GAP ----------
def slide_researchgap():
    body=[
        para("Few well-powered randomized controlled trials; small samples dominate.", sz=1300, bullet=True, space_after=260),
        para("Inconsistent definitions of self-harm and differing outcome measures.", sz=1300, bullet=True, space_after=260),
        para("Varied MBT formats (individual, group, MBT-A) and treatment durations.", sz=1300, bullet=True, space_after=260),
        para("Heterogeneous control conditions (TAU, SCM, waiting list).", sz=1300, bullet=True, space_after=260),
        para("Limited long-term follow-up and inconsistent adherence/fidelity reporting.", sz=1300, bullet=True, space_after=260),
        para("Little evidence from culturally diverse / low-resource settings; unclear mechanisms of change.", sz=1300, bullet=True, space_after=260),
        para("Limited head-to-head comparison with other evidence-based therapies (DBT, CBT).", sz=1300, bullet=True, space_after=0),
    ]
    return content_slide("Research Gap", "WHAT PRIOR LITERATURE LEAVES UNRESOLVED", body, 11, cardfill=LGREY)
add(slide_researchgap(),
    "Consolidate the gaps that justify a synthesis: too few adequately powered RCTs, inconsistent self-harm "
    "definitions and outcome tools, mixed MBT formats and durations, heterogeneous controls, short follow-up, and "
    "weak reporting of adherence. There is little evidence outside Western settings and unclear mechanisms of change. "
    "These gaps are precisely why the authors pooled the trials \u2014 and also why we must temper confidence in any "
    "single pooled number.")

# ---------- SLIDE 12 : RATIONALE ----------
def slide_rationale():
    body=[
        para("Synthesise the scattered evidence into a single pooled estimate of MBT efficacy for self-harm.", sz=1350, bullet=True, space_after=300),
        para("Quantify effects on self-harm and related outcomes (BPD symptoms, depression).", sz=1350, bullet=True, space_after=300),
        para("Assess consistency across studies and identify sources of heterogeneity.", sz=1350, bullet=True, space_after=300),
        para("Inform clinical decision-making about when and for whom to offer MBT.", sz=1350, bullet=True, space_after=300),
        para("Define priorities for future, better-designed trials.", sz=1350, bullet=True, space_after=0),
    ]
    return content_slide("Rationale of the Study", "WHY A SYSTEMATIC REVIEW & META-ANALYSIS", body, 12, cardfill=MINT)
add(slide_rationale(),
    "State why the review was needed: to convert a set of small, conflicting trials into one pooled, more precise "
    "estimate; to quantify MBT's effect not only on self-harm but on BPD symptoms and depression; to test how "
    "consistent the trials are; and to give clinicians and future researchers a clearer evidence signal. Emphasise "
    "the dual clinical and research payoff.")

# ---------- SLIDE 13 : AIM, RESEARCH QUESTION, HYPOTHESIS ----------
def slide_aim():
    s, ytop = banner("Aim, Research Question & Review Objective", "OBJECTIVES")
    y=ytop+120000
    ch=(H-y-620000-2*180000)//3
    s.append(roundrect(CX_L, y, CONTENT_W, ch, LAVENDER, [
        multi_run_para([("Aim:  ",{"bold":True,"color":ACCENT,"sz":1450}),
                        ("To systematically review and quantitatively synthesise the evidence on the efficacy of MBT in reducing self-harm and improving related clinical outcomes.",{"sz":1300})], space_after=0)], line=LILAC, anchor="ctr"))
    s.append(roundrect(CX_L, y+ch+180000, CONTENT_W, ch, POWDER, [
        multi_run_para([("Research question:  ",{"bold":True,"color":ACCENT,"sz":1450}),
                        ("Is MBT effective in reducing self-harm and related outcomes compared with control / treatment-as-usual conditions?",{"sz":1300})], space_after=0)], line=ACCENT2, anchor="ctr"))
    s.append(roundrect(CX_L, y+2*(ch+180000), CONTENT_W, ch, MINT, [
        multi_run_para([("Review objective (in lieu of a formal hypothesis):  ",{"bold":True,"color":ACCENT,"sz":1400}),
                        ("MBT is expected to be associated with a significant reduction in self-harm and related symptoms vs control. As a meta-analysis, the study used predefined review questions rather than an experimental hypothesis.",{"sz":1250})], space_after=0)], line=SAGE, anchor="ctr"))
    s += footer(13)
    return s
add(slide_aim(),
    "Present the aim, question and objective together. Aim: synthesise the efficacy evidence for MBT on self-harm "
    "and related outcomes. Research question: does MBT reduce self-harm versus control/TAU? Because this is a "
    "systematic review, there is no conventional experimental hypothesis \u2014 the expectation is a significant "
    "reduction favouring MBT, tested through predefined review questions. Adapt wording to the paper's exact "
    "statement of objectives.")

print("slides through 13:", len(SLIDES))

# ---------- SLIDE 14 : PICO ----------
def slide_pico():
    s, ytop = banner("Research Question in PICO Format", "FRAMING")
    y=ytop+120000
    cw=(CONTENT_W-3*220000)//4
    ch=H-y-620000
    cards=[("P \u2014 Population","Individuals with self-harm, BPD or emotional dysregulation (adolescents & adults).",BLUSH),
           ("I \u2014 Intervention","Mentalization-based therapy (MBT / MBT-A), individual or group.",LAVENDER),
           ("C \u2014 Comparison","Treatment as usual, structured clinical management, waiting list or another therapy.",POWDER),
           ("O \u2014 Outcomes","Self-harm frequency/severity, suicidal behaviour, BPD symptoms, depression, functioning, retention.",MINT)]
    cx=CX_L
    for head,txt,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [
            para(head, sz=1500, color=ACCENT, bold=True, space_after=300),
            para(txt, sz=1200, space_after=0)], anchor="t"))
        cx+=cw+220000
    s += footer(14)
    return s
add(slide_pico(),
    "Structure the question with PICO. Population: people who self-harm, including those with BPD or emotional "
    "dysregulation, spanning adolescents and adults. Intervention: MBT or its adolescent form MBT-A. Comparison: "
    "usual care, structured clinical management, waiting list, or another active therapy. Outcomes: self-harm "
    "frequency and severity as primary, with suicidal behaviour, BPD symptoms, depression, functioning and "
    "treatment retention as secondary. This mapping mirrors the review's eligibility logic.")

# ---------- SLIDE 15 : METHODOLOGY overview ----------
def slide_method():
    s, ytop = banner("Methodology Overview", "SYSTEMATIC REVIEW WORKFLOW")
    y=ytop+120000
    stages=[("Protocol &\nsearch strategy",LAVENDER),("Screening &\nstudy selection",POWDER),
            ("Data\nextraction",MINT),("Risk-of-bias\nassessment",PEACH),("Meta-analysis\n(random effects)",SAGE)]
    bw=1950000; gap=280000; bh=1000000
    total=len(stages)*bw+(len(stages)-1)*gap
    sx=(W-total)//2; sy=y+120000
    for i,(t,c) in enumerate(stages):
        lines=t.split("\n")
        ps=[para(lines[0], sz=1250, bold=True, align="ctr", space_after=60)]
        for extra in lines[1:]:
            ps.append(para(extra, sz=1250, bold=True, align="ctr", space_after=0))
        s.append(roundrect(sx, sy, bw, bh, c, ps, anchor="ctr", line=ACCENT))
        if i<len(stages)-1:
            s.append(arrow_right(sx+bw+30000, sy+bh/2-90000, gap-60000, 180000, ACCENT))
        sx+=bw+gap
    s.append(roundrect(CX_L, sy+bh+280000, CONTENT_W, 1100000, CREAM, [
        para("Analytic choices", sz=1350, color=ACCENT, bold=True, space_after=200),
        multi_run_para([("Random-effects model \u2022 effect size = Hedges' g (standardized mean difference) \u2022 heterogeneity via I\u00B2 \u2022 sensitivity & subgroup checks \u2022 publication-bias assessment.",{"sz":1250})], space_after=0)], line=LILAC, anchor="t"))
    s += footer(15)
    return s
add(slide_method(),
    "Walk the standard SR/MA pipeline left to right: register a protocol and run a systematic multi-database search; "
    "screen titles/abstracts then full texts against eligibility criteria; extract data; assess risk of bias; and "
    "pool with a random-effects meta-analysis. Highlight the analytic choices box: the effect metric is Hedges' g "
    "(a standardized mean difference suited to different measurement scales), heterogeneity is quantified with I\u00B2, "
    "and the authors used random effects because the trials differ. Note exact databases/dates should be taken from "
    "the paper's methods.")

# ---------- SLIDE 16 : ELIGIBILITY ----------
def slide_eligibility():
    return two_col("Inclusion & Exclusion Criteria","ELIGIBILITY",
        "Inclusion",
        ["Participants with documented self-harm or a relevant clinical diagnosis (e.g., BPD)",
         "Studies evaluating MBT / MBT-A",
         "Report a self-harm or related clinical outcome",
         "Controlled or eligible comparative designs",
         "Sufficient quantitative data to compute an effect size",
         "Published within the specified period & eligible language"],
        SAGE,
        "Exclusion",
        ["No mentalization-based intervention",
         "No relevant self-harm / clinical outcome",
         "Single-case reports or case series",
         "Conference abstracts lacking extractable data",
         "Duplicate publications",
         "Studies without extractable outcome data"],
        BLUSH, 16)
add(slide_eligibility(),
    "Present eligibility as the filter that defines the evidence base. Inclusion: participants who self-harm or "
    "carry a relevant diagnosis, an MBT/MBT-A intervention, a self-harm or related outcome, a comparative design, "
    "and enough numeric data to compute Hedges' g. Exclusion: anything without MBT, without a relevant outcome, or "
    "without extractable data, plus case reports, bare abstracts and duplicates. Tell the audience to map these onto "
    "the paper's exact wording, as small criterion differences change which trials qualify.")

print("slides through 16:", len(SLIDES))

# ---------- SLIDE 17 : SEARCH STRATEGY ----------
def slide_search():
    s, ytop = banner("Search Strategy", "IDENTIFYING STUDIES")
    y=ytop+120000
    lw=5400000
    lp=[para("Databases (typical for this field)", sz=1400, color=ACCENT, bold=True, space_after=240)]
    for d in ["PubMed / MEDLINE","PsycINFO","Embase","Cochrane CENTRAL","Reference-list & hand-searching"]:
        lp.append(para(d, sz=1250, bullet=True, space_after=200))
    lp.append(para("Confirm the exact databases, date range and final search date from the paper.", sz=1050, italic=True, color="9AA0A6", space_after=0))
    s.append(roundrect(CX_L, y, lw, H-y-620000, POWDER, lp, anchor="t"))
    rx=CX_L+lw+300000; rw=CONTENT_W-lw-300000
    s.append(roundrect(rx, y, rw, H-y-620000, CREAM, [
        para("Example concept blocks (Boolean)", sz=1350, color=ACCENT, bold=True, space_after=260),
        multi_run_para([('"mentalization-based therapy" OR "MBT"',{"sz":1250,"bold":True})], space_after=120),
        para("AND", sz=1250, color=ACCENT2, bold=True, align="ctr", space_after=120),
        multi_run_para([('"self-harm" OR "self-injury" OR "suicidal behaviour"',{"sz":1250,"bold":True})], space_after=120),
        para("AND", sz=1250, color=ACCENT2, bold=True, align="ctr", space_after=120),
        multi_run_para([('"borderline personality disorder" OR "emotional dysregulation"',{"sz":1250,"bold":True})], space_after=200),
        para("Illustrative only \u2014 not the verified search string.", sz=1000, italic=True, color="9AA0A6", space_after=0),
    ], line=LILAC, anchor="t"))
    s += footer(17)
    return s
add(slide_search(),
    "Describe how studies were found: systematic searching of the major psychology/medicine databases (PubMed, "
    "PsycINFO, Embase, Cochrane CENTRAL), supplemented by reference-list and hand-searching. Show the three Boolean "
    "concept blocks \u2014 intervention terms AND self-harm terms AND diagnosis terms. Be explicit that this string is "
    "illustrative; read the actual databases, date range and final-search date from the methods so nothing is "
    "misattributed to the authors.")

# ---------- SLIDE 18 : PRISMA ----------
def slide_prisma():
    s, ytop = banner("PRISMA Study-Selection Flow", "TEMPLATE \u2014 INSERT VERIFIED NUMBERS")
    y=ytop+120000
    # left column: identification -> screening -> eligibility -> included
    lx=CX_L+300000; lw=5200000
    stages=[("Records identified through database searching (n = __)",LAVENDER),
            ("Records after duplicates removed (n = __)",POWDER),
            ("Records screened by title/abstract (n = __)",MINT),
            ("Full-text articles assessed for eligibility (n = __)",PEACH),
            ("Studies included in qualitative synthesis (n = __)",SAGE),
            ("Studies included in meta-analysis (n = __)",LAVENDER)]
    bh=560000; gap=180000; sy=y
    for i,(t,c) in enumerate(stages):
        s.append(roundrect(lx, sy, lw, bh, c, [para(t, sz=1200, bold=True, align="ctr", space_after=0)], line=ACCENT))
        if i<len(stages)-1:
            s.append(arrow_down(lx+lw/2-80000, sy+bh+4000, 160000, gap-16000, ACCENT))
        sy+=bh+gap
    # right column: exclusions
    ex=lx+lw+500000; ew=CONTENT_W-(lx-CX_L)-lw-500000
    s.append(roundrect(ex, y+2*(bh+gap), ew, bh, BLUSH,
        [para("Excluded after screening (n = __)", sz=1150, bold=True, align="ctr", space_after=0)], line="D98B9E"))
    s.append(arrow_right(ex-460000, y+2*(bh+gap)+bh/2-80000, 420000, 160000, "D98B9E"))
    s.append(roundrect(ex, y+3*(bh+gap), ew, bh+120000, BLUSH,
        [para("Full texts excluded, with reasons (n = __): no MBT / no relevant outcome / no extractable data / duplicate",
              sz=1050, align="ctr", space_after=0)], line="D98B9E"))
    s.append(arrow_right(ex-460000, y+3*(bh+gap)+bh/2-80000, 420000, 160000, "D98B9E"))
    s += footer(18)
    return s
add(slide_prisma(),
    "This is the PRISMA flow the review reports. Fill each 'n = __' from the paper's PRISMA diagram: records "
    "identified, duplicates removed, records screened, full texts assessed, and studies included in the qualitative "
    "and quantitative syntheses. The right-hand pink boxes capture exclusions with reasons. Do NOT invent these "
    "counts \u2014 transcribe them exactly. Verbally note the funnel from many records down to the small number of "
    "trials that were actually poolable.")

# ---------- SLIDE 19 : PARTICIPANTS ----------
def slide_participants():
    s, ytop = banner("Participant Characteristics", "WHO WAS STUDIED")
    y=ytop+120000
    cw=(CONTENT_W-3*220000)//4
    ch=1500000
    cards=[("Age","Adolescents (MBT-A) & adults (MBT); confirm ranges from paper",POWDER),
           ("Sex","Predominantly female in the constituent trials (e.g., ~85% in Rossouw & Fonagy 2012)",BLUSH),
           ("Diagnosis","Self-harm \u00B1 BPD / emerging BPD traits; emotional dysregulation",LAVENDER),
           ("Setting","Mostly outpatient mental-health services in high-income (largely European) countries",MINT)]
    cx=CX_L
    for h,t,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [para(h, sz=1500, color=ACCENT, bold=True, space_after=240),
                                              para(t, sz=1150, space_after=0)], anchor="t"))
        cx+=cw+220000
    s.append(roundrect(CX_L, y+ch+220000, CONTENT_W, H-(y+ch+220000)-560000, CREAM, [
        multi_run_para([("Total studies and pooled N: ",{"bold":True,"color":ACCENT,"sz":1300}),
                        ("insert from the paper's characteristics table. Samples are small and skew female, adolescent/young-adult, and Western \u2014 note this for generalisability.",{"sz":1250})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(19)
    return s
add(slide_participants(),
    "Describe the pooled sample. Ages span adolescents (MBT-A trials) and adults (MBT trials). Samples are "
    "predominantly female \u2014 for example about 85% female in Rossouw & Fonagy (2012). Most participants have "
    "self-harm with BPD or emerging BPD traits, treated in outpatient services in high-income, largely European "
    "settings. Insert the exact number of studies and total N from the characteristics table, and flag the female, "
    "young, Western skew because it limits generalisability.")

print("slides through 19:", len(SLIDES))

# ---------- SLIDE 20 : MBT PROTOCOL + CORE PRINCIPLES ----------
def slide_protocol():
    s, ytop = banner("MBT Treatment Protocol & Core Principles", "THE INTERVENTION")
    y=ytop+120000
    lw=5600000
    lp=[para("Delivery (as reported across trials)", sz=1400, color=ACCENT, bold=True, space_after=240)]
    for t in ["Individual and/or group MBT (MBT-A adds a parent/family strand)",
              "Typically ~12 months in landmark adolescent trial; frequency varies",
              "Delivered by trained therapists with supervision",
              "Outpatient delivery; crisis management alongside therapy",
              "Fidelity/adherence monitoring inconsistently reported"]:
        lp.append(para(t, sz=1200, bullet=True, space_after=200))
    lp.append(para("Confirm session number, duration and format from each included study.", sz=1000, italic=True, color="9AA0A6", space_after=0))
    s.append(roundrect(CX_L, y, lw, H-y-620000, POWDER, lp, anchor="t"))
    rx=CX_L+lw+300000; rw=CONTENT_W-lw-300000
    rp=[para("Core therapeutic stance", sz=1400, color=ACCENT, bold=True, space_after=240)]
    for t in ["Curiosity & 'not-knowing' stance","Focus on current mental states & affect",
              "Clarify, explore alternative perspectives","Regulate arousal; repair misunderstandings",
              "Stay in the interpersonal context","Avoid premature interpretation"]:
        rp.append(para(t, sz=1200, bullet=True, space_after=180))
    s.append(roundrect(rx, y, rw, H-y-620000, MINT, rp, anchor="t"))
    s += footer(20)
    return s
add(slide_protocol(),
    "Describe what MBT actually involves. Delivery ranges across trials: individual and/or group sessions, with "
    "MBT-A adding a family/parent component; the landmark adolescent trial ran about 12 months. Therapists are "
    "trained and supervised, and treatment is usually outpatient with crisis management running alongside. Note that "
    "fidelity reporting is patchy. On the right, summarise the core stance: curiosity, not-knowing, focus on present "
    "affect and mental states, exploring alternatives, regulating arousal and repairing ruptures \u2014 deliberately "
    "avoiding premature interpretation. Confirm session specifics per study.")

# ---------- SLIDE 21 : CONTROL CONDITIONS ----------
def slide_controls():
    s, ytop = banner("Control / Comparison Conditions", "WHAT MBT WAS COMPARED AGAINST")
    y=ytop+120000
    headers=["Comparator","What it involves","Implication for interpretation"]
    cw=[3000000, 4200000, CONTENT_W-3000000-4200000]
    rows=[
        ["Treatment as usual (TAU)","Routine care, often variable & non-specific","Most common comparator; effect may partly reflect TAU quality"],
        ["Structured clinical management","Manualised generic support (active control)","Sterner test \u2014 smaller MBT advantage expected"],
        ["Waiting list / no treatment","No active therapy","Tends to inflate apparent effect sizes"],
        ["Another psychotherapy","e.g., DBT / supportive therapy","Head-to-head; rare in the MBT self-harm literature"],
    ]
    s += table(CX_L, y, cw, 700000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1100, head_font=1200)
    s.append(textbox(CX_L, y+700000*5+40000, CONTENT_W, 360000,
        [para("Heterogeneous comparators are a key driver of between-study variance and complicate a single pooled estimate.",
              sz=1050, italic=True, color="9AA0A6", space_after=0)]))
    s += footer(21)
    return s
add(slide_controls(),
    "Explain why the comparator matters. Most trials use treatment as usual, which is variable and non-specific, so "
    "part of MBT's apparent benefit may reflect weak usual care. A structured-clinical-management control is a much "
    "sterner test and typically shrinks the MBT advantage. Waiting-list/no-treatment controls tend to inflate effect "
    "sizes. Head-to-head comparisons with DBT are rare. Because the control conditions differ so much across trials, "
    "they are a major source of heterogeneity in the pooled estimate.")

# ---------- SLIDE 22 : OUTCOME MEASURES ----------
def slide_outcomes():
    s, ytop = banner("Outcome Measures", "PRIMARY & SECONDARY")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-620000
    lp=[para("Primary", sz=1600, color=ACCENT, bold=True, space_after=300)]
    for t in ["Frequency / number of self-harm episodes","Self-harm severity","Time to recurrence","Suicidal behaviour"]:
        lp.append(para(t, sz=1300, bullet=True, space_after=240))
    rp=[para("Secondary", sz=1600, color=ACCENT, bold=True, space_after=300)]
    for t in ["BPD symptoms","Depression & anxiety","Emotion dysregulation & impulsivity","Global functioning / quality of life","Treatment retention & hospital use"]:
        rp.append(para(t, sz=1300, bullet=True, space_after=210))
    s.append(roundrect(CX_L, y, cw, ch, BLUSH, lp, anchor="t"))
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, POWDER, rp, anchor="t"))
    s += footer(22)
    return s
add(slide_outcomes(),
    "Organise outcomes into primary and secondary. Primary: self-harm frequency/number of episodes, severity, time "
    "to recurrence, and suicidal behaviour. Secondary: BPD symptoms, depression and anxiety, emotion dysregulation "
    "and impulsivity, functioning/quality of life, and treatment retention or hospital use. Note that trials used "
    "different instruments for the same construct \u2014 one reason a standardized effect size (Hedges' g) was needed "
    "to pool them. List the specific scales the paper reports where available.")

# ---------- SLIDE 23 : RISK OF BIAS (traffic light) ----------
def tl_dot(x,y,d,color):
    return oval(x,y,d,d,color,[],line="FFFFFF")
def slide_rob():
    s, ytop = banner("Risk-of-Bias Assessment", "TEMPLATE \u2014 CONFIRM RATINGS FROM PAPER")
    y=ytop+160000
    domains=["Random sequence generation","Allocation concealment","Blinding of outcome assessment",
             "Incomplete outcome data","Selective reporting","Treatment fidelity"]
    # illustrative pattern per study
    GREEN="8FCB9B"; YELLOW="F2D479"; RED="E39B9B"
    studies=["Rossouw 2012","Bateman 2009","Laurenssen 2018","Beck 2020"]
    pattern={
      "Rossouw 2012":[GREEN,YELLOW,YELLOW,GREEN,GREEN,YELLOW],
      "Bateman 2009":[GREEN,GREEN,YELLOW,YELLOW,GREEN,GREEN],
      "Laurenssen 2018":[GREEN,GREEN,GREEN,YELLOW,GREEN,YELLOW],
      "Beck 2020":[GREEN,GREEN,YELLOW,GREEN,YELLOW,GREEN],
    }
    x0=CX_L+2600000; colw=1500000; d=300000; rh=560000
    # domain labels
    for i,dom in enumerate(domains):
        s.append(textbox(CX_L, y+ (i+1)*rh-40000, 2500000, rh, [para(dom, sz=1050, align="r", space_after=0)], anchor="ctr"))
    # study headers
    for j,st in enumerate(studies):
        s.append(textbox(x0+j*colw-200000, y-20000, colw, rh, [para(st, sz=1050, bold=True, align="ctr", space_after=0)], anchor="ctr"))
    for i,dom in enumerate(domains):
        for j,st in enumerate(studies):
            c=pattern[st][i]
            s.append(tl_dot(x0+j*colw+colw/2-d/2-200000, y+(i+1)*rh+rh/2-d/2-40000, d, c))
    # legend
    ly=y+(len(domains)+1)*rh+40000
    s.append(tl_dot(CX_L, ly, 260000, GREEN)); s.append(textbox(CX_L+320000, ly-30000, 1400000, 320000,[para("Low risk", sz=1100, space_after=0)],anchor="ctr"))
    s.append(tl_dot(CX_L+1800000, ly, 260000, YELLOW)); s.append(textbox(CX_L+2120000, ly-30000, 1600000, 320000,[para("Some concerns", sz=1100, space_after=0)],anchor="ctr"))
    s.append(tl_dot(CX_L+3800000, ly, 260000, RED)); s.append(textbox(CX_L+4120000, ly-30000, 1600000, 320000,[para("High risk", sz=1100, space_after=0)],anchor="ctr"))
    s.append(textbox(CX_L+6000000, ly-30000, CONTENT_W-6000000, 320000, [para("Illustrative pattern \u2014 replace with the paper's actual RoB table.", sz=1000, italic=True, color="9AA0A6", space_after=0)], anchor="ctr"))
    s += footer(23)
    return s
add(slide_rob(),
    "Explain the risk-of-bias appraisal. Domains follow the Cochrane approach: randomisation, allocation "
    "concealment, blinding of outcome assessment, incomplete outcome data, selective reporting, and treatment "
    "fidelity. The recurring weak spots in psychotherapy trials are blinding (you cannot blind participants to "
    "talking therapy) and attrition. This traffic-light grid is illustrative \u2014 replace each dot with the paper's "
    "actual rating. Use it to argue that even a strong pooled effect carries moderate risk of bias.")

print("slides through 23:", len(SLIDES))

# ---------- SLIDE 24 : STATISTICAL ANALYSIS ----------
def slide_stats():
    s, ytop = banner("Statistical Analysis (in Plain Language)", "HOW EFFECTS WERE POOLED")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-980000
    lp=[para("Key concepts", sz=1450, color=ACCENT, bold=True, space_after=260)]
    for t in ["Effect size = Hedges' g (a standardized mean difference); negative g favours MBT",
              "Random-effects model \u2014 assumes true effects vary across trials",
              "95% confidence interval shows the plausible range of the true effect",
              "I\u00B2 quantifies heterogeneity (how much studies disagree)",
              "Sensitivity / subgroup analyses & publication-bias checks"]:
        lp.append(para(t, sz=1200, bullet=True, space_after=200))
    s.append(roundrect(CX_L, y, cw, ch, LGREY, lp, anchor="t"))
    rp=[para("Interpreting Hedges' g", sz=1450, color=ACCENT, bold=True, space_after=260),
        multi_run_para([("\u2248 0.2 small   \u2022   \u2248 0.5 medium   \u2022   \u2248 0.8 large",{"sz":1300,"bold":True})], space_after=220),
        para("A pooled estimate summarises the average treatment effect across studies; high I\u00B2 warns that this average may hide real differences between trials.", sz=1200, space_after=0)]
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, CREAM, rp, anchor="t", line=LILAC))
    s.append(roundrect(CX_L, y+ch+120000, CONTENT_W, 620000, LAVENDER, [
        multi_run_para([("Reminder:  ",{"bold":True,"color":ACCENT,"sz":1300}),
                        ("statistical significance (does the CI exclude 0?) is not the same as clinical meaningfulness \u2014 always report both.",{"sz":1250})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(24)
    return s
add(slide_stats(),
    "Demystify the statistics for a clinical audience. The effect size is Hedges' g, a standardized mean difference "
    "that lets us combine trials using different scales; a negative g means MBT reduced the outcome. A random-effects "
    "model is used because the true effect probably varies across these heterogeneous trials. The confidence interval "
    "gives the plausible range, and I\u00B2 tells us how much the trials disagree. Anchor g values: 0.2 small, 0.5 "
    "medium, 0.8 large. Close by stressing that significance and clinical importance are different questions.")

# ---------- SLIDE 25 : RESULTS - SELF-HARM (forest style) ----------
def forest_row(x, y, w, label, g, lo, hi, gmin=-2.0, gmax=1.0, color=ACCENT):
    """Draw a forest-plot style row. Scale g in [gmin,gmax] onto width w."""
    shapes=[]
    def sx(v): return x + int(w*(v-gmin)/(gmax-gmin))
    # label
    shapes.append(textbox(x-4200000, y-40000, 4100000, 420000, [para(label, sz=1150, align="r", space_after=0)], anchor="ctr"))
    # CI line
    x1=sx(lo); x2=sx(hi); cy=y+140000
    shapes.append(line_shape(x1, cy, x2-x1, 0, color, 28575))
    # point marker (square)
    d=180000
    shapes.append(rect(sx(g)-d//2, cy-d//2, d, d, color))
    # value text
    shapes.append(textbox(x+w+80000, y-40000, 2600000, 420000,
        [para(f"g = {g:.2f}  [{lo:.2f}, {hi:.2f}]", sz=1100, bold=True, space_after=0)], anchor="ctr"))
    return shapes, sx

def slide_res_selfharm():
    s, ytop = banner("Results \u2014 Effect on Self-Harm", "PRIMARY OUTCOME (VERIFIED)")
    y=ytop+240000
    px=CX_L+4300000; pw=3600000
    gmin,gmax=-2.0,1.0
    def sx(v): return px+int(pw*(v-gmin)/(gmax-gmin))
    # axis
    s.append(line_shape(px, y+40000, pw, 0, "9AA0A6", 12700))
    # zero line (no effect)
    zx=sx(0.0)
    s.append(line_shape(zx, y+40000, 0, 1900000, "B0B4B8", 12700))
    s.append(textbox(zx-500000, y+1950000, 1000000, 300000, [para("0 (no effect)", sz=950, color="9AA0A6", align="ctr", space_after=0)], anchor="ctr"))
    s.append(textbox(px-500000, y-260000, 1000000, 260000, [para("\u2190 favours MBT", sz=1000, color=ACCENT2, bold=True, align="ctr", space_after=0)], anchor="ctr"))
    rows=[("Self-harm (pooled, MBT & MBT-A)",-0.82,-1.15,-0.50,BLUSH),
          ("Reference: medium effect (g=-0.5)",-0.50,-0.50,-0.50,"CFD3D8"),
          ("Reference: large effect (g=-0.8)",-0.80,-0.80,-0.80,"CFD3D8")]
    ry=y+320000
    for lbl,g,lo,hi,c in rows[:1]:
        sh,_=forest_row(px, ry, pw, lbl, g, lo, hi, gmin, gmax, "C25E7A"); s+=sh
        ry+=520000
    # interpretation card
    s.append(roundrect(CX_L, y+2100000, CONTENT_W, H-(y+2100000)-560000, MINT, [
        multi_run_para([("Pooled effect on self-harm:  ",{"bold":True,"color":ACCENT,"sz":1400}),
                        ("Hedges' g = \u22120.82 (95% CI \u22121.15 to \u22120.50).",{"sz":1350,"bold":True})], space_after=180),
        multi_run_para([("Interpretation:  ",{"bold":True,"color":ACCENT,"sz":1300}),
                        ("A large reduction favouring MBT/MBT-A; the CI excludes 0, so the effect is statistically significant. Wide CImirrors small, heterogeneous trials \u2014 interpret with caution.",{"sz":1250})], space_after=0)], line=SAGE, anchor="ctr"))
    s += footer(25)
    return s
add(slide_res_selfharm(),
    "This is the headline result. Pooled across MBT and MBT-A trials, self-harm fell with a Hedges' g of \u22120.82, "
    "95% CI \u22121.15 to \u22120.50. The square is the point estimate and the line is the confidence interval; because it "
    "sits entirely left of zero and left of the g=\u22120.8 reference, this is a large, statistically significant "
    "reduction favouring MBT. Immediately temper it: the interval is wide, reflecting few, small, heterogeneous "
    "trials, so the true effect could plausibly be medium rather than large. These numbers are verified from the "
    "published abstract.")

print("slides through 25:", len(SLIDES))

# ---------- SLIDE 26 : RESULTS - SUICIDAL BEHAVIOUR ----------
def slide_res_suicide():
    s, ytop = banner("Results \u2014 Suicidal Behaviour", "DISTINGUISH FROM SELF-HARM")
    y=ytop+120000
    s.append(roundrect(CX_L, y, CONTENT_W, 900000, PEACH, [
        multi_run_para([("Definitional care:  ",{"bold":True,"color":"B5651D","sz":1350}),
                        ("self-harm \u2260 suicidal ideation \u2260 suicide attempt \u2260 suicide death. Many trials measured self-harm broadly rather than suicide-specific endpoints.",{"sz":1250})], space_after=0)], line="E0A96D", anchor="ctr"))
    s.append(roundrect(CX_L, y+1020000, CONTENT_W, H-(y+1020000)-560000, CREAM, [
        para("What the review can and cannot say", sz=1400, color=ACCENT, bold=True, space_after=260),
        para("The abstract reports pooled effects for self-harm, BPD symptoms and depression \u2014 not a separate pooled estimate for suicide attempts or deaths.", sz=1250, bullet=True, space_after=220),
        para("If the paper provides no pooled estimate for a suicide-specific outcome, state that explicitly: 'No pooled estimate was available for this outcome.'", sz=1250, bullet=True, space_after=220),
        para("Absence of evidence for a suicide-specific effect is NOT evidence of no effect \u2014 avoid over-claiming on suicide prevention.", sz=1250, bullet=True, space_after=0),
    ], line=LILAC, anchor="t"))
    s += footer(26)
    return s
add(slide_res_suicide(),
    "Handle suicidal behaviour carefully and honestly. Draw the ladder: self-harm, suicidal ideation, suicide "
    "attempt, and suicide death are distinct. The verified pooled results cover self-harm, BPD symptoms and "
    "depression \u2014 the abstract does not report a separate pooled effect for suicide attempts or deaths. If the full "
    "paper gives no pooled suicide-specific estimate, say so plainly and do not extrapolate a self-harm effect into "
    "a claim about suicide prevention. Remember: absence of evidence is not evidence of absence.")

# ---------- SLIDE 27 : RESULTS - BPD & DEPRESSION (bars with real g) ----------
def slide_res_related():
    s, ytop = banner("Results \u2014 BPD Symptoms & Depression", "SECONDARY OUTCOMES (VERIFIED)")
    y=ytop+280000
    # bars representing |g| on a 0..1.6 scale
    def gbar(x,y,maxw,h,g,lo,hi,label,color):
        sh=[]
        scale=1.6
        bw=int(maxw*abs(g)/scale)
        sh.append(textbox(x-3000000, y-30000, 2900000, h+60000, [para(label, sz=1200, align="r", space_after=0)], anchor="ctr"))
        sh.append(rect(x,y,maxw,h,WHITE,line="E0E4E8"))
        sh.append(roundrect(x,y,max(bw,60000),h,color,[],rad=6000,shadow=False))
        sh.append(textbox(x+max(bw,60000)+50000, y-30000, 3200000, h+60000,
            [para(f"g = {g:.2f}  [{lo:.2f}, {hi:.2f}]", sz=1150, bold=True, space_after=0)], anchor="ctr"))
        return sh
    bx=CX_L+3100000; bmax=CONTENT_W-3100000-3400000; bh=620000
    s+=gbar(bx, y, bmax, bh, -0.82, -1.15, -0.50, "Self-harm", "C25E7A")
    s+=gbar(bx, y+900000, bmax, bh, -1.08, -1.38, -0.77, "BPD symptoms", ACCENT)
    s+=gbar(bx, y+1800000, bmax, bh, -1.10, -1.52, -0.68, "Depression", ACCENT2)
    s.append(roundrect(CX_L, y+2700000, CONTENT_W, 760000, LAVENDER, [
        multi_run_para([("All three pooled effects are large and statistically significant (CIs exclude 0). ",{"sz":1250,"bold":True}),
                        ("Bar length = magnitude of |g|; longer = larger reduction favouring MBT.",{"sz":1200})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(27)
    return s
add(slide_res_related(),
    "Present the two secondary outcomes alongside self-harm for context. BPD symptoms improved with g = \u22121.08 "
    "(95% CI \u22121.38 to \u22120.77) and depression with g = \u22121.10 (\u22121.52 to \u22120.68). All three effects are large "
    "and statistically significant. The bars show magnitude only. Note that these very large effect sizes from small "
    "trials are somewhat implausibly large and may reflect small-study effects or optimistic early trials \u2014 a point "
    "to develop in the appraisal. Numbers are verified from the published abstract.")

# ---------- SLIDE 28 : RETENTION / ACCEPTABILITY ----------
def slide_retention():
    body=[
        para("Report dropout, completion and attendance for MBT vs control in each trial (insert values from the paper).", sz=1300, bullet=True, space_after=280),
        para("Feasibility trials (e.g., Griffiths et al., 2019) reported acceptable group attendance and safety.", sz=1300, bullet=True, space_after=280),
        para("Consider reasons for dropout and any adverse events, if reported.", sz=1300, bullet=True, space_after=280),
        para("Discuss whether MBT is feasible and acceptable for people with recurrent self-harm and emotional dysregulation.", sz=1300, bullet=True, space_after=280),
        para("Retention data are often incompletely reported \u2014 note this as a limitation.", sz=1300, bullet=True, space_after=0),
    ]
    return content_slide("Treatment Retention & Acceptability","FEASIBILITY", body, 28, cardfill=LGREY)
add(slide_retention(),
    "Address feasibility, which matters as much as efficacy in this hard-to-engage group. Report the dropout and "
    "completion rates for MBT versus control from the paper. Feasibility work such as Griffiths et al. (2019) found "
    "group MBT acceptable and safe with reasonable attendance. Discuss reasons for dropout and any adverse events. "
    "Conclude cautiously on acceptability and flag that retention and adverse-event reporting is often incomplete "
    "across the trials.")

# ---------- SLIDE 29 : GRAPHICAL RESULTS SUMMARY ----------
def slide_gfx_summary():
    s, ytop = banner("Results at a Glance", "VERIFIED POOLED EFFECTS")
    y=ytop+160000
    cards=[("Self-harm","g = \u22120.82","95% CI \u22121.15, \u22120.50",BLUSH),
           ("BPD symptoms","g = \u22121.08","95% CI \u22121.38, \u22120.77",LAVENDER),
           ("Depression","g = \u22121.10","95% CI \u22121.52, \u22120.68",POWDER)]
    cw=(CONTENT_W-2*300000)//3; ch=2000000
    cx=CX_L
    for h,g,ci,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [
            para(h, sz=1500, color=ACCENT, bold=True, align="ctr", space_after=260),
            para(g, sz=2600, color=CHARCOAL, bold=True, align="ctr", space_after=200),
            para(ci, sz=1200, align="ctr", space_after=160),
            para("large \u2022 significant", sz=1100, color=ACCENT2, bold=True, align="ctr", space_after=0)], anchor="ctr", line=LILAC))
        cx+=cw+300000
    s.append(roundrect(CX_L, y+ch+220000, CONTENT_W, H-(y+ch+220000)-560000, CREAM, [
        multi_run_para([("All effects favour MBT and are statistically significant, but derive from few, small trials with wide CIs and heterogeneous designs \u2014 ",{"sz":1250}),
                        ("treat as promising, not definitive.",{"sz":1250,"bold":True,"color":ACCENT})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(29)
    return s
add(slide_gfx_summary(),
    "A one-glance summary of the three verified pooled effects: self-harm g=\u22120.82, BPD symptoms g=\u22121.08, "
    "depression g=\u22121.10, all large and significant. Use this as the pivot from results to discussion: the signal is "
    "consistent and favourable, but it comes from a small, heterogeneous evidence base, so the honest headline is "
    "'promising, not definitive.'")

print("slides through 29:", len(SLIDES))

# ---------- SLIDE 30 : DISCUSSION - MAIN FINDINGS ----------
def slide_disc_main():
    body=[
        para("MBT / MBT-A was associated with large, statistically significant reductions in self-harm (g=\u22120.82).", sz=1300, bullet=True, space_after=260),
        para("Parallel large improvements in BPD symptoms (g=\u22121.08) and depression (g=\u22121.10).", sz=1300, bullet=True, space_after=260),
        para("Direction of effect was consistent across the pooled outcomes \u2014 all favour MBT.", sz=1300, bullet=True, space_after=260),
        para("However, effects rest on a small number of small trials with heterogeneous designs and controls.", sz=1300, bullet=True, space_after=260),
        para("Overall evidence quality is best described as promising but limited / moderate-to-low confidence.", sz=1300, bullet=True, space_after=0),
    ]
    return content_slide("Discussion \u2014 Main Findings","INTERPRETATION", body, 30, cardfill=MINT)
add(slide_disc_main(),
    "Summarise what was found: consistent, large, significant reductions in self-harm, BPD symptoms and depression, "
    "all favouring MBT. Then give the balanced verdict \u2014 the effects are encouraging and internally consistent, but "
    "they come from few small trials with mixed designs and controls, so overall confidence in the effect size is "
    "moderate at best. Avoid overselling.")

# ---------- SLIDE 31 : DISCUSSION - MECHANISMS ----------
def slide_disc_mech():
    s, ytop = banner("Discussion \u2014 How Might MBT Work?", "MECHANISMS (CAUTIOUS)")
    y=ytop+120000
    mechs=[("Improved reflective functioning","Better reading of own & others' mental states",LAVENDER),
           ("Reduced misinterpretation","Fewer catastrophic readings of interpersonal events",POWDER),
           ("Better affect regulation","Tolerating arousal without acting on it",MINT),
           ("Fewer impulsive responses","Pause between urge and action",PEACH),
           ("Stronger therapeutic alliance","Engagement supports change",BLUSH),
           ("Recognising triggers & alternatives","New coping instead of self-harm",SAGE)]
    cw=(CONTENT_W-2*260000)//3; ch=(H-y-620000-260000)//2
    for i,(h,t,c) in enumerate(mechs):
        r=i//3; col=i%3
        s.append(roundrect(CX_L+col*(cw+260000), y+r*(ch+260000), cw, ch, c,
            [para(h, sz=1300, color=ACCENT, bold=True, space_after=200), para(t, sz=1150, space_after=0)], anchor="t"))
    s += footer(31)
    return s
add(slide_disc_mech(),
    "Offer plausible mechanisms, framed tentatively. MBT is theorised to work by rebuilding reflective functioning "
    "so people read mental states more accurately, misinterpret interpersonal events less, regulate affect better, "
    "and insert a pause between urge and action; a strong alliance and better trigger-recognition support new "
    "coping. In Rossouw & Fonagy (2012) the effect was actually mediated by improved mentalizing and reduced "
    "attachment avoidance, which supports this model. Stress that a meta-analysis of outcomes cannot by itself prove "
    "mechanism \u2014 use 'may'.")

# ---------- SLIDE 32 : DISCUSSION - COMPARISON ----------
def slide_disc_compare():
    s, ytop = banner("Discussion \u2014 Comparison with Other Evidence", "CONTEXT")
    y=ytop+120000
    headers=["Therapy","Evidence for self-harm","Note"]
    cw=[2600000, 4600000, CONTENT_W-2600000-4600000]
    rows=[
        ["MBT / MBT-A","Large pooled effects here; early positive, later mixed trials","This review's focus"],
        ["DBT / DBT-A","Strongest evidence base; 'well-established' for adolescent self-harm","Common comparator/benchmark"],
        ["CBT-based","Moderate evidence for reducing repetition","Widely available"],
        ["TAU / SCM","Active controls shrink apparent MBT advantage","Comparator quality matters"],
    ]
    s+=table(CX_L, y, cw, 700000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1100, head_font=1200)
    s.append(textbox(CX_L, y+700000*5+40000, CONTENT_W, 420000, [
        para("Inconsistencies across MBT trials likely reflect differences in sample severity, treatment intensity, outcome measures and follow-up length.",
             sz=1050, italic=True, color="9AA0A6", space_after=0)]))
    s += footer(32)
    return s
add(slide_disc_compare(),
    "Place MBT among alternatives. DBT has the strongest, best-replicated evidence for adolescent self-harm and is "
    "the natural benchmark; CBT-based approaches show moderate benefit for repetition. MBT's pooled effects here are "
    "large but rest on fewer trials with mixed later results. Explain the inconsistencies by differences in sample "
    "severity, treatment intensity, outcome measurement and follow-up. The clinical takeaway: MBT is a reasonable "
    "option, but the comparative evidence does not yet establish superiority over DBT.")

# ---------- SLIDE 33 : CLINICAL MEANING ----------
def slide_clinical():
    body=[
        para("MBT may be considered for individuals with recurrent self-harm and BPD/emerging BPD features.", sz=1300, bullet=True, space_after=250),
        para("Treatment should be individualized; ongoing risk assessment remains essential.", sz=1300, bullet=True, space_after=250),
        para("MBT complements \u2014 does not replace \u2014 crisis planning and safety management.", sz=1300, bullet=True, space_after=250),
        para("Medication may be needed for comorbid conditions; engagement & continuity of care matter.", sz=1300, bullet=True, space_after=250),
        para("Therapists require appropriate training and supervision; match setting to risk level.", sz=1300, bullet=True, space_after=250),
        para("Involve family/carers where appropriate (especially in adolescent MBT-A).", sz=1300, bullet=True, space_after=0),
    ]
    return content_slide("Clinical Meaning","IMPLICATIONS FOR PRACTICE", body, 33, cardfill=POWDER)
add(slide_clinical(),
    "Translate findings into practice. MBT is a reasonable, evidence-supported option for recurrent self-harm with "
    "BPD features, but it must be individualized and sits within \u2014 never instead of \u2014 continuous risk assessment "
    "and crisis planning. Comorbidities may need medication; engagement and continuity are decisive in this "
    "population. Emphasise that MBT requires trained, supervised therapists and that setting should match risk. In "
    "adolescents, involving family (MBT-A) is important.")

print("slides through 33:", len(SLIDES))

# ---------- SLIDE 34 : LIMITATIONS (matrix) ----------
def slide_limits():
    s, ytop = banner("Limitations of the Review","CRITICAL APPRAISAL")
    y=ytop+120000
    headers=["Limitation","Possible impact"]
    cw=[5200000, CONTENT_W-5200000]
    rows=[
        ["Few, small studies","Reduced power; wide confidence intervals"],
        ["Heterogeneity (design, format, controls)","Lower confidence in the pooled estimate"],
        ["Short / variable follow-up","Long-term durability of effects unclear"],
        ["Different outcome measures","Harder to compare and combine"],
        ["Attrition & incomplete reporting","Risk of attrition / reporting bias"],
        ["Limited diversity (female, Western, young)","Restricted generalisability"],
        ["Self-harm vs suicide not always separated","Ambiguity in what was reduced"],
    ]
    s+=table(CX_L, y, cw, 560000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1150, head_font=1250)
    s += footer(34)
    return s
add(slide_limits(),
    "Be explicit and balanced about limitations \u2014 examiners reward this. The evidence base is few small trials, so "
    "power is low and CIs wide. Heterogeneity in design, MBT format and control conditions undermines a single "
    "pooled number. Follow-up is short, outcome measures differ, and attrition/reporting is inconsistent. Samples "
    "are predominantly female, young and Western, limiting generalisability, and self-harm and suicidal behaviour "
    "are not always clearly separated. These caveats are why the large effect sizes should be read as promising "
    "rather than conclusive.")

# ---------- SLIDE 35 : SAMPLE & DESIGN LIMITATIONS ----------
def slide_sample_design():
    return two_col("Sample & Design Limitations","APPRAISAL",
        "Sample",
        ["Predominantly female participants","Under-representation of males & gender-diverse people",
         "Limited older-adult representation","Largely Western / high-income settings",
         "Diagnostic overlap; varying baseline severity"],
        BLUSH,
        "Design",
        ["Incomplete randomization in some trials","Lack of assessor blinding",
         "Variable / non-specific TAU comparators","High or differential attrition",
         "Limited fidelity monitoring & preregistration"],
        POWDER, 35)
add(slide_sample_design(),
    "Separate the appraisal into sample and design. Sample-side: participants are mostly female, young and Western, "
    "with limited male, gender-diverse and older-adult representation and variable baseline severity \u2014 so who these "
    "results apply to is narrow. Design-side: some trials had weak randomisation, no assessor blinding, "
    "poorly specified TAU comparators, notable attrition, and little fidelity monitoring or preregistration. "
    "Together these lower our certainty and point to what better trials must fix.")

# ---------- SLIDE 36 : GENERALIZABILITY ----------
def slide_general():
    s, ytop = banner("Generalizability","TO WHOM DO RESULTS APPLY?")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-620000
    lp=[para("More applicable to", sz=1500, color=ACCENT, bold=True, space_after=300)]
    for t in ["Outpatient mental-health services","Adolescents & adults with self-harm + BPD features","High-income / Western settings","Female-predominant clinical samples"]:
        lp.append(para(t, sz=1300, bullet=True, space_after=240))
    rp=[para("Caution / limited evidence", sz=1500, color=ACCENT, bold=True, space_after=300)]
    for t in ["Emergency-department & inpatient settings","Low-resource / non-Western populations","Older adults; males & gender-diverse people","Self-harm without BPD; comorbid psychosis or severe cognitive impairment"]:
        rp.append(para(t, sz=1300, bullet=True, space_after=210))
    s.append(roundrect(CX_L, y, cw, ch, SAGE, lp, anchor="t"))
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, PEACH, rp, anchor="t"))
    s += footer(36)
    return s
add(slide_general(),
    "State the boundaries of generalisation. The findings map best onto outpatient services treating "
    "female-predominant adolescents and adults with self-harm and BPD features in Western settings. Be cautious "
    "applying them to emergency or inpatient contexts, low-resource or non-Western populations, older adults, males "
    "and gender-diverse people, or to self-harm without BPD and to comorbid psychosis or severe cognitive "
    "impairment \u2014 groups the trials barely represent.")

# ---------- SLIDE 37 : FUTURE APPLICATIONS ----------
def slide_future():
    s, ytop = banner("Future Applications & Implications","LOOKING FORWARD")
    y=ytop+120000
    cw=(CONTENT_W-3*220000)//4; ch=H-y-620000
    cards=[("Clinical practice","Integrate mentalizing-informed principles; assess mentalizing in crises; collaborative formulation; individualized safety plans",LAVENDER),
           ("Services","Structured MBT programmes; clinician training & supervision; stepped care; continuity after discharge",POWDER),
           ("Relapse prevention","Identify early warning signs; build coping alternatives; plan for interpersonal crises; monitor recurrence",MINT),
           ("Psychosocial","Involve family/carers; address trauma & attachment; support social functioning; multidisciplinary care",PEACH)]
    cx=CX_L
    for h,t,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [para(h, sz=1400, color=ACCENT, bold=True, space_after=240), para(t, sz=1150, space_after=0)], anchor="t"))
        cx+=cw+220000
    s += footer(37)
    return s
add(slide_future(),
    "Lay out implications across four levels. Practice: embed mentalizing-informed principles, assess mentalizing "
    "during crises, use collaborative formulation and individualized safety plans. Services: set up structured MBT "
    "programmes with training, supervision, stepped care and post-discharge continuity. Relapse prevention: identify "
    "early warning signs, rehearse coping alternatives, plan for interpersonal triggers and monitor recurrence. "
    "Psychosocial: involve family/carers, address trauma and attachment, and coordinate multidisciplinary care.")

# ---------- SLIDE 38 : RESEARCHER REFLECTION ----------
def slide_reflection():
    s, ytop = banner("What Could I Have Done as a Researcher?","REFLECTIVE CRITICAL APPRAISAL")
    y=ytop+120000
    items=["Recruit a larger, more diverse sample (age, sex, culture)",
           "Use longer, standardized follow-up",
           "Clearly separate suicidal from non-suicidal self-harm",
           "Employ rigorous randomization & assessor blinding",
           "Include active psychotherapy comparators (e.g., DBT)",
           "Monitor treatment fidelity; report therapist training",
           "Examine mechanisms of change; run subgroup analyses",
           "Preregister the protocol; report adverse events & dropout fully",
           "Assess cost-effectiveness & post-discharge outcomes"]
    cw=(CONTENT_W-260000)//2; ch=H-y-620000
    half=(len(items)+1)//2
    lp=[para(t, sz=1250, bullet=True, space_after=250) for t in items[:half]]
    rp=[para(t, sz=1250, bullet=True, space_after=250) for t in items[half:]]
    s.append(roundrect(CX_L, y, cw, ch, CREAM, lp, anchor="t", line=LILAC))
    s.append(roundrect(CX_L+cw+260000, y, cw, ch, CREAM, rp, anchor="t", line=LILAC))
    s += footer(38)
    return s
add(slide_reflection(),
    "Show independent critical thinking. If you had run this programme, you would recruit larger, more diverse "
    "samples with longer standardized follow-up, cleanly separate suicidal from non-suicidal self-harm, use rigorous "
    "randomisation and blinded outcome assessment, and include an active comparator like DBT rather than only TAU. "
    "You would monitor fidelity, report therapist training, test mechanisms and subgroups, preregister the protocol, "
    "and report adverse events, dropout and cost-effectiveness fully. This directly answers the review's own "
    "limitations.")

print("slides through 38:", len(SLIDES))

# ---------- SLIDE 39 : KEY TAKEAWAYS ----------
def slide_takeaways():
    s, ytop = banner("Key Takeaways","HIGH-YIELD POINTS")
    y=ytop+120000
    pts=["Self-harm is recurrent and clinically complex \u2014 comprehensive risk assessment is essential.",
         "MBT targets difficulties in understanding self and others during emotional arousal.",
         "Pooled evidence shows large, significant reductions: self-harm g=\u22120.82, BPD g=\u22121.08, depression g=\u22121.10.",
         "Small samples, heterogeneity and short follow-up limit confidence \u2014 promising, not definitive.",
         "MBT is especially relevant for interpersonal sensitivity, emotional dysregulation & BPD features.",
         "More rigorous, diverse, long-term trials (ideally vs active comparators) are needed."]
    ch=(H-y-620000-5*140000)//6
    colors=[BLUSH,PEACH,SAGE,POWDER,LAVENDER,MINT]
    sy=y
    for i,(p,c) in enumerate(zip(pts,colors)):
        s.append(roundrect(CX_L, sy, CONTENT_W, ch, c, [
            multi_run_para([(f"{i+1}.  ",{"bold":True,"color":ACCENT,"sz":1300}),(p,{"sz":1250})], space_after=0)], anchor="ctr"))
        sy+=ch+140000
    s += footer(39)
    return s
add(slide_takeaways(),
    "Deliver the six high-yield messages slowly. Self-harm is recurrent and needs thorough risk assessment. MBT "
    "targets mentalizing failure under arousal. The verified pooled effects are large and significant across "
    "self-harm, BPD symptoms and depression. But small, heterogeneous, short-follow-up trials mean the honest "
    "verdict is 'promising, not definitive.' MBT fits patients with interpersonal sensitivity and BPD features "
    "best. And the field needs bigger, more diverse, longer trials against active comparators. End on this balanced "
    "note.")

# ---------- SLIDE 40 : CRITICAL APPRAISAL SUMMARY ----------
def slide_appraisal():
    s, ytop = banner("Critical Appraisal Summary","OVERALL JUDGEMENT (template)")
    y=ytop+120000
    headers=["Appraisal domain","Judgement"]
    cw=[6800000, CONTENT_W-6800000]
    rows=[
        ["Clear research question","Yes"],
        ["Appropriate design (SR & meta-analysis)","Yes"],
        ["Comprehensive search","Confirm from paper"],
        ["Appropriate inclusion criteria","Partly / Yes"],
        ["Risk-of-bias assessment conducted","Confirm from paper"],
        ["Appropriate statistical analysis (random effects, Hedges' g)","Yes"],
        ["Heterogeneity addressed","Partly"],
        ["Clinical relevance","Moderate\u2013High"],
        ["Applicability / generalisability","Limited"],
        ["Overall confidence in conclusions","Moderate\u2013Low"],
    ]
    s+=table(CX_L, y, cw, 430000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1150, head_font=1250)
    s += footer(40)
    return s
add(slide_appraisal(),
    "Give a structured verdict, like a CASP checklist. The question and design are appropriate and the statistics "
    "(random effects, Hedges' g) are suitable. Confirm from the full paper whether the search was comprehensive and "
    "whether risk of bias was formally assessed. Heterogeneity is only partly addressed. Clinical relevance is "
    "moderate-to-high, but applicability is limited by the narrow samples, so overall confidence in the conclusions "
    "is moderate-to-low. Adjust each judgement once you have read the full methods.")

# ---------- SLIDE 41 & 42 : REFERENCES ----------
def slide_refs(part, refs, page):
    body=[]
    for r in refs:
        body.append(para(r, sz=1150, bullet=False, space_after=280))
    s, ytop = banner("References" + (" (cont.)" if part==2 else ""), "APA 7th EDITION \u2014 VERIFY BEFORE USE")
    s.append(textbox(CX_L, ytop+120000, CONTENT_W, H-ytop-560000, body, anchor="t"))
    s += footer(page)
    return s
refs1=[
 "Primary paper: (2024). Efficacy of mentalization-based therapy in treating self-harm: A systematic review and meta-analysis. Suicide and Life-Threatening Behavior. https://doi.org/10.1111/sltb.13044  [insert full author list from the paper]",
 "Rossouw, T. I., & Fonagy, P. (2012). Mentalization-based treatment for self-harm in adolescents: A randomized controlled trial. Journal of the American Academy of Child & Adolescent Psychiatry, 51(12), 1304\u20131313.",
 "Bateman, A., & Fonagy, P. (2009). Randomized controlled trial of outpatient mentalization-based treatment versus structured clinical management for borderline personality disorder. American Journal of Psychiatry, 166(12), 1355\u20131364.",
 "Bateman, A., & Fonagy, P. (2016). Mentalization-based treatment for personality disorders: A practical guide. Oxford University Press.",
 "Laurenssen, E. M. P., et al. (2018). Day hospital mentalization-based treatment for adolescents (MBT-A). [confirm citation details].",
]
refs2=[
 "Beck, E., et al. (2020). Mentalization-based treatment in groups for adolescents with borderline personality disorder: A randomized controlled trial. [confirm citation details].",
 "Griffiths, H., et al. (2019). Group mentalization-based treatment for adolescents (MBT-Ai): Feasibility RCT. [confirm citation details].",
 "Page, M. J., et al. (2021). The PRISMA 2020 statement: An updated guideline for reporting systematic reviews. BMJ, 372, n71.",
 "National Institute for Health and Care Excellence (NICE). Self-harm: assessment, management and preventing recurrence (NG225).",
 "World Health Organization. (2019). International Classification of Diseases (11th ed.).",
 "American Psychiatric Association. (2022). Diagnostic and statistical manual of mental disorders (5th ed., text rev.).",
 "Note: verify every citation against the paper's reference list before presenting; do not present unverified details as the authors' own.",
]
add(slide_refs(1, refs1, 41),
    "Present the key references. Lead with the primary paper (DOI 10.1111/sltb.13044) \u2014 insert the full author list "
    "from the article. Then the landmark constituent trials (Rossouw & Fonagy 2012; Bateman & Fonagy 2009) and the "
    "MBT manual. Tell the audience these are cross-checked but a few constituent-trial details are marked 'confirm' "
    "and must be verified against the paper's reference list before the talk.")
add(slide_refs(2, refs2, 42),
    "Continue with the remaining trials, PRISMA reporting guidance, NICE self-harm guidance, and the diagnostic "
    "systems (ICD-11, DSM-5-TR). Reiterate the accuracy rule: verify each citation against the paper before "
    "presenting, and never attribute unverified specifics to the authors.")

# ---------- SLIDE 43 : QUESTIONS ----------
def slide_questions():
    s=[]
    s.append(rect(0,0,W,H,ACCENT))
    s.append(oval(-600000,-600000,2400000,2400000,ACCENT2))
    s.append(oval(W-1800000,H-1800000,2400000,2400000,"5A4A8A"))
    s.append(textbox(1200000, 2400000, W-2400000, 1400000, [
        para("Thank you \u2014 Questions & Discussion", sz=4000, color=WHITE, bold=True, align="ctr", space_after=260),
        para("Efficacy of Mentalization-Based Therapy in Treating Self-Harm: A Systematic Review & Meta-Analysis",
             sz=1500, color=LAVENDER, italic=True, align="ctr", space_after=0)]))
    s.append(textbox(1200000, 4300000, W-2400000, 600000, [
        para("Sonal  \u2022  M.Phil. Clinical Psychology  \u2022  DOI: 10.1111/sltb.13044", sz=1300, color="D6C7EE", align="ctr", space_after=0)]))
    return s
add(slide_questions(),
    "Close and open the floor. Offer two or three prompts to seed discussion: (1) Given the large but uncertain "
    "effects, would you offer MBT or DBT first for adolescent self-harm with BPD features, and why? (2) How should "
    "we weigh a large pooled effect from small, heterogeneous trials in clinical decisions? (3) What would convince "
    "us MBT reduces suicide-specific outcomes, not just self-harm? Thank the audience.")

print("TOTAL SLIDES:", len(SLIDES))

# ============================================================
# OOXML PACKAGING
# ============================================================
NS_P='http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A='http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def slide_xml(shapes):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{shapes}
</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>'''

def notes_xml(notestext, idx):
    safe=esc(notestext)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{safe}</a:t></a:r></a:p>
</p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>'''

SLIDE_LAYOUT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>'''%(NS_A,NS_R,NS_P)

def color_scheme():
    return ('<a:clrScheme name="Custom"><a:dk1><a:srgbClr val="2E2E38"/></a:dk1>'
            '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="6E5AA6"/></a:dk2>'
            '<a:lt2><a:srgbClr val="FBF6EC"/></a:lt2>'
            '<a:accent1><a:srgbClr val="6E5AA6"/></a:accent1><a:accent2><a:srgbClr val="3E7CB1"/></a:accent2>'
            '<a:accent3><a:srgbClr val="D9E8D2"/></a:accent3><a:accent4><a:srgbClr val="FBE3D2"/></a:accent4>'
            '<a:accent5><a:srgbClr val="F7D9E0"/></a:accent5><a:accent6><a:srgbClr val="D7E8F4"/></a:accent6>'
            '<a:hlink><a:srgbClr val="3E7CB1"/></a:hlink><a:folHlink><a:srgbClr val="6E5AA6"/></a:folHlink></a:clrScheme>')

SLIDE_MASTER='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>'''%(NS_A,NS_R,NS_P)

THEME='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="%s" name="MBT Pastel">
<a:themeElements>%s
<a:fontScheme name="Calibri"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>'''%(NS_A, color_scheme())

print("packaging helpers ready")

# ============================================================
# ASSEMBLE THE ZIP
# ============================================================
def build():
    n = len(SLIDES)
    files = {}

    # [Content_Types].xml
    overrides = []
    overrides.append('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>')
    overrides.append('<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>')
    overrides.append('<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>')
    overrides.append('<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
    overrides.append('<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>')
    for i in range(1,n+1):
        overrides.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
        overrides.append(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>')
    overrides.append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    overrides.append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    files['[Content_Types].xml']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides) + '</Types>')

    # _rels/.rels
    files['_rels/.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>')

    # docProps
    today=datetime.date.today().isoformat()
    files['docProps/core.xml']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Efficacy of MBT in Treating Self-Harm: A Systematic Review and Meta-Analysis</dc:title>'
        '<dc:creator>Journal Club (Sonal)</dc:creator><cp:keywords>MBT; self-harm; meta-analysis</cp:keywords>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{today}T00:00:00Z</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{today}T00:00:00Z</dcterms:modified>'
        '</cp:coreProperties>')
    files['docProps/app.xml']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        f'<Application>Kiro OOXML Generator</Application><Slides>{n}</Slides></Properties>')

    # theme, master, layout, notesMaster
    files['ppt/theme/theme1.xml']=THEME
    files['ppt/slideMasters/slideMaster1.xml']=SLIDE_MASTER
    files['ppt/slideMasters/_rels/slideMaster1.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>')
    files['ppt/slideLayouts/slideLayout1.xml']=SLIDE_LAYOUT
    files['ppt/slideLayouts/_rels/slideLayout1.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>')
    # notes master (minimal)
    files['ppt/notesMasters/notesMaster1.xml']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:notesMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/></p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '</p:notesMaster>')
    files['ppt/notesMasters/_rels/notesMaster1.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>')

    # presentation.xml
    sldIds=""
    for i in range(1,n+1):
        sldIds+=f'<p:sldId id="{255+i}" r:id="rId{i}"/>'
    files['ppt/presentation.xml']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" saveSubsetFonts="1">'
        f'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{n+1}"/></p:sldMasterIdLst>'
        f'<p:notesMasterIdLst><p:notesMasterId r:id="rId{n+2}"/></p:notesMasterIdLst>'
        f'<p:sldIdLst>{sldIds}</p:sldIdLst>'
        f'<p:sldSz cx="{W}" cy="{H}"/><p:notesSz cx="{H}" cy="{W}"/></p:presentation>')

    # presentation rels
    prels=""
    for i in range(1,n+1):
        prels+=f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
    prels+=f'<Relationship Id="rId{n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    prels+=f'<Relationship Id="rId{n+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>'
    prels+=f'<Relationship Id="rId{n+3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
    files['ppt/_rels/presentation.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+prels+'</Relationships>')

    # slides + notes
    for i,(shapes,notes) in enumerate(SLIDES, start=1):
        files[f'ppt/slides/slide{i}.xml']=slide_xml(shapes)
        files[f'ppt/slides/_rels/slide{i}.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{i}.xml"/>'
            '</Relationships>')
        files[f'ppt/notesSlides/notesSlide{i}.xml']=notes_xml(notes, i)
        files[f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels']=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{i}.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
            '</Relationships>')

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        # content types first is conventional
        z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for name, content in files.items():
            z.writestr(name, content)
    print(f"WROTE {OUT} with {n} slides")

build()
