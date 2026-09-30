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
# SLIDE CONTENT  (body font = 14pt / 1400; simple language;
# hard terms explained in brackets; every graph has a plain caption)
# ============================================================
SLIDES = []
def add(shapes, notes):
    SLIDES.append(("".join(shapes), notes))

CX_L = 560000
CONTENT_W = W - 2*CX_L
BODY = 1400
SUB  = 1400

def caption(x, y, cx, text):
    return roundrect(x, y, cx, 680000, CREAM, [
        multi_run_para([("What this shows:  ",{"bold":True,"color":ACCENT2,"sz":1400}),
                        (text,{"sz":1400})], space_after=0)], line=LILAC, anchor="ctr")

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
    lp = [para(left_head, sz=1700, color=ACCENT, bold=True, space_after=360)]
    lp += [para(t, sz=BODY, bullet=True, space_after=280) for t in left_items]
    rp = [para(right_head, sz=1700, color=ACCENT, bold=True, space_after=360)]
    rp += [para(t, sz=BODY, bullet=True, space_after=280) for t in right_items]
    s.append(roundrect(CX_L, y, colw, ch, left_fill, lp, anchor="t"))
    s.append(roundrect(CX_L+colw+300000, y, colw, ch, right_fill, rp, anchor="t"))
    s += footer(page)
    return s

def table(x, y, col_w, row_h, headers, rows, head_fill=ACCENT, head_color=WHITE,
          zebra=(WHITE, LGREY), font=1400, head_font=1400):
    shapes=[]
    cx=x
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

def slide_title():
    s = []
    s.append(rect(0,0,W,H,CREAM))
    s.append(rect(0,0,W,120000,ACCENT))
    s.append(rect(0,H-120000,W,120000,ACCENT))
    s.append(oval(-500000,-500000,2200000,2200000,LAVENDER))
    s.append(oval(W-1600000,H-1600000,2200000,2200000,POWDER))
    s.append(textbox(900000, 850000, W-1800000, 500000,
        [para("M.Phil. CLINICAL PSYCHOLOGY  \u2022  JOURNAL CLUB", sz=1500, color=ACCENT2, bold=True, align="ctr", space_after=0)]))
    s.append(roundrect(900000, 1480000, W-1800000, 1650000, WHITE,
        [para("Efficacy of Mentalization-Based Therapy in Treating Self-Harm",
              sz=3200, color=ACCENT, bold=True, align="ctr", space_after=200),
         para("A Systematic Review and Meta-Analysis", sz=2200, color=CHARCOAL, italic=True, align="ctr", space_after=0)],
        line=LILAC, anchor="ctr"))
    s.append(textbox(900000, 3180000, W-1800000, 640000, [
        multi_run_para([("Journal: ",{"bold":True,"color":ACCENT2,"sz":1300}),
                        ("Suicide and Life-Threatening Behavior (2024)  \u2022  DOI: 10.1111/sltb.13044",{"sz":1300}),
                        ("      What kind of study: ",{"bold":True,"color":ACCENT2,"sz":1300}),
                        ("systematic review + meta-analysis (combines earlier studies into one overall answer).",{"sz":1300})], align="ctr", space_after=0),
    ]))
    cardw = (W-1800000-300000)//2
    cardx = 900000
    cardy = 3900000
    cardh = 2350000
    s.append(roundrect(cardx, cardy, cardw, cardh, LAVENDER, [
        para("Presented by", sz=1500, color=ACCENT, bold=True, align="ctr", space_after=200),
        para("Ms. Sonal Tripathi", sz=1400, bold=True, align="ctr", space_after=140),
        para("M.Phil. Year II Trainee", sz=1400, align="ctr", space_after=140),
        para("Department of Clinical Psychology", sz=1400, align="ctr", space_after=140),
        para("MAN College of Special Education and Psychological Studies, Guna, M.P.", sz=1400, align="ctr", space_after=0),
    ], line=LILAC, anchor="ctr"))
    s.append(roundrect(cardx+cardw+300000, cardy, cardw, cardh, MINT, [
        para("Supervised by", sz=1500, color=ACCENT, bold=True, align="ctr", space_after=200),
        para("Dr. Ajay Sharma", sz=1400, bold=True, align="ctr", space_after=140),
        para("Professor & Head of Department", sz=1400, align="ctr", space_after=140),
        para("Department of Clinical Psychology", sz=1400, align="ctr", space_after=140),
        para("MAN College of Special Education and Psychological Studies, Guna, M.P.", sz=1400, align="ctr", space_after=0),
    ], line=SAGE, anchor="ctr"))
    s.append(textbox(900000, cardy+cardh+120000, W-1800000, 320000, [
        para("Journal Club Date: ______________", sz=1300, color="6B6B75", align="ctr", space_after=0)]))
    notes = ("Open simply: 'Today I am presenting a 2024 paper that asks one clear question - does mentalization-based "
             "therapy (a talking therapy that helps people understand thoughts and feelings) actually reduce self-harm?' "
             "Explain it is a systematic review and meta-analysis: the authors did not run a new experiment; they "
             "collected earlier good-quality studies and combined them to get one overall result. Introduce yourself "
             "(Ms. Sonal Tripathi, M.Phil. Year II) and acknowledge your supervisor (Dr. Ajay Sharma, Professor & "
             "HoD). Give the plan: the problem, why this therapy, how the review was done, what it found, and how "
             "strong the evidence really is.")
    return s, notes
s,n = slide_title(); add(s,n)

def slide_why():
    s, ytop = banner("Why This Paper Matters", "CLINICAL & ACADEMIC RELEVANCE")
    y = ytop+120000
    left_w = 5600000
    lp = [
        para("Self-harm is common and tends to happen again and again (high recurrence).", sz=BODY, bullet=True, space_after=300),
        para("It is closely linked to poor emotion control (emotional dysregulation), past trauma, personality difficulties, and a higher risk of suicide.", sz=BODY, bullet=True, space_after=300),
        para("Because it repeats so often, we badly need talking therapies that actually work.", sz=BODY, bullet=True, space_after=300),
        para("Therapies that fix the root cause (how a person handles emotions and relationships) are more useful than ones that only stop the behaviour briefly.", sz=BODY, bullet=True, space_after=300),
        para("A meta-analysis (combining many studies) gives one clearer, stronger answer than any single small study.", sz=BODY, bullet=True, space_after=0),
    ]
    s.append(roundrect(CX_L, y, left_w, H-y-720000, CREAM, lp, anchor="t", line=LILAC))
    rx = CX_L + left_w + 300000
    rw = CONTENT_W - left_w - 300000
    steps = [("Self-harm", BLUSH), ("It repeats", PEACH), ("Suicide risk rises", LILAC),
             ("Treatment needed", POWDER), ("Use proven therapy", SAGE)]
    sy = y+20000
    bh = 560000; gap = 170000
    for i,(t,c) in enumerate(steps):
        s.append(roundrect(rx, sy, rw, bh, c, [para(t, sz=BODY, bold=True, align="ctr", space_after=0)]))
        if i < len(steps)-1:
            s.append(arrow_down(rx+rw/2-120000, sy+bh+8000, 240000, gap-16000, ACCENT))
        sy += bh+gap
    s.append(caption(rx, sy+10000, rw, "The chain: self-harm usually repeats, repeating raises suicide risk, so treatment is needed - ideally a therapy already proven to work."))
    s += footer(2)
    return s
add(slide_why(),
    "Explain why the topic matters in everyday words. Self-harm is common and keeps coming back. It travels with "
    "trouble managing emotions, past trauma, personality difficulties, and a higher chance of suicide later. Because "
    "it repeats, we need therapies that genuinely help - ideally ones that treat the underlying cause. Then explain "
    "why a meta-analysis is valuable: one small study can be a fluke, but combining many gives a steadier answer. "
    "Walk down the arrow-chain on the right.")

def slide_problem():
    s, ytop = banner("The Clinical Problem: What Is Self-Harm?", "BACKGROUND")
    y = ytop+120000
    colw = (CONTENT_W-300000)//2
    ch = H-y-1080000
    lp = [para("Definition", sz=1600, color=ACCENT, bold=True, space_after=320),
          para("Deliberately hurting or poisoning oneself, whether or not the person wants to die.", sz=BODY, bullet=True, space_after=300),
          para("It overlaps with, but is not the same as, suicidal behaviour - non-suicidal self-injury (hurting self with no wish to die) and suicidal self-injury (with a wish to die) can both occur.", sz=BODY, bullet=True, space_after=300),
          para("It serves a purpose for the person: to feel relief, to cope, to communicate distress, or to punish oneself.", sz=BODY, bullet=True, space_after=0)]
    rp = [para("Common forms", sz=1600, color=ACCENT, bold=True, space_after=320)]
    for f in ["Cutting","Burning","Hitting oneself","Scratching","Self-poisoning (taking an overdose)","Interfering with wound healing (stopping wounds from healing)"]:
        rp.append(para(f, sz=BODY, bullet=True, space_after=250))
    s.append(roundrect(CX_L, y, colw, ch, POWDER, lp, anchor="t"))
    s.append(roundrect(CX_L+colw+300000, y, colw, ch, MINT, rp, anchor="t"))
    s.append(roundrect(CX_L, y+ch+120000, CONTENT_W, 720000, PEACH, [
        multi_run_para([("\u26A0  Always assess safety:  ",{"bold":True,"color":"B5651D","sz":BODY}),
                        ("check for suicidal intent (wish to die), lethality (how dangerous), how often it happens, what triggers it, what purpose it serves, and whether the person is safe right now.",
                         {"sz":BODY})], space_after=0)], line="E0A96D", anchor="ctr"))
    s += footer(3)
    return s
add(slide_problem(),
    "Define self-harm plainly: deliberately hurting or poisoning yourself, regardless of intent. Make the key "
    "distinction: self-harm and suicide overlap but are not identical; a person can hurt themselves with no wish to "
    "die, or with a wish to die. Stress that self-harm serves a function - relief, coping, communicating pain, "
    "self-punishment - and those functions are what this therapy addresses. Read the orange safety box aloud.")

print("content slides 1-3 done:", len(SLIDES))

def slide_cycle():
    s, ytop = banner("Why Self-Harm Repeats: The Cycle", "RECURRENCE")
    y = ytop+120000
    lw = 4400000
    lp = [para("Common triggers (what sets it off)", sz=SUB, color=ACCENT, bold=True, space_after=240)]
    for t in ["Conflict or rejection by others","Fear of being abandoned","Shame and feeling overwhelmed","Dissociation (feeling cut off / unreal) or trauma reminders"]:
        lp.append(para(t, sz=BODY, bullet=True, space_after=200))
    lp.append(para("What keeps it going (maintaining factors)", sz=SUB, color=ACCENT, bold=True, space_after=240))
    for t in ["It brings quick emotional relief","Hard to name what one is feeling","Acting on impulse (little pause before acting)","Others respond, which unintentionally reinforces it"]:
        lp.append(para(t, sz=BODY, bullet=True, space_after=200))
    s.append(roundrect(CX_L, y, lw, H-y-720000, CREAM, lp, anchor="t", line=LILAC))
    rx = CX_L+lw+300000
    rw = CONTENT_W-lw-300000
    steps=[("Trigger",BLUSH),("Emotions spike",PEACH),("Can't think clearly\n(mentalizing drops)",LILAC),
           ("Urge \u2192 self-harm",POWDER),("Quick relief",MINT),("Shame \u2192 repeat",BLUSH)]
    sy=y+10000; bh=470000; gap=120000
    for i,(t,c) in enumerate(steps):
        lines=t.split("\n")
        ps=[para(lines[0], sz=1300, bold=True, align="ctr", space_after=(40 if len(lines)>1 else 0))]
        for extra in lines[1:]:
            ps.append(para(extra, sz=1100, bold=True, align="ctr", space_after=0))
        s.append(roundrect(rx, sy, rw, bh, c, ps))
        if i<len(steps)-1:
            s.append(arrow_down(rx+rw/2-100000, sy+bh+2000, 200000, gap-8000, ACCENT))
        sy+=bh+gap
    s.append(caption(rx-200000, sy+10000, rw+400000, "The loop: a trigger raises emotion, thinking clearly gets harder, self-harm gives quick relief, then shame restarts the loop."))
    s += footer(4)
    return s
add(slide_cycle(),
    "Explain recurrence as a vicious cycle. A trigger (often a relationship problem) raises emotion; when emotion is "
    "high, thinking clearly about feelings drops; an urge appears; self-harm brings quick relief - which is why the "
    "brain wants to repeat it; then shame becomes the next trigger. Point to the maintaining factors on the left: "
    "relief acts as a reward, so the loop keeps running. That is why we need a therapy that works at the 'can't "
    "think clearly' step.")

def slide_gap():
    body = [
        para("Getting the right help", sz=SUB, color=ACCENT, bold=True, space_after=320),
        para("Specialist talking therapies are hard to access; people respond differently, and many drop out before finishing.", sz=BODY, bullet=True, space_after=300),
        para("People who struggle most with emotions are often the hardest to keep in treatment.", sz=BODY, bullet=True, space_after=300),
        para("What the research is missing", sz=SUB, color=ACCENT, bold=True, space_after=320),
        para("Studies disagree with each other, use different treatment set-ups, and rarely follow people for a long time.", sz=BODY, bullet=True, space_after=300),
        para("There was no single combined answer on self-harm, suicidal behaviour, borderline personality symptoms, and day-to-day functioning - which is the gap this paper fills.", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("The Gap This Study Fills", "RATIONALE", body, 5, cardfill=LGREY)
add(slide_gap(),
    "Explain the problem the review solves. In real life, specialist therapies for self-harm are scarce, people "
    "respond unevenly, and dropout is high - especially in those with the biggest emotional struggles. On the "
    "research side, individual studies contradict each other, use different versions of the therapy, and rarely "
    "follow people long enough. Nobody had combined them into one clear answer - which is what this meta-analysis "
    "does.")

def slide_whymbt():
    s, ytop = banner("Why Mentalization-Based Therapy (MBT)?", "THE MAIN IDEA")
    y = ytop+120000
    s.append(roundrect(CX_L, y, CONTENT_W, 800000, LAVENDER, [
        multi_run_para([("Mentalizing = ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("the ability to understand your own and other people's behaviour by thinking about the feelings, thoughts and intentions behind it ('reading minds - your own and others').",{"sz":BODY})], space_after=0)], line=LILAC, anchor="ctr"))
    my = y+920000
    steps=[("Relationship\nstress",BLUSH),("Can't 'read'\nfeelings",PEACH),("Misreads self\n& others",LILAC),
           ("Emotions get\nout of control",POWDER),("Self-harm",BLUSH)]
    bw=2000000; gap=320000; bh=820000
    total = len(steps)*bw + (len(steps)-1)*gap
    sx = (W-total)//2
    for i,(t,c) in enumerate(steps):
        lines=t.split("\n")
        ps=[para(lines[0], sz=1300, bold=True, align="ctr", space_after=40)]
        for extra in lines[1:]:
            ps.append(para(extra, sz=1300, bold=True, align="ctr", space_after=0))
        s.append(roundrect(sx, my, bw, bh, c, ps))
        if i<len(steps)-1:
            s.append(arrow_right(sx+bw+40000, my+bh/2-90000, gap-80000, 180000, ACCENT))
        sx+=bw+gap
    s.append(caption(CX_L, my+bh+140000, CONTENT_W, "Under relationship stress, the ability to 'read' feelings drops, the person misjudges what's happening, emotions overflow, and self-harm follows. MBT works on this weak link."))
    s.append(roundrect(CX_L, my+bh+800000, CONTENT_W, H-(my+bh+800000)-560000, MINT, [
        multi_run_para([("MBT helps a person: ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("understand feelings better, become more aware of emotions, be less sure their negative reading of a situation is correct, calm strong emotions, get on better with others, and act less on impulse.",{"sz":BODY})], space_after=0)],
        line=SAGE, anchor="ctr"))
    s += footer(6)
    return s
add(slide_whymbt(),
    "Give the core idea plainly. 'Mentalizing' means reading the feelings and intentions behind behaviour - in "
    "yourself and others. The horizontal chain shows the theory: relationship stress makes this mind-reading break "
    "down, so the person misjudges what's happening, emotions overflow, and self-harm follows. MBT targets that "
    "breakdown. List what MBT builds: better understanding of feelings, more emotional awareness, less certainty "
    "that 'they hate me' is true, calmer reactions, better relationships, fewer impulsive acts. Give a quick "
    "everyday example if time allows.")

def slide_intro():
    s, ytop = banner("Introduction: Three Key Terms", "BACKGROUND")
    y = ytop+120000
    cw=(CONTENT_W-2*260000)//3
    ch=H-y-620000
    c1=[para("Self-harm", sz=SUB, color=ACCENT, bold=True, space_after=260),
        para("Deliberately hurting or poisoning oneself. Comes in many forms, serves a purpose, and is a strong warning sign for later suicide.", sz=BODY, bullet=True, space_after=0)]
    c2=[para("Mentalizing", sz=SUB, color=ACCENT, bold=True, space_after=260),
        para("Understanding behaviour through the feelings behind it. It grows in safe relationships and breaks down when emotions run high.", sz=BODY, bullet=True, space_after=0)]
    c3=[para("MBT", sz=SUB, color=ACCENT, bold=True, space_after=260),
        para("A structured talking therapy (Bateman & Fonagy). The therapist stays curious and 'not-knowing' (asks rather than assumes) and focuses on present feelings and relationships.", sz=BODY, bullet=True, space_after=0)]
    s.append(roundrect(CX_L, y, cw, ch, POWDER, c1, anchor="t"))
    s.append(roundrect(CX_L+cw+260000, y, cw, ch, LAVENDER, c2, anchor="t"))
    s.append(roundrect(CX_L+2*(cw+260000), y, cw, ch, MINT, c3, anchor="t"))
    s += footer(7)
    return s
add(slide_intro(),
    "Give three simple definitions. Self-harm: deliberately hurting yourself; a strong warning sign for suicide. "
    "Mentalizing: understanding behaviour by the feelings behind it; develops in safe relationships and switches "
    "off under stress. MBT: a structured talking therapy where the therapist stays curious and avoids assuming "
    "('not-knowing stance'), focusing on present feelings and relationships. Also remind them why we combine "
    "studies: a more reliable answer than any single small study.")

def slide_framework():
    s, ytop = banner("The Whole Idea in One Picture", "CONCEPTUAL FRAMEWORK")
    y = ytop+60000
    steps=[("Insecure relationships / relationship stress",LILAC),
           ("Ability to 'read' feelings drops when emotions rise",PEACH),
           ("Person misreads self and others",BLUSH),
           ("Emotions get out of control; acts on impulse",POWDER),
           ("Self-harm \u2192 quick relief",MINT),
           ("MBT strengthens mentalizing, emotion control & understanding others",SAGE),
           ("Less self-harm & better day-to-day functioning",LAVENDER)]
    bh=470000; gap=80000
    sx=CX_L+1500000; bw=CONTENT_W-3000000
    sy=y
    for i,(t,c) in enumerate(steps):
        bold = i>=5
        s.append(roundrect(sx, sy, bw, bh, c, [para(t, sz=1300, bold=bold, align="ctr", space_after=0)], line=(ACCENT if bold else None)))
        if i<len(steps)-1:
            s.append(arrow_down(sx+bw/2-90000, sy+bh+2000, 180000, gap-8000, ACCENT))
        sy+=bh+gap
    s.append(caption(CX_L, sy+6000, CONTENT_W, "Read top to bottom: stress breaks down 'mind-reading', leading to self-harm. The two highlighted boxes show where MBT steps in to break the chain."))
    s += footer(8)
    return s
add(slide_framework(),
    "This single picture ties the talk together - read top to bottom. Relationship stress makes reading feelings "
    "harder; the person misjudges; emotions overflow and they act on impulse; self-harm brings quick relief. The "
    "last two highlighted boxes are the hopeful part: MBT strengthens mind-reading and emotion-control, which should "
    "break the chain and reduce self-harm while improving daily life. Keep this in mind for the results.")

def slide_litreview():
    s, ytop = banner("What Earlier Studies Found", "REVIEW OF LITERATURE")
    y = ytop+100000
    headers=["Study (year)","Who / how many","Compared","Main finding (plain)"]
    cw=[2500000, 2350000, 2650000, CONTENT_W-2500000-2350000-2650000]
    rows=[
        ["Rossouw & Fonagy (2012)","UK; 80 teenagers","MBT-A vs usual care","MBT helped more - less self-harm & depression; worked by improving mind-reading"],
        ["Bateman & Fonagy (2009)","UK; adults with BPD","MBT vs generic support","Fewer suicide attempts & self-harm; symptoms improved"],
        ["Laurenssen et al. (2018)","Netherlands; teenagers","MBT-A vs usual care","No clear extra benefit - mixed result"],
        ["Beck et al. (2020)","Denmark; teenagers, BPD","MBT-A vs usual care","No significant difference in self-harm"],
        ["Griffiths et al. (2019)","UK; group MBT","Small pilot study","Doable & acceptable, but too small to prove it works"],
    ]
    s += table(CX_L, y, cw, 600000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1200, head_font=1250)
    s.append(caption(CX_L, y+600000*6+20000, CONTENT_W, "Early studies were positive, but later ones were mixed or found no clear benefit - which is why one combined answer (this meta-analysis) is needed. (Confirm details against the paper.)"))
    s += footer(9)
    return s
add(slide_litreview(),
    "Walk through the earlier studies plainly. The landmark is Rossouw & Fonagy (2012): 80 teenagers; MBT beat usual "
    "care on self-harm and depression, working specifically by improving mind-reading. Bateman & Fonagy's adult work "
    "is the foundation. Be honest: later studies (Laurenssen 2018, Beck 2020) found no clear extra benefit, and some "
    "were only small pilots. This disagreement is the key reason we need a combined answer, and it explains why the "
    "final result must be read carefully. Say details should be checked against the paper.")

print("content slides through 9:", len(SLIDES))

def hbar(x, y, max_w, h, val, maxval, fill, label, valtxt):
    shapes=[]
    bw=int(max_w*val/maxval)
    shapes.append(textbox(x-2100000, y-30000, 2050000, h+60000, [para(label, sz=1250, align="r", space_after=0)], anchor="ctr"))
    shapes.append(rect(x, y, max_w, h, WHITE, line="E0E4E8"))
    shapes.append(roundrect(x, y, max(bw,60000), h, fill, None, rad=8000, shadow=False))
    shapes.append(textbox(x+max(bw,60000)+40000, y-30000, 900000, h+60000, [para(valtxt, sz=1250, bold=True, space_after=0)], anchor="ctr"))
    return shapes

def slide_compare():
    s, ytop = banner("How Big Were the Studies?", "COMPARISON OF SAMPLE SIZE")
    y=ytop+200000
    data=[("Rossouw & Fonagy 2012",80,BLUSH),("Bateman & Fonagy 2009",134,PEACH),
          ("Laurenssen 2018",109,LILAC),("Beck 2020",111,POWDER),("Griffiths 2019 (pilot)",53,SAGE)]
    maxv=140; bx=CX_L+2200000; bmax=CONTENT_W-2200000-1000000; bh=520000; gap=230000
    sy=y
    for lbl,v,c in data:
        s += hbar(bx, sy, bmax, bh, v, maxv, c, lbl, f"{v} people")
        sy+=bh+gap
    s.append(caption(CX_L, sy+10000, CONTENT_W, "Every study was small (under 150 people). Small studies give a less reliable answer, so even the combined result carries a wide margin of uncertainty. (Numbers from trial reports - confirm with the paper.)"))
    s += footer(10)
    return s
add(slide_compare(),
    "This bar chart makes one point visually: every study was small - all under about 150 people. Say why it "
    "matters: small studies are more easily thrown off by chance, so their results are less reliable, and even "
    "combined the answer has a wide margin of uncertainty. Use it to prepare for the appraisal: a big-looking "
    "benefit built from small studies deserves caution. Check exact numbers against the paper.")

def slide_researchgap():
    body=[
        para("Very few large, high-quality trials - most studies were small.", sz=BODY, bullet=True, space_after=280),
        para("Studies defined self-harm differently and measured outcomes with different tools.", sz=BODY, bullet=True, space_after=280),
        para("The therapy was delivered in different ways (one-to-one, group, teen version) and for different lengths of time.", sz=BODY, bullet=True, space_after=280),
        para("The comparison groups differed (usual care, structured support, or a waiting list).", sz=BODY, bullet=True, space_after=280),
        para("People were rarely followed for long, and adherence (whether the therapy was delivered as intended) was poorly reported.", sz=BODY, bullet=True, space_after=280),
        para("Almost no evidence from non-Western or low-resource settings; little direct comparison with other proven therapies (like DBT or CBT).", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("What Was Still Unknown", "RESEARCH GAP", body, 11, cardfill=LGREY)
add(slide_researchgap(),
    "List the gaps in everyday language: too few large, good-quality trials; different definitions and measures; "
    "different forms and lengths of therapy; different comparison groups; short follow-up; and often unclear whether "
    "the therapy was even delivered correctly. Almost nothing from non-Western settings and little head-to-head "
    "comparison with other proven therapies. These gaps justify combining the studies - and remind us to stay "
    "cautious about the final number.")

def slide_rationale():
    body=[
        para("Bring together all the scattered evidence into one overall answer.", sz=BODY, bullet=True, space_after=320),
        para("Measure how much MBT helps with self-harm and related problems (borderline personality symptoms, depression).", sz=BODY, bullet=True, space_after=320),
        para("Check whether the studies agree with each other, and why they might differ.", sz=BODY, bullet=True, space_after=320),
        para("Help clinicians decide when and for whom to use MBT.", sz=BODY, bullet=True, space_after=320),
        para("Point out what better future studies should do.", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("Why the Review Was Done", "RATIONALE", body, 12, cardfill=MINT)
add(slide_rationale(),
    "State the purpose simply: turn many small, conflicting studies into one clearer answer; measure how much MBT "
    "helps self-harm, borderline symptoms and depression; check whether studies agree; guide clinicians on when to "
    "use it; and flag what future research must improve. Emphasise the practical value for real decisions.")

def slide_aim():
    s, ytop = banner("Aim, Question & What They Expected", "OBJECTIVES")
    y=ytop+120000
    ch=(H-y-620000-2*180000)//3
    s.append(roundrect(CX_L, y, CONTENT_W, ch, LAVENDER, [
        multi_run_para([("Aim:  ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("to gather and combine the evidence on whether MBT reduces self-harm and improves related problems.",{"sz":BODY})], space_after=0)], line=LILAC, anchor="ctr"))
    s.append(roundrect(CX_L, y+ch+180000, CONTENT_W, ch, POWDER, [
        multi_run_para([("Question:  ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("Does MBT reduce self-harm (and related problems) more than usual care or other comparison conditions?",{"sz":BODY})], space_after=0)], line=ACCENT2, anchor="ctr"))
    s.append(roundrect(CX_L, y+2*(ch+180000), CONTENT_W, ch, MINT, [
        multi_run_para([("What they expected:  ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("MBT would lower self-harm and related symptoms compared with the control group. (A review like this uses set questions rather than a formal experimental hypothesis.)",{"sz":BODY})], space_after=0)], line=SAGE, anchor="ctr"))
    s += footer(13)
    return s
add(slide_aim(),
    "Keep this short. Aim: combine the evidence on whether MBT reduces self-harm. Question: does MBT help more than "
    "usual care or another comparison? What they expected: MBT would reduce self-harm and related symptoms. Explain "
    "the technical point simply - because this reviews existing studies rather than running a new experiment, it "
    "uses pre-set review questions instead of a classic hypothesis. Match wording to the paper.")

def slide_pico():
    s, ytop = banner("The Question in PICO Form", "A SIMPLE FRAMEWORK")
    y=ytop+120000
    s.append(textbox(CX_L, y, CONTENT_W, 360000, [para("PICO = a standard way to frame a clinical question: Population, Intervention, Comparison, Outcomes.", sz=1250, italic=True, color="6B6B75", space_after=0)]))
    y2=y+380000
    cw=(CONTENT_W-3*220000)//4
    ch=H-y2-620000
    cards=[("P - Population","(who was studied) People who self-harm - including those with borderline traits or poor emotion control; teenagers and adults.",BLUSH),
           ("I - Intervention","(the treatment) Mentalization-based therapy (MBT), including the teen version MBT-A; one-to-one or in a group.",LAVENDER),
           ("C - Comparison","(compared against) Usual care, structured support, a waiting list, or another therapy.",POWDER),
           ("O - Outcomes","(what was measured) Self-harm, suicidal behaviour, borderline symptoms, depression, functioning, and staying in treatment.",MINT)]
    cx=CX_L
    for head,txt,c in cards:
        ps=[para(head, sz=SUB, color=ACCENT, bold=True, space_after=240), para(txt, sz=BODY, space_after=0)]
        s.append(roundrect(cx, y2, cw, ch, c, ps, anchor="t"))
        cx+=cw+220000
    s += footer(14)
    return s
add(slide_pico(),
    "Introduce PICO as a simple recipe. Population: people who self-harm, including borderline traits, teens and "
    "adults. Intervention: MBT (and teen MBT-A), one-to-one or group. Comparison: usual care, structured support, "
    "waiting list, or another therapy. Outcomes: self-harm and suicidal behaviour first, then borderline symptoms, "
    "depression, functioning and staying in treatment. Just a tidy way to see what was tested.")

def slide_method():
    s, ytop = banner("How the Review Was Done", "METHOD OVERVIEW")
    y=ytop+120000
    stages=[("Plan &\nsearch",LAVENDER),("Screen &\nselect studies",POWDER),
            ("Pull out\nthe data",MINT),("Judge study\nquality",PEACH),("Combine results\n(meta-analysis)",SAGE)]
    bw=1950000; gap=280000; bh=980000
    total=len(stages)*bw+(len(stages)-1)*gap
    sx=(W-total)//2; sy=y+80000
    for i,(t,c) in enumerate(stages):
        lines=t.split("\n")
        ps=[para(lines[0], sz=1300, bold=True, align="ctr", space_after=60)]
        for extra in lines[1:]:
            ps.append(para(extra, sz=1300, bold=True, align="ctr", space_after=0))
        s.append(roundrect(sx, sy, bw, bh, c, ps, anchor="ctr", line=ACCENT))
        if i<len(stages)-1:
            s.append(arrow_right(sx+bw+30000, sy+bh/2-90000, gap-60000, 180000, ACCENT))
        sx+=bw+gap
    s.append(caption(CX_L, sy+bh+150000, CONTENT_W, "Five steps: plan and search for studies, pick the ones that fit, extract their numbers, judge how trustworthy each is, then combine them into one result."))
    s.append(roundrect(CX_L, sy+bh+820000, CONTENT_W, H-(sy+bh+820000)-560000, CREAM, [
        multi_run_para([("Key terms: ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("effect size = a number showing how big the benefit is; random-effects model = a fair way to combine studies that differ; I\u00B2 = how much the studies disagree.",{"sz":BODY})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(15)
    return s
add(slide_method(),
    "Explain the method as five plain steps: plan and search databases; screen and select studies that fit; pull "
    "out their numbers; judge how trustworthy each is; then combine them statistically. Read the key-terms box "
    "slowly - effect size is how big the benefit is; random-effects is a fair way to combine non-identical studies; "
    "I-squared measures how much they disagree. Say exact databases and dates come from the paper.")

def slide_eligibility():
    return two_col("Which Studies Were Let In (and Kept Out)","ELIGIBILITY RULES",
        "Included if...",
        ["People had self-harm or a relevant diagnosis (e.g., borderline personality disorder)",
         "The study tested MBT (or its teen version MBT-A)",
         "It reported self-harm or a related result",
         "It had a comparison group",
         "It gave enough numbers to calculate a benefit",
         "It was published in the set time period and language"],
        SAGE,
        "Excluded if...",
        ["No mentalization-based therapy was used",
         "No relevant self-harm result was reported",
         "It was a single case report (just one patient)",
         "It was only a conference summary with too little data",
         "It was a repeat publication of the same study",
         "The numbers needed could not be extracted"],
        BLUSH, 16)
add(slide_eligibility(),
    "Explain these as the entry rules deciding which studies count. In: people who self-harm, tested MBT, reported "
    "a relevant result, had a comparison group, and gave enough numbers. Out: no MBT, no relevant result, single "
    "case, brief conference summary, duplicate, or unusable numbers. Small differences in these rules change which "
    "studies qualify, so match them to the paper.")

print("content slides through 16:", len(SLIDES))

def slide_search():
    s, ytop = banner("How Studies Were Found", "SEARCH STRATEGY")
    y=ytop+120000
    lw=5400000
    lp=[para("Where they looked (research databases)", sz=SUB, color=ACCENT, bold=True, space_after=260)]
    for d in ["PubMed / MEDLINE","PsycINFO","Embase","Cochrane CENTRAL","Checking reference lists of found papers"]:
        lp.append(para(d, sz=BODY, bullet=True, space_after=220))
    lp.append(para("Confirm the exact databases and dates from the paper.", sz=1150, italic=True, color="9AA0A6", space_after=0))
    s.append(roundrect(CX_L, y, lw, H-y-720000, POWDER, lp, anchor="t"))
    rx=CX_L+lw+300000; rw=CONTENT_W-lw-300000
    s.append(roundrect(rx, y, rw, H-y-720000, CREAM, [
        para("How search words were combined", sz=SUB, color=ACCENT, bold=True, space_after=240),
        multi_run_para([('"mentalization-based therapy" OR "MBT"',{"sz":1250,"bold":True})], space_after=120),
        para("AND  (all groups must appear)", sz=1200, color=ACCENT2, bold=True, align="ctr", space_after=120),
        multi_run_para([('"self-harm" OR "self-injury"',{"sz":1250,"bold":True})], space_after=120),
        para("AND", sz=1200, color=ACCENT2, bold=True, align="ctr", space_after=120),
        multi_run_para([('"borderline personality" OR "emotion dysregulation"',{"sz":1250,"bold":True})], space_after=0),
    ], line=LILAC, anchor="t"))
    s.append(caption(CX_L, H-660000, CONTENT_W, "'OR' widens the search (any of these words); 'AND' narrows it (all groups must be present). This example is for illustration - use the paper's real search terms."))
    s += footer(17)
    return s
add(slide_search(),
    "Describe searching plainly: they looked through the main research databases and checked reference lists of "
    "found papers. Explain the logic using the caption - 'OR' widens by accepting any listed word, 'AND' narrows by "
    "requiring all groups together. Stress this string is only an example; real terms and dates come from the paper.")

def slide_prisma():
    s, ytop = banner("How Studies Were Narrowed Down", "PRISMA FLOW - ADD REAL NUMBERS")
    y=ytop+120000
    lx=CX_L+300000; lw=5000000
    stages=[("Studies found in databases (n = __)",LAVENDER),
            ("After removing duplicates (n = __)",POWDER),
            ("Titles/abstracts screened (n = __)",MINT),
            ("Full papers read in detail (n = __)",PEACH),
            ("Studies used in the review (n = __)",SAGE),
            ("Studies combined in the meta-analysis (n = __)",LAVENDER)]
    bh=520000; gap=150000; sy=y
    for i,(t,c) in enumerate(stages):
        s.append(roundrect(lx, sy, lw, bh, c, [para(t, sz=1250, bold=True, align="ctr", space_after=0)], line=ACCENT))
        if i<len(stages)-1:
            s.append(arrow_down(lx+lw/2-80000, sy+bh+2000, 160000, gap-12000, ACCENT))
        sy+=bh+gap
    ex=lx+lw+500000; ew=CONTENT_W-(lx-CX_L)-lw-500000
    s.append(roundrect(ex, y+2*(bh+gap), ew, bh, BLUSH,
        [para("Removed at screening (n = __)", sz=1200, bold=True, align="ctr", space_after=0)], line="D98B9E"))
    s.append(arrow_right(ex-460000, y+2*(bh+gap)+bh/2-80000, 420000, 160000, "D98B9E"))
    s.append(roundrect(ex, y+3*(bh+gap), ew, bh+140000, BLUSH,
        [para("Full papers removed, with reasons (n = __): no MBT / no relevant result / not enough data / duplicate",
              sz=1150, align="ctr", space_after=0)], line="D98B9E"))
    s.append(arrow_right(ex-460000, y+3*(bh+gap)+bh/2-80000, 420000, 160000, "D98B9E"))
    s.append(caption(ex-100000, y+4*(bh+gap)+bh, ew+200000, "PRISMA is a standard diagram showing how many studies started, how many were dropped and why, and how many were finally used."))
    s += footer(18)
    return s
add(slide_prisma(),
    "Explain PRISMA: a standard picture showing how you go from thousands of studies down to the few actually used, "
    "and why studies were dropped. Read the left column top to bottom - found, duplicates removed, screened, full "
    "papers read, then included and combined. The pink boxes show exclusions and reasons. Fill every 'n = __' from "
    "the paper's own diagram - do not guess.")

def slide_participants():
    s, ytop = banner("Who Were the People Studied?", "PARTICIPANTS")
    y=ytop+120000
    cw=(CONTENT_W-3*220000)//4
    ch=1520000
    cards=[("Age","Teenagers (MBT-A) and adults (MBT) - confirm exact ages from the paper",POWDER),
           ("Sex","Mostly female (e.g., about 85% female in Rossouw & Fonagy, 2012)",BLUSH),
           ("Diagnosis","Self-harm, often with borderline personality traits and poor emotion control",LAVENDER),
           ("Setting","Mostly outpatient clinics (not admitted to hospital) in wealthy, mainly European countries",MINT)]
    cx=CX_L
    for h,t,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [para(h, sz=SUB, color=ACCENT, bold=True, space_after=260),
                                              para(t, sz=BODY, space_after=0)], anchor="t"))
        cx+=cw+220000
    s.append(caption(CX_L, y+ch+150000, CONTENT_W, "The people studied were mostly young, mostly female, and mostly from Western countries - so results may not apply equally to everyone. Insert the total studies and people from the paper."))
    s += footer(19)
    return s
add(slide_participants(),
    "Describe the people plainly. They span teenagers (MBT-A) and adults (MBT). Most were female - about 85% in the "
    "2012 trial. Most had self-harm with borderline traits and emotion-control difficulty, treated in outpatient "
    "clinics in wealthy, mainly European countries. Use the caption to make the honest point about limited "
    "generalisability. Insert the exact totals from the paper.")

def slide_protocol():
    s, ytop = banner("What MBT Actually Involves", "THE TREATMENT")
    y=ytop+120000
    lw=5600000
    lp=[para("How it is delivered", sz=SUB, color=ACCENT, bold=True, space_after=260)]
    for t in ["One-to-one and/or group sessions (teen MBT-A also involves parents/family)",
              "Often about 12 months in the main teen study; session frequency varies",
              "Given by trained therapists who receive supervision (guidance from a senior therapist)",
              "Usually outpatient; safety/crisis support runs alongside the therapy",
              "Whether the therapy was delivered exactly as intended (fidelity) was not always reported"]:
        lp.append(para(t, sz=BODY, bullet=True, space_after=220))
    s.append(roundrect(CX_L, y, lw, H-y-620000, POWDER, lp, anchor="t"))
    rx=CX_L+lw+300000; rw=CONTENT_W-lw-300000
    rp=[para("The therapist's stance", sz=SUB, color=ACCENT, bold=True, space_after=260)]
    for t in ["Curious, 'not-knowing' (asks, doesn't assume)","Stays on present feelings",
              "Explores other ways to see a situation","Helps calm strong emotions",
              "Repairs misunderstandings","Avoids jumping to conclusions"]:
        rp.append(para(t, sz=BODY, bullet=True, space_after=210))
    s.append(roundrect(rx, y, rw, H-y-620000, MINT, rp, anchor="t"))
    s += footer(20)
    return s
add(slide_protocol(),
    "Describe what happens in MBT. Given one-to-one and/or in groups; the teen version adds work with parents. In "
    "the main teen study it ran about a year. Therapists are trained and supervised (guided by a senior colleague), "
    "usually outpatient with crisis support alongside. Note honestly that fidelity - whether it was delivered as "
    "designed - wasn't always reported. On the right, sum up the therapist's style: curious, not assuming, focused "
    "on present feelings, exploring other viewpoints, calming emotions, repairing misunderstandings, not jumping to "
    "conclusions.")

def slide_controls():
    s, ytop = banner("What MBT Was Compared Against", "COMPARISON GROUPS")
    y=ytop+120000
    headers=["Comparison group","What it means","Why it matters"]
    cw=[3000000, 4200000, CONTENT_W-3000000-4200000]
    rows=[
        ["Usual care (TAU)","The routine care people would normally get","Most common; if usual care is weak, MBT looks better than it is"],
        ["Structured support","Organised general support (a fair, active comparison)","A tougher test - MBT's extra benefit looks smaller"],
        ["Waiting list","No active therapy yet","Makes any therapy look better than it really is"],
        ["Another therapy","e.g., DBT or supportive therapy","Direct comparison; rare in this area"],
    ]
    s += table(CX_L, y, cw, 720000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1200, head_font=1250)
    s.append(caption(CX_L, y+720000*5+30000, CONTENT_W, "The result depends heavily on the comparison. Beating a waiting list is easy; beating real structured support is a much harder, more meaningful test."))
    s += footer(21)
    return s
add(slide_controls(),
    "Explain that the comparison group changes the meaning of the result. Most studies compared MBT with usual "
    "care, which is often weak - so part of MBT's apparent benefit may reflect poor usual care. Structured support "
    "is a fairer, tougher test and shrinks MBT's advantage. A waiting list flatters any therapy. Direct DBT "
    "comparisons are rare. Takeaway: beating a waiting list is easy; beating real structured support is the "
    "meaningful test.")

def slide_outcomes():
    s, ytop = banner("What Was Measured", "OUTCOMES")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-620000
    lp=[para("Main outcomes", sz=1650, color=ACCENT, bold=True, space_after=320),
        para("How often self-harm happened (number of episodes)", sz=BODY, bullet=True, space_after=260),
        para("How severe the self-harm was", sz=BODY, bullet=True, space_after=260),
        para("How long until it happened again", sz=BODY, bullet=True, space_after=260),
        para("Suicidal behaviour", sz=BODY, bullet=True, space_after=0)]
    rp=[para("Other outcomes", sz=1650, color=ACCENT, bold=True, space_after=320),
        para("Borderline personality symptoms", sz=BODY, bullet=True, space_after=240),
        para("Depression and anxiety", sz=BODY, bullet=True, space_after=240),
        para("Emotion control and impulsivity (acting without thinking)", sz=BODY, bullet=True, space_after=240),
        para("General functioning / quality of life", sz=BODY, bullet=True, space_after=240),
        para("Staying in treatment; use of hospital services", sz=BODY, bullet=True, space_after=0)]
    s.append(roundrect(CX_L, y, cw, ch, BLUSH, lp, anchor="t"))
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, POWDER, rp, anchor="t"))
    s += footer(22)
    return s
add(slide_outcomes(),
    "Split what was measured into main and other. Main: how often self-harm happened, how severe, how long until it "
    "returned, and suicidal behaviour. Other: borderline symptoms, depression and anxiety, emotion control and "
    "impulsivity, functioning and quality of life, and staying in treatment or needing hospital. Mention different "
    "studies used different questionnaires for the same thing - why a common effect size was needed to combine them.")

print("content slides through 22:", len(SLIDES))

def tl_dot(x,y,d,color):
    return oval(x,y,d,d,color,None,line="FFFFFF")
def slide_rob():
    s, ytop = banner("How Trustworthy Were the Studies?", "RISK OF BIAS - CONFIRM FROM PAPER")
    y=ytop+180000
    domains=["Fair random grouping","Group hidden from staff","Assessor didn't know the group",
             "Few people dropped out","All results reported","Therapy given as intended"]
    GREEN="8FCB9B"; YELLOW="F2D479"; RED="E39B9B"
    studies=["Rossouw 2012","Bateman 2009","Laurenssen 2018","Beck 2020"]
    pattern={
      "Rossouw 2012":[GREEN,YELLOW,YELLOW,GREEN,GREEN,YELLOW],
      "Bateman 2009":[GREEN,GREEN,YELLOW,YELLOW,GREEN,GREEN],
      "Laurenssen 2018":[GREEN,GREEN,GREEN,YELLOW,GREEN,YELLOW],
      "Beck 2020":[GREEN,GREEN,YELLOW,GREEN,YELLOW,GREEN],
    }
    x0=CX_L+3000000; colw=1350000; d=300000; rh=470000
    for i,dom in enumerate(domains):
        s.append(textbox(CX_L, y+ (i+1)*rh-40000, 2900000, rh, [para(dom, sz=1200, align="r", space_after=0)], anchor="ctr"))
    for j,st in enumerate(studies):
        s.append(textbox(x0+j*colw-200000, y-20000, colw, rh, [para(st, sz=1150, bold=True, align="ctr", space_after=0)], anchor="ctr"))
    for i,dom in enumerate(domains):
        for j,st in enumerate(studies):
            c=pattern[st][i]
            s.append(tl_dot(x0+j*colw+colw/2-d/2-200000, y+(i+1)*rh+rh/2-d/2-40000, d, c))
    ly=y+(len(domains)+1)*rh+30000
    s.append(tl_dot(CX_L, ly, 260000, GREEN)); s.append(textbox(CX_L+320000, ly-30000, 1500000, 320000,[para("Low risk (good)", sz=1150, space_after=0)],anchor="ctr"))
    s.append(tl_dot(CX_L+2000000, ly, 260000, YELLOW)); s.append(textbox(CX_L+2320000, ly-30000, 1800000, 320000,[para("Some concerns", sz=1150, space_after=0)],anchor="ctr"))
    s.append(tl_dot(CX_L+4200000, ly, 260000, RED)); s.append(textbox(CX_L+4520000, ly-30000, 1600000, 320000,[para("High risk (weak)", sz=1150, space_after=0)],anchor="ctr"))
    s.append(caption(CX_L, ly+400000, CONTENT_W, "Green = well done, yellow = some doubts, red = weak. In talking-therapy trials, hiding which group a person is in is hard, so 'blinding' is often the weak spot. This grid is an example - use the paper's real ratings."))
    s += footer(23)
    return s
add(slide_rob(),
    "Explain risk of bias simply: how much can we trust each study? Each dot rates one quality check - green good, "
    "yellow some doubts, red weak. Explain the checks plainly: was grouping truly random and fair, hidden from "
    "staff and assessors, did few drop out, were all results reported, was the therapy delivered as intended. Point "
    "out the common weak spot: in talking therapies you can't hide which treatment someone got, so blinding often "
    "scores yellow. This grid is illustrative - replace with the paper's real ratings.")

def slide_stats():
    s, ytop = banner("The Statistics, in Plain English", "HOW RESULTS WERE COMBINED")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-1080000
    lp=[para("Key ideas", sz=1650, color=ACCENT, bold=True, space_after=280),
        para("Effect size (Hedges' g) = one number for how big the benefit is; a negative number means self-harm went down (good).", sz=BODY, bullet=True, space_after=240),
        para("Confidence interval (CI) = the range the true result probably sits in; a narrow range = more certain.", sz=BODY, bullet=True, space_after=240),
        para("Random-effects model = a fair way to combine studies that are not identical.", sz=BODY, bullet=True, space_after=240),
        para("I\u00B2 = how much the studies disagree (higher = more disagreement).", sz=BODY, bullet=True, space_after=0)]
    s.append(roundrect(CX_L, y, cw, ch, LGREY, lp, anchor="t"))
    rp=[para("How big is 'big'? (effect size g)", sz=1650, color=ACCENT, bold=True, space_after=280),
        multi_run_para([("about 0.2 = small   \u2022   0.5 = medium   \u2022   0.8 = large",{"sz":BODY,"bold":True})], space_after=260),
        para("A combined score summarises the average benefit across all studies. But if the studies disagree a lot (high I\u00B2), that single average can hide real differences between them.", sz=BODY, space_after=0)]
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, CREAM, rp, anchor="t", line=LILAC))
    s.append(roundrect(CX_L, y+ch+120000, CONTENT_W, 720000, LAVENDER, [
        multi_run_para([("Remember:  ",{"bold":True,"color":ACCENT,"sz":BODY}),
                        ("'statistically significant' (unlikely to be chance) is not the same as 'clinically meaningful' (makes a real difference to the patient). Always mention both.",{"sz":BODY})], space_after=0)], line=LILAC, anchor="ctr"))
    s += footer(24)
    return s
add(slide_stats(),
    "Demystify the numbers. Effect size (Hedges' g) is one number for how big the benefit is; here negative is good "
    "(self-harm down). The confidence interval is the likely range - narrow means more certain. Random-effects is a "
    "fair way to combine differing studies; I-squared measures disagreement. Anchor the scale: 0.2 small, 0.5 "
    "medium, 0.8 large. Stress the difference between statistically significant (probably not chance) and clinically "
    "meaningful (actually helps) - mention both.")

def slide_res_selfharm():
    s, ytop = banner("Result 1 - Self-Harm", "MAIN OUTCOME (VERIFIED FROM THE PAPER)")
    y=ytop+220000
    px=CX_L+4300000; pw=3400000
    gmin,gmax=-2.0,1.0
    def sx(v): return px+int(pw*(v-gmin)/(gmax-gmin))
    s.append(line_shape(px, y+40000, pw, 0, "9AA0A6", 12700))
    zx=sx(0.0)
    s.append(line_shape(zx, y+40000, 0, 1450000, "B0B4B8", 12700))
    s.append(textbox(zx-560000, y+1470000, 1120000, 300000, [para("0 = no effect", sz=1050, color="9AA0A6", align="ctr", space_after=0)], anchor="ctr"))
    s.append(textbox(px-620000, y-300000, 1240000, 300000, [para("\u2190 favours MBT (better)", sz=1100, color=ACCENT2, bold=True, align="ctr", space_after=0)], anchor="ctr"))
    g,lo,hi=-0.82,-1.15,-0.50
    x1=sx(lo); x2=sx(hi); cy=y+300000
    s.append(textbox(px-4200000, cy-90000, 4100000, 420000, [para("Self-harm (all MBT & MBT-A studies combined)", sz=1200, align="r", space_after=0)], anchor="ctr"))
    s.append(line_shape(x1, cy, x2-x1, 0, "C25E7A", 28575))
    d=180000
    s.append(rect(sx(g)-d//2, cy-d//2, d, d, "C25E7A"))
    s.append(textbox(px+pw+80000, cy-90000, 2600000, 420000, [para("g = -0.82  [-1.15, -0.50]", sz=1200, bold=True, space_after=0)], anchor="ctr"))
    s.append(caption(CX_L, y+1900000, CONTENT_W, "The square is the average benefit; the line is the uncertainty range. Both sit fully on the 'better' side of 0, so MBT clearly lowered self-harm - a large effect (g = -0.82). The line is wide because the studies were few and small, so read it as 'a real but not exact benefit'."))
    s += footer(25)
    return s
add(slide_res_selfharm(),
    "This is the headline. Combining all studies, self-harm went down with an effect size of g = -0.82 (likely range "
    "-1.15 to -0.50). Explain the picture: the square is the average benefit, the line is the uncertainty range; "
    "because both sit left of zero (the 'better' side), MBT clearly reduced self-harm, and by the scale this is a "
    "large effect. Be honest: the line is wide because studies were few and small, so the true benefit could range "
    "from medium to very large. Verified from the paper.")

print("content slides through 25:", len(SLIDES))

def slide_res_suicide():
    s, ytop = banner("Result 2 - Suicidal Behaviour", "KEEP IT SEPARATE FROM SELF-HARM")
    y=ytop+120000
    s.append(roundrect(CX_L, y, CONTENT_W, 940000, PEACH, [
        multi_run_para([("Be precise:  ",{"bold":True,"color":"B5651D","sz":BODY}),
                        ("self-harm (may have no wish to die) is different from suicidal thoughts, a suicide attempt, and death by suicide. Many studies measured self-harm broadly, not suicide specifically.",{"sz":BODY})], space_after=0)], line="E0A96D", anchor="ctr"))
    s.append(roundrect(CX_L, y+1060000, CONTENT_W, H-(y+1060000)-560000, CREAM, [
        para("What the paper can and cannot say", sz=SUB, color=ACCENT, bold=True, space_after=280),
        para("The paper gives combined results for self-harm, borderline symptoms and depression - not a separate combined number just for suicide attempts or deaths.", sz=BODY, bullet=True, space_after=240),
        para("If there is no combined number for a suicide-specific outcome, say so plainly: 'No combined estimate was available for this.'", sz=BODY, bullet=True, space_after=240),
        para("Important: no proof of a suicide-specific benefit does NOT mean there is none - it just wasn't measured well enough. So we should not claim MBT prevents suicide.", sz=BODY, bullet=True, space_after=0),
    ], line=LILAC, anchor="t"))
    s += footer(26)
    return s
add(slide_res_suicide(),
    "Handle suicide carefully. First separate the terms: self-harm (may have no wish to die) is not the same as "
    "suicidal thoughts, a suicide attempt, or death by suicide. The verified combined results cover self-harm, "
    "borderline symptoms and depression - the paper does not give a separate combined number just for suicide "
    "attempts or deaths. If missing, say so directly. Make the logic clear: absence of a proven suicide-specific "
    "benefit does not prove there is none. So do not overclaim that MBT prevents suicide.")

def slide_res_related():
    s, ytop = banner("Result 3 - Personality Symptoms & Depression", "OTHER OUTCOMES (VERIFIED)")
    y=ytop+240000
    def gbar(x,y,maxw,h,g,lo,hi,label,color):
        sh=[]
        scale=1.6
        bw=int(maxw*abs(g)/scale)
        sh.append(textbox(x-3000000, y-30000, 2900000, h+60000, [para(label, sz=1250, align="r", space_after=0)], anchor="ctr"))
        sh.append(rect(x,y,maxw,h,WHITE,line="E0E4E8"))
        sh.append(roundrect(x,y,max(bw,60000),h,color,None,rad=6000,shadow=False))
        sh.append(textbox(x+max(bw,60000)+50000, y-30000, 3200000, h+60000,
            [para(f"g = {g:.2f}  [{lo:.2f}, {hi:.2f}]", sz=1250, bold=True, space_after=0)], anchor="ctr"))
        return sh
    bx=CX_L+3100000; bmax=CONTENT_W-3100000-3400000; bh=560000
    s+=gbar(bx, y, bmax, bh, -0.82, -1.15, -0.50, "Self-harm", "C25E7A")
    s+=gbar(bx, y+820000, bmax, bh, -1.08, -1.38, -0.77, "Borderline symptoms", ACCENT)
    s+=gbar(bx, y+1640000, bmax, bh, -1.10, -1.52, -0.68, "Depression", ACCENT2)
    s.append(caption(CX_L, y+2450000, CONTENT_W, "Longer bar = bigger improvement. MBT reduced self-harm, borderline personality symptoms and depression, all by a large amount. But these came from small studies, so treat the exact size cautiously."))
    s += footer(27)
    return s
add(slide_res_related(),
    "Show the two other verified results next to self-harm. Borderline personality symptoms improved by g = -1.08 "
    "(range -1.38 to -0.77) and depression by g = -1.10 (range -1.52 to -0.68). Explain the bars: longer means "
    "bigger improvement, and all three are large. Add the balanced note - these very large numbers come from small "
    "studies, which tend to produce bigger-looking effects, so treat the exact size cautiously. Verified from the "
    "paper.")

def slide_retention():
    body=[
        para("Report how many people finished the therapy versus dropped out, for MBT and the comparison group (add the paper's numbers).", sz=BODY, bullet=True, space_after=300),
        para("Small pilot studies (e.g., Griffiths et al., 2019) found group MBT was acceptable and safe, with reasonable attendance.", sz=BODY, bullet=True, space_after=300),
        para("Look at why people dropped out, and whether there were any harms (adverse events).", sz=BODY, bullet=True, space_after=300),
        para("The key question: is MBT practical and acceptable for people who repeatedly self-harm and struggle with emotions?", sz=BODY, bullet=True, space_after=300),
        para("Drop-out and safety information was often incompletely reported - a real weakness to flag.", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("Did People Stay in Treatment?","RETENTION & ACCEPTABILITY", body, 28, cardfill=LGREY)
add(slide_retention(),
    "Feasibility matters as much as effectiveness here. Report how many finished versus dropped out in each arm from "
    "the paper. Small pilot work (Griffiths 2019) suggests group MBT is acceptable and safe with reasonable "
    "attendance. Look at reasons for dropout and any harms. The real-world question is whether MBT is practical for "
    "people who repeatedly self-harm - and flag honestly that dropout and safety reporting was often incomplete.")

def slide_gfx_summary():
    s, ytop = banner("Results at a Glance", "THE THREE VERIFIED NUMBERS")
    y=ytop+160000
    cards=[("Self-harm","g = -0.82","range -1.15 to -0.50",BLUSH),
           ("Borderline symptoms","g = -1.08","range -1.38 to -0.77",LAVENDER),
           ("Depression","g = -1.10","range -1.52 to -0.68",POWDER)]
    cw=(CONTENT_W-2*300000)//3; ch=1980000
    cx=CX_L
    for h,g,ci,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [
            para(h, sz=SUB, color=ACCENT, bold=True, align="ctr", space_after=240),
            para(g, sz=2600, color=CHARCOAL, bold=True, align="ctr", space_after=200),
            para(ci, sz=1250, align="ctr", space_after=140),
            para("large & unlikely to be chance", sz=1100, color=ACCENT2, bold=True, align="ctr", space_after=0)], anchor="ctr", line=LILAC))
        cx+=cw+300000
    s.append(caption(CX_L, y+ch+150000, CONTENT_W, "All three point the same way: MBT helped, and by a large amount. But because the studies were few, small and different from each other, the honest summary is 'promising, not yet proven'."))
    s += footer(29)
    return s
add(slide_gfx_summary(),
    "Give the one-glance summary: self-harm -0.82, borderline symptoms -1.08, depression -1.10 - all large, all "
    "unlikely to be chance, all favouring MBT. Use it to bridge to the discussion: the signal is encouraging and "
    "consistent, but rests on few, small, differing studies. So the honest, examiner-friendly summary is "
    "'promising, not yet proven'.")

print("content slides through 29:", len(SLIDES))

def slide_disc_main():
    body=[
        para("MBT reduced self-harm by a large amount (g = -0.82) - a real, unlikely-to-be-chance benefit.", sz=BODY, bullet=True, space_after=280),
        para("It also improved borderline personality symptoms (g = -1.08) and depression (g = -1.10) by a large amount.", sz=BODY, bullet=True, space_after=280),
        para("All the combined results pointed the same way - in favour of MBT.", sz=BODY, bullet=True, space_after=280),
        para("But the results come from only a few small studies that differed from each other.", sz=BODY, bullet=True, space_after=280),
        para("So the fairest verdict: the evidence is promising but limited - we can be only moderately confident.", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("Discussion - What Did We Find?","INTERPRETATION", body, 30, cardfill=MINT)
add(slide_disc_main(),
    "Sum up plainly. MBT reduced self-harm by a large, real amount, and improved borderline symptoms and depression "
    "a lot. Every combined result favoured MBT and was unlikely to be chance. Then the balanced verdict: because "
    "results come from only a few small, differing studies, be only moderately confident - promising but not yet "
    "solid. Don't oversell.")

def slide_disc_mech():
    s, ytop = banner("How Might MBT Work?", "LIKELY REASONS (SAID CAUTIOUSLY)")
    y=ytop+120000
    mechs=[("Better 'mind-reading'","Understands own & others' feelings more accurately",LAVENDER),
           ("Fewer misreadings","Less likely to assume the worst about others",POWDER),
           ("Calmer emotions","Can sit with distress without acting on it",MINT),
           ("Less impulsive","A pause between the urge and the action",PEACH),
           ("Stronger bond with therapist","Feeling understood keeps people engaged",BLUSH),
           ("Spots triggers early","Uses new coping instead of self-harm",SAGE)]
    cw=(CONTENT_W-2*260000)//3; ch=(H-y-720000-260000)//2
    for i,(h,t,c) in enumerate(mechs):
        r=i//3; col=i%3
        s.append(roundrect(CX_L+col*(cw+260000), y+r*(ch+260000), cw, ch, c,
            [para(h, sz=SUB, color=ACCENT, bold=True, space_after=220), para(t, sz=BODY, space_after=0)], anchor="t"))
    s.append(caption(CX_L, y+2*ch+260000+30000, CONTENT_W, "These are the likely reasons MBT helps. In the 2012 teen study, the benefit really did come from better 'mind-reading'. But a review of results cannot fully prove the reason - so we say 'may'."))
    s += footer(31)
    return s
add(slide_disc_mech(),
    "Offer likely reasons, tentatively. MBT probably helps by improving mind-reading so people understand feelings "
    "better and misjudge others less; by staying calmer and sitting with distress; by adding a pause between urge "
    "and action; by building a bond that keeps them in therapy; and by spotting triggers and coping differently. "
    "Supportive evidence: in the 2012 study the benefit flowed through improved mind-reading. But combining outcome "
    "results can't fully prove mechanism - hence 'may'.")

def slide_disc_compare():
    s, ytop = banner("How Does MBT Compare With Other Therapies?", "CONTEXT")
    y=ytop+120000
    headers=["Therapy","Evidence for self-harm","Note"]
    cw=[2600000, 4600000, CONTENT_W-2600000-4600000]
    rows=[
        ["MBT / MBT-A","Large combined benefit here; early studies positive, later ones mixed","This paper's focus"],
        ["DBT (dialectical behaviour therapy)","Strongest, best-repeated evidence for teen self-harm","The main benchmark"],
        ["CBT-based","Moderate evidence for reducing repeat self-harm","Widely available"],
        ["Usual care / structured support","A fair comparison shrinks MBT's apparent lead","The comparison matters a lot"],
    ]
    s+=table(CX_L, y, cw, 720000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1200, head_font=1250)
    s.append(caption(CX_L, y+720000*5+30000, CONTENT_W, "MBT looks helpful, but DBT still has the strongest, most repeated evidence. Differences between MBT studies are likely due to who was treated, how intense the therapy was, and how outcomes were measured."))
    s += footer(32)
    return s
add(slide_disc_compare(),
    "Place MBT among alternatives fairly. DBT has the strongest, best-replicated evidence for teen self-harm and is "
    "the natural benchmark; CBT-based approaches show moderate benefit. MBT's combined benefit here looks large but "
    "rests on fewer studies with mixed later results. Explain disagreements plainly: different patients, different "
    "intensity, different outcome measures. Takeaway: MBT is a reasonable option, but this evidence doesn't yet show "
    "it beats DBT.")

def slide_clinical():
    body=[
        para("MBT is worth considering for people who repeatedly self-harm and have borderline personality traits.", sz=BODY, bullet=True, space_after=270),
        para("Tailor treatment to the person; keep assessing risk throughout.", sz=BODY, bullet=True, space_after=270),
        para("MBT adds to - it does not replace - safety planning and crisis support.", sz=BODY, bullet=True, space_after=270),
        para("Medication may still be needed for other conditions; keeping the person engaged and in continuous care matters a lot.", sz=BODY, bullet=True, space_after=270),
        para("Therapists need proper training and supervision; match the setting to how high the risk is.", sz=BODY, bullet=True, space_after=270),
        para("Involve family or carers where helpful (especially in the teen version, MBT-A).", sz=BODY, bullet=True, space_after=0),
    ]
    return content_slide("What This Means for Practice","CLINICAL IMPLICATIONS", body, 33, cardfill=POWDER)
add(slide_clinical(),
    "Translate into practice. MBT is worth considering for people who repeatedly self-harm with borderline traits, "
    "but tailor it and always keep it alongside - never instead of - ongoing risk assessment and crisis planning. "
    "Other conditions may need medication, and keeping the person engaged in continuous care is decisive. Emphasise "
    "trained, supervised therapists, matching setting to risk, and involving family, especially for teenagers.")

print("content slides through 33:", len(SLIDES))

def slide_limits():
    s, ytop = banner("Weaknesses of the Review", "LIMITATIONS")
    y=ytop+120000
    headers=["Weakness","Why it matters"]
    cw=[5200000, CONTENT_W-5200000]
    rows=[
        ["Few, small studies","Less reliable; wider uncertainty range"],
        ["Studies differed a lot (setup, format, comparison)","Harder to trust one combined number"],
        ["Short or varied follow-up","We don't know if the benefit lasts"],
        ["Different measuring tools","Harder to compare and combine fairly"],
        ["People dropped out; gaps in reporting","Results may be skewed"],
        ["Mostly female, young, Western","May not apply to everyone"],
        ["Self-harm & suicide not always separated","Unclear exactly what improved"],
    ]
    s+=table(CX_L, y, cw, 560000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1300, head_font=1350)
    s += footer(34)
    return s
add(slide_limits(),
    "Be open about weaknesses - examiners reward balance. Few, small studies mean a less reliable answer and wide "
    "uncertainty. Studies differed a lot in setup, format and comparison, making one combined number harder to "
    "trust. Follow-up was short or inconsistent, so we don't know if benefit lasts. Different tools, dropout, "
    "reporting gaps, a mostly young/female/Western sample, and not always separating self-harm from suicide all "
    "lower certainty. This is why the large effects are 'promising, not proven'.")

def slide_sample_design():
    return two_col("Weaknesses in the People & the Study Design","LIMITATIONS (DETAIL)",
        "The people (sample)",
        ["Mostly female participants","Few males and gender-diverse people",
         "Few older adults","Almost all from Western, wealthy countries",
         "Different starting severity between studies"],
        BLUSH,
        "The study design",
        ["Grouping not always fully random","Assessors often knew who got MBT",
         "'Usual care' meant different things in different studies","High or uneven drop-out",
         "Little checking that therapy was delivered as intended"],
        POWDER, 35)
add(slide_sample_design(),
    "Split weaknesses into who was studied and how. People: mostly female, few males or gender-diverse people, few "
    "older adults, almost all Western, differing severity - a fairly narrow group. Design: grouping not always fully "
    "random, assessors often knew who got MBT (can bias ratings), 'usual care' varied, dropout high or uneven, and "
    "rarely checked whether therapy was delivered properly. These show exactly what better trials must fix.")

def slide_general():
    s, ytop = banner("Who Do These Results Apply To?", "GENERALIZABILITY")
    y=ytop+120000
    cw=(CONTENT_W-300000)//2; ch=H-y-720000
    lp=[para("More likely to apply to", sz=1600, color=ACCENT, bold=True, space_after=320)]
    for t in ["Outpatient clinics (people not admitted to hospital)","Teens & adults who self-harm with borderline traits","Western, higher-income settings","Groups that are mostly female"]:
        lp.append(para(t, sz=BODY, bullet=True, space_after=270))
    rp=[para("Apply with caution to", sz=1600, color=ACCENT, bold=True, space_after=320)]
    for t in ["Emergency departments and hospital wards","Low-resource or non-Western settings","Older adults; males & gender-diverse people","Self-harm without borderline traits; people with psychosis or major memory/thinking problems"]:
        rp.append(para(t, sz=BODY, bullet=True, space_after=250))
    s.append(roundrect(CX_L, y, cw, ch, SAGE, lp, anchor="t"))
    s.append(roundrect(CX_L+cw+300000, y, cw, ch, PEACH, rp, anchor="t"))
    s.append(caption(CX_L, H-660000, CONTENT_W, "The findings fit the kind of people who were studied. For groups barely included (older adults, non-Western, males), we simply don't have enough evidence yet."))
    s += footer(36)
    return s
add(slide_general(),
    "Explain generalisability as 'who can we safely apply this to?'. The findings fit the people studied: mostly "
    "female teens and adults with self-harm and borderline traits, in outpatient clinics in Western countries. Be "
    "cautious applying them to emergency or hospital settings, low-resource or non-Western populations, older "
    "adults, males and gender-diverse people, or self-harm without borderline traits and people with psychosis - "
    "these were barely included.")

def slide_future():
    s, ytop = banner("What Next? Uses & Implications", "LOOKING FORWARD")
    y=ytop+120000
    cw=(CONTENT_W-3*220000)//4; ch=H-y-620000
    cards=[("In the clinic","Use MBT ideas in everyday care; check 'mind-reading' during crises; build the picture together with the patient; make personal safety plans",LAVENDER),
           ("In services","Set up proper MBT programmes; train and supervise therapists; step up care as needed; keep support going after discharge",POWDER),
           ("Preventing relapse","Spot early warning signs; build coping alternatives; plan for relationship crises; track over time",MINT),
           ("With family/society","Involve family or carers; address past trauma and attachment; support social life; coordinate the whole care team",PEACH)]
    cx=CX_L
    for h,t,c in cards:
        s.append(roundrect(cx, y, cw, ch, c, [para(h, sz=SUB, color=ACCENT, bold=True, space_after=260), para(t, sz=BODY, space_after=0)], anchor="t"))
        cx+=cw+220000
    s += footer(37)
    return s
add(slide_future(),
    "Lay out implications at four levels. In the clinic: use MBT ideas day to day, check mind-reading during "
    "crises, build the formulation with the patient, and make personal safety plans. In services: set up proper MBT "
    "programmes with training and supervision, step care up when needed, and keep support going after discharge. "
    "Relapse prevention: spot early warning signs, build coping alternatives, track over time. With family/society: "
    "involve carers, address trauma and attachment, coordinate the team.")

def slide_reflection():
    s, ytop = banner("If I Were the Researcher, I Would...", "MY CRITICAL REFLECTION")
    y=ytop+120000
    items=["Study more people, and a more varied group (age, sex, culture)",
           "Follow them for longer, using the same measures each time",
           "Clearly separate suicidal from non-suicidal self-harm",
           "Group people truly randomly, and keep assessors 'blind' to the group",
           "Compare MBT against another real therapy (like DBT), not just usual care",
           "Check the therapy was delivered properly; report therapist training",
           "Test how it works and for whom (subgroups)",
           "Register the plan in advance; report all drop-outs and harms fully",
           "Check whether it is good value for money; follow up after discharge"]
    cw=(CONTENT_W-260000)//2; ch=H-y-620000
    half=(len(items)+1)//2
    lp=[para(t, sz=BODY, bullet=True, space_after=280) for t in items[:half]]
    rp=[para(t, sz=BODY, bullet=True, space_after=280) for t in items[half:]]
    s.append(roundrect(CX_L, y, cw, ch, CREAM, lp, anchor="t", line=LILAC))
    s.append(roundrect(CX_L+cw+260000, y, cw, ch, CREAM, rp, anchor="t", line=LILAC))
    s += footer(38)
    return s
add(slide_reflection(),
    "Show your own critical thinking. If you ran this, you would study more people and a more varied group, follow "
    "them longer with consistent measures, and clearly separate suicidal from non-suicidal self-harm. You would "
    "group people truly at random, keep assessors blind, and compare against a real therapy like DBT rather than "
    "only usual care. You would check the therapy was delivered properly, report training, test how and for whom it "
    "works, register the plan in advance, report all dropouts and harms, and look at value for money and "
    "after-discharge outcomes. This answers the review's own weaknesses.")

print("content slides through 38:", len(SLIDES))

def slide_takeaways():
    s, ytop = banner("Key Takeaways", "THE 6 THINGS TO REMEMBER")
    y=ytop+120000
    pts=["Self-harm keeps coming back - always assess safety and risk carefully.",
         "MBT targets the breakdown in 'reading feelings' that happens when emotions run high.",
         "Combined results show large benefits: self-harm g=-0.82, borderline symptoms g=-1.08, depression g=-1.10.",
         "But few, small, differing studies with short follow-up mean: promising, not yet proven.",
         "MBT fits best for people with relationship sensitivity, poor emotion control & borderline traits.",
         "We need bigger, more varied, longer studies - ideally compared with other real therapies."]
    ch=(H-y-620000-5*140000)//6
    colors=[BLUSH,PEACH,SAGE,POWDER,LAVENDER,MINT]
    sy=y
    for i,(p,c) in enumerate(zip(pts,colors)):
        s.append(roundrect(CX_L, sy, CONTENT_W, ch, c, [
            multi_run_para([(f"{i+1}.  ",{"bold":True,"color":ACCENT,"sz":BODY}),(p,{"sz":BODY})], space_after=0)], anchor="ctr"))
        sy+=ch+140000
    s += footer(39)
    return s
add(slide_takeaways(),
    "Deliver these six messages slowly and plainly. Self-harm keeps coming back, so safety assessment is essential. "
    "MBT targets the 'can't read feelings' breakdown under stress. Combined results show large benefits across "
    "self-harm, borderline symptoms and depression. But few, small, differing, short studies mean 'promising, not "
    "yet proven'. MBT fits best for people with relationship sensitivity and borderline traits. And the field needs "
    "bigger, more varied, longer studies against other real therapies. End balanced.")

def slide_appraisal():
    s, ytop = banner("Overall Quality Check", "CRITICAL APPRAISAL (template)")
    y=ytop+120000
    headers=["Question","Answer"]
    cw=[6800000, CONTENT_W-6800000]
    rows=[
        ["Was the question clear?","Yes"],
        ["Was the study type right (review + meta-analysis)?","Yes"],
        ["Was the search thorough?","Confirm from paper"],
        ["Were the entry rules sensible?","Mostly yes"],
        ["Was study quality checked?","Confirm from paper"],
        ["Were the right statistics used?","Yes"],
        ["Was disagreement between studies handled?","Partly"],
        ["Is it useful for real patients?","Moderate-High"],
        ["Does it apply widely?","Limited"],
        ["Overall, how confident can we be?","Moderate-Low"],
    ]
    s+=table(CX_L, y, cw, 430000, headers, rows, head_fill=ACCENT, zebra=(CREAM, LAVENDER), font=1300, head_font=1350)
    s += footer(40)
    return s
add(slide_appraisal(),
    "Give a simple structured verdict, like ticking a checklist. Question and study type were right; statistics "
    "appropriate. Confirm from the full paper whether the search was thorough and study quality formally checked. "
    "Disagreement between studies was only partly handled. Moderately-to-highly useful for patients, but applies to "
    "a narrow group, so overall confidence is moderate-to-low. Adjust each answer after reading the full methods.")

def slide_refs(part, refs, page):
    body=[]
    for r in refs:
        body.append(para(r, sz=1250, bullet=False, space_after=300))
    s, ytop = banner("References" + (" (cont.)" if part==2 else ""), "APA 7TH EDITION - VERIFY BEFORE USE")
    s.append(textbox(CX_L, ytop+120000, CONTENT_W, H-ytop-560000, body, anchor="t"))
    s += footer(page)
    return s
refs1=[
 "Primary paper: (2024). Efficacy of mentalization-based therapy in treating self-harm: A systematic review and meta-analysis. Suicide and Life-Threatening Behavior. https://doi.org/10.1111/sltb.13044  [insert the full author list from the paper]",
 "Rossouw, T. I., & Fonagy, P. (2012). Mentalization-based treatment for self-harm in adolescents: A randomized controlled trial. Journal of the American Academy of Child & Adolescent Psychiatry, 51(12), 1304-1313.",
 "Bateman, A., & Fonagy, P. (2009). Randomized controlled trial of outpatient mentalization-based treatment versus structured clinical management for borderline personality disorder. American Journal of Psychiatry, 166(12), 1355-1364.",
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
 "Note: check every reference against the paper's own reference list before presenting; do not present unverified details as the authors' own.",
]
add(slide_refs(1, refs1, 41),
    "Present the key references. Lead with the primary paper (DOI 10.1111/sltb.13044) and insert its full author "
    "list. Then the landmark studies (Rossouw & Fonagy 2012; Bateman & Fonagy 2009) and the MBT manual. These are "
    "cross-checked, but a few details are marked 'confirm' and must be verified against the paper before the talk.")
add(slide_refs(2, refs2, 42),
    "Continue with the other studies, the PRISMA guideline, NICE self-harm guidance, and the diagnostic manuals "
    "(ICD-11, DSM-5-TR). Repeat the accuracy rule: verify each reference against the paper, and never attribute "
    "unverified specifics to the authors.")

def slide_questions():
    s=[]
    s.append(rect(0,0,W,H,ACCENT))
    s.append(oval(-600000,-600000,2400000,2400000,ACCENT2))
    s.append(oval(W-1800000,H-1800000,2400000,2400000,"5A4A8A"))
    s.append(textbox(1200000, 2400000, W-2400000, 1400000, [
        para("Thank You - Questions & Discussion", sz=4000, color=WHITE, bold=True, align="ctr", space_after=260),
        para("Efficacy of Mentalization-Based Therapy in Treating Self-Harm: A Systematic Review & Meta-Analysis",
             sz=1500, color=LAVENDER, italic=True, align="ctr", space_after=0)]))
    s.append(textbox(1200000, 4300000, W-2400000, 600000, [
        para("Sonal  \u2022  M.Phil. Clinical Psychology  \u2022  DOI: 10.1111/sltb.13044", sz=1300, color="D6C7EE", align="ctr", space_after=0)]))
    return s
add(slide_questions(),
    "Close and open the floor with simple prompts: (1) For a teenager who self-harms with borderline traits, would "
    "you choose MBT or DBT first, and why? (2) How much should a large benefit change practice when it comes from "
    "small, differing studies? (3) What evidence would convince us MBT reduces suicide itself, not just self-harm? "
    "Thank the audience.")

print("TOTAL CONTENT SLIDES:", len(SLIDES))
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
