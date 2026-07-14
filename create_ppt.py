#!/usr/bin/env python3
"""
Create a visually attractive PowerPoint presentation for MPhil Research Proposal:
Rorschach Cognitive Triad as Predictor of Subjective Loneliness in Schizophrenia
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# Color palette - Professional deep blue/teal theme
PRIMARY = RGBColor(0x1B, 0x2A, 0x4A)      # Deep navy
SECONDARY = RGBColor(0x2E, 0x86, 0xAB)    # Teal blue
ACCENT = RGBColor(0xE8, 0x4D, 0x4D)       # Warm red accent
LIGHT_BG = RGBColor(0xF0, 0xF4, 0xF8)     # Light blue-grey
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)
GOLD = RGBColor(0xD4, 0xA5, 0x37)         # Gold accent
GREEN = RGBColor(0x27, 0xAE, 0x60)        # Success green
PURPLE = RGBColor(0x6C, 0x5C, 0xE7)       # Purple

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def add_gradient_bg(slide, color1, color2):
    """Add a colored rectangle as background"""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color1
    bg.line.fill.background()
    # Move to back
    sp = bg._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)

def add_accent_bar(slide, top, width=Inches(2), color=SECONDARY):
    """Add a decorative accent bar"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top, width, Pt(5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

def add_circle_icon(slide, left, top, size, color, text=""):
    """Add a circle with text inside as an icon"""
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.fill.background()
    if text:
        tf = circle.text_frame
        tf.text = text
        tf.paragraphs[0].font.size = Pt(14)
        tf.paragraphs[0].font.color.rgb = WHITE
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.word_wrap = True

def make_title_slide():
    """Slide 1: Title Slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_gradient_bg(slide, PRIMARY, PRIMARY)
    
    # Decorative top bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = GOLD
    bar.line.fill.background()
    
    # Title
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11), Inches(2.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Rorschach Indices of the Cognitive Triad"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = "(Processing, Mediation, and Ideation)"
    p2.font.size = Pt(28)
    p2.font.color.rgb = SECONDARY
    p2.alignment = PP_ALIGN.CENTER
    
    p3 = tf.add_paragraph()
    p3.text = "as Predictors of Subjective Loneliness"
    p3.font.size = Pt(32)
    p3.font.bold = True
    p3.font.color.rgb = WHITE
    p3.alignment = PP_ALIGN.CENTER
    
    p4 = tf.add_paragraph()
    p4.text = "in Patients with Schizophrenia"
    p4.font.size = Pt(32)
    p4.font.bold = True
    p4.font.color.rgb = WHITE
    p4.alignment = PP_ALIGN.CENTER
    
    # Subtitle info
    txBox2 = slide.shapes.add_textbox(Inches(2), Inches(5.2), Inches(9), Inches(1.8))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p5 = tf2.paragraphs[0]
    p5.text = "MPhil Clinical Psychology Research Proposal"
    p5.font.size = Pt(18)
    p5.font.color.rgb = GOLD
    p5.alignment = PP_ALIGN.CENTER
    
    p6 = tf2.add_paragraph()
    p6.text = "\nCandidate: [Your Name]  |  Guide: [Guide Name]"
    p6.font.size = Pt(14)
    p6.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
    p6.alignment = PP_ALIGN.CENTER
    
    p7 = tf2.add_paragraph()
    p7.text = "[Institution Name]  |  2026-2027"
    p7.font.size = Pt(14)
    p7.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
    p7.alignment = PP_ALIGN.CENTER

make_title_slide()


def make_content_slide(title, bullets, accent_color=SECONDARY):
    """Generic content slide with styled bullets"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, WHITE, WHITE)
    
    # Left accent stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.15), prs.slide_height)
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_color
    stripe.line.fill.background()
    
    # Title area
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    
    add_accent_bar(slide, Inches(1.1), Inches(3), accent_color)
    
    # Bullets
    txBox2 = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.5))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf2.add_paragraph() if i > 0 else tf2.paragraphs[0]
        p.text = bullet
        p.font.size = Pt(16)
        p.font.color.rgb = DARK_TEXT
        p.space_after = Pt(10)
        p.level = 0
    return slide

def make_two_column_slide(title, left_title, left_items, right_title, right_items):
    """Two-column comparison slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, LIGHT_BG, LIGHT_BG)
    
    # Title
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    
    add_accent_bar(slide, Inches(0.95), Inches(2.5), SECONDARY)
    
    # Left column box
    left_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 
                                       Inches(0.5), Inches(1.5), Inches(5.8), Inches(5.5))
    left_box.fill.solid()
    left_box.fill.fore_color.rgb = WHITE
    left_box.line.color.rgb = SECONDARY
    left_box.line.width = Pt(2)
    
    # Left content
    txL = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.3), Inches(5.3))
    tfL = txL.text_frame
    tfL.word_wrap = True
    pL = tfL.paragraphs[0]
    pL.text = left_title
    pL.font.size = Pt(18)
    pL.font.bold = True
    pL.font.color.rgb = SECONDARY
    for item in left_items:
        p = tfL.add_paragraph()
        p.text = item
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT
        p.space_after = Pt(6)
    
    # Right column box
    right_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Inches(6.8), Inches(1.5), Inches(5.8), Inches(5.5))
    right_box.fill.solid()
    right_box.fill.fore_color.rgb = WHITE
    right_box.line.color.rgb = ACCENT
    right_box.line.width = Pt(2)
    
    # Right content
    txR = slide.shapes.add_textbox(Inches(7.1), Inches(1.6), Inches(5.3), Inches(5.3))
    tfR = txR.text_frame
    tfR.word_wrap = True
    pR = tfR.paragraphs[0]
    pR.text = right_title
    pR.font.size = Pt(18)
    pR.font.bold = True
    pR.font.color.rgb = ACCENT
    for item in right_items:
        p = tfR.add_paragraph()
        p.text = item
        p.font.size = Pt(13)
        p.font.color.rgb = DARK_TEXT
        p.space_after = Pt(6)
    return slide


# ============ SLIDE 2: INTRODUCTION ============
make_content_slide(
    "Introduction",
    [
        "* Schizophrenia affects 0.3-0.7% of the global population with pervasive cognitive,",
        "   social, and functional impairments that persist despite pharmacological treatment.",
        "",
        "* Loneliness -- the painful gap between desired and actual social connection --",
        "   is one of the most neglected yet devastating experiences in schizophrenia.",
        "",
        "* Meta-analysis (2018): Significant positive relationship between loneliness & psychosis.",
        "   Loneliness linked to hospitalization, cognitive decline, and mortality (Green et al., 2023).",
        "",
        "* The Cognitive Model of Loneliness (Cacioppo & Hawkley, 2009) proposes that",
        "   loneliness arises from HOW one processes and perceives the social world --",
        "   not merely from being objectively alone.",
        "",
        "* The Rorschach Cognitive Triad (Processing -> Mediation -> Ideation) captures",
        "   EXACTLY these cognitive operations through a performance-based assessment."
    ]
)

# ============ SLIDE 3: THE PROBLEM ============
def make_problem_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY)
    
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "The Problem: Why Are Patients With Schizophrenia Lonely?"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = GOLD
    
    # Three boxes representing the cognitive pathway
    colors = [SECONDARY, ACCENT, PURPLE]
    labels = [
        ("POOR PROCESSING", "They scan their social\nworld hastily or\noversimplistically"),
        ("MISPERCEPTION", "They misread neutral\nsocial cues as\nthreatening/rejecting"),
        ("DISTORTED THINKING", "They reason about\nrelationships in\nillogical ways")
    ]
    
    for i, (title, desc) in enumerate(labels):
        left = Inches(0.8 + i * 4.2)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      left, Inches(1.8), Inches(3.8), Inches(2.5))
        box.fill.solid()
        box.fill.fore_color.rgb = colors[i]
        box.line.fill.background()
        
        tx = slide.shapes.add_textbox(left + Inches(0.2), Inches(1.9), Inches(3.4), Inches(2.3))
        t = tx.text_frame
        t.word_wrap = True
        p1 = t.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(15)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p1.alignment = PP_ALIGN.CENTER
        p2 = t.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(13)
        p2.font.color.rgb = WHITE
        p2.alignment = PP_ALIGN.CENTER
    
    # Arrow connections
    for i in range(2):
        arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                        Inches(4.4 + i * 4.2), Inches(2.8), Inches(0.6), Inches(0.4))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = GOLD
        arrow.line.fill.background()
    
    # Bottom result
    result_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                         Inches(3.5), Inches(4.8), Inches(6), Inches(1.2))
    result_box.fill.solid()
    result_box.fill.fore_color.rgb = ACCENT
    result_box.line.fill.background()
    
    tx2 = slide.shapes.add_textbox(Inches(3.7), Inches(4.9), Inches(5.6), Inches(1))
    t2 = tx2.text_frame
    t2.word_wrap = True
    p3 = t2.paragraphs[0]
    p3.text = "RESULT: Profound Subjective Loneliness"
    p3.font.size = Pt(18)
    p3.font.bold = True
    p3.font.color.rgb = WHITE
    p3.alignment = PP_ALIGN.CENTER
    p4 = t2.add_paragraph()
    p4.text = "\"I am surrounded by people but feel utterly alone\""
    p4.font.size = Pt(13)
    p4.font.italic = True
    p4.font.color.rgb = WHITE
    p4.alignment = PP_ALIGN.CENTER
    
    # Big downward arrow
    darrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                                     Inches(6.2), Inches(4.3), Inches(0.5), Inches(0.5))
    darrow.fill.solid()
    darrow.fill.fore_color.rgb = GOLD
    darrow.line.fill.background()
    
    # Footer
    tx3 = slide.shapes.add_textbox(Inches(1), Inches(6.3), Inches(11), Inches(0.8))
    t3 = tx3.text_frame
    p5 = t3.paragraphs[0]
    p5.text = "The Rorschach Cognitive Triad measures ALL THREE steps in a single 30-min assessment"
    p5.font.size = Pt(14)
    p5.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
    p5.alignment = PP_ALIGN.CENTER

make_problem_slide()


# ============ SLIDE 4: REVIEW OF LITERATURE ============
make_content_slide(
    "Review of Literature: Key Findings (2016-2026)",
    [
        "1. Michalska da Rocha et al. (2018) - Meta-analysis: Significant positive relationship",
        "   between loneliness and psychosis. Called for cognitive mechanism research.",
        "",
        "2. Thibaudeau et al. (2023) - Loneliness linked to impaired mentalizing & emotion",
        "   recognition ONLY in patients with social-cognitive deficits.",
        "",
        "3. Hajduk et al. (2023) - Social Threat Bias directly predicts loneliness in SCZ.",
        "   Misperception of neutral as threatening -> withdrawal -> loneliness.",
        "",
        "4. Green et al. (2023) - 'Social homeostasis' model: Cognitive-perceptual impairment",
        "   is the primary mechanism of chronic social disconnection in SMI.",
        "",
        "5. Ilonen et al. (2012) - Rorschach psychological complexity predicted functional",
        "   capacity (r=.30-.35) BEYOND neurocognition & symptoms in schizophrenia.",
        "",
        "6. Singh et al. (2021) - Rorschach cognitive variables correlate with psychopathology",
        "   in Indian schizophrenia sample. No loneliness outcome examined."
    ],
    SECONDARY
)

# ============ SLIDE 5: MORE ROL ============
make_content_slide(
    "Review of Literature: Additional Evidence",
    [
        "7. Hsu et al. (2024) - Longitudinal: Social isolation -> poor language & memory",
        "   in schizophrenia inpatients (2-year follow-up, N=166).",
        "",
        "8. Wang et al. (2026) - Loneliness directly associated with cognitive impairment;",
        "   sleep quality and anxiety mediate the pathway.",
        "",
        "9. Regev et al. (2024) - Patients with high social exclusion were 4.24x more likely",
        "   to have cognitive impairment vs. low exclusion group.",
        "",
        "10. Jo et al. (2024) - Rorschach ideation variables (DV2, D score) differentiate",
        "    Kraepelinian vs. DSM schizophrenia severity.",
        "",
        "11. Lin et al. (2022) - UCLA Loneliness Scale V3 validated in schizophrenia.",
        "",
        "12. Kimoto et al. (2016) - Rorschach mediation (X-%, XA%) distinctively impaired",
        "    in schizophrenia vs. ASD: stronger perception distortions."
    ],
    PURPLE
)

# ============ SLIDE 6: RESEARCH GAP ============
def make_gap_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, LIGHT_BG, LIGHT_BG)
    
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Research Gap"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    add_accent_bar(slide, Inches(1.0), Inches(2.5), ACCENT)
    
    gaps = [
        "No study globally has linked Rorschach cognitive triad to loneliness",
        "Existing loneliness research uses computerized tasks unavailable in India",
        "Nobody has tested WHICH cognitive step predicts loneliness most",
        "No Indian data on loneliness & its cognitive correlates in schizophrenia",
        "Rorschach cognition studied for functioning but NEVER for subjective experience",
        "Ilonen (2012) used R-PAS, not Exner CS taught in Indian MPhil programs"
    ]
    
    for i, gap in enumerate(gaps):
        top = Inches(1.5 + i * 0.9)
        # Red X icon
        icon = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), top, Inches(0.45), Inches(0.45))
        icon.fill.solid()
        icon.fill.fore_color.rgb = ACCENT
        icon.line.fill.background()
        tx_i = slide.shapes.add_textbox(Inches(0.65), top + Pt(2), Inches(0.4), Inches(0.4))
        ti = tx_i.text_frame
        ti.paragraphs[0].text = "X"
        ti.paragraphs[0].font.size = Pt(14)
        ti.paragraphs[0].font.bold = True
        ti.paragraphs[0].font.color.rgb = WHITE
        ti.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Gap text
        tx = slide.shapes.add_textbox(Inches(1.3), top, Inches(11), Inches(0.7))
        t = tx.text_frame
        t.paragraphs[0].text = gap
        t.paragraphs[0].font.size = Pt(16)
        t.paragraphs[0].font.color.rgb = DARK_TEXT

make_gap_slide()


# ============ SLIDE 7: RATIONALE ============
make_content_slide(
    "Rationale of the Study",
    [
        "1. CLINICAL URGENCY: Loneliness predicts hospitalization, suicide risk,",
        "   cognitive decline & mortality in schizophrenia (Green et al., 2023; Yen et al., 2023)",
        "",
        "2. THEORETICAL BASIS: Cognitive Model of Loneliness (Cacioppo & Hawkley, 2009)",
        "   establishes that misperception of social environment CAUSES loneliness.",
        "",
        "3. PERFECT TOOL MATCH: Rorschach Cognitive Triad measures exactly the cognitive",
        "   operations (processing, perception accuracy, reasoning) implicated in loneliness.",
        "",
        "4. PRACTICAL ADVANTAGE: Single 30-min projective assessment replaces multi-hour",
        "   neuropsych batteries; available in every clinical psychology department in India.",
        "",
        "5. INTERVENTION IMPLICATION: Identifying WHICH cognitive step predicts loneliness",
        "   enables targeted cognitive remediation planning for lonely patients."
    ],
    GREEN
)

# ============ SLIDE 8: AIM & OBJECTIVES ============
def make_aim_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, WHITE, WHITE)
    
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.15), prs.slide_height)
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = SECONDARY
    stripe.line.fill.background()
    
    # Title
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Aim & Objectives"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    add_accent_bar(slide, Inches(1.0), Inches(2.5), SECONDARY)
    
    # Aim box
    aim_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(0.5), Inches(1.4), Inches(12), Inches(1.2))
    aim_box.fill.solid()
    aim_box.fill.fore_color.rgb = PRIMARY
    aim_box.line.fill.background()
    
    tx = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(1))
    t = tx.text_frame
    t.word_wrap = True
    p1 = t.paragraphs[0]
    p1.text = "AIM: To examine whether Rorschach indices of the Cognitive Triad (Processing,"
    p1.font.size = Pt(15)
    p1.font.color.rgb = WHITE
    p1.font.bold = True
    p2 = t.add_paragraph()
    p2.text = "Mediation, Ideation) predict subjective loneliness in patients with schizophrenia."
    p2.font.size = Pt(15)
    p2.font.color.rgb = WHITE
    p2.font.bold = True
    
    # Objectives
    objectives = [
        "O1: Assess Rorschach Cognitive Triad profile of schizophrenia patients",
        "O2: Assess level of subjective loneliness using UCLA Loneliness Scale V3",
        "O3: Examine correlation between Processing variables & loneliness",
        "O4: Examine correlation between Mediation variables & loneliness",
        "O5: Examine correlation between Ideation variables & loneliness",
        "O6: Determine which cluster is the STRONGEST predictor after controlling PANSS-N"
    ]
    
    txObj = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), Inches(11.5), Inches(4))
    tObj = txObj.text_frame
    tObj.word_wrap = True
    for i, obj in enumerate(objectives):
        p = tObj.add_paragraph() if i > 0 else tObj.paragraphs[0]
        p.text = obj
        p.font.size = Pt(14)
        p.font.color.rgb = DARK_TEXT
        p.space_after = Pt(8)

make_aim_slide()


# ============ SLIDE 9: HYPOTHESES ============
make_content_slide(
    "Hypotheses",
    [
        "H1: Higher Processing efficiency (normal Zd, higher DQ+) will predict",
        "    LESS loneliness -- patients who process efficiently connect better.",
        "",
        "H2: Greater Mediation distortion (higher X-%, lower XA%) will predict",
        "    MORE loneliness -- patients who misperceive reality feel more isolated.",
        "",
        "H3: Greater Ideation disturbance (higher WSum6, more M-) will predict",
        "    MORE loneliness -- disorganized thinking impairs social connection.",
        "",
        "H4: Cognitive Triad variables will predict loneliness ABOVE AND BEYOND",
        "    negative symptoms (PANSS-N) -- showing unique predictive value.",
        "",
        "H5: The MEDIATION cluster (reality testing) will be the STRONGEST predictor",
        "    of loneliness -- because misperception most directly disrupts connection."
    ],
    GOLD
)

# ============ SLIDE 10: METHODOLOGY ============
def make_methodology_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, WHITE, WHITE)
    
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.15), prs.slide_height)
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = GREEN
    stripe.line.fill.background()
    
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Methodology"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    add_accent_bar(slide, Inches(1.0), Inches(2.5), GREEN)
    
    # Method cards
    cards = [
        ("Research Design", "Cross-sectional\nCorrelational", SECONDARY),
        ("Sample", "N = 40-45\nSchizophrenia (ICD-10/11)\nClinically stable", GREEN),
        ("Sampling", "Purposive\nPsychiatry OPD\nFollow-up clinic", PURPLE),
        ("Duration", "4-5 months\nData collection\n+ Analysis", GOLD),
    ]
    
    for i, (title, desc, color) in enumerate(cards):
        left = Inches(0.5 + i * 3.2)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       left, Inches(1.5), Inches(2.9), Inches(2.5))
        card.fill.solid()
        card.fill.fore_color.rgb = color
        card.line.fill.background()
        
        tx = slide.shapes.add_textbox(left + Inches(0.2), Inches(1.6), Inches(2.5), Inches(2.3))
        t = tx.text_frame
        t.word_wrap = True
        p1 = t.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p1.alignment = PP_ALIGN.CENTER
        p2 = t.add_paragraph()
        p2.text = "\n" + desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = WHITE
        p2.alignment = PP_ALIGN.CENTER
    
    # Inclusion/Exclusion
    inc_exc = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), Inches(11.5), Inches(2.8))
    tf2 = inc_exc.text_frame
    tf2.word_wrap = True
    p_inc = tf2.paragraphs[0]
    p_inc.text = "Inclusion: Schizophrenia F20.x | Age 18-55 | Stable meds 8+ wks | PANSS <80 | R>=14"
    p_inc.font.size = Pt(13)
    p_inc.font.color.rgb = GREEN
    p_inc.font.bold = True
    
    p_exc = tf2.add_paragraph()
    p_exc.text = "Exclusion: Active substance use | ID (IQ<70) | Organic brain | ECT <3 months | Acute episode"
    p_exc.font.size = Pt(13)
    p_exc.font.color.rgb = ACCENT
    p_exc.font.bold = True

make_methodology_slide()


# ============ SLIDE 11: TOOLS ============
make_two_column_slide(
    "Tools",
    "Rorschach Inkblot Test (Exner CS)",
    [
        "",
        "* 10 inkblot cards, standard admin",
        "* Processing: Zf, Zd, DQ+, DQv, W:D:Dd",
        "* Mediation: XA%, WDA%, X-%, X+%, Xu%, P",
        "* Ideation: WSum6, Lv2, Ma:Mp, a:p, M quality",
        "* Inter-rater kappa: .85-.97",
        "* Indian norms: Dubey (2011)",
        "* Administration: 30-40 minutes"
    ],
    "UCLA Loneliness Scale V3 + PANSS-N",
    [
        "",
        "UCLA Loneliness Scale (Russell, 1996):",
        "* 20 items, 4-point Likert (20-80)",
        "* Alpha: .89-.94",
        "* Validated in SCZ (Lin et al., 2022)",
        "* Indian validation: Suri et al. (2020)",
        "",
        "PANSS-N (Kay et al., 1987):",
        "* 7-item Negative Subscale (covariate)",
        "* ICC: .83-.87"
    ]
)

# ============ SLIDE 12: CONCEPTUAL MODEL ============
def make_model_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY)
    
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.2), Inches(11), Inches(0.7))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Conceptual Model: How Cognitive Triad Predicts Loneliness"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = GOLD
    
    # Three cognitive boxes at top
    labels = [
        ("PROCESSING", "Zd, Zf, DQ+\nDQv, W:D:Dd", SECONDARY),
        ("MEDIATION", "XA%, WDA%\nX-%, P", ACCENT),
        ("IDEATION", "WSum6, Lv2\nMa:Mp, M quality", PURPLE)
    ]
    
    for i, (title, vars_text, color) in enumerate(labels):
        left = Inches(1 + i * 4)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      left, Inches(1.2), Inches(3.5), Inches(1.8))
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.fill.background()
        
        tx = slide.shapes.add_textbox(left + Inches(0.2), Inches(1.3), Inches(3.1), Inches(1.6))
        t = tx.text_frame
        t.word_wrap = True
        p1 = t.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p1.alignment = PP_ALIGN.CENTER
        p2 = t.add_paragraph()
        p2.text = vars_text
        p2.font.size = Pt(12)
        p2.font.color.rgb = WHITE
        p2.alignment = PP_ALIGN.CENTER
    
    # Arrows down
    for i in range(3):
        arr = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                                      Inches(2.5 + i * 4), Inches(3.1), Inches(0.5), Inches(0.6))
        arr.fill.solid()
        arr.fill.fore_color.rgb = GOLD
        arr.line.fill.background()
    
    # DV box
    dv_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(3.5), Inches(3.9), Inches(6), Inches(1.3))
    dv_box.fill.solid()
    dv_box.fill.fore_color.rgb = GOLD
    dv_box.line.fill.background()
    
    tx_dv = slide.shapes.add_textbox(Inches(3.7), Inches(4.0), Inches(5.6), Inches(1.1))
    t_dv = tx_dv.text_frame
    t_dv.word_wrap = True
    p_dv = t_dv.paragraphs[0]
    p_dv.text = "SUBJECTIVE LONELINESS"
    p_dv.font.size = Pt(18)
    p_dv.font.bold = True
    p_dv.font.color.rgb = PRIMARY
    p_dv.alignment = PP_ALIGN.CENTER
    p_dv2 = t_dv.add_paragraph()
    p_dv2.text = "(UCLA Loneliness Scale V3)"
    p_dv2.font.size = Pt(13)
    p_dv2.font.color.rgb = PRIMARY
    p_dv2.alignment = PP_ALIGN.CENTER
    
    # Covariate
    cov_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(4.5), Inches(5.6), Inches(4), Inches(0.9))
    cov_box.fill.solid()
    cov_box.fill.fore_color.rgb = RGBColor(0x55, 0x55, 0x77)
    cov_box.line.fill.background()
    
    tx_cov = slide.shapes.add_textbox(Inches(4.7), Inches(5.7), Inches(3.6), Inches(0.7))
    t_cov = tx_cov.text_frame
    p_cov = t_cov.paragraphs[0]
    p_cov.text = "Controlled for: PANSS-N (Negative Sx)"
    p_cov.font.size = Pt(12)
    p_cov.font.color.rgb = WHITE
    p_cov.alignment = PP_ALIGN.CENTER

make_model_slide()


# ============ SLIDE 13: DATA ANALYSIS ============
make_content_slide(
    "Data Analysis Plan",
    [
        "Step 1: Descriptive Statistics (Mean, SD, frequencies for all variables)",
        "",
        "Step 2: Normality Testing (Shapiro-Wilk) -> parametric vs. non-parametric",
        "",
        "Step 3: Inter-Rater Reliability (Cohen's Kappa/ICC on 20% of protocols)",
        "",
        "Step 4: Bivariate Correlations (Pearson/Spearman)",
        "        Each Rorschach variable x UCLA Loneliness Score",
        "",
        "Step 5: HIERARCHICAL MULTIPLE REGRESSION (Primary Analysis)",
        "        Block 1: PANSS-N (covariate)",
        "        Block 2: Processing composite (Zd, DQ+, Lambda)",
        "        Block 3: Mediation composite (XA%, X-%, P)",
        "        Block 4: Ideation composite (WSum6, M quality)",
        "        DV: UCLA Loneliness Score",
        "",
        "Key Output: Delta-R-squared at each step -> which cluster adds most power?"
    ],
    SECONDARY
)

# ============ SLIDE 14: SIGNIFICANCE ============
def make_significance_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, LIGHT_BG, LIGHT_BG)
    
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Clinical Significance & Contribution"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    add_accent_bar(slide, Inches(0.95), Inches(2.5), GREEN)
    
    items = [
        ("First study GLOBALLY", "linking Rorschach cognitive triad to loneliness in any population", SECONDARY),
        ("Identifies cognitive TARGETS", "for loneliness intervention in schizophrenia", GREEN),
        ("Performance-based assessment", "that cannot be faked (unlike self-report)", PURPLE),
        ("Practical for Indian settings", "single test, 30 min, universally available", GOLD),
        ("Informs rehabilitation", "which cognitive step to prioritize in cognitive remediation", ACCENT),
    ]
    
    for i, (title, desc, color) in enumerate(items):
        top = Inches(1.4 + i * 1.15)
        icon = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), top, Inches(0.5), Inches(0.5))
        icon.fill.solid()
        icon.fill.fore_color.rgb = color
        icon.line.fill.background()
        
        tx = slide.shapes.add_textbox(Inches(1.4), top, Inches(11), Inches(0.9))
        t = tx.text_frame
        t.word_wrap = True
        p1 = t.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = t.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(14)
        p2.font.color.rgb = DARK_TEXT

make_significance_slide()

# ============ SLIDE 15: REFERENCES ============
make_content_slide(
    "Key References",
    [
        "* Cacioppo, J.T. & Hawkley, L.C. (2009). Trends in Cognitive Sciences, 13(10), 447-454.",
        "* Exner, J.E. (2003). The Rorschach: A Comprehensive System (4th ed.). Wiley.",
        "* Green, M.F. et al. (2023). Schizophrenia Bulletin, 49(5), 1112-1126.",
        "* Hajduk, M. et al. (2023). Schizophrenia Research, 256, 38-46.",
        "* Ilonen, T. et al. (2012). Psychological Assessment, 25(1), 253-263.",
        "* Lin, C.Y. et al. (2022). Int J Environ Res Public Health, 19(14), 8443.",
        "* Michalska da Rocha, B. et al. (2018). Schizophrenia Bulletin, 44(1), 114-125.",
        "* Russell, D.W. (1996). J Personality Assessment, 66(1), 20-40.",
        "* Singh, G. et al. (2021). Indian J Psychiatric Nursing, 30(1), 49-55.",
        "* Thibaudeau, E. et al. (2023). Schizophrenia Research, 256, 29-37.",
        "* Wang, X. et al. (2026). European Archives of Psychiatry & Clinical Neuroscience.",
        "* Yen, C.F. et al. (2023). npj Schizophrenia, 9, Article 40."
    ],
    RGBColor(0x55, 0x55, 0x77)
)

# ============ SLIDE 16: THANK YOU ============
def make_thankyou_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY)
    
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.35), prs.slide_width, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = GOLD
    bar.line.fill.background()
    
    txBox = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(9), Inches(3))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Thank You"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = "\nQuestions & Suggestions Welcome"
    p2.font.size = Pt(22)
    p2.font.color.rgb = WHITE
    p2.alignment = PP_ALIGN.CENTER
    
    p3 = tf.add_paragraph()
    p3.text = "\n\n[Candidate Name]"
    p3.font.size = Pt(16)
    p3.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
    p3.alignment = PP_ALIGN.CENTER
    
    p4 = tf.add_paragraph()
    p4.text = "MPhil Clinical Psychology | [Institution] | 2026-2027"
    p4.font.size = Pt(14)
    p4.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
    p4.alignment = PP_ALIGN.CENTER

make_thankyou_slide()

# ============ SAVE ============
output_path = "/projects/sandbox/Dango-kiro/Research_Proposal_Rorschach_Loneliness_PPT.pptx"
prs.save(output_path)
print(f"Presentation saved to: {output_path}")
print(f"Total slides: {len(prs.slides)}")
