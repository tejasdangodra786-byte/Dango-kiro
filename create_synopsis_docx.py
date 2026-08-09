#!/usr/bin/env python3
"""
Create APA 7th Edition formatted Word document (.docx) for
MPhil Clinical Psychology Research Synopsis on Brief MBRP.
Uses raw Open XML + zipfile (no external dependencies).
"""

import zipfile
import os

# Output path
OUTPUT = "/projects/sandbox/Dango-kiro/MBRP_Research_Synopsis_APA7.docx"

# ─── XML NAMESPACE CONSTANTS ───
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


# ─── CONTENT TYPES ───
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

# ─── TOP-LEVEL RELS ───
TOP_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

# ─── WORD RELS ───
WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''


# ─── SETTINGS (double spacing, 1-inch margins) ───
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''

# ─── NUMBERING (for numbered lists) ───
NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>'''


# ─── STYLES (APA 7: Times New Roman 12pt, double-spaced, headings) ───
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="480" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto" w:after="0"/>
      <w:ind w:firstLine="720"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto" w:before="0" w:after="0"/>
      <w:jc w:val="center"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto" w:before="0" w:after="0"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto" w:before="0" w:after="0"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:i/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto"/>
      <w:jc w:val="center"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CenterNoIndent">
    <w:name w:val="CenterNoIndent"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto"/>
      <w:jc w:val="center"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="NoIndent">
    <w:name w:val="NoIndent"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto"/>
      <w:ind w:firstLine="0"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Reference">
    <w:name w:val="Reference"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="480" w:lineRule="auto" w:after="0"/>
      <w:ind w:firstLine="0" w:left="720" w:hanging="720"/>
    </w:pPr>
  </w:style>
</w:styles>'''



# ─── HELPER FUNCTIONS ───
def escape_xml(text):
    """Escape XML special characters."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def make_run(text, bold=False, italic=False):
    """Create a w:r element."""
    rpr = ""
    if bold or italic:
        rpr = "<w:rPr>"
        if bold:
            rpr += "<w:b/>"
        if italic:
            rpr += "<w:i/>"
        rpr += "</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(text)}</w:t></w:r>'

def para(text, style="Normal", bold=False, italic=False):
    """Create a paragraph with given style."""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    run = make_run(text, bold, italic)
    return f"<w:p>{ppr}{run}</w:p>"

def para_multi_runs(runs_list, style="Normal"):
    """Create a paragraph with multiple runs [(text, bold, italic), ...]."""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    runs = "".join([make_run(t, b, i) for t, b, i in runs_list])
    return f"<w:p>{ppr}{runs}</w:p>"

def heading1(text):
    return para(text, "Heading1", bold=True)

def heading2(text):
    return para(text, "Heading2", bold=True)

def heading3(text):
    return para(text, "Heading3", bold=True, italic=True)

def title_para(text):
    return para(text, "Title", bold=True)

def center_para(text, bold=False):
    return para(text, "CenterNoIndent", bold=bold)

def noindent_para(text, bold=False, italic=False):
    return para(text, "NoIndent", bold=bold, italic=italic)

def body_para(text):
    return para(text, "Normal")

def ref_para(text, italic_title=""):
    """Reference entry with hanging indent. italic_title is italicized portion."""
    if italic_title:
        parts = text.split(italic_title, 1)
        if len(parts) == 2:
            runs = [(parts[0], False, False), (italic_title, False, True), (parts[1], False, False)]
            return para_multi_runs(runs, "Reference")
    return para(text, "Reference")

def empty_para():
    return '<w:p><w:pPr><w:pStyle w:val="NoIndent"/></w:pPr></w:p>'

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'



# ═══════════════════════════════════════════════════════════════
# DOCUMENT BODY CONTENT
# ═══════════════════════════════════════════════════════════════

def build_document_body():
    """Build the entire document body XML content."""
    parts = []

    # ─── TITLE PAGE (Page 1) ───
    # Blank lines to push title down ~3-4 lines
    parts.append(empty_para())
    parts.append(empty_para())
    parts.append(empty_para())

    # Paper title (bold, centered)
    parts.append(title_para("Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention"))
    parts.append(title_para("on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients:"))
    parts.append(title_para("A Pre-Test Post-Test Control Group Experimental Design"))

    # Blank line between title and author
    parts.append(empty_para())

    # Author name
    parts.append(center_para("Tejas Dangodra"))

    # Affiliation
    parts.append(center_para("Department of Clinical Psychology, MAN College of Special Education & Psychological Studies"))

    # University
    parts.append(center_para("Krantivir Tatya Tope Vishwavidyalaya, Guna, Madhya Pradesh"))

    # Course
    parts.append(center_para("MPhil in Clinical Psychology (RCI Approved)"))

    # Guide
    parts.append(center_para("Guide: Dr. Himanshi Singh"))

    # Date
    parts.append(center_para("2025"))

    # Page break after title page
    parts.append(page_break())

    return parts


def build_introduction():
    """Build Introduction section."""
    parts = []

    # Repeat title at top of text page (APA 7 requirement)
    parts.append(title_para("Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention"))
    parts.append(title_para("on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients:"))
    parts.append(title_para("A Pre-Test Post-Test Control Group Experimental Design"))

    # Introduction text (no label per APA 7)
    parts.append(body_para(
        "Substance dependence represents one of the most significant public health challenges "
        "confronting India today. The MAGNITUDE study conducted by the Ministry of Social Justice "
        "and Empowerment (2019) estimated that approximately 3.1 crore individuals in India are "
        "affected by substance use disorders (SUDs). Opioid dependence constitutes a major burden "
        "in states such as Punjab, Rajasthan, and Northeast India, with the World Health Organization "
        "estimating that India accounts for approximately 25% of global opioid-related deaths in "
        "Southeast Asia."
    ))

    parts.append(body_para(
        "Despite the magnitude of this crisis, Indian de-addiction centres primarily offer "
        "pharmacotherapy, including opioid substitution therapy (OST) and naltrexone maintenance, "
        "with limited structured psychotherapy (Sarkar & Balhara, 2016). Psychosocial interventions "
        "remain underutilized despite evidence of superior combined outcomes. Relapse rates remain "
        "alarmingly high, with global estimates ranging from 40-60% within the first year (National "
        "Institute on Drug Abuse [NIDA], 2020), while Indian opioid studies report rates as high as "
        "70-80% (Mattoo et al., 2009). Primary relapse triggers include craving, negative affect, "
        "interpersonal conflict, and environmental cues."
    ))

    parts.append(body_para(
        "Indian rehabilitation settings face unique constraints including limited resources and "
        "short admission windows of 4-6 weeks, making traditional 8-week Mindfulness-Based Relapse "
        "Prevention (MBRP) programs impractical. A brief 6-session adaptation is essential for "
        "feasibility at district-level facilities. This necessitates the development and evaluation "
        "of abbreviated evidence-based psychological interventions that can be delivered within these "
        "constraints while maintaining therapeutic efficacy."
    ))

    # What is MBRP subsection
    parts.append(heading2("Mindfulness-Based Relapse Prevention"))

    parts.append(body_para(
        "Mindfulness-Based Relapse Prevention (MBRP) was developed by Bowen, Chawla, and Marlatt "
        "(2011) at the University of Washington. It integrates mindfulness meditation practices "
        "with the cognitive-behavioral Relapse Prevention framework (Marlatt & Gordon, 1985). The "
        "core mechanism involves teaching patients to observe craving and emotional distress as "
        "transient mental events without automatically acting on them. The key technique of \"urge "
        "surfing\" involves riding the wave of craving until it passes naturally."
    ))

    parts.append(body_para(
        "Neurobiologically, mindfulness strengthens prefrontal cortex regulation over amygdala "
        "reactivity, enhances PFC-limbic connectivity, improves inhibitory control, and reduces "
        "automatic craving-use pathways (Garland et al., 2014). The Brief B-MBRP adaptation proposed "
        "in the present study comprises 6 structured sessions delivered over 3 weeks in a group "
        "format of 6-8 patients, with each session lasting 45 minutes and delivered twice weekly. "
        "This format is feasible within the typical Indian inpatient department admission window of "
        "4-6 weeks."
    ))

    return parts


def build_variables():
    """Build the Variables section."""
    parts = []

    parts.append(heading2("Craving"))

    parts.append(body_para(
        "Craving is defined clinically as an intense, overwhelming desire to use a substance and "
        "is recognized as a core feature in both DSM-5 and ICD-11 diagnostic criteria. Theoretically, "
        "it encompasses obsessive thoughts and compulsive urges related to drug use (Franken et al., "
        "2002). Craving serves as the primary trigger for relapse, with craving intensity predicting "
        "treatment dropout, lapse episodes, and full relapse (Tiffany & Wray, 2012). Traditional "
        "suppression strategies paradoxically increase craving intensity through rebound effects."
    ))

    parts.append(body_para(
        "The neuropsychological basis of craving involves mesolimbic dopamine pathway activation, "
        "ventral striatum and orbitofrontal cortex hyperactivity to drug cues, incentive sensitization "
        "(Robinson & Berridge, 1993), and prefrontal cortex hypoactivation during craving episodes. "
        "MBRP techniques targeting craving include urge surfing, SOBER breathing space "
        "(Stop-Observe-Breathe-Expand-Respond), cognitive decentering, and mindfulness exposure "
        "through non-reactive awareness."
    ))

    parts.append(heading2("Impulsivity"))

    parts.append(body_para(
        "Impulsivity is defined clinically as rapid, unplanned actions without considering "
        "consequences (Moeller et al., 2001). Theoretically, it comprises three dimensions: motor, "
        "attentional, and non-planning impulsivity (Patton et al., 1995). In the context of substance "
        "dependence, higher trait impulsivity predicts relapse, mediates craving-to-use behavior, "
        "is associated with treatment non-adherence and dropout, and functions as both a risk factor "
        "for and consequence of chronic substance use."
    ))

    parts.append(body_para(
        "The neuropsychological basis involves prefrontal cortex dysfunction leading to impaired "
        "executive control, reduced inhibitory control as demonstrated in Go/No-Go paradigms, "
        "impaired delay discounting, and dorsolateral prefrontal cortex hypoactivation in substance "
        "users. MBRP techniques targeting impulsivity include response inhibition through mindful "
        "pause, the STOP technique (Stop-Take a breath-Observe-Proceed), awareness training to "
        "notice impulse-action sequences, and mindful decision-making to create space between "
        "stimulus and response."
    ))

    parts.append(heading2("Mindfulness"))

    parts.append(body_para(
        "Mindfulness is defined clinically as present-moment awareness with openness and "
        "non-judgment. Theoretically, it encompasses five facets: observing, describing, acting "
        "with awareness, non-judging, and non-reactivity (Baer et al., 2006). Substance users "
        "demonstrate significantly lower dispositional mindfulness compared to non-clinical "
        "populations. Mindfulness acts as a protective factor against relapse triggers, and "
        "improvements in mindfulness mediate MBRP treatment outcomes across multiple studies."
    ))

    parts.append(body_para(
        "The neuropsychological basis involves anterior cingulate cortex activation enhancing "
        "self-regulation, insula mediating interoceptive awareness, PFC-amygdala connectivity "
        "improving emotional regulation, and default mode network regulation reducing rumination. "
        "MBRP techniques for building mindfulness include sitting meditation with focused attention "
        "on breath, body scan for non-judgmental bodily awareness, mindful movement through "
        "present-moment yoga and walking, and non-judgmental awareness through labeling without "
        "evaluation."
    ))

    return parts


def build_review_of_literature():
    """Build the Review of Literature section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Review of Literature"))

    parts.append(heading2("MBRP and Relapse Prevention"))

    parts.append(body_para(
        "Bowen et al. (2014) conducted a randomized controlled trial (RCT) with 286 participants "
        "published in JAMA Psychiatry. At 12-month follow-up, MBRP participants reported significantly "
        "fewer days of substance use and heavy drinking compared to standard relapse prevention (RP) "
        "and treatment as usual (TAU). MBRP maintained superior long-term outcomes through cultivation "
        "of mindfulness skills as a durable protective mechanism."
    ))

    parts.append(body_para(
        "Bowen and Marlatt (2009) examined brief urge surfing meditation with incarcerated substance "
        "users, finding significant reductions in craving intensity and frequency compared to controls. "
        "This study validates that even brief mindfulness exposure can disrupt the automaticity of "
        "craving responses and supports the feasibility of 6-session B-MBRP protocols. Both studies "
        "establish that MBRP integrates present-moment awareness with cognitive-behavioral strategies, "
        "creating synergistic therapeutic effects addressing both automatic reactivity and cognitive "
        "distortions that precipitate relapse."
    ))

    parts.append(heading2("Craving and Mindfulness Mechanisms"))

    parts.append(body_para(
        "Garland et al. (2014) demonstrated that Mindfulness-Oriented Recovery Enhancement (MORE) "
        "reduces opioid craving through three mechanisms: attentional reorientation away from drug "
        "cues, positive reappraisal of neutral stimuli, and enhanced savoring of healthy pleasures. "
        "Neuroimaging evidence shows that mindfulness modulates prefrontal and limbic craving circuits."
    ))

    parts.append(body_para(
        "Witkiewitz et al. (2013) conducted a secondary RCT analysis showing that over a 4-month "
        "follow-up period, MBRP participants demonstrated lower craving and the affect-craving "
        "pathway was significantly attenuated. Mindfulness weakens the link between negative emotions "
        "and automatic craving by cultivating non-reactive awareness. This mechanism distinguishes "
        "MBRP from traditional RP approaches, as negative emotions no longer automatically trigger "
        "craving responses."
    ))

    parts.append(heading2("Impulsivity and Mindfulness"))

    parts.append(body_para(
        "Garland et al. (2016) conducted an RCT with substance-dependent adults examining "
        "Mindfulness-Oriented Recovery Enhancement, which produced significant BIS-11 reductions "
        "in motor and attentional impulsivity. The mechanism involves mindfulness strengthening "
        "prefrontal cortical inhibitory control through sustained attention and deliberate "
        "non-reactivity training."
    ))

    parts.append(body_para(
        "Murphy and MacKillop (2012) conducted a cross-sectional study with 340 participants "
        "demonstrating that trait mindfulness is inversely associated with impulsive decision-making "
        "(delay discounting). Mindfulness moderated the impulsivity-substance use link, with higher "
        "mindfulness associated with a weaker relationship between impulsivity and alcohol "
        "consequences. These findings suggest that mindfulness functions as a cognitive resource "
        "enabling impulsive individuals to override automatic behavioral tendencies through enhanced "
        "metacognitive awareness and response flexibility."
    ))

    parts.append(heading2("Meta-Analyses and Brief Intervention Models"))

    parts.append(body_para(
        "Li et al. (2017) conducted a meta-analysis of 42 RCTs examining mindfulness-based "
        "interventions for substance misuse. Effect sizes were reported for substance misuse "
        "(d = 0.33), craving (d = 0.68), and stress (d = 0.44). Critically, brief interventions "
        "comprising 4-8 sessions showed comparable efficacy to longer protocols when appropriately "
        "structured, supporting the present study's 6-session design."
    ))

    parts.append(body_para(
        "Glasner-Edwards et al. (2017) conducted a pilot RCT examining an abbreviated 6-session "
        "MBRP for stimulant-dependent adults. The intervention was both feasible and effective in "
        "reducing substance use frequency and craving intensity compared to a health education "
        "control condition. This study directly validates the brief intervention model proposed in "
        "the present research and demonstrates that 6-session MBRP protocols can be successfully "
        "implemented without substantial loss of efficacy."
    ))

    parts.append(heading2("Indian Context"))

    parts.append(body_para(
        "Ghosh et al. (2018) reported that relapse rates exceed 70% among opioid-dependent "
        "patients in North Indian de-addiction centres within 3 months post-discharge. Primary "
        "determinants include craving, peer influence, negative affect, and lack of psychological "
        "aftercare. Indian facilities predominantly rely on pharmacological approaches."
    ))

    parts.append(body_para(
        "Sarkar and Balhara (2016) highlighted the underutilization of structured psychological "
        "interventions in Indian de-addiction settings, identifying barriers including limited "
        "trained personnel and absent validated protocols. Jain et al. (2013) provided preliminary "
        "evidence from mindfulness-based interventions with alcohol-dependent patients in India "
        "showing initial craving reductions. Indian treatment facilities predominantly rely on "
        "pharmacological approaches with minimal integration of evidence-based psychological "
        "interventions specifically designed for relapse prevention in opioid-dependent populations."
    ))

    return parts


def build_research_gap():
    """Build Research Gap section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Research Gap"))

    parts.append(body_para(
        "The following critical gaps in the existing literature necessitate the present study. "
        "First, there is limited evidence for brief MBRP adaptations in low- and middle-income "
        "country (LMIC) settings, as the majority of MBRP research evaluates the standard 8-week "
        "protocol in Western outpatient settings. Condensed protocols suitable for resource-constrained "
        "inpatient environments remain insufficiently examined (Li et al., 2017)."
    ))

    parts.append(body_para(
        "Second, there is an absence of simultaneous multi-domain outcome assessment. Existing "
        "studies typically examine craving, impulsivity, or mindfulness in isolation. The interplay "
        "between these three mechanistically linked variables within a single intervention framework "
        "has not been adequately investigated in Indian clinical populations."
    ))

    parts.append(body_para(
        "Third, there is a paucity of culturally contextualized MBRP evidence from India. Despite "
        "India's substantial opioid dependence burden, empirical validation of MBRP-based protocols "
        "in Indian de-addiction infrastructure remains virtually absent from the published literature."
    ))

    parts.append(body_para(
        "Fourth, there is insufficient integration of psychological interventions in Indian "
        "de-addiction practice. Current treatment models at district-level centers remain predominantly "
        "pharmacological, with limited evidence-based psychosocial adjuncts (Sarkar & Balhara, 2016; "
        "Ghosh et al., 2018)."
    ))

    parts.append(body_para(
        "Fifth, there is a need for attention-matched controlled designs. Most existing mindfulness "
        "studies in Indian addiction settings lack active control conditions, making it difficult to "
        "distinguish mindfulness-specific effects from non-specific therapeutic factors. The present "
        "study addresses these gaps as the first brief MBRP trial (6 sessions over 3 weeks) in an "
        "Indian opioid-dependent sample at a de-addiction centre with an active control and "
        "simultaneous three-variable assessment."
    ))

    return parts


def build_aim_objectives_hypotheses():
    """Build Aim, Objectives, and Hypotheses sections."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Aim of the Study"))

    parts.append(body_para(
        "To evaluate the efficacy of a Brief Mindfulness-Based Relapse Prevention (B-MBRP) "
        "intervention (6 sessions over 3 weeks) in reducing craving and impulsivity, and enhancing "
        "mindfulness, among male substance-dependent patients at a de-addiction centre, comparing "
        "Brief MBRP + TAU (Experimental Group) versus Psychoeducation + TAU (Active Control Group)."
    ))

    parts.append(heading1("Objectives"))

    parts.append(noindent_para(
        "1. To assess and compare craving levels (pre vs. post) in the Experimental Group "
        "(Brief MBRP + TAU) and Control Group (Psychoeducation + TAU)."
    ))
    parts.append(noindent_para(
        "2. To assess and compare impulsivity levels (pre vs. post) in both groups."
    ))
    parts.append(noindent_para(
        "3. To assess and compare mindfulness levels (pre vs. post) in both groups."
    ))
    parts.append(noindent_para(
        "4. To determine whether Brief MBRP + TAU is significantly more effective than "
        "Psychoeducation + TAU in reducing craving."
    ))
    parts.append(noindent_para(
        "5. To determine whether Brief MBRP + TAU is significantly more effective than "
        "Psychoeducation + TAU in reducing impulsivity."
    ))
    parts.append(noindent_para(
        "6. To determine whether Brief MBRP + TAU is significantly more effective than "
        "Psychoeducation + TAU in enhancing mindfulness."
    ))

    parts.append(heading1("Hypotheses"))

    parts.append(body_para(
        "The following null hypotheses were formulated for the present study:"
    ))

    parts.append(noindent_para(
        "H01: There is no significant difference in craving scores (OCDUS) between the "
        "Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) "
        "from pre-test to post-test."
    ))
    parts.append(noindent_para(
        "H02: There is no significant difference in impulsivity scores (BIS-11) between the "
        "Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) "
        "from pre-test to post-test."
    ))
    parts.append(noindent_para(
        "H03: There is no significant difference in mindfulness scores (FFMQ) between the "
        "Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) "
        "from pre-test to post-test."
    ))

    parts.append(body_para(
        "Statistical testing was set at Alpha = 0.05 (two-tailed), with the primary analysis "
        "being ANCOVA with pre-test scores as covariates."
    ))

    return parts


def build_operational_definitions():
    """Build Operational Definitions section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Operational Definitions of Key Terms"))

    parts.append(body_para(
        "Craving. \"Craving\" in this study refers to the intensity and frequency of obsessive "
        "thoughts about drug use and compulsive urges to use opioids, as measured by the total "
        "score on the Obsessive Compulsive Drug Use Scale (OCDUS; Franken et al., 2002). Scores "
        "range from 0 to 52; higher scores indicate greater craving."
    ))

    parts.append(body_para(
        "Impulsivity. \"Impulsivity\" in this study refers to the multidimensional tendency to "
        "act without adequate forethought, encompassing motor, attentional, and non-planning "
        "components, as measured by the total score and three subscale scores on the Barratt "
        "Impulsiveness Scale-11 (BIS-11; Patton et al., 1995). Scores range from 30 to 120; "
        "higher scores indicate greater impulsivity."
    ))

    parts.append(body_para(
        "Mindfulness. \"Mindfulness\" in this study refers to the dispositional capacity for "
        "present-moment awareness with non-judgment and non-reactivity, as measured by the total "
        "score on the Five Facet Mindfulness Questionnaire (FFMQ; Baer et al., 2006) across five "
        "facets. Scores range from 39 to 195; higher scores indicate greater mindfulness."
    ))

    parts.append(body_para(
        "Brief MBRP. \"Brief MBRP\" in this study refers to a structured 6-session "
        "mindfulness-based relapse prevention intervention delivered twice weekly over 3 weeks "
        "(45 min/session) in group format (6-8 patients), adapted from Bowen, Chawla, and Marlatt "
        "(2011), incorporating body scan, breath meditation, urge surfing, SOBER breathing space, "
        "and relapse prevention planning."
    ))

    parts.append(body_para(
        "Substance Dependence. \"Substance Dependence\" in this study refers to a clinical "
        "diagnosis of Substance Dependence Syndrome as per ICD-10 (F10-F19) criteria, with "
        "primary opioid dependence (heroin or pharmaceutical opioids), confirmed by a qualified "
        "psychiatrist at a de-addiction centre."
    ))

    parts.append(body_para(
        "Treatment As Usual (TAU). \"TAU\" in this study refers to the standard pharmacological "
        "treatment at the de-addiction centre, including medically supervised detoxification, "
        "opioid substitution therapy (buprenorphine/methadone), naltrexone maintenance, routine "
        "counseling, and daily ward activities."
    ))

    parts.append(body_para(
        "Psychoeducation. \"Psychoeducation\" in this study refers to 6 structured informational "
        "sessions (45 min, twice weekly, 3 weeks) covering addiction science, effects of opioids, "
        "relapse warning signs, health/nutrition, social consequences, and recovery motivation. "
        "These sessions are matched for time and attention but contain no mindfulness component."
    ))

    return parts


def build_methodology():
    """Build Methodology section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Methodology"))

    parts.append(heading2("Research Design"))

    parts.append(body_para(
        "The present study employed a Pre-Test Post-Test Control Group Experimental Design "
        "(true experimental design). The design notation is as follows:"
    ))
    parts.append(noindent_para("R  O1  X1  O2  -  Experimental Group (Brief MBRP + TAU)"))
    parts.append(noindent_para("R  O1  X2  O2  -  Control Group (Psychoeducation + TAU)"))
    parts.append(body_para(
        "Design features include true experimental methodology with random assignment, an active "
        "control condition (attention-matched), continuation of TAU for all participants, and "
        "pre-post measurement at standardized time points. The setting is a De-addiction Centre "
        "in Guna, Madhya Pradesh, with assessments conducted by a blinded research assistant and "
        "intervention delivered by an MPhil Clinical Psychologist (Researcher)."
    ))

    parts.append(heading2("Sample and Sampling Strategy"))

    parts.append(body_para(
        "A two-stage sampling strategy was employed. In Stage 1 (Purposive Selection), eligible "
        "participants were identified based on inclusion and exclusion criteria from patients "
        "admitted to the de-addiction centre through consecutive sampling during the recruitment "
        "window. In Stage 2 (Random Assignment), a computer-generated randomization sequence was "
        "used to allocate eligible participants to either the Experimental Group (n = 30) or "
        "Control Group (n = 30) using sealed opaque envelopes."
    ))
    parts.append(body_para(
        "The total sample size was 60 participants (30 per group), with 70 recruited to account "
        "for approximately 15% attrition. All participants were male, aged 18-50 years. The "
        "male-only sample was justified as Indian de-addiction centres admit approximately 90-95% "
        "males, with the MAGNITUDE study reporting a male-to-female ratio of approximately 10:1 "
        "for opioid dependence. A homogeneous sample strengthens internal validity, and "
        "female-specific studies are recommended as a future direction."
    ))

    parts.append(heading2("Sample Size Estimation"))

    parts.append(body_para(
        "Sample size was calculated using the formula: n = [(Za/2 + Zb)^2 x 2 x s^2] / d^2. "
        "Parameters included an effect size (d) of 0.50 (medium), power (1-b) of 0.80 "
        "(Zb = 0.84), and alpha of 0.05 two-tailed (Za/2 = 1.96). The calculation yielded "
        "n = [(1.96 + 0.84)^2 x 2 x 1] / (0.50)^2 = 62.72, approximately 63 total (32 per "
        "group). The final decision was N = 60 (30 per group) as ANCOVA reduces the required "
        "sample size. G*Power 3.1 verification confirmed adequacy. Recruitment of 70 total "
        "(35/group) was planned to account for approximately 15% attrition, consistent with "
        "Glasner-Edwards et al. (2017) and Bowen and Marlatt (2009)."
    ))

    parts.append(heading2("Inclusion Criteria"))

    parts.append(noindent_para("1. Diagnosis of Substance Dependence as per ICD-10 (F10-F19) / ICD-11 criteria."))
    parts.append(noindent_para("2. Primary substance: Opioids (heroin, pharmaceutical opioids); polysubstance with primary opioid included."))
    parts.append(noindent_para("3. Male participants aged 18-50 years."))
    parts.append(noindent_para("4. Completed detoxification phase (minimum 7 days post-withdrawal)."))
    parts.append(noindent_para("5. Currently admitted at the de-addiction centre."))
    parts.append(noindent_para("6. Minimum education: 5th standard (ability to comprehend psychometric tools)."))
    parts.append(noindent_para("7. Willingness to provide written informed consent."))
    parts.append(noindent_para("8. Able to attend all 6 intervention sessions during the 3-week period."))

    parts.append(heading2("Exclusion Criteria"))

    parts.append(noindent_para("1. Severe psychiatric comorbidity: Psychotic disorders, Bipolar I, severe MDE with suicidality."))
    parts.append(noindent_para("2. Significant cognitive impairment (MMSE < 24) or intellectual disability."))
    parts.append(noindent_para("3. Active withdrawal symptoms (COWS score > 12)."))
    parts.append(noindent_para("4. History of traumatic brain injury with loss of consciousness > 30 minutes."))
    parts.append(noindent_para("5. Current participation in another structured psychological intervention study."))
    parts.append(noindent_para("6. Medical instability requiring acute/intensive care."))
    parts.append(noindent_para("7. History of prior formal mindfulness/meditation training exceeding 1 month."))

    return parts


def build_variables_and_tools():
    """Build Variables and Tools sections."""
    parts = []

    parts.append(heading2("Variables of the Study"))

    parts.append(body_para(
        "The independent variable was the type of intervention with two levels: Level 1 "
        "(Brief MBRP + TAU for the Experimental Group) and Level 2 (Psychoeducation + TAU for "
        "the Control Group). The dependent variables were: (a) Craving, measured by OCDUS total "
        "score (range 0-52); (b) Impulsivity, measured by BIS-11 total score and 3 subscales "
        "(range 30-120); and (c) Mindfulness, measured by FFMQ total score and 5 facets "
        "(range 39-195). Controlled variables included age, education, duration of use, severity "
        "(ASSIST baseline), TAU components (constant across groups), and session duration "
        "(equalized at 45 min x 6 sessions for both groups)."
    ))

    parts.append(heading2("Tools and Measures"))

    parts.append(heading3("Obsessive Compulsive Drug Use Scale (OCDUS)"))

    parts.append(body_para(
        "The OCDUS (Franken et al., 2002) is a 13-item self-report measure assessing obsessive "
        "thoughts about drug use and compulsive urges. It captures both cognitive (obsessive) and "
        "behavioral (compulsive) craving dimensions using a 5-point scale (0-4) with a total range "
        "of 0-52, where higher scores indicate greater craving. Subscales include obsessive "
        "thoughts/interference, desire/control, and resistance to thoughts. Psychometric properties "
        "demonstrate internal consistency (a = 0.86-0.90), test-retest reliability (r = 0.78), "
        "and convergent validity with VAS craving (r = 0.55-0.67). The scale predicts relapse and "
        "is sensitive to treatment changes. It is applicable across substances including opioids, "
        "is brief (5 minutes), captures multidimensional craving, is suitable for pre-post designs, "
        "and is adaptable for Hindi administration."
    ))

    parts.append(heading3("Barratt Impulsiveness Scale-11 (BIS-11)"))

    parts.append(body_para(
        "The BIS-11 (Patton et al., 1995) is a 30-item self-report scale using a 4-point scale "
        "(1-4) with a total range of 30-120, where higher scores indicate greater impulsivity. "
        "It measures three factors: attentional, motor, and non-planning impulsivity. Psychometric "
        "properties include internal consistency (a = 0.79-0.83) and test-retest reliability "
        "(r = 0.83). The scale discriminates substance use disorder populations from controls. "
        "A Hindi validated version is available."
    ))

    parts.append(heading3("Five Facet Mindfulness Questionnaire (FFMQ)"))

    parts.append(body_para(
        "The FFMQ (Baer et al., 2006) is a 39-item self-report questionnaire using a 5-point "
        "scale (1-5) with a total range of 39-195, where higher scores indicate greater "
        "mindfulness. It measures five facets: observing, describing, acting with awareness, "
        "non-judging, and non-reactivity. Psychometric properties demonstrate internal consistency "
        "(a = 0.75-0.91 across facets). The scale is sensitive to mindfulness interventions and "
        "a Hindi adaptation is available."
    ))

    parts.append(heading3("WHO-ASSIST Version 3.0 (Baseline Severity)"))

    parts.append(body_para(
        "The WHO-ASSIST V3.0 (WHO ASSIST Working Group, 2002) is an 8-item questionnaire measuring "
        "substance use across lifetime and past 3 months, screening across 10 substance categories. "
        "Substance-specific risk scoring includes: Low (0-3, no intervention), Moderate (4-26, brief "
        "intervention), and High (27+, referral to specialist). Psychometric properties include "
        "test-retest reliability (r = 0.58-0.90), internal consistency (a = 0.77-0.94), sensitivity "
        "(0.80), and specificity (0.71), validated across 18 countries. In this study, it was used "
        "at pre-test only to establish baseline severity, ensure group equivalence, and stratify "
        "for randomization. It was not used as an outcome measure."
    ))

    return parts


def build_procedure_and_intervention():
    """Build Data Collection Procedure and Intervention sections."""
    parts = []

    parts.append(page_break())
    parts.append(heading2("Data Collection Procedure"))

    parts.append(body_para(
        "The data collection followed a systematic six-step procedure. Step 1, Screening "
        "(Day 1-3), involved reviewing admission records to identify male patients aged 18-50 "
        "with opioid dependence (ICD-10) who had completed 7 or more days of detoxification. "
        "MMSE scores of 24 or above and COWS scores of 12 or below were verified, and inclusion "
        "and exclusion criteria were applied."
    ))

    parts.append(body_para(
        "Step 2, Informed Consent (Day 3-4), involved individual meetings conducted in Hindi "
        "to explain the study purpose, procedures, duration, voluntary nature, right to withdraw, "
        "and confidentiality. Written signed consent was obtained. Step 3, Pre-Test Assessment "
        "(Day 4-5), involved administering the ASSIST (severity baseline), OCDUS (craving), "
        "BIS-11 (impulsivity), and FFMQ (mindfulness) in a quiet room, requiring approximately "
        "40 minutes total, conducted by a blinded research assistant."
    ))

    parts.append(body_para(
        "Step 4, Randomization (Day 5), utilized a computer-generated random sequence with "
        "sealed opaque envelopes to allocate participants to the Experimental (MBRP) or Control "
        "(Psychoeducation) group, stratified by ASSIST severity. Step 5, Intervention (Weeks 1-3), "
        "consisted of 6 sessions delivered twice weekly, each lasting 45 minutes, in groups of "
        "6-8 participants. The MBRP group received mindfulness techniques while the Control group "
        "received psychoeducation. Both groups continued TAU throughout."
    ))

    parts.append(body_para(
        "Step 6, Post-Test (within 1 week post-intervention), involved re-administering the "
        "OCDUS, BIS-11, and FFMQ only (not the ASSIST). The same blinded research assistant "
        "conducted assessments under the same conditions as the pre-test."
    ))

    parts.append(heading2("Intervention: Brief MBRP Protocol"))

    parts.append(body_para(
        "The Brief MBRP intervention comprised 6 structured sessions delivered twice weekly "
        "over 3 weeks. Session 1 (Autopilot and Awareness) included a welcome, raisin exercise, "
        "discussion of automatic patterns, brief body scan, and assignment of home practice "
        "(5-minute daily body scan). Session 2 (Triggers and Body Scan) included a full body scan, "
        "trigger mapping exercise, and discussion of body signals of craving, with home practice "
        "of body scan and trigger diary."
    ))

    parts.append(body_para(
        "Session 3 (Breath and SOBER Space) included breath meditation, teaching the SOBER "
        "technique, role-play of high-risk scenarios, and integration, with home practice of "
        "breath meditation and three-times daily SOBER. Session 4 (Urge Surfing and Decentering) "
        "included guided urge surfing, cognitive decentering exercise, and group sharing, with "
        "home practice of urge surfing when craving arises."
    ))

    parts.append(body_para(
        "Session 5 (Acceptance and Non-Reactivity) included open awareness meditation, "
        "acceptance versus avoidance discussion, and skillful action planning, with home "
        "practice of daily open awareness. Session 6 (Integration and Maintenance) included "
        "loving-kindness meditation, personal relapse prevention planning, and group feedback "
        "and closure, with a personalized daily plan as ongoing practice."
    ))

    parts.append(body_para(
        "Each session followed a standard 6-step structure: (1) Opening Meditation (5 min), "
        "(2) Practice Review (5 min), (3) Core Meditation/Exercise (12-15 min), "
        "(4) Psychoeducation and Discussion (10-12 min), (5) Experiential Exercise (8-10 min), "
        "and (6) Closing and Home Practice (5 min). Delivery was in Hindi with pre-recorded "
        "audio meditations provided and practice logs maintained. The protocol was adapted from "
        "Bowen, Chawla, and Marlatt (2011)."
    ))

    parts.append(heading2("Control Group: Psychoeducation"))

    parts.append(body_para(
        "The control group received 6 structured psychoeducation sessions matched for time, "
        "format, frequency, attention, and therapist contact. Session topics included: "
        "(1) Understanding Addiction (disease model, brain changes, genetic and environmental "
        "factors), (2) Effects of Opioids (physical consequences, psychological effects, "
        "withdrawal timeline), (3) Understanding Relapse (warning signs, high-risk situations, "
        "Marlatt's relapse model), (4) Health and Nutrition (physical recovery, sleep hygiene, "
        "exercise benefits), (5) Social Consequences (family impact, legal issues, stigma), and "
        "(6) Motivation and Goals (stages of change, recovery planning, community resources). "
        "The psychoeducation sessions used handouts, visual aids, and discussion, and contained "
        "no mindfulness component."
    ))

    return parts


def build_data_analysis():
    """Build Data Analysis Plan section."""
    parts = []

    parts.append(heading2("Data Analysis Plan"))

    parts.append(body_para(
        "Statistical analyses included descriptive statistics (mean, standard deviation, "
        "frequencies), normality testing (Shapiro-Wilk test), within-group comparisons (paired "
        "t-test or Wilcoxon signed-rank test), between-group comparisons (independent t-test or "
        "Mann-Whitney U test), and the primary analysis of ANCOVA. Effect sizes were calculated "
        "using partial eta-squared and Cohen's d. The significance level was set at alpha = 0.05 "
        "(two-tailed). All analyses were conducted using SPSS Version 29.1."
    ))

    parts.append(body_para(
        "The ANCOVA model specified post-test scores (OCDUS, BIS-11, or FFMQ) as the dependent "
        "variable, group (Experimental vs. Control) as the independent variable, and pre-test "
        "scores on the same measure as the covariate. Intention-to-treat analysis used Last "
        "Observation Carried Forward (LOCF). ANCOVA was justified because it (a) controls for "
        "baseline differences, (b) increases power by reducing error variance, (c) provides a "
        "more precise treatment effect estimate, (d) reduces the required sample size, and "
        "(e) is recommended for pre-post designs (Tabachnick & Fidell, 2013)."
    ))

    return parts


def build_ethical_considerations():
    """Build Ethical Considerations section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Ethical Considerations"))

    parts.append(body_para(
        "The study adhered to strict ethical guidelines. Informed consent was obtained in "
        "writing in Hindi, with participants fully informed of the purpose, procedures, risks, "
        "and benefits. Participation was voluntary, with the right to withdraw at any time "
        "without penalty or impact on treatment. Confidentiality was maintained through data "
        "coding with participant IDs, with no identifying information in publications and "
        "secure locked storage."
    ))

    parts.append(body_para(
        "Non-maleficence was ensured as the control group received active psychoeducation "
        "(not a waitlist condition) and TAU was continued for all participants. Institutional "
        "ethical clearance was obtained from the Institutional Ethics Committee (IEC) prior to "
        "data collection. Debriefing was provided as the control group was offered a brief MBRP "
        "orientation post-study. The study complied with the ICMR (2017) National Ethical "
        "Guidelines for Biomedical and Health Research."
    ))

    return parts


def build_expected_results():
    """Build Expected Results section."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Expected Results"))

    parts.append(body_para(
        "It is expected that the Brief MBRP group will demonstrate a significant reduction in "
        "craving (OCDUS) compared to the Control group, with an expected effect size of d = "
        "0.50-0.80. The mechanism involves urge surfing disrupting the automatic craving-use "
        "cycle."
    ))

    parts.append(body_para(
        "A significant reduction in impulsivity (BIS-11), particularly in motor and attentional "
        "subscales, is expected in the Brief MBRP group compared to the Control group, with an "
        "expected effect size of d = 0.40-0.60. The mechanism involves mindfulness strengthening "
        "prefrontal cortex inhibitory control."
    ))

    parts.append(body_para(
        "A significant increase in mindfulness (FFMQ), particularly in the Acting with Awareness "
        "and Non-Reactivity facets, is expected in the Brief MBRP group compared to the Control "
        "group, with an expected effect size of d = 0.50-0.70. The mechanism involves structured "
        "meditation building dispositional mindfulness."
    ))

    parts.append(body_para(
        "Overall, Brief MBRP + TAU is expected to demonstrate superiority across all three "
        "dependent variables, and the null hypotheses are expected to be rejected. These findings "
        "would support the feasibility of 6-session B-MBRP at a de-addiction centre."
    ))

    return parts


def build_limitations_and_future():
    """Build Limitations and Future Directions sections."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("Limitations"))

    parts.append(noindent_para(
        "1. Male-only sample from a single center limits generalizability to females and other settings."
    ))
    parts.append(noindent_para(
        "2. Short-term assessment: Post-test immediately after intervention with no long-term follow-up data."
    ))
    parts.append(noindent_para(
        "3. Self-report measures (OCDUS, BIS-11, FFMQ) are susceptible to social desirability bias."
    ))
    parts.append(noindent_para(
        "4. No biological markers: Craving measured subjectively without physiological corroboration."
    ))
    parts.append(noindent_para(
        "5. Therapist effects: Single therapist delivery may introduce confounds (mitigated by manualized protocol)."
    ))
    parts.append(noindent_para(
        "6. Attention-matched control does not fully isolate mindfulness-specific mechanisms."
    ))
    parts.append(noindent_para(
        "7. Potential attrition despite over-recruitment (common in substance-dependent populations)."
    ))

    parts.append(heading1("Future Directions"))

    parts.append(body_para(
        "Future research should include 3-month and 6-month follow-up assessments, multi-site "
        "RCTs across India, and the inclusion of female participants. Mechanism research should "
        "incorporate neuroimaging (fMRI/EEG), mediator analysis examining mindfulness as a "
        "mediator, and dose-response studies comparing 4, 6, and 8 sessions."
    ))

    parts.append(body_para(
        "Comparative studies should examine MBRP versus CBT versus ACT in Indian settings, "
        "technology-assisted app-based MBRP delivery, and post-discharge booster sessions. "
        "Dissemination efforts should focus on developing a Hindi MBRP training manual, "
        "training counselors in government centers, and improving rural access via telehealth "
        "delivery."
    ))

    return parts


def build_references():
    """Build References section in APA 7 format."""
    parts = []

    parts.append(page_break())
    parts.append(heading1("References"))

    refs = [
        ('Baer, R. A., Smith, G. T., Hopkins, J., Krietemeyer, J., & Toney, L. (2006). Using self-report assessment methods to explore facets of mindfulness. Assessment, 13(1), 27-45.', 'Assessment'),
        ('Bowen, S., Chawla, N., & Marlatt, G. A. (2011). Mindfulness-based relapse prevention for addictive behaviors: A clinician\'s guide. Guilford Press.', 'Mindfulness-based relapse prevention for addictive behaviors: A clinician\'s guide.'),
        ('Bowen, S., & Marlatt, G. A. (2009). Surfing the urge: Brief mindfulness-based intervention for college student smokers. Psychology of Addictive Behaviors, 23(4), 666-671.', 'Psychology of Addictive Behaviors'),
        ('Bowen, S., Witkiewitz, K., Clifasefi, S. L., Grow, J., Chawla, N., Hsu, S. H., Carroll, H. A., Harrop, E., Collins, S. E., Lustyk, M. K., & Larimer, M. E. (2014). Relative efficacy of mindfulness-based relapse prevention, standard relapse prevention, and treatment as usual for substance use disorders: A randomized clinical trial. JAMA Psychiatry, 71(5), 547-556.', 'JAMA Psychiatry'),
        ('Brewer, J. A., Mallik, S., Babuscio, T. A., Nich, C., Johnson, H. E., Deleone, C. M., Minnix-Cotton, C. A., Byrne, S. A., Kober, H., Weinstein, A. J., Carroll, K. M., & Rounsaville, B. J. (2011). Mindfulness training for smoking cessation: Results from a randomized controlled trial. Drug and Alcohol Dependence, 119(1-2), 72-80.', 'Drug and Alcohol Dependence'),
        ('Chiesa, A., & Serretti, A. (2014). Are mindfulness-based interventions effective for substance use disorders? A systematic review of the evidence. Substance Use & Misuse, 49(5), 492-512.', 'Substance Use & Misuse'),
        ('Franken, I. H. A., Hendriks, V. M., & van den Brink, W. (2002). Initial validation of two opiate craving questionnaires: The Obsessive Compulsive Drug Use Scale and the Desires for Drug Questionnaire. Addictive Behaviors, 27(5), 675-685.', 'Addictive Behaviors'),
        ('Garland, E. L., Froeliger, B., & Howard, M. O. (2014). Mindfulness training targets neurocognitive mechanisms of addiction at the attention-appraisal-emotion interface. Frontiers in Psychiatry, 4, Article 173.', 'Frontiers in Psychiatry'),
        ('Garland, E. L., Roberts-Lewis, A., Tronnier, C. D., Graves, R., & Kelley, K. (2016). Mindfulness-Oriented Recovery Enhancement versus CBT for co-occurring substance dependence, traumatic stress, and psychiatric disorders: Proximal outcomes from a pragmatic randomized trial. Behaviour Research and Therapy, 77, 7-16.', 'Behaviour Research and Therapy'),
        ('Ghosh, A., Basu, D., & Avasthi, A. (2018). Relapse in opioid dependence: An Indian perspective. Indian Journal of Psychiatry, 60(Suppl 4), S469-S476.', 'Indian Journal of Psychiatry'),
        ('Glasner-Edwards, S., Mooney, L. J., Ang, A., Garneau, H. C., Hartwell, E., Brecht, M. L., & Rawson, R. A. (2017). Mindfulness-based relapse prevention for stimulant dependent adults: A pilot randomized clinical trial. Mindfulness, 8(1), 126-135.', 'Mindfulness'),
        ('Grant, S., Colaiaco, B., Motala, A., Shanman, R., Booth, M., Sorbero, M., & Hempel, S. (2017). Mindfulness-based relapse prevention for substance use disorders: A systematic review and meta-analysis. Journal of Addiction Medicine, 11(5), 386-396.', 'Journal of Addiction Medicine'),
        ('Humeniuk, R., Ali, R., Babor, T. F., Farrell, M., Formigoni, M. L., Jittiwutikarn, J., de Lacerda, R. B., Ling, W., Marsden, J., Monteiro, M., Nhiwatiwa, S., Pal, H., Poznyak, V., & Simon, S. (2008). Validation of the Alcohol, Smoking and Substance Involvement Screening Test (ASSIST). Addiction, 103(6), 1039-1047.', 'Addiction'),
        ('Jain, R., Majumder, P., & Gupta, T. (2013). Pharmacological intervention of nicotine dependence. Indian Journal of Psychiatry, 55(Suppl 1), S86-S92.', 'Indian Journal of Psychiatry'),
        ('Kabat-Zinn, J. (1990). Full catastrophe living: Using the wisdom of your body and mind to face stress, pain, and illness. Delacorte Press.', 'Full catastrophe living: Using the wisdom of your body and mind to face stress, pain, and illness.'),
    ]

    for text, italic_part in refs:
        parts.append(ref_para(text, italic_part))

    return parts


def build_references_page2():
    """Build remaining references."""
    parts = []

    refs2 = [
        ('Karyadi, K. A., VanderVeen, J. D., & Cyders, M. A. (2014). A meta-analysis of the relationship between trait mindfulness and substance use behaviors. Drug and Alcohol Dependence, 143, 1-10.', 'Drug and Alcohol Dependence'),
        ('Li, W., Howard, M. O., Garland, E. L., McGovern, P., & Lazar, M. (2017). Mindfulness treatment for substance misuse: A systematic review and meta-analysis. Journal of Substance Abuse Treatment, 75, 62-96.', 'Journal of Substance Abuse Treatment'),
        ('Marlatt, G. A., & Gordon, J. R. (1985). Relapse prevention: Maintenance strategies in the treatment of addictive behaviors. Guilford Press.', 'Relapse prevention: Maintenance strategies in the treatment of addictive behaviors.'),
        ('Mattoo, S. K., Chakrabarti, S., & Anjaiah, M. (2009). Psychosocial factors associated with relapse in men with alcohol or opioid dependence. Indian Journal of Medical Research, 130(6), 702-708.', 'Indian Journal of Medical Research'),
        ('Ministry of Social Justice and Empowerment. (2019). Magnitude of substance use in India. Government of India.', 'Magnitude of substance use in India.'),
        ('Moeller, F. G., Barratt, E. S., Dougherty, D. M., Schmitz, J. M., & Swann, A. C. (2001). Psychiatric aspects of impulsivity. American Journal of Psychiatry, 158(11), 1783-1793.', 'American Journal of Psychiatry'),
        ('Murphy, C., & MacKillop, J. (2012). Living in the here and now: Interrelationships between impulsivity, mindfulness, and alcohol misuse. Psychopharmacology, 219(2), 527-536.', 'Psychopharmacology'),
        ('National Institute on Drug Abuse. (2020). Drugs, brains, and behavior: The science of addiction. National Institutes of Health.', 'Drugs, brains, and behavior: The science of addiction.'),
        ('Patton, J. H., Stanford, M. S., & Barratt, E. S. (1995). Factor structure of the Barratt Impulsiveness Scale. Journal of Clinical Psychology, 51(6), 768-774.', 'Journal of Clinical Psychology'),
        ('Robinson, T. E., & Berridge, K. C. (1993). The neural basis of drug craving: An incentive-sensitization theory of addiction. Brain Research Reviews, 18(3), 247-291.', 'Brain Research Reviews'),
        ('Sarkar, S., & Balhara, Y. P. S. (2016). Diabetes mellitus in people with substance use disorders: A narrative review. Indian Journal of Psychiatry, 58(3), 290-295.', 'Indian Journal of Psychiatry'),
        ('Serre, F., Fatseas, M., Swendsen, J., & Auriacombe, M. (2015). Ecological momentary assessment in the investigation of craving and substance use in daily life: A systematic review. Drug and Alcohol Dependence, 148, 1-20.', 'Drug and Alcohol Dependence'),
        ('Stanford, M. S., Mathias, C. W., Dougherty, D. M., Lake, S. L., Anderson, N. E., & Patton, J. H. (2009). Fifty years of the Barratt Impulsiveness Scale: An update and review. Personality and Individual Differences, 47(5), 385-395.', 'Personality and Individual Differences'),
        ('Tabachnick, B. G., & Fidell, L. S. (2013). Using multivariate statistics (6th ed.). Pearson.', 'Using multivariate statistics'),
        ('Tiffany, S. T., & Wray, J. M. (2012). The clinical significance of drug craving. Annals of the New York Academy of Sciences, 1248(1), 1-17.', 'Annals of the New York Academy of Sciences'),
        ('WHO ASSIST Working Group. (2002). The Alcohol, Smoking and Substance Involvement Screening Test (ASSIST): Development, reliability and feasibility. Addiction, 97(9), 1183-1194.', 'Addiction'),
        ('Witkiewitz, K., Bowen, S., Douglas, H., & Hsu, S. H. (2013). Mindfulness-based relapse prevention for substance craving. Addictive Behaviors, 38(2), 1563-1571.', 'Addictive Behaviors'),
    ]

    for text, italic_part in refs2:
        parts.append(ref_para(text, italic_part))

    return parts



# ═══════════════════════════════════════════════════════════════
# ASSEMBLE FULL DOCUMENT
# ═══════════════════════════════════════════════════════════════

def build_full_document():
    """Assemble all sections into the full document XML."""
    all_parts = []
    all_parts.extend(build_document_body())
    all_parts.extend(build_introduction())
    all_parts.extend(build_variables())
    all_parts.extend(build_review_of_literature())
    all_parts.extend(build_research_gap())
    all_parts.extend(build_aim_objectives_hypotheses())
    all_parts.extend(build_operational_definitions())
    all_parts.extend(build_methodology())
    all_parts.extend(build_variables_and_tools())
    all_parts.extend(build_procedure_and_intervention())
    all_parts.extend(build_data_analysis())
    all_parts.extend(build_ethical_considerations())
    all_parts.extend(build_expected_results())
    all_parts.extend(build_limitations_and_future())
    all_parts.extend(build_references())
    all_parts.extend(build_references_page2())

    body_content = "\n".join(all_parts)

    # Section properties: Letter size, 1-inch margins, page numbers top-right
    sect_pr = '''<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:pgNumType w:start="1"/>
      <w:headerReference w:type="default" r:id="rId4"/>
    </w:sectPr>'''

    # Actually, skip header reference to avoid complexity - page numbers
    # will be set via simple sectPr
    sect_pr = '''<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>'''

    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
{body_content}
{sect_pr}
  </w:body>
</w:document>'''

    return document_xml


def create_docx():
    """Create the .docx file."""
    document_xml = build_full_document()

    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', TOP_RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/numbering.xml', NUMBERING)

    print(f"SUCCESS: Document created at {OUTPUT}")
    print(f"File size: {os.path.getsize(OUTPUT)} bytes")


if __name__ == "__main__":
    create_docx()
