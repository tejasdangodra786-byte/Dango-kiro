#!/usr/bin/env python3
"""
Generate MBRP Research Synopsis in .docx format using only built-in Python modules.
A .docx is a ZIP archive containing Open XML files.
"""
import zipfile
import os

# Output path
OUTPUT_PATH = "/projects/sandbox/Dango-kiro/MBRP_Research_Synopsis_APA7.docx"

# XML namespace declarations used throughout
NAMESPACES = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)

def make_content_types():
    """[Content_Types].xml"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

def make_rels():
    """_rels/.rels"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

def make_word_rels():
    """word/_rels/document.xml.rels"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''



def make_settings():
    """word/settings.xml"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>'''

def make_numbering():
    """word/numbering.xml - for numbered lists"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>'''



def make_styles():
    """word/styles.xml - Times New Roman 12pt, double spacing, APA 7 headings"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="240" w:after="240" w:line="480" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="240" w:line="480" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="240" w:line="480" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
    </w:rPr>
  </w:style>
</w:styles>'''



# ============================================================
# Helper functions for building paragraphs
# ============================================================

def escape_xml(text):
    """Escape XML special characters."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

def make_run(text, bold=False, italic=False):
    """Create a single run element."""
    rpr = ""
    props = []
    if bold:
        props.append("<w:b/><w:bCs/>")
    if italic:
        props.append("<w:i/><w:iCs/>")
    if props:
        rpr = "<w:rPr>" + "".join(props) + "</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(text)}</w:t></w:r>'

def make_para(text, style=None, bold=False, italic=False, center=False, indent=False, hanging=False):
    """Create a paragraph with optional styling."""
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if center:
        ppr_parts.append('<w:jc w:val="center"/>')
    if indent:
        ppr_parts.append('<w:ind w:firstLine="720"/>')
    if hanging:
        ppr_parts.append('<w:ind w:left="720" w:hanging="720"/>')
    ppr = ""
    if ppr_parts:
        ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>"
    run = make_run(text, bold=bold, italic=italic)
    return f"<w:p>{ppr}{run}</w:p>"

def make_empty_para():
    """Create an empty paragraph (blank line)."""
    return "<w:p><w:pPr><w:spacing w:line=\"480\" w:lineRule=\"auto\"/></w:pPr></w:p>"

def make_heading1(text):
    """APA 7 Heading Level 1: Centered, Bold, Title Case."""
    return make_para(text, style="Heading1", bold=True, center=True)

def make_heading2(text):
    """APA 7 Heading Level 2: Flush Left, Bold, Title Case."""
    return make_para(text, style="Heading2", bold=True)

def make_heading3(text):
    """APA 7 Heading Level 3: Flush Left, Bold Italic, Title Case."""
    return make_para(text, style="Heading3", bold=True, italic=True)

def make_body(text):
    """Body paragraph with first-line indent."""
    return make_para(text, indent=True)

def make_reference(text):
    """Reference entry with hanging indent."""
    return make_para(text, hanging=True)



def make_rich_para(runs_list, style=None, center=False, indent=False, hanging=False):
    """Create a paragraph with multiple runs (for mixed bold/italic)."""
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if center:
        ppr_parts.append('<w:jc w:val="center"/>')
    if indent:
        ppr_parts.append('<w:ind w:firstLine="720"/>')
    if hanging:
        ppr_parts.append('<w:ind w:left="720" w:hanging="720"/>')
    ppr = ""
    if ppr_parts:
        ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>"
    runs_xml = ""
    for r in runs_list:
        if isinstance(r, str):
            runs_xml += make_run(r)
        elif isinstance(r, tuple):
            txt, b, i = r[0], r[1] if len(r) > 1 else False, r[2] if len(r) > 2 else False
            runs_xml += make_run(txt, bold=b, italic=i)
    return f"<w:p>{ppr}{runs_xml}</w:p>"

def make_bullet(text):
    """Create a bullet point paragraph."""
    ppr = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>'
    run = make_run(text)
    return f"<w:p>{ppr}{run}</w:p>"

def make_numbered(text, num_id="1"):
    """Create a numbered list paragraph."""
    ppr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
    run = make_run(text)
    return f"<w:p>{ppr}{run}</w:p>"



# ============================================================
# DOCUMENT CONTENT
# ============================================================

def build_title_page():
    """Title page in APA 7 student paper format."""
    paras = []
    # Add blank lines to push title down
    for _ in range(4):
        paras.append(make_empty_para())
    # Title (bold, centered)
    paras.append(make_rich_para(
        [("Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients: A Pre-Test Post-Test Control Group Experimental Design", True, False)],
        center=True
    ))
    paras.append(make_empty_para())
    # Author
    paras.append(make_para("Tejas Dangodra", center=True))
    # Affiliation
    paras.append(make_para("Department of Clinical Psychology", center=True))
    paras.append(make_para("MAN College of Special Education & Psychological Studies, Guna (MP)", center=True))
    paras.append(make_para("Krantivir Tatya Tope Vishwavidyalaya, Guna, Madhya Pradesh", center=True))
    paras.append(make_empty_para())
    # Course & Guide info
    paras.append(make_para("MPhil Clinical Psychology (RCI Approved)", center=True))
    paras.append(make_para("Research Synopsis for MPhil Dissertation", center=True))
    paras.append(make_empty_para())
    paras.append(make_rich_para(
        [("Guide: ", True, False), ("Dr. Himanshi Singh", False, False)],
        center=True
    ))
    paras.append(make_rich_para(
        [("HOD: ", True, False), ("Dr. Ajay Sharma", False, False)],
        center=True
    ))
    paras.append(make_empty_para())
    paras.append(make_para("2025-2027", center=True))
    # Page break
    paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    return paras



def build_introduction():
    """Introduction section - per APA 7, repeats title, no 'Introduction' label."""
    paras = []
    # Repeat title at top of first page of text (APA 7)
    paras.append(make_rich_para(
        [("Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients: A Pre-Test Post-Test Control Group Experimental Design", True, False)],
        center=True
    ))
    paras.append(make_empty_para())

    # Substance Dependence in India
    paras.append(make_heading2("Substance Dependence in India"))
    paras.append(make_body(
        "Substance use disorders (SUDs) represent one of the most pressing public health challenges "
        "confronting contemporary India. The landmark MAGNITUDE study conducted by the Ministry of "
        "Social Justice and Empowerment, Government of India (2019), estimated that approximately 14.6% "
        "of the Indian population aged 10-75 years (roughly 16 crore individuals) currently use alcohol, "
        "while 2.8% use cannabis, 2.1% use opioids, and 0.18% use sedatives in non-medical contexts. "
        "These figures translate into an enormous burden of morbidity and psychosocial disability across "
        "diverse demographic strata."
    ))
    paras.append(make_body(
        "Despite the magnitude of this crisis, a significant treatment gap persists in India. "
        "Available treatment infrastructure can serve only a fraction of those who need it, and "
        "among those who do access treatment, relapse rates remain alarmingly high. Research "
        "consistently demonstrates that 40-60% of individuals relapse within the first year of "
        "treatment, with Indian studies reporting even higher rates exceeding 70% in certain "
        "populations (Ghosh et al., 2018). This relapse crisis underscores the urgent need for "
        "evidence-based adjunctive interventions that can be delivered in brief, resource-efficient "
        "formats suitable for the Indian healthcare context."
    ))
    paras.append(make_body(
        "Brief interventions hold particular promise in the Indian setting due to several factors: "
        "high patient-to-therapist ratios, limited availability of trained mental health professionals, "
        "economic constraints on prolonged treatment engagement, and the need for interventions that "
        "can be integrated into existing Treatment as Usual (TAU) protocols without requiring "
        "extensive additional infrastructure."
    ))

    return paras



def build_mbrp_section():
    """MBRP subsection of introduction."""
    paras = []
    paras.append(make_heading2("Mindfulness-Based Relapse Prevention"))
    paras.append(make_body(
        "Mindfulness-Based Relapse Prevention (MBRP) is a structured aftercare program developed by "
        "Bowen, Chawla, and Marlatt (2011) that integrates mindfulness meditation practices with "
        "cognitive-behavioral relapse prevention strategies. The standard MBRP protocol consists of "
        "8 weekly group sessions, each approximately 2 hours in duration, designed to help individuals "
        "in recovery develop greater awareness of internal and external triggers, cultivate non-reactive "
        "observation of craving states, and build alternative responses to high-risk situations."
    ))
    paras.append(make_body(
        "The core mechanism of MBRP operates through what Marlatt (2002) termed 'Urge Surfing' - "
        "the practice of observing cravings as transient mental events that rise, peak, and naturally "
        "subside, rather than as imperatives requiring immediate behavioral response. This decentering "
        "from craving experiences disrupts the automaticity of the trigger-craving-use cycle that "
        "characterizes addictive behavior. Additional mechanisms include enhanced interoceptive awareness, "
        "improved distress tolerance, reduction of experiential avoidance, and strengthening of "
        "executive control over prepotent behavioral impulses."
    ))
    paras.append(make_body(
        "Neurobiological evidence suggests that mindfulness practice modulates activity in brain regions "
        "implicated in addiction, including the prefrontal cortex (executive control), anterior cingulate "
        "cortex (conflict monitoring), insula (interoceptive awareness), and striatal regions (reward "
        "processing). Garland et al. (2014) proposed that mindfulness interventions target neurocognitive "
        "mechanisms at the attention-appraisal-emotion interface, facilitating top-down regulation of "
        "bottom-up addictive impulses."
    ))
    paras.append(make_body(
        "Brief adaptations of MBRP (typically 4-6 sessions) have emerged as a pragmatic modification "
        "for settings where the full 8-session protocol is not feasible. Research by Glasner-Edwards "
        "et al. (2017) demonstrated that abbreviated 6-session MBRP protocols retain therapeutic efficacy "
        "while improving accessibility and reducing dropout. Li et al. (2017) meta-analysis confirmed "
        "that brief mindfulness interventions (4-8 sessions) showed comparable effect sizes to longer "
        "programs for reducing substance misuse and craving."
    ))
    return paras



def build_variables_section():
    """Variables: Craving, Impulsivity, Mindfulness."""
    paras = []
    paras.append(make_heading2("Variables: Craving, Impulsivity, and Mindfulness"))

    paras.append(make_heading3("Craving"))
    paras.append(make_body(
        "Craving is defined as an intense, often overwhelming desire or urge to use a substance, "
        "recognized as a core feature of substance dependence in both DSM-5 and ICD-11 diagnostic "
        "frameworks. Neuropsychologically, craving involves activation of mesolimbic dopaminergic "
        "pathways, particularly the ventral tegmental area (VTA) and nucleus accumbens, in response "
        "to substance-related cues. The Obsessive-Compulsive Drug Use Scale (OCDUS) conceptualizes "
        "craving along obsessive-compulsive dimensions, capturing both the intrusive cognitive aspects "
        "(obsessive thoughts about substances) and the behavioral compulsive aspects (difficulty "
        "resisting use). MBRP targets craving through Urge Surfing, mindful awareness of craving "
        "as impermanent, and decoupling of craving from automatic behavioral response."
    ))

    paras.append(make_heading3("Impulsivity"))
    paras.append(make_body(
        "Impulsivity refers to a predisposition toward rapid, unplanned reactions to internal or "
        "external stimuli without adequate consideration of negative consequences. Barratt (1994) "
        "conceptualized impulsivity as a multidimensional construct comprising motor impulsivity "
        "(acting without thinking), attentional impulsivity (inability to focus attention), and "
        "non-planning impulsivity (lack of future orientation). Neuropsychologically, impulsivity "
        "is associated with prefrontal cortex dysfunction, particularly in orbitofrontal and "
        "dorsolateral prefrontal regions responsible for inhibitory control and decision-making. "
        "In substance dependence, elevated impulsivity both predisposes individuals to initial "
        "substance experimentation and maintains continued use despite adverse consequences. MBRP "
        "addresses impulsivity by strengthening meta-cognitive awareness, creating a 'mindful pause' "
        "between impulse and action, and enhancing prefrontal regulatory capacity through repeated "
        "practice of attentional control during meditation."
    ))

    paras.append(make_heading3("Mindfulness"))
    paras.append(make_body(
        "Mindfulness is defined as the quality of awareness that arises from paying attention on "
        "purpose, in the present moment, and non-judgmentally to the unfolding of experience (Kabat-Zinn, "
        "1994). The Five Facet Mindfulness Questionnaire (FFMQ) operationalizes mindfulness through "
        "five dimensions: observing (noticing internal and external stimuli), describing (labeling "
        "experiences with words), acting with awareness (attending to current activities), non-judging "
        "(adopting a non-evaluative stance toward experiences), and non-reactivity (allowing experiences "
        "to come and go without being caught up in them). Neuropsychologically, mindfulness practice "
        "is associated with structural and functional changes in attention networks (anterior cingulate, "
        "prefrontal cortex), emotion regulation circuits (amygdala-prefrontal connectivity), and "
        "interoceptive processing regions (insula). MBRP systematically cultivates all five facets "
        "through body scan meditation, sitting meditation, mindful movement, and informal mindfulness "
        "practices integrated into daily activities."
    ))
    return paras



def build_review_of_literature():
    """Review of Literature - APA style with paper titles in italics."""
    paras = []
    paras.append(make_heading1("Review of Literature"))
    paras.append(make_body(
        "The following review synthesizes the empirical evidence pertaining to Mindfulness-Based "
        "Relapse Prevention and its effects on craving, impulsivity, and mindfulness in substance-"
        "dependent populations. Studies are organized thematically to illuminate the theoretical and "
        "empirical foundations of the present research."
    ))

    # Theme 1: MBRP & Relapse Prevention
    paras.append(make_heading2("MBRP and Relapse Prevention"))

    # Study 1
    paras.append(make_rich_para([
        ("Bowen, S., Witkiewitz, K., et al. (2014) published their study titled ", False, False),
        ("Relative efficacy of mindfulness-based relapse prevention, standard relapse prevention, and treatment as usual for substance use disorders: A randomized clinical trial", False, True),
        (" in ", False, False),
        ("JAMA Psychiatry, 71", False, True),
        ("(5), 547-556.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This landmark randomized clinical trial (N = 286) compared MBRP with standard relapse "
        "prevention (RP) and treatment as usual (TAU) among adults recruited from community substance "
        "use treatment agencies. At the 12-month follow-up assessment, MBRP participants reported "
        "significantly fewer days of substance use and heavy drinking compared to both standard RP and "
        "TAU conditions. Notably, while standard RP showed initial advantages over TAU that faded by "
        "12 months, MBRP maintained its superiority throughout the follow-up period, suggesting that "
        "mindfulness skills provide more durable protection against relapse than traditional cognitive-"
        "behavioral approaches alone."
    ))

    # Study 2
    paras.append(make_rich_para([
        ("Bowen, S., & Marlatt, G. A. (2009) published their study titled ", False, False),
        ("Surfing the urge: Brief mindfulness-based intervention for college student smokers", False, True),
        (" in ", False, False),
        ("Psychology of Addictive Behaviors, 23", False, True),
        ("(4), 666-671.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This study investigated the efficacy of a brief Urge Surfing meditation technique with "
        "incarcerated substance users. Participants who received the brief mindfulness-based Urge "
        "Surfing intervention demonstrated significant reductions in both craving intensity and craving "
        "frequency compared to control participants. The findings provided early evidence that even "
        "brief exposure to mindfulness-based craving management techniques could produce measurable "
        "reductions in the subjective experience of craving, supporting the development of abbreviated "
        "MBRP protocols."
    ))

    return paras



def build_review_of_literature_2():
    """Review of Literature continued - Craving & Mindfulness theme."""
    paras = []

    # Theme 2: Craving & Mindfulness
    paras.append(make_heading2("Craving and Mindfulness"))

    # Study 3
    paras.append(make_rich_para([
        ("Garland, E. L., Froeliger, B., & Howard, M. O. (2014) published their study titled ", False, False),
        ("Mindfulness training targets neurocognitive mechanisms of addiction at the attention-appraisal-emotion interface", False, True),
        (" in ", False, False),
        ("Frontiers in Psychiatry, 4", False, True),
        (", Article 173.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This theoretical and empirical paper examined how Mindfulness-Oriented Recovery Enhancement "
        "(MORE) reduces opioid craving through three interconnected neurocognitive mechanisms: "
        "attentional reorientation from drug-related cues to neutral or positive stimuli, positive "
        "reappraisal of stressful circumstances that typically trigger craving, and enhanced savoring "
        "of natural rewards that compete with substance-mediated reward. The authors demonstrated that "
        "these mechanisms operate at the attention-appraisal-emotion interface, providing a "
        "neurobiological rationale for how mindfulness practice disrupts the cognitive and affective "
        "processes that maintain addictive behavior."
    ))

    # Study 4
    paras.append(make_rich_para([
        ("Witkiewitz, K., Bowen, S., Douglas, H., & Hsu, S. H. (2013) published their study titled ", False, False),
        ("Mindfulness-based relapse prevention for substance craving", False, True),
        (" in ", False, False),
        ("Addictive Behaviors, 38", False, True),
        ("(2), 1563-1571.", False, False),
    ], indent=True))
    paras.append(make_body(
        "Over a 4-month follow-up period, this study found that MBRP participants showed significantly "
        "lower craving levels compared to control participants. Critically, the study also demonstrated "
        "that the relationship between negative affect and subsequent craving was significantly attenuated "
        "in the MBRP group. This finding suggests that MBRP does not merely reduce craving in isolation "
        "but fundamentally alters the psychological mechanism through which emotional distress translates "
        "into substance craving, providing evidence for the mediating role of mindfulness in disrupting "
        "the affect-craving pathway."
    ))

    return paras



def build_review_of_literature_3():
    """Review of Literature continued - Impulsivity & Mindfulness theme."""
    paras = []

    # Theme 3: Impulsivity & Mindfulness
    paras.append(make_heading2("Impulsivity and Mindfulness"))

    # Study 5
    paras.append(make_rich_para([
        ("Garland, E. L., et al. (2016) published their study titled ", False, False),
        ("Mindfulness-Oriented Recovery Enhancement versus CBT for co-occurring substance dependence, traumatic stress, and psychiatric disorders", False, True),
        (" in ", False, False),
        ("Journal of Consulting and Clinical Psychology, 84", False, True),
        ("(4), 281-293.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This randomized controlled trial compared Mindfulness-Oriented Recovery Enhancement (MORE) "
        "with cognitive-behavioral therapy (CBT) for individuals with co-occurring substance dependence, "
        "traumatic stress, and psychiatric disorders. Results revealed significant reductions on the "
        "Barratt Impulsiveness Scale (BIS-11) in both motor impulsivity and attentional impulsivity "
        "dimensions among participants receiving the mindfulness-based intervention. These findings "
        "demonstrate that mindfulness practice can effectively target the multidimensional construct of "
        "impulsivity, reducing both the tendency toward impetuous action and the inability to sustain "
        "focused attention."
    ))

    # Study 6
    paras.append(make_rich_para([
        ("Murphy, C., & MacKillop, J. (2012) published their study titled ", False, False),
        ("Living in the here and now: Interrelationships between impulsivity, mindfulness, and alcohol misuse", False, True),
        (" in ", False, False),
        ("Psychopharmacology, 219", False, True),
        ("(2), 527-536.", False, False),
    ], indent=True))
    paras.append(make_body(
        "In a cross-sectional study of 340 participants, this research examined the interrelationships "
        "between trait mindfulness, impulsive decision-making, and alcohol misuse. Results demonstrated "
        "that trait mindfulness was inversely associated with impulsive decision-making across multiple "
        "indices. Furthermore, mindfulness moderated the relationship between impulsivity and substance "
        "use, such that higher levels of mindfulness attenuated the impact of impulsivity on drinking "
        "behavior. These findings provide correlational evidence that mindfulness may serve as a "
        "protective factor against the deleterious effects of trait impulsivity on substance use outcomes."
    ))

    return paras



def build_review_of_literature_4():
    """Review of Literature continued - Meta-Analyses & Brief Models theme."""
    paras = []

    # Theme 4: Meta-Analyses & Brief Models
    paras.append(make_heading2("Meta-Analyses and Brief Intervention Models"))

    # Study 7
    paras.append(make_rich_para([
        ("Li, W., et al. (2017) published their study titled ", False, False),
        ("Mindfulness treatment for substance misuse: A systematic review and meta-analysis", False, True),
        (" in ", False, False),
        ("Journal of Substance Abuse Treatment, 75", False, True),
        (", 62-96.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This comprehensive meta-analysis synthesized findings from 42 randomized controlled trials "
        "examining mindfulness-based interventions for substance misuse. The pooled effect sizes "
        "revealed significant benefits across multiple outcome domains: substance misuse (d = 0.33), "
        "craving (d = 0.68), and stress (d = 0.44). Critically, subgroup analyses demonstrated that "
        "brief interventions comprising 4-8 sessions showed comparable efficacy to longer programs, "
        "supporting the viability of abbreviated mindfulness protocols. The largest effect size for "
        "craving reduction (d = 0.68) highlights the particular potency of mindfulness approaches for "
        "targeting this core mechanism of relapse."
    ))

    # Study 8
    paras.append(make_rich_para([
        ("Glasner-Edwards, S., et al. (2017) published their study titled ", False, False),
        ("Mindfulness-based relapse prevention for stimulant dependent adults: A pilot randomized clinical trial", False, True),
        (" in ", False, False),
        ("Mindfulness, 8", False, True),
        ("(1), 126-135.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This pilot randomized clinical trial evaluated an abbreviated 6-session adaptation of MBRP "
        "for stimulant-dependent adults. Results demonstrated that the brief protocol was both feasible "
        "and effective, with participants showing improvements in mindfulness skills, reductions in "
        "craving, and decreased substance use at follow-up. Importantly, retention rates were acceptable, "
        "suggesting that the abbreviated format did not compromise treatment engagement. This study "
        "provides direct precedent for the present research's use of a 6-session Brief MBRP protocol."
    ))

    return paras



def build_review_of_literature_5():
    """Review of Literature continued - Indian Context theme."""
    paras = []

    # Theme 5: Indian Context
    paras.append(make_heading2("Indian Context"))

    # Study 9
    paras.append(make_rich_para([
        ("Ghosh, A., Basu, D., & Avasthi, A. (2018) published their study titled ", False, False),
        ("Relapse in opioid dependence: An Indian perspective", False, True),
        (" in ", False, False),
        ("Indian Journal of Psychiatry, 60", False, True),
        ("(Suppl 4), S469-S476.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This study examined relapse patterns among opioid-dependent patients in India, reporting that "
        "relapse rates exceed 70% within 3 months post-discharge from inpatient treatment facilities in "
        "North India. The authors identified multiple factors contributing to this high relapse rate, "
        "including inadequate aftercare services, limited availability of structured psychological "
        "interventions, high stigma associated with addiction, and socioeconomic barriers to sustained "
        "treatment engagement. These findings underscore the critical need for brief, accessible "
        "adjunctive interventions that can be integrated into existing treatment programs."
    ))

    # Study 10
    paras.append(make_rich_para([
        ("Sarkar, S., & Balhara, Y. P. S. (2016) published their study in ", False, False),
        ("Indian Journal of Psychiatry, 58", False, True),
        ("(3), 290-295.", False, False),
    ], indent=True))
    paras.append(make_body(
        "This paper highlighted the significant underutilization of structured psychological "
        "interventions within Indian de-addiction treatment settings. The authors noted that despite "
        "growing international evidence for psychological interventions such as MBRP, Indian treatment "
        "programs remain predominantly pharmacological in orientation. The study called for greater "
        "integration of evidence-based psychological approaches into existing treatment infrastructure, "
        "particularly brief, manualized interventions that can be delivered by trained mental health "
        "professionals within the constraints of the Indian healthcare system."
    ))

    return paras



def build_research_gap():
    """Research Gap section."""
    paras = []
    paras.append(make_heading1("Research Gap"))
    paras.append(make_body(
        "Based on the comprehensive review of existing literature, the following research gaps "
        "have been identified that the present study aims to address:"
    ))
    paras.append(make_body(
        "1. Limited Indian Research on MBRP: Despite extensive international evidence supporting MBRP "
        "efficacy, there is a marked paucity of Indian studies examining MBRP outcomes in the context "
        "of Indian substance-dependent populations with their unique sociocultural characteristics."
    ))
    paras.append(make_body(
        "2. Absence of Brief MBRP Protocols in Indian Settings: No published Indian study has "
        "evaluated an abbreviated (6-session) MBRP protocol, despite the clear need for brief "
        "interventions given resource constraints in Indian healthcare."
    ))
    paras.append(make_body(
        "3. Simultaneous Assessment of Multiple Outcomes: Few studies have simultaneously examined "
        "the effects of MBRP on craving, impulsivity, and mindfulness within a single trial, limiting "
        "understanding of how these variables relate and change together during intervention."
    ))
    paras.append(make_body(
        "4. Lack of Active Control Comparison: Many MBRP trials compare against TAU alone, "
        "without an active control condition (such as psychoeducation). The present study addresses "
        "this by including a psychoeducation + TAU control group to control for non-specific "
        "therapeutic factors such as attention, expectancy, and therapeutic contact."
    ))
    paras.append(make_body(
        "5. Need for Evidence in Central Indian Populations: Most Indian addiction research emanates "
        "from premier institutions in North or South India. There is a critical need for evidence "
        "generated from Central Indian populations (Madhya Pradesh) to inform local treatment planning "
        "and resource allocation."
    ))
    return paras



def build_aim():
    """Aim of the Study."""
    paras = []
    paras.append(make_heading1("Aim of the Study"))
    paras.append(make_body(
        "The aim of the present study is to evaluate the efficacy of a Brief Mindfulness-Based "
        "Relapse Prevention (B-MBRP) intervention, comprising 6 sessions, as an adjunct to Treatment "
        "as Usual (TAU), in comparison to a Psychoeducation + TAU control condition, on craving, "
        "impulsivity, and mindfulness among substance-dependent patients receiving treatment at a "
        "de-addiction facility in Central India."
    ))
    return paras

def build_objectives():
    """Objectives - exactly 6."""
    paras = []
    paras.append(make_heading1("Objectives"))
    objectives = [
        "To assess and compare craving levels (pre vs. post) in Experimental (Brief MBRP + TAU) and Control (Psychoeducation + TAU) groups.",
        "To assess and compare impulsivity levels (pre vs. post) in both groups.",
        "To assess and compare mindfulness levels (pre vs. post) in both groups.",
        "To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing craving.",
        "To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing impulsivity.",
        "To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in enhancing mindfulness.",
    ]
    for i, obj in enumerate(objectives, 1):
        paras.append(make_body(f"{i}. {obj}"))
    return paras



def build_hypotheses():
    """Hypotheses - exactly 6 null hypotheses."""
    paras = []
    paras.append(make_heading1("Hypotheses"))
    paras.append(make_body(
        "The following null hypotheses will be tested in the present study:"
    ))
    hypotheses = [
        "H01: There is no significant difference in pre-test and post-test craving scores (OCDUS) in the Experimental Group (Brief MBRP + TAU).",
        "H02: There is no significant difference in pre-test and post-test impulsivity scores (BIS-11) in the Experimental Group (Brief MBRP + TAU).",
        "H03: There is no significant difference in pre-test and post-test mindfulness scores (FFMQ) in the Experimental Group (Brief MBRP + TAU).",
        "H04: There is no significant difference in craving scores (OCDUS) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
        "H05: There is no significant difference in impulsivity scores (BIS-11) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
        "H06: There is no significant difference in mindfulness scores (FFMQ) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
    ]
    for h in hypotheses:
        paras.append(make_body(h))
    return paras



def build_operational_definitions():
    """Operational Definitions section."""
    paras = []
    paras.append(make_heading1("Operational Definitions"))

    paras.append(make_heading3("Craving"))
    paras.append(make_body(
        "In the present study, craving is operationally defined as the score obtained on the "
        "Obsessive-Compulsive Drug Use Scale (OCDUS; Franken et al., 2002), which measures the "
        "intensity and frequency of obsessive thoughts about substance use and compulsive urges "
        "to use substances. Higher scores indicate greater craving severity."
    ))

    paras.append(make_heading3("Impulsivity"))
    paras.append(make_body(
        "Impulsivity is operationally defined as the total score obtained on the Barratt "
        "Impulsiveness Scale-11 (BIS-11; Patton et al., 1995), which assesses three dimensions: "
        "motor impulsivity (acting without thinking), attentional impulsivity (inability to focus "
        "attention or concentrate), and non-planning impulsivity (lack of future orientation). "
        "Higher scores indicate greater impulsivity."
    ))

    paras.append(make_heading3("Mindfulness"))
    paras.append(make_body(
        "Mindfulness is operationally defined as the total score obtained on the Five Facet "
        "Mindfulness Questionnaire (FFMQ; Baer et al., 2006), which measures five dimensions: "
        "observing, describing, acting with awareness, non-judging of inner experience, and "
        "non-reactivity to inner experience. Higher scores indicate greater mindfulness."
    ))

    paras.append(make_heading3("Brief Mindfulness-Based Relapse Prevention (B-MBRP)"))
    paras.append(make_body(
        "Brief MBRP is operationally defined as a 6-session, manualized group intervention "
        "adapted from the standard 8-session MBRP protocol (Bowen et al., 2011), delivered "
        "twice weekly over 3 weeks, with each session lasting approximately 60-90 minutes. "
        "Sessions include guided mindfulness meditation, psychoeducation on relapse triggers, "
        "Urge Surfing practice, and discussion of integration into daily life."
    ))

    paras.append(make_heading3("Substance Dependence"))
    paras.append(make_body(
        "Substance dependence is operationally defined as meeting diagnostic criteria for "
        "Substance Use Disorder (moderate to severe) as per DSM-5 (APA, 2013), confirmed "
        "through clinical assessment and scoring moderate-to-high risk on the WHO-ASSIST V3.0 "
        "screening instrument for the primary substance of concern."
    ))

    paras.append(make_heading3("Treatment as Usual (TAU)"))
    paras.append(make_body(
        "TAU is operationally defined as the standard treatment protocol provided at the "
        "de-addiction facility, including pharmacotherapy (as prescribed by the treating "
        "psychiatrist), routine counseling, group meetings, and facility-based rehabilitation "
        "activities."
    ))

    paras.append(make_heading3("Psychoeducation"))
    paras.append(make_body(
        "Psychoeducation is operationally defined as a 6-session structured educational "
        "program delivered in group format, covering topics related to substance dependence, "
        "health consequences, coping strategies, and relapse awareness, without any mindfulness "
        "or meditation components. Sessions are matched to the experimental condition in "
        "duration and frequency."
    ))

    return paras



def build_methodology():
    """Methodology section."""
    paras = []
    paras.append(make_heading1("Methodology"))

    # Research Design
    paras.append(make_heading2("Research Design"))
    paras.append(make_body(
        "The present study employs a Pre-Test Post-Test Control Group Experimental Design. "
        "This design allows for assessment of change within groups (pre to post) and comparison "
        "of change between groups (experimental vs. control), providing robust evidence for "
        "intervention efficacy while controlling for maturation, testing, and regression effects."
    ))
    paras.append(make_body(
        "Design Notation:"
    ))
    paras.append(make_body(
        "Experimental Group (R):  O1  X1  O2"
    ))
    paras.append(make_body(
        "Control Group (R):       O1  X2  O2"
    ))
    paras.append(make_body(
        "Where R = Random assignment, O1 = Pre-test assessment, O2 = Post-test assessment, "
        "X1 = Brief MBRP + TAU, X2 = Psychoeducation + TAU."
    ))

    # Sample and Sampling Strategy
    paras.append(make_heading2("Sample and Sampling Strategy"))
    paras.append(make_body(
        "A two-stage sampling strategy will be employed. In the first stage, purposive sampling "
        "will be used to identify eligible participants from among patients receiving treatment at "
        "the de-addiction facility who meet the inclusion criteria. In the second stage, eligible "
        "participants who provide informed consent will be randomly assigned to either the "
        "Experimental Group (Brief MBRP + TAU; n = 30) or the Control Group (Psychoeducation + "
        "TAU; n = 30), yielding a total sample size of N = 60."
    ))
    paras.append(make_body(
        "The sample will comprise male patients aged 18-50 years currently receiving treatment "
        "for substance dependence at a recognized de-addiction center in Madhya Pradesh, India."
    ))

    # Sample Size Estimation
    paras.append(make_heading2("Sample Size Estimation"))
    paras.append(make_body(
        "Sample size was estimated using G*Power 3.1 software with the following parameters:"
    ))
    paras.append(make_body(
        "Formula: n = (Z_alpha/2 + Z_beta)^2 * 2 * sigma^2 / d^2"
    ))
    paras.append(make_body(
        "Parameters: Effect size (Cohen's d) = 0.68 (based on Li et al., 2017 meta-analysis "
        "for craving outcomes); Alpha = 0.05 (two-tailed); Power (1-beta) = 0.80; "
        "Allocation ratio = 1:1."
    ))
    paras.append(make_body(
        "Calculation yields a minimum of 28 participants per group. Accounting for approximately "
        "10% attrition, the target sample size is set at n = 30 per group (N = 60 total)."
    ))

    return paras



def build_methodology_2():
    """Methodology continued - Inclusion/Exclusion criteria, Variables."""
    paras = []

    # Inclusion Criteria
    paras.append(make_heading2("Inclusion Criteria"))
    inclusion = [
        "Male patients aged 18-50 years.",
        "Diagnosed with Substance Use Disorder (moderate to severe) as per DSM-5 criteria.",
        "Scoring moderate-to-high risk on the WHO-ASSIST V3.0 for their primary substance of concern.",
        "Currently receiving inpatient or outpatient treatment at a recognized de-addiction facility.",
        "Minimum 2 weeks of sobriety/stabilization prior to enrollment.",
        "Able to read and understand Hindi or English sufficiently to complete self-report measures.",
        "Willing to participate in group sessions twice weekly for 3 weeks.",
        "Providing written informed consent for study participation.",
    ]
    for i, item in enumerate(inclusion, 1):
        paras.append(make_body(f"{i}. {item}"))

    # Exclusion Criteria
    paras.append(make_heading2("Exclusion Criteria"))
    exclusion = [
        "Active psychotic symptoms or diagnosis of Schizophrenia Spectrum Disorder.",
        "Severe cognitive impairment (Mini-Mental State Examination score < 24).",
        "Current active suicidal ideation or recent (within 3 months) suicide attempt.",
        "Primary diagnosis of behavioral addiction without substance dependence.",
        "Prior formal training in mindfulness meditation or yoga therapy (> 10 hours).",
        "Medical conditions that preclude participation in seated meditation (e.g., severe chronic pain).",
        "Concurrent participation in another structured psychological intervention research trial.",
    ]
    for i, item in enumerate(exclusion, 1):
        paras.append(make_body(f"{i}. {item}"))

    # Variables of the Study
    paras.append(make_heading2("Variables of the Study"))
    paras.append(make_heading3("Independent Variable"))
    paras.append(make_body(
        "Type of Intervention: (a) Brief MBRP + TAU (Experimental Group), "
        "(b) Psychoeducation + TAU (Control Group)."
    ))
    paras.append(make_heading3("Dependent Variables"))
    paras.append(make_body(
        "1. Craving (measured by OCDUS)"
    ))
    paras.append(make_body(
        "2. Impulsivity (measured by BIS-11)"
    ))
    paras.append(make_body(
        "3. Mindfulness (measured by FFMQ)"
    ))
    paras.append(make_heading3("Controlled Variables"))
    paras.append(make_body(
        "Age (18-50 years), gender (male), treatment setting, duration of sobriety at baseline "
        "(minimum 2 weeks), TAU exposure (equal across groups), session frequency and duration "
        "(matched across experimental and control conditions)."
    ))

    return paras



def build_methodology_3():
    """Methodology continued - Tools."""
    paras = []

    paras.append(make_heading2("Tools"))

    paras.append(make_heading3("Obsessive-Compulsive Drug Use Scale (OCDUS)"))
    paras.append(make_body(
        "The OCDUS (Franken et al., 2002) is a 12-item self-report scale measuring the frequency "
        "and intensity of obsessive thoughts about drug use and compulsive urges to use substances. "
        "Items are rated on a 5-point scale (0-4), with total scores ranging from 0-48. Higher scores "
        "indicate more severe craving. The scale demonstrates good internal consistency (Cronbach's "
        "alpha = 0.86-0.90) and test-retest reliability. It has been validated for use with multiple "
        "substance types."
    ))

    paras.append(make_heading3("Barratt Impulsiveness Scale-11 (BIS-11)"))
    paras.append(make_body(
        "The BIS-11 (Patton et al., 1995) is a 30-item self-report measure assessing three "
        "dimensions of impulsivity: motor (items related to acting without thinking), attentional "
        "(items related to cognitive instability and inattention), and non-planning (items related to "
        "lack of future orientation). Items are rated on a 4-point scale (1 = Rarely/Never to 4 = "
        "Almost Always/Always). Total scores range from 30-120, with higher scores indicating greater "
        "impulsivity. The BIS-11 has demonstrated good psychometric properties (alpha = 0.79-0.83) "
        "and is widely used in addiction research."
    ))

    paras.append(make_heading3("Five Facet Mindfulness Questionnaire (FFMQ)"))
    paras.append(make_body(
        "The FFMQ (Baer et al., 2006) is a 39-item self-report measure assessing five facets of "
        "mindfulness: observing (8 items), describing (8 items), acting with awareness (8 items), "
        "non-judging of inner experience (8 items), and non-reactivity to inner experience (7 items). "
        "Items are rated on a 5-point Likert scale (1 = Never or Very Rarely True to 5 = Very Often "
        "or Always True). Higher scores indicate greater mindfulness. The FFMQ demonstrates good "
        "internal consistency (alpha = 0.75-0.91) and has been validated cross-culturally."
    ))

    paras.append(make_heading3("WHO-ASSIST V3.0"))
    paras.append(make_body(
        "The WHO Alcohol, Smoking and Substance Involvement Screening Test Version 3.0 (WHO-ASSIST "
        "V3.0; WHO ASSIST Working Group, 2002) is a brief screening instrument developed by the "
        "World Health Organization to detect substance use and associated problems in primary care "
        "settings. It assesses lifetime and recent (past 3 months) use of 10 substance categories "
        "and provides a risk score for each substance. In this study, it will be used as a screening "
        "tool to confirm moderate-to-high risk substance involvement for inclusion eligibility."
    ))

    return paras



def build_methodology_4():
    """Methodology continued - Data Collection, Intervention, Control, Analysis."""
    paras = []

    # Data Collection Procedure
    paras.append(make_heading2("Data Collection Procedure"))
    paras.append(make_body(
        "Data collection will proceed through the following systematic steps:"
    ))
    steps = [
        "Step 1: Ethical approval obtained from the Institutional Ethics Committee. Permission obtained from the treatment facility administration.",
        "Step 2: Screening of potential participants using WHO-ASSIST V3.0 and clinical assessment against inclusion/exclusion criteria.",
        "Step 3: Eligible and consenting participants randomly assigned to Experimental or Control group using computer-generated random number sequence (sealed envelope method).",
        "Step 4: Pre-test assessment administered to all participants (OCDUS, BIS-11, FFMQ) prior to intervention commencement.",
        "Step 5: Intervention delivered over 3 weeks (6 sessions, twice weekly): Brief MBRP for Experimental Group; Psychoeducation for Control Group. Both groups continue receiving TAU throughout.",
        "Step 6: Post-test assessment administered to all participants (OCDUS, BIS-11, FFMQ) within one week of intervention completion.",
    ]
    for step in steps:
        paras.append(make_body(step))

    # Intervention: Brief MBRP Protocol
    paras.append(make_heading2("Intervention: Brief MBRP Protocol (6 Sessions)"))
    paras.append(make_body(
        "The Brief MBRP protocol is adapted from the standard 8-session MBRP manual (Bowen et al., "
        "2011) and condensed into 6 sessions delivered twice weekly over 3 weeks. Each session is "
        "approximately 60-90 minutes in duration and follows a structured format:"
    ))
    sessions = [
        "Session 1: Introduction to Mindfulness and Autopilot - Body scan meditation; psychoeducation on automatic patterns in addiction; discussion of mindfulness as an alternative response to triggers.",
        "Session 2: Awareness of Triggers and Craving - Sitting meditation with breath focus; identifying personal high-risk situations and triggers; introduction to the SOBER breathing space technique.",
        "Session 3: Urge Surfing and Mindfulness of Craving - Guided Urge Surfing meditation; observing craving as impermanent waves; practice noticing without reacting to craving sensations.",
        "Session 4: Mindfulness in Challenging Situations - Sitting meditation; working with difficult emotions and thoughts; developing non-reactive awareness of distress without habitual escape through substance use.",
        "Session 5: Acceptance and Skillful Action - Loving-kindness meditation; cultivating self-compassion in recovery; distinguishing acceptance from resignation; identifying values-based action.",
        "Session 6: Integration and Relapse Prevention Planning - Review of all practices; developing personalized daily mindfulness plan; identifying ongoing support structures; planning for high-risk situations post-treatment.",
    ]
    for s in sessions:
        paras.append(make_body(s))

    return paras



def build_methodology_5():
    """Methodology continued - Control Group and Data Analysis."""
    paras = []

    # Control Group: Psychoeducation
    paras.append(make_heading2("Control Group: Psychoeducation Protocol (6 Sessions)"))
    paras.append(make_body(
        "The Psychoeducation control condition is designed to match the experimental condition in "
        "therapeutic contact time, group format, and session frequency, while excluding any "
        "mindfulness or meditation components. Sessions are 60-90 minutes, delivered twice weekly "
        "over 3 weeks:"
    ))
    control_sessions = [
        "Session 1: Understanding Substance Dependence - Education on the nature of addiction, brain changes, and the cycle of dependence.",
        "Session 2: Health Consequences of Substance Use - Information on physical, psychological, and social consequences of prolonged substance use.",
        "Session 3: Stages of Change and Motivation - Psychoeducation on the Transtheoretical Model, stages of recovery, and motivational enhancement.",
        "Session 4: Coping Strategies and Life Skills - Discussion of general coping strategies, stress management (without meditation), and life skills for recovery.",
        "Session 5: Social Support and Communication - Education on building healthy social networks, communication skills, and managing relationships in recovery.",
        "Session 6: Relapse Awareness and Planning - Information on relapse warning signs, high-risk situations, and general planning for sustained recovery.",
    ]
    for s in control_sessions:
        paras.append(make_body(s))

    # Data Analysis Plan
    paras.append(make_heading2("Data Analysis Plan"))
    paras.append(make_body(
        "Data will be analyzed using the following statistical procedures:"
    ))
    paras.append(make_body(
        "1. Descriptive Statistics: Mean, standard deviation, and frequency distributions for "
        "demographic and clinical variables in both groups."
    ))
    paras.append(make_body(
        "2. Preliminary Analyses: Assessment of normality (Shapiro-Wilk test), homogeneity of "
        "variance (Levene's test), and baseline equivalence between groups (independent samples t-test)."
    ))
    paras.append(make_body(
        "3. Primary Analysis: Analysis of Covariance (ANCOVA) with post-test scores as the dependent "
        "variable, group (Experimental vs. Control) as the independent variable, and pre-test scores "
        "as the covariate. Separate ANCOVA models will be conducted for each dependent variable "
        "(craving, impulsivity, mindfulness)."
    ))
    paras.append(make_body(
        "4. Within-Group Analyses: Paired samples t-tests to examine pre-to-post changes within "
        "each group separately."
    ))
    paras.append(make_body(
        "5. Effect Sizes: Cohen's d will be calculated for all significant effects to quantify "
        "the magnitude of intervention effects."
    ))
    paras.append(make_body(
        "6. Significance Level: Alpha will be set at 0.05 (two-tailed) for all analyses. "
        "Bonferroni correction will be applied where multiple comparisons are conducted."
    ))
    paras.append(make_body(
        "Statistical analyses will be performed using SPSS Version 26.0 or JASP."
    ))

    return paras



def build_ethical_considerations():
    """Ethical Considerations section."""
    paras = []
    paras.append(make_heading1("Ethical Considerations"))
    paras.append(make_body(
        "The present study will adhere to the following ethical principles and guidelines:"
    ))
    paras.append(make_body(
        "1. Ethical Approval: The study protocol will be submitted to and approved by the Institutional "
        "Ethics Committee prior to participant recruitment."
    ))
    paras.append(make_body(
        "2. Informed Consent: Written informed consent will be obtained from all participants after "
        "providing a comprehensive explanation of the study purpose, procedures, potential risks and "
        "benefits, voluntary nature of participation, and right to withdraw at any time without "
        "penalty or impact on their ongoing treatment."
    ))
    paras.append(make_body(
        "3. Confidentiality: All participant data will be coded with unique identification numbers. "
        "Identifying information will be stored separately from research data in a locked cabinet "
        "accessible only to the principal investigator. Electronic data will be password-protected."
    ))
    paras.append(make_body(
        "4. Right to Withdraw: Participants may withdraw from the study at any point without any "
        "negative consequences for their ongoing treatment."
    ))
    paras.append(make_body(
        "5. No Harm Principle: The intervention (Brief MBRP) is a well-established, low-risk "
        "psychological intervention. If any participant experiences significant psychological "
        "distress during sessions, appropriate clinical support will be provided immediately."
    ))
    paras.append(make_body(
        "6. Equity: Control group participants will be offered the Brief MBRP intervention "
        "after completion of the study period (waitlist provision)."
    ))
    paras.append(make_body(
        "7. Compliance with Guidelines: The study will comply with the Declaration of Helsinki "
        "(2013), ICMR National Ethical Guidelines for Biomedical and Health Research Involving "
        "Human Participants (2017), and RCI Code of Ethics for Psychologists."
    ))
    return paras



def build_expected_results():
    """Expected Results section."""
    paras = []
    paras.append(make_heading1("Expected Results"))
    paras.append(make_body(
        "Based on the existing empirical evidence reviewed above, the following results are expected:"
    ))
    paras.append(make_body(
        "1. Participants in the Experimental Group (Brief MBRP + TAU) are expected to show significant "
        "reductions in craving scores (OCDUS) from pre-test to post-test, with the magnitude of "
        "reduction being significantly greater than that observed in the Control Group (Psychoeducation "
        "+ TAU)."
    ))
    paras.append(make_body(
        "2. Participants in the Experimental Group are expected to demonstrate significant reductions "
        "in impulsivity scores (BIS-11) from pre-test to post-test, with greater reduction compared "
        "to the Control Group."
    ))
    paras.append(make_body(
        "3. Participants in the Experimental Group are expected to show significant increases in "
        "mindfulness scores (FFMQ) from pre-test to post-test, with greater enhancement compared "
        "to the Control Group."
    ))
    paras.append(make_body(
        "4. The effect sizes for Brief MBRP are expected to be in the medium range (d = 0.50-0.80), "
        "consistent with meta-analytic findings (Li et al., 2017)."
    ))
    paras.append(make_body(
        "5. Craving reduction is expected to show the largest effect size among the three dependent "
        "variables, consistent with the meta-analytic finding of d = 0.68 for craving outcomes."
    ))
    return paras



def build_limitations():
    """Limitations section."""
    paras = []
    paras.append(make_heading1("Limitations"))
    paras.append(make_body(
        "The following limitations of the present study are acknowledged:"
    ))
    paras.append(make_body(
        "1. Gender Restriction: The sample is limited to male patients, which restricts "
        "generalizability to female substance-dependent populations who may respond differently "
        "to mindfulness interventions."
    ))
    paras.append(make_body(
        "2. Single-Site Design: Data collection from a single treatment facility limits "
        "generalizability to other settings and populations."
    ))
    paras.append(make_body(
        "3. Absence of Long-Term Follow-Up: The study assesses outcomes immediately post-intervention "
        "without longer-term follow-up, precluding conclusions about durability of effects."
    ))
    paras.append(make_body(
        "4. Self-Report Measures: Reliance on self-report instruments for all dependent variables "
        "introduces potential social desirability bias and shared method variance."
    ))
    paras.append(make_body(
        "5. Heterogeneous Substance Types: Including participants with dependence on different "
        "substances may increase within-group variability and mask substance-specific effects."
    ))
    paras.append(make_body(
        "6. Brief Protocol Duration: The condensed 3-week, 6-session format may not allow "
        "sufficient time for deeper mindfulness skills to develop compared to the standard "
        "8-session protocol."
    ))
    paras.append(make_body(
        "7. Non-Blinding: Due to the nature of psychotherapeutic interventions, participant "
        "blinding is not possible, potentially introducing expectancy effects. However, assessor "
        "blinding will be maintained where feasible."
    ))
    return paras



def build_future_directions():
    """Future Directions section."""
    paras = []
    paras.append(make_heading1("Future Directions"))
    paras.append(make_body(
        "Future research extending the present study may consider the following directions:"
    ))
    paras.append(make_body(
        "1. Longitudinal follow-up assessments at 3, 6, and 12 months to evaluate the durability "
        "of Brief MBRP effects and actual relapse rates."
    ))
    paras.append(make_body(
        "2. Inclusion of female participants and examination of gender differences in response to "
        "mindfulness-based interventions for substance dependence."
    ))
    paras.append(make_body(
        "3. Multi-site replication across diverse geographic and cultural contexts within India."
    ))
    paras.append(make_body(
        "4. Integration of neurobiological measures (e.g., fMRI, EEG) to examine neural correlates "
        "of Brief MBRP effects on craving and impulsivity."
    ))
    paras.append(make_body(
        "5. Dose-response studies comparing 4-session, 6-session, and 8-session protocols to "
        "identify the optimal intervention length."
    ))
    paras.append(make_body(
        "6. Mediational analyses to identify specific mindfulness facets that drive craving "
        "reduction and impulsivity changes."
    ))
    paras.append(make_body(
        "7. Technology-enhanced delivery (app-based guided meditation supplements) to enhance "
        "home practice compliance and extend intervention reach."
    ))
    return paras



def build_references():
    """References section - APA 7, alphabetical, hanging indent, italicized titles."""
    paras = []
    paras.append(make_heading1("References"))

    # Each reference: author portion normal, title in italics, then journal italics
    refs = [
        [("American Psychiatric Association. (2013). ", False, False),
         ("Diagnostic and statistical manual of mental disorders", False, True),
         (" (5th ed.). American Psychiatric Publishing.", False, False)],

        [("Baer, R. A., Smith, G. T., Hopkins, J., Krietemeyer, J., & Toney, L. (2006). Using self-report assessment methods to explore facets of mindfulness. ", False, False),
         ("Assessment, 13", False, True),
         ("(1), 27-45.", False, False)],

        [("Barratt, E. S. (1994). Impulsiveness and aggression. In J. Monahan & H. J. Steadman (Eds.), ", False, False),
         ("Violence and mental disorder: Developments in risk assessment", False, True),
         (" (pp. 61-79). University of Chicago Press.", False, False)],

        [("Bowen, S., Chawla, N., & Marlatt, G. A. (2011). ", False, False),
         ("Mindfulness-based relapse prevention for addictive behaviors: A clinician's guide", False, True),
         (". Guilford Press.", False, False)],

        [("Bowen, S., & Marlatt, G. A. (2009). Surfing the urge: Brief mindfulness-based intervention for college student smokers. ", False, False),
         ("Psychology of Addictive Behaviors, 23", False, True),
         ("(4), 666-671.", False, False)],

        [("Bowen, S., Witkiewitz, K., Clifasefi, S. L., Grow, J., Chawla, N., Hsu, S. H., Carroll, H. A., Harrop, E., Collins, S. E., Lustyk, M. K., & Larimer, M. E. (2014). Relative efficacy of mindfulness-based relapse prevention, standard relapse prevention, and treatment as usual for substance use disorders: A randomized clinical trial. ", False, False),
         ("JAMA Psychiatry, 71", False, True),
         ("(5), 547-556.", False, False)],

        [("Franken, I. H. A., Hendriks, V. M., & van den Brink, W. (2002). New perspectives on the Obsessive-Compulsive Drug Use Scale. ", False, False),
         ("European Addiction Research, 8", False, True),
         ("(4), 200-204.", False, False)],

        [("Garland, E. L., Froeliger, B., & Howard, M. O. (2014). Mindfulness training targets neurocognitive mechanisms of addiction at the attention-appraisal-emotion interface. ", False, False),
         ("Frontiers in Psychiatry, 4", False, True),
         (", Article 173.", False, False)],

        [("Garland, E. L., Roberts-Lewis, A., Tronnier, C. D., Graves, R., & Kelley, K. (2016). Mindfulness-Oriented Recovery Enhancement versus CBT for co-occurring substance dependence, traumatic stress, and psychiatric disorders. ", False, False),
         ("Journal of Consulting and Clinical Psychology, 84", False, True),
         ("(4), 281-293.", False, False)],

        [("Ghosh, A., Basu, D., & Avasthi, A. (2018). Relapse in opioid dependence: An Indian perspective. ", False, False),
         ("Indian Journal of Psychiatry, 60", False, True),
         ("(Suppl 4), S469-S476.", False, False)],
    ]

    for ref in refs:
        paras.append(make_rich_para(ref, hanging=True))

    return paras



def build_references_2():
    """References continued."""
    paras = []

    refs = [
        [("Glasner-Edwards, S., Mooney, L. J., Ang, A., Garneau, H. C., Hartwell, E., Brecht, M. L., & Rawson, R. A. (2017). Mindfulness-based relapse prevention for stimulant dependent adults: A pilot randomized clinical trial. ", False, False),
         ("Mindfulness, 8", False, True),
         ("(1), 126-135.", False, False)],

        [("Kabat-Zinn, J. (1994). ", False, False),
         ("Wherever you go, there you are: Mindfulness meditation in everyday life", False, True),
         (". Hyperion.", False, False)],

        [("Li, W., Howard, M. O., Garland, E. L., McGovern, P., & Lazar, M. (2017). Mindfulness treatment for substance misuse: A systematic review and meta-analysis. ", False, False),
         ("Journal of Substance Abuse Treatment, 75", False, True),
         (", 62-96.", False, False)],

        [("Marlatt, G. A. (2002). Buddhist philosophy and the treatment of addictive behavior. ", False, False),
         ("Cognitive and Behavioral Practice, 9", False, True),
         ("(1), 44-50.", False, False)],

        [("Marlatt, G. A., & Gordon, J. R. (1985). ", False, False),
         ("Relapse prevention: Maintenance strategies in the treatment of addictive behaviors", False, True),
         (". Guilford Press.", False, False)],

        [("Ministry of Social Justice and Empowerment. (2019). ", False, False),
         ("Magnitude of substance use in India", False, True),
         (". Government of India.", False, False)],

        [("Murphy, C., & MacKillop, J. (2012). Living in the here and now: Interrelationships between impulsivity, mindfulness, and alcohol misuse. ", False, False),
         ("Psychopharmacology, 219", False, True),
         ("(2), 527-536.", False, False)],

        [("Patton, J. H., Stanford, M. S., & Barratt, E. S. (1995). Factor structure of the Barratt Impulsiveness Scale. ", False, False),
         ("Journal of Clinical Psychology, 51", False, True),
         ("(6), 768-774.", False, False)],

        [("Sarkar, S., & Balhara, Y. P. S. (2016). Systematic review of mindfulness-based interventions for substance use in the Indian context. ", False, False),
         ("Indian Journal of Psychiatry, 58", False, True),
         ("(3), 290-295.", False, False)],

        [("Witkiewitz, K., Bowen, S., Douglas, H., & Hsu, S. H. (2013). Mindfulness-based relapse prevention for substance craving. ", False, False),
         ("Addictive Behaviors, 38", False, True),
         ("(2), 1563-1571.", False, False)],
    ]

    for ref in refs:
        paras.append(make_rich_para(ref, hanging=True))

    # Additional references
    more_refs = [
        [("WHO ASSIST Working Group. (2002). The Alcohol, Smoking and Substance Involvement Screening Test (ASSIST): Development, reliability and feasibility. ", False, False),
         ("Addiction, 97", False, True),
         ("(9), 1183-1194.", False, False)],

        [("Witkiewitz, K., & Bowen, S. (2010). Depression, craving, and substance use following a randomized trial of mindfulness-based relapse prevention. ", False, False),
         ("Journal of Consulting and Clinical Psychology, 78", False, True),
         ("(3), 362-374.", False, False)],

        [("Witkiewitz, K., Marlatt, G. A., & Walker, D. (2005). Mindfulness-based relapse prevention for alcohol and substance use disorders. ", False, False),
         ("Journal of Cognitive Psychotherapy, 19", False, True),
         ("(3), 211-228.", False, False)],

        [("Zgierska, A., Rabago, D., Chawla, N., Kushner, K., Koehler, R., & Marlatt, A. (2009). Mindfulness meditation for substance use disorders: A systematic review. ", False, False),
         ("Substance Abuse, 30", False, True),
         ("(4), 266-294.", False, False)],

        [("Brewer, J. A., Elwafi, H. M., & Davis, J. H. (2013). Craving to quit: Psychological models and neurobiological mechanisms of mindfulness training as treatment for addictions. ", False, False),
         ("Psychology of Addictive Behaviors, 27", False, True),
         ("(2), 366-379.", False, False)],

        [("Chiesa, A., & Serretti, A. (2014). Are mindfulness-based interventions effective for substance use disorders? A systematic review of the evidence. ", False, False),
         ("Substance Use & Misuse, 49", False, True),
         ("(5), 492-512.", False, False)],

        [("de Dios, M. A., Herman, D. S., Britton, W. B., Hagerty, C. E., Anderson, B. J., & Stein, M. D. (2012). Motivational and mindfulness intervention for young adult female marijuana users. ", False, False),
         ("Journal of Substance Abuse Treatment, 42", False, True),
         ("(1), 56-64.", False, False)],

        [("Grant, S., Colaiaco, B., Motala, A., Shanman, R., Booth, M., Sorbero, M., & Hempel, S. (2017). Mindfulness-based relapse prevention for substance use disorders: A systematic review and meta-analysis. ", False, False),
         ("Journal of Addiction Medicine, 11", False, True),
         ("(5), 386-396.", False, False)],

        [("Hsu, S. H., Collins, S. E., & Marlatt, G. A. (2013). Examining psychometric properties of distress tolerance and its moderation of mindfulness-based relapse prevention effects on alcohol and other drug use outcomes. ", False, False),
         ("Addictive Behaviors, 38", False, True),
         ("(3), 1852-1858.", False, False)],

        [("Katz, D., & Toner, B. (2013). A systematic review of gender differences in the effectiveness of mindfulness-based treatments for substance use disorders. ", False, False),
         ("Mindfulness, 4", False, True),
         ("(4), 318-331.", False, False)],

        [("Priddy, S. E., Howard, M. O., Hanley, A. W., Riquino, M. R., Friberg-Felsted, K., & Garland, E. L. (2018). Mindfulness meditation in the treatment of substance use disorders and preventing future relapse: Neurocognitive mechanisms and clinical implications. ", False, False),
         ("Substance Abuse and Rehabilitation, 9", False, True),
         (", 103-114.", False, False)],

        [("Tang, Y. Y., Tang, R., & Posner, M. I. (2016). Mindfulness meditation improves emotion regulation and reduces drug abuse. ", False, False),
         ("Drug and Alcohol Dependence, 163", False, True),
         ("(Suppl 1), S13-S18.", False, False)],
    ]

    for ref in more_refs:
        paras.append(make_rich_para(ref, hanging=True))

    return paras



# ============================================================
# ASSEMBLE DOCUMENT
# ============================================================

def build_document_xml():
    """Assemble all sections into the document.xml content."""
    all_paras = []

    # 1. Title Page
    all_paras.extend(build_title_page())

    # 2. Introduction (no "Introduction" heading per APA 7)
    all_paras.extend(build_introduction())
    all_paras.extend(build_mbrp_section())
    all_paras.extend(build_variables_section())

    # Page break before ROL
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 3. Review of Literature
    all_paras.extend(build_review_of_literature())
    all_paras.extend(build_review_of_literature_2())
    all_paras.extend(build_review_of_literature_3())
    all_paras.extend(build_review_of_literature_4())
    all_paras.extend(build_review_of_literature_5())

    # Page break
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 4. Research Gap
    all_paras.extend(build_research_gap())

    # 5. Aim
    all_paras.extend(build_aim())

    # 6. Objectives
    all_paras.extend(build_objectives())

    # 7. Hypotheses
    all_paras.extend(build_hypotheses())

    # Page break
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 8. Operational Definitions
    all_paras.extend(build_operational_definitions())

    # Page break
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 9. Methodology
    all_paras.extend(build_methodology())
    all_paras.extend(build_methodology_2())
    all_paras.extend(build_methodology_3())
    all_paras.extend(build_methodology_4())
    all_paras.extend(build_methodology_5())

    # Page break
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 10. Ethical Considerations
    all_paras.extend(build_ethical_considerations())

    # 11. Expected Results
    all_paras.extend(build_expected_results())

    # 12. Limitations
    all_paras.extend(build_limitations())

    # 13. Future Directions
    all_paras.extend(build_future_directions())

    # Page break
    all_paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    # 14. References
    all_paras.extend(build_references())
    all_paras.extend(build_references_2())

    # Join all paragraphs
    body_content = "\n".join(all_paras)

    # Build complete document.xml
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NAMESPACES}>
  <w:body>
    {body_content}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    return doc_xml



# ============================================================
# CREATE THE .DOCX FILE
# ============================================================

def create_docx():
    """Create the .docx file as a ZIP archive with Open XML content."""
    print("Generating MBRP Research Synopsis .docx...")

    # Build all XML content
    content_types = make_content_types()
    rels = make_rels()
    word_rels = make_word_rels()
    styles = make_styles()
    settings = make_settings()
    numbering = make_numbering()
    document = build_document_xml()

    # Create the ZIP file (.docx)
    with zipfile.ZipFile(OUTPUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('word/_rels/document.xml.rels', word_rels)
        zf.writestr('word/document.xml', document)
        zf.writestr('word/styles.xml', styles)
        zf.writestr('word/settings.xml', settings)
        zf.writestr('word/numbering.xml', numbering)

    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"Successfully created: {OUTPUT_PATH}")
    print(f"File size: {file_size:,} bytes ({file_size/1024:.1f} KB)")
    print("Document includes:")
    print("  - Title Page")
    print("  - Introduction (Substance Dependence, MBRP, Variables)")
    print("  - Review of Literature (10 studies, APA style with paper titles)")
    print("  - Research Gap (5 gaps)")
    print("  - Aim of the Study")
    print("  - Objectives (6 objectives)")
    print("  - Hypotheses (6 null hypotheses)")
    print("  - Operational Definitions")
    print("  - Methodology (Design, Sample, Criteria, Tools, Procedure, Intervention, Analysis)")
    print("  - Ethical Considerations")
    print("  - Expected Results")
    print("  - Limitations (7)")
    print("  - Future Directions")
    print("  - References (32+ entries, APA 7 format)")


if __name__ == "__main__":
    create_docx()
