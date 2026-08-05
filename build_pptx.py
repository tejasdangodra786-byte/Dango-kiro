#!/usr/bin/env python3
"""
Build a new MBRP Research Synopsis PPTX from scratch using raw XML.
No external dependencies needed - uses only zipfile and xml.etree.
"""
import zipfile
import os
import copy

# EMU constants (English Metric Units)
EMU_PER_INCH = 914400
SLIDE_WIDTH = 12192000   # 13.33 inches
SLIDE_HEIGHT = 6858000   # 7.5 inches



def make_slide_xml(title, bullets, title_color="1B2A4A", accent_color="2E86AB"):
    """Create a slide XML with title and bullet points."""
    # Build bullet paragraphs
    bullet_paras = ""
    for b in bullets:
        # Check if it's a sub-bullet (starts with spaces or tab)
        level = "0"
        text = b
        if b.startswith("  ") or b.startswith("\t"):
            level = "1"
            text = b.strip()
        bold_start = ""
        bold_end = ""
        # If text has ** markers for bold
        if "**" in text:
            parts = text.split("**")
            para_runs = ""
            for i, part in enumerate(parts):
                if i % 2 == 1:  # bold part
                    para_runs += f'<a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(part)}</a:t></a:r>'
                else:
                    if part:
                        para_runs += f'<a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="2C3E50"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(part)}</a:t></a:r>'
            bullet_paras += f'<a:p><a:pPr lvl="{level}" marL="{int(level)*457200}" indent="-228600"><a:buChar char="&#x2022;"/></a:pPr>{para_runs}</a:p>'
        else:
            bullet_paras += f'<a:p><a:pPr lvl="{level}" marL="{int(level)*457200}" indent="-228600"><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="2C3E50"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(text)}</a:t></a:r></a:p>'

    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="457200" y="152400"/><a:ext cx="11277600" cy="685800"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
</p:spPr>
<p:txBody>
<a:bodyPr anchor="ctr"/>
<a:lstStyle/>
<a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="2200" b="1" dirty="0"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(title)}</a:t></a:r></a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Accent Bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="457200" y="838200"/><a:ext cx="11277600" cy="45720"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>
</p:spPr>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="4" name="Content"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="457200" y="1000000"/><a:ext cx="11277600" cy="5500000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
</p:spPr>
<p:txBody>
<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:normAutofit/></a:bodyPr>
<a:lstStyle/>
{bullet_paras}
</p:txBody>
</p:sp>
</p:spTree>
</p:cSld>
</p:sld>'''
    return xml



def escape_xml(text):
    """Escape XML special characters."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

def make_title_slide_xml(title, subtitle, meta_lines):
    """Create a title/cover slide with gradient-like dark background."""
    meta_paras = ""
    for line in meta_lines:
        meta_paras += f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="B0BEC5"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(line)}</a:t></a:r></a:p>'
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1B2A4A"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="1371600"/><a:ext cx="10363200" cy="1828800"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
</p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2600" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(title)}</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="3200400"/><a:ext cx="10363200" cy="457200"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
</p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1600" dirty="0"><a:solidFill><a:srgbClr val="7DD3C0"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>{escape_xml(subtitle)}</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="4" name="Meta"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="4114800"/><a:ext cx="10363200" cy="2286000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
</p:spPr>
<p:txBody><a:bodyPr anchor="t"/><a:lstStyle/>
{meta_paras}
</p:txBody></p:sp>
</p:spTree>
</p:cSld>
</p:sld>'''
    return xml



# Define all slides content
slides_data = []

# SLIDE 1: Title
slides_data.append(("title", "Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients",
    "A Pre-Test Post-Test Control Group Experimental Design",
    ["Researcher: Tejas Dangodra  |  Guide: Dr. Himanshi Singh",
     "Department of Clinical Psychology",
     "MAN College of Special Education & Psychological Studies | Man Nasha Mukti Kendra",
     "Krantivir Tatya Tope Vishwavidyalaya, Guna, MP | 2025-2027"]))

# SLIDE 2: Introduction 1
slides_data.append(("content", "Introduction: Substance Dependence in India", [
    "**Indian Context:** MAGNITUDE study (2019) estimated ~3.1 crore individuals affected by substance use disorders in India",
    "Opioid dependence is a major public health burden, particularly in Punjab, Rajasthan, NE India, and metropolitan areas",
    "WHO estimates India accounts for approximately 25% of global opioid-related deaths in South-East Asia",
    "**Treatment Gap:** Indian de-addiction centers (incl. Man Nasha Mukti Kendra) primarily offer pharmacotherapy (OST, naltrexone) with limited structured psychotherapy",
    "Psychosocial interventions remain under-utilized despite evidence of superior combined treatment outcomes",
    "**Relapse Crisis:** Relapse rates range from 40-60% within first year globally (NIDA, 2020); Indian opioid studies report 70-80% (Mattoo et al., 2009)",
    "Primary triggers: craving, negative affect, interpersonal conflict, environmental cues",
    "Brief interventions essential for Indian rehab settings with limited resources and 4-6 week admission windows"
]))

# SLIDE 3: Introduction - What is MBRP
slides_data.append(("content", "Introduction: What is MBRP?", [
    "**Mindfulness-Based Relapse Prevention (MBRP)** developed by Bowen, Chawla & Marlatt (2011) at University of Washington",
    "Integrates mindfulness meditation with cognitive-behavioral Relapse Prevention framework",
    "**Core Mechanism:** Teaches patients to observe craving and emotional distress as transient mental events without automatically acting on them",
    "**Key Technique - Urge Surfing:** Riding the wave of craving until it passes naturally",
    "**Neurobiological Basis:** Strengthens prefrontal cortex regulation over amygdala reactivity; enhances PFC-limbic connectivity",
    "**Brief B-MBRP Adaptation:** 6 structured sessions over 3 weeks, group format (6-8 patients), 45 min/session, twice weekly",
    "Feasible within typical Indian IPD admission window of 4-6 weeks",
    "**Three Targets:** Craving (urge surfing), Impulsivity (mindful pause), Mindfulness (structured meditation)"
]))

# SLIDE 4: Variable 1 - Craving
slides_data.append(("content", "Variable 1: Craving", [
    "**Clinical Definition:** Intense, often overwhelming desire to use a substance (DSM-5 & ICD-11 core feature)",
    "**Theoretical:** Obsessive thoughts + compulsive urges related to drug use (Franken et al., 2002)",
    "**Relevance:** Primary trigger for relapse; craving intensity predicts treatment dropout, lapse episodes, and full relapse (Tiffany & Wray, 2012)",
    "Traditional suppression paradoxically increases craving intensity through rebound effects",
    "**Neuropsychological Basis:** Mesolimbic dopamine pathway activation; Ventral striatum & OFC hyperactivity to drug cues",
    "Incentive sensitization (Robinson & Berridge, 1993); PFC hypoactivation during craving",
    "**MBRP Techniques Targeting Craving:**",
    "  Urge Surfing: Observe craving as transient wave that rises, peaks, and falls",
    "  SOBER Space: Stop-Observe-Breathe-Expand-Respond",
    "  Cognitive Decentering: 'I am having a craving' (distancing from automatic thought)",
    "  Mindfulness Exposure: Non-reactive awareness of craving sensations"
]))

# SLIDE 5: Variable 2 - Impulsivity
slides_data.append(("content", "Variable 2: Impulsivity", [
    "**Clinical Definition:** Rapid, unplanned actions without considering consequences (Moeller et al., 2001)",
    "**Theoretical:** Three dimensions - Motor + Attentional + Non-Planning Impulsivity (Patton et al., 1995)",
    "**Relevance in Substance Dependence:**",
    "  Higher trait impulsivity predicts relapse",
    "  Mediates craving-to-use behavior",
    "  Associated with treatment non-adherence and dropout",
    "  Both a risk factor for and a consequence of chronic substance use",
    "**Neuropsychological Basis:** PFC dysfunction leading to impaired executive control; Reduced inhibitory control (Go/No-Go paradigms); DLPFC hypoactivation",
    "**MBRP Techniques Targeting Impulsivity:**",
    "  Response Inhibition: Mindful pause before action",
    "  STOP Technique: Stop - Take a breath - Observe - Proceed",
    "  Awareness Training: Notice impulse-action sequences in real-time",
    "  Mindful Decision-Making: Creating space between stimulus and response"
]))



# SLIDE 6: Variable 3 - Mindfulness
slides_data.append(("content", "Variable 3: Mindfulness", [
    "**Clinical Definition:** Present-moment awareness with openness and non-judgment",
    "**Theoretical:** 5 facets - Observing, Describing, Acting with Awareness, Non-Judging, Non-Reactivity (Baer et al., 2006)",
    "**Relevance in Substance Dependence:**",
    "  Substance users show significantly lower dispositional mindfulness",
    "  Acts as protective factor against relapse triggers",
    "  Improvements in mindfulness mediate MBRP treatment outcomes",
    "**Neuropsychological Basis:** ACC activation enhances self-regulation; Insula mediates interoceptive awareness; PFC-amygdala connectivity improves emotional regulation; DMN regulation reduces rumination",
    "**MBRP Techniques for Building Mindfulness:**",
    "  Sitting Meditation: Focused attention on breath",
    "  Body Scan: Non-judgmental awareness of bodily sensations",
    "  Mindful Movement: Present-moment yoga/walking",
    "  Non-Judgmental Awareness: Labeling experiences without evaluation"
]))

# SLIDE 7-10: Review of Literature
slides_data.append(("content", "Review of Literature (1/5): MBRP & Relapse Prevention", [
    "**Bowen et al. (2014) - JAMA Psychiatry:** N=286 RCT. At 12-month follow-up, MBRP participants reported significantly fewer days of substance use and heavy drinking compared to standard RP and TAU. MBRP maintained superior long-term outcomes through cultivation of mindfulness skills as durable protective mechanism.",
    "",
    "**Bowen & Marlatt (2009) - Psychology of Addictive Behaviors:** Brief urge surfing meditation with incarcerated substance users. Significant reductions in craving intensity and frequency vs. controls. Validates that even brief mindfulness exposure can disrupt automaticity of craving responses.",
    "",
    "**Key Insight:** Both studies establish that MBRP integrates present-moment awareness with cognitive-behavioral strategies, creating synergistic therapeutic effects. Supports feasibility of 6-session B-MBRP protocols."
]))

slides_data.append(("content", "Review of Literature (2/5): Craving & Mindfulness Mechanisms", [
    "**Garland et al. (2014) - Frontiers in Psychiatry:** Demonstrates MORE reduces opioid craving through: (1) attentional reorientation away from drug cues, (2) positive reappraisal of neutral stimuli, (3) enhanced savoring of healthy pleasures. Neuroimaging shows mindfulness modulates prefrontal and limbic craving circuits.",
    "",
    "**Witkiewitz et al. (2013) - Addictive Behaviors:** Secondary RCT analysis. Over 4-month follow-up, MBRP participants showed lower craving AND the affect-craving pathway was significantly attenuated. Mindfulness weakens the link between negative emotions and automatic craving.",
    "",
    "**Key Mechanism:** MBRP decouples the affect-craving pathway. Negative emotions no longer automatically trigger craving responses - this distinguishes MBRP from traditional RP approaches."
]))

slides_data.append(("content", "Review of Literature (3/5): Impulsivity & Mindfulness", [
    "**Garland et al. (2016) - J Consulting & Clinical Psychology:** RCT with substance-dependent adults. Mindfulness-Oriented Recovery Enhancement produced significant BIS-11 reductions in motor and attentional impulsivity. Mechanism: mindfulness strengthens prefrontal cortical inhibitory control through sustained attention and deliberate non-reactivity training.",
    "",
    "**Murphy & MacKillop (2012) - Psychopharmacology:** N=340 cross-sectional. Trait mindfulness inversely associated with impulsive decision-making (delay discounting). Mindfulness moderated the impulsivity-substance use link. Higher mindfulness = weaker association between impulsivity and alcohol consequences.",
    "",
    "**Implication:** Mindfulness functions as a cognitive resource enabling impulsive individuals to override automatic behavioral tendencies through enhanced metacognitive awareness and response flexibility."
]))

slides_data.append(("content", "Review of Literature (4/5): Meta-Analyses & Brief Models", [
    "**Li et al. (2017) - J Substance Abuse Treatment:** Meta-analysis of 42 RCTs. Effect sizes: substance misuse (d=0.33), craving (d=0.68), stress (d=0.44). Brief interventions (4-8 sessions) showed comparable efficacy to longer protocols when appropriately structured.",
    "",
    "**Glasner-Edwards et al. (2017) - Mindfulness:** Pilot RCT of abbreviated 6-session MBRP for stimulant-dependent adults. Both feasible and effective in reducing substance use frequency and craving intensity vs. health education control.",
    "",
    "**Critical Evidence:** 6-session MBRP protocols can be successfully implemented without substantial loss of efficacy, making them appropriate for Indian settings with 4-6 week admission durations. Directly validates the brief intervention model proposed in present research."
]))

slides_data.append(("content", "Review of Literature (5/5): Indian Context", [
    "**Ghosh et al. (2018) - Indian J Psychiatry:** Relapse rates exceed 70% among opioid-dependent patients in North Indian de-addiction centers within 3 months post-discharge. Primary determinants: craving, peer influence, negative affect, lack of psychological aftercare.",
    "",
    "**Sarkar & Balhara (2016):** Highlights underutilization of structured psychological interventions in Indian de-addiction settings. Barriers: limited trained personnel, absent validated protocols, institutional emphasis on pharmacological management.",
    "",
    "**Jain et al. (2013):** Preliminary evidence from MBI with alcohol-dependent patients in India showing initial craving reductions.",
    "",
    "**Critical Gap:** Indian treatment facilities predominantly rely on pharmacological approaches with minimal integration of evidence-based psychological interventions specifically designed for relapse prevention."
]))



# SLIDE 12: Research Gap (REWRITTEN - more academic, less direct)
slides_data.append(("content", "Research Gap", [
    "**1. Limited evidence for brief MBRP adaptations in LMIC settings:** The majority of MBRP research evaluates the standard 8-week protocol in Western outpatient settings. Condensed protocols suitable for resource-constrained inpatient environments remain insufficiently examined (Li et al., 2017)",
    "",
    "**2. Absence of simultaneous multi-domain outcome assessment:** Existing studies typically examine craving, impulsivity, or mindfulness in isolation. The interplay between these three mechanistically linked variables within a single intervention framework has not been adequately investigated in Indian clinical populations",
    "",
    "**3. Paucity of culturally contextualized MBRP evidence from India:** Despite India's substantial opioid dependence burden, empirical validation of MBRP-based protocols in Indian de-addiction infrastructure remains virtually absent from the published literature",
    "",
    "**4. Insufficient integration of psychological interventions in Indian de-addiction practice:** Current treatment models at district-level centers remain predominantly pharmacological, with limited evidence-based psychosocial adjuncts (Sarkar & Balhara, 2016; Ghosh et al., 2018)",
    "",
    "**5. Need for attention-matched controlled designs:** Most existing mindfulness studies in Indian addiction settings lack active control conditions, making it difficult to distinguish mindfulness-specific effects from non-specific therapeutic factors (therapist contact, group support, attention)",
    "",
    "**PRESENT STUDY:** Addresses these gaps as the first brief MBRP trial (6 sessions/3 weeks) in an Indian opioid-dependent sample with active control and simultaneous three-variable assessment"
]))

# SLIDE 13: Aim
slides_data.append(("content", "Aim of the Study", [
    "",
    "**PRIMARY AIM:**",
    "",
    "To evaluate the efficacy of a Brief Mindfulness-Based Relapse Prevention (B-MBRP) intervention (6 sessions over 3 weeks) in reducing craving and impulsivity, and enhancing mindfulness, among male substance-dependent patients at Man Nasha Mukti Kendra",
    "",
    "**Comparing:** Brief MBRP + TAU (Experimental) vs. Psychoeducation + TAU (Active Control)",
    "",
    "**Key Parameters:** N = 60 (30 per group) | 6 B-MBRP Sessions | 3 Weeks Duration | 3 Outcome Variables"
]))

# SLIDE 14: Objectives
slides_data.append(("content", "Objectives", [
    "**1.** To assess and compare craving levels (pre vs. post) in Experimental (Brief MBRP + TAU) and Control (Psychoeducation + TAU) groups",
    "",
    "**2.** To assess and compare impulsivity levels (pre vs. post) in both groups",
    "",
    "**3.** To assess and compare mindfulness levels (pre vs. post) in both groups",
    "",
    "**4.** To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing craving",
    "",
    "**5.** To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing impulsivity",
    "",
    "**6.** To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in enhancing mindfulness"
]))

# SLIDE 15: Hypotheses (NULL ONLY)
slides_data.append(("content", "Hypotheses (Null Hypotheses Only)", [
    "**H01 - Craving:**",
    "There is no significant difference in craving scores (OCDUS) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
    "",
    "**H02 - Impulsivity:**",
    "There is no significant difference in impulsivity scores (BIS-11) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
    "",
    "**H03 - Mindfulness:**",
    "There is no significant difference in mindfulness scores (FFMQ) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.",
    "",
    "**Statistical Testing:** Alpha = 0.05 (two-tailed) | Primary Analysis: ANCOVA with pre-test scores as covariates"
]))

# SLIDE 16: Operational Definitions 1
slides_data.append(("content", "Operational Definitions of Key Terms (1/2)", [
    "**Craving** in this study refers to the intensity and frequency of obsessive thoughts about drug use and compulsive urges to use opioids, as measured by the total score on the Obsessive Compulsive Drug Use Scale (OCDUS; Franken et al., 2002). Scores range from 0-48; higher scores indicate greater craving.",
    "",
    "**Impulsivity** in this study refers to the multidimensional tendency to act without adequate forethought, encompassing motor, attentional, and non-planning components, as measured by the total score and three subscale scores on the Barratt Impulsiveness Scale-11 (BIS-11; Patton et al., 1995). Scores range from 30-120; higher scores indicate greater impulsivity.",
    "",
    "**Mindfulness** in this study refers to the dispositional capacity for present-moment awareness with non-judgment and non-reactivity, as measured by the total score on the Five Facet Mindfulness Questionnaire (FFMQ; Baer et al., 2006) across five facets: Observing, Describing, Acting with Awareness, Non-Judging, Non-Reactivity. Scores range from 39-195; higher scores indicate greater mindfulness."
]))

# SLIDE 17: Operational Definitions 2
slides_data.append(("content", "Operational Definitions of Key Terms (2/2)", [
    "**Brief MBRP (B-MBRP)** in this study refers to a structured 6-session mindfulness-based relapse prevention intervention delivered twice weekly over 3 weeks (45 min/session) in group format (6-8 patients), adapted from Bowen, Chawla & Marlatt (2011), incorporating body scan, breath meditation, urge surfing, SOBER breathing space, and relapse prevention planning.",
    "",
    "**Substance Dependence** in this study refers to a clinical diagnosis of Substance Dependence Syndrome as per ICD-10 (F10-F19) criteria, with primary opioid dependence (heroin or pharmaceutical opioids), confirmed by a qualified psychiatrist at Man Nasha Mukti Kendra.",
    "",
    "**Treatment As Usual (TAU)** in this study refers to the standard pharmacological treatment at Man Nasha Mukti Kendra, including medically supervised detoxification, Opioid Substitution Therapy (buprenorphine/methadone), naltrexone maintenance, routine counseling, and daily ward activities.",
    "",
    "**Psychoeducation (Active Control)** in this study refers to 6 structured informational sessions (45 min, twice weekly, 3 weeks) covering addiction science, effects of opioids, relapse warning signs, health/nutrition, social consequences, and recovery motivation - matched for time and attention but containing NO mindfulness component."
]))



# SLIDE 18: Research Design
slides_data.append(("content", "Research Design: Pre-Test Post-Test Control Group", [
    "**Design Notation:**",
    "  R  O1  X1  O2  -->  Experimental Group (Brief MBRP + TAU)",
    "  R  O1  X2  O2  -->  Control Group (Psychoeducation + TAU)",
    "",
    "**Flow:** Randomization (N=60) -> Pre-Test (OCDUS + BIS-11 + FFMQ + ASSIST) -> Intervention (3 Weeks) -> Post-Test (OCDUS + BIS-11 + FFMQ)",
    "",
    "**Design Features:**",
    "  True experimental design with random assignment",
    "  Active control condition (attention-matched psychoeducation)",
    "  TAU continued for all participants throughout",
    "  Pre-post measurement at standardized time points",
    "",
    "**Setting:** Man Nasha Mukti Kendra, Guna, MP | Assessments by blinded Research Assistant | Intervention by MPhil Clinical Psychologist (Researcher)"
]))

# SLIDE 19: Sample
slides_data.append(("content", "Sample & Sampling Strategy", [
    "**Stage 1 - Purposive Selection:** Eligible participants identified based on inclusion/exclusion criteria from patients admitted to Man Nasha Mukti Kendra. Consecutive sampling of all eligible patients.",
    "",
    "**Stage 2 - Random Assignment:** Computer-generated randomization allocating eligible participants to Experimental (n=30) or Control (n=30) using sealed opaque envelopes.",
    "",
    "**Sample:** N = 60 total (30 per group) | Recruit 70 (35/group) for attrition | Male | Age 18-50 | Opioid-dependent | Detoxified",
    "",
    "**Male-Only Justification:**",
    "  Indian de-addiction centers (incl. Man Nasha Mukti Kendra) admit ~90-95% male patients",
    "  MAGNITUDE study reports male:female ratio ~10:1 for opioid dependence in India",
    "  Gender differences in craving, impulsivity, mindfulness may confound if mixed sample used",
    "  Homogeneous sample strengthens internal validity for this initial efficacy trial",
    "  Female-specific MBRP studies recommended as future direction"
]))

# SLIDE 20: Sample Size
slides_data.append(("content", "Sample Size Estimation", [
    "**Formula:** n = [(Za/2 + Zb)^2 x 2 x sigma^2] / d^2",
    "",
    "**Parameters:** Effect size (d) = 0.50 (medium) | Power (1-beta) = 0.80 -> Zb = 0.84 | Alpha = 0.05 (two-tailed) -> Za/2 = 1.96",
    "",
    "**Calculation:** n = [(1.96 + 0.84)^2 x 2 x 1] / (0.50)^2 = [7.84 x 2] / 0.25 = 62.72 ~ 63 total (~32 per group)",
    "",
    "**FINAL DECISION: N = 60 (30 per group)**",
    "  ANCOVA as primary analysis reduces required n",
    "  G*Power 3.1 verification: ANCOVA with 1 covariate -> ~34/group sufficient",
    "  Recruit 70 total (35/group) to account for ~15% attrition",
    "  Consistent with: Glasner-Edwards et al. (2017), Bowen & Marlatt (2009)"
]))

# SLIDE 21: Inclusion Criteria
slides_data.append(("content", "Inclusion Criteria", [
    "**1.** Diagnosis of Substance Dependence as per ICD-10 (F10-F19) / ICD-11 criteria",
    "**2.** Primary substance: Opioids (heroin, pharmaceutical opioids); polysubstance with primary opioid dependence included",
    "**3.** Male participants aged 18-50 years",
    "**4.** Completed detoxification phase (minimum 7 days post-withdrawal)",
    "**5.** Currently admitted at Man Nasha Mukti Kendra",
    "**6.** Minimum education: 5th standard (ability to comprehend psychometric tools)",
    "**7.** Willingness to provide written informed consent",
    "**8.** Able to attend all 6 intervention sessions during 3-week period"
]))

# SLIDE 22: Exclusion Criteria
slides_data.append(("content", "Exclusion Criteria", [
    "**1.** Severe psychiatric comorbidity: Psychotic disorders, Bipolar I, severe MDE with suicidality",
    "**2.** Significant cognitive impairment (MMSE < 24) or intellectual disability",
    "**3.** Active withdrawal symptoms (COWS score > 12)",
    "**4.** History of traumatic brain injury with LOC > 30 minutes",
    "**5.** Current participation in another structured psychological intervention study",
    "**6.** Medical instability requiring acute/intensive care",
    "**7.** History of prior formal mindfulness/meditation training exceeding 1 month"
]))

# SLIDE 23: Variables
slides_data.append(("content", "Variables of the Study", [
    "**INDEPENDENT VARIABLE (IV):** Type of Intervention",
    "  Level 1: Brief MBRP + TAU (Experimental Group)",
    "  Level 2: Psychoeducation + TAU (Control Group)",
    "",
    "**DEPENDENT VARIABLES:**",
    "  DV 1: Craving - measured by OCDUS total score (range 0-48)",
    "  DV 2: Impulsivity - measured by BIS-11 total + 3 subscales (range 30-120)",
    "  DV 3: Mindfulness - measured by FFMQ total + 5 facets (range 39-195)",
    "",
    "**CONTROLLED VARIABLES:** Age, education, duration of use, severity (ASSIST baseline), TAU components constant across groups, session duration equalized (45 min x 6 sessions for both groups)"
]))

# SLIDE 24-26: Tools
slides_data.append(("content", "Tool 1: OCDUS (Craving Measure)", [
    "**Obsessive Compulsive Drug Use Scale (OCDUS) - Franken et al. (2002)**",
    "",
    "Description: 12-item self-report measuring obsessive thoughts about drug use and compulsive urges",
    "Scoring: 5-point scale (0-4) | Total range: 0-48 | Higher = greater craving",
    "Subscales: (1) Obsessive thoughts/interference, (2) Desire/control, (3) Resistance to thoughts",
    "",
    "**Psychometric Properties:**",
    "  Internal consistency: alpha = 0.86-0.90",
    "  Test-retest reliability: r = 0.78",
    "  Convergent validity with VAS craving: r = 0.55-0.67",
    "  Predicts relapse; sensitive to treatment changes",
    "",
    "**Justification:** Captures both cognitive (obsessive) and behavioral (compulsive) craving dimensions. Applicable across substances including opioids. Brief (5 min). Adaptable for Hindi."
]))

slides_data.append(("content", "Tool 2: BIS-11 & Tool 3: FFMQ", [
    "**Barratt Impulsiveness Scale (BIS-11) - Patton et al., 1995:**",
    "  30 items | 4-point scale (1-4) | Range: 30-120",
    "  3 factors: Attentional, Motor, Non-Planning Impulsivity",
    "  Higher = greater impulsivity | alpha = 0.79-0.83; test-retest r = 0.83",
    "  Discriminates SUD from controls | Hindi validated version available",
    "",
    "**Five Facet Mindfulness Questionnaire (FFMQ) - Baer et al., 2006:**",
    "  39 items | 5-point scale (1-5) | Range: 39-195",
    "  5 facets: Observing, Describing, Acting with Awareness, Non-Judging, Non-Reactivity",
    "  Higher = greater mindfulness | alpha = 0.75-0.91 across facets",
    "  Sensitive to mindfulness interventions | Hindi adaptation available",
    "",
    "**Both tools:** Self-report | Validated in substance use populations | Hindi available | Administered at PRE-TEST and POST-TEST | 10-15 min each"
]))

slides_data.append(("content", "Tool 4: WHO-ASSIST (Baseline Severity)", [
    "**Alcohol, Smoking and Substance Involvement Screening Test (ASSIST) - WHO, 2002:**",
    "",
    "Description: 8-item screening for risk level across 10 substance categories",
    "Scoring: Substance-specific risk -> Low (0-3), Moderate (4-26), High (27+)",
    "Reliability: Test-retest r = 0.58-0.90; alpha = 0.77-0.94",
    "Validity: Sensitivity = 0.80, Specificity = 0.71 for substance dependence",
    "",
    "**Indian Usability:** WHO-validated; Hindi version available; used in NDDTC studies; 5-10 min administration",
    "",
    "**PURPOSE IN THIS STUDY:** Used at PRE-TEST ONLY to establish baseline severity and ensure group equivalence. NOT used as an outcome measure."
]))



# SLIDE 27: Data Collection Procedure
slides_data.append(("content", "Data Collection Procedure: Step-by-Step", [
    "**STEP 1 - Screening (Day 1-3):** Review admission records. Identify male patients (18-50) with opioid dependence (ICD-10) who have completed 7+ days detox. Check MMSE >= 24, COWS <= 12. Apply inclusion/exclusion criteria.",
    "**STEP 2 - Informed Consent (Day 3-4):** Individual meeting in Hindi. Explain study purpose, procedures, duration, voluntary nature, right to withdraw, confidentiality. Obtain written signed consent.",
    "**STEP 3 - Pre-Test Assessment (Day 4-5):** Administer in quiet room: ASSIST (severity baseline) + OCDUS (craving) + BIS-11 (impulsivity) + FFMQ (mindfulness). ~40 min total. Conducted by blinded Research Assistant.",
    "**STEP 4 - Randomization (Day 5):** Computer-generated random sequence. Sealed opaque envelopes. Allocate to Experimental (MBRP) or Control (Psychoeducation). Stratified by ASSIST severity.",
    "**STEP 5 - Intervention (Weeks 1-3):** 6 sessions, twice weekly, 45 min each. Groups of 6-8. MBRP group: mindfulness techniques. Control: psychoeducation. Both continue TAU.",
    "**STEP 6 - Post-Test (Within 1 week of final session):** Re-administer OCDUS + BIS-11 + FFMQ only (no ASSIST). Same blinded RA. Same conditions as pre-test."
]))

# SLIDE 28: Data Collection Details
slides_data.append(("content", "Data Collection: Techniques & Conditions", [
    "**Assessment Setting:**",
    "  Quiet, private room within Man Nasha Mukti Kendra",
    "  Comfortable seating, adequate lighting, no distractions",
    "  Morning hours (9-12 AM) to avoid medication effects",
    "  Same room and conditions for pre and post assessment",
    "",
    "**Administration Protocol:**",
    "  Researcher-assisted (read aloud if literacy issues)",
    "  Hindi language throughout | Standardized instructions read verbatim",
    "  Unlimited time (typically 35-45 min total) | Practice items completed first",
    "",
    "**Blinding & Bias Control:**",
    "  Assessor (RA) blinded to group allocation",
    "  Participants instructed not to reveal group assignment",
    "  Standardized order: OCDUS -> BIS-11 -> FFMQ",
    "  Social desirability managed via anonymous coding | No researcher present during assessment",
    "",
    "**Data Management:** Participant ID codes (no names on forms) | Double data entry | Locked cabinet for paper forms | Electronic data password-protected"
]))

# SLIDE 29: B-MBRP Techniques Table
slides_data.append(("content", "B-MBRP Techniques Used in the Study", [
    "**Body Scan Meditation:** Systematic attention to bodily sensations head to toe without judgment -> Targets: Mindfulness (Observing)",
    "**Breath Awareness:** Focused attention on natural breathing rhythm as anchor -> Targets: Mindfulness (Acting with Awareness)",
    "**Urge Surfing:** Observing craving as wave that rises, peaks, falls without acting -> Targets: Craving reduction",
    "**SOBER Breathing Space:** Stop-Observe-Breathe-Expand-Respond (3-min emergency technique) -> Targets: Impulsivity (response inhibition)",
    "**Trigger Mapping:** Identifying personal high-risk situations, people, places, emotions -> Targets: Craving awareness",
    "**Cognitive Decentering:** Labeling thoughts: 'I am having the thought that...' -> Targets: Impulsivity + Craving",
    "**Loving-Kindness Meditation:** Cultivating self-compassion, positive affect -> Targets: Mindfulness (Non-Judging)",
    "**Mindful Movement:** Gentle yoga/walking with full present-moment attention -> Targets: Mindfulness (Observing)"
]))

# SLIDE 30: Session-by-Session Protocol
slides_data.append(("content", "B-MBRP Session-by-Session Protocol (Experimental Group)", [
    "**Session 1 - Autopilot & Awareness:** Welcome (5 min) -> Raisin exercise (10 min) -> Discussion on automatic patterns in addiction (10 min) -> Brief body scan (12 min) -> Assign practice (8 min) | Home: 5 min daily body scan",
    "**Session 2 - Triggers & Body Scan:** Practice review (5 min) -> Full body scan (15 min) -> Personal trigger mapping (15 min) -> Discussion on body signals of craving (10 min) | Home: Body scan + trigger diary",
    "**Session 3 - Breath & SOBER Space:** Practice review (5 min) -> 10 min breath meditation -> Teach SOBER (10 min) -> Role-play in high-risk scenario (10 min) -> Integration (10 min) | Home: Breath meditation + 3x daily SOBER",
    "**Session 4 - Urge Surfing & Decentering:** Practice review (5 min) -> Guided urge surfing (15 min) -> Cognitive decentering exercise (10 min) -> Group sharing (15 min) | Home: Urge surfing when craving arises",
    "**Session 5 - Acceptance & Non-Reactivity:** Practice review (5 min) -> Open awareness meditation (12 min) -> Acceptance vs. avoidance discussion (10 min) -> Skillful action planning (10 min) -> Wrap-up (8 min) | Home: Daily open awareness",
    "**Session 6 - Integration & Maintenance:** Practice review (5 min) -> Loving-kindness meditation (10 min) -> Personal relapse prevention plan with mindfulness tools (15 min) -> Group feedback & closure (15 min) | Home: Personalized daily plan",
    "",
    "**Delivery:** Twice weekly | 45 min/session | Group (6-8) | Hindi audio meditations | Practice logs maintained"
]))

# SLIDE 31: How Each Session is Conducted
slides_data.append(("content", "How Each B-MBRP Session is Conducted (Standard Structure)", [
    "**Every session follows this 6-step format:**",
    "",
    "**Step 1 - Opening Meditation (5 min):** Brief breath awareness to settle the group. Participants sit comfortably with eyes closed. Therapist guides attention to breath in Hindi.",
    "**Step 2 - Practice Review (5 min):** Participants share home practice experiences. Troubleshoot difficulties. Normalize challenges. Non-judgmental discussion.",
    "**Step 3 - Core Meditation/Exercise (12-15 min):** Session-specific guided meditation (body scan, urge surfing, etc.). Conducted in Hindi using pre-recorded audio + live guidance by therapist.",
    "**Step 4 - Psychoeducation & Discussion (10-12 min):** Brief teaching on session theme. Group discussion linking mindfulness to personal recovery. Culturally relevant examples used.",
    "**Step 5 - Experiential Exercise (8-10 min):** Active practice or role-play (SOBER in high-risk scenario, trigger mapping on paper, decentering exercise with craving thoughts).",
    "**Step 6 - Closing & Home Practice Assignment (5 min):** Summary of key learning. Assign specific daily practice (5-10 min). Distribute audio recordings. Record attendance.",
    "",
    "**Therapist:** MPhil Clinical Psychologist trained in MBRP | Manualized protocol ensures fidelity"
]))

# SLIDE 32: Control Group
slides_data.append(("content", "Control Group: Psychoeducation (6 Sessions / 3 Weeks)", [
    "**Session 1 - Understanding Addiction:** Disease model; brain changes with chronic use; genetic & environmental factors",
    "**Session 2 - Effects of Opioids:** Physical consequences (liver, heart, immune); psychological effects; withdrawal timeline",
    "**Session 3 - Understanding Relapse:** Warning signs; high-risk situations; Marlatt's relapse model; cognitive distortions",
    "**Session 4 - Health & Nutrition:** Physical recovery during treatment; sleep hygiene; exercise benefits; nutrition",
    "**Session 5 - Social Consequences:** Family impact; legal issues; stigma; workplace; rehabilitation resources",
    "**Session 6 - Motivation & Goals:** Stages of change; recovery planning; goal setting; community resources; discharge planning",
    "",
    "**Active Control Design:** Matched for time (45 min), format (group 6-8), frequency (2x/week), attention, and therapist contact. Uses handouts + visual aids + discussion. Contains NO mindfulness component."
]))



# SLIDE 33: Data Analysis
slides_data.append(("content", "Data Analysis Plan", [
    "**Statistical Tests:** Descriptive (Mean, SD, frequencies) | Normality (Shapiro-Wilk) | Within-group (Paired t-test / Wilcoxon) | Between-group (Independent t-test / Mann-Whitney) | Primary: ANCOVA | Effect size: Partial eta-squared, Cohen's d | Alpha = 0.05 | SPSS 26.0",
    "",
    "**ANCOVA Model:** DV: Post-test scores (OCDUS / BIS-11 / FFMQ) | IV: Group (Exp vs. Control) | Covariate: Pre-test scores | ITT: Last Observation Carried Forward (LOCF)",
    "",
    "**ANCOVA Justification:**",
    "  (a) Controls for pre-existing baseline differences even after randomization",
    "  (b) Increases statistical power by reducing within-group error variance",
    "  (c) Provides more precise treatment effect estimate by adjusting post-test for baseline",
    "  (d) Reduces required sample size compared to independent t-test",
    "  (e) Recommended for pre-post designs by Tabachnick & Fidell (2013) and Field (2018)"
]))

# SLIDE 34: Ethics
slides_data.append(("content", "Ethical Considerations", [
    "**1. Informed Consent:** Written consent in Hindi; participants fully informed of purpose, procedures, risks, benefits",
    "**2. Voluntary Participation:** Right to withdraw anytime without penalty or impact on treatment at Man Nasha Mukti Kendra",
    "**3. Confidentiality:** Data coded with participant IDs; no identifying information in publications; secure locked storage",
    "**4. Non-Maleficence:** Control receives active psychoeducation (not waitlist/no treatment); TAU continued for all",
    "**5. Institutional Approval:** Ethical clearance from Institutional Ethics Committee (IEC) prior to data collection",
    "**6. Debriefing:** Control group offered brief MBRP orientation post-study completion",
    "**7. Compliance:** ICMR (2017) National Ethical Guidelines for Biomedical and Health Research"
]))

# SLIDE 35: Expected Results
slides_data.append(("content", "Expected Results", [
    "**CRAVING (OCDUS):** Significant REDUCTION expected in MBRP group vs. Control. Effect: d = 0.50-0.80. Mechanism: Urge surfing disrupts automatic craving-use cycle.",
    "",
    "**IMPULSIVITY (BIS-11):** Significant REDUCTION expected (Motor + Attentional subscales). Effect: d = 0.40-0.60. Mechanism: Mindfulness strengthens prefrontal inhibitory control.",
    "",
    "**MINDFULNESS (FFMQ):** Significant INCREASE expected in MBRP group (Acting with Awareness & Non-Reactivity facets). Effect: d = 0.50-0.70. Mechanism: Structured meditation cultivates dispositional mindfulness.",
    "",
    "**OVERALL:** Brief MBRP + TAU expected to demonstrate superiority across all 3 DVs. Null hypotheses expected to be rejected. Supports feasibility of 6-session B-MBRP at Man Nasha Mukti Kendra."
]))

# SLIDE 36: Clinical Implications
slides_data.append(("content", "Clinical Implications", [
    "**1.** Validates a brief MBRP model (6 sessions/3 weeks) feasible for Indian de-addiction settings with limited resources",
    "**2.** Provides evidence-based psychological intervention to complement pharmacotherapy (OST, naltrexone) in routine care",
    "**3.** Demonstrates mindfulness-based approaches are culturally compatible with Indian populations",
    "**4.** Addresses multiple relapse risk factors simultaneously (craving + impulsivity + mindfulness) through single protocol",
    "**5.** Supports task-shifting: Brief MBRP deliverable by MPhil-trained Clinical Psychologists in district-level settings",
    "**6.** Scalable model: If effective, can be disseminated to government de-addiction centers across MP and India"
]))

# SLIDE 37: Limitations
slides_data.append(("content", "Limitations", [
    "**1.** Male-only sample from single center (Man Nasha Mukti Kendra) limits generalizability to females and other settings",
    "**2.** Short-term assessment: Post-test immediately after 3-week intervention; no long-term follow-up",
    "**3.** Self-report measures (OCDUS, BIS-11, FFMQ) susceptible to social desirability bias",
    "**4.** No biological markers: Craving measured subjectively without physiological corroboration (cortisol, HRV)",
    "**5.** Therapist effects: Single therapist delivery may introduce confounds (mitigated by manualized protocol)",
    "**6.** Attention-matched control does not fully isolate mindfulness-specific mechanisms from non-specific factors",
    "**7.** Potential attrition despite over-recruitment planning (common in substance-dependent populations)"
]))

# SLIDE 38: Future Directions
slides_data.append(("content", "Future Directions", [
    "**Follow-Up & Replication:** 3-month and 6-month follow-up assessments; Multi-site RCTs across Indian de-addiction centers; Include female participants for gender-specific effects",
    "**Mechanism Research:** Neuroimaging (fMRI/EEG) for neural mechanisms; Mediator analysis: mindfulness as mediator; Dose-response: Compare 4 vs. 6 vs. 8 sessions",
    "**Comparative Studies:** MBRP vs. CBT vs. ACT in Indian samples; Technology-assisted app-based digital MBRP; Post-discharge booster sessions",
    "**Dissemination:** Develop Hindi MBRP training manual; Train counselors in government centers; Rural access via telehealth MBRP"
]))

# SLIDE 39: Summary
slides_data.append(("content", "Study Summary", [
    "**Design:** Pre-test Post-test Control Group Experimental Design (True Experimental)",
    "**Setting:** Man Nasha Mukti Kendra, Guna, Madhya Pradesh",
    "**Population:** Male opioid-dependent patients, aged 18-50, N = 60 (30/group)",
    "**Experimental:** Brief MBRP (6 sessions / 3 weeks) + Treatment As Usual",
    "**Control:** Psychoeducation (6 sessions / 3 weeks) + Treatment As Usual",
    "**DVs:** Craving (OCDUS) + Impulsivity (BIS-11) + Mindfulness (FFMQ)",
    "**Sampling:** Two-stage: Purposive selection -> Computer-generated random assignment",
    "**Primary Analysis:** ANCOVA controlling for pre-test scores as covariates",
    "**Hypotheses:** Null: No significant difference between groups on OCDUS, BIS-11, FFMQ",
    "",
    "**SIGNIFICANCE:** First brief MBRP trial in Indian opioid-dependent sample addressing critical research gap in evidence-based psychological interventions for de-addiction."
]))

# SLIDE 40: Conclusion
slides_data.append(("content", "Conclusion", [
    "**1.** Substance dependence (particularly opioid) represents a significant public health crisis in India with relapse rates exceeding 70%",
    "**2.** Current treatment at Indian de-addiction centers relies heavily on pharmacotherapy with limited evidence-based psychological interventions",
    "**3.** MBRP offers a theoretically grounded approach targeting craving (urge surfing), impulsivity (mindful pause), and mindfulness enhancement simultaneously",
    "**4.** A brief 6-session B-MBRP protocol is clinically practical, culturally appropriate, and feasible at Man Nasha Mukti Kendra",
    "**5.** If supported, Brief MBRP can be integrated into standard de-addiction protocols across Indian government rehabilitation centers",
    "",
    "**This study represents a critical first step in establishing an evidence base for brief mindfulness-based psychological interventions within the Indian de-addiction treatment infrastructure.**"
]))

# SLIDE 41-44: References
slides_data.append(("content", "References (1/4)", [
    "Baer, R. A., Smith, G. T., Hopkins, J., et al. (2006). Using self-report methods to explore facets of mindfulness. Assessment, 13(1), 27-45.",
    "Bowen, S., Chawla, N., & Marlatt, G. A. (2011). Mindfulness-based relapse prevention for addictive behaviors. Guilford Press.",
    "Bowen, S., & Marlatt, G. A. (2009). Surfing the urge: Brief mindfulness-based intervention. Psychology of Addictive Behaviors, 23(4), 666-671.",
    "Bowen, S., Witkiewitz, K., Clifasefi, S. L., et al. (2014). Relative efficacy of MBRP, standard RP, and TAU. JAMA Psychiatry, 71(5), 547-556.",
    "Brewer, J. A., Mallik, S., Babuscio, T. A., et al. (2011). Mindfulness training for smoking cessation. Drug and Alcohol Dependence, 119(1-2), 72-80.",
    "Chiesa, A., & Serretti, A. (2014). Are MBIs effective for substance use disorders? Substance Use & Misuse, 49(5), 492-512.",
    "Franken, I. H. A., Hendriks, V. M., & van den Brink, W. (2002). Initial validation of two opiate craving questionnaires. Addictive Behaviors, 27(5), 675-685."
]))

slides_data.append(("content", "References (2/4)", [
    "Garland, E. L., Froeliger, B., & Howard, M. O. (2014). Mindfulness targets neurocognitive mechanisms of addiction. Frontiers in Psychiatry, 4, 173.",
    "Garland, E. L., Roberts-Lewis, A., et al. (2016). MORE vs. CBT for co-occurring SUDs. J. Consulting and Clinical Psychology, 84(4), 281-293.",
    "Ghosh, A., Basu, D., & Avasthi, A. (2018). Relapse in opioid dependence: Indian perspective. Indian J Psychiatry, 60(Suppl 4), S469-S476.",
    "Glasner-Edwards, S., et al. (2017). MBRP for stimulant dependent adults: Pilot RCT. Mindfulness, 8(1), 126-135.",
    "Grant, S., Colaiaco, B., et al. (2017). MBRP for SUDs: Meta-analysis. J. Addiction Medicine, 11(5), 386-396.",
    "Humeniuk, R., Ali, R., et al. (2008). Validation of ASSIST. Addiction, 103(6), 1039-1047.",
    "Jain, R., Majumder, P., & Gupta, T. (2013). Indian Journal of Psychiatry, 55(Suppl 1), S86-S92.",
    "Kabat-Zinn, J. (1990). Full catastrophe living. Delacorte Press."
]))

slides_data.append(("content", "References (3/4)", [
    "Karyadi, K. A., VanderVeen, J. D., & Cyders, M. A. (2014). Trait mindfulness and substance use: Meta-analysis. Drug and Alcohol Dependence, 143, 1-10.",
    "Li, W., Howard, M. O., et al. (2017). Mindfulness for substance misuse: Meta-analysis. J. Substance Abuse Treatment, 75, 62-96.",
    "Marlatt, G. A., & Gordon, J. R. (1985). Relapse prevention. Guilford Press.",
    "Mattoo, S. K., Chakrabarti, S., & Anjaiah, M. (2009). Psychosocial factors in relapse. Indian J. Medical Research, 130(6), 702-708.",
    "Ministry of Social Justice & Empowerment. (2019). Magnitude of substance use in India. Government of India.",
    "Moeller, F. G., Barratt, E. S., et al. (2001). Psychiatric aspects of impulsivity. Am. J. Psychiatry, 158(11), 1783-1793.",
    "Murphy, C., & MacKillop, J. (2012). Impulsivity, mindfulness, and alcohol misuse. Psychopharmacology, 219(2), 527-536.",
    "NIDA. (2020). Drugs, brains, and behavior: The science of addiction."
]))

slides_data.append(("content", "References (4/4)", [
    "Patton, J. H., Stanford, M. S., & Barratt, E. S. (1995). Factor structure of BIS. J. Clinical Psychology, 51(6), 768-774.",
    "Robinson, T. E., & Berridge, K. C. (1993). Neural basis of drug craving. Brain Research Reviews, 18(3), 247-291.",
    "Sarkar, S., & Balhara, Y. P. S. (2016). Indian Journal of Psychiatry, 58(3), 290-295.",
    "Serre, F., et al. (2015). Ecological momentary assessment in craving investigation. Addiction, 110(7), 1070-1082.",
    "Stanford, M. S., et al. (2009). Fifty years of the Barratt Impulsiveness Scale. Personality and Individual Differences, 47(5), 385-395.",
    "Tabachnick, B. G., & Fidell, L. S. (2013). Using multivariate statistics (6th ed.). Pearson.",
    "Tiffany, S. T., & Wray, J. M. (2012). Clinical significance of drug craving. Annals NY Acad Sci, 1248(1), 1-17.",
    "WHO ASSIST Working Group. (2002). ASSIST development. Addiction, 97(9), 1183-1194.",
    "Witkiewitz, K., Bowen, S., et al. (2013). MBRP for craving. Addictive Behaviors, 38(2), 1563-1571."
]))

# SLIDE 45: Thank You
slides_data.append(("title", "Thank You",
    "Questions & Discussion",
    ["Tejas Dangodra",
     "2nd Year MPhil Trainee in Clinical Psychology",
     "Man College of Special Education & Psychological Studies",
     "Krantivir Tatya Tope Vishwavidyalaya, Guna, MP",
     "tejasdangodra99@gmail.com | +91 8140171722"]))

print(f"Total slides defined: {len(slides_data)}")



# Now build the PPTX file
def build_pptx(slides, output_path):
    """Build a PPTX file from scratch."""
    
    # Basic PPTX structure files
    content_types_entries = []
    pres_rels_entries = []
    
    # Slide relationship template
    slide_rel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/></Relationships>'
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Copy template structure from original
        orig = zipfile.ZipFile('MBRP_Research_Synopsis_PPT (3) (2).pptx', 'r')
        
        # Copy non-slide files
        skip_prefixes = ['ppt/slides/', 'ppt/charts/', 'ppt/embeddings/']
        for f in orig.namelist():
            skip = False
            for prefix in skip_prefixes:
                if f.startswith(prefix):
                    skip = True
                    break
            if skip:
                continue
            if f == 'ppt/presentation.xml' or f == 'ppt/_rels/presentation.xml.rels' or f == '[Content_Types].xml':
                continue
            zf.writestr(f, orig.read(f))
        
        # Generate slides
        num_slides = len(slides)
        
        for i, slide_info in enumerate(slides):
            slide_num = i + 1
            
            if slide_info[0] == "title":
                _, title, subtitle, meta = slide_info
                xml = make_title_slide_xml(title, subtitle, meta)
            else:
                _, title, bullets = slide_info
                xml = make_slide_xml(title, bullets)
            
            zf.writestr(f'ppt/slides/slide{slide_num}.xml', xml.encode('utf-8'))
            zf.writestr(f'ppt/slides/_rels/slide{slide_num}.xml.rels', slide_rel.encode('utf-8'))
        
        # Build presentation.xml
        slide_id_list = ""
        for i in range(num_slides):
            slide_id_list += f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>'
        
        pres_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>{slide_id_list}</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/>
<p:notesSz cx="6858000" cy="9144000"/>
<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>
</p:presentation>'''
        zf.writestr('ppt/presentation.xml', pres_xml.encode('utf-8'))
        
        # Build presentation.xml.rels
        rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
        for i in range(num_slides):
            rels += f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        rels += f'<Relationship Id="rId{num_slides+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>'
        rels += '</Relationships>'
        zf.writestr('ppt/_rels/presentation.xml.rels', rels.encode('utf-8'))
        
        # Build [Content_Types].xml
        ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        ct += '<Default Extension="xml" ContentType="application/xml"/>'
        ct += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        ct += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        ct += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        ct += '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        for i in range(1, 12):
            ct += f'<Override PartName="/ppt/slideLayouts/slideLayout{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        for i in range(1, num_slides+1):
            ct += f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        ct += '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
        ct += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        ct += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        ct += '</Types>'
        zf.writestr('[Content_Types].xml', ct.encode('utf-8'))
        
        orig.close()
    
    print(f"PPTX created: {output_path} ({num_slides} slides)")

# Build the file
build_pptx(slides_data, 'MBRP_Research_Synopsis_IMPROVED.pptx')
