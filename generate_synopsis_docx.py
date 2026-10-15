#!/usr/bin/env python3
"""
Generates MBRP_Research_Synopsis_Detailed.docx
using raw Office Open XML (no external libraries).
"""
import zipfile
import os
import textwrap

def esc(text):
    """Escape XML special characters."""
    if text is None:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

class DocxBuilder:
    def __init__(self):
        self.body_xml = []
    
    def add_para(self, text, bold=False, italic=False, size=24, font="Times New Roman", 
                 alignment=None, spacing_after=200, first_line_indent=None, line_spacing=None):
        """Add a paragraph. size is in half-points (24=12pt, 28=14pt, 32=16pt)."""
        ppr_parts = []
        if alignment:
            ppr_parts.append(f'<w:jc w:val="{alignment}"/>')
        
        spacing_attrs = f'w:after="{spacing_after}"'
        if line_spacing:
            spacing_attrs += f' w:line="{line_spacing}" w:lineRule="auto"'
        ppr_parts.append(f'<w:spacing {spacing_attrs}/>')
        
        if first_line_indent:
            ppr_parts.append(f'<w:ind w:firstLine="{first_line_indent}"/>')
        
        ppr = f'<w:pPr>{"".join(ppr_parts)}</w:pPr>' if ppr_parts else ''
        
        rpr_parts = []
        if bold:
            rpr_parts.append('<w:b/><w:bCs/>')
        if italic:
            rpr_parts.append('<w:i/><w:iCs/>')
        rpr_parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
        rpr_parts.append(f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>')
        rpr = f'<w:rPr>{"".join(rpr_parts)}</w:rPr>'
        
        self.body_xml.append(f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')
    
    def add_heading(self, text, level=1):
        """Add heading (level 1=16pt bold, 2=14pt bold, 3=13pt bold)."""
        sizes = {1: 32, 2: 28, 3: 26}
        size = sizes.get(level, 24)
        spacing = 300 if level == 1 else 240
        self.add_para(text, bold=True, size=size, spacing_after=spacing, 
                     alignment="left" if level > 1 else "center" if level == 0 else "left")
    
    def add_title(self, text):
        """Add centered title."""
        self.add_para(text, bold=True, size=36, alignment="center", spacing_after=100)
    
    def add_subtitle(self, text):
        """Add centered subtitle."""
        self.add_para(text, bold=False, size=24, alignment="center", spacing_after=100, italic=True)
    
    def add_body(self, text, indent=True):
        """Add body paragraph with first-line indent and 1.5 line spacing."""
        fi = "720" if indent else None
        self.add_para(text, size=24, first_line_indent=fi, spacing_after=200, 
                     alignment="justify", line_spacing="360")
    
    def add_bullet(self, text):
        """Add a bullet point paragraph."""
        ppr = '<w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/></w:pPr>'
        rpr = '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr>'
        bullet_text = f"\u2022  {text}"
        self.body_xml.append(f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{esc(bullet_text)}</w:t></w:r></w:p>')
    
    def add_empty_line(self):
        """Add an empty paragraph for spacing."""
        self.body_xml.append('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>')
    
    def add_page_break(self):
        """Add a page break."""
        self.body_xml.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    
    def build(self, output_path):
        """Build the .docx file."""
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
        
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
        
        doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
        
        styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
        <w:spacing w:after="200" w:line="360" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>'''
        
        body_content = "\n".join(self.body_xml)
        document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body_content}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>'''
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', content_types)
            zf.writestr('_rels/.rels', rels)
            zf.writestr('word/_rels/document.xml.rels', doc_rels)
            zf.writestr('word/styles.xml', styles)
            zf.writestr('word/document.xml', document)
        
        print(f"Created: {output_path} ({os.path.getsize(output_path)} bytes)")


# ═══════════════════════════════════════════════════════════════
# CONTENT GENERATION
# ═══════════════════════════════════════════════════════════════

doc = DocxBuilder()


# ═══════════════ TITLE PAGE ═══════════════
doc.add_empty_line()
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("MAN College of Special Education & Psychological Studies, Guna (MP)", bold=True, size=24, alignment="center", spacing_after=100)
doc.add_para("Department of Clinical Psychology", size=22, alignment="center", spacing_after=100)
doc.add_para("Krantivir Tatya Tope Vishwavidyalaya, Guna, Madhya Pradesh", size=22, alignment="center", spacing_after=100)
doc.add_para("Approved by Rehabilitation Council of India (RCI)", size=20, alignment="center", spacing_after=400, italic=True)
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("RESEARCH SYNOPSIS", bold=True, size=36, alignment="center", spacing_after=300)
doc.add_empty_line()
doc.add_para("Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients", bold=True, size=28, alignment="center", spacing_after=200)
doc.add_empty_line()
doc.add_para("A Pre-Test Post-Test Control Group Experimental Design", italic=True, size=24, alignment="center", spacing_after=400)
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("Researcher: Tejas Dangodra", bold=True, size=24, alignment="center", spacing_after=100)
doc.add_para("Guide: Dr. Himanshi Singh", bold=True, size=24, alignment="center", spacing_after=100)
doc.add_empty_line()
doc.add_para("MPhil Clinical Psychology (2025-2027)", size=22, alignment="center", spacing_after=100)
doc.add_para("Session: 2025-2027", size=22, alignment="center", spacing_after=100)

doc.add_page_break()


# ═══════════════ TABLE OF CONTENTS ═══════════════
doc.add_heading("TABLE OF CONTENTS", level=1)
doc.add_empty_line()
toc_items = [
    "1. Introduction: Substance Dependence in India",
    "2. Introduction: What is MBRP?",
    "3. Variable 1: Craving",
    "4. Variable 2: Impulsivity",
    "5. Variable 3: Mindfulness",
    "6. Review of Literature",
    "7. Research Gap",
    "8. Aim of the Study",
    "9. Objectives",
    "10. Hypotheses",
    "11. Operational Definitions of Key Terms",
    "12. Research Design",
    "13. Sample and Sampling Strategy",
    "14. Sample Size Estimation",
    "15. Inclusion and Exclusion Criteria",
    "16. Variables of the Study",
    "17. Assessment Tools",
    "18. Data Collection Procedure",
    "19. B-MBRP Techniques Used in the Study",
    "20. B-MBRP Session-by-Session Protocol with Hinglish Scripts",
    "21. How Each B-MBRP Session is Conducted",
    "22. Control Group: Psychoeducation Protocol",
    "23. Data Analysis Plan",
    "24. Ethical Considerations",
    "25. Expected Results",
    "26. Clinical Implications",
    "27. Limitations",
    "28. Future Directions",
    "29. Study Summary",
    "30. Conclusion",
    "31. References",
    "Appendix A: Informed Consent Form (Editable)",
    "Appendix B: Assessment Tools (Editable)",
]
for item in toc_items:
    doc.add_para(item, size=24, spacing_after=80)

doc.add_page_break()


# ═══════════════ SECTION 1: INTRODUCTION ═══════════════
doc.add_heading("1. INTRODUCTION: SUBSTANCE DEPENDENCE IN INDIA", level=1)
doc.add_empty_line()

doc.add_heading("1.1 The Indian Context and Magnitude of the Problem", level=2)
doc.add_body("Substance use disorders have emerged as one of the most pressing public health challenges facing the Indian healthcare system today. The Magnitude of Substance Use in India (MAGNITUDE) study, conducted by the National Drug Dependence Treatment Centre (NDDTC) under the Ministry of Social Justice and Empowerment in 2019, reported that approximately 3.1 crore individuals across the country are affected by substance use disorders of varying severity. This figure, already staggering in its magnitude, is widely acknowledged to be an underestimate given the pervasive stigma associated with substance use in Indian society, which discourages many from seeking help or even disclosing their patterns of use during survey research.")

doc.add_body("Opioid dependence, in particular, constitutes a major component of this burden. States such as Punjab, Rajasthan, and several northeastern states report disproportionately high rates of opioid use, including heroin, pharmaceutical opioids such as tramadol and codeine-based preparations, and raw opium derivatives like doda (opium husk). The World Health Organization has estimated that India accounts for roughly 25 percent of all opioid-related deaths in the South-East Asia region, a figure that speaks to the severity of the crisis. In the central Indian state of Madhya Pradesh, where the present study is situated, opioid use patterns are increasingly documented in both urban and semi-urban populations, with heroin and pharmaceutical opioids emerging as the primary substances of dependence at district-level de-addiction facilities.")

doc.add_body("What makes the Indian scenario particularly concerning is not merely the prevalence of use, but the treatment gap that accompanies it. A substantial proportion of individuals meeting clinical criteria for substance dependence never access formal treatment services. Among those who do seek treatment, the available interventions at district-level facilities remain heavily weighted toward pharmacological management, with limited availability of structured, evidence-based psychological interventions. This gap between what is known to be effective (combined pharmacological and psychosocial treatment) and what is actually delivered in routine practice is the fundamental backdrop against which the present study is conceived.")

doc.add_heading("1.2 The Relapse Crisis", level=2)
doc.add_body("Perhaps no single statistic captures the inadequacy of current treatment approaches more starkly than the relapse rate. The National Institute on Drug Abuse (NIDA, 2020) reports global relapse rates of 40 to 60 percent within the first year following treatment completion. Indian data paints an even more discouraging picture: studies from de-addiction settings in North India have consistently reported relapse rates of 70 to 80 percent among opioid-dependent patients within the first three months of discharge (Mattoo et al., 2009; Ghosh, Basu, & Avasthi, 2018). This means that for every ten patients discharged from a typical Indian de-addiction facility, seven or more will have returned to active substance use within 90 days.")

doc.add_body("The triggers for relapse are well-documented in both the international and Indian literature. Craving, which refers to the intense and often overwhelming desire to use a substance, is the most consistently identified proximal trigger. Negative affective states such as anxiety, depression, boredom, loneliness, and shame represent the most common emotional antecedents. Interpersonal conflict, social pressure from substance-using peers, and environmental cues associated with prior use (places, paraphernalia, routines) complete the picture of a complex, multi-determined relapse process that pharmacotherapy alone cannot adequately address.")

doc.add_heading("1.3 The Treatment Gap: Why Pharmacotherapy Alone is Insufficient", level=2)
doc.add_body("Indian de-addiction centres, including the present study setting, primarily offer pharmacological interventions as the cornerstone of treatment. These typically include medically supervised detoxification, Opioid Substitution Therapy (OST) with buprenorphine or methadone, naltrexone maintenance for relapse prevention, and in some cases, symptomatic management of co-occurring psychiatric symptoms. While these pharmacological approaches effectively manage the biological dimension of dependence, they do not, by themselves, equip patients with the psychological skills necessary to navigate craving, regulate difficult emotions, inhibit impulsive responses to triggers, or develop alternative coping repertoires.")

doc.add_body("Psychosocial interventions, despite robust evidence supporting their efficacy as adjuncts to pharmacotherapy, remain grossly underutilized in Indian de-addiction practice (Sarkar & Balhara, 2016). The barriers are multiple: limited availability of trained clinical psychologists at district-level facilities, absence of validated and culturally adapted intervention protocols in Hindi, lack of institutional mandates for psychological treatment, and the pervasive belief among medical staff that detoxification and medication are sufficient. The result is a treatment model that produces biological sobriety at discharge but leaves the psychological vulnerabilities that drive relapse entirely unaddressed.")

doc.add_heading("1.4 The Rationale for Brief Interventions", level=2)
doc.add_body("Indian rehabilitation and de-addiction settings typically operate with admission windows of four to six weeks. Traditional evidence-based interventions such as the standard 8-week MBRP protocol developed by Bowen, Chawla, and Marlatt (2011), or the 12-session CBT-based relapse prevention programmes, are often impractical within these time constraints. Patients are discharged before such protocols can be completed, and outpatient follow-up attendance in Indian settings is notoriously poor, particularly among substance-dependent populations who may live far from the treatment facility, lack transportation, or face financial constraints.")

doc.add_body("This practical reality necessitates the development and evaluation of brief intervention protocols that can be completed within the available admission window while retaining therapeutic potency. A 6-session adaptation delivered twice weekly over three weeks represents the optimal balance between comprehensiveness and feasibility. Such a protocol can be initiated after the acute detoxification phase is complete (typically by day seven of admission) and concluded before the patient's anticipated discharge, ensuring that every enrolled participant has the opportunity to complete the full intervention package.")

doc.add_page_break()


# ═══════════════ SECTION 2: WHAT IS MBRP ═══════════════
doc.add_heading("2. INTRODUCTION: WHAT IS MBRP?", level=1)
doc.add_empty_line()

doc.add_heading("2.1 Development and Background", level=2)
doc.add_body("Mindfulness-Based Relapse Prevention (MBRP) was developed by Sarah Bowen, Neha Chawla, and G. Alan Marlatt at the Addictive Behaviors Research Center, University of Washington, with the full clinical manual published by Guilford Press in 2011. The intervention represents a thoughtful integration of mindfulness meditation practices drawn from Mindfulness-Based Stress Reduction (MBSR; Kabat-Zinn, 1990) and Mindfulness-Based Cognitive Therapy (MBCT; Segal, Williams, & Teasdale, 2002) with the cognitive-behavioural relapse prevention framework originally formulated by Marlatt and Gordon in 1985. It was specifically designed for individuals with substance use disorders, addressing the unique psychological processes that maintain addictive behaviour and precipitate relapse.")

doc.add_body("The development of MBRP was motivated by a fundamental observation: traditional relapse prevention approaches, while effective in identifying high-risk situations and building coping skills, relied primarily on cognitive strategies (thought challenging, coping rehearsal, lifestyle modification) that require a level of executive functioning often compromised in early recovery from substance dependence. Mindfulness offers a complementary and in some ways more fundamental approach. Rather than asking patients to think their way out of craving or argue with permission-giving thoughts, it teaches them to observe these mental events with a quality of detached awareness that naturally reduces their behavioural pull.")

doc.add_heading("2.2 The Core Mechanism of MBRP", level=2)
doc.add_body("The central therapeutic mechanism of MBRP can be stated simply: it teaches patients to observe craving and emotional distress as transient mental events without automatically acting on them. This sounds straightforward, but it represents a radical departure from how most substance-dependent individuals have learned to relate to their internal experience. For the typical patient entering de-addiction treatment, craving has always been experienced as a command that must be obeyed, negative emotions as states that must be immediately escaped, and the impulse to use as a reflex that happens faster than conscious choice can intervene.")

doc.add_body("MBRP disrupts this automatic cycle by training what might be called a 'mindful pause' between stimulus and response. The key technique that exemplifies this mechanism is Urge Surfing, a practice developed by Alan Marlatt himself, in which patients learn to observe craving as one might observe an ocean wave, noticing it rise in intensity, reach a peak, and then naturally subside without any action on their part. Through repeated practice, patients discover experientially that craving is not, as they had believed, a force that will continue escalating until they use. It is a time-limited phenomenon that typically peaks within 15 to 20 minutes and recedes on its own if not reinforced by substance use.")

doc.add_heading("2.3 Neurobiological Basis", level=2)
doc.add_body("The neurobiological rationale for applying mindfulness to addiction is compelling. Chronic substance use produces well-documented changes in brain circuits governing executive control, emotional regulation, and reward processing. Specifically, the prefrontal cortex (PFC), which mediates impulse inhibition, decision-making, and top-down regulation of subcortical impulses, shows reduced activation and structural thinning in substance-dependent individuals. Simultaneously, the amygdala and ventral striatum, which drive emotional reactivity and reward-seeking behaviour respectively, show heightened activation in response to drug cues.")

doc.add_body("Mindfulness practice has been shown to strengthen precisely those circuits that substance use weakens. Neuroimaging research demonstrates that regular meditation practice is associated with increased prefrontal cortex thickness, enhanced PFC-limbic connectivity (allowing better top-down regulation of emotional and craving responses), reduced amygdala reactivity to stress and substance-related cues, and improved functioning of the anterior cingulate cortex, which mediates conflict monitoring and response inhibition. In essence, mindfulness practice provides a form of neural rehabilitation that directly counteracts the neurobiological damage produced by chronic substance use.")

doc.add_heading("2.4 The Brief B-MBRP Adaptation", level=2)
doc.add_body("The Brief MBRP (B-MBRP) protocol employed in the present study represents a carefully structured adaptation of the original 8-session group protocol for the Indian de-addiction inpatient setting. It consists of 6 structured sessions delivered twice weekly over 3 weeks, in group format with 6 to 8 patients per group, with each session lasting approximately 45 minutes. The sessions are designed to be delivered in Hindi, with all meditation scripts, psychoeducational materials, and home practice instructions prepared in a language accessible to the target population.")

doc.add_body("The adaptation retains all core MBRP components, including body scan meditation, breath awareness, urge surfing, the SOBER breathing space, trigger mapping, cognitive decentering, and relapse prevention planning, but condenses them into a format feasible within the typical Indian IPD admission window of 4 to 6 weeks. The group format offers additional therapeutic benefits including normalisation of experiences, peer support, vicarious learning, and cost-effectiveness. The twice-weekly frequency allows adequate time for home practice between sessions while maintaining therapeutic momentum.")

doc.add_body("The three primary targets of B-MBRP align with the three dependent variables of the present study: craving (targeted primarily through urge surfing and SOBER breathing space), impulsivity (targeted through the mindful pause and response inhibition training inherent in all mindfulness practices), and mindfulness itself (built through structured formal meditation practice across all six sessions).")

doc.add_page_break()


# ═══════════════ SECTION 3: VARIABLE 1 - CRAVING ═══════════════
doc.add_heading("3. VARIABLE 1: CRAVING", level=1)
doc.add_empty_line()

doc.add_heading("3.1 Definition of Craving", level=2)
doc.add_body("Craving is defined clinically as an intense, often overwhelming desire or urge to use a substance. Both the DSM-5 (American Psychiatric Association, 2013) and the ICD-11 (World Health Organization, 2019) include craving as a core diagnostic feature of substance use disorders, reflecting the consensus that craving is not merely a symptom but a central driving force in the maintenance of addictive behaviour. Theoretically, craving has been conceptualised in multiple ways across different models of addiction. Franken and colleagues (2002), whose conceptualisation informs the measurement tool used in the present study (OCDUS), define craving as comprising two interrelated dimensions: obsessive thoughts about drug use (the cognitive component) and compulsive urges to use (the behavioural-motivational component). This dual-process conceptualisation captures both the involuntary, intrusive nature of craving thoughts and the sense of compelling urgency to act on them.")

doc.add_heading("3.2 Relevance of Craving in Substance Dependence and Relapse", level=2)
doc.add_body("The clinical significance of craving in substance dependence cannot be overstated. Craving intensity has been consistently identified as the single strongest proximal predictor of relapse across substances, treatment settings, and measurement approaches (Tiffany & Wray, 2012). Patients who report higher craving at discharge are significantly more likely to relapse within the first months of community re-entry. Craving intensity also predicts treatment dropout, with patients experiencing severe craving more likely to leave treatment prematurely, and predicts the progression from a single lapse episode to a full-blown relapse.")

doc.add_body("What makes craving particularly treacherous from a treatment perspective is that traditional suppression-based approaches, in which patients are instructed to 'fight' or 'resist' their craving through willpower, paradoxically tend to increase craving intensity through rebound effects. This is analogous to the well-documented 'white bear' phenomenon in cognitive psychology: the instruction to not think about something makes that very thing more salient and intrusive. MBRP offers a fundamentally different approach, teaching patients to observe craving without either acting on it or fighting it, allowing it to follow its natural wave pattern without reinforcement.")

doc.add_heading("3.3 Neuropsychological Basis of Craving", level=2)
doc.add_body("At the neurobiological level, craving involves activation of the mesolimbic dopamine pathway, with particular involvement of the ventral striatum (nucleus accumbens) and the orbitofrontal cortex (OFC). When a substance-dependent individual encounters drug-related cues, whether external (places, people, paraphernalia) or internal (negative emotions, physical discomfort), these brain regions show hyperactivation, generating the subjective experience of intense wanting. This process is explained by Robinson and Berridge's (1993) incentive sensitisation theory, which proposes that repeated substance use sensitises the wanting system, making it progressively more reactive to drug cues even as the liking or pleasure derived from actual use diminishes.")

doc.add_body("Critically, during episodes of craving, the prefrontal cortex shows relative hypoactivation, meaning that the very brain region responsible for inhibiting impulsive responses and making deliberate choices is least available precisely when it is most needed. This creates the experience that many patients describe: knowing intellectually that using is harmful, but feeling unable to resist the pull. Mindfulness practice, by strengthening prefrontal regulatory capacity and reducing limbic reactivity, addresses this neurobiological vulnerability directly.")

doc.add_heading("3.4 MBRP Techniques Specifically Targeting Craving", level=2)
doc.add_body("Several specific MBRP techniques are designed to target craving:")
doc.add_bullet("Urge Surfing: The patient learns to observe craving as a transient wave-like phenomenon, noticing where it is located in the body, describing its quality and intensity, and watching it change moment-to-moment without acting on it. Through repeated practice, patients discover that craving peaks within 15-20 minutes and subsides naturally.")
doc.add_bullet("SOBER Breathing Space: A structured emergency technique (Stop-Observe-Breathe-Expand-Respond) that can be deployed in real-time when craving is triggered, creating a 2-3 minute pause between the craving impulse and any behavioural response.")
doc.add_bullet("Cognitive Decentering: Patients learn to rephrase their craving experience in observational language ('I am having a craving' rather than 'I need to use'), creating psychological distance between the self and the mental event.")
doc.add_bullet("Mindfulness Exposure: Non-reactive awareness of craving sensations during meditation provides a form of exposure therapy, gradually reducing the intensity of the craving response through habituation and reconditioning.")

doc.add_page_break()


# ═══════════════ SECTION 4: VARIABLE 2 - IMPULSIVITY ═══════════════
doc.add_heading("4. VARIABLE 2: IMPULSIVITY", level=1)
doc.add_empty_line()

doc.add_heading("4.1 Definition of Impulsivity", level=2)
doc.add_body("Impulsivity is defined clinically as a predisposition toward rapid, unplanned actions without adequate consideration of their consequences (Moeller et al., 2001). Unlike a momentary lapse in judgment, impulsivity represents a stable personality-level trait that predisposes individuals toward acting before thinking across a range of situations. The construct is understood to be multidimensional rather than unitary. Patton, Stanford, and Barratt (1995), whose theoretical framework informs the primary measurement instrument of the present study (BIS-11), identify three distinct but related dimensions of impulsivity: Motor Impulsivity (acting without thinking, making rapid motor responses), Attentional Impulsivity (difficulty sustaining attention, racing thoughts, making quick cognitive decisions), and Non-Planning Impulsivity (lack of future orientation, failure to consider consequences, living for the present moment without forethought).")

doc.add_heading("4.2 Relevance of Impulsivity in Substance Dependence", level=2)
doc.add_body("The relationship between impulsivity and substance dependence is bidirectional and self-perpetuating. Higher trait impulsivity is both a pre-existing vulnerability factor that increases the risk of developing substance dependence, and a consequence of chronic substance use that worsens over time as prefrontal regulatory circuits deteriorate. In the context of relapse specifically, impulsivity operates as a critical mediating variable between craving and actual substance use. When craving is experienced, it is the impulsive individual who acts on that craving immediately, before coping strategies can be deployed or the natural subsidence of craving can occur.")

doc.add_body("Research has consistently demonstrated that higher impulsivity scores predict relapse across substances, mediate the craving-to-use behavioural pathway, are associated with treatment non-adherence and premature dropout, and predict the escalation of use frequency and quantity over time. In the motor impulsivity domain specifically, substance-dependent individuals show pronounced deficits on laboratory measures of response inhibition such as Go/No-Go paradigms and Stop-Signal tasks, reflecting a fundamental difficulty in withholding prepotent responses, precisely the capacity needed to refrain from use when triggered.")

doc.add_heading("4.3 Neuropsychological Basis of Impulsivity in Substance Use", level=2)
doc.add_body("The neuropsychological substrate of impulsivity in substance dependence centres on prefrontal cortex dysfunction. The dorsolateral prefrontal cortex (DLPFC), which mediates working memory, planning, and deliberate decision-making, shows consistent hypoactivation in substance-dependent populations. The ventromedial prefrontal cortex, which integrates emotional and value-based information into decision-making, shows altered functioning that manifests as impaired delay discounting, meaning that substance users disproportionately prefer smaller immediate rewards over larger delayed ones.")

doc.add_body("These prefrontal deficits interact with hyperactive subcortical reward systems to create a neural architecture that strongly favours impulsive responding: strong bottom-up signals from the limbic system (craving, emotional urgency) meet weakened top-down inhibitory control from the prefrontal cortex, resulting in the characteristic pattern of knowing that use is harmful but being unable to stop in the moment. This is the precise neural mechanism that mindfulness-based interventions are designed to address.")

doc.add_heading("4.4 MBRP Techniques Specifically Targeting Impulsivity", level=2)
doc.add_body("MBRP targets impulsivity through several mechanisms:")
doc.add_bullet("Response Inhibition Training: Every mindfulness practice inherently involves noticing impulses (to move, to scratch, to open the eyes, to stop meditating) and choosing not to act on them. This builds the general capacity for impulse inhibition that transfers to substance-related situations.")
doc.add_bullet("The STOP Technique: Stop, Take a breath, Observe, Proceed with awareness. A brief metacognitive intervention that inserts a conscious pause into the stimulus-response chain.")
doc.add_bullet("Awareness of Impulse-Action Sequences: Patients learn to notice the micro-sequence of trigger, impulse, action in daily life, developing the metacognitive capacity to observe the impulse before it becomes an action.")
doc.add_bullet("Mindful Decision-Making: Creating space between stimulus and response through deliberate attentional anchoring (breath awareness), allowing the prefrontal cortex time to come online and participate in the decision rather than being bypassed by automatic subcortical responding.")

doc.add_page_break()


# ═══════════════ SECTION 5: VARIABLE 3 - MINDFULNESS ═══════════════
doc.add_heading("5. VARIABLE 3: MINDFULNESS", level=1)
doc.add_empty_line()

doc.add_heading("5.1 Definition of Mindfulness", level=2)
doc.add_body("Mindfulness, in its clinical and psychological usage, refers to the capacity for present-moment awareness characterised by openness, curiosity, and non-judgmental attention to one's ongoing experience. Jon Kabat-Zinn (1990), who is widely credited with introducing mindfulness into Western clinical practice, defined it as 'paying attention, in a particular way, on purpose, in the present moment, and non-judgmentally.' This definition highlights three essential qualities: intentionality (it is a deliberate act of attention, not accidental), present-moment focus (attention is directed to what is happening now, not to memories or future plans), and non-judgment (whatever is observed is met with acceptance rather than evaluation as good or bad).")

doc.add_body("Theoretically, the most comprehensive psychological model of mindfulness is the five-facet model proposed by Baer and colleagues (2006), which was derived from a factor analysis of five pre-existing mindfulness questionnaires and forms the basis of the FFMQ used in the present study. The five facets are: Observing (noticing internal and external experiences such as sensations, cognitions, emotions, sounds, and smells), Describing (the ability to label internal experiences with words), Acting with Awareness (attending fully to one's current activity rather than operating on automatic pilot), Non-Judging of Inner Experience (refraining from evaluating one's thoughts and feelings as good or bad), and Non-Reactivity to Inner Experience (allowing thoughts and feelings to come and go without being swept away by them or compelled to act on them).")

doc.add_heading("5.2 Relevance of Mindfulness in Substance Dependence", level=2)
doc.add_body("Research has consistently demonstrated that individuals with substance use disorders show significantly lower dispositional mindfulness compared to matched controls, across all five facets but particularly on Acting with Awareness and Non-Reactivity subscales (Karyadi et al., 2014). This deficit is not merely an epiphenomenon of substance use; low mindfulness appears to function as both a vulnerability factor that predisposes individuals toward substance use and a maintaining factor that perpetuates the addictive cycle.")

doc.add_body("The protective function of mindfulness against relapse has been documented across multiple studies. Higher mindfulness, particularly on the Non-Reactivity facet, is associated with reduced craving intensity, better emotional regulation, lower impulsive responding to triggers, and reduced probability of relapse following discharge. Critically, improvements in mindfulness have been shown to mediate the treatment outcomes of MBRP interventions, meaning that MBRP reduces craving and substance use specifically through its effect on increasing mindfulness capacity. This establishes mindfulness not merely as a desirable outcome but as the active mechanism through which therapeutic change occurs.")

doc.add_heading("5.3 Neuropsychological Basis of Mindfulness", level=2)
doc.add_body("The neuroscience of mindfulness reveals a pattern of brain changes that directly counteract the neural alterations produced by chronic substance use. The anterior cingulate cortex (ACC), which mediates self-regulation, error detection, and conflict monitoring, shows enhanced activation during and after mindfulness practice. The insula, which mediates interoceptive awareness (the ability to sense internal bodily states), shows increased connectivity, supporting the enhanced body awareness that allows practitioners to notice craving as a physical sensation rather than being overwhelmed by it.")

doc.add_body("The prefrontal cortex shows increased thickness and activation following mindfulness training, directly strengthening the executive control circuits weakened by substance use. PFC-amygdala connectivity improves, allowing better top-down regulation of emotional responses. The Default Mode Network (DMN), which is associated with mind-wandering, rumination, and self-referential processing, shows altered patterns of activation, with experienced meditators showing reduced DMN activity during rest, which correlates with reduced rumination and reduced automatic engagement with craving-related thoughts.")

doc.add_heading("5.4 MBRP Techniques for Building Mindfulness", level=2)
doc.add_body("Each of the core MBRP techniques contributes to building one or more facets of mindfulness:")
doc.add_bullet("Sitting Meditation (Breath Awareness): Trains focused attention and acting with awareness. The patient directs sustained attention to the physical sensations of breathing, notices when the mind wanders, and gently returns attention to the breath without self-criticism.")
doc.add_bullet("Body Scan: Develops the Observing facet through systematic, non-judgmental attention to bodily sensations from head to toe. Builds interoceptive awareness, which is the foundation for recognising craving as a physical sensation.")
doc.add_bullet("Mindful Movement: Gentle yoga or mindful walking with full present-moment attention to the body in motion. Builds acting with awareness in a physical modality that is accessible to patients who find sitting meditation difficult.")
doc.add_bullet("Non-Judgmental Awareness Practice: Labelling experiences without evaluation (noting 'tension' rather than 'this terrible feeling'). Builds the Non-Judging facet that reduces emotional reactivity to internal states.")

doc.add_page_break()


# ═══════════════ SECTION 6: REVIEW OF LITERATURE ═══════════════
doc.add_heading("6. REVIEW OF LITERATURE", level=1)
doc.add_empty_line()

doc.add_heading("6.1 MBRP and Relapse Prevention", level=2)
doc.add_body("Bowen and colleagues (2014), in what remains the most methodologically rigorous evaluation of MBRP to date, conducted a three-arm randomised controlled trial with 286 participants comparing MBRP, standard cognitive-behavioural Relapse Prevention (RP), and Treatment As Usual (TAU). Published in JAMA Psychiatry, this landmark study demonstrated that at 12-month follow-up, MBRP participants reported significantly fewer days of substance use and significantly decreased heavy drinking episodes compared to both standard RP and TAU. What was particularly notable was that while standard RP showed initial superiority over TAU that diminished over time, MBRP maintained and even strengthened its advantage at the 12-month time point. The authors attributed this to the cultivation of mindfulness skills as a durable protective mechanism that continues to develop with ongoing practice, unlike specific coping skills that may decay without reinforcement.")

doc.add_body("Bowen and Marlatt (2009), in an earlier study published in Psychology of Addictive Behaviors, demonstrated the efficacy of a brief urge surfing meditation with incarcerated substance users. This study is particularly relevant to the present research because it established that even brief mindfulness exposure, far shorter than the standard 8-week protocol, can significantly reduce craving intensity and frequency compared to controls. The findings validated that the core MBRP mechanism of disrupting the automaticity of craving responses can be activated even within condensed intervention formats, providing empirical support for the 6-session B-MBRP protocol proposed in the present study.")

doc.add_heading("6.2 Craving and Mindfulness Mechanisms", level=2)
doc.add_body("Garland and colleagues (2014), publishing in Frontiers in Psychiatry, provided critical mechanistic evidence for how mindfulness reduces opioid craving specifically. Their research on Mindfulness-Oriented Recovery Enhancement (MORE) demonstrated that the intervention reduces opioid craving through three complementary mechanisms: attentional reorientation away from drug cues toward natural reward stimuli, positive reappraisal of neutral or mildly pleasant stimuli that had lost their salience through hedonic habituation, and enhanced capacity for savouring healthy pleasures that provides an alternative source of reward. Neuroimaging data from the same research group showed that mindfulness practice modulates activation patterns in prefrontal and limbic circuits during craving exposure, providing biological evidence for the subjective reports of reduced craving.")

doc.add_body("Witkiewitz and colleagues (2013), in a secondary analysis of RCT data published in Addictive Behaviors, examined the mediational pathway through which MBRP reduces substance use. Over a 4-month follow-up period, MBRP participants showed not only lower craving scores but also a significantly attenuated affect-craving pathway. This means that the link between negative emotions and automatic craving responses, which in untreated individuals is strong and automatic, was substantially weakened in MBRP participants. Mindfulness appears to create a buffer between the experience of negative affect and the triggering of craving, precisely by cultivating non-reactive awareness. This finding is of particular relevance because it identifies the specific mechanism through which MBRP prevents relapse: not by eliminating negative emotions (which are inevitable in recovery), but by preventing those emotions from automatically activating craving.")

doc.add_heading("6.3 Impulsivity and Mindfulness", level=2)
doc.add_body("Garland and colleagues (2016), in a study published in the Journal of Consulting and Clinical Psychology, conducted a randomised controlled trial comparing Mindfulness-Oriented Recovery Enhancement (MORE) with CBT in substance-dependent adults. The findings demonstrated that mindfulness intervention produced significant reductions on the BIS-11, particularly in the motor and attentional impulsivity subscales. The mechanism proposed was that sustained attention training inherent in mindfulness practice strengthens prefrontal cortical inhibitory control, essentially rehabilitating the neural circuits responsible for impulse inhibition that chronic substance use has damaged.")

doc.add_body("Murphy and MacKillop (2012), in a cross-sectional study of 340 participants published in Psychopharmacology, examined the relationship between trait mindfulness and impulsive decision-making as measured by delay discounting tasks. They found that trait mindfulness was inversely associated with impulsive decision-making patterns, and importantly, that mindfulness moderated the relationship between impulsivity and alcohol-related consequences. In other words, among individuals with high impulsivity, those with higher mindfulness showed weaker associations between their impulsive tendencies and actual harmful substance use. This suggests that mindfulness functions as a cognitive resource that enables impulsive individuals to override their automatic behavioural tendencies through enhanced metacognitive awareness and response flexibility.")

doc.add_heading("6.4 Meta-Analyses and Brief Intervention Models", level=2)
doc.add_body("Li and colleagues (2017) conducted a comprehensive meta-analysis of 42 randomised controlled trials examining mindfulness-based interventions for substance misuse, published in the Journal of Substance Abuse Treatment. The overall effect sizes were encouraging: substance misuse reduction (d = 0.33), craving reduction (d = 0.68), and stress reduction (d = 0.44). Critically for the present study, their analysis of intervention duration revealed that brief interventions consisting of 4 to 8 sessions showed comparable efficacy to longer protocols when appropriately structured with clear session-by-session learning objectives, home practice assignments, and progressive skill building. This meta-analytic evidence directly supports the feasibility and expected efficacy of the 6-session B-MBRP protocol proposed in the present research.")

doc.add_body("Glasner-Edwards and colleagues (2017), publishing in Mindfulness, conducted a pilot RCT of an abbreviated 6-session MBRP protocol specifically designed for stimulant-dependent adults. The intervention was found to be both feasible (high attendance rates, positive participant feedback) and effective in reducing substance use frequency and craving intensity compared to a health education control condition. This study provides the most direct empirical precedent for the present research, demonstrating that a 6-session MBRP format retains therapeutic potency and is acceptable to substance-dependent populations.")

doc.add_heading("6.5 Indian Context", level=2)
doc.add_body("Ghosh, Basu, and Avasthi (2018), in a review published in the Indian Journal of Psychiatry, documented that relapse rates exceed 70 percent among opioid-dependent patients in North Indian de-addiction centres within three months of discharge. The primary determinants identified were craving, peer influence, negative affect, and the absence of structured psychological aftercare following discharge. The authors noted that Indian facilities rely predominantly on pharmacological approaches, with minimal integration of structured, evidence-based psychological interventions.")

doc.add_body("Sarkar and Balhara (2016) highlighted the systematic underutilisation of structured psychological interventions in Indian de-addiction settings, identifying barriers including limited availability of trained clinical psychologists, absence of validated Hindi-language protocols, and institutional cultures that prioritise pharmacological management over psychological treatment. Jain and colleagues (2013) provided preliminary evidence from a mindfulness-based intervention with alcohol-dependent patients in an Indian setting, showing initial reductions in craving, but the study lacked a control group and used a small sample, limiting the conclusions that could be drawn.")

doc.add_page_break()


# ═══════════════ SECTION 7: RESEARCH GAP ═══════════════
doc.add_heading("7. RESEARCH GAP", level=1)
doc.add_empty_line()
doc.add_body("A careful review of the existing literature, both international and Indian, reveals several specific and significant gaps that the present study is designed to address:")

doc.add_body("First, there is limited evidence for brief MBRP adaptations in low- and middle-income country (LMIC) settings. The vast majority of MBRP research has evaluated the standard 8-week protocol in Western outpatient settings with relatively well-resourced populations. Condensed protocols suitable for the resource-constrained, time-limited inpatient environments characteristic of Indian public sector de-addiction facilities remain insufficiently examined in the published literature (Li et al., 2017).")

doc.add_body("Second, existing studies have typically examined craving, impulsivity, or mindfulness in isolation, measuring only one or at most two of these variables as outcomes. The simultaneous assessment of all three mechanistically linked variables within a single intervention framework has not been adequately investigated in Indian clinical populations. Given the theoretical interrelationships between these constructs, wherein mindfulness is hypothesised to reduce craving through non-reactive awareness and to reduce impulsivity through strengthened inhibitory control, a multi-domain assessment approach is essential for understanding the full spectrum of MBRP effects and their potential interactions.")

doc.add_body("Third, despite India's substantial opioid dependence burden, empirical validation of MBRP-based protocols in Indian de-addiction infrastructure remains virtually absent from the published literature. The handful of Indian studies that have examined mindfulness-based approaches have used MBSR rather than MBRP, have focused on alcohol rather than opioid dependence, and have been conducted at tertiary care institutions whose resources and patient populations differ substantially from district-level facilities where the majority of Indian patients actually receive treatment.")

doc.add_body("Fourth, current treatment models at district-level Indian centres remain predominantly pharmacological, with limited evidence-based psychosocial adjuncts. There is an urgent need for research that demonstrates the feasibility and efficacy of brief psychological interventions that can be realistically integrated into existing treatment infrastructure without requiring resources (specialised training, extended treatment duration, sophisticated technology) that are unavailable at this level of the healthcare system (Sarkar & Balhara, 2016; Ghosh et al., 2018).")

doc.add_body("Fifth, most existing mindfulness studies in Indian addiction settings lack active control conditions, making it difficult to distinguish mindfulness-specific effects from non-specific therapeutic factors such as therapist attention, group support, expectancy effects, and the passage of time. Without an attention-matched control condition, any observed improvements cannot be confidently attributed to the mindfulness component specifically.")

doc.add_body("The present study directly addresses all five of these gaps. It represents the first brief MBRP trial (6 sessions over 3 weeks) conducted with an Indian opioid-dependent sample at a de-addiction centre, employing an active attention-matched control group (psychoeducation) and simultaneous three-variable assessment (craving, impulsivity, and mindfulness).", indent=False)

doc.add_page_break()

# ═══════════════ SECTION 8: AIM ═══════════════
doc.add_heading("8. AIM OF THE STUDY", level=1)
doc.add_empty_line()
doc.add_body("The aim of the present study is to evaluate the efficacy of a Brief Mindfulness-Based Relapse Prevention (B-MBRP) intervention consisting of 6 sessions delivered over 3 weeks in reducing craving and impulsivity, and enhancing mindfulness, among male substance-dependent patients undergoing treatment at a de-addiction centre. The study compares an experimental group receiving Brief MBRP plus Treatment As Usual (TAU) with an active control group receiving matched Psychoeducation plus TAU, employing a pre-test post-test control group experimental design with random assignment.")

doc.add_body("The study involves a total of 60 participants (30 per group), uses validated psychometric instruments (OCDUS for craving, BIS-11 for impulsivity, FFMQ for mindfulness), and employs ANCOVA as the primary statistical analysis to control for pre-test differences and provide the most precise estimate of the intervention effect.")

doc.add_page_break()

# ═══════════════ SECTION 9: OBJECTIVES ═══════════════
doc.add_heading("9. OBJECTIVES", level=1)
doc.add_empty_line()
doc.add_body("The following specific objectives guide the present investigation:")

doc.add_body("Objective 1: To assess and compare craving levels (as measured by OCDUS total scores) at pre-test and post-test in the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU).")

doc.add_body("Objective 2: To assess and compare impulsivity levels (as measured by BIS-11 total and subscale scores) at pre-test and post-test in both groups.")

doc.add_body("Objective 3: To assess and compare mindfulness levels (as measured by FFMQ total and facet scores) at pre-test and post-test in both groups.")

doc.add_body("Objective 4: To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing craving scores from pre-test to post-test.")

doc.add_body("Objective 5: To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in reducing impulsivity scores from pre-test to post-test.")

doc.add_body("Objective 6: To determine whether Brief MBRP + TAU is significantly more effective than Psychoeducation + TAU in enhancing mindfulness scores from pre-test to post-test.")

doc.add_page_break()


# ═══════════════ SECTION 10: HYPOTHESES ═══════════════
doc.add_heading("10. HYPOTHESES", level=1)
doc.add_empty_line()
doc.add_body("The present study tests the following null hypotheses at an alpha level of 0.05 (two-tailed):")
doc.add_empty_line()

doc.add_para("H01 (Craving):", bold=True, size=24, spacing_after=100)
doc.add_body("There is no significant difference in craving scores (OCDUS) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.")
doc.add_empty_line()

doc.add_para("H02 (Impulsivity):", bold=True, size=24, spacing_after=100)
doc.add_body("There is no significant difference in impulsivity scores (BIS-11) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.")
doc.add_empty_line()

doc.add_para("H03 (Mindfulness):", bold=True, size=24, spacing_after=100)
doc.add_body("There is no significant difference in mindfulness scores (FFMQ) between the Experimental Group (Brief MBRP + TAU) and the Control Group (Psychoeducation + TAU) from pre-test to post-test.")
doc.add_empty_line()

doc.add_body("The primary statistical analysis for testing these hypotheses is one-way ANCOVA, with post-test scores as the dependent variable, group (Experimental vs. Control) as the independent variable, and pre-test scores as the covariate. This analytical approach controls for any baseline differences between groups and provides the most precise estimate of the intervention effect.")

doc.add_page_break()

# ═══════════════ SECTION 11: OPERATIONAL DEFINITIONS ═══════════════
doc.add_heading("11. OPERATIONAL DEFINITIONS OF KEY TERMS", level=1)
doc.add_empty_line()

doc.add_para("Craving:", bold=True, size=24, spacing_after=100)
doc.add_body("'Craving' in the present study refers to the intensity and frequency of obsessive thoughts about drug use and compulsive urges to use opioids, as measured by the total score on the Obsessive Compulsive Drug Use Scale (OCDUS; Franken et al., 2002). The OCDUS comprises 13 items scored on a 5-point scale (0-4), yielding a total score range of 0 to 52. Higher scores indicate greater craving severity. The scale captures both the cognitive dimension (obsessive preoccupation with drug-related thoughts) and the behavioural-motivational dimension (compulsive urges and difficulty resisting thoughts of use).")
doc.add_empty_line()

doc.add_para("Impulsivity:", bold=True, size=24, spacing_after=100)
doc.add_body("'Impulsivity' in the present study refers to the multidimensional tendency to act without adequate forethought, encompassing motor, attentional, and non-planning components, as measured by the total score and three subscale scores on the Barratt Impulsiveness Scale-11 (BIS-11; Patton, Stanford, & Barratt, 1995). The BIS-11 comprises 30 items scored on a 4-point scale (1-4), yielding a total score range of 30 to 120. Higher scores indicate greater impulsivity. The three subscales measure distinct dimensions: Motor Impulsivity (acting without thinking), Attentional Impulsivity (difficulty concentrating and cognitive instability), and Non-Planning Impulsivity (lack of future orientation and failure to plan ahead).")
doc.add_empty_line()

doc.add_para("Mindfulness:", bold=True, size=24, spacing_after=100)
doc.add_body("'Mindfulness' in the present study refers to the dispositional capacity for present-moment awareness characterised by non-judgment and non-reactivity, as measured by the total score on the Five Facet Mindfulness Questionnaire (FFMQ; Baer et al., 2006) across five facets: Observing, Describing, Acting with Awareness, Non-Judging, and Non-Reactivity. The FFMQ comprises 39 items scored on a 5-point scale (1-5), yielding a total score range of 39 to 195. Higher scores indicate greater dispositional mindfulness.")
doc.add_empty_line()

doc.add_para("Brief MBRP:", bold=True, size=24, spacing_after=100)
doc.add_body("'Brief MBRP' in the present study refers to a structured 6-session mindfulness-based relapse prevention intervention delivered twice weekly over 3 weeks (45 minutes per session) in group format (6-8 patients per group), adapted from Bowen, Chawla, and Marlatt (2011). The protocol incorporates body scan meditation, breath awareness, urge surfing, SOBER breathing space, trigger mapping, cognitive decentering, loving-kindness meditation, and relapse prevention planning, delivered in Hindi with culturally appropriate examples and metaphors.")
doc.add_empty_line()

doc.add_para("Substance Dependence:", bold=True, size=24, spacing_after=100)
doc.add_body("'Substance Dependence' in the present study refers to a clinical diagnosis of Substance Dependence Syndrome as per ICD-10 (F10-F19) criteria, with primary opioid dependence (heroin or pharmaceutical opioids), confirmed by a qualified psychiatrist at the de-addiction centre. The diagnosis requires the presence of at least three of the following during the previous year: compulsion to use, impaired control, physiological withdrawal, tolerance, progressive neglect of alternatives, and continued use despite harm.")
doc.add_empty_line()

doc.add_para("Treatment As Usual (TAU):", bold=True, size=24, spacing_after=100)
doc.add_body("'TAU' in the present study refers to the standard pharmacological treatment at the de-addiction centre, including medically supervised detoxification, Opioid Substitution Therapy (buprenorphine or methadone), naltrexone maintenance for relapse prevention, routine counselling, and daily ward activities. TAU is continued for all participants in both experimental and control groups throughout the study period.")
doc.add_empty_line()

doc.add_para("Psychoeducation (Active Control):", bold=True, size=24, spacing_after=100)
doc.add_body("'Psychoeducation' in the present study refers to 6 structured informational sessions (45 minutes each, twice weekly, over 3 weeks) covering addiction science, effects of opioids on body and mind, relapse warning signs, health and nutrition in recovery, social consequences of substance use, and recovery motivation and goal-setting. These sessions are matched to the experimental group for time, attention, format (group of 6-8), and therapist contact, but contain no mindfulness component whatsoever.")

doc.add_page_break()


# ═══════════════ SECTION 12: RESEARCH DESIGN ═══════════════
doc.add_heading("12. RESEARCH DESIGN", level=1)
doc.add_empty_line()

doc.add_body("The present study employs a Pre-Test Post-Test Control Group Experimental Design, which is classified as a true experimental design. The design notation is as follows:")
doc.add_empty_line()
doc.add_para("R   O1   X1   O2   -->  Experimental Group (Brief MBRP + TAU)", size=24, spacing_after=80)
doc.add_para("R   O1   X2   O2   -->  Control Group (Psychoeducation + TAU)", size=24, spacing_after=200)
doc.add_empty_line()

doc.add_body("Where R denotes random assignment, O1 represents the pre-test assessment (administration of OCDUS, BIS-11, FFMQ, and ASSIST), X1 represents the experimental intervention (6 sessions of Brief MBRP + TAU), X2 represents the active control condition (6 sessions of Psychoeducation + TAU), and O2 represents the post-test assessment (re-administration of OCDUS, BIS-11, and FFMQ).")

doc.add_body("This design was chosen for several methodological reasons. First, the inclusion of random assignment to groups elevates the study to a true experimental design, which provides the strongest basis for causal inference regarding the intervention's effects. Second, the use of an active control condition (psychoeducation matched for time, attention, format, and therapist contact) allows the study to isolate the specific effects of the mindfulness component from non-specific therapeutic factors. If the experimental group shows significantly greater improvement than the control group, this improvement can be attributed specifically to the mindfulness training rather than to therapist attention, group support, or structured activity alone.")

doc.add_body("Third, the pre-test measurement allows assessment of baseline equivalence between groups (essential for confirming that randomisation was effective) and enables the use of ANCOVA with pre-test scores as covariates, which increases statistical power and provides more precise effect estimates. Fourth, the standardised timing of assessments (pre-test before intervention begins, post-test within one week of intervention completion) ensures that all participants are measured at equivalent time points relative to their intervention exposure.")

doc.add_body("The study is conducted at the De-addiction Centre, Guna, Madhya Pradesh. All assessments are administered by a blinded research assistant who is unaware of participants' group allocation. The intervention is delivered by the researcher (MPhil Clinical Psychologist) following a manualized protocol to ensure treatment fidelity.")

doc.add_page_break()

# ═══════════════ SECTION 13: SAMPLE AND SAMPLING ═══════════════
doc.add_heading("13. SAMPLE AND SAMPLING STRATEGY", level=1)
doc.add_empty_line()

doc.add_heading("13.1 Sampling Procedure", level=2)
doc.add_body("The sampling procedure for the present study follows a two-stage approach. In the first stage, purposive selection is employed to identify eligible participants from among patients admitted to the de-addiction centre. All patients meeting the inclusion criteria during the recruitment window are approached for participation. This consecutive sampling of all eligible patients ensures that the sample is representative of the treatment-seeking population at this facility rather than a selectively chosen subset.")

doc.add_body("In the second stage, random assignment is used to allocate eligible consenting participants to either the Experimental Group (n=30) or the Control Group (n=30). Randomisation is achieved through computer-generated random sequences, with allocation concealment maintained through sealed opaque envelopes that are opened only after the participant has completed the pre-test assessment and confirmed their eligibility and consent. This procedure prevents selection bias in group assignment and ensures that both groups are equivalent on all known and unknown confounding variables at baseline.")

doc.add_heading("13.2 Sample Characteristics", level=2)
doc.add_body("The total sample comprises 60 male participants, with 30 allocated to each group. An additional 10 participants (5 per group) are recruited to account for anticipated attrition of approximately 15 percent, which is common in substance-dependent populations due to discharge against medical advice, medical complications, or withdrawal of consent. The target is a final analysable sample of 30 per group (total N = 60).")

doc.add_body("The sample is restricted to male participants for the following reasons: Indian de-addiction centres admit approximately 90 to 95 percent male patients, reflecting the heavily gendered pattern of substance use in India. The MAGNITUDE study (2019) reported a male-to-female ratio of approximately 10:1 for opioid dependence. Restricting the sample to males provides a homogeneous sample that strengthens internal validity and avoids the confounding effects of gender on all three dependent variables. Studies specifically examining female substance-dependent populations are recommended as an important future direction.")

doc.add_body("Participants are aged 18 to 50 years, currently admitted at the de-addiction centre, and have completed the acute detoxification phase (minimum 7 days post-admission). This ensures that participants are medically stable, not experiencing active withdrawal symptoms, and cognitively capable of engaging with the intervention and assessment procedures.")

doc.add_page_break()

# ═══════════════ SECTION 14: SAMPLE SIZE ═══════════════
doc.add_heading("14. SAMPLE SIZE ESTIMATION", level=1)
doc.add_empty_line()

doc.add_body("The sample size for the present study was determined using the standard formula for comparing two independent group means:")
doc.add_empty_line()
doc.add_para("n = [(Z_alpha/2 + Z_beta)^2 x 2 x sigma^2] / d^2", bold=True, size=24, alignment="center", spacing_after=200)
doc.add_empty_line()

doc.add_body("The following parameters were used: an expected medium effect size of d = 0.50, based on the meta-analytic findings of Li et al. (2017) who reported effect sizes of d = 0.33 to 0.68 for mindfulness-based interventions on substance use outcomes; statistical power of 0.80 (Z_beta = 0.84); and alpha level of 0.05 two-tailed (Z_alpha/2 = 1.96). Substituting these values:")
doc.add_empty_line()
doc.add_para("n = [(1.96 + 0.84)^2 x 2 x 1] / (0.50)^2 = [7.84 x 2] / 0.25 = 62.72", size=24, alignment="center", spacing_after=200)
doc.add_empty_line()

doc.add_body("This yields a requirement of approximately 63 participants total, or approximately 32 per group. The decision to use N = 60 (30 per group) is justified by the fact that ANCOVA, the primary analysis, reduces the required sample size compared to simple t-tests by controlling for pre-test variance, effectively increasing statistical power. Verification using G*Power 3.1 software confirms that N = 30 per group provides adequate power (0.80) to detect a medium effect size (f = 0.25) with ANCOVA (one covariate). To account for expected attrition of approximately 15 percent, a total of 70 participants (35 per group) will be recruited initially.")

doc.add_page_break()


# ═══════════════ SECTION 15: INCLUSION/EXCLUSION ═══════════════
doc.add_heading("15. INCLUSION AND EXCLUSION CRITERIA", level=1)
doc.add_empty_line()

doc.add_heading("15.1 Inclusion Criteria", level=2)
doc.add_bullet("Diagnosis of Substance Dependence as per ICD-10 (F10-F19) or ICD-11 criteria, confirmed by the treating psychiatrist")
doc.add_bullet("Primary substance of dependence: Opioids (heroin, pharmaceutical opioids); polysubstance dependence with primary opioid included")
doc.add_bullet("Male participants aged 18 to 50 years")
doc.add_bullet("Completed detoxification phase with minimum 7 days post-withdrawal management")
doc.add_bullet("Currently admitted at the de-addiction centre")
doc.add_bullet("Minimum education of 5th standard (sufficient literacy to comprehend psychometric tool items when read aloud)")
doc.add_bullet("Willingness to provide written informed consent")
doc.add_bullet("Able to attend all 6 intervention sessions during the 3-week intervention period")

doc.add_heading("15.2 Exclusion Criteria", level=2)
doc.add_bullet("Severe psychiatric comorbidity: Active psychotic disorders, Bipolar I disorder, severe Major Depressive Episode with active suicidality")
doc.add_bullet("Significant cognitive impairment as assessed by Mini Mental State Examination (MMSE score below 24) or intellectual disability")
doc.add_bullet("Active withdrawal symptoms at the time of pre-test assessment (Clinical Opiate Withdrawal Scale score above 12)")
doc.add_bullet("History of traumatic brain injury with loss of consciousness exceeding 30 minutes")
doc.add_bullet("Current participation in another structured psychological intervention study")
doc.add_bullet("Medical instability requiring acute or intensive medical care")
doc.add_bullet("History of prior formal mindfulness or meditation training exceeding one month of regular practice")

doc.add_page_break()

# ═══════════════ SECTION 16: VARIABLES ═══════════════
doc.add_heading("16. VARIABLES OF THE STUDY", level=1)
doc.add_empty_line()

doc.add_heading("16.1 Independent Variable", level=2)
doc.add_body("The independent variable in the present study is the Type of Intervention, which has two levels: Level 1 consists of Brief MBRP plus Treatment As Usual (Experimental Group), and Level 2 consists of Psychoeducation plus Treatment As Usual (Active Control Group). Both interventions are matched for duration (45 minutes per session), frequency (twice weekly), number of sessions (6 sessions), total intervention period (3 weeks), format (group of 6-8 participants), and therapist contact time.")

doc.add_heading("16.2 Dependent Variables", level=2)
doc.add_body("The study has three dependent variables:")
doc.add_bullet("Dependent Variable 1 - Craving: Measured by the OCDUS total score (range 0-52; higher scores indicate greater craving)")
doc.add_bullet("Dependent Variable 2 - Impulsivity: Measured by the BIS-11 total score and three subscale scores (total range 30-120; higher scores indicate greater impulsivity)")
doc.add_bullet("Dependent Variable 3 - Mindfulness: Measured by the FFMQ total score and five facet scores (total range 39-195; higher scores indicate greater mindfulness)")

doc.add_heading("16.3 Controlled Variables", level=2)
doc.add_body("The following variables are controlled in the present study to minimise their confounding influence on the dependent variables: age (restricted to 18-50), education level (minimum 5th standard), duration of substance use (assessed at baseline), severity of dependence (assessed using WHO-ASSIST V3.0 at baseline), TAU components (identical for both groups throughout the study), session duration (equalised at 45 minutes for both groups), number of sessions (6 for both groups), and therapist (same researcher delivers both interventions to control for therapist effects).")

doc.add_page_break()


# ═══════════════ SECTION 17: ASSESSMENT TOOLS ═══════════════
doc.add_heading("17. ASSESSMENT TOOLS", level=1)
doc.add_empty_line()

doc.add_heading("17.1 Obsessive Compulsive Drug Use Scale (OCDUS)", level=2)
doc.add_body("The Obsessive Compulsive Drug Use Scale (OCDUS), developed by Franken and colleagues (2002), is a 13-item self-report instrument designed to measure the obsessive and compulsive dimensions of craving for drugs. The scale was modelled on the Yale-Brown Obsessive Compulsive Scale (Y-BOCS) and captures both the cognitive dimension of craving (intrusive, obsessive thoughts about drug use, difficulty resisting such thoughts, interference with daily functioning) and the behavioural-motivational dimension (compulsive urges to use, perceived loss of control over use-related behaviour, desire to use).")

doc.add_body("Each item is scored on a 5-point scale ranging from 0 to 4, yielding a total score range of 0 to 52. Higher scores indicate greater craving severity. The scale comprises three subscales: Obsessive Thoughts and Interference (items assessing the frequency, duration, and disruptiveness of drug-related thoughts), Desire and Control (items assessing the intensity of desire to use and perceived ability to resist), and Resistance to Thoughts (items assessing efforts and ability to resist drug-related cognitions).")

doc.add_body("Psychometric properties are strong: internal consistency (Cronbach's alpha) ranges from 0.86 to 0.90 across validation studies, test-retest reliability is r = 0.78, and convergent validity with Visual Analogue Scale craving measures is r = 0.55 to 0.67. The OCDUS has demonstrated sensitivity to treatment-related changes and is applicable across substance types including opioids. It is brief (approximately 5 minutes to administer), suitable for pre-post designs, and adaptable for Hindi administration through read-aloud procedures for participants with literacy limitations.")

doc.add_heading("17.2 Barratt Impulsiveness Scale-11 (BIS-11)", level=2)
doc.add_body("The Barratt Impulsiveness Scale-11 (BIS-11), developed by Patton, Stanford, and Barratt (1995), is the most widely used self-report measure of impulsivity in clinical and research settings. It comprises 30 items, each rated on a 4-point scale (1 = rarely/never to 4 = almost always/always), with 11 items reverse-scored. The total score ranges from 30 to 120, with higher scores indicating greater impulsivity.")

doc.add_body("The BIS-11 yields a total impulsivity score and three factor-analytically derived subscale scores: Attentional Impulsivity (8 items measuring difficulty focusing attention, cognitive instability, and racing thoughts), Motor Impulsivity (11 items measuring acting without thinking, restlessness, and making rapid responses), and Non-Planning Impulsivity (11 items measuring lack of future orientation, failure to consider consequences, and preference for immediate over delayed gratification). Of these, Motor Impulsivity is most directly relevant to relapse, as it captures the tendency to act on craving impulses without adequate reflection.")

doc.add_body("Psychometric properties are well-established: internal consistency is 0.79 to 0.83, test-retest reliability is r = 0.83, and the factor structure has been replicated across multiple languages including a Hindi-validated version. The BIS-11 discriminates reliably between substance-dependent individuals and healthy controls, and has been shown to be sensitive to intervention effects in mindfulness-based treatment studies.")

doc.add_heading("17.3 Five Facet Mindfulness Questionnaire (FFMQ)", level=2)
doc.add_body("The Five Facet Mindfulness Questionnaire (FFMQ), developed by Baer and colleagues (2006), is the most comprehensive and widely validated multidimensional self-report measure of mindfulness. It was derived from a factor analysis combining items from five pre-existing mindfulness scales (Mindful Attention Awareness Scale, Freiburg Mindfulness Inventory, Kentucky Inventory of Mindfulness Skills, Cognitive and Affective Mindfulness Scale, and Southampton Mindfulness Questionnaire), yielding a five-factor structure that captures the full breadth of the mindfulness construct.")

doc.add_body("The FFMQ comprises 39 items rated on a 5-point Likert scale (1 = never or very rarely true to 5 = very often or always true), with several items reverse-scored. It yields a total score (range 39-195) and five facet scores: Observing (8 items; noticing internal and external stimuli), Describing (8 items; labelling experiences with words), Acting with Awareness (8 items; attending to present activity vs. autopilot), Non-Judging of Inner Experience (8 items; accepting thoughts and feelings without evaluation), and Non-Reactivity to Inner Experience (7 items; allowing thoughts and feelings to pass without being swept away).")

doc.add_body("The Acting with Awareness and Non-Reactivity facets are of particular relevance to the present study, as they correspond most directly to the capacities that MBRP is designed to build. Internal consistency ranges from 0.75 to 0.91 across facets. The FFMQ has been shown to be sensitive to mindfulness intervention effects and is the primary self-report outcome in all published MBRP trials. A Hindi adaptation is available for use with Indian populations.")

doc.add_heading("17.4 WHO-ASSIST Version 3.0 (Baseline Severity Measure)", level=2)
doc.add_body("The Alcohol, Smoking and Substance Involvement Screening Test Version 3.0 (WHO-ASSIST V3.0), developed by the WHO ASSIST Working Group (2002), is an 8-item questionnaire measuring substance involvement across 10 substance categories (tobacco, alcohol, cannabis, cocaine, amphetamines, inhalants, sedatives, hallucinogens, opioids, and other drugs). It assesses lifetime use, past 3-month use frequency, and associated problems.")

doc.add_body("Scoring yields substance-specific risk levels: Low risk (0-3, no intervention needed), Moderate risk (4-26, brief intervention indicated), and High risk (27 and above, specialist referral indicated). Psychometric properties are strong: test-retest reliability ranges from r = 0.58 to 0.90, internal consistency is 0.77 to 0.94, sensitivity is 0.80, and specificity is 0.71. The ASSIST has been validated across 18 countries and is available in Hindi.")

doc.add_body("In the present study, WHO-ASSIST V3.0 is administered at pre-test only. It serves to establish baseline severity of substance involvement, ensure group equivalence at baseline, and provide a stratification variable for randomisation. It is NOT used as an outcome measure, as severity of dependence does not change meaningfully within a 3-week inpatient admission during which patients are not using substances.")

doc.add_page_break()


# ═══════════════ SECTION 18: DATA COLLECTION ═══════════════
doc.add_heading("18. DATA COLLECTION PROCEDURE", level=1)
doc.add_empty_line()

doc.add_body("The data collection procedure follows a systematic, step-by-step protocol designed to ensure methodological rigour, participant safety, and ethical compliance:")

doc.add_heading("Step 1: Screening (Day 1-3 of Recruitment)", level=2)
doc.add_body("The researcher reviews admission records in coordination with the ward in-charge and treating psychiatrist to identify male patients aged 18 to 50 years with a confirmed diagnosis of opioid dependence (ICD-10) who have completed at least 7 days of detoxification. The Mini Mental State Examination is administered to confirm cognitive adequacy (MMSE score of 24 or above), and the Clinical Opiate Withdrawal Scale (COWS) is checked to confirm absence of active withdrawal symptoms (COWS score of 12 or below). Inclusion and exclusion criteria are systematically applied.")

doc.add_heading("Step 2: Informed Consent (Day 3-4)", level=2)
doc.add_body("Eligible patients are approached individually in a private space within the ward. The study is explained in Hindi, covering the purpose, procedures, duration of participation, voluntary nature, right to withdraw at any time without impact on treatment, and confidentiality protections. Patients are given adequate time to ask questions and consider their decision. Written informed consent is obtained, with the consent form available in both Hindi and English. For patients who cannot write, a thumb impression witnessed by a ward staff member is acceptable.")

doc.add_heading("Step 3: Pre-Test Assessment (Day 4-5)", level=2)
doc.add_body("Pre-test assessment is conducted in a quiet, private room within the de-addiction centre during morning hours (9 AM to 12 noon) to minimise medication effects. The blinded research assistant administers the full battery in standardised order: WHO-ASSIST V3.0 (baseline severity), OCDUS (craving), BIS-11 (impulsivity), and FFMQ (mindfulness). Total administration time is approximately 40 minutes. For participants with literacy difficulties, items are read aloud and responses recorded by the assessor. Practice items are completed first to ensure comprehension.")

doc.add_heading("Step 4: Randomisation (Day 5)", level=2)
doc.add_body("Following completion of pre-test assessment, participants are allocated to either the Experimental Group or the Control Group using computer-generated random sequences contained in sealed opaque envelopes. Randomisation is stratified by ASSIST severity score (moderate vs. high risk) to ensure balanced groups on this important prognostic variable. The allocating researcher is different from the blinded assessor.")

doc.add_heading("Step 5: Intervention Period (Weeks 1-3)", level=2)
doc.add_body("Both groups receive their respective 6-session interventions twice weekly over 3 weeks. Sessions are conducted in groups of 6 to 8 participants, each lasting 45 minutes. The Experimental Group receives Brief MBRP (mindfulness techniques as detailed in the session protocol), while the Control Group receives matched Psychoeducation (informational sessions without any mindfulness component). Both groups continue to receive Treatment As Usual (pharmacotherapy, routine counselling, ward activities) throughout the intervention period.")

doc.add_heading("Step 6: Post-Test Assessment (Within 1 Week Post-Intervention)", level=2)
doc.add_body("Post-test assessment is conducted within one week of completing the final intervention session, under identical conditions to pre-test (same room, same time of day, same blinded assessor, same standardised instructions). Only the three outcome measures are re-administered: OCDUS, BIS-11, and FFMQ. The ASSIST is not re-administered as it measures baseline severity rather than treatment response. Participants are instructed not to reveal their group allocation to the assessor.")

doc.add_heading("Assessment Conditions and Bias Control", level=2)
doc.add_body("All assessments are conducted in a quiet, private room with comfortable seating and adequate lighting, with no other patients present. The same room and conditions are maintained for pre-test and post-test. The research assistant administering assessments is blinded to group allocation. Standardised instructions are read verbatim. Social desirability effects are managed through anonymous participant ID coding (no names on any assessment form). Data is entered into a password-protected electronic database within 48 hours of each assessment, with double data entry for accuracy verification. Paper forms are stored in a locked cabinet accessible only to the research team.")

doc.add_page_break()


# ═══════════════ SECTION 19: B-MBRP TECHNIQUES ═══════════════
doc.add_heading("19. B-MBRP TECHNIQUES USED IN THE STUDY", level=1)
doc.add_empty_line()

doc.add_body("The following core techniques constitute the active therapeutic ingredients of the B-MBRP intervention. Each technique is described in terms of its procedure, therapeutic rationale, and the specific dependent variable it targets:")

doc.add_heading("19.1 Body Scan Meditation", level=2)
doc.add_body("The body scan involves systematic, deliberate attention to physical sensations in each part of the body, progressing from head to toe (or toe to head) without judgment. The patient is guided to simply notice whatever sensations are present in each area, whether pleasant, unpleasant, or neutral, without trying to change them. When the mind wanders to thoughts, plans, or memories, the patient gently redirects attention back to the body part being scanned. The body scan builds the Observing facet of mindfulness and provides the experiential foundation for urge surfing, as it teaches patients to relate to bodily sensations (including craving sensations) with curiosity rather than reactivity. Duration: 10-15 minutes in sessions, building to 20 minutes in home practice.")

doc.add_heading("19.2 Breath Awareness Meditation", level=2)
doc.add_body("Breath awareness involves directing sustained, focused attention to the natural rhythm of breathing, typically noticing the sensations of air entering and leaving the nostrils, or the rise and fall of the abdomen. The breath serves as a present-moment anchor; when the mind wanders (which it inevitably will), the instruction is simply to notice the wandering without judgment and gently return attention to the breath. This practice primarily builds the Acting with Awareness facet of mindfulness, training the capacity to maintain attention on a chosen object rather than being carried away by automatic thought patterns. In the context of substance dependence, the breath becomes a portable, always-available resource that patients can access in any high-risk situation.")

doc.add_heading("19.3 Urge Surfing", level=2)
doc.add_body("Urge surfing, developed by Alan Marlatt, is the signature technique of MBRP and the primary intervention for craving. The metaphor is that craving is like an ocean wave: it rises in intensity, reaches a peak, and then naturally subsides, following a predictable wave pattern that typically peaks within 15 to 20 minutes and diminishes without action. The patient is guided to observe the craving as a physical sensation in the body, noting its location (chest, stomach, throat), its quality (pulling, pressure, heat, restlessness), and its changing intensity moment by moment, without either acting on it or fighting it. Through repeated practice, patients discover experientially that craving is time-limited and self-resolving, dismantling the belief that craving will continue escalating indefinitely unless satisfied through substance use.")

doc.add_heading("19.4 SOBER Breathing Space", level=2)
doc.add_body("The SOBER breathing space is a structured 2-3 minute emergency mindfulness technique designed for use in real-time high-risk situations. The acronym guides the patient through five steps: Stop (pause whatever you are doing, step out of automatic pilot), Observe (notice what is present, including thoughts, feelings, and body sensations, without judgment), Breathe (take three slow, deliberate breaths, using the breath as an anchor to the present moment), Expand (widen awareness from the breath to the whole body and the surrounding environment), and Respond (from this place of awareness, choose a skilful response rather than an automatic reaction). The SOBER space directly targets impulsivity by inserting a structured pause between trigger and response, allowing the prefrontal cortex time to come online and participate in decision-making.")

doc.add_heading("19.5 Trigger Mapping", level=2)
doc.add_body("Trigger mapping involves identifying and documenting the patient's personal high-risk situations, divided into internal triggers (emotions such as boredom, loneliness, anger, shame, anxiety) and external triggers (people, places, times of day, environmental cues associated with prior use). This exercise is conducted collaboratively, with the therapist guiding the patient to identify their specific, personal triggers rather than generic ones. The resulting trigger map becomes a living document that the patient carries forward into recovery, serving as an early warning system. Awareness of one's triggers allows anticipatory rather than reactive responses to high-risk situations.")

doc.add_heading("19.6 Cognitive Decentering", level=2)
doc.add_body("Cognitive decentering, sometimes called cognitive defusion in the ACT literature, involves changing one's relationship to thoughts rather than changing the content of thoughts. In MBRP, patients learn to rephrase their craving-related and permission-giving thoughts using observational language: instead of 'I need to use,' the patient practises saying 'I am having the thought that I need to use.' This small linguistic shift creates psychological distance between the self and the thought, reducing the thought's power to automatically drive behaviour. The technique targets both craving (by reducing the compelling quality of craving thoughts) and impulsivity (by creating a metacognitive space between thought and action).")

doc.add_heading("19.7 Loving-Kindness Meditation (Metta)", level=2)
doc.add_body("Loving-kindness meditation involves the systematic cultivation of feelings of warmth, care, and goodwill, initially directed toward oneself and progressively extended to loved ones, neutral persons, and eventually all beings. In the context of substance dependence treatment, loving-kindness practice serves several functions: it counters the pervasive shame and self-criticism that characterise early recovery, builds self-compassion which reduces the Abstinence Violation Effect (the tendency to catastrophise after a lapse), and cultivates positive affective states that serve as alternatives to substance-induced pleasure. This technique primarily builds the Non-Judging facet of mindfulness.")

doc.add_heading("19.8 Mindful Movement", level=2)
doc.add_body("Mindful movement involves gentle physical activity, typically walking or simple yoga stretches, performed with full present-moment attention to bodily sensations, balance, and the experience of movement. Unlike exercise for fitness, the purpose is entirely attentional: maintaining moment-to-moment awareness of the body in motion. This technique is particularly valuable for patients who find sitting meditation difficult due to restlessness (common in early recovery), as it provides an accessible entry point into mindfulness practice that engages the body directly. It builds the Observing facet and provides variety in the practice repertoire.")

doc.add_page_break()


# ═══════════════ SECTION 20: SESSION PROTOCOL WITH HINGLISH SCRIPTS ═══════════════
doc.add_heading("20. B-MBRP SESSION-BY-SESSION PROTOCOL WITH HINGLISH SCRIPTS", level=1)
doc.add_empty_line()

doc.add_heading("Session 1: Autopilot and Awareness", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Raisin/Murmura Exercise + Brief Body Scan", italic=True, size=22, spacing_after=200)

doc.add_body("The first session introduces the foundational concept of automatic pilot and establishes mindfulness as a psychological skill. The session begins with a 5-minute welcome and orientation, establishing group norms (confidentiality, non-judgment, participation). The therapist explains that this is not religious meditation but a psychological training that strengthens the mind, similar to how physical exercise strengthens the body.")

doc.add_body("The Raisin/Murmura Exercise (10 minutes) is the first experiential demonstration of mindfulness. Each participant is given a single piece of murmura (puffed rice) or a raisin. The therapist guides them to explore it as if encountering it for the first time:")

doc.add_para("Hinglish Script for Murmura Exercise:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Sabse pehle, ise apne haath mein rakhiye. Dekhiye ise dhyan se -- iska rang, iska shape, koi lines hain kya iske upar? Ab ise apni ungliyon mein ghuma ke dekhiye -- kaisa lagta hai touch karne mein? Halka hai ya bhaari? Ab ise apne kaan ke paas le jaiye aur halka sa dabaye -- koi awaaz aati hai? Ab ise apni naak ke paas laiye -- koi smell aati hai? Bahut halki bhi ho sakti hai. Ab dheere se ise apne hothon pe rakhiye -- abhi mat khaiye. Bas notice kijiye ki aapki jeebh kya kar rahi hai, muh mein kya ho raha hai. Ab dheere se muh mein rakhiye -- abhi mat chabaiye. Bas feel kijiye. Ab ek baar chabaiye -- kaisa taste aa raha hai? Kaise texture hai? Ab dheere dheere chabaiye aur nigal jaiye. Notice kijiye ki kaise yeh poora experience alag tha uss tarah se jaise hum normally khaana khaate hain.\"")

doc.add_body("After the exercise, the therapist connects it to substance use: \"Jaise aapne abhi dekha ki normally hum khaana autopilot pe khaate hain -- dekhte bhi nahi ki kya kha rahe hain. Usi tarah substance use bhi aksar autopilot pe hota hai. Trigger aata hai, craving hoti hai, use ho jaata hai -- bina soche samjhe. Jo humne abhi kiya -- dhyan se, hosh mein, present moment mein rehna -- yahi mindfulness hai. Aur yahi seekhna hai humein aane wale sessions mein.\"")

doc.add_body("The session concludes with a Brief Body Scan (12 minutes) where the therapist guides attention systematically through the body:")

doc.add_para("Hinglish Script for Brief Body Scan:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Aaraam se baith jaiye. Aankhein band kar lijiye. Teen lambi saansein lijiye... Ab apna dhyan apne pair ki ungliyon pe le jaiye. Kya feel ho raha hai wahaan? Koi sensation hai? Tingling, garmi, thanda, ya kuch bhi nahi -- jo bhi hai, bas notice kijiye. Badalne ki zaroorat nahi. Ab dheere se dhyan upar laiye -- talwe, ankle, pindli... Bas observe kijiye. Koi bhi feeling aaye -- acchi ya buri -- judgement mat kijiye. Bas dekh rahe hain... Ab ghutne, jaaghein... Ab pet ka area -- yahan aksar craving feel hoti hai. Jo bhi sensation hai, bas notice kijiye. Chhati... kaandhe... baajuein... haath... gardan... chehra... aur sir ka upar ka hissa. Ab poore sharir ko ek saath feel kijiye -- saas le rahe hain, poora sharir saas le raha hai. Teen saansein lijiye aur dheere se aankhein kholiye.\"")

doc.add_body("Home Practice assigned: 5-minute body scan daily, and notice one moment per day when you were operating on autopilot.")
doc.add_empty_line()

doc.add_heading("Session 2: Triggers and Body Scan", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Full Body Scan + Trigger Mapping + SOBER Introduction", italic=True, size=22, spacing_after=200)

doc.add_body("Session 2 begins with a brief review of home practice (5 minutes), normalising difficulties and celebrating any moments of awareness. The Full Body Scan (15 minutes) extends the practice from Session 1, spending more time in areas where patients commonly hold tension or craving sensations (stomach, chest, hands).")

doc.add_body("The Trigger Mapping Exercise (15 minutes) is conducted collaboratively. The therapist draws a simple framework on paper: Internal Triggers (emotions like boredom, akela hona, gussa, sharam, tension) and External Triggers (jagah, log, time of day, paisa milna, jhagda ke baad). Each participant identifies their personal triggers specific to their own life history:")

doc.add_para("Hinglish Script for Trigger Mapping:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Ab hum milke dekhenge ki aapke liye woh kaunsi situations hain jab craving sabse zyada aati hai. Do tarah ke triggers hote hain -- andar ke aur baahar ke. Andar ke matlab emotions -- jaise bore ho rahe hain, akele feel ho raha hai, gussa aa raha hai, ya tension hai. Baahar ke matlab log, jagah, samay -- jaise koi purana dost mil gaya, us gali se guzre jahan pehle lete the, ya raat ko neend nahi aa rahi. Aap sochiye -- aapke liye kaunse hain? Ek ek karke bataiye, hum likhte hain. Yeh aapka personal map hai -- yeh kisi aur ka nahi hai, sirf aapka.\"")

doc.add_body("The session concludes with a brief introduction to the SOBER technique (10 minutes), which will be practised fully in Session 3. The therapist explains the acronym and demonstrates one quick example.")

doc.add_body("Home Practice: Full body scan daily. Maintain a trigger diary: when craving came, what was the trigger (internal/external), what did you notice in your body.")
doc.add_empty_line()

doc.add_heading("Session 3: Breath Meditation and SOBER Space", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Breath Meditation + SOBER Role-Play", italic=True, size=22, spacing_after=200)

doc.add_body("After practice review (5 minutes), the therapist introduces formal Breath Meditation (10 minutes) as the core daily practice that patients will carry beyond discharge:")

doc.add_para("Hinglish Script for Breath Meditation:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Seedhe baith jaiye, lekin aaraam se. Aankhein band. Ab apna saara dhyan apni saas pe le aiye. Jab saas andar aa rahi hai -- feel kijiye naak mein thandi hawa. Jab baahar ja rahi hai -- feel kijiye garmi. Pet upar uth raha hai saas mein... neeche ja raha hai baahar. Bas yahi karna hai -- saas ko feel karna. Kuch badalna nahi hai, kuch control nahi karna. Jab mann bhatak jaaye -- aur bhatkega, yeh normal hai -- toh bas notice kijiye ki bhatka hai, aur dheere se wapas saas pe le aiye. Koi judgement nahi. Bhatak gaya? Theek hai. Wapas aao. Yahi practice hai. Har baar wapas aana -- yahi mindfulness ban rahi hai. Agar 10 baar bhatka aur 10 baar wapas laaye, toh aaj 10 baar mindfulness practice hui.\"")

doc.add_body("The SOBER Breathing Space is then taught formally (10 minutes) and practised through role-play (10 minutes):")

doc.add_para("Hinglish Script for SOBER:", bold=True, size=24, spacing_after=100)
doc.add_body("\"SOBER ka matlab hai -- S for Stop: ruk jao. Jo bhi kar rahe ho, ek second ke liye ruk jao. O for Observe: dekho kya ho raha hai andar. Kya thought aa raha hai? Kya feeling hai? Body mein kahan tension hai? B for Breathe: teen lambi saansein lo. Dheere se. Saas pe dhyan. E for Expand: ab poore sharir ko feel karo, apne aas paas ko dekho, yahaan ho tum abhi. R for Respond: ab hosh mein choose karo -- kya karna hai. Autopilot pe mat jao. Ab iska use karke dekhte hain -- socho tumhara wo purana dost milta hai aur kehta hai 'chal ek baar...' -- us moment mein SOBER karo.\"")

doc.add_body("The role-play involves one participant describing a personal high-risk scenario while the group practises deploying SOBER in response.")

doc.add_body("Home Practice: 10-minute breath meditation daily. Use SOBER at least 3 times daily (even for small urges or stressful moments). Record in diary.")
doc.add_empty_line()


doc.add_heading("Session 4: Urge Surfing and Decentering", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Guided Urge Surfing + Cognitive Decentering", italic=True, size=22, spacing_after=200)

doc.add_body("This is the most critical therapeutic session of the protocol, directly targeting the craving mechanism. After practice review (5 minutes), the therapist introduces the wave metaphor and guides a full Urge Surfing exercise (15 minutes):")

doc.add_para("Hinglish Script for Urge Surfing:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Aankhein band kijiye. Ab main chahta hoon ki aap apne mann mein woh situation laiye jab aapko sabse zyada craving hoti hai. Shayad woh jagah, woh log, ya woh feeling... Dheere dheere us scene ko yaad kijiye... Ab notice kijiye -- kya ho raha hai aapke sharir mein abhi? Kahaan feel ho rahi hai craving? Pet mein? Chhaati mein? Gale mein? Haathon mein?... Us jagah pe dhyan le jaiye. Kaisa lag raha hai? Khinchav? Dabav? Garmi? Becheni?... Ab bus isko dekhiye. Ladna nahi hai isse. Bhagana nahi hai. Bas dekh rahe hain jaise kinare pe khade hokar samundar ki lahar dekh rahe hain. Lahar uthti hai... upar jaati hai... aur phir neeche aati hai. Craving bhi lahar hai -- uthegi, peak pe jaayegi, aur khud kam ho jaayegi. Aapko kuch karna nahi hai. Bas surf karna hai -- is lahar ke upar rehna hai. Teen saansein lijiye... Ab notice kijiye -- kya intensity wahi hai jo pehle thi? Ya kuch badla hai?\"")

doc.add_body("The Cognitive Decentering exercise (10 minutes) follows, teaching patients to observe thoughts as mental events rather than commands:")

doc.add_para("Hinglish Script for Cognitive Decentering:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Ab ek exercise aur karte hain. Jab craving aati hai toh mann mein kuch thoughts aate hain -- jaise 'mujhe lena hai,' 'nahi reh sakta bina iske,' 'bas ek baar.' Yeh thoughts bahut powerful lagte hain. Lekin yeh thoughts hain -- facts nahi. Ab ise aise boliye: 'Mujhe ek thought aa raha hai ki mujhe lena hai.' Fark samjhe? Pehle thought aur aap ek the. Ab aap alag hain, thought alag hai. Aap thought nahi ho. Aap woh insaan ho jo thought ko dekh raha hai. Thought patte ki tarah hai jo paani mein beh raha hai -- aata hai, jaata hai. Aapko uske peeche nahi jaana.\"")

doc.add_body("Group sharing (15 minutes) follows where participants discuss their experience of urge surfing and decentering, with the therapist normalising the difficulty and reinforcing the key learning.")

doc.add_body("Home Practice: When craving arises, use urge surfing. Observe the craving in the body for 5 minutes without acting. Notice what happens to its intensity. Record in diary.")
doc.add_empty_line()

doc.add_heading("Session 5: Acceptance and Non-Reactivity", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Open Awareness Meditation + Acceptance vs. Avoidance Discussion", italic=True, size=22, spacing_after=200)

doc.add_body("Session 5 deepens the work begun in Session 4, focusing on the broader principle of accepting difficult experiences rather than escaping them through substance use. After practice review (5 minutes), the therapist guides an Open Awareness Meditation (12 minutes):")

doc.add_para("Hinglish Script for Open Awareness Meditation:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Aaraam se baith jaiye. Aankhein band. Pehle kuch saansein saas pe dhyan dete hue... Ab dheere se apna dhyan khol dijiye -- saas ke alawa bhi jo kuch aa raha hai, use aane dijiye. Koi awaaz sunaai de rahi hai? Aane dijiye, jaane dijiye. Koi thought aa raha hai? Aane dijiye, jaane dijiye. Koi feeling -- boredom, becheni, ya kuch aur? Aane dijiye. Kuch bhi rokna nahi hai, kuch bhi pakadna nahi hai. Aap bus ek khuli jagah hain jahan sab kuch aa sakta hai aur ja sakta hai. Agar koi mushkil feeling aaye -- tension, gussa, udaasi -- toh usse bhi aane dijiye. Dekhiye kahan feel ho rahi hai body mein. Us feeling ke saath ek minute rahiye. Woh aapko nuksan nahi pahuncha sakti. Woh feeling hai -- woh aap nahi hain. Aap pahaad hain -- feelings mausam hain. Mausam badalta rehta hai, pahaad waheen rehta hai.\"")

doc.add_body("The Acceptance vs. Avoidance Discussion (10 minutes) addresses the core insight that substance use is emotional avoidance made chemical. The therapist draws the avoidance cycle:")

doc.add_body("\"Mushkil feeling aati hai --> asahaniya lagti hai --> use karte hain --> 2 ghante relief --> phir guilt, sharam --> aur mushkil feelings --> phir use. Yeh chakkar kabhi khatam nahi hota agar hum feelings se bhaag rahe hain. MBRP mein hum ulta karte hain -- feelings ki taraf muh karke khade hote hain. Dekhte hain ki yeh feeling itni khatarnak nahi hai jitni lagti thi. Yeh bardaasht hoti hai. Aur dheere dheere kamzor padti hai.\"")

doc.add_body("Skilful Action Planning (10 minutes) helps patients identify concrete alternative actions for each of their mapped triggers, creating an 'if-then' plan: if this trigger occurs, then I will do this specific action (call someone, go for a walk, use SOBER, sit with the feeling for 5 minutes).")

doc.add_body("Home Practice: Open awareness meditation daily. When a difficult emotion arises, sit with it for 3 minutes before doing anything. Notice if it changes on its own.")
doc.add_empty_line()

doc.add_heading("Session 6: Integration and Maintenance", level=2)
doc.add_para("Duration: 45 minutes | Techniques: Loving-Kindness Meditation + Relapse Prevention Plan", italic=True, size=22, spacing_after=200)

doc.add_body("The final session consolidates all learnings and prepares the patient for post-discharge independent practice. After practice review (5 minutes), the therapist guides a Loving-Kindness Meditation (10 minutes):")

doc.add_para("Hinglish Script for Loving-Kindness Meditation:", bold=True, size=24, spacing_after=100)
doc.add_body("\"Aankhein band kijiye. Aaraam se saas lijiye. Ab apne baare mein sochiye -- jaisa bhi hoon, jo bhi galtiyan ki hain, jo bhi hua hai -- abhi is moment mein main apne aap ko acchi feelings bhejta hoon. Mann mein boliye: 'Main sukhi rahoon. Main swasth rahoon. Main surakshit rahoon. Mera dukh kam ho.' Dheere dheere... baar baar... Ab apne kisi apne ke baare mein sochiye -- maa, biwi, bhai, bachcha -- unke liye boliye: 'Woh sukhi rahen. Woh swasth rahen. Unka dukh kam ho.' Ab thoda aur badhate hain -- is ward mein jo log hain, sab ke liye: 'Sab sukhi rahen. Sab swasth rahen. Sabka dukh kam ho.' Aur last mein -- apne aap ke liye phir se: 'Main theek hoon. Main is raaste pe hoon. Main apna khayal rakh sakta hoon.'\"")

doc.add_body("The Personal Relapse Prevention Plan (15 minutes) is created collaboratively in session, documented on paper that the patient takes home:")

doc.add_body("The plan includes: (a) My personal high-risk situations and triggers (from trigger map), (b) My early warning signs that indicate I am moving toward relapse, (c) My SOBER practice plan for when triggered, (d) My daily mindfulness practice plan (which practice, what time, how long), (e) Two support contacts (name and phone number) to call when in difficulty, (f) What I will do if a lapse occurs: 'Ek baar use karna matlab poori tarah fail nahi hai. Main SOBER karoonga, apne support person ko call karoonga, aur wapas practice pe aaoonga.'")

doc.add_body("Group Feedback and Closure (15 minutes) allows each participant to share what was most useful, what they will continue practising, and their commitment to recovery. The therapist closes with warmth and confidence in each participant's capacity to use these tools independently.")

doc.add_body("Home Practice (Post-Discharge): 15 minutes daily practice (any combination of breath, body scan, or open awareness). Use SOBER whenever triggered. Keep practice log for accountability. Follow the personal relapse prevention plan.")

doc.add_page_break()


# ═══════════════ SECTION 21: HOW EACH SESSION IS CONDUCTED ═══════════════
doc.add_heading("21. HOW EACH B-MBRP SESSION IS CONDUCTED", level=1)
doc.add_empty_line()

doc.add_body("Every B-MBRP session follows a standardised 6-step structure to ensure consistency, predictability for participants, and treatment fidelity:")

doc.add_body("Step 1 - Opening Meditation (5 minutes): Each session begins with a brief breath awareness exercise to settle the group. Participants sit comfortably with eyes closed while the therapist guides attention to the breath in Hindi. This serves as a transition from the ward environment to the therapeutic space and immediately activates the mindful awareness that will be built upon throughout the session.")

doc.add_body("Step 2 - Practice Review (5 minutes): Participants share their experiences with home practice since the last session. The therapist enquires about what was practised, any difficulties encountered, any moments of mindfulness noticed in daily life, and any questions. Difficulties are normalised (e.g., 'mann bhatak jaana normal hai, wapas laana hi practice hai'), and small successes are acknowledged. This builds accountability and allows troubleshooting of barriers to practice.")

doc.add_body("Step 3 - Core Meditation or Exercise (12-15 minutes): This is the session-specific guided meditation or experiential exercise (body scan, urge surfing, open awareness, etc.). The therapist uses pre-prepared Hindi scripts delivered in a calm, unhurried voice. Audio recordings in Hindi are available for home practice. The pace is slow and the language is simple, accessible, and free of technical jargon.")

doc.add_body("Step 4 - Psychoeducation and Discussion (10-12 minutes): A brief teaching on the session theme (e.g., the wave model of craving, how automatic pilot works, the avoidance cycle) is followed by group discussion linking the teaching to participants' personal recovery experiences. Culturally relevant examples and metaphors are used throughout (e.g., river/nadi for flow of thoughts, mountain/pahaad for stability, wave/lahar for craving).")

doc.add_body("Step 5 - Experiential Exercise (8-10 minutes): An active practice component such as SOBER role-play in a high-risk scenario, trigger mapping on paper, or decentering practice with actual craving thoughts. This ensures that learning is not merely cognitive but experiential and embodied.")

doc.add_body("Step 6 - Closing and Home Practice Assignment (5 minutes): The session key learning is summarised in one or two sentences. A specific daily practice is assigned (5-15 minutes depending on the session). Audio recordings are provided for guided practices. Attendance is recorded. The session closes with three collective breaths.")

doc.add_body("All sessions are delivered by the researcher (MPhil Clinical Psychologist) following a manualized protocol. Supervision is provided by the research guide. Session logs are maintained documenting date, duration, attendance, adherence to protocol, and any significant observations.")

doc.add_page_break()

# ═══════════════ SECTION 22: CONTROL GROUP ═══════════════
doc.add_heading("22. CONTROL GROUP: PSYCHOEDUCATION PROTOCOL (6 SESSIONS)", level=1)
doc.add_empty_line()

doc.add_body("The control group receives 6 sessions of structured psychoeducation, matched to the experimental group for time (45 minutes per session), frequency (twice weekly), total duration (3 weeks), format (group of 6-8), and therapist contact. The sessions are informational and discussion-based, containing NO mindfulness, meditation, or contemplative practice component whatsoever. This attention-matched design allows the study to isolate the specific effects of mindfulness training from non-specific therapeutic factors.")

doc.add_heading("Session 1: Understanding Addiction", level=3)
doc.add_body("Content covers the disease model of addiction, brain changes produced by chronic substance use, genetic and environmental risk factors, and how dependence develops as a chronic relapsing condition. Discussion addresses myths about addiction (e.g., 'it is just a matter of willpower') and provides accurate psychoeducation.")

doc.add_heading("Session 2: Effects of Opioids", level=3)
doc.add_body("Content covers physical consequences of chronic opioid use (hepatic, cardiac, immune, respiratory), psychological effects (cognitive impairment, emotional blunting, depression), the withdrawal timeline, and the concept of post-acute withdrawal syndrome (PAWS) that may persist for months after cessation.")

doc.add_heading("Session 3: Understanding Relapse", level=3)
doc.add_body("Content covers relapse warning signs, high-risk situations (Marlatt's taxonomy), the relapse process as a gradual rather than sudden event, common cognitive distortions that precede relapse (apparently irrelevant decisions, permission-giving thoughts), and the distinction between lapse and relapse.")

doc.add_heading("Session 4: Health and Nutrition", level=3)
doc.add_body("Content covers physical recovery during treatment, the importance of sleep hygiene, benefits of regular exercise for mood and craving management, nutritional needs during early recovery, and managing common physical complaints (fatigue, appetite changes, pain).")

doc.add_heading("Session 5: Social Consequences", level=3)
doc.add_body("Content covers the impact of substance use on family relationships, legal issues, stigma in Indian society, workplace and financial consequences, and available rehabilitation resources and support systems.")

doc.add_heading("Session 6: Motivation and Goals", level=3)
doc.add_body("Content covers the Stages of Change model (Prochaska & DiClemente), recovery planning, realistic goal-setting for the first month post-discharge, identifying community resources (NA meetings, OPD follow-up, family support), and discharge planning.")

doc.add_page_break()


# ═══════════════ SECTION 23: DATA ANALYSIS ═══════════════
doc.add_heading("23. DATA ANALYSIS PLAN", level=1)
doc.add_empty_line()

doc.add_body("The data analysis proceeds through a systematic sequence designed to address the study hypotheses with appropriate statistical rigour:")

doc.add_heading("23.1 Descriptive Analysis", level=2)
doc.add_body("Descriptive statistics including mean, standard deviation, frequencies, and range will be computed for all sociodemographic variables (age, education, duration of use, number of previous admissions) and for all pre-test and post-test scores on the three outcome measures. This provides a comprehensive description of the sample and allows readers to assess the generalisability of findings.")

doc.add_heading("23.2 Normality Testing", level=2)
doc.add_body("The Shapiro-Wilk test will be used to assess the normality of distribution for all continuous outcome variables. Given the sample size of 30 per group, the Shapiro-Wilk test is preferred over the Kolmogorov-Smirnov test as it provides greater statistical power for smaller samples. If normality assumptions are violated, non-parametric alternatives will be employed.")

doc.add_heading("23.3 Baseline Equivalence", level=2)
doc.add_body("Independent samples t-tests (or Mann-Whitney U if non-normal) will be conducted on all pre-test scores to verify that randomisation produced equivalent groups at baseline. This is essential for confirming that any post-test differences can be attributed to the intervention rather than pre-existing group differences.")

doc.add_heading("23.4 Primary Analysis: ANCOVA", level=2)
doc.add_body("The primary analysis for testing all three null hypotheses is one-way Analysis of Covariance (ANCOVA). For each dependent variable separately, the model includes: Dependent Variable (post-test score on OCDUS, BIS-11, or FFMQ), Independent Variable (Group: Experimental vs. Control), and Covariate (pre-test score on the same measure). ANCOVA is chosen over simple t-tests or repeated measures ANOVA for several reasons: it statistically controls for any baseline differences between groups, it reduces error variance by accounting for individual variation in pre-test scores, it provides a more precise estimate of the treatment effect, it reduces the required sample size compared to simpler analyses, and it is the recommended analysis for pre-test post-test designs in intervention research (Tabachnick & Fidell, 2013).")

doc.add_heading("23.5 Within-Group Analysis", level=2)
doc.add_body("Paired samples t-tests (or Wilcoxon signed-rank tests if non-normal) will be conducted separately for each group to examine within-group change from pre-test to post-test on each outcome variable. This shows whether each group improved individually, independent of the between-group comparison.")

doc.add_heading("23.6 Effect Size Reporting", level=2)
doc.add_body("Effect sizes will be reported for all comparisons: partial eta-squared for ANCOVA models (small = 0.01, medium = 0.06, large = 0.14) and Cohen's d for t-test comparisons (small = 0.20, medium = 0.50, large = 0.80). Effect size reporting allows comparison with published MBRP studies and informs sample size calculations for future confirmatory trials.")

doc.add_heading("23.7 Significance Level", level=2)
doc.add_body("All tests will use an alpha level of 0.05 (two-tailed). Given that three primary ANCOVAs are conducted (one per dependent variable), Bonferroni correction will be applied, yielding an adjusted alpha of 0.017 for each primary test. Both uncorrected and Bonferroni-corrected p-values will be reported. All analyses will be conducted using IBM SPSS Statistics version 26.0.")

doc.add_heading("23.8 Intention-to-Treat Analysis", level=2)
doc.add_body("For participants who complete pre-test but not post-test (due to premature discharge or withdrawal), Last Observation Carried Forward (LOCF) will be used as the primary intention-to-treat approach. A sensitivity analysis comparing ITT results with per-protocol results (completers only) will be reported to assess the robustness of findings.")

doc.add_page_break()

# ═══════════════ SECTION 24: ETHICAL CONSIDERATIONS ═══════════════
doc.add_heading("24. ETHICAL CONSIDERATIONS", level=1)
doc.add_empty_line()

doc.add_body("The present study adheres strictly to the ICMR (2017) National Ethical Guidelines for Biomedical and Health Research Involving Human Participants, and the following ethical principles are observed:")

doc.add_body("Informed Consent: Written informed consent is obtained from all participants in Hindi (with English version available). The consent process explains the study purpose, procedures, duration, potential benefits and risks, voluntary nature of participation, and right to withdraw at any time without penalty or impact on treatment received. Participants are given adequate time to consider their decision and ask questions. For participants who cannot write, a thumb impression witnessed by a ward staff member is accepted.")

doc.add_body("Voluntary Participation: Participation is entirely voluntary. Refusal to participate or withdrawal at any point does not affect the treatment received at the de-addiction centre. This is explicitly stated in the consent form and verbally during the consent process.")

doc.add_body("Confidentiality: All data is stored using participant ID codes only. No identifying information (names, addresses, phone numbers) appears on any assessment form or data file. A master code key linking participant codes to names is stored separately in a locked location accessible only to the researcher and guide. Electronic data is password-protected.")

doc.add_body("Non-Maleficence: The control group receives active psychoeducation (not a no-treatment waitlist), ensuring all participants receive a structured psychological intervention. TAU continues for all participants throughout the study. Additionally, following completion of data collection, the control group is offered a brief MBRP orientation session as a debriefing measure.")

doc.add_body("Institutional Approval: Ethical clearance from the Institutional Ethics Committee (IEC) is obtained prior to any data collection. The full study protocol, consent form, assessment tools, and intervention outline are submitted for review.")

doc.add_body("Compliance: The study complies with all applicable sections of the ICMR (2017) guidelines, the Declaration of Helsinki, and the institutional requirements of the college and university.")

doc.add_page_break()


# ═══════════════ SECTION 25: EXPECTED RESULTS ═══════════════
doc.add_heading("25. EXPECTED RESULTS", level=1)
doc.add_empty_line()

doc.add_body("Based on the theoretical rationale, existing evidence from MBRP trials, and the mechanisms of action described above, the following results are anticipated:")

doc.add_body("Craving (OCDUS): A statistically significant reduction in craving scores is expected in the Experimental Group (Brief MBRP + TAU) compared to the Control Group (Psychoeducation + TAU) at post-test, with an expected effect size of d = 0.50 to 0.80. The mechanism for this expected change is the disruption of the automatic craving-use cycle through urge surfing and SOBER breathing space practice. Patients in the MBRP group are expected to develop the capacity to observe craving as a transient phenomenon rather than an imperative to act, reducing both the subjective intensity and the behavioural pull of craving episodes.")

doc.add_body("Impulsivity (BIS-11): A statistically significant reduction in impulsivity scores, particularly on the Motor Impulsivity and Attentional Impulsivity subscales, is expected in the Experimental Group compared to the Control Group, with an anticipated effect size of d = 0.40 to 0.60. The mechanism is the strengthening of prefrontal inhibitory control through repeated mindfulness practice, which trains the capacity to notice an impulse and choose not to act on it. The mindful pause inherent in all MBRP techniques directly exercises the neural circuits responsible for response inhibition.")

doc.add_body("Mindfulness (FFMQ): A statistically significant increase in mindfulness scores, particularly on the Acting with Awareness and Non-Reactivity facets, is expected in the Experimental Group compared to the Control Group, with an anticipated effect size of d = 0.50 to 0.70. This is the most directly expected outcome, as the intervention specifically trains mindfulness capacity through 6 sessions of structured meditation practice and daily home practice. The FFMQ Non-Reactivity subscale is predicted to show the largest change, as it directly measures the capacity that urge surfing and SOBER breathing space are designed to develop.")

doc.add_body("Overall, the three null hypotheses are expected to be rejected, supporting the efficacy of Brief MBRP as an adjunct to TAU in reducing craving and impulsivity while enhancing mindfulness in substance-dependent patients at a de-addiction centre.")

doc.add_page_break()

# ═══════════════ SECTION 26: CLINICAL IMPLICATIONS ═══════════════
doc.add_heading("26. CLINICAL IMPLICATIONS", level=1)
doc.add_empty_line()

doc.add_body("The findings of the present study, if supported, carry several important clinical implications for the practice of de-addiction treatment in India:")

doc.add_body("First, the study validates a brief MBRP model (6 sessions over 3 weeks) that is feasible for Indian de-addiction settings operating with limited resources, short admission windows, and heavy patient loads. Unlike the standard 8-week outpatient protocol, this brief adaptation can be completed within a typical Indian IPD admission, making it practically implementable without requiring changes to existing treatment infrastructure.")

doc.add_body("Second, the study provides evidence for a structured, evidence-based psychological intervention that complements pharmacotherapy (OST, naltrexone) in routine de-addiction care. This addresses the critical gap between pharmacological and psychological treatment that characterises current Indian practice, offering a concrete protocol that can be delivered alongside existing medical management.")

doc.add_body("Third, the study demonstrates the cultural compatibility of mindfulness approaches with Indian populations. While mindfulness has ancient roots in Indian contemplative traditions, its clinical application in modern healthcare settings requires careful adaptation and empirical validation. The present study, conducted in Hindi with culturally appropriate metaphors and examples, establishes this compatibility.")

doc.add_body("Fourth, the multi-target nature of the protocol is clinically significant: a single integrated intervention addresses craving, impulsivity, and mindfulness simultaneously, rather than requiring separate interventions for each domain. This is both clinically efficient and theoretically coherent, given the mechanistic interconnections between these three constructs.")

doc.add_body("Fifth, the protocol is deliverable by MPhil-trained Clinical Psychologists, making it appropriate for task-shifting at district-level facilities where specialist availability is limited. This has implications for scalability: if effective, the protocol can potentially be disseminated to government de-addiction centres across India through training programmes for clinical psychology trainees and counsellors.")

doc.add_page_break()

# ═══════════════ SECTION 27: LIMITATIONS ═══════════════
doc.add_heading("27. LIMITATIONS", level=1)
doc.add_empty_line()

doc.add_body("The present study acknowledges the following limitations:")

doc.add_body("First, the sample is restricted to male participants from a single de-addiction centre, which limits generalisability to female populations and to other treatment settings with different patient demographics, treatment cultures, or resource availability. However, this restriction enhances internal validity through sample homogeneity.")

doc.add_body("Second, the assessment is limited to immediate post-test measurement (within one week of intervention completion), without a longer-term follow-up. This means the study cannot determine whether intervention effects are maintained over weeks or months in the community, which is the critical question for relapse prevention. A follow-up assessment at 3 or 6 months post-discharge would strengthen the study's clinical relevance considerably.")

doc.add_body("Third, all outcome measures are self-report instruments, which are susceptible to social desirability bias, demand characteristics, and limited insight. Participants in the experimental group may report improvements because they believe they should, rather than because genuine change has occurred. This limitation is partially addressed through blinded assessment, but cannot be fully eliminated without objective measures.")

doc.add_body("Fourth, no biological markers of craving are included. Craving is measured entirely through subjective self-report, without physiological corroboration (e.g., heart rate variability, skin conductance, or cortisol reactivity to cue exposure). Future studies should consider multimodal assessment of craving.")

doc.add_body("Fifth, the same researcher delivers both interventions (MBRP and psychoeducation), which introduces the potential for therapist allegiance effects. The researcher's investment in the mindfulness protocol may subtly influence delivery quality, enthusiasm, or non-verbal communication in ways that favour the experimental condition. This is a common limitation in MPhil-level intervention research and is partially mitigated through manualized delivery.")

doc.add_body("Sixth, while the attention-matched control design is a methodological strength, it does not fully isolate mindfulness-specific mechanisms from all non-specific factors. The experimental group receives a qualitatively different type of experience (meditation, body-based practice) that cannot be perfectly matched by informational sessions. Future studies might employ an active control with relaxation training to more fully isolate the mindfulness-specific contribution.")

doc.add_body("Seventh, potential attrition remains a concern despite over-recruitment, as substance-dependent populations are characterised by high rates of premature discharge (LAMA), treatment non-adherence, and loss to follow-up.")

doc.add_page_break()


# ═══════════════ SECTION 28: FUTURE DIRECTIONS ═══════════════
doc.add_heading("28. FUTURE DIRECTIONS", level=1)
doc.add_empty_line()

doc.add_body("The present study opens several avenues for future research:")

doc.add_heading("Follow-Up and Replication", level=2)
doc.add_bullet("Conduct 3-month and 6-month post-discharge follow-up assessments to determine whether B-MBRP effects are maintained in the community")
doc.add_bullet("Replicate the study across multiple sites in different geographical and cultural contexts within India")
doc.add_bullet("Include female participants in dedicated studies to examine gender-specific effects and barriers")

doc.add_heading("Mechanism Research", level=2)
doc.add_bullet("Incorporate neuroimaging (fMRI or EEG) to examine neural changes associated with B-MBRP, particularly in prefrontal and limbic circuits")
doc.add_bullet("Conduct mediator analyses to formally test whether changes in mindfulness mediate changes in craving and impulsivity")
doc.add_bullet("Examine dose-response relationships: compare 4-session, 6-session, and 8-session protocols")

doc.add_heading("Comparative Studies", level=2)
doc.add_bullet("Compare B-MBRP with other evidence-based approaches (CBT, ACT, Motivational Enhancement Therapy) in Indian de-addiction settings")
doc.add_bullet("Develop and evaluate technology-assisted delivery (smartphone app-based guided meditations for post-discharge support)")
doc.add_bullet("Examine the added value of post-discharge booster sessions via phone or OPD")

doc.add_heading("Dissemination", level=2)
doc.add_bullet("Develop a complete Hindi B-MBRP training manual for dissemination to other Indian de-addiction centres")
doc.add_bullet("Train counsellors and clinical psychologists in government centres through structured training workshops")
doc.add_bullet("Explore rural access through telehealth-delivered mindfulness sessions")

doc.add_page_break()

# ═══════════════ SECTION 29: STUDY SUMMARY ═══════════════
doc.add_heading("29. STUDY SUMMARY", level=1)
doc.add_empty_line()

doc.add_body("The following table summarises the key components of the present study:")
doc.add_empty_line()

summary_items = [
    ("Design", "Pre-test Post-test Control Group Experimental Design (True Experimental)"),
    ("Setting", "De-addiction Centre, Guna, Madhya Pradesh"),
    ("Population", "Male opioid-dependent patients, aged 18-50 years, N = 60 (30 per group)"),
    ("Experimental Group", "Brief MBRP (6 sessions, twice weekly, 3 weeks, 45 min each) + Treatment As Usual"),
    ("Control Group", "Psychoeducation (6 sessions, twice weekly, 3 weeks, 45 min each) + Treatment As Usual"),
    ("Dependent Variables", "Craving (OCDUS) + Impulsivity (BIS-11) + Mindfulness (FFMQ)"),
    ("Sampling", "Two-stage: Purposive selection followed by computer-generated random assignment"),
    ("Primary Analysis", "ANCOVA controlling for pre-test scores as covariates"),
    ("Hypotheses", "Null hypotheses: No significant difference between groups on OCDUS, BIS-11, FFMQ"),
    ("Effect Sizes Expected", "Craving d=0.50-0.80; Impulsivity d=0.40-0.60; Mindfulness d=0.50-0.70"),
    ("Timeline", "12 months (Ethics Month 1-2, Data Collection Months 3-8, Writing Months 9-12)"),
]

for label, value in summary_items:
    doc.add_para(f"{label}: {value}", bold=False, size=24, spacing_after=120)

doc.add_page_break()

# ═══════════════ SECTION 30: CONCLUSION ═══════════════
doc.add_heading("30. CONCLUSION", level=1)
doc.add_empty_line()

doc.add_body("Substance dependence, particularly opioid dependence, represents a significant and growing public health crisis in India, with relapse rates exceeding 70 percent within the first three months of discharge from de-addiction treatment. Current Indian de-addiction practice relies heavily on pharmacotherapy, with limited integration of evidence-based psychological interventions specifically designed for relapse prevention. This treatment gap leaves patients discharged with biological sobriety but without the psychological skills necessary to navigate craving, regulate difficult emotions, and inhibit impulsive responses to triggers in the community.")

doc.add_body("Mindfulness-Based Relapse Prevention (MBRP) offers a theoretically grounded, empirically supported approach that directly addresses the three psychological mechanisms most strongly implicated in relapse: craving (through urge surfing and non-reactive awareness), impulsivity (through the mindful pause and strengthened inhibitory control), and low mindfulness (through structured meditation training that builds present-moment awareness and non-reactivity). The international evidence base for MBRP is strong, with multiple randomised controlled trials demonstrating superiority over both standard relapse prevention and treatment as usual.")

doc.add_body("The present study represents a critical first step in establishing an evidence base for brief mindfulness-based psychological interventions within the Indian de-addiction treatment infrastructure. By evaluating a 6-session B-MBRP protocol that is feasible within the typical Indian IPD admission window, employing an attention-matched active control, and simultaneously assessing three theoretically linked outcome variables, the study addresses a specific, documented, and clinically important gap in the existing literature.")

doc.add_body("If the hypothesised results are supported, the findings will have direct clinical utility for the study setting and broader implications for Indian de-addiction practice. A brief, deliverable, culturally adapted mindfulness-based intervention that reduces craving and impulsivity while building protective mindfulness capacity would represent a valuable addition to the treatment armamentarium of Indian de-addiction centres. This study provides the foundational evidence for such an integration.")

doc.add_page_break()


# ═══════════════ SECTION 31: REFERENCES ═══════════════
doc.add_heading("31. REFERENCES", level=1)
doc.add_empty_line()

refs = [
    "Baer, R. A., Smith, G. T., Hopkins, J., Krietemeyer, J., & Toney, L. (2006). Using self-report assessment methods to explore facets of mindfulness. Assessment, 13(1), 27-45.",
    "Bowen, S., Chawla, N., & Marlatt, G. A. (2011). Mindfulness-based relapse prevention for addictive behaviors: A clinician's guide. Guilford Press.",
    "Bowen, S., & Marlatt, G. A. (2009). Surfing the urge: Brief mindfulness-based intervention for college student smokers. Psychology of Addictive Behaviors, 23(4), 666-671.",
    "Bowen, S., Witkiewitz, K., Clifasefi, S. L., Grow, J., Chawla, N., Hsu, S. H., ... & Larimer, M. E. (2014). Relative efficacy of mindfulness-based relapse prevention, standard relapse prevention, and treatment as usual for substance use disorders: A randomized clinical trial. JAMA Psychiatry, 71(5), 547-556.",
    "Brewer, J. A., Mallik, S., Desai, R., Kober, H., & Potenza, M. N. (2011). Mindfulness training for smoking cessation: Results from a randomized controlled trial. Drug and Alcohol Dependence, 119(1-2), 72-80.",
    "Chiesa, A., & Serretti, A. (2014). Are mindfulness-based interventions effective for substance use disorders? A systematic review of the evidence. Substance Use and Misuse, 49(5), 492-512.",
    "Cohen, S., Kamarck, T., & Mermelstein, R. (1983). A global measure of perceived stress. Journal of Health and Social Behavior, 24(4), 385-396.",
    "Franken, I. H. A., Hendriks, V. M., & van den Brink, W. (2002). Initial validation of two opiate craving questionnaires: The Obsessive Compulsive Drug Use Scale and the Desires for Drug Questionnaire. Addictive Behaviors, 27(5), 675-685.",
    "Garland, E. L., Froeliger, B., & Howard, M. O. (2014). Mindfulness training targets neurocognitive mechanisms of addiction at the attention-appraisal-emotion interface. Frontiers in Psychiatry, 4, 173.",
    "Garland, E. L., Roberts-Lewis, A., Tronnier, C. D., Graves, R., & Kelley, K. (2016). Mindfulness-Oriented Recovery Enhancement versus CBT for co-occurring substance dependence, traumatic stress, and psychiatric disorders. Journal of Consulting and Clinical Psychology, 84(4), 281-293.",
    "Ghosh, A., Basu, D., & Avasthi, A. (2018). Relapse in opioid use disorder: An Indian perspective. Indian Journal of Psychiatry, 60(Suppl 4), S469-S476.",
    "Glasner-Edwards, S., Mooney, L. J., Ang, A., Garneau, H. C., Hartwell, E., Brecht, M. L., & Rawson, R. A. (2017). Mindfulness-based relapse prevention for stimulant dependent adults: A pilot randomized clinical trial. Mindfulness, 8(1), 126-135.",
    "Grant, S., Colaiaco, B., Motala, A., Shanman, R., Booth, M., Sorbero, M., & Hempel, S. (2017). Mindfulness-based relapse prevention for substance use disorders: A systematic review and meta-analysis. Journal of Addiction Medicine, 11(5), 386-396.",
    "Humeniuk, R., Ali, R., Babor, T. F., Farrell, M., Formigoni, M. L., Jittiwutikarn, J., ... & Simon, S. (2008). Validation of the Alcohol, Smoking and Substance Involvement Screening Test (ASSIST). Addiction, 103(6), 1039-1047.",
    "Jain, R., Majumder, P., & Gupta, T. (2013). Pharmacological intervention of nicotine dependence. Indian Journal of Psychiatry, 55(Suppl 1), S86-S92.",
    "Kabat-Zinn, J. (1990). Full catastrophe living: Using the wisdom of your body and mind to face stress, pain, and illness. Delacorte Press.",
    "Karyadi, K. A., VanderVeen, J. D., & Cyders, M. A. (2014). A meta-analysis of the relationship between trait mindfulness and substance use behaviors. Drug and Alcohol Dependence, 143, 1-10.",
    "Li, W., Howard, M. O., Garland, E. L., McGovern, P., & Lazar, M. (2017). Mindfulness treatment for substance misuse: A systematic review and meta-analysis. Journal of Substance Abuse Treatment, 75, 62-96.",
    "Marlatt, G. A., & Gordon, J. R. (1985). Relapse prevention: Maintenance strategies in the treatment of addictive behaviors. Guilford Press.",
    "Mattoo, S. K., Chakrabarti, S., & Anjaiah, M. (2009). Psychosocial factors associated with relapse in men with alcohol or opioid dependence. Indian Journal of Medical Research, 130(6), 702-708.",
    "Ministry of Social Justice and Empowerment. (2019). Magnitude of substance use in India. Government of India.",
    "Moeller, F. G., Barratt, E. S., Dougherty, D. M., Schmitz, J. M., & Swann, A. C. (2001). Psychiatric aspects of impulsivity. American Journal of Psychiatry, 158(11), 1783-1793.",
    "Murphy, C., & MacKillop, J. (2012). Living in the here and now: Interrelationships between impulsivity, mindfulness, and alcohol misuse. Psychopharmacology, 219(2), 527-536.",
    "National Institute on Drug Abuse. (2020). Drugs, brains, and behavior: The science of addiction. National Institutes of Health.",
    "Patton, J. H., Stanford, M. S., & Barratt, E. S. (1995). Factor structure of the Barratt Impulsiveness Scale. Journal of Clinical Psychology, 51(6), 768-774.",
    "Robinson, T. E., & Berridge, K. C. (1993). The neural basis of drug craving: An incentive-sensitization theory of addiction. Brain Research Reviews, 18(3), 247-291.",
    "Sarkar, S., & Balhara, Y. P. S. (2016). Diabetes mellitus in alcohol use disorder: Where do we stand? Indian Journal of Psychiatry, 58(3), 290-295.",
    "Segal, Z. V., Williams, J. M. G., & Teasdale, J. D. (2002). Mindfulness-based cognitive therapy for depression: A new approach to preventing relapse. Guilford Press.",
    "Serre, F., Fatseas, M., Swendsen, J., & Auriacombe, M. (2015). Ecological momentary assessment in the investigation of craving and substance use in daily life: A systematic review. Drug and Alcohol Dependence, 148, 1-20.",
    "Stanford, M. S., Mathias, C. W., Dougherty, D. M., Lake, S. L., Anderson, N. E., & Patton, J. H. (2009). Fifty years of the Barratt Impulsiveness Scale: An update and review. Personality and Individual Differences, 47(5), 385-395.",
    "Tabachnick, B. G., & Fidell, L. S. (2013). Using multivariate statistics (6th ed.). Pearson.",
    "Tiffany, S. T., & Wray, J. M. (2012). The clinical significance of drug craving. Annals of the New York Academy of Sciences, 1248(1), 1-17.",
    "WHO ASSIST Working Group. (2002). The Alcohol, Smoking and Substance Involvement Screening Test (ASSIST): Development, reliability and feasibility. Addiction, 97(9), 1183-1194.",
    "Witkiewitz, K., Bowen, S., Douglas, H., & Hsu, S. H. (2013). Mindfulness-based relapse prevention for substance craving. Addictive Behaviors, 38(2), 1563-1571.",
]

for ref in refs:
    doc.add_para(ref, size=22, spacing_after=100)

doc.add_page_break()


# ═══════════════ APPENDIX A: CONSENT FORM ═══════════════
doc.add_heading("APPENDIX A: INFORMED CONSENT FORM", level=1)
doc.add_para("[EDITABLE - Modify as per your institutional requirements]", italic=True, size=22, spacing_after=200)
doc.add_empty_line()

doc.add_para("INFORMED CONSENT FOR PARTICIPATION IN RESEARCH", bold=True, size=28, alignment="center", spacing_after=200)
doc.add_empty_line()

doc.add_body("Title of Study: Efficacy of Brief Mindfulness-Based Relapse Prevention (B-MBRP) Intervention on Craving, Impulsivity, and Mindfulness in Substance-Dependent Patients", indent=False)
doc.add_body("Researcher: Tejas Dangodra, MPhil Clinical Psychology Trainee", indent=False)
doc.add_body("Guide: Dr. Himanshi Singh", indent=False)
doc.add_body("Institution: MAN College of Special Education & Psychological Studies, Guna (MP)", indent=False)
doc.add_empty_line()

doc.add_body("I, _________________________ (participant name), have been informed about the following:", indent=False)
doc.add_empty_line()
doc.add_bullet("The purpose of this study is to evaluate the effectiveness of a brief mindfulness-based intervention for substance dependence")
doc.add_bullet("My participation involves completing questionnaires (approximately 40 minutes) on two occasions and attending 6 group sessions over 3 weeks")
doc.add_bullet("My participation is entirely voluntary and I may withdraw at any time without giving any reason and without any effect on my treatment")
doc.add_bullet("All information collected will be kept strictly confidential and my identity will not be revealed in any publication")
doc.add_bullet("There are no known risks associated with participation in this study")
doc.add_bullet("I will not receive any monetary compensation for my participation")
doc.add_bullet("I have had the opportunity to ask questions and have received satisfactory answers")
doc.add_empty_line()

doc.add_body("I hereby voluntarily consent to participate in this research study.", indent=False)
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("Participant Signature / Thumb Impression: _________________________", size=24, spacing_after=100)
doc.add_para("Date: _________________________", size=24, spacing_after=100)
doc.add_para("Witness Name and Signature: _________________________", size=24, spacing_after=100)
doc.add_para("Researcher Signature: _________________________", size=24, spacing_after=100)

doc.add_page_break()

# ═══════════════ APPENDIX A (HINDI) ═══════════════
doc.add_heading("APPENDIX A (Hindi): Sahmati Patra", level=1)
doc.add_para("[EDITABLE - Modify as per your institutional requirements]", italic=True, size=22, spacing_after=200)
doc.add_empty_line()

doc.add_para("SHODH MEIN BHAGIDAARI KE LIYE SAHMATI PATRA", bold=True, size=26, alignment="center", spacing_after=200)
doc.add_empty_line()

doc.add_body("Shodh ka Sheerashak: Brief Mindfulness-Based Relapse Prevention (B-MBRP) ka Prabhav -- Craving, Impulsivity, aur Mindfulness par", indent=False)
doc.add_body("Shodhakarti: Tejas Dangodra, MPhil Clinical Psychology", indent=False)
doc.add_body("Margadarshak: Dr. Himanshi Singh", indent=False)
doc.add_body("Sanstha: MAN College of Special Education & Psychological Studies, Guna (MP)", indent=False)
doc.add_empty_line()

doc.add_body("Main, _________________________ (pratyashi ka naam), yeh sahmati deta hoon ki:", indent=False)
doc.add_empty_line()
doc.add_bullet("Mujhe is shodh ka uddeshya samjhaya gaya hai")
doc.add_bullet("Meri bhagidaari mein 2 baar prashnavali bharna (lagbhag 40 minute) aur 3 hafton mein 6 group sessions mein shamil hona hai")
doc.add_bullet("Meri bhagidaari poori tarah se swaichchhik hai aur main kisi bhi samay bina karan bataye apni bhagidaari wapas le sakta hoon")
doc.add_bullet("Mere ilaj par koi asar nahi padega chahe main bhag loon ya na loon")
doc.add_bullet("Meri saari jaankari guptiya rakhi jaayegi aur mere naam ka kisi bhi jagah zikr nahi hoga")
doc.add_bullet("Is shodh mein koi hani ka khatara nahi hai")
doc.add_bullet("Mujhe prashna poochhne ka avsar diya gaya aur santooshjanak uttar mile")
doc.add_empty_line()

doc.add_body("Main apni swaichchha se is shodh mein bhag lene ke liye sahmati deta hoon.", indent=False)
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("Pratyashi ke Hastaakshar / Angutha Nisshan: _________________________", size=24, spacing_after=100)
doc.add_para("Dinnank: _________________________", size=24, spacing_after=100)
doc.add_para("Saakshi ka Naam aur Hastaakshar: _________________________", size=24, spacing_after=100)
doc.add_para("Shodhakarti ke Hastaakshar: _________________________", size=24, spacing_after=100)

doc.add_page_break()


# ═══════════════ APPENDIX B: ASSESSMENT TOOLS ═══════════════
doc.add_heading("APPENDIX B: ASSESSMENT TOOLS", level=1)
doc.add_para("[EDITABLE - Add full tool items here. Space provided for pasting tools.]", italic=True, size=22, spacing_after=200)
doc.add_empty_line()

doc.add_heading("Tool 1: Obsessive Compulsive Drug Use Scale (OCDUS)", level=2)
doc.add_body("[Paste the 13 OCDUS items here with response options 0-4]", indent=False)
doc.add_empty_line()
doc.add_body("Scoring: Sum all 13 items. Total score range: 0-52. Higher scores = greater craving.", indent=False)
doc.add_empty_line()
doc.add_empty_line()

doc.add_heading("Tool 2: Barratt Impulsiveness Scale-11 (BIS-11)", level=2)
doc.add_body("[Paste the 30 BIS-11 items here with response options 1-4]", indent=False)
doc.add_empty_line()
doc.add_body("Scoring: Reverse-score items 1, 7, 8, 9, 10, 12, 13, 15, 20, 29, 30. Sum all items. Range: 30-120. Higher = more impulsive.", indent=False)
doc.add_empty_line()
doc.add_body("Subscales: Attentional (items 5, 9, 11, 20, 24, 26, 28), Motor (items 2, 3, 4, 16, 17, 19, 21, 22, 23, 25, 30), Non-Planning (items 1, 7, 8, 10, 12, 13, 14, 15, 18, 27, 29).", indent=False)
doc.add_empty_line()
doc.add_empty_line()

doc.add_heading("Tool 3: Five Facet Mindfulness Questionnaire (FFMQ)", level=2)
doc.add_body("[Paste the 39 FFMQ items here with response options 1-5]", indent=False)
doc.add_empty_line()
doc.add_body("Scoring: Reverse-score designated items. Sum facets: Observing (8 items), Describing (8), Acting with Awareness (8), Non-Judging (8), Non-Reactivity (7). Total range: 39-195. Higher = greater mindfulness.", indent=False)
doc.add_empty_line()
doc.add_empty_line()

doc.add_heading("Tool 4: WHO-ASSIST V3.0", level=2)
doc.add_body("[Paste the 8 ASSIST items here]", indent=False)
doc.add_empty_line()
doc.add_body("Scoring: Substance-specific risk scores. Low: 0-3, Moderate: 4-26, High: 27+. Used at PRE-TEST ONLY for baseline severity characterisation.", indent=False)
doc.add_empty_line()
doc.add_empty_line()

doc.add_heading("Sociodemographic and Clinical Proforma", level=2)
doc.add_body("[Editable - Design your proforma with the following variables]", indent=False)
doc.add_empty_line()
doc.add_body("Sociodemographic: Age, Gender, Education, Marital Status, Occupation, SES, Religion, Residence (Urban/Rural)", indent=False)
doc.add_empty_line()
doc.add_body("Clinical: Primary substance, ICD-10 diagnosis, Duration of use (years), Age at first use, Age at onset of dependence, Number of previous admissions, Number of previous relapses, Days since last use, Current medications, Family history of SUD, Motivation rating (1-10)", indent=False)

doc.add_page_break()

# ═══════════════ FINAL PAGE ═══════════════
doc.add_empty_line()
doc.add_empty_line()
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("--- END OF SYNOPSIS ---", bold=True, size=24, alignment="center", spacing_after=200)
doc.add_empty_line()
doc.add_empty_line()
doc.add_para("Researcher: Tejas Dangodra", size=24, alignment="center", spacing_after=100)
doc.add_para("MPhil Clinical Psychology (2025-2027)", size=22, alignment="center", spacing_after=100)
doc.add_para("MAN College of Special Education & Psychological Studies, Guna (MP)", size=22, alignment="center", spacing_after=100)
doc.add_para("tejasdangodra99@gmail.com | +91 8140171722", size=22, alignment="center", spacing_after=100)

# ═══════════════ BUILD THE DOCUMENT ═══════════════
output_file = "/projects/sandbox/Dango-kiro/MBRP_Research_Synopsis_Detailed.docx"
doc.build(output_file)
print(f"\nDocument generated successfully!")
print(f"File: {output_file}")
