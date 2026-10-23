# -*- coding: utf-8 -*-
"""
Generator for the self-contained offline HTML study-notes file:
"Biological Foundations of Behavior - Unit I: Nomenclature, Signs and Symptoms".

This module defines ONE CONTENT data structure (a list of 12 topic dicts) that is
the single source of truth for the HTML deliverable rendered here.

Markup conventions used inside point / intro / flow text:
  **word**                 -> a bold memory-hook keyword (a few per sentence only)
  [[plain-language text]]  -> a distinct bracketed plain-language explanation of a
                              difficult term, shown in its own visible style.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 generate_unit1_nomenclature_notes.py
"""

import base64
import html as _html
import os
import re

TITLE = ("Biological Foundations of Behavior \u2014 Unit I: "
         "Nomenclature, Signs and Symptoms")
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"
SUBTITLE = ("Paper I: Biological Foundations of Behavior \u00b7 Descriptive psychopathology "
            "after Fish \u00b7 Exam-ready study notes")
OUTPUT_HTML = "Biological_Foundations_Unit1_Nomenclature_and_Signs_Symptoms.html"

# ---------------------------------------------------------------------------
# HOW TO READ THESE NOTES (short usage guide; NOT about how the notes were made)
# ---------------------------------------------------------------------------
HOW_TO_READ = [
    "Each topic opens with a one-line **intro** that frames the idea and then an "
    "**Answer flow** box, a small ladder of stages (definition, background, key concepts, "
    "clinical examples, significance, one-line summary); write your exam answer in exactly that order.",
    "The numbered points below the Answer flow follow those same stages, so reading them "
    "top to bottom gives you a ready-made **5-mark answer** rather than a list of scattered notes.",
    "Hard psychopathology terms are kept in full, but each is followed by a plain-language "
    "meaning inside a shaded box like this [[a simple everyday explanation that sits right next "
    "to the difficult word]] so you keep the clinical vocabulary while understanding it.",
    "Only a **few keywords** in each sentence are shown in bold; treat those bold words as "
    "memory hooks and rebuild the rest of the sentence around them while you revise.",
    "Every topic carries exactly one **flowchart** of boxes joined by arrows to show a sequence, "
    "and exactly one **mnemonic** whose letters each map to a point you must remember.",
    "Several topics carry a **labelled diagram**; use each figure as a visual map, matching its "
    "labels to the numbered points and the flowchart so you can both picture and describe it.",
    "Revise by first reading the Answer flow to fix the structure, then the flowchart and "
    "mnemonic to lock the sequence and key points, and finally the full sentences for exam wording.",
]

# ---------------------------------------------------------------------------
# Inline labelled SVG diagrams (fully offline, no data URI needed).
# Returned as literal <svg> markup embedded directly in the page.
# ---------------------------------------------------------------------------
SVG_DSM_ICD_TIMELINE = """
<svg viewBox="0 0 900 320" role="img" aria-label="Parallel timelines of DSM and ICD editions"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="900" height="320" fill="#ffffff"/>
  <text x="450" y="26" text-anchor="middle" font-weight="bold" font-size="14" fill="#0f2744">
    Two parallel classifications: APA's DSM (top) and WHO's ICD (bottom)</text>
  <line x1="40" y1="110" x2="860" y2="110" stroke="#1d5fa6" stroke-width="4"/>
  <polygon points="860,110 846,102 846,118" fill="#1d5fa6"/>
  <line x1="40" y1="240" x2="860" y2="240" stroke="#0e7c6b" stroke-width="4"/>
  <polygon points="860,240 846,232 846,248" fill="#0e7c6b"/>
  <g font-size="12" fill="#10243f">
    <text x="20" y="92" font-weight="bold" fill="#1d5fa6">DSM (APA)</text>
    <circle cx="110" cy="110" r="6" fill="#1d5fa6"/>
    <text x="110" y="82" text-anchor="middle" font-weight="bold">DSM-I</text>
    <text x="110" y="140" text-anchor="middle">1952</text>
    <circle cx="260" cy="110" r="6" fill="#1d5fa6"/>
    <text x="260" y="82" text-anchor="middle" font-weight="bold">DSM-II</text>
    <text x="260" y="140" text-anchor="middle">1968</text>
    <circle cx="420" cy="110" r="6" fill="#1d5fa6"/>
    <text x="420" y="82" text-anchor="middle" font-weight="bold">DSM-III</text>
    <text x="420" y="140" text-anchor="middle">1980 (criteria)</text>
    <circle cx="600" cy="110" r="6" fill="#1d5fa6"/>
    <text x="600" y="82" text-anchor="middle" font-weight="bold">DSM-IV</text>
    <text x="600" y="140" text-anchor="middle">1994</text>
    <circle cx="790" cy="110" r="6" fill="#1d5fa6"/>
    <text x="790" y="82" text-anchor="middle" font-weight="bold">DSM-5</text>
    <text x="790" y="140" text-anchor="middle">2013 (no axes)</text>

    <text x="20" y="222" font-weight="bold" fill="#0e7c6b">ICD (WHO)</text>
    <circle cx="110" cy="240" r="6" fill="#0e7c6b"/>
    <text x="110" y="212" text-anchor="middle" font-weight="bold">ICD-7</text>
    <text x="110" y="270" text-anchor="middle">1955</text>
    <circle cx="260" cy="240" r="6" fill="#0e7c6b"/>
    <text x="260" y="212" text-anchor="middle" font-weight="bold">ICD-8</text>
    <text x="260" y="270" text-anchor="middle">1965</text>
    <circle cx="420" cy="240" r="6" fill="#0e7c6b"/>
    <text x="420" y="212" text-anchor="middle" font-weight="bold">ICD-9</text>
    <text x="420" y="270" text-anchor="middle">1978</text>
    <circle cx="600" cy="240" r="6" fill="#0e7c6b"/>
    <text x="600" y="212" text-anchor="middle" font-weight="bold">ICD-10</text>
    <text x="600" y="270" text-anchor="middle">1992</text>
    <circle cx="790" cy="240" r="6" fill="#0e7c6b"/>
    <text x="790" y="212" text-anchor="middle" font-weight="bold">ICD-11</text>
    <text x="790" y="270" text-anchor="middle">2022 (Ch. 6)</text>
  </g>
  <text x="450" y="306" text-anchor="middle" font-size="12" fill="#40506a" font-style="italic">
    Both are syndrome-based catalogues; DSM is operational and research-led, ICD is descriptive and clinical.</text>
</svg>
"""

SVG_PERCEPTION_TREE = """
<svg viewBox="0 0 860 340" role="img" aria-label="Tree of perceptual disorders: distortions, illusions and hallucinations"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="860" height="340" fill="#ffffff"/>
  <g font-size="12" fill="#10243f">
    <rect x="330" y="20" width="200" height="48" rx="10" fill="#f0eafa" stroke="#4a2080"/>
    <text x="430" y="40" text-anchor="middle" font-weight="bold" fill="#4a2080">Disorders of perception</text>
    <text x="430" y="58" text-anchor="middle">a false or altered percept</text>

    <rect x="40" y="130" width="220" height="60" rx="10" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="150" y="152" text-anchor="middle" font-weight="bold">Sensory distortions</text>
    <text x="150" y="170" text-anchor="middle">a real object, wrongly</text>
    <text x="150" y="184" text-anchor="middle">perceived (intensity/quality)</text>

    <rect x="320" y="130" width="220" height="60" rx="10" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="430" y="152" text-anchor="middle" font-weight="bold">Illusions</text>
    <text x="430" y="170" text-anchor="middle">a real object is</text>
    <text x="430" y="184" text-anchor="middle">misinterpreted</text>

    <rect x="600" y="130" width="220" height="60" rx="10" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="710" y="152" text-anchor="middle" font-weight="bold">Hallucinations</text>
    <text x="710" y="170" text-anchor="middle">perception without</text>
    <text x="710" y="184" text-anchor="middle">an object</text>

    <rect x="560" y="250" width="130" height="56" rx="8" fill="#fff" stroke="#8a5c00"/>
    <text x="625" y="272" text-anchor="middle" font-weight="bold">True</text>
    <text x="625" y="290" text-anchor="middle">outer space, real</text>

    <rect x="710" y="250" width="130" height="56" rx="8" fill="#fff" stroke="#8a5c00"/>
    <text x="775" y="268" text-anchor="middle" font-weight="bold">Pseudo-</text>
    <text x="775" y="284" text-anchor="middle">inner space,</text>
    <text x="775" y="298" text-anchor="middle">insight kept</text>
  </g>
  <g stroke="#0f2744" stroke-width="2" fill="none" marker-end="url(#arP)">
    <defs><marker id="arP" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto">
      <path d="M0,0 L7,3 L0,6 Z" fill="#0f2744"/></marker></defs>
    <path d="M380,68 L170,128"/>
    <path d="M430,68 L430,128"/>
    <path d="M480,68 L700,128"/>
    <path d="M680,190 L640,248"/>
    <path d="M740,190 L772,248"/>
  </g>
</svg>
"""

SVG_SELF_ASPECTS = """
<svg viewBox="0 0 860 300" role="img" aria-label="Jaspers' four formal aspects of the experience of self"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="860" height="300" fill="#ffffff"/>
  <circle cx="430" cy="150" r="60" fill="#0f2744"/>
  <text x="430" y="146" text-anchor="middle" font-weight="bold" fill="#fff" font-size="14"
        font-family="Arial,Helvetica,sans-serif">The</text>
  <text x="430" y="164" text-anchor="middle" font-weight="bold" fill="#fff" font-size="14"
        font-family="Arial,Helvetica,sans-serif">self (I)</text>
  <g font-size="12" fill="#10243f">
    <rect x="60" y="40" width="210" height="64" rx="10" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="165" y="62" text-anchor="middle" font-weight="bold">1. Activity / existence</text>
    <text x="165" y="80" text-anchor="middle">"I am doing / I exist"</text>
    <text x="165" y="96" text-anchor="middle">lost in depersonalisation</text>

    <rect x="590" y="40" width="210" height="64" rx="10" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="695" y="62" text-anchor="middle" font-weight="bold">2. Unity</text>
    <text x="695" y="80" text-anchor="middle">"I am one at this moment"</text>
    <text x="695" y="96" text-anchor="middle">lost in "double self"</text>

    <rect x="60" y="196" width="210" height="64" rx="10" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="165" y="218" text-anchor="middle" font-weight="bold">3. Continuity</text>
    <text x="165" y="236" text-anchor="middle">"same person over time"</text>
    <text x="165" y="252" text-anchor="middle">lost after "break" in life</text>

    <rect x="590" y="196" width="210" height="64" rx="10" fill="#f0eafa" stroke="#4a2080"/>
    <text x="695" y="218" text-anchor="middle" font-weight="bold">4. Boundaries</text>
    <text x="695" y="236" text-anchor="middle">self vs the outer world</text>
    <text x="695" y="252" text-anchor="middle">lost in passivity phenomena</text>
  </g>
  <g stroke="#5a6a7e" stroke-width="2">
    <line x1="270" y1="72" x2="375" y2="128"/>
    <line x1="590" y1="72" x2="485" y2="128"/>
    <line x1="270" y1="228" x2="375" y2="172"/>
    <line x1="590" y1="228" x2="485" y2="172"/>
  </g>
</svg>
"""

SVG_THOUGHT_MAP = """
<svg viewBox="0 0 860 300" role="img" aria-label="Map of disorders of thought by stream, possession, content and form"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="860" height="300" fill="#ffffff"/>
  <rect x="340" y="18" width="180" height="46" rx="10" fill="#0f2744"/>
  <text x="430" y="38" text-anchor="middle" font-weight="bold" fill="#fff">Disorders of thought</text>
  <text x="430" y="55" text-anchor="middle" fill="#fff" font-size="11">examined four ways</text>
  <g font-size="12" fill="#10243f">
    <rect x="20" y="130" width="185" height="90" rx="10" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="112" y="152" text-anchor="middle" font-weight="bold">Stream (tempo)</text>
    <text x="112" y="172" text-anchor="middle">flight of ideas,</text>
    <text x="112" y="188" text-anchor="middle">inhibition, blocking,</text>
    <text x="112" y="204" text-anchor="middle">circumstantiality</text>

    <rect x="233" y="130" width="185" height="90" rx="10" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="325" y="152" text-anchor="middle" font-weight="bold">Possession</text>
    <text x="325" y="172" text-anchor="middle">obsessions,</text>
    <text x="325" y="188" text-anchor="middle">compulsions,</text>
    <text x="325" y="204" text-anchor="middle">thought alienation</text>

    <rect x="446" y="130" width="185" height="90" rx="10" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="538" y="152" text-anchor="middle" font-weight="bold">Content</text>
    <text x="538" y="172" text-anchor="middle">delusions</text>
    <text x="538" y="188" text-anchor="middle">(primary vs</text>
    <text x="538" y="204" text-anchor="middle">secondary), themes</text>

    <rect x="659" y="130" width="185" height="90" rx="10" fill="#f0eafa" stroke="#4a2080"/>
    <text x="751" y="152" text-anchor="middle" font-weight="bold">Form</text>
    <text x="751" y="172" text-anchor="middle">formal thought</text>
    <text x="751" y="188" text-anchor="middle">disorder, loosening,</text>
    <text x="751" y="204" text-anchor="middle">knight's move</text>
  </g>
  <g stroke="#0f2744" stroke-width="2" fill="none">
    <line x1="400" y1="64" x2="112" y2="128"/>
    <line x1="415" y1="64" x2="325" y2="128"/>
    <line x1="445" y1="64" x2="538" y2="128"/>
    <line x1="460" y1="64" x2="751" y2="128"/>
  </g>
</svg>
"""

SVG_MEMORY_CHAIN = """
<svg viewBox="0 0 880 210" role="img" aria-label="The memory chain: registration, retention, retrieval, recall"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="880" height="210" fill="#ffffff"/>
  <g font-size="12" fill="#10243f">
    <rect x="20" y="70" width="180" height="66" rx="10" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="110" y="96" text-anchor="middle" font-weight="bold">Registration</text>
    <text x="110" y="114" text-anchor="middle">taking information in</text>
    <text x="110" y="128" text-anchor="middle">(attention needed)</text>

    <rect x="240" y="70" width="180" height="66" rx="10" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="330" y="96" text-anchor="middle" font-weight="bold">Retention</text>
    <text x="330" y="114" text-anchor="middle">storing it over time</text>

    <rect x="460" y="70" width="180" height="66" rx="10" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="550" y="96" text-anchor="middle" font-weight="bold">Retrieval</text>
    <text x="550" y="114" text-anchor="middle">finding it again</text>

    <rect x="680" y="70" width="180" height="66" rx="10" fill="#0f2744" stroke="#0f2744"/>
    <text x="770" y="96" text-anchor="middle" font-weight="bold" fill="#fff">Recall</text>
    <text x="770" y="114" text-anchor="middle" fill="#fff">bringing it to mind</text>
  </g>
  <g stroke="#0f2744" stroke-width="2" fill="none" marker-end="url(#arM)">
    <defs><marker id="arM" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto">
      <path d="M0,0 L7,3 L0,6 Z" fill="#0f2744"/></marker></defs>
    <path d="M200,103 L238,103"/>
    <path d="M420,103 L458,103"/>
    <path d="M640,103 L678,103"/>
  </g>
  <text x="440" y="175" text-anchor="middle" font-size="12" fill="#40506a" font-style="italic">
    A failure at any stage shows as amnesia; the chain is the frame for testing memory at the bedside.</text>
</svg>
"""

SVG_EMOTION_LEVELS = """
<svg viewBox="0 0 860 300" role="img" aria-label="Five levels of disorders of emotion from normal to morbid"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="860" height="300" fill="#ffffff"/>
  <text x="430" y="26" text-anchor="middle" font-weight="bold" font-size="14" fill="#0f2744">
    Emotion graded along five levels (normal &#8594; morbid)</text>
  <g font-size="12" fill="#10243f">
    <rect x="60" y="56" width="740" height="34" rx="6" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="78" y="78" font-weight="bold">1. Normal emotions</text>
    <text x="430" y="78" text-anchor="middle">appropriate joy, fear, sadness in daily life</text>
    <rect x="60" y="100" width="740" height="34" rx="6" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="78" y="122" font-weight="bold">2. Abnormal emotional reactions</text>
    <text x="430" y="122" text-anchor="middle">too strong or too long for the trigger</text>
    <rect x="60" y="144" width="740" height="34" rx="6" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="78" y="166" font-weight="bold">3. Abnormal expressions of emotion</text>
    <text x="430" y="166" text-anchor="middle">incongruity, blunting, lability, perplexity</text>
    <rect x="60" y="188" width="740" height="34" rx="6" fill="#f0eafa" stroke="#4a2080"/>
    <text x="78" y="210" font-weight="bold">4. Morbid expressions</text>
    <text x="430" y="210" text-anchor="middle">apathy, flat affect, loss of feeling</text>
    <rect x="60" y="232" width="740" height="34" rx="6" fill="#7a1f3d" stroke="#7a1f3d"/>
    <text x="78" y="254" font-weight="bold" fill="#fff">5. Morbid disorders of emotion</text>
    <text x="430" y="254" text-anchor="middle" fill="#fff">depression, mania, pathological anxiety, ecstasy</text>
  </g>
</svg>
"""

# ---------------------------------------------------------------------------
# CONTENT = single source of truth (12 topics, in the required order)
# ---------------------------------------------------------------------------
CONTENT = [
    {
        "anchor": "nomenclature",
        "number": 1,
        "color": "blue",
        "title": "Nomenclature and the Classificatory Systems in Current Use",
        "intro": "Before any symptom can be named, psychiatry needs an agreed vocabulary and a shared "
                 "catalogue of disorders, so this topic sets out what classification is, traces the "
                 "parallel histories of the DSM and the ICD, and shows how the two systems resemble and "
                 "differ from each other.",
        "answer_flow": [
            "Definition of nomenclature & classification",
            "Background / why we classify",
            "DSM history and ICD history",
            "What each is & how they differ",
            "Interview schedules (SCID/CIDI/SCAN)",
            "One-line summary",
        ],
        "points": [
            "**Nomenclature** [[the agreed set of names for disorders]] is the shared vocabulary of "
            "psychiatry, while **classification** is the orderly grouping of those named disorders into "
            "categories, so that clinicians everywhere can mean the same thing by a diagnosis.",
            "Classification exists chiefly for **communication**, letting clinicians and researchers "
            "talk precisely, pool data on cause and outcome, and decide treatment, so a good system is "
            "judged by how useful and reliable it proves.",
            "The two systems in current use are the **DSM** [[the Diagnostic and Statistical Manual]] of "
            "the American Psychiatric Association and the **ICD** [[the International Classification of "
            "Diseases]] of the World Health Organization, and both are built around clinical syndromes.",
            "The DSM lineage runs from **DSM-I in 1952** and DSM-II in 1968, which were brief and "
            "psychoanalytic, to the landmark **DSM-III in 1980**, which introduced explicit operational "
            "criteria and set the template every later edition has followed.",
            "**DSM-IV** appeared in 1994 and the current **DSM-5** in 2013, a revision that notably "
            "**dropped the five-axis system** [[the old habit of coding illness, personality, medical, "
            "social and functioning problems on separate lines]] used since DSM-III.",
            "The ICD lineage reaches modern psychiatry with **ICD-7 in 1955**, then ICD-8 in 1965 and "
            "ICD-9 in 1978, each slowly widening and sharpening the mental-disorders section as the WHO "
            "sought a single language for health statistics worldwide.",
            "**ICD-10** in 1992 gave mental and behavioural disorders their own detailed chapter with "
            "clinical descriptions and diagnostic guidelines, and the current **ICD-11**, adopted in "
            "2022, places them in **Chapter 6** with a fully updated, dimensional-leaning structure.",
            "A key difference is purpose: the DSM is an **operationalised research manual** for a largely "
            "specialist readership, whereas the ICD is a broader **descriptive** system for public-health "
            "and general clinical use worldwide.",
            "They also differ in scope, since the ICD is a whole-body disease classification with mental "
            "disorders as just one chapter, while the DSM deals **solely with mental disorders** and so "
            "carries more detailed criteria for each.",
            "In style the DSM tends to favour **tight operational criteria** with counts and thresholds, "
            "while ICD-10 offered looser **clinical descriptions and guidelines**, a difference that can "
            "make the same patient meet one system's threshold but not the other's.",
            "Despite these contrasts both systems are **syndrome-based** [[they group symptoms that "
            "cluster together rather than grouping by proven cause]], because psychiatry still lacks the "
            "biological markers that would let it classify most disorders by underlying pathology.",
            "The two bodies now work for **harmonisation**, so that DSM-5 and ICD-11 categories and codes "
            "line up far better than earlier editions did, reducing the confusion of a patient being "
            "labelled differently on either side of an ocean.",
            "To apply any system reliably clinicians use **structured interview schedules**, chief among "
            "them the **SCID** [[the Structured Clinical Interview for DSM, run by a clinician]], which "
            "walks through DSM criteria disorder by disorder to yield a standardised diagnosis.",
            "The **CIDI** [[the Composite International Diagnostic Interview]] is a fully structured "
            "schedule designed for trained lay interviewers and large epidemiological surveys, giving "
            "ICD and DSM diagnoses without needing clinical judgement at each question.",
            "The **SCAN** [[the Schedules for Clinical Assessment in Neuropsychiatry, built around the "
            "Present State Examination]] is a semi-structured instrument that relies on the clinician's "
            "trained judgement to rate the presence and severity of each symptom.",
            "These schedules matter because they raise **reliability** [[the chance that two clinicians "
            "reach the same diagnosis]], turning a classification on paper into something that can be "
            "applied consistently in research and audited in everyday clinical work.",
            "In one line, nomenclature gives psychiatry its shared names and the DSM and ICD its two "
            "great syndrome-based catalogues, one research-led and specialist, the other descriptive and "
            "global, applied reliably through structured interviews.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Classification in psychiatry",
            "text": "A classification in psychiatry is a system that groups mental disorders into "
                    "categories on the basis of shared clinical features (a syndrome), its chief "
                    "purposes being reliable communication between clinicians, the gathering of "
                    "comparable statistics, and the guidance of prognosis and treatment, as set out in "
                    "the opening chapters of Fish's Clinical Psychopathology.",
            "cite": "Paraphrased from Casey, P. & Kelly, B., Fish's Clinical Psychopathology, 5th ed. "
                    "(introductory chapter on classification and nomenclature).",
        },
        "svg": SVG_DSM_ICD_TIMELINE,
        "svg_caption": ("Figure 1. Parallel timelines of the two classifications \u2014 the APA's DSM "
                        "(DSM-I 1952 to DSM-5 2013, which dropped the five axes) and the WHO's ICD "
                        "(ICD-7 1955 to ICD-11 2022, Chapter 6), both syndrome-based."),
        "figures": [],
        "flow": {
            "title": "From symptoms to a shared diagnosis",
            "steps": [
                ("Observe symptoms", "the raw clinical data", False),
                ("Group as a syndrome", "symptoms that cluster", False),
                ("Apply DSM or ICD", "named category + code", False),
                ("Use a schedule", "SCID / CIDI / SCAN", False),
                ("Reliable diagnosis", "shared, comparable", True),
            ],
        },
        "mnemonic": {
            "word": "CLASSIFY",
            "line": "CLASSIFY reminds you why and how psychiatry orders its disorders.",
            "items": [
                ("C", "**Communication** is the first purpose of any system."),
                ("L", "**Lineages**: DSM from 1952, ICD-7 from 1955."),
                ("A", "**APA owns DSM**, WHO owns the ICD."),
                ("S", "**Syndrome-based**: both group clustering symptoms."),
                ("S", "**Statistics** and prognosis flow from good classes."),
                ("I", "**ICD-11 (2022) Chapter 6** for mental disorders."),
                ("F", "**Five axes dropped** by DSM-5 in 2013."),
                ("Y", "**Yardsticks**: SCID, CIDI and SCAN apply them reliably."),
            ],
        },
    },
    {
        "anchor": "limitations",
        "number": 2,
        "color": "teal",
        "title": "Limitations of Classification",
        "intro": "No classification is a perfect mirror of nature, so this topic examines the "
                 "philosophical and practical limits of diagnosis, from Jaspers' caution and the realism "
                 "debate, through validity, reliability and the clinical-significance criterion, to the "
                 "stubborn problem of comorbidity.",
        "answer_flow": [
            "Definition of the problem",
            "Philosophical status (realism debate)",
            "Validity vs reliability",
            "Criterion & threshold problems",
            "Comorbidity & category limits",
            "One-line summary",
        ],
        "points": [
            "Classification is useful but **provisional**, because psychiatric categories are working "
            "tools rather than proven natural kinds, so every system carries limits that a thoughtful "
            "clinician must weigh rather than treating a diagnosis as a fixed fact.",
            "**Karl Jaspers** warned that diagnostic categories are **ideal types** [[useful mental "
            "models that real patients only ever approximate]], so no actual person fits a textbook "
            "label exactly, and forcing a messy case into a neat box can distort the clinical picture.",
            "Kendler's 2016 analysis frames three philosophical stances: **realism** [[disorders are "
            "real natural entities waiting to be discovered]], **constructivism** [[categories are "
            "invented by committees]], and **pragmatism**, the middle view that categories are useful "
            "tools judged by what they achieve.",
            "A core limit is the gap between **validity** [[whether a category marks a genuinely "
            "distinct, real disorder]] and **reliability** [[whether clinicians agree on applying it]], "
            "because a label can be highly reliable yet still not carve nature at a true joint.",
            "Operational criteria greatly improved reliability after DSM-III, but critics note this can "
            "buy agreement at the cost of validity, pinning down a definition precisely without proving "
            "that the thing defined is a **real and separate illness**.",
            "The **clinical-significance criterion** [[the rule that symptoms must cause distress or "
            "impairment to count as a disorder]] tries to keep normal reactions out, yet Allen Frances "
            "called it the **seeds of tautology** because it can be circular.",
            "That circularity arises because distress or impairment is used both to define the disorder "
            "and to prove it matters, so the criterion risks simply restating that a problem is a "
            "problem rather than adding genuinely new information.",
            "The systems even word this threshold differently: the DSM typically requires distress "
            "**OR** impairment, whereas ICD-10 often leant toward requiring **BOTH**, so the same "
            "person can cross one manual's line into caseness while staying below the other's.",
            "Categories may be **monothetic** [[every listed feature is required]] or **polythetic** "
            "[[any sufficient subset of features will do]], and modern manuals are largely polythetic, "
            "which means two people with the same diagnosis can share very few symptoms.",
            "This polythetic design produces great **within-category heterogeneity**, so a single label "
            "like depression covers patients whose presentations overlap only partly, weakening the "
            "assumption that one diagnosis implies one cause, course or best treatment.",
            "A major limitation is **comorbidity** [[two or more disorders diagnosed in the same person "
            "at once]], which is so common that Tyrer's 2017 work argues it often reflects artificial "
            "splitting of what is really one broad underlying problem.",
            "Tyrer distinguishes true **co-occurrence** from mere **consanguinity** [[categories that "
            "look separate but are really close relatives carved from the same territory]], so high "
            "comorbidity rates may be an artefact of how the boundaries were drawn.",
            "Field data expose the fragility of boundaries, with reported **ICD-10 versus DSM-IV "
            "concordance** [[the rate at which the two systems agree on a case]] ranging widely, from "
            "about **33% to 87%** depending on the disorder studied.",
            "Such variable agreement shows that for some conditions the two manuals are nearly "
            "interchangeable while for others they disagree sharply, undermining any claim that either "
            "system has simply found the one correct way to divide mental illness.",
            "Classification can also create **stigma** and self-fulfilling expectations, since a label "
            "shapes how others treat a person and how the person sees themselves, a social cost that "
            "sits alongside the scientific limits of validity and reliability.",
            "Despite all this, the pragmatic defence holds that an imperfect, agreed classification is "
            "**better than none**, because without shared categories research could not accumulate and "
            "clinicians could not communicate, so the goal is steady improvement, not abandonment.",
            "In one line, classification is limited by shaky validity, circular thresholds, heterogeneous "
            "polythetic categories and pervasive comorbidity, so diagnoses are best treated as useful, "
            "revisable tools rather than discovered natural facts.",
        ],
        "definition": None,
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "Why a diagnosis may mislead",
            "steps": [
                ("Ideal type", "patients only approximate it", False),
                ("Reliable label", "clinicians agree", False),
                ("But is it valid?", "a real, separate illness?", False),
                ("Polythetic + comorbid", "fuzzy, overlapping", False),
                ("Treat as a tool", "useful, revisable", True),
            ],
        },
        "mnemonic": {
            "word": "LIMITS",
            "line": "LIMITS flags the weak points of any diagnostic system.",
            "items": [
                ("L", "**Labels** are ideal types (Jaspers), not exact fits."),
                ("I", "**Invalidity** can hide behind high reliability."),
                ("M", "**Monothetic vs polythetic** creates heterogeneity."),
                ("I", "**ICD vs DSM** concordance ranges ~33-87%."),
                ("T", "**Tautology**: the clinical-significance 'seeds of tautology'."),
                ("S", "**Splitting** drives comorbidity (Tyrer 2017)."),
            ],
        },
    },
    {
        "anchor": "consciousness",
        "number": 3,
        "color": "amber",
        "title": "Disorders of Consciousness",
        "intro": "Consciousness is the stage on which all other mental functions appear, so this topic "
                 "defines awareness of self and environment, sets out the states in which it is clouded, "
                 "lowered, narrowed or dissociated, and explains why such change is usually a red flag "
                 "for organic illness.",
        "answer_flow": [
            "Definition of consciousness",
            "Background (a continuum of arousal)",
            "Types of disordered consciousness",
            "Clinical examples (delirium, fugue)",
            "Significance (organic until proven)",
            "One-line summary",
        ],
        "points": [
            "**Consciousness** is best defined as **awareness of the self and of the environment**, a "
            "background state of alertness that makes ordinary experience possible, so when it is "
            "disturbed every other mental function is inevitably affected too.",
            "Consciousness is not all-or-none but a **continuum** running from full alertness, through "
            "drowsiness and clouding, down to stupor and finally coma, so clinicians describe where on "
            "this gradient a patient sits rather than using a single on-off label.",
            "**Clouding of consciousness** [[a mild dulling where awareness and attention are slightly "
            "reduced]] is the gentlest change, leaving the patient muddled, slow and poorly focused, "
            "and is the earliest warning that the brain's overall function is failing.",
            "A **dream-like (oneiroid) state** blends vivid inner imagery with the outer world, and is "
            "the quality typical of **delirium** [[an acute, fluctuating confusional state]], in which "
            "the patient is disoriented, frightened and often troubled by vivid hallucinations.",
            "A **lowering of consciousness (torpor)** [[reduced wakefulness with slow, sluggish "
            "thinking]] leaves the patient drowsy, easily losing the thread, drifting toward sleep, and "
            "able to give only poor and muddled answers to questions.",
            "A **restriction or narrowing of consciousness (twilight state)** keeps a small band of "
            "awareness intact while the rest is shut out, so the person may carry out semi-purposeful "
            "acts yet have no proper memory of them afterwards.",
            "A **fugue** [[a sudden wandering away from home with loss of personal memory]] is a "
            "dissociative narrowing in which a person may travel far, sometimes taking on a new "
            "identity, and later cannot recall the episode at all.",
            "These states are distinguished partly by **mood and content**: the twilight state often "
            "carries a tense, irritable mood, while the oneiroid state is dominated by dream-like "
            "imagery, helping the clinician tell one narrowing from another.",
            "**Delirium** deserves special emphasis as the commonest disorder of consciousness in "
            "general hospitals, marked by a fluctuating course, worse confusion at night, altered "
            "sleep-wake cycle and frequent visual hallucinations and misperceptions.",
            "A cardinal teaching point is that disturbed consciousness is **prima facie organic** "
            "[[on the face of it caused by physical brain disease]], so any clouding should trigger a "
            "search for infection, drugs, metabolic upset or other bodily causes.",
            "This organic rule matters because missing a delirium can be fatal, so a new confusional "
            "state is treated as a **medical emergency** until a physical cause is found or firmly "
            "excluded, not dismissed as a purely psychological reaction.",
            "By contrast, the **dissociative** narrowings such as fugue can occur with a clear sensorium "
            "and are psychogenic, which is why distinguishing an organic clouding from a dissociative "
            "narrowing is one of the first judgements the clinician must make.",
            "Assessment relies on **simple bedside tests** of orientation, attention and the ability to "
            "sustain a conversation, because a patient who cannot hold attention or stay oriented is "
            "showing, by definition, a disorder of consciousness.",
            "The level of consciousness typically **fluctuates** in organic states, so repeated short "
            "examinations across the day are more revealing than a single assessment, and night-time "
            "worsening is a classic clue to an underlying delirium.",
            "Severe lowering ends in **stupor**, a state of near-unresponsiveness with preserved "
            "posture and some reaction to strong stimuli, and then **coma**, the complete and "
            "unrousable loss of awareness at the far end of the continuum.",
            "Recovery usually leaves a **patchy or absent memory** for the period of altered "
            "consciousness, and this amnesic gap is itself diagnostic evidence that consciousness, not "
            "merely mood or thought, was the function disturbed.",
            "In one line, disorders of consciousness are disturbances of awareness ranging from clouding "
            "to coma and including dream-like, lowered, narrowed and fugue states, and they should be "
            "read as organic until proven otherwise.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Consciousness",
            "text": "Consciousness is described in Fish as a state of awareness of the self and of the "
                    "environment; its disorders are graded along a continuum (clouding, drowsiness, "
                    "stupor, coma) and qualitative changes such as dream-like, lowered, restricted and "
                    "dissociative states, a disturbance of which is regarded as prima facie evidence of "
                    "organic cerebral disorder.",
            "cite": "Paraphrased from Casey, P. & Kelly, B., Fish's Clinical Psychopathology, 5th ed. "
                    "(chapter on disorders of consciousness and attention).",
        },
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "The continuum of consciousness",
            "steps": [
                ("Alert", "full awareness", False),
                ("Clouding", "mild dulling", False),
                ("Torpor / twilight", "lowered / narrowed", False),
                ("Stupor", "near-unresponsive", False),
                ("Coma", "unrousable", True),
            ],
        },
        "mnemonic": {
            "word": "CLOUDS",
            "line": "CLOUDS gathers the ways consciousness can be disturbed.",
            "items": [
                ("C", "**Clouding**: mild dulling of awareness."),
                ("L", "**Lowering** (torpor): drowsy, sluggish thinking."),
                ("O", "**Oneiroid / delirium**: dream-like, fluctuating."),
                ("U", "**Unrousable** end-point is coma."),
                ("D", "**Dissociative** fugue: wandering, memory lost."),
                ("S", "**Suspect organic** cause first (prima facie)."),
            ],
        },
    },
    {
        "anchor": "attention",
        "number": 4,
        "color": "rose",
        "title": "Disorders of Attention and Concentration",
        "intro": "Attention is the spotlight that selects what we process, so this topic separates active "
                 "from passive attention, surveys how fatigue, anxiety, mood, psychosis and organic "
                 "disease each disturb it, and shows why attentional change is both common and "
                 "diagnostically informative.",
        "answer_flow": [
            "Definition of attention & concentration",
            "Active vs passive attention",
            "Causes of distractibility",
            "Clinical patterns by disorder",
            "Significance for assessment",
            "One-line summary",
        ],
        "points": [
            "**Attention** is the active direction of awareness onto a chosen object, while "
            "**concentration** is the effort of holding that focus over time, so together they form the "
            "gateway through which information must pass to be remembered.",
            "A useful split is between **active attention** [[deliberately choosing what to attend to]] "
            "and **passive attention** [[being automatically drawn by a sudden or striking stimulus]], "
            "and many disorders tip the balance between these two kinds.",
            "**Distractibility** [[the tendency for attention to be pulled away by any passing "
            "stimulus]] is the most frequent attentional symptom, appearing whenever active attention "
            "weakens and passive attention takes over the direction of the mind.",
            "Ordinary **fatigue** [[tiredness]] and boredom reduce attention in everyone, so a tired or "
            "under-stimulated person attends poorly, which means the clinician must gauge whether poor "
            "attention is situational or a genuine sign of illness.",
            "In **anxiety** the mind is repeatedly dragged back to the feared theme, so attention "
            "narrows onto threat and concentration on neutral tasks suffers, a pattern patients often "
            "describe as being unable to keep their mind on anything.",
            "In **depression** attention and concentration are slowed and effortful, so patients "
            "complain they cannot read a page or follow a conversation, and this subjective poor "
            "concentration is one of the recognised diagnostic features of the illness.",
            "In **mania** attention becomes highly **distractible**, flitting from one external stimulus "
            "to the next, so the patient cannot sustain any single task and their talk jumps in step "
            "with whatever happens to catch their eye.",
            "In **schizophrenia** attention is often disorganised, with patients unable to filter out "
            "irrelevant stimuli, so they feel flooded by input and their concentration fragments, "
            "mirroring the broader disturbance of thought and perception.",
            "In **organic states** such as delirium and dementia, impaired attention is a core feature, "
            "and indeed the inability to sustain and shift attention is one of the defining signs that "
            "separates delirium from a purely functional illness.",
            "A special pattern is the **amnestic patient's rigid set**, where severe memory loss leaves "
            "the person clinging to a narrow fixed focus, unable to redirect attention because they "
            "cannot update what they hold in mind.",
            "Attention and consciousness are intimately linked, because a patient whose consciousness is "
            "clouded cannot by definition attend well, so attentional testing doubles as a sensitive "
            "**screen for early clouding** of consciousness.",
            "Bedside tests of attention include **serial sevens** [[subtracting seven repeatedly from "
            "one hundred]], reciting the months of the year backwards, or the digit span, each "
            "demanding sustained active attention that organic and severe functional illness disrupts.",
            "It is important to separate **subjective** poor concentration, which the patient reports, "
            "from **objective** inattention that the examiner can demonstrate on testing, because the "
            "two do not always agree and each carries different weight.",
            "Attentional symptoms are highly **non-specific**, occurring across almost every psychiatric "
            "and many physical illnesses, so poor attention alone never fixes a diagnosis and must "
            "always be read in the context of the whole clinical picture.",
            "Because attention underpins memory **registration**, a patient who cannot attend will "
            "appear to have a memory problem when the true fault lies one step earlier, a distinction "
            "the clinician must keep clearly in mind when testing.",
            "Culturally and educationally fair testing matters, since tasks like serial sevens assume "
            "numeracy, so the examiner chooses an attention test suited to the patient to avoid "
            "mistaking lack of schooling for genuine inattention.",
            "In one line, disorders of attention range from everyday distractibility in fatigue and "
            "anxiety to the gross inattention of delirium, and because they are common but non-specific "
            "they guide rather than make a diagnosis.",
        ],
        "definition": None,
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "How attention breaks down",
            "steps": [
                ("Active attention", "chosen focus", False),
                ("Load: fatigue / anxiety", "focus weakens", False),
                ("Distractibility", "passive pull wins", False),
                ("Disorder pattern", "mania, depression, organic", False),
                ("Read in context", "non-specific sign", True),
            ],
        },
        "mnemonic": {
            "word": "FOCUS",
            "line": "FOCUS lists what disturbs attention and why it matters.",
            "items": [
                ("F", "**Fatigue** and boredom lower attention in all."),
                ("O", "**Organic** states: core inattention (delirium)."),
                ("C", "**Concentration** slowed in depression."),
                ("U", "**Unfiltered** flooding in schizophrenia."),
                ("S", "**Switching** wildly: distractibility of mania."),
            ],
        },
    },
    {
        "anchor": "orientation",
        "number": 5,
        "color": "purple",
        "title": "Disorders of Orientation",
        "intro": "Orientation anchors a person in time, place and identity, so this topic defines it, "
                 "explains the characteristic order in which it is lost, distinguishes disorientation "
                 "from confusion, and shows why disordered orientation is one of psychiatry's clearest "
                 "pointers to organic disease.",
        "answer_flow": [
            "Definition of orientation",
            "The three spheres (time, place, person)",
            "Order of loss",
            "Clinical causes (organic)",
            "Confusion vs disorientation",
            "One-line summary",
        ],
        "points": [
            "**Orientation** is a person's awareness of their situation in **time, place and person**, "
            "so an oriented patient knows the date, knows where they are, and knows who they and the "
            "people around them are, grounding them in reality.",
            "Orientation depends on intact **consciousness, attention and memory**, so it fails when any "
            "of these underlying functions fails, which is why testing orientation is such an efficient "
            "screen for a disturbed brain.",
            "**Orientation to time** is usually the **first** sphere to be lost, so a patient who "
            "becomes vague about the day, date, month or even the year is showing the earliest and most "
            "sensitive sign of developing disorientation.",
            "**Orientation to place** is lost next, so the patient may mistake the hospital for home or "
            "be unable to name their town or ward, a sign of more advanced disturbance than time errors "
            "alone.",
            "**Orientation to person** is the **last** to go and the most robust, so losing awareness of "
            "one's own identity is rare and usually indicates either very severe organic disease or a "
            "dissociative, psychogenic disturbance rather than ordinary confusion.",
            "This stepwise order, **time lost first and person lost last**, is a reliable clinical ruler, "
            "letting the examiner gauge the depth of disturbance from which spheres remain intact and "
            "which have given way.",
            "**Disorientation** [[not knowing the time, place or who people are]] arising acutely is "
            "regarded as **prima facie organic**, so like clouding of consciousness it should prompt a "
            "search for a physical cause before any functional explanation is accepted.",
            "Disorientation is a hallmark of **delirium and dementia**, so its presence strongly tilts "
            "the diagnosis toward an organic disorder and away from conditions such as anxiety or "
            "depression, where orientation is characteristically preserved.",
            "**Confusion** [[muddled, unclear thinking]] often accompanies disorientation but is not "
            "identical to it, because confusion describes disordered thinking in general whereas "
            "disorientation specifically means losing the bearings of time, place or person.",
            "A patient can be transiently **confused** under stress yet remain fully oriented, so the "
            "clinician separates a muddled train of thought from true disorientation, since only the "
            "latter carries the strong organic warning.",
            "Orientation is tested with **simple direct questions** about the day, date, season, place "
            "and the identity of people present, a quick examination that nonetheless yields some of "
            "the most diagnostically weighty information at the bedside.",
            "Errors should be judged **sensibly**, because being a day out on the date in a long "
            "hospital stay is trivial, whereas naming the wrong year or decade is a serious sign of a "
            "significant disturbance of orientation.",
            "In acute organic states orientation typically **fluctuates**, being clearer in the morning "
            "and worse at night, so repeated testing across the day reveals the instability that is "
            "itself characteristic of delirium.",
            "A **false or confabulated orientation** can occur, where an amnesic patient confidently but "
            "wrongly states the date or place, so apparent certainty does not guarantee that the answer "
            "is correct or that orientation is intact.",
            "Because orientation integrates several functions, it serves as a compact **index of global "
            "brain function**, which is why it features in almost every cognitive screening test used "
            "in psychiatry and general medicine.",
            "Disorientation to person in a patient otherwise alert and oriented to time and place is a "
            "**red flag for a dissociative state**, since this reversed pattern breaks the usual organic "
            "order and suggests a psychogenic cause.",
            "In one line, disorders of orientation are losses of awareness of time, place and person, "
            "usually in that order, and their acute appearance is one of psychiatry's clearest signs of "
            "organic cerebral disturbance.",
        ],
        "definition": None,
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "The order of lost orientation",
            "steps": [
                ("Time", "lost first", False),
                ("Place", "lost next", False),
                ("Person", "lost last", False),
                ("Suspect organic", "prima facie", True),
            ],
        },
        "mnemonic": {
            "word": "TPP",
            "line": "TPP fixes the order and meaning of disordered orientation.",
            "items": [
                ("T", "**Time** is lost first (day, date, year)."),
                ("P", "**Place** is lost next (ward, town)."),
                ("P", "**Person** is lost last (own identity)."),
            ],
        },
    },
    {
        "anchor": "motor",
        "number": 6,
        "color": "navy",
        "title": "Disorders of Motor Behavior",
        "intro": "Movement can itself become a symptom, so this topic classifies disordered motor "
                 "behaviour from the patient's and observer's view, details the classic catatonic signs, "
                 "and distinguishes illness-driven movements from those produced by the drugs used to "
                 "treat it.",
        "answer_flow": [
            "Definition of motor disorder",
            "Subjective vs observed; classes of movement",
            "Catatonic phenomena",
            "Stupor and excitement",
            "Drug-induced movement disorders",
            "One-line summary",
        ],
        "points": [
            "**Disorders of motor behaviour** are abnormalities of movement, posture and activity that "
            "arise from mental illness rather than from simple weakness or paralysis, so they are read "
            "as outward signs of an inner psychological or organic disturbance.",
            "Movements can be described from the **subjective** side, how the patient experiences them, "
            "and the **objective** side, what the observer sees, and the two together build the full "
            "clinical picture of a motor disorder.",
            "A key subjective experience is **passivity** [[the feeling that one's movements are made by "
            "an outside force]], where the patient believes their limbs are being moved for them, a "
            "phenomenon strongly associated with schizophrenia.",
            "Movements can be grouped as **adaptive** [[purposeful, goal-directed acts]], **non-adaptive** "
            "[[purposeless movements serving no goal]], and **induced** movements that appear only in "
            "response to the examiner, and this grouping organises the whole examination.",
            "**Catatonia** [[a syndrome of marked motor disturbance]] is the richest source of these "
            "signs, classically occurring in severe psychosis and mood disorder and, importantly, also "
            "arising from organic brain disease, so it always warrants a medical work-up.",
            "**Echopraxia** [[automatically copying the examiner's movements]] and **echolalia** "
            "[[automatically repeating the examiner's words]] are catatonic signs in which the patient "
            "imitates without being asked and seemingly without being able to stop.",
            "**Automatic obedience** is the exaggerated, uncritical carrying out of every instruction, "
            "so the patient complies with even trivial or odd commands, a loss of the normal filtering "
            "of what one chooses to do.",
            "**Negativism** [[motiveless resistance to all instructions, often doing the opposite]] is "
            "the mirror image of automatic obedience, where the patient actively opposes or ignores "
            "what is asked without any apparent reason.",
            "**Ambitendency** [[starting a movement then checking it, caught between doing and not "
            "doing]] shows as a hesitant, to-and-fro motion, as when a patient half-extends a hand to "
            "shake and then withdraws it repeatedly.",
            "**Waxy flexibility (cerea flexibilitas)** [[limbs that can be moulded like wax and held in "
            "position]] lets the examiner place the patient's arm in an awkward posture that is then "
            "maintained for long periods without apparent discomfort.",
            "**Stupor** is a marked reduction or absence of movement and speech with the patient still "
            "conscious, so unlike coma the person remains aware and may later recall events, and it can "
            "be organic, schizophrenic or depressive.",
            "At the opposite pole, **excitement** is a state of excessive, often purposeless overactivity, "
            "ranging from the goal-driven overactivity of mania to the chaotic, bizarre excitement seen "
            "in catatonic and organic states.",
            "**Mannerisms** [[odd but goal-directed movements, like an elaborate way of saluting]] are "
            "distinguished from **stereotypies** [[repetitive, uniform movements with no goal, such as "
            "constant rocking]], a distinction that helps localise the kind of disturbance.",
            "**Tics** [[sudden, brief, repetitive involuntary movements or sounds]] and **compulsive "
            "acts** are further motor phenomena, the former usually neurological and the latter driven "
            "by an inner urge the patient feels compelled to obey.",
            "Many motor signs today are **drug-induced**, because antipsychotic medicines block dopamine "
            "and so produce movement disorders that the clinician must not mistake for the illness "
            "itself when assessing the patient.",
            "The most important of these is **tardive dyskinesia** [[late-appearing, often irreversible "
            "involuntary movements of the face, tongue and limbs]], alongside acute dystonia, "
            "parkinsonism and the restless legs of **akathisia** caused by antipsychotic treatment.",
            "In one line, disorders of motor behaviour span passivity, the catatonic signs, stupor and "
            "excitement, and the movement disorders induced by antipsychotic drugs, each requiring the "
            "clinician to separate illness from treatment and functional from organic causes.",
        ],
        "definition": None,
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "Reading abnormal movement",
            "steps": [
                ("Subjective / observed", "two viewpoints", False),
                ("Adaptive / non-adaptive / induced", "classify it", False),
                ("Catatonic signs", "echo-, negativism, waxy", False),
                ("Stupor vs excitement", "too little / too much", False),
                ("Rule out drug effect", "e.g. tardive dyskinesia", True),
            ],
        },
        "mnemonic": {
            "word": "CATATONIA",
            "line": "CATATONIA captures the classic motor signs plus the drug caveat.",
            "items": [
                ("C", "**Copying**: echopraxia and echolalia."),
                ("A", "**Automatic obedience** to every command."),
                ("T", "**Tension between acts**: ambitendency."),
                ("A", "**Agitation**: catatonic excitement."),
                ("T", "**Tardive dyskinesia**: drug-induced, rule it out."),
                ("O", "**Opposition**: negativism."),
                ("N", "**No movement**: stupor."),
                ("I", "**Induced** postures held (waxy flexibility)."),
                ("A", "**Also organic**: catatonia needs a medical work-up."),
            ],
        },
    },
    {
        "anchor": "experience-of-self",
        "number": 7,
        "color": "blue",
        "title": "Disorders of the Experience of Self",
        "intro": "The sense of being a single, continuous, self-directed person with firm boundaries is "
                 "normally taken for granted, so this topic uses Jaspers' four aspects of the self to "
                 "map how that sense breaks down in depersonalisation, passivity phenomena and related "
                 "states.",
        "answer_flow": [
            "Definition of the experience of self",
            "Jaspers' four aspects",
            "Depersonalisation & derealisation",
            "Passivity phenomena",
            "Significance (self vs world)",
            "One-line summary",
        ],
        "points": [
            "The **experience of self (ego)** is the immediate awareness that it is *I* who think, feel "
            "and act, a certainty so basic it is never noticed until illness makes it falter and the "
            "patient feels strangely unreal.",
            "**Karl Jaspers** analysed this awareness into **four formal aspects**, a framework that "
            "lets the clinician say exactly which part of the sense of self has broken down rather than "
            "describing the disturbance only in vague terms.",
            "The first aspect is the **awareness of activity and existence** [[the sense that I am the "
            "one doing my actions and that I exist]], which is weakened in depersonalisation when "
            "patients feel detached, automatic or as if they are not really there.",
            "The second aspect is the **awareness of unity** [[the sense of being one single person at "
            "any moment]], whose disturbance can produce the rare experience of a **double self**, as "
            "though two people inhabited the same body at once.",
            "The third aspect is the **awareness of continuity** [[the sense of being the same person "
            "over time]], which can be disrupted after a psychotic break so that patients feel they are "
            "no longer the person they were before the illness.",
            "The fourth aspect is the **awareness of the boundaries of the self** [[a clear line between "
            "me and the outside world]], whose breakdown is central to the **passivity phenomena** of "
            "schizophrenia, where the border between self and world dissolves.",
            "**Depersonalisation** [[a distressing feeling of being unreal, detached or watching oneself "
            "from outside]] is the commonest disorder of the self, and crucially the patient keeps "
            "**insight**, knowing the feeling is strange and not literally true.",
            "**Derealisation** [[a feeling that the surrounding world is unreal, flat or dreamlike]] "
            "usually accompanies depersonalisation, so the patient reports that both themselves and "
            "their environment seem artificial, distant or lacking in their usual vividness.",
            "Depersonalisation and derealisation are **non-specific**, arising in anxiety, depression, "
            "tiredness, epilepsy and normal people under stress, so they point to the quality of the "
            "experience rather than to any single diagnosis.",
            "**Passivity phenomena** are disorders of the self-boundary in which the patient experiences "
            "their own thoughts, feelings, impulses or movements as **made by an external force**, so "
            "the normal ownership of inner life is lost.",
            "These include **made actions, made feelings and made impulses**, where patients believe an "
            "outside agency is controlling what they do, feel or want, and such experiences were "
            "classically treated as first-rank symptoms of schizophrenia.",
            "**Thought alienation** is a related loss of boundary, covering **thought insertion** "
            "[[thoughts put into the mind by others]], **withdrawal** [[thoughts taken out]] and "
            "**broadcasting** [[one's thoughts escaping to others]], all breaching the privacy of the self.",
            "A useful contrast is that in depersonalisation insight is **retained**, whereas in "
            "passivity phenomena insight is **lost** and the experience is held with delusional "
            "conviction, which places the two on very different levels of severity.",
            "Disturbances of the self also bear on **theory of mind** [[the ability to understand that "
            "others have their own separate thoughts and feelings]], which is impaired in schizophrenia "
            "and autism and helps explain some difficulties in reading other people.",
            "A clinical warning is that depersonalisation is often **under-reported**, because patients "
            "find the feeling so hard to put into words and fear sounding mad, so the clinician may "
            "need to ask about it directly and sympathetically.",
            "The self-disorders form a rough **hierarchy of severity**, from transient depersonalisation "
            "with full insight, through persistent depersonalisation, to the insight-less passivity "
            "phenomena that signal severe psychotic illness.",
            "In one line, disorders of the experience of self are breakdowns in Jaspers' four aspects of "
            "activity, unity, continuity and boundaries, ranging from insightful depersonalisation to "
            "the boundary-dissolving passivity phenomena of schizophrenia.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Depersonalisation",
            "text": "Depersonalisation is described as an unpleasant state in which the patient feels "
                    "changed, unreal or detached from their own experience, as if observing themselves "
                    "from outside, while retaining insight that the change is subjective and not an "
                    "actual alteration of the self.",
            "cite": "Paraphrased from Casey, P. & Kelly, B., Fish's Clinical Psychopathology, 5th ed. "
                    "(chapter on disorders of the experience of self), after Jaspers.",
        },
        "svg": SVG_SELF_ASPECTS,
        "svg_caption": ("Figure 2. Jaspers' four formal aspects of the experience of self \u2014 "
                        "awareness of activity/existence, unity, continuity and boundaries \u2014 each "
                        "of which can break down in a characteristic disorder of the self."),
        "flow": {
            "title": "From intact self to passivity",
            "steps": [
                ("Normal sense of 'I'", "activity, unity, continuity, boundary", False),
                ("Feel unreal", "depersonalisation (insight kept)", False),
                ("World unreal", "derealisation", False),
                ("Boundary dissolves", "made acts, thought alienation", False),
                ("Passivity phenomena", "insight lost", True),
            ],
        },
        "mnemonic": {
            "word": "AUCB",
            "line": "AUCB locks in Jaspers' four aspects of the self.",
            "items": [
                ("A", "**Activity / existence**: 'I am the one acting'."),
                ("U", "**Unity**: 'I am one person now'."),
                ("C", "**Continuity**: 'the same person over time'."),
                ("B", "**Boundaries**: 'me versus the world' (lost in passivity)."),
            ],
        },
    },
    {
        "anchor": "speech",
        "number": 8,
        "color": "teal",
        "title": "Disorders of Speech",
        "intro": "Speech is the chief window onto thought, so this topic separates disorders of the "
                 "mechanics of talking from disorders of its content, covering stammering and mutism, "
                 "the odd productions of psychosis, and the organic language failures known as the "
                 "aphasias.",
        "answer_flow": [
            "Definition of speech disorder",
            "Rhythm & output problems",
            "Mutism (its types)",
            "Psychotic speech productions",
            "Aphasias (receptive vs expressive)",
            "One-line summary",
        ],
        "points": [
            "**Disorders of speech** are abnormalities of the spoken output itself, and although speech "
            "and thought are closely tied, disordered speech can arise either from disturbed thinking "
            "or from a fault in the language machinery of the brain.",
            "**Stammering (stuttering)** [[the halting repetition or blocking of sounds and syllables]] "
            "is a disorder of the rhythm of speech, often worsened by anxiety, in which the flow of "
            "words is broken even though the person knows exactly what they wish to say.",
            "**Mutism** [[a complete absence of speech with the organs of speech intact]] is a striking "
            "sign that can arise in several very different conditions, so its type must be identified "
            "before its cause can be understood.",
            "**Elective (selective) mutism** describes a person, often a child, who speaks normally in "
            "some settings but refuses to speak in others, so the silence is situational rather than a "
            "loss of the capacity for speech.",
            "**Catatonic mutism** occurs as part of the catatonic syndrome, where the patient does not "
            "speak despite being conscious, while **akinetic mutism** [[silent immobility from damage "
            "near the brain's midline]] is an organic state of near-total lack of speech and movement.",
            "**Vorbeireden**, or **talking past the point** [[giving answers that are obviously wrong "
            "but clearly related to the question]], is seen in the Ganser state, as when a patient asked "
            "the colour of grass replies that it is blue.",
            "**Neologisms** [[brand-new words invented by the patient with a private meaning]] are a "
            "sign of severe disturbance, typically schizophrenia, where the patient creates words that "
            "carry no shared meaning for the listener.",
            "**Word salad (schizophasia)** [[a jumble of real words and phrases strung together without "
            "grammar or sense]] is an extreme disorganisation of speech in which the connections "
            "between words are so lost that no meaning can be followed.",
            "**Perseveration** of speech, the inappropriate repetition of a previous word or phrase, and "
            "**echolalia**, the automatic echoing of the examiner's words, are further signs that may "
            "point to either organic disease or severe psychosis.",
            "A separate major group is the **aphasias** [[losses of language caused by damage to the "
            "brain's language areas]], which are organic disorders of comprehension or production rather "
            "than disturbances of the content of thought.",
            "**Receptive (Wernicke's) aphasia** follows damage to the temporal language area, producing "
            "**fluent** speech that flows easily but is empty of meaning, together with a severe failure "
            "to understand what others say.",
            "**Expressive (Broca's) aphasia** follows damage to the frontal language area, producing "
            "**non-fluent**, effortful, telegraphic speech while comprehension is relatively preserved, "
            "so the patient understands but cannot fluently produce words.",
            "Distinguishing an **aphasia** from a thought disorder is a key clinical skill, because "
            "fluent empty Wernicke's speech can superficially resemble the word salad of schizophrenia "
            "yet points to a focal brain lesion needing neurological investigation.",
            "**Dysarthria** [[slurred speech from poor control of the speech muscles]] must also be "
            "separated from aphasia, because dysarthria disturbs the clarity of articulation while "
            "leaving the choice and meaning of words intact.",
            "The **tempo** of speech is itself informative, being fast and hard to interrupt as "
            "**pressure of speech** in mania, and slow, sparse and delayed as **poverty of speech** in "
            "depression and in the negative symptoms of schizophrenia.",
            "Careful observation of speech therefore lets the examiner infer the underlying process, "
            "since the **form** of the output, fluent or halting, pressured or impoverished, novel or "
            "echoed, often reveals more than the literal content of what is said.",
            "In one line, disorders of speech range from stammering and the several mutisms, through the "
            "neologisms and word salad of psychosis, to the receptive and expressive aphasias of focal "
            "brain damage.",
        ],
        "definition": None,
        "svg": None,
        "svg_caption": "",
        "figures": [],
        "flow": {
            "title": "Locating a speech problem",
            "steps": [
                ("Is there output?", "mutism if none", False),
                ("Rhythm / articulation", "stammer, dysarthria", False),
                ("Novel / jumbled words", "neologism, word salad", False),
                ("Fluent-empty vs non-fluent", "Wernicke vs Broca", False),
                ("Infer the process", "thought vs lesion", True),
            ],
        },
        "mnemonic": {
            "word": "SPEAKS",
            "line": "SPEAKS runs through the main disorders of spoken output.",
            "items": [
                ("S", "**Stammering**: broken rhythm of speech."),
                ("P", "**Past the point**: Vorbeireden (Ganser)."),
                ("E", "**Expressive aphasia**: non-fluent Broca's."),
                ("A", "**Absent speech**: elective, catatonic, akinetic mutism."),
                ("K", "**Konstructed words**: neologisms and word salad."),
                ("S", "**Sensory/receptive aphasia**: fluent, empty Wernicke's."),
            ],
        },
    },
    {
        "anchor": "thought",
        "number": 9,
        "color": "amber",
        "title": "Disorders of Thought",
        "intro": "Thought is examined along four separate lines, so this topic sets out disorders of the "
                 "stream, the possession, the content and the form of thinking, keeping delusions, "
                 "obsessions, flight of ideas and formal thought disorder clearly apart.",
        "answer_flow": [
            "Definition & the four lines",
            "Disorders of stream (tempo)",
            "Disorders of possession",
            "Disorders of content (delusions)",
            "Disorders of form (FTD)",
            "One-line summary",
        ],
        "points": [
            "**Disorders of thought** are best organised under four separate headings, the **stream, the "
            "possession, the content and the form** of thinking, because a patient may have a disorder "
            "of one of these while the others remain entirely normal.",
            "Disorders of the **stream** concern the **tempo and flow** of thought, that is how fast "
            "thoughts come and how smoothly one leads to the next, independent of whether the thoughts "
            "themselves are true or false.",
            "**Flight of ideas** [[thoughts racing so fast that speech jumps between loosely linked "
            "topics]] is the classic stream disorder of mania, where the connections, though rapid, can "
            "usually still be traced through rhymes, puns and chance associations.",
            "**Inhibition or retardation of thinking** is the opposite, a slowing of the stream seen in "
            "depression, where thoughts come sparsely and effortfully and the patient describes their "
            "mind as empty, heavy or stuck.",
            "**Circumstantiality** [[reaching the point only after many unnecessary details and "
            "digressions]] keeps the goal of thought in view but approaches it by a long, winding "
            "route, eventually arriving where the question intended.",
            "**Perseveration** [[senseless repetition of a thought or word beyond the point of "
            "relevance]] and **thought blocking** [[a sudden halt where the thread of thought vanishes "
            "entirely]] are further stream disorders, the latter strongly suggestive of schizophrenia.",
            "Disorders of the **possession** of thought concern whether thinking feels **owned and "
            "controlled** by the self, and they cover both the obsessional phenomena and the alienation "
            "of thought seen in psychosis.",
            "**Obsessions** [[recurrent unwanted thoughts the person recognises as their own and "
            "resists]] and **compulsions** [[repetitive acts done to relieve the anxiety of an "
            "obsession]] are possession disorders in which control, not ownership, is lost.",
            "By contrast, **thought alienation** is a loss of **ownership**, covering insertion, "
            "withdrawal and broadcasting, where the patient feels thoughts are not their own, a far "
            "more severe disturbance typical of schizophrenia.",
            "Disorders of the **content** of thought concern **what** is believed, and chief among them "
            "is the **delusion** [[a fixed false belief held firmly despite clear contrary evidence and "
            "out of keeping with the person's culture]].",
            "A **primary (autochthonous) delusion** arises fully formed out of nowhere, not understandable "
            "from the patient's mood or situation, whereas a **secondary delusion** grows understandably "
            "from a preceding mood state, perception or other experience.",
            "Delusions are further grouped by **theme**, including persecutory, grandiose, of reference, "
            "of guilt, nihilistic, hypochondriacal and of jealousy or love, and the dominant theme "
            "often reflects the underlying illness.",
            "**Overvalued ideas** [[strongly held, understandable beliefs pursued beyond reason but "
            "short of delusional conviction]] lie between normal beliefs and delusions, and "
            "distinguishing them matters because they imply a different diagnosis and prognosis.",
            "Disorders of the **form** of thought, or **formal thought disorder** [[a breakdown in the "
            "logical connections between thoughts]], show in disorganised speech and are inferred "
            "because thought itself cannot be observed directly.",
            "Signs of formal thought disorder include **loosening of associations** [[ideas that drift "
            "apart with no clear link]] and the **knight's move (derailment)** [[sudden jumps from one "
            "topic to an unrelated one]], both pointing toward schizophrenia.",
            "Keeping these four lines apart is a core exam skill, since flight of ideas is a **stream** "
            "disorder of mania whereas loosening of associations is a **form** disorder of "
            "schizophrenia, and confusing them points to the wrong diagnosis.",
            "In one line, disorders of thought divide into the stream, the possession, the content and "
            "the form, housing respectively flight of ideas, obsessions and alienation, delusions, and "
            "formal thought disorder.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Delusion",
            "text": "A delusion is a false, unshakeable belief that is held with absolute conviction, is "
                    "not amenable to reason or contrary evidence, and is out of keeping with the "
                    "patient's social, cultural and educational background; it may be primary "
                    "(arising inexplicably) or secondary (arising understandably from another "
                    "experience).",
            "cite": "Paraphrased from Casey, P. & Kelly, B., Fish's Clinical Psychopathology, 5th ed. "
                    "(chapter on disorders of thought and belief), after Jaspers.",
        },
        "svg": SVG_THOUGHT_MAP,
        "svg_caption": ("Figure 3. The four lines along which thought is examined \u2014 stream (tempo), "
                        "possession (ownership and control), content (beliefs, delusions) and form "
                        "(logical connections) \u2014 each with its own disorders."),
        "flow": {
            "title": "Examining thought in four lines",
            "steps": [
                ("Stream", "tempo: flight, block", False),
                ("Possession", "obsession, alienation", False),
                ("Content", "delusions, themes", False),
                ("Form", "loosening, knight's move", False),
                ("Name the disorder", "match line to sign", True),
            ],
        },
        "mnemonic": {
            "word": "SPCF",
            "line": "SPCF keeps the four lines of thought disorder separate.",
            "items": [
                ("S", "**Stream**: flight of ideas, inhibition, blocking."),
                ("P", "**Possession**: obsessions and thought alienation."),
                ("C", "**Content**: delusions (primary vs secondary) and themes."),
                ("F", "**Form**: formal thought disorder, loosening, knight's move."),
            ],
        },
    },
    {
        "anchor": "perception",
        "number": 10,
        "color": "rose",
        "title": "Disorders of Perception",
        "intro": "Perception can go wrong in graded ways, so this topic moves from sensory distortions of "
                 "real objects, through illusions, to hallucinations proper, and sorts them by modality "
                 "while distinguishing true hallucinations from pseudohallucinations.",
        "answer_flow": [
            "Definition of perceptual disorder",
            "Sensory distortions",
            "Illusions",
            "Hallucinations (by modality)",
            "Pseudohallucinations & special types",
            "One-line summary",
        ],
        "points": [
            "**Disorders of perception** are faults in how the mind builds its picture of the world from "
            "the senses, graded from distortions of a real object, through misinterpretations, to "
            "perceptions with no object at all.",
            "**Sensory distortions** are changes in how a **real** object is perceived, altering its "
            "intensity, quality or spatial form, as in **hyperacusis** [[sounds seeming abnormally "
            "loud]] or **micropsia** [[objects seeming smaller than they are]].",
            "Distortions of **intensity** make stimuli seem stronger or weaker, distortions of "
            "**quality** change colour or tone, and **spatial** distortions change apparent size, shape "
            "or distance, each pointing to specific drug or organic causes.",
            "An **illusion** [[a real external object that is misperceived or misinterpreted]] differs "
            "from a distortion because the whole object is read wrongly, as when a shadow is taken for "
            "an intruder, and illusions are commoner when attention is low or affect is strong.",
            "Illusions are usually **normal phenomena**, arising in poor light, fatigue, fear or "
            "inattention, so they carry far less diagnostic weight than hallucinations and often "
            "resolve as soon as attention is properly directed to the object.",
            "A **hallucination** is the key perceptual disorder, classically defined as a **perception "
            "without an object**, arising in external space with the full force of a true perception "
            "yet corresponding to nothing actually present.",
            "**Auditory hallucinations** are the commonest in psychiatry, ranging from simple noises to "
            "**voices**, and in schizophrenia the voices may give a running commentary, speak in the "
            "third person, or repeat the patient's own thoughts aloud.",
            "**Visual hallucinations** are more suggestive of an **organic or drug-induced** state such "
            "as delirium tremens, so their presence should prompt a search for a physical cause rather "
            "than being assumed to indicate a primary psychiatric illness.",
            "**Olfactory and gustatory hallucinations** of smell and taste, often unpleasant, point "
            "toward temporal-lobe epilepsy or organic disease, while **tactile (haptic)** hallucinations "
            "include the sensation of insects crawling on or under the skin.",
            "Hallucinations are further classed by **form**: an **elementary** hallucination is a simple "
            "flash or bang, whereas a **complex** one is a fully formed voice, figure or scene, the "
            "latter carrying more specific diagnostic meaning.",
            "**Pseudohallucinations** [[vivid images experienced in inner subjective space and "
            "recognised as not real]] differ from true hallucinations because they lack the quality of "
            "an external percept and the patient retains insight into their unreal nature.",
            "Several **special hallucinations** are defined by their trigger or timing, such as "
            "**functional** hallucinations [[a voice heard only while a real sound is present]] and the "
            "**hypnagogic** and **hypnopompic** hallucinations that occur while falling asleep or waking.",
            "**Charles Bonnet syndrome** is the occurrence of complex visual hallucinations in people "
            "with serious **eye disease** but a clear mind and full insight, a benign condition that "
            "must not be mistaken for a psychotic illness.",
            "**Autoscopy** [[seeing a double of oneself]] and **extracampine** hallucinations [[a "
            "perception located outside the possible sensory field, such as hearing a voice from "
            "hundreds of miles away]] are further special forms worth recognising.",
            "**Disorders of body image** belong here too, including the **phantom limb** [[feeling a "
            "limb that has been amputated]] and distortions in which parts of the body feel changed in "
            "size, shape or position.",
            "Distinguishing an **illusion from a hallucination** is a core skill, since an illusion needs "
            "a real object to misinterpret whereas a hallucination needs none, and only the latter, "
            "especially when complex and insight-less, strongly signals serious illness.",
            "In one line, disorders of perception run from distortions of real objects, through "
            "illusions, to true hallucinations sorted by modality, with pseudohallucinations and "
            "special types such as Charles Bonnet rounding out the picture.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Hallucination",
            "text": "A hallucination is a perception without an object \u2014 a percept experienced in "
                    "external space, with the full force and clarity of a true perception, but arising "
                    "in the absence of any actual external stimulus and not subject to conscious "
                    "manipulation by the perceiver.",
            "cite": "Paraphrased from Casey, P. & Kelly, B., Fish's Clinical Psychopathology, 5th ed. "
                    "(chapter on disorders of perception); the 'perception without an object' "
                    "formulation is attributed there to Esquirol / Jaspers.",
        },
        "svg": SVG_PERCEPTION_TREE,
        "svg_caption": ("Figure 4. The tree of perceptual disorders \u2014 sensory distortions (a real "
                        "object wrongly perceived), illusions (a real object misinterpreted) and "
                        "hallucinations (perception without an object), split into true and "
                        "pseudohallucinations."),
        "flow": {
            "title": "Grading a perceptual disorder",
            "steps": [
                ("Is there a real object?", "yes = distortion/illusion", False),
                ("Distortion", "intensity/quality/space", False),
                ("Illusion", "object misinterpreted", False),
                ("No object", "hallucination", False),
                ("True vs pseudo", "outer vs inner space", True),
            ],
        },
        "mnemonic": {
            "word": "SIGHT",
            "line": "SIGHT orders perception from distortion to hallucination.",
            "items": [
                ("S", "**Sensory distortion**: real object, wrong intensity/quality."),
                ("I", "**Illusion**: real object misinterpreted."),
                ("G", "**Genuine hallucination**: perception without an object."),
                ("H", "**Hearing voices** commonest; vision suggests organic."),
                ("T", "**Two insights**: pseudohallucination keeps insight, true loses it."),
            ],
        },
    },
    {
        "anchor": "emotion",
        "number": 11,
        "color": "purple",
        "title": "Disorders of Emotion",
        "intro": "Emotional life has several layers and many ways of going wrong, so this topic first "
                 "separates feeling, emotion, mood and affect, then grades disturbance across five "
                 "levels from normal reactions to the morbid disorders of depression, mania and anxiety.",
        "answer_flow": [
            "Definitions (feeling/emotion/mood/affect)",
            "The five levels of disturbance",
            "Abnormal reactions & expressions",
            "Morbid disorders (depression/mania)",
            "Anxiety and ecstasy",
            "One-line summary",
        ],
        "points": [
            "Clear terms matter: a **feeling** is a brief positive or negative state, an **emotion** a "
            "stirred-up feeling with bodily change, **mood** the prevailing background over time, and "
            "**affect** the emotion shown in the moment.",
            "A helpful image is that **mood** is the climate while **affect** is the weather, so the "
            "examiner records both the sustained background the patient reports and the moment-to-moment "
            "emotional expression that is observed during the interview.",
            "Fish grades emotional disturbance across **five levels**, running from normal emotions, "
            "through abnormal reactions and abnormal expressions, to morbid expressions and finally the "
            "morbid disorders of emotion, giving a clear ladder of increasing severity.",
            "Level one is simply **normal emotion**, appropriate joy, fear, sadness or anger that fits "
            "its trigger in kind, strength and duration, providing the baseline against which every "
            "abnormality is judged.",
            "Level two, **abnormal emotional reactions**, are responses that are understandable in kind "
            "but **too intense or too prolonged** for their trigger, such as grief that is far deeper "
            "or far longer than the loss would ordinarily warrant.",
            "Level three, **abnormal expressions of emotion**, includes **incongruity of affect** "
            "[[emotion that clashes with the situation, such as laughing at sad news]], emotional "
            "**lability** [[rapid shifts of emotion]], and **perplexity**, a bewildered emotional state.",
            "Level four, **morbid expressions**, covers **apathy** [[a loss of feeling and interest]] "
            "and **blunting or flattening of affect** [[a marked reduction in the range and intensity "
            "of emotional expression]], both prominent in the negative symptoms of schizophrenia.",
            "Level five, the **morbid disorders of emotion**, are the sustained pathological mood states "
            "of depression, mania, pathological anxiety and ecstasy, which form the core of the "
            "affective disorders seen in psychiatry.",
            "In **depression** the mood is persistently low with **anhedonia** [[loss of the ability to "
            "feel pleasure]], hopelessness, guilt and often diurnal variation, so the whole emotional "
            "field is pulled down and fails to lift with good events.",
            "In **mania** the mood is abnormally elevated, expansive or irritable, with **elation** "
            "that is often infectious but brittle, tipping quickly into irritability when the patient's "
            "grandiose plans are thwarted.",
            "**Pathological anxiety** is fear or dread out of proportion to any real threat, with bodily "
            "symptoms such as palpitations, sweating and trembling, and it becomes a disorder when it is "
            "persistent, distressing and impairs the person's functioning.",
            "**Ecstasy** [[an intense feeling of rapture, bliss or oneness with the world]] is a rarer "
            "morbid emotional state that can occur in mania, in some epilepsies and in certain religious "
            "or drug-induced experiences, and is often held with conviction.",
            "A distinction worth stressing is between **endogenous**-type mood change, seeming to arise "
            "from within with biological features, and **reactive** mood change clearly triggered by "
            "events, a contrast that historically guided treatment choices.",
            "The clinician assesses emotion through **subjective report**, what the patient says they "
            "feel, and **objective observation** of facial expression, posture, tears and movement, "
            "because these two sources can diverge, as in the smiling, hidden depression.",
            "Emotional symptoms must be judged against **cultural norms** for expressing feeling, since "
            "what looks like blunting or over-expression in one culture may be entirely normal in "
            "another, so context guards against over-diagnosis.",
            "Because emotional change colours thought, perception and behaviour, a **secondary "
            "delusion** or an illusion may grow understandably out of a morbid mood, which is why mood "
            "is always assessed alongside the other domains rather than in isolation.",
            "In one line, disorders of emotion are graded across five levels from abnormal reactions to "
            "morbid disorders, the latter comprising depression, mania, pathological anxiety and "
            "ecstasy, each judged against normal, culturally appropriate feeling.",
        ],
        "definition": None,
        "svg": SVG_EMOTION_LEVELS,
        "svg_caption": ("Figure 5. Emotion graded across five levels in Fish \u2014 from normal "
                        "emotions, through abnormal reactions and abnormal expressions, to morbid "
                        "expressions and the morbid disorders of emotion (depression, mania, anxiety, "
                        "ecstasy)."),
        "figures": [],
        "flow": {
            "title": "Five levels of emotional disturbance",
            "steps": [
                ("Normal emotion", "fits the trigger", False),
                ("Abnormal reaction", "too strong / long", False),
                ("Abnormal expression", "incongruity, lability", False),
                ("Morbid expression", "apathy, flat affect", False),
                ("Morbid disorder", "depression, mania, anxiety", True),
            ],
        },
        "mnemonic": {
            "word": "MOODS",
            "line": "MOODS tracks emotion from the normal to the morbid.",
            "items": [
                ("M", "**Mood** is the climate, affect the weather."),
                ("O", "**Over-reaction**: abnormal reactions (too strong/long)."),
                ("O", "**Odd expression**: incongruity, lability, perplexity."),
                ("D", "**Dampened**: apathy and blunted/flat affect."),
                ("S", "**Severe/morbid**: depression, mania, anxiety, ecstasy."),
            ],
        },
    },
    {
        "anchor": "memory",
        "number": 12,
        "color": "navy",
        "title": "Disorders of Memory",
        "intro": "Memory is a chain of linked stages, so this topic defines registration, retention, "
                 "retrieval and recall, then surveys the amnesias, the distinction between organic and "
                 "psychogenic memory loss, and the paramnesias in which memory is falsified rather than "
                 "lost.",
        "answer_flow": [
            "Definition & the memory stages",
            "The amnesias",
            "Organic vs psychogenic loss",
            "Ribot's law",
            "Paramnesias (falsified memory)",
            "One-line summary",
        ],
        "points": [
            "**Memory** is the capacity to store and later use information, and it is usefully broken "
            "into four linked stages, **registration, retention, retrieval and recall**, since a fault "
            "at any one stage will present as a memory complaint.",
            "**Registration** [[taking information in]] depends on intact attention and consciousness, "
            "so a patient who cannot attend never truly registers material, and what was never "
            "registered cannot later be remembered however hard they try.",
            "**Retention** [[holding information in store over time]] and **retrieval** [[finding "
            "stored information again]] are the middle stages, and disorders can strike either the "
            "storing of new memories or the reaching back for old ones.",
            "**Recall** [[bringing stored information back to mind]] is the final, testable stage, and "
            "because it is what the examiner directly observes, a failure of recall must be traced back "
            "to see which earlier stage actually broke down.",
            "**Amnesia** [[loss of memory]] is the central disorder, and a key split is between "
            "**anterograde amnesia** [[failure to lay down new memories after an event]] and "
            "**retrograde amnesia** [[loss of memories formed before the event]].",
            "In the classic **amnesic (Korsakoff) syndrome**, severe anterograde amnesia from thiamine "
            "deficiency leaves the patient unable to form new memories while older memories and "
            "immediate attention are relatively spared, often with confabulation.",
            "A crucial clinical distinction is between **organic** memory loss, which follows the "
            "pattern of brain pathology, and **psychogenic** memory loss, which is driven by "
            "psychological factors and often spares the kind of information organic amnesia would lose.",
            "**Psychogenic amnesia** characteristically targets **personal, emotionally charged** "
            "memories, as in dissociative amnesia and fugue, so a patient may forget their own identity "
            "yet retain general knowledge, a pattern organic disease rarely produces.",
            "By contrast, organic amnesia typically spares deep **personal identity** until very late "
            "while heavily damaging the laying down of new facts, so the two patterns are in many ways "
            "mirror images of each other.",
            "**Ribot's law** states that in organic memory loss the **most recent memories are lost "
            "first** and the oldest, best-rehearsed memories survive longest, giving a predictable "
            "gradient that helps confirm an organic cause.",
            "This recency gradient explains why a patient with dementia may vividly recount their "
            "childhood while being unable to recall the morning's conversation, a pattern that is "
            "diagnostically reassuring of an organic rather than a feigned loss.",
            "A separate group are the **paramnesias** [[distortions or falsifications of memory]], in "
            "which memory is not simply lost but altered, so the patient remembers things that did not "
            "happen or misremembers things that did.",
            "**Confabulation** [[filling gaps in memory with invented but sincerely believed accounts]] "
            "is the classic paramnesia of the amnesic syndrome, where the patient unknowingly produces "
            "false memories to bridge what they cannot recall.",
            "**Deja vu** [[the false sense that a new situation has been experienced before]] and its "
            "opposite **jamais vu** [[a familiar situation feeling strange and new]] are paramnesias "
            "that occur normally but may be prominent in temporal-lobe epilepsy.",
            "The **Capgras** delusion, in which a familiar person is believed to be an identical "
            "impostor, is a disorder at the border of memory and recognition, reflecting a split "
            "between knowing a face and feeling its familiarity.",
            "At the opposite extreme are **hyperamnesia** [[abnormally intense, detailed memory]] and "
            "the vivid **flashbulb memory** of an emotionally charged event, showing that memory can be "
            "pathologically heightened as well as lost or falsified.",
            "In one line, disorders of memory are failures or falsifications along the chain of "
            "registration, retention, retrieval and recall, divided into organic and psychogenic "
            "amnesias, Ribot's recency gradient, and the paramnesias such as confabulation.",
        ],
        "definition": None,
        "svg": SVG_MEMORY_CHAIN,
        "svg_caption": ("Figure 6. The memory chain \u2014 registration, retention, retrieval and recall "
                        "\u2014 the four linked stages, a failure in any one of which presents as "
                        "amnesia and which frames memory testing at the bedside."),
        "figures": [],
        "flow": {
            "title": "Tracing a memory failure",
            "steps": [
                ("Registration", "needs attention", False),
                ("Retention", "storage over time", False),
                ("Retrieval / recall", "finding it again", False),
                ("Organic or psychogenic?", "Ribot vs personal loss", False),
                ("Lost or falsified?", "amnesia vs paramnesia", True),
            ],
        },
        "mnemonic": {
            "word": "RECALL",
            "line": "RECALL maps the stages and the disorders of memory.",
            "items": [
                ("R", "**Registration** first (needs attention)."),
                ("E", "**Encoding/retention** stores it over time."),
                ("C", "**Confabulation** fills gaps with false memories."),
                ("A", "**Anterograde vs retrograde** amnesia."),
                ("L", "**Law of Ribot**: recent lost first, old survives."),
                ("L", "**Lost vs falsified**: amnesia versus paramnesia."),
            ],
        },
    },
]

REFERENCES = {
    "books": [
        "Casey, Patricia, and Brendan Kelly, Fish's Clinical Psychopathology: Signs and Symptoms in "
        "Psychiatry, 5th ed. (Royal College of Psychiatrists / Cambridge University Press, 2024) "
        "\u2014 the primary source and organising spine for every topic in these notes, from "
        "classification and its limitations through all ten domains of signs and symptoms.",
        "Oyebode, Femi, Sims' Symptoms in the Mind: Textbook of Descriptive Psychopathology \u2014 "
        "used as a standard cross-reference for the descriptions of hallucinations, delusions, "
        "disorders of the self and the terminology of emotion.",
    ],
    "online": [],
    "flagged": [
        "The historical edition dates for the DSM (DSM-I 1952, DSM-II 1968, DSM-III 1980, DSM-IV 1994, "
        "DSM-5 2013) and the ICD (ICD-7 1955, ICD-8 1965, ICD-9 1978, ICD-10 1992, ICD-11 adopted 2022) "
        "are well-established publication dates; they are stated here as general reference facts and "
        "were not re-derived from a single cited page.",
        "The attributions to Jaspers (ideal types; the four aspects of the self), Kendler 2016 "
        "(realism/constructivism/pragmatism), Frances ('seeds of tautology'), Tyrer 2017 "
        "(comorbidity/consanguinity/co-occurrence) and Ribot (law of regression) follow the way these "
        "authors are discussed within Fish's Clinical Psychopathology; the primary papers were not "
        "accessed directly for these notes.",
    ],
}


def figure_credits():
    """Build the figure-source list from the figures declared in CONTENT.

    Every topic uses only inline labelled SVG schematics (figures lists are empty),
    so this function must not crash on empty figures lists; it simply records the
    original-schematic note for the inline diagrams.
    """
    credits = []
    seen = set()
    for topic in CONTENT:
        for fig in topic.get("figures", []):
            caption = fig["caption"]
            fig_label = caption.split(".", 1)[0].strip() if "." in caption else caption
            if fig.get("credit"):
                entry = "%s \u2014 %s" % (fig_label, fig["credit"])
            else:
                entry = ("%s \u2014 %s by %s. License: %s."
                         % (fig_label, fig.get("source", ""), fig.get("author", ""),
                            fig.get("license", "")))
            if entry not in seen:
                seen.add(entry)
                credits.append(entry)
    credits.append(
        "The DSM/ICD dual timeline, the illusions-and-hallucinations tree, Jaspers' four aspects of "
        "the self, the four-line map of thought, the registration-to-recall memory chain and the "
        "five-levels-of-emotion diagram are original labelled schematics drawn for these notes.")
    return credits


REFERENCES["figures"] = figure_credits()

# ---------------------------------------------------------------------------
# Rendering helpers (shared formatting markup: **bold** and [[plain]])
# ---------------------------------------------------------------------------
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_PLAIN_RE = re.compile(r"\[\[(.+?)\]\]")


def render_inline_html(text):
    """Convert **bold** -> <strong> and [[plain]] -> distinct span, with escaping."""
    out = []
    i = 0
    token_re = re.compile(r"\*\*(.+?)\*\*|\[\[(.+?)\]\]")
    for m in token_re.finditer(text):
        out.append(_html.escape(text[i:m.start()]))
        if m.group(1) is not None:
            out.append("<strong>" + _html.escape(m.group(1)) + "</strong>")
        else:
            out.append('<span class="plain">' + _html.escape(m.group(2)) + "</span>")
        i = m.end()
    out.append(_html.escape(text[i:]))
    return "".join(out)


def strip_markup(text):
    """Plain text version of the markup (kept for completeness)."""
    text = _BOLD_RE.sub(r"\1", text)
    text = _PLAIN_RE.sub(r"(\1)", text)
    return text


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------
CSS = """
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  background: #f4f7fb;
  color: #1a2535;
  font-family: Georgia, 'Times New Roman', serif;
  font-size: 16px;
  line-height: 1.75;
}
.cover {
  min-height: 100vh;
  display: flex; flex-direction: column; justify-content: center; align-items: center;
  text-align: center;
  background: linear-gradient(135deg, #0f2744 0%, #1d5fa6 55%, #0e7c6b 100%);
  color: #fff; padding: 60px 28px;
}
.cover .label {
  font-family: Arial, Helvetica, sans-serif;
  letter-spacing: 4px; text-transform: uppercase; font-size: 13px;
  color: rgba(255,255,255,0.7); margin-bottom: 22px;
}
.cover h1 {
  font-size: clamp(28px, 5vw, 46px); line-height: 1.25; max-width: 900px;
  margin-bottom: 20px;
}
.cover .subtitle {
  font-size: 17px; color: rgba(255,255,255,0.85); max-width: 660px; margin-bottom: 38px;
  font-style: italic;
}
.cover .credit {
  font-family: Arial, Helvetica, sans-serif;
  display: inline-block; border: 1px solid rgba(255,255,255,0.4);
  border-radius: 30px; padding: 10px 24px; font-size: 15px; font-weight: 600;
  background: rgba(255,255,255,0.08);
}
.wrap { max-width: 960px; margin: 0 auto; padding: 48px 26px 70px; }
.toc {
  background: #fff; border: 1px solid #dce6f0; border-radius: 12px;
  padding: 26px 30px; margin-bottom: 40px; box-shadow: 0 2px 10px rgba(15,39,68,0.05);
}
.toc h2 { font-size: 22px; color: #0f2744; margin-bottom: 14px; }
.toc ol { margin-left: 22px; }
.toc li { margin: 6px 0; }
.toc a { color: #1d5fa6; text-decoration: none; }
.toc a:hover { text-decoration: underline; }
.guide {
  background: #eef6ff; border: 1px solid #cfe2f7; border-left: 5px solid #1d5fa6;
  border-radius: 10px; padding: 22px 26px; margin-bottom: 44px;
}
.guide h2 { color: #0f2744; font-size: 22px; margin-bottom: 12px; }
.guide ul { margin-left: 22px; }
.guide li { margin: 8px 0; }
.card {
  background: #fff; border: 1px solid #dce6f0; border-radius: 12px;
  padding: 28px 30px; margin-bottom: 40px; box-shadow: 0 2px 12px rgba(15,39,68,0.06);
  border-left: 6px solid #1d5fa6;
}
.card.blue   { border-left-color: #1d5fa6; }
.card.teal   { border-left-color: #0e7c6b; }
.card.amber  { border-left-color: #8a5c00; }
.card.rose   { border-left-color: #7a1f3d; }
.card.purple { border-left-color: #4a2080; }
.card.navy   { border-left-color: #0f2744; }
.card-head { display: flex; align-items: center; gap: 14px; margin-bottom: 8px; }
.sec-num {
  width: 38px; height: 38px; border-radius: 50%; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  font-family: Arial, Helvetica, sans-serif; font-weight: 700; color: #fff; font-size: 16px;
  background: #1d5fa6;
}
.card.teal   .sec-num { background: #0e7c6b; }
.card.amber  .sec-num { background: #8a5c00; }
.card.rose   .sec-num { background: #7a1f3d; }
.card.purple .sec-num { background: #4a2080; }
.card.navy   .sec-num { background: #0f2744; }
.card-head h2 { font-size: 24px; color: #0f2744; }
.intro { font-style: italic; color: #40506a; margin: 10px 0 18px; font-size: 16.5px; }
.answer-flow {
  background: #fffaf0; border: 1px solid #e7c766; border-left: 5px solid #8a5c00;
  border-radius: 10px; padding: 16px 20px; margin: 0 0 24px;
}
.answer-flow .af-label {
  font-family: Arial, Helvetica, sans-serif; font-size: 12.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 1px; color: #8a5c00; margin-bottom: 10px;
}
.answer-flow .af-steps {
  display: flex; flex-wrap: wrap; align-items: center; gap: 6px;
  font-family: Arial, Helvetica, sans-serif;
}
.answer-flow .af-step {
  background: #fff; border: 1px solid #e0c38f; border-radius: 20px;
  padding: 5px 13px; font-size: 13.5px; color: #5a4300; font-weight: 600;
}
.answer-flow .af-arrow { color: #8a5c00; font-weight: 700; font-size: 15px; }
.points { counter-reset: pt; list-style: none; }
.points li {
  position: relative; padding: 10px 0 10px 40px; border-bottom: 1px dashed #e4ebf3;
}
.points li:before {
  counter-increment: pt; content: counter(pt);
  position: absolute; left: 0; top: 10px;
  width: 26px; height: 26px; border-radius: 50%;
  background: #e1edf8; color: #1d5fa6; font-family: Arial, sans-serif;
  font-size: 13px; font-weight: 700; display: flex; align-items: center; justify-content: center;
}
.plain {
  background: #fff3d6; color: #6b4e00; border: 1px solid #e7c766;
  border-radius: 6px; padding: 1px 8px; font-size: 14.5px;
  font-family: Arial, Helvetica, sans-serif; font-style: normal;
  display: inline-block; line-height: 1.5;
}
strong { color: #10243f; font-weight: 700; }
.alert-box {
  background: #f0eafa; border: 1px solid #d6c7f0; border-left: 5px solid #4a2080;
  border-radius: 10px; padding: 18px 22px; margin: 22px 0;
}
.alert-box .alert-label {
  font-family: Arial, Helvetica, sans-serif; font-size: 12.5px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 1px; color: #4a2080; margin-bottom: 8px;
}
.alert-box .cite { display: block; margin-top: 10px; font-size: 13.5px; color: #5a6a7e; }
.block-title {
  font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 1px; color: #5a6a7e; margin: 26px 0 12px;
}
.flow {
  display: flex; flex-wrap: wrap; align-items: stretch; gap: 8px;
  background: #f7fafd; border: 1px dashed #c6d6e6; border-radius: 10px; padding: 18px;
}
.flow-box {
  background: #fff; border: 1px solid #bcd0e6; border-radius: 8px;
  padding: 10px 14px; font-size: 14px; font-family: Arial, Helvetica, sans-serif;
  display: flex; flex-direction: column; justify-content: center; min-width: 110px; text-align: center;
}
.flow-box small { color: #5a6a7e; font-size: 12px; }
.flow-box.end { background: #0f2744; color: #fff; border-color: #0f2744; }
.flow-box.end small { color: rgba(255,255,255,0.8); }
.flow-arrow { display: flex; align-items: center; color: #1d5fa6; font-size: 22px; font-weight: 700; }
.mnemonic {
  background: #e3f5f1; border: 1px solid #b5e0d6; border-left: 5px solid #0e7c6b;
  border-radius: 10px; padding: 18px 22px; margin-top: 10px;
}
.mnemonic .mword {
  font-family: Arial, Helvetica, sans-serif; font-weight: 700; font-size: 20px;
  color: #0e7c6b; letter-spacing: 2px;
}
.mnemonic .mline { font-style: italic; color: #33514a; margin: 6px 0 12px; }
.mnemonic ul { list-style: none; }
.mnemonic li { margin: 5px 0; }
.mnemonic .ml {
  display: inline-block; width: 26px; font-family: Arial, sans-serif; font-weight: 700;
  color: #0e7c6b; font-size: 16px;
}
.fig {
  margin: 22px 0; padding: 16px; background: #f7fafd;
  border: 1px solid #d3e1ee; border-radius: 10px; text-align: center;
}
.fig img, .fig svg {
  max-width: 100%; height: auto; border: 1px solid #dce6f0;
  border-radius: 8px; background: #fff;
}
.fig figcaption {
  margin-top: 10px; font-size: 14px; color: #40506a; line-height: 1.5; text-align: left;
}
.fig figcaption .credit {
  display: block; margin-top: 6px; font-size: 12.5px; color: #5a6a7e;
  font-family: Arial, Helvetica, sans-serif;
}
.refs { background: #fff; border: 1px solid #dce6f0; border-radius: 12px; padding: 26px 30px; }
.refs h2 { color: #0f2744; font-size: 24px; margin-bottom: 14px; }
.refs h3 { color: #1d5fa6; font-size: 16px; margin: 16px 0 8px; font-family: Arial, sans-serif; }
.refs ul { margin-left: 22px; }
.refs li { margin: 8px 0; }
.footer {
  background: #0f2744; color: rgba(255,255,255,0.9); text-align: center;
  padding: 26px 20px; font-family: Arial, Helvetica, sans-serif; font-size: 14px;
}
.footer .credit { font-weight: 700; color: #fff; }
.footer .title-line { display: block; margin-bottom: 8px; font-size: 14.5px; color: rgba(255,255,255,0.92); }
@media (max-width: 640px) {
  .flow { flex-direction: column; }
  .flow-arrow { transform: rotate(90deg); }
}
"""


def render_answer_flow(stages):
    """Render the exam 'Answer flow' skeleton as a small visible box."""
    parts = ['<div class="answer-flow">',
             '<div class="af-label">Answer flow &middot; write your answer in this order</div>',
             '<div class="af-steps">']
    n = len(stages)
    for idx, stage in enumerate(stages):
        parts.append('<span class="af-step">%s</span>' % _html.escape(stage))
        if idx < n - 1:
            parts.append('<span class="af-arrow">&rarr;</span>')
    parts.append("</div></div>")
    return "\n".join(parts)


def render_flow(flow):
    parts = ['<div class="block-title">Flowchart &middot; ' + _html.escape(flow["title"]) + "</div>",
             '<div class="flow">']
    n = len(flow["steps"])
    for idx, (label, sub, is_end) in enumerate(flow["steps"]):
        cls = "flow-box end" if is_end else "flow-box"
        box = ('<div class="%s">%s<br><small>%s</small></div>'
               % (cls, _html.escape(label), _html.escape(sub)))
        parts.append(box)
        if idx < n - 1:
            parts.append('<div class="flow-arrow">&rarr;</div>')
    parts.append("</div>")
    return "\n".join(parts)


def render_mnemonic(mn):
    parts = ['<div class="block-title">Mnemonic</div>', '<div class="mnemonic">',
             '<div class="mword">' + _html.escape(mn["word"]) + "</div>",
             '<div class="mline">' + render_inline_html(mn["line"]) + "</div>", "<ul>"]
    for letter, text in mn["items"]:
        parts.append('<li><span class="ml">%s</span> %s</li>'
                     % (_html.escape(letter), render_inline_html(text)))
    parts.append("</ul></div>")
    return "\n".join(parts)


def render_definition(d):
    return ('<div class="alert-box">'
            '<div class="alert-label">%s</div>'
            '<div>%s</div>'
            '<span class="cite">Source: %s</span>'
            '</div>') % (_html.escape(d["label"]), _html.escape(d["text"]), _html.escape(d["cite"]))


_IMG_MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif"}


def _data_uri(file_rel):
    """Read an image file and return an inline base64 data: URI (keeps HTML offline)."""
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, file_rel)
    ext = os.path.splitext(file_rel)[1].lower()
    mime = _IMG_MIME.get(ext, "image/png")
    with open(path, "rb") as fh:
        encoded = base64.b64encode(fh.read()).decode("ascii")
    return "data:%s;base64,%s" % (mime, encoded)


def render_svg_figure(topic):
    """Render an inline labelled SVG (fully offline, no data URI) with a caption."""
    svg = topic.get("svg")
    if not svg:
        return ""
    caption = topic.get("svg_caption", "")
    return ('<figure class="fig">%s<figcaption>%s'
            '<span class="credit">Original labelled schematic drawn for these notes.</span>'
            '</figcaption></figure>' % (svg, _html.escape(caption)))


def render_figures(topic):
    """Render each raster figure as an inline base64 <figure> (no external resources).

    All topics in this unit use inline SVG only, so figures lists are empty and this
    returns an empty string; the loop is kept so future raster figures stay offline.
    """
    figs = topic.get("figures")
    if not figs:
        return ""
    parts = []
    for fig in figs:
        if fig.get("credit"):
            credit = fig["credit"]
        else:
            credit = "Source: %s by %s. License: %s." % (
                fig.get("source", ""), fig.get("author", ""), fig.get("license", ""))
        parts.append(
            '<figure class="fig">'
            '<img alt="%s" src="%s">'
            '<figcaption>%s<span class="credit">%s</span></figcaption>'
            '</figure>'
            % (_html.escape(fig["alt"]), _data_uri(fig["file"]),
               _html.escape(fig["caption"]), _html.escape(credit)))
    return "\n".join(parts)


def render_topic(t):
    parts = ['<section class="card %s" id="%s">' % (t["color"], t["anchor"]),
             '<div class="card-head"><div class="sec-num">%d</div><h2>%s</h2></div>'
             % (t["number"], _html.escape(t["title"])),
             '<p class="intro">%s</p>' % render_inline_html(t["intro"])]
    parts.append(render_answer_flow(t["answer_flow"]))
    parts.append('<ol class="points">')
    for p in t["points"]:
        parts.append("<li>" + render_inline_html(p) + "</li>")
    parts.append("</ol>")
    svg_html = render_svg_figure(t)
    if svg_html:
        parts.append(svg_html)
    fig_html = render_figures(t)
    if fig_html:
        parts.append(fig_html)
    if t.get("definition"):
        parts.append(render_definition(t["definition"]))
    parts.append(render_flow(t["flow"]))
    parts.append(render_mnemonic(t["mnemonic"]))
    parts.append("</section>")
    return "\n".join(parts)


def render_html():
    toc_items = "\n".join(
        '<li><a href="#%s">%s</a></li>' % (t["anchor"], _html.escape(t["title"]))
        for t in CONTENT)
    guide_items = "\n".join("<li>" + render_inline_html(g) + "</li>" for g in HOW_TO_READ)
    topics_html = "\n".join(render_topic(t) for t in CONTENT)

    ref_books = "\n".join("<li>" + _html.escape(b) + "</li>" for b in REFERENCES["books"])
    ref_online_items = REFERENCES.get("online") or []
    ref_online = "\n".join("<li>" + _html.escape(o) + "</li>" for o in ref_online_items)
    ref_flagged = "\n".join("<li>" + _html.escape(f) + "</li>" for f in REFERENCES["flagged"])
    ref_figures = "\n".join("<li>" + _html.escape(f) + "</li>" for f in REFERENCES["figures"])

    online_section = ""
    if ref_online_items:
        online_section = ("<h3>Named online sources used</h3>\n<ul>\n%s\n</ul>" % ref_online)

    doc = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>{css}</style>
</head>
<body>

<header class="cover">
  <div class="label">M.Phil. Clinical Psychology &middot; Study Notes</div>
  <h1>{title}</h1>
  <div class="subtitle">{subtitle}</div>
  <div class="credit">{credit}</div>
</header>

<div class="wrap">

  <nav class="toc">
    <h2>Contents</h2>
    <ol>
      {toc}
    </ol>
  </nav>

  <section class="guide">
    <h2>How to read these notes</h2>
    <ul>
      {guide}
    </ul>
  </section>

  {topics}

  <section class="refs" id="references">
    <h2>References</h2>
    <h3>Books used</h3>
    <ul>
      {ref_books}
    </ul>
    {online_section}
    <h3>Figure sources</h3>
    <ul>
      {ref_figures}
    </ul>
    <h3>Secondary sources flagged (not accessed directly)</h3>
    <ul>
      {ref_flagged}
    </ul>
  </section>

</div>

<footer class="footer">
  <span class="title-line">{title}</span>
  <div class="credit">{credit}</div>
</footer>

</body>
</html>
""".format(title=_html.escape(TITLE), css=CSS, subtitle=_html.escape(SUBTITLE),
           credit=_html.escape(CREDIT), toc=toc_items, guide=guide_items,
           topics=topics_html, ref_books=ref_books, online_section=online_section,
           ref_figures=ref_figures, ref_flagged=ref_flagged)
    return doc


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(here, OUTPUT_HTML)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(render_html())
    print("Wrote", out_path)
    print("Topics:", len(CONTENT))
    for t in CONTENT:
        print("  - %s: %d points" % (t["title"], len(t["points"])))


if __name__ == "__main__":
    main()
