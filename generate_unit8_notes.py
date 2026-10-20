# -*- coding: utf-8 -*-
"""
Generator for the self-contained offline HTML study-notes file:
"Biological Foundations of Behavior - Unit VIII: Introduction to Neuropsychology".

This module defines ONE CONTENT data structure (a list of 7 topic dicts) that is
the single source of truth for the HTML deliverable rendered here.

Markup conventions used inside point / intro / flow text:
  **word**                 -> a bold memory-hook keyword (a few per sentence only)
  [[plain-language text]]  -> a distinct bracketed plain-language explanation of a
                              difficult term, shown in its own visible style.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 generate_unit8_notes.py
"""

import base64
import html as _html
import os
import re

TITLE = "Biological Foundations of Behavior \u2014 Unit VIII: Introduction to Neuropsychology"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"
SUBTITLE = ("Paper I: Biological Foundations of Behavior \u00b7 Part B (Neuropsychology) \u00b7 "
            "Exam-ready study notes")
OUTPUT_HTML = "Biological_Foundations_Unit8_Introduction_to_Neuropsychology.html"

# ---------------------------------------------------------------------------
# HOW TO READ THESE NOTES (short usage guide; NOT about how the notes were made)
# ---------------------------------------------------------------------------
HOW_TO_READ = [
    "Each topic opens with a one-line **intro** that frames the idea and then an "
    "**Answer flow** box, a small ladder of stages (definition, history, key concepts, "
    "evidence, significance, one-line summary); write your exam answer in exactly that order.",
    "The numbered points below the Answer flow follow those same stages, so reading them "
    "top to bottom gives you a ready-made **5-mark answer** rather than a list of scattered notes.",
    "Hard neuropsychological terms are kept in full, but each is followed by a plain-language "
    "meaning inside a shaded box like this [[a simple everyday explanation that sits right next "
    "to the difficult word]] so you keep the technical vocabulary while understanding it.",
    "Only a **few keywords** in each sentence are shown in bold; treat those bold words as "
    "memory hooks and rebuild the rest of the sentence around them while you revise.",
    "Every topic carries exactly one **flowchart** of boxes joined by arrows to show a sequence, "
    "and exactly one **mnemonic** whose letters each map to a point you must remember.",
    "Most topics carry a **labelled diagram**; use each figure as a visual map, matching its "
    "labels to the numbered points and the flowchart so you can both picture and describe it.",
    "Revise by first reading the Answer flow to fix the structure, then the flowchart and "
    "mnemonic to lock the sequence and key points, and finally the full sentences for exam wording.",
]

# ---------------------------------------------------------------------------
# Inline labelled SVG diagrams (fully offline, no data URI needed).
# Returned as literal <svg> markup embedded directly in the page.
# ---------------------------------------------------------------------------
SVG_TIMELINE = """
<svg viewBox="0 0 900 280" role="img" aria-label="Timeline of the rise of neuropsychology as a discipline"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="900" height="280" fill="#ffffff"/>
  <line x1="40" y1="150" x2="860" y2="150" stroke="#0e7c6b" stroke-width="4"/>
  <polygon points="860,150 846,142 846,158" fill="#0e7c6b"/>
  <g font-size="12" fill="#1a2535">
    <circle cx="90" cy="150" r="7" fill="#0f2744"/>
    <text x="90" y="110" text-anchor="middle" font-weight="bold">Egypt / Aristotle</text>
    <text x="90" y="178" text-anchor="middle">heart = seat of soul</text>
    <circle cx="230" cy="150" r="7" fill="#1d5fa6"/>
    <text x="230" y="110" text-anchor="middle" font-weight="bold">Hippocrates</text>
    <text x="230" y="178" text-anchor="middle">brain = seat of mind</text>
    <circle cx="370" cy="150" r="7" fill="#1d5fa6"/>
    <text x="370" y="128" text-anchor="middle" font-weight="bold">Willis / Gall</text>
    <text x="370" y="196" text-anchor="middle">lobes; phrenology</text>
    <circle cx="510" cy="150" r="7" fill="#7a1f3d"/>
    <text x="510" y="110" text-anchor="middle" font-weight="bold">Broca 1861</text>
    <text x="510" y="178" text-anchor="middle">speech in left frontal</text>
    <circle cx="640" cy="150" r="7" fill="#7a1f3d"/>
    <text x="640" y="128" text-anchor="middle" font-weight="bold">Wernicke 1874</text>
    <text x="640" y="196" text-anchor="middle">comprehension model</text>
    <circle cx="770" cy="150" r="7" fill="#4a2080"/>
    <text x="770" y="110" text-anchor="middle" font-weight="bold">Lashley</text>
    <text x="770" y="178" text-anchor="middle">mass action</text>
    <circle cx="845" cy="150" r="7" fill="#0e7c6b"/>
    <text x="845" y="128" text-anchor="middle" font-weight="bold">Kolb &amp;</text>
    <text x="845" y="210" text-anchor="middle">Whishaw 1980</text>
  </g>
</svg>
"""

SVG_LANGUAGE_LOOP = """
<svg viewBox="0 0 820 300" role="img" aria-label="Wernicke-Geschwind language loop in the left hemisphere"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="820" height="300" fill="#ffffff"/>
  <ellipse cx="410" cy="150" rx="380" ry="120" fill="#eef6ff" stroke="#bcd0e6"/>
  <g font-size="12" fill="#10243f">
    <rect x="70" y="190" width="150" height="46" rx="8" fill="#fff" stroke="#1d5fa6"/>
    <text x="145" y="210" text-anchor="middle" font-weight="bold">Primary auditory</text>
    <text x="145" y="226" text-anchor="middle">cortex (heard word)</text>
    <rect x="270" y="200" width="150" height="46" rx="8" fill="#fff" stroke="#0e7c6b"/>
    <text x="345" y="220" text-anchor="middle" font-weight="bold">Wernicke's area</text>
    <text x="345" y="236" text-anchor="middle">comprehension</text>
    <rect x="470" y="200" width="150" height="46" rx="8" fill="#fff" stroke="#8a5c00"/>
    <text x="545" y="220" text-anchor="middle" font-weight="bold">Arcuate</text>
    <text x="545" y="236" text-anchor="middle">fasciculus</text>
    <rect x="590" y="110" width="150" height="46" rx="8" fill="#fff" stroke="#7a1f3d"/>
    <text x="665" y="130" text-anchor="middle" font-weight="bold">Broca's area</text>
    <text x="665" y="146" text-anchor="middle">speech plan</text>
    <rect x="360" y="40" width="150" height="46" rx="8" fill="#fff" stroke="#4a2080"/>
    <text x="435" y="60" text-anchor="middle" font-weight="bold">Primary motor</text>
    <text x="435" y="76" text-anchor="middle">cortex (articulation)</text>
    <rect x="90" y="90" width="150" height="46" rx="8" fill="#fff" stroke="#1d5fa6"/>
    <text x="165" y="110" text-anchor="middle" font-weight="bold">Angular gyrus</text>
    <text x="165" y="126" text-anchor="middle">reading (visual code)</text>
  </g>
  <g stroke="#0f2744" stroke-width="2" fill="none" marker-end="url(#ar8)">
    <defs><marker id="ar8" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto">
      <path d="M0,0 L7,3 L0,6 Z" fill="#0f2744"/></marker></defs>
    <path d="M220,213 L270,218"/>
    <path d="M420,222 L470,222"/>
    <path d="M560,200 L620,156"/>
    <path d="M620,110 L510,78"/>
  </g>
</svg>
"""

SVG_IMAGING = """
<svg viewBox="0 0 860 300" role="img" aria-label="Comparison of brain imaging methods by what they record"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="860" height="300" fill="#ffffff"/>
  <g font-size="12">
    <text x="430" y="26" text-anchor="middle" font-weight="bold" font-size="14" fill="#0f2744">
      Electrical signal (fast, ms) &#8594; Blood/metabolism signal (slow, seconds)</text>
    <rect x="30" y="60" width="180" height="150" rx="10" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="120" y="86" text-anchor="middle" font-weight="bold" fill="#0e7c6b">EEG / ERP / MEG</text>
    <text x="120" y="112" text-anchor="middle" fill="#10243f">records electrical</text>
    <text x="120" y="130" text-anchor="middle" fill="#10243f">activity directly</text>
    <text x="120" y="156" text-anchor="middle" fill="#10243f">excellent TIME</text>
    <text x="120" y="174" text-anchor="middle" fill="#10243f">(N100, P300, N400)</text>
    <text x="120" y="196" text-anchor="middle" fill="#5a6a7e">poor location</text>
    <rect x="245" y="60" width="180" height="150" rx="10" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="335" y="86" text-anchor="middle" font-weight="bold" fill="#1d5fa6">PET / SPECT</text>
    <text x="335" y="112" text-anchor="middle" fill="#10243f">radioactive tracer</text>
    <text x="335" y="130" text-anchor="middle" fill="#10243f">blood flow / glucose</text>
    <text x="335" y="156" text-anchor="middle" fill="#10243f">subtraction method</text>
    <text x="335" y="178" text-anchor="middle" fill="#5a6a7e">good location,</text>
    <text x="335" y="196" text-anchor="middle" fill="#5a6a7e">slow (40-60 s)</text>
    <rect x="460" y="60" width="180" height="150" rx="10" fill="#f0eafa" stroke="#4a2080"/>
    <text x="550" y="86" text-anchor="middle" font-weight="bold" fill="#4a2080">fMRI (BOLD)</text>
    <text x="550" y="112" text-anchor="middle" fill="#10243f">blood-oxygen signal</text>
    <text x="550" y="130" text-anchor="middle" fill="#10243f">no radiation</text>
    <text x="550" y="156" text-anchor="middle" fill="#10243f">good LOCATION</text>
    <text x="550" y="178" text-anchor="middle" fill="#5a6a7e">slower than</text>
    <text x="550" y="196" text-anchor="middle" fill="#5a6a7e">electrical methods</text>
    <rect x="675" y="60" width="155" height="150" rx="10" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="752" y="86" text-anchor="middle" font-weight="bold" fill="#8a5c00">CT / MRI</text>
    <text x="752" y="112" text-anchor="middle" fill="#10243f">structure only</text>
    <text x="752" y="130" text-anchor="middle" fill="#10243f">(anatomy picture)</text>
    <text x="752" y="156" text-anchor="middle" fill="#10243f">lesion location</text>
    <text x="752" y="182" text-anchor="middle" fill="#5a6a7e">no function</text>
    <text x="430" y="250" text-anchor="middle" fill="#40506a" font-style="italic">
      Rule of thumb: electrical methods win on timing, blood-flow methods win on pin-pointing place.</text>
  </g>
</svg>
"""

SVG_STRUCTURE_FUNCTION = """
<svg viewBox="0 0 820 260" role="img" aria-label="Structure to function: levels of brain organisation"
     xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;font-family:Arial,Helvetica,sans-serif;">
  <rect x="0" y="0" width="820" height="260" fill="#ffffff"/>
  <g font-size="12" fill="#10243f">
    <rect x="20" y="100" width="120" height="60" rx="8" fill="#eef6ff" stroke="#1d5fa6"/>
    <text x="80" y="126" text-anchor="middle" font-weight="bold">Neurons</text>
    <text x="80" y="144" text-anchor="middle">cells</text>
    <rect x="190" y="100" width="130" height="60" rx="8" fill="#e3f5f1" stroke="#0e7c6b"/>
    <text x="255" y="126" text-anchor="middle" font-weight="bold">Local circuits</text>
    <text x="255" y="144" text-anchor="middle">columns / areas</text>
    <rect x="370" y="100" width="130" height="60" rx="8" fill="#fff3d6" stroke="#8a5c00"/>
    <text x="435" y="126" text-anchor="middle" font-weight="bold">Lobes</text>
    <text x="435" y="144" text-anchor="middle">gyri &amp; sulci</text>
    <rect x="550" y="100" width="130" height="60" rx="8" fill="#f0eafa" stroke="#4a2080"/>
    <text x="615" y="126" text-anchor="middle" font-weight="bold">Networks</text>
    <text x="615" y="144" text-anchor="middle">distributed</text>
    <rect x="700" y="100" width="100" height="60" rx="8" fill="#0f2744" stroke="#0f2744"/>
    <text x="750" y="126" text-anchor="middle" font-weight="bold" fill="#fff">Behaviour</text>
    <text x="750" y="144" text-anchor="middle" fill="#fff">&amp; cognition</text>
  </g>
  <g stroke="#0f2744" stroke-width="2" fill="none" marker-end="url(#ar8b)">
    <defs><marker id="ar8b" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto">
      <path d="M0,0 L7,3 L0,6 Z" fill="#0f2744"/></marker></defs>
    <path d="M140,130 L188,130"/>
    <path d="M320,130 L368,130"/>
    <path d="M500,130 L548,130"/>
    <path d="M680,130 L698,130"/>
  </g>
  <text x="410" y="210" text-anchor="middle" font-size="12" fill="#40506a" font-style="italic"
        font-family="Arial,Helvetica,sans-serif">
    Structure at each level constrains the function at the next: damage anywhere reshapes behaviour.</text>
</svg>
"""

# ---------------------------------------------------------------------------
# CONTENT = single source of truth (7 topics, in the required order)
# Each topic dict:
#   anchor, number, color, title, intro, answer_flow[list], points[list],
#   definition(dict or None), figures[list], svg(optional literal), flow(dict),
#   mnemonic(dict)
# ---------------------------------------------------------------------------
CONTENT = [
    {
        "anchor": "structure-function",
        "number": 1,
        "color": "blue",
        "title": "Relationship Between Structure and Function of the Brain",
        "intro": "Neuropsychology begins with one central claim, that the physical structure of the "
                 "brain and the mental functions it supports are two sides of one coin, so this topic "
                 "moves from what structure and function mean, through the levels that link cells to "
                 "whole networks, to why a clinician reads behaviour as a readout of the tissue beneath.",
        "answer_flow": [
            "Definition of structure and function",
            "Background / anatomical levels",
            "Key concepts (localization vs networks)",
            "Evidence / examples (Broca, lobes, connectivity)",
            "Clinical significance",
            "One-line summary",
        ],
        "points": [
            "The **structure-function relationship** is the idea that each part of the brain's physical "
            "**architecture** [[the way the tissue is built and wired]] gives rise to a particular mental "
            "or behavioural function, so anatomy and psychology are studied together, not apart.",
            "In practice this means a neuropsychologist treats observable behaviour as a **window** onto "
            "the underlying tissue, reading changes in sensing, moving, remembering or feeling as clues "
            "to where and how the brain's structure has been altered.",
            "The brain is organised in **levels** [[layers of scale from the smallest to the largest]], "
            "running from single neurons, through local circuits and cortical areas, up to the lobes and "
            "the distributed networks that together produce complex behaviour and cognition.",
            "At the largest scale the brain has three main parts, the **cerebrum**, the **cerebellum** "
            "that coordinates movement, posture and balance, and the **brainstem** that drives automatic "
            "functions such as breathing, heart rate and swallowing.",
            "The wrinkled surface of the cerebrum is the **cortex** [[the outer grey sheet of nerve cell "
            "bodies]], whose folds, a ridge called a **gyrus** and a groove called a **sulcus**, pack far "
            "more neurons into the skull by greatly increasing the surface area.",
            "Beneath the grey cortex lie long **axon** tracts that look pale and are called **white "
            "matter** [[bundles of insulated nerve fibres that wire one brain region to another]], so "
            "structure includes not just regions but the cabling that connects them.",
            "The cerebrum is split into a right and a left **hemisphere** joined by the **corpus "
            "callosum** [[the thick bridge of fibres that lets the two halves share information]], and "
            "each hemisphere mainly controls the opposite side of the body.",
            "Each hemisphere is divided into four **lobes**, the **frontal** lobe for movement, planning "
            "and personality, the **parietal** lobe for touch and spatial sense, the **temporal** lobe "
            "for hearing, language and memory, and the **occipital** lobe for vision.",
            "A classic structure-function demonstration is **Broca's area** in the left frontal lobe, "
            "where damage leaves a patient able to understand speech yet unable to produce fluent words, "
            "showing that one small region carries a specific job.",
            "Its partner is **Wernicke's area** in the left temporal lobe, where damage produces fluent "
            "but meaningless speech and poor comprehension, confirming that producing and understanding "
            "language depend on physically separate pieces of tissue.",
            "Yet no lobe works alone, because the lobes and the two hemispheres are tied together by "
            "**very complex relationships**, so most real functions emerge from several regions acting "
            "in concert rather than from one isolated spot.",
            "Modern work therefore stresses **brain connectivity** [[the pattern of anatomical and "
            "functional links between regions]], arguing that cognition arises from the dynamic "
            "interaction of distributed areas working as large-scale networks for vision, memory and "
            "attention.",
            "This connectivity comes in kinds, **structural connectivity** meaning the physical white "
            "matter wiring, and **functional connectivity** meaning which regions switch on together in "
            "time, and the two together explain how fixed structure supports flexible function.",
            "The link is not static, since the **structure-function coupling** [[how tightly wiring "
            "predicts joint activity]] strengthens with age and is disrupted in disorders, so this "
            "relationship itself can be measured and tracked in the clinic.",
            "A useful everyday analogy is a **city map**: the roads and buildings are the fixed structure, "
            "the daily traffic flowing along them is the function, and a blocked road (a lesion) forces "
            "the whole pattern of movement to change, just as brain damage reshapes behaviour.",
            "Clinically this coupling is powerful, because decoupling of structure and function has been "
            "linked to **schizophrenia**, epilepsy and major depression, so a disturbed relationship "
            "between wiring and activity can itself be a marker of psychological disorder.",
            "The practical payoff for a trainee is that symptoms become a **localising language**: a "
            "specific loss of naming, movement or spatial sense points toward a specific damaged "
            "structure, which is the first inference a neuropsychological assessment is built to make.",
            "In one line, the brain's structure and function are inseparable, so the clinician reads "
            "behaviour as the visible output of the tissue, its regions and the networks that join them "
            "underneath.",
        ],
        "definition": None,
        "svg": SVG_STRUCTURE_FUNCTION,
        "svg_caption": ("Figure 1a. Levels of brain organisation, from single neurons up through "
                        "circuits, lobes and distributed networks to behaviour, showing how structure "
                        "at each level constrains the function at the next."),
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig3_13_gyri_sulci.png",
                "alt": "Four-panel diagram of the cortex showing lateral and medial views of the gyri "
                       "and of the sulci, with the major lobes and folds of the cerebral cortex labelled",
                "caption": "Figure 1b. Gyri and sulci of the cortex \u2014 lateral and medial views of "
                           "the ridges (gyri) and grooves (sulci) that fold the cortex and define the "
                           "lobes, the physical structure behind the functions set out above.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 3.13.",
            },
        ],
        "flow": {
            "title": "From structure to behaviour",
            "steps": [
                ("Neurons", "cells", False),
                ("Circuits / areas", "local processing", False),
                ("Lobes", "gyri & sulci", False),
                ("Networks", "distributed links", False),
                ("Behaviour & cognition", "the output", True),
            ],
        },
        "mnemonic": {
            "word": "SHAPES",
            "line": "SHAPES links each piece of structure to the function it serves.",
            "items": [
                ("S", "**Structure** of tissue gives rise to function."),
                ("H", "**Hemispheres** joined by the corpus callosum."),
                ("A", "**Areas** and lobes each serve specific jobs."),
                ("P", "**Pathways** of white matter connect the regions."),
                ("E", "**Emergent** networks produce complex cognition."),
                ("S", "**Symptoms** localise where the structure is damaged."),
            ],
        },
    },
    {
        "anchor": "rise-discipline",
        "number": 2,
        "color": "teal",
        "title": "The Rise of Neuropsychology as a Distinct Discipline",
        "intro": "Neuropsychology did not appear overnight but grew over centuries from superstition to "
                 "science, so this topic traces its history from ancient Egypt and Aristotle, through "
                 "Hippocrates, Willis, Gall, Broca and Lashley, to the 1980 textbook that finally named "
                 "and defined the field as a discipline in its own right.",
        "answer_flow": [
            "Definition of the discipline",
            "Background / ancient roots",
            "Key figures in sequence",
            "Evidence / turning points",
            "Significance (field is named)",
            "One-line summary",
        ],
        "points": [
            "**Neuropsychology** is the branch of psychology concerned with how the brain and nervous "
            "system shape **cognition** [[thinking, memory and perception]] and behaviour, and especially "
            "with how injury or illness of the brain changes mental functioning.",
            "It is both an **experimental and a clinical** field, bridging classical neurology, which "
            "studies nervous-system disease, and classical psychology, which had grown largely divorced "
            "from the brain, by studying neurological patients to link mind with brain.",
            "The discipline is **relatively young**, and the first textbook that actually defined the "
            "field, Fundamentals of Human Neuropsychology, was published by **Kolb and Whishaw in 1980**, "
            "a landmark usually taken to mark its formal arrival.",
            "Its roots run far deeper, however, traceable back to the **Third Dynasty of ancient Egypt**, "
            "where the priest-physician Imhotep's writings described the brain, trauma and remedies in a "
            "surprisingly scientific way for future doctors.",
            "For centuries the brain was badly misjudged, since the Egyptians and then **Aristotle** held "
            "that the **heart** was the seat of the soul and thought, treating the brain as little more "
            "than a radiator that cooled the blood.",
            "The turning point in antiquity was **Hippocrates**, who moved the seat of the mind from the "
            "heart to the **brain**, writing that the brain exercises the greatest power in a person and "
            "so steering later science toward the correct organ.",
            "The philosopher **Rene Descartes** sharpened the debate with the **mind-body problem** "
            "[[the question of how an immaterial mind relates to the physical body]], proposing a "
            "dualism in which mind and body interact, which he located at the pineal gland.",
            "In the mid-seventeenth century **Thomas Willis** took a physiological approach, coined the "
            "very words **hemisphere and lobe**, was among the first to use neurology and psychology, "
            "and argued that higher brain structures handle complex functions.",
            "**Franz Joseph Gall** proposed that personality was tied to brain structure and invented "
            "**phrenology** [[the mistaken idea that bumps on the skull reveal character and ability]], "
            "which was wrong in detail yet rightly pushed the idea of localised cortical function.",
            "**Jean-Baptiste Bouillaud** extended Gall by arguing that distinct cortical regions have "
            "independent jobs, focusing on speech in the front of the brain and, unusually for the time, "
            "basing his case on over a hundred clinical case studies.",
            "The decisive figure was **Paul Broca**, who in 1861 showed from autopsy of aphasic patients "
            "that articulate speech depends on the **left hemisphere**, work widely seen as the moment "
            "neuropsychology took form as a respected discipline.",
            "In 1874 **Carl Wernicke** added a second language region and a **connectionist model** "
            "[[an account in which functions arise from signals passing between linked regions]], "
            "predicting conduction aphasia from damage to the fibres joining the two speech areas.",
            "The American **Karl Lashley** then studied memory by training rats and making lesions, "
            "proposing **mass action** [[the idea that loss depends on how much cortex is removed, not "
            "exactly where]] and equipotentiality, which we now read as early evidence of plasticity.",
            "Alongside these localisers ran a **debate** between strict localisation, which assigns each "
            "function to one spot, and more holistic views, and the modern field keeps both in tension "
            "by seeing functions as distributed yet regionally anchored.",
            "A helpful analogy is the writing of a **national constitution**: many thinkers over "
            "generations each add a clause, so that by the time the document is formally signed in 1980 "
            "it already rests on centuries of accumulated argument and evidence.",
            "Clinically this history still matters, because the very syndromes these pioneers described, "
            "**Broca's and Wernicke's aphasia**, remain core assessment findings a trainee must "
            "recognise and interpret at the bedside today.",
            "The discipline now splits into recognised **approaches**, experimental, clinical and "
            "cognitive neuropsychology, each a direct descendant of the lesion studies and case methods "
            "that these historical figures pioneered.",
            "In one line, neuropsychology rose slowly from the heart-centred guesses of antiquity to a "
            "brain-centred science, crystallising as a named discipline with Broca's localisation and "
            "the 1980 Kolb and Whishaw textbook.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Neuropsychology",
            "text": "\u201cNeuropsychology is a branch of psychology that is concerned with how the "
                    "brain and the rest of the nervous system influence a person's cognition and "
                    "behaviors... It is both an experimental and clinical field of psychology that aims "
                    "to understand how behavior and cognition are influenced by brain functioning and is "
                    "concerned with the diagnosis and treatment of behavioral and cognitive effects of "
                    "neurological disorders. The first textbook defining the field, Fundamentals of "
                    "Human Neuropsychology, was initially published by Kolb and Whishaw in 1980.\u201d",
            "cite": "\u201cNeuropsychology\u201d, Wikipedia (CC BY-SA), History section, as extracted "
                    "for this unit; the 1980 first-textbook fact is attributed there to Kolb & Whishaw.",
        },
        "svg": SVG_TIMELINE,
        "svg_caption": ("Figure 2a. A timeline of the rise of neuropsychology, from the heart-centred "
                        "views of Egypt and Aristotle, through Hippocrates, Willis and Gall, to Broca, "
                        "Wernicke, Lashley and the 1980 Kolb & Whishaw textbook that defined the field."),
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig1_6_wernicke_model.png",
                "alt": "Wernicke's 1874 language model drawn on a chimpanzee brain and a human brain, "
                       "labelling Broca's area, Wernicke's area and the arcuate fasciculus with the "
                       "sound to comprehension to articulation flow",
                "caption": "Figure 2b. Wernicke's 1874 model of language \u2014 drawn on (A) a chimpanzee "
                           "and (B) a human brain, showing Broca's area, Wernicke's area and the arcuate "
                           "fasciculus, an early connectionist account from the history above.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 1.6.",
            },
        ],
        "flow": {
            "title": "How the discipline grew over time",
            "steps": [
                ("Egypt / Aristotle", "heart = soul", False),
                ("Hippocrates", "brain = mind", False),
                ("Willis / Gall", "lobes; phrenology", False),
                ("Broca 1861 / Wernicke 1874", "speech localised", False),
                ("Lashley", "mass action", False),
                ("Kolb & Whishaw 1980", "field defined", True),
            ],
        },
        "mnemonic": {
            "word": "HE-WBLK",
            "line": "\"Hearts Explained, Willis Broca Lashley Kolb\" walks the key figures in order.",
            "items": [
                ("H", "**Hippocrates** moved the mind from heart to brain."),
                ("E", "**Egypt** and Aristotle first placed it in the heart."),
                ("W", "**Willis** coined 'hemisphere' and 'lobe'; **Gall** began phrenology."),
                ("B", "**Broca** (1861) localised speech to the left hemisphere."),
                ("L", "**Lashley** proposed mass action and equipotentiality."),
                ("K", "**Kolb** & Whishaw (1980) wrote the defining textbook."),
            ],
        },
    },
    {
        "anchor": "cerebral-organization",
        "number": 3,
        "color": "amber",
        "title": "The Logic of Cerebral Organization",
        "intro": "The cortex is not a random jumble but is laid out by a clear logic, so this topic "
                 "explains the principles that organise it, from primary sensory and motor zones through "
                 "association areas and cytoarchitectonic maps, to the hierarchy and symmetry that let a "
                 "clinician predict what a given region should do.",
        "answer_flow": [
            "Definition (what cerebral organization means)",
            "Background (why a logic is needed)",
            "Key organising principles",
            "Evidence / maps (Brodmann)",
            "Clinical significance",
            "One-line summary",
        ],
        "points": [
            "**Cerebral organization** means the orderly set of principles by which the cortex is laid "
            "out, so that position, cell structure and connections together predict what each region "
            "does rather than functions being scattered at random.",
            "The first principle is **division into lobes**, with the frontal, parietal, temporal and "
            "occipital lobes marked off by deep **fissures** [[the largest grooves that split the "
            "cortex]], giving a coarse map before any finer detail is added.",
            "A second principle is the split into **primary** areas that receive raw sensation or send "
            "raw movement commands, and **association** areas that integrate and interpret, so simple "
            "input is progressively turned into meaningful experience.",
            "Each sense has a **primary cortex**, vision in the occipital lobe, hearing in the temporal "
            "lobe, body sensation in the parietal **sensory strip**, while the frontal **motor strip** "
            "issues movement, giving the cortex an input-processing-output layout.",
            "A third principle is **topographic mapping** [[neighbouring parts of the body or visual "
            "field map onto neighbouring patches of cortex]], so the sensory and motor strips carry an "
            "orderly body map rather than a shuffled one.",
            "A fourth principle is **hierarchy**, drawn out by Willis long ago, in which lower structures "
            "handle automatic reactions and higher cortical structures handle complex functions, so "
            "processing climbs from reflex toward reasoning.",
            "The deep structures obey this logic too, with the **thalamus** acting as a relay station "
            "for almost all information travelling to the cortex and the **limbic system** handling "
            "emotion and memory, so organisation runs inward as well as across the surface.",
            "The strongest evidence for an orderly cortex came from **cytoarchitectonics** [[mapping the "
            "cortex by the microscopic arrangement of its cell layers]], the method by which Korbinian "
            "Brodmann divided the cortex into numbered areas.",
            "**Brodmann (1868-1918)** established the comparative cytoarchitectonics of the mammalian "
            "cortex and produced his famous map of about **47 areas**, still the most widely used scheme "
            "for naming cortical regions by number.",
            "Others drew finer maps, with **Vogt and Vogt (1919)** describing over two hundred areas and "
            "**von Economo (1929)** about a hundred and nine, showing that the cortex can be parcellated "
            "ever more finely yet still lawfully.",
            "Experimental stimulation confirmed the map, as **Fritsch and Hitzig (1870)** found the motor "
            "cortex, **Ferrier (1873)** the auditory cortex, and Penfield and Rasmussen later outlined "
            "the motor and sensory **homunculus** [[the distorted body map on the strips]].",
            "A further principle is **two near-mirror hemispheres** that are organised alike in broad "
            "layout yet differ in detail, so the logic includes a left-right symmetry that is real but "
            "not perfect.",
            "The modern refinement is that regions are also organised into **networks**, so cerebral "
            "organisation is not only a patchwork of areas but a wiring diagram in which distributed "
            "regions cooperate to produce a single function.",
            "A clear everyday analogy is a **well-planned office building**: the ground floor handles "
            "incoming deliveries (primary sensory areas), workshops process them (association areas), "
            "the top floor makes decisions (frontal cortex), and lifts and corridors are the connecting "
            "tracts.",
            "Clinically this logic is what makes **lesion prediction** possible, because knowing the "
            "orderly map lets a neuropsychologist forecast that occipital damage harms vision while "
            "parietal damage harms touch and spatial sense.",
            "It also guides **neurosurgery**, since mapping eloquent or 'no-go' areas before operating "
            "depends entirely on trusting that cortical functions sit in lawful, predictable places.",
            "The same orderly map underlies standard **assessment**, because tests are chosen to probe "
            "specific regions, and a patient's profile of spared and impaired scores is read back onto "
            "this cortical layout.",
            "In one line, the cortex is organised by lawful principles of lobes, primary-versus-"
            "association zones, topographic maps, hierarchy and networks, which together let the clinician "
            "predict function from position.",
        ],
        "definition": {
            "label": "Sourced definition \u2014 Cerebral localization (mapping sense)",
            "text": "\u201cThe mapping of the cerebral cortex into areas, and the correlation of these "
                    "areas with cerebral function.\u201d",
            "cite": "Cerebral Localization teaching presentation (H. Mishra), Definition slide, as "
                    "extracted for this unit.",
        },
        "svg": None,
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig10_7_brodmann_map.png",
                "alt": "Brodmann's cytoarchitectonic map of the cortex showing numbered areas in clean "
                       "lateral and medial views of the human brain",
                "caption": "Figure 3a. Brodmann's map of the cortex \u2014 clean lateral and medial views "
                           "of the numbered cytoarchitectonic areas, the orderly cell-based parcellation "
                           "that is the strongest evidence for a logic of cerebral organization.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 10.7.",
            },
        ],
        "flow": {
            "title": "Principles that organise the cortex",
            "steps": [
                ("Lobes & fissures", "coarse map", False),
                ("Primary areas", "raw in / out", False),
                ("Association areas", "integrate", False),
                ("Cytoarchitectonic map", "Brodmann areas", False),
                ("Networks", "regions cooperate", True),
            ],
        },
        "mnemonic": {
            "word": "LAMPH",
            "line": "LAMPH lights up the logic of how the cortex is organised.",
            "items": [
                ("L", "**Lobes** divided by deep fissures give the coarse map."),
                ("A", "**Association** areas integrate what primary areas receive."),
                ("M", "**Maps** are topographic (orderly body and field maps)."),
                ("P", "**Parcellation** by Brodmann numbers areas by cell structure."),
                ("H", "**Hierarchy** and networks build complex function from simple."),
            ],
        },
    },
    {
        "anchor": "localization",
        "number": 4,
        "color": "rose",
        "title": "Localization of Functions",
        "intro": "Localization is the claim that particular mental functions live in particular places "
                 "in the brain, so this topic defines it, traces its evidence from phrenology through "
                 "Broca and Brodmann, surveys the lobe syndromes it predicts, and weighs it against the "
                 "fact that many functions are in truth distributed.",
        "answer_flow": [
            "Definition of localization",
            "Background / history of the idea",
            "Key evidence (speech, maps, stimulation)",
            "Examples (lobe syndromes, aphasias)",
            "Significance and the distributed caveat",
            "One-line summary",
        ],
        "points": [
            "**Localization of function** is the principle that a specific cognitive or behavioural "
            "function is carried out by a specific, identifiable region of the brain, so that damage "
            "there produces a predictable and characteristic deficit.",
            "Formally it also names a clinical task, the **diagnosis of the location** of a brain lesion "
            "from the signs and symptoms a patient shows, so localization is both a theory of the brain "
            "and a method of bedside reasoning.",
            "The idea has a long history, beginning in distorted form with the **phrenology** of Gall "
            "and Spurzheim around 1781, who wrongly read skull bumps but rightly insisted that different "
            "faculties sit in different parts of the brain.",
            "The decisive evidence was **Paul Broca's** 1861 case of the patient 'Tan', whose left "
            "frontal lesion abolished fluent speech, giving convincing proof of **speech laterality** "
            "[[the siting of language on one side]] and a localised speech-production centre.",
            "**Carl Wernicke (1874)** added that a left temporal lesion disturbs **comprehension**, "
            "built a connectionist model of language, and predicted **conduction aphasia** from damage "
            "to the fibres linking the two language areas.",
            "Experimental physiology confirmed localisation, as **Fritsch and Hitzig (1870)** located "
            "the motor cortex, **von Gudden** the visual cortex and **Ferrier (1873)** the auditory "
            "cortex, each by stimulating or lesioning a specific region.",
            "**Brodmann's** cytoarchitectonic map then gave localisation an anatomical backbone, numbering "
            "areas such as the **primary motor area (4)**, somatosensory areas (3,1,2), primary visual "
            "area (17) and primary auditory area (41,42).",
            "Direct cortical mapping sealed the case, with **Harvey Cushing** stimulating the cortex of "
            "conscious patients and **Penfield and Rasmussen** outlining the motor and sensory "
            "**homunculus** that shows how much cortex each body part commands.",
            "Localisation predicts clear **frontal lobe signs**, including impaired planning, emotional "
            "lability, disinhibition and the executive syndromes, famously illustrated by the personality "
            "change of **Phineas Gage** after a frontal injury.",
            "It predicts distinct **parietal lobe signs** such as **astereognosis** [[inability to "
            "recognise objects by touch]], neglect of one side of space and constructional difficulty, "
            "especially after right-hemisphere damage.",
            "It predicts **temporal lobe signs** including auditory agnosia, memory disturbance and, with "
            "left-sided damage, impaired discrimination of words, confirming that hearing, language and "
            "memory are regionally anchored.",
            "It predicts **occipital lobe signs** such as **hemianopia** [[loss of half the visual "
            "field]] and the visual agnosias, so that losing a defined patch of cortex removes a defined "
            "slice of visual experience.",
            "The aphasias map neatly onto areas, with **receptive (Wernicke) aphasia** around area 22 "
            "giving poor comprehension but fluent speech, and **motor (Broca) aphasia** around areas 44 "
            "and 45 giving good comprehension but halting, telegraphic speech.",
            "Even reading has localised breakdowns, since damage near the left **angular gyrus** can "
            "cause an inability to read with intact spoken language, a dissociation that pins a complex "
            "skill onto a small region.",
            "The strong caveat, however, is that function is often **distributed**, because Lashley's "
            "mass action and modern network views show many tasks need several cooperating areas, so "
            "pure, single-spot localisation is the exception rather than the rule.",
            "A clear everyday analogy is a **kitchen in a restaurant**: there is a definite station for "
            "chopping, one for the stove and one for plating (localised jobs), yet producing a full meal "
            "still needs all the stations working together (a distributed network).",
            "Clinically, localisation underpins **lesion-deficit reasoning**, the everyday inference by "
            "which a neuropsychologist moves from a pattern of test failures to a best guess about which "
            "brain region has been harmed.",
            "In one line, localization says specific functions sit in specific places, proven by speech "
            "areas, cytoarchitectonic maps and stimulation, but tempered by the truth that many functions "
            "are distributed across cooperating regions.",
        ],
        "definition": {
            "label": "Sourced definition \u2014 Cerebral localization (lesion-diagnosis sense)",
            "text": "\u201cThe diagnosis of the location in the cerebrum of a brain lesion, made either "
                    "from the signs and symptoms manifested by the patient or from an investigation "
                    "modality.\u201d",
            "cite": "Cerebral Localization teaching presentation (H. Mishra), Definition slide, as "
                    "extracted for this unit.",
        },
        "svg": None,
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig3_16_brodmann_areas.png",
                "alt": "Brodmann's areas shown as lateral and medial cytoarchitectonic maps together "
                       "with a reference table linking function, map code and Brodmann area number",
                "caption": "Figure 4a. Brodmann's areas with the function / map-code / area-number "
                           "reference table \u2014 the anatomical backbone that lets specific functions "
                           "be localised to specific numbered regions of cortex.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 3.16.",
            },
            {
                "file": "assets/unit8_figures/from_books/kolb_fig1_6_wernicke_model.png",
                "alt": "Wernicke's language model on chimpanzee and human brains showing Broca's area, "
                       "Wernicke's area and the arcuate fasciculus",
                "caption": "Figure 4b. Wernicke's model locating speech production (Broca's area) and "
                           "comprehension (Wernicke's area) in separate left-hemisphere regions joined "
                           "by the arcuate fasciculus \u2014 a textbook case of localization.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 1.6.",
            },
        ],
        "flow": {
            "title": "How a function is localised to a region",
            "steps": [
                ("Specific lesion", "injury to one area", False),
                ("Specific deficit", "one function lost", False),
                ("Map to Brodmann area", "e.g. 44/45 speech", False),
                ("Confirm by stimulation", "homunculus", False),
                ("Predict syndrome", "lobe signs", True),
            ],
        },
        "mnemonic": {
            "word": "BROCA",
            "line": "BROCA holds the pillars of localization of function.",
            "items": [
                ("B", "**Broca** localised speech production to the left frontal lobe."),
                ("R", "**Regions** differ by cell structure (Brodmann areas)."),
                ("O", "**Occipital, parietal, temporal** lobes give distinct signs."),
                ("C", "**Comprehension** sits in Wernicke's area (area 22)."),
                ("A", "**And yet** many functions are distributed, not single-spot."),
            ],
        },
    },
    {
        "anchor": "lateralization",
        "number": 5,
        "color": "purple",
        "title": "Lateralization of Functions",
        "intro": "Lateralization is the fact that the two hemispheres are not identical twins but "
                 "specialists, so this topic defines cerebral dominance, presents the split-brain "
                 "evidence, contrasts the left and right hemispheres, and shows why a clinician must "
                 "always ask not just where damage is but on which side.",
        "answer_flow": [
            "Definition of lateralization / dominance",
            "Background (Dax, Broca, callosum)",
            "Key evidence (split-brain studies)",
            "Left vs right specialisations",
            "Clinical significance",
            "One-line summary",
        ],
        "points": [
            "**Lateralization of function**, also called **cerebral dominance** or asymmetry, is the "
            "fact that the two hemispheres are specialised for different functions rather than being "
            "mirror-image duplicates of each other.",
            "The two halves are joined and normally kept in step by the **corpus callosum** [[the great "
            "bridge of fibres carrying messages between the hemispheres]], so each side can share what "
            "it knows with the other.",
            "The first hint came in **1836**, when **Dax** reported that none of his roughly forty "
            "patients with speech problems had damage restricted to the right hemisphere, pointing to a "
            "left-sided speech specialisation.",
            "Twenty-five years later **Broca's** autopsies of aphasic patients confirmed that diffuse "
            "**left-hemisphere** damage, centred in the inferior frontal lobe, abolished speech, firmly "
            "establishing the left side as dominant for language.",
            "The **left hemisphere** is therefore the usually **dominant hemisphere** for language, "
            "speech and writing, for analytical and mathematical skills, and for the **temporal "
            "sequencing** [[ordering events in time]] of stimuli.",
            "The **right hemisphere** is specialised instead for **spatial perception** of "
            "three-dimensional material, for singing and for playing musical instruments, so it is far "
            "from the 'minor' hemisphere it was once thought to be.",
            "Handedness matters, since the left hemisphere is dominant for language in about **ninety-"
            "two percent** of people, while in roughly a third of left-handers speech may sit on the "
            "right, a fact checked before brain surgery.",
            "The clinching evidence came from **split-brain** studies, which began when **Myers and "
            "Sperry (1953)** cut the corpus callosum in cats and revealed the role of the brain's "
            "largest commissure in sharing learning between the hemispheres.",
            "In humans, **corpus callosotomy** [[surgical cutting of the callosum, once used to control "
            "severe epilepsy]] creates a split-brain patient in whom the two hemispheres can no longer "
            "directly exchange information.",
            "When a colour is flashed only in the **left visual field**, which projects to the mute "
            "right hemisphere, a split-brain patient's talking left hemisphere must guess the name, "
            "exposing that each hemisphere now knows different things.",
            "Strikingly, such a patient can **learn two things at once**, reaching with each hand into a "
            "different bag to pull out two different objects shown separately to each visual field, "
            "because the hemispheres act independently.",
            "The brain also shows **anatomical asymmetries**, with the **planum temporale** and frontal "
            "operculum, both language-related, larger in the left hemisphere, while Heschl's gyrus is "
            "larger on the right, so structure mirrors the functional split.",
            "Lateralisation is finely wired within language itself, for damage to the **arcuate "
            "fasciculus** joining Broca's and Wernicke's areas leaves comprehension and speech intact "
            "but makes the patient unable to repeat words just heard.",
            "Reading too is lateralised, since left angular-gyrus damage can cause **surface dyslexia**, "
            "losing the memory-based reading of irregular words, or **deep dyslexia**, losing the "
            "rule-based sounding-out of new or non-words.",
            "A clear everyday analogy is a **two-person workshop with a shared doorway**: one worker "
            "specialises in words and logic, the other in space and music, and the corpus callosum is "
            "the doorway through which they pass work; seal the door and each carries on alone.",
            "Clinically, lateralisation means a neuropsychologist must always ask **which side** is "
            "damaged, because a left-hemisphere stroke threatens language while a matching right-sided "
            "stroke threatens spatial awareness and may cause neglect.",
            "It also drives the pre-surgical **Wada procedure** [[injecting a short-acting anaesthetic "
            "into one hemisphere to find which side controls speech]], so that dominant language cortex "
            "can be protected during an operation.",
            "In one line, lateralization means the left and right hemispheres are specialists linked by "
            "the corpus callosum, proven by split-brain studies, so the side of a lesion matters as much "
            "as its site.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Cerebral dominance (lateralization)",
            "text": "\u201cCerebral Dominance (Lateralization, Asymmetry): Dominant Hemisphere (LEFT) \u2014 "
                    "language (speech, writing), analytical and mathematical skills, temporal sequencing "
                    "of stimuli. Non-dominant Hemisphere (RIGHT) \u2014 spatial perception (3D), singing, "
                    "playing musical instruments.\u201d",
            "cite": "Cerebral Localization teaching presentation (H. Mishra), Cerebral Dominance slide, "
                    "as extracted for this unit.",
        },
        "svg": None,
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig11_5_visual_pathways.png",
                "alt": "Diagram showing that each visual field is represented in the opposite visual "
                       "cortex via the optic chiasm, the crossed visual pathways",
                "caption": "Figure 5a. Visual pathways are crossed \u2014 each visual field projects to "
                           "the opposite hemisphere via the optic chiasm, the wiring that lets split-"
                           "brain testing show each hemisphere's separate knowledge.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 11.5.",
            },
            {
                "file": "assets/unit8_figures/from_books/kolb_fig11_6_corpus_callosum.png",
                "alt": "The split-brain key demonstration: a word in the left visual field reaches the "
                       "right hemisphere first, then crosses the corpus callosum to the left hemisphere",
                "caption": "Figure 5b. Connection between the two sides \u2014 a word seen in the left "
                           "visual field goes first to the right hemisphere and crosses the corpus "
                           "callosum to the left; cutting the callosum blocks the transfer.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 11.6.",
            },
            {
                "file": "assets/unit8_figures/from_books/kolb_fig11_13_motor_asymmetry.png",
                "alt": "Successive video frames showing the right side of the mouth opening wider during "
                       "production of the syllable ma, an example of motor asymmetry",
                "caption": "Figure 5c. Motor asymmetry \u2014 successive frames show the right side of "
                           "the mouth opening wider while speaking, a subtle behavioural sign of "
                           "left-hemisphere control of speech output.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 11.13.",
            },
        ],
        "flow": {
            "title": "What split-brain testing reveals",
            "steps": [
                ("Stimulus to one visual field", "e.g. LVF", False),
                ("Reaches opposite hemisphere", "crossed pathway", False),
                ("Corpus callosum cut", "no transfer", False),
                ("Each hemisphere acts alone", "different knowledge", False),
                ("Reveals lateralization", "L vs R specialities", True),
            ],
        },
        "mnemonic": {
            "word": "L-SAT",
            "line": "The dominant Left does \"L-SAT\": Language, Sequencing, Analysis, Talk/write.",
            "items": [
                ("L", "**Language**, speech and writing sit in the left hemisphere."),
                ("S", "**Sequencing** of stimuli in time is a left-hemisphere skill."),
                ("A", "**Analytical** and mathematical reasoning are left-sided."),
                ("T", "**Right** instead handles space, singing and music (the other side)."),
            ],
        },
    },
    {
        "anchor": "clinical-methods",
        "number": 6,
        "color": "navy",
        "title": "Approaches and Methodologies of Clinical Neuropsychologists",
        "intro": "Clinical neuropsychology takes the science of brain-behaviour relationships to the "
                 "bedside, so this topic defines the approach, surveys the standardised tests, "
                 "electrophysiological and imaging tools the clinician uses, and shows how these combine "
                 "to assess, diagnose and rehabilitate real patients.",
        "answer_flow": [
            "Definition of clinical neuropsychology",
            "Background / aim (assessment & rehab)",
            "Key methods (tests, EEG, imaging)",
            "Examples (test batteries, Wada)",
            "Clinical significance",
            "One-line summary",
        ],
        "points": [
            "**Clinical neuropsychology** is the application of neuropsychological knowledge to the "
            "assessment, management and rehabilitation of people who have suffered illness or injury, "
            "particularly to the brain, that has caused **neurocognitive problems**.",
            "Its distinctive contribution is to bring a **psychological viewpoint** to such cases, asking "
            "how brain illness both affects and is affected by emotional and psychological factors, not "
            "just the physical lesion itself.",
            "A key clinical skill is **differential diagnosis**, judging whether a difficulty arises from "
            "**brain pathology** [[physical disease of the brain tissue]] or from a potentially "
            "reversible emotional cause, or from both acting together.",
            "A classic example is distinguishing types of memory loss, where offering a **categorical "
            "cue** (for instance, telling the patient the item is a fruit) that restores naming points "
            "to a reversible vascular cause rather than a fixed dementia.",
            "The core method is the **standardized neuropsychological test**, a task designed so that "
            "performance links to a specific neurocognitive process and that has been given to reference "
            "groups to create **normative data** [[average scores from a comparison sample]].",
            "Named examples the trainee should know include the **Wechsler Adult Intelligence Scale** and "
            "Wechsler Memory Scale, the **Wisconsin Card Sorting Test** of executive function, the Boston "
            "Naming Test and the Benton Visual Retention Test.",
            "The clinician reads a patient's **profile** of spared and impaired scores against the norms, "
            "turning a scatter of results into an inference about which cognitive systems, and so which "
            "brain regions, are working or failing.",
            "Bedside **clinical methods** also look for lobar signs and syndromes, screening for the "
            "**aphasias** of language, the **apraxias** [[loss of skilled, planned movement]], the "
            "**agnosias** [[failure to recognise despite intact senses]] and visual field loss.",
            "Clinical neuropsychologists use **electrophysiological** tools such as "
            "**electroencephalography (EEG)**, which records the brain's electrical activity through "
            "scalp electrodes and is valuable in epilepsy and in assessing overall brain state.",
            "They draw on **radiological** methods, distinguishing **structural** imaging such as CT and "
            "MRI that pictures the anatomy and any lesion, from **functional** imaging that shows which "
            "regions are active during a task.",
            "**PET** [[positron emission tomography, which tracks blood flow or glucose use with a "
            "radioactive tracer]] images function, and is often read with the **subtraction method** "
            "that removes resting activity to reveal the regions a task specifically switches on.",
            "**fMRI** [[functional magnetic resonance imaging, which detects the blood-oxygen BOLD "
            "signal]] shows activation without radiation and is used to map language, memory and emotion, "
            "complementing the structural picture from ordinary MRI.",
            "For surgical planning the clinician contributes **intra-operative** and specialised methods, "
            "including cortical stimulation mapping and the **Wada procedure** that anaesthetises one "
            "hemisphere to find where language and memory live before an operation.",
            "The guiding logic is **functional localization** combined with careful behavioural testing, "
            "so that a specific deficit after injury to a specific area implicates that region, exactly "
            "the lesion-deficit reasoning the pioneers established.",
            "A clear everyday analogy is a **car mechanic's full service**: road-testing the car is the "
            "cognitive testing, the diagnostic computer readout is the EEG or scan, and only by combining "
            "both does the mechanic locate the fault and plan the repair (rehabilitation).",
            "Settings matter clinically, since these professionals work in hospitals within "
            "**interdisciplinary medical teams**, in private practice, and in **medico-legal** work where "
            "an expert opinion on brain function is required.",
            "The end goal is not merely a label but **management and rehabilitation**, using the "
            "assessment to design therapy, support recovery and harness the brain's plasticity so the "
            "patient regains as much function as possible.",
            "In one line, clinical neuropsychology blends standardized testing with EEG and structural "
            "and functional imaging to assess, localise, diagnose and rehabilitate brain-injured patients "
            "with a psychological eye.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Clinical neuropsychology",
            "text": "\u201cClinical neuropsychology is the application of neuropsychological knowledge to "
                    "the assessment... management, and rehabilitation of people who have suffered illness "
                    "or injury (particularly to the brain) which has caused neurocognitive problems. In "
                    "particular they bring a psychological viewpoint to treatment, to understand how such "
                    "illness and injury may affect and be affected by psychological factors.\u201d",
            "cite": "\u201cNeuropsychology\u201d, Wikipedia (CC BY-SA), Approaches section, as extracted "
                    "for this unit.",
        },
        "svg": SVG_IMAGING,
        "svg_caption": ("Figure 6a. The main investigative methods compared \u2014 electrical methods "
                        "(EEG, ERP, MEG) excel at timing, blood-flow and metabolic methods (PET, SPECT, "
                        "fMRI) excel at pinpointing place, while CT and MRI show structure alone."),
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig11_6_corpus_callosum.png",
                "alt": "Split-brain testing apparatus showing a word presented to one visual field and "
                       "the hemisphere it reaches, the kind of controlled clinical test used to probe "
                       "hemispheric function",
                "caption": "Figure 6b. A controlled split-brain testing arrangement \u2014 the kind of "
                           "carefully designed behavioural method a clinical neuropsychologist uses to "
                           "isolate what one hemisphere can do on its own.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 11.6.",
            },
        ],
        "flow": {
            "title": "A clinical neuropsychology assessment pathway",
            "steps": [
                ("Referral & history", "the question", False),
                ("Standardised tests", "vs normative data", False),
                ("EEG / imaging", "electrical & structural", False),
                ("Localise & diagnose", "lesion-deficit logic", False),
                ("Manage & rehabilitate", "plan therapy", True),
            ],
        },
        "mnemonic": {
            "word": "TESTER",
            "line": "A clinical neuropsychologist is a careful TESTER of the injured brain.",
            "items": [
                ("T", "**Tests**, standardised and compared against normative data."),
                ("E", "**EEG** and electrophysiology record electrical activity."),
                ("S", "**Structural** imaging (CT, MRI) shows the lesion."),
                ("T", "**Tomography** of function (PET, fMRI) shows activity."),
                ("E", "**Examine** lobar signs: aphasia, apraxia, agnosia."),
                ("R", "**Rehabilitate**, the goal of assessment and management."),
            ],
        },
    },
    {
        "anchor": "cognitive-methods",
        "number": 7,
        "color": "blue",
        "title": "Approaches and Methodologies of Cognitive Neuropsychologists",
        "intro": "Cognitive neuropsychology uses brain-damaged patients to reveal the hidden "
                 "architecture of the normal mind, so this topic defines the approach, explains its "
                 "logic of dissociations and models, and surveys the ERP, functional imaging and "
                 "computational tools it uses to connect cognition to the brain.",
        "answer_flow": [
            "Definition of cognitive neuropsychology",
            "Background (distilled from experimental + clinical)",
            "Key logic (models, dissociations)",
            "Methods (ERP, fMRI, connectionism)",
            "Significance",
            "One-line summary",
        ],
        "points": [
            "**Cognitive neuropsychology** is a relatively new approach that has emerged as a "
            "**distillation** [[a blend of the best parts]] of experimental and clinical "
            "neuropsychology, aiming to understand the mind and brain by studying brain-injured people.",
            "Its central aim is to use the **pattern of breakdown** after brain injury to reveal the "
            "normal architecture of cognition, so patients act as natural experiments that expose how a "
            "healthy mind is organised.",
            "One guiding model is **functional localization**, the principle that if a specific cognitive "
            "problem follows injury to a specific area, that part of the brain is probably involved in "
            "that cognitive function.",
            "The approach relies heavily on **dissociations** [[cases where one ability is lost while "
            "another is spared]], as when a reader loses irregular-word reading in surface dyslexia but "
            "keeps rule-based sounding-out, implying two separate reading routes.",
            "A famous pair is **surface versus deep dyslexia**, where surface dyslexics can sound out "
            "non-words like 'wug' but stumble on irregular words, while deep dyslexics manage irregular "
            "words yet cannot pronounce non-words, mapping two distinct mechanisms.",
            "The field is careful about its models, recognising that the mind-brain link may not be "
            "simple one-to-one, so alternatives such as **parallel processing** [[many operations "
            "running at once rather than in a single chain]] may better explain some functions.",
            "A further, model-light strategy studies how the **pattern of errors** a patient makes "
            "constrains theories of mental representation and process, letting cognition be mapped even "
            "without pinning each step to a precise brain location.",
            "A major tool is the **event-related potential (ERP)** [[a tiny brainwave averaged from many "
            "repeats of a stimulus to pull the signal out of background noise]], which is an "
            "electrophysiological measure with excellent timing.",
            "ERP waves are named by polarity and latency, such as the **N100, P200 and P300**, and the "
            "**P300** in particular is widely studied as an index of attention and the updating of "
            "working memory during a task.",
            "The **N400** was the first ERP tied to **language processing**, growing larger when the "
            "final word of a sentence is semantically surprising or inappropriate, so its amplitude "
            "tracks the difficulty of making meaning.",
            "ERP has been used to study memory, language and attention, and to index the cognitive "
            "deterioration seen in **neuropsychiatric and degenerative disorders** such as schizophrenia, "
            "depression and Alzheimer's disease.",
            "The field also uses **functional neuroimaging**, taking PET or fMRI readings while a person "
            "performs a task to relate the activation of particular brain areas to that specific "
            "cognitive operation.",
            "**PET** studies apply the **subtraction method**, subtracting resting blood flow from "
            "task blood flow so the remaining image highlights only the regions a cognitive task "
            "genuinely switches on, giving good spatial but poor temporal resolution.",
            "**fMRI** detects the **BOLD signal** [[a rise in blood oxygen near active neurons]] without "
            "radiation, and the growth of cognitive testing inside fMRI has had a notable influence on "
            "studying brain-behaviour relations in the intact brain.",
            "**Connectionism** adds a modelling method, using artificial **neural networks** [[computer "
            "models of simplified neurons]] that are trained on a task and then deliberately 'lesioned' "
            "to see whether their breakdown mimics that of brain-damaged patients.",
            "A clear everyday analogy is understanding a **vending machine by its faults**: if pressing "
            "one button fails while others work, you infer the machine has separate internal circuits, "
            "just as a selective cognitive deficit reveals a separate mental module.",
            "A related offshoot is **cognitive neuropsychiatry**, which turns the same logic onto "
            "psychiatric illness, studying disorders like delusions to understand the normal workings of "
            "belief, perception and the mind.",
            "In one line, cognitive neuropsychology reads the selective breakdowns of brain-injured "
            "patients, backed by ERP, functional imaging and connectionist models, to map the hidden "
            "architecture of the normal mind.",
        ],
        "definition": {
            "label": "Sourced characterisation \u2014 Cognitive neuropsychology",
            "text": "\u201cCognitive neuropsychology... has emerged as a distillation of the "
                    "complementary approaches of both experimental and clinical neuropsychology. It "
                    "seeks to understand the mind and brain by studying people who have suffered brain "
                    "injury or neurological illness. One model of neuropsychological functioning is "
                    "known as functional localization.\u201d",
            "cite": "\u201cNeuropsychology\u201d, Wikipedia (CC BY-SA), Approaches section, as extracted "
                    "for this unit.",
        },
        "svg": SVG_LANGUAGE_LOOP,
        "svg_caption": ("Figure 7a. The Wernicke-Geschwind language loop \u2014 a cognitive model of how "
                        "a heard or read word flows through auditory cortex, Wernicke's area, the arcuate "
                        "fasciculus, Broca's area and motor cortex to produce speech."),
        "figures": [
            {
                "file": "assets/unit8_figures/from_books/kolb_fig1_6_wernicke_model.png",
                "alt": "Wernicke's connectionist language model on chimpanzee and human brains, a "
                       "forerunner of modern cognitive models of information flow",
                "caption": "Figure 7b. Wernicke's connectionist model of language \u2014 an early example "
                           "of the cognitive-neuropsychology style of explaining a function as "
                           "information flowing between linked processing stages.",
                "credit": "From Kolb & Whishaw, Fundamentals of Human Neuropsychology, Figure 1.6.",
            },
        ],
        "flow": {
            "title": "The reasoning of cognitive neuropsychology",
            "steps": [
                ("Study brain-injured patient", "natural experiment", False),
                ("Find dissociation", "one ability lost", False),
                ("Infer a module", "separate mechanism", False),
                ("Test with ERP / fMRI", "N400, BOLD", False),
                ("Model it (connectionism)", "lesion the network", False),
                ("Map the normal mind", "the goal", True),
            ],
        },
        "mnemonic": {
            "word": "MINDER",
            "line": "A cognitive neuropsychologist is a MINDER of the mind's hidden parts.",
            "items": [
                ("M", "**Models** of cognition are built and tested."),
                ("I", "**Injury** patterns in patients are the key data."),
                ("N", "**N400** and other ERPs index cognitive processing."),
                ("D", "**Dissociations** reveal separate mental modules."),
                ("E", "**Experimental** roots: controlled tasks and reaction time."),
                ("R", "**Resonance** (fMRI BOLD) and PET map activity to tasks."),
            ],
        },
    },
]

REFERENCES = {
    "books": [
        "Kolb, Bryan, and Ian Q. Whishaw, Fundamentals of Human Neuropsychology (Worth Publishers) "
        "\u2014 the source of the gyri-and-sulci, Wernicke-model, Brodmann-area, Brodmann-map, "
        "visual-pathway, corpus-callosum and motor-asymmetry figures reproduced in these notes, and "
        "the standard reference for the structure-function, localization and lateralization material.",
        "Examining Biological Foundations of Human Behavior, edited by Barre Vijaya Prasad (2020), "
        "IGI Global \u2014 used as a key Paper I source and cross-checked against the standard texts "
        "below for the anatomy and brain-behaviour material.",
        "Lezak, Muriel D., and colleagues, Neuropsychological Assessment \u2014 the standard clinical "
        "reference, cross-checked for the logic and purpose of standardized neuropsychological testing.",
        "Luria, Alexander R., Higher Cortical Functions in Man \u2014 the classic reference on "
        "localization and the clinical investigation of brain-damaged patients, cross-checked for the "
        "clinical-methods material.",
    ],
    "online": [
        "\u201cNeuropsychology\u201d, Wikipedia (text available under the Creative Commons "
        "Attribution-ShareAlike License) \u2014 used for the history of the discipline, the "
        "characterisations of neuropsychology, and the clinical and cognitive approaches.",
        "Mayfield Clinic, \u201cAnatomy of the Brain\u201d (mayfieldclinic.com) \u2014 used for the "
        "lobes, cortex (gyri and sulci), Broca's and Wernicke's areas and the deep structures.",
        "Cerebral Localization teaching presentation (H. Mishra) \u2014 used for the definitions of "
        "cerebral localization and cerebral dominance and the survey of localization methods.",
    ],
    "flagged": [
        "The localisation/lateralisation and brain-imaging teaching materials used here are secondary "
        "lecture summaries; the primary studies they cite (for example Dax 1836, Broca 1861, Myers & "
        "Sperry 1953, and the ERP/PET studies) were referenced only as reported in those summaries and "
        "were not accessed directly for these notes.",
        "Standard texts by Lezak and Luria, and the popular writing of Oliver Sacks and the "
        "split-brain research of Michael Gazzaniga, were drawn on as well-known background where "
        "genuinely applicable; no specific page numbers or quotations are attributed to them.",
    ],
}


def figure_credits():
    """Build the figure-image source list from the figures declared in CONTENT.

    Each entry names the figure and its honest provenance; kept in sync with the
    per-figure caption credits. All Unit VIII raster figures are book crops that
    carry a verbatim 'credit' string, so no open-license wording is fabricated.
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
                suffix = "" if fig["source"].startswith("NIH BioArt") else " Via Wikimedia Commons."
                entry = ("%s \u2014 %s by %s. License: %s.%s"
                         % (fig_label, fig["source"], fig["author"], fig["license"], suffix))
            if entry not in seen:
                seen.add(entry)
                credits.append(entry)
    credits.append(
        "The timeline, Wernicke-Geschwind language loop, imaging-methods comparison and "
        "levels-of-organisation diagrams are original labelled schematics drawn for these notes.")
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
    """Render each raster figure as an inline base64 <figure> (no external resources)."""
    figs = topic.get("figures")
    if not figs:
        return ""
    parts = []
    for fig in figs:
        if fig.get("credit"):
            credit = fig["credit"]
        else:
            credit = "Source: %s by %s. License: %s." % (
                fig["source"], fig["author"], fig["license"])
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
    ref_online = "\n".join("<li>" + _html.escape(o) + "</li>" for o in REFERENCES["online"])
    ref_flagged = "\n".join("<li>" + _html.escape(f) + "</li>" for f in REFERENCES["flagged"])
    ref_figures = "\n".join("<li>" + _html.escape(f) + "</li>" for f in REFERENCES["figures"])

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
    <h3>Named online sources used</h3>
    <ul>
      {ref_online}
    </ul>
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
  <div class="credit">{credit}</div>
</footer>

</body>
</html>
""".format(title=_html.escape(TITLE), css=CSS, subtitle=_html.escape(SUBTITLE),
           credit=_html.escape(CREDIT), toc=toc_items, guide=guide_items,
           topics=topics_html, ref_books=ref_books, ref_online=ref_online,
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
