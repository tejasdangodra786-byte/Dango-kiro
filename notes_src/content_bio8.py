# -*- coding: utf-8 -*-
"""Single source of truth for Biological Foundations of Behavior, Unit VIII
(Introduction to Neuropsychology).

The HTML build (build_bio8.py -> parts) imports SECTIONS and the front-matter
from this module. A byte-identical copy lives in the build directory so the
renderer can import it; the two copies are kept in sync (diff -q clean).

Inline markup tokens kept in text strings (renderers interpret them):
  {{LAW}} {{CODE}} {{GUIDE}} {{THEORY}} {{EG}} {{SEC}}  -> evidence labels
  {{c:TEXT}}                                            -> citation [TEXT]
  [[term|gloss]]                                        -> glossary term (gloss)
  {{fig:NAME}}                                          -> inline SVG figure (HTML only)

Block schema (dicts):
  {"t":"lead","text":str}
  {"t":"h3","text":str}
  {"t":"h4","text":str}
  {"t":"ul","items":[str | (str,[str,...]) ]}      # optional one level of sub-bullets
  {"t":"ol","items":[...]}
  {"t":"table","caption":str|None,"head":[...],"rows":[[...],...]}
  {"t":"box","kind":"def|key|law|ex|warn|info","title":str,"items":[...] , "text":str|None}
  {"t":"fig","key":str,"caption":str}
  {"t":"mn","word":str,"text":str}                  # mnemonic
  {"t":"note","text":str}

TITLE and CREDIT are the byte-exact brand strings.
"""

TITLE = "Biological Foundations of Behavior \u2014 Unit VIII: Introduction to Neuropsychology"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

SUBTITLE = "M.Phil. Clinical Psychology study notes (bullet style) introducing neuropsychology"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens the notes.
INTRO = [
    {"t": "lead", "text": "These are concise, exam-ready revision notes for Unit VIII of the Biological Foundations of Behavior paper (M.Phil. Clinical Psychology). They introduce neuropsychology topic by topic, from the relationship between brain structure and function and the rise of the discipline, through the logic of cerebral organization, localization and lateralization, to the working methods of clinical and cognitive neuropsychologists, so a whole 5-mark answer can be revised from a single topic."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one Unit VIII syllabus topic, written as a 5-mark exam answer in full sentences rather than terse fragments, so you can revise straight from the point and reproduce it under exam conditions.",
        "Every topic opens with an Answer-flow box that gives the ordered skeleton to write, carries a flowchart figure that pictures the core idea, gives a short definition box drawn from standard neuropsychology texts, and closes with a mnemonic that packs the key points into a single memorable word.",
        "A short source note and a references list sit at the end; the references name only the books (full title and author) and the online sources that were actually used.",
    ]},
    {"t": "h3", "text": "How the coloured gloss works (live example)"},
    {"t": "ul", "items": [
        "Hard neuropsychology terms are kept, but each one is explained right after it in a coloured bracket, like this: [[aphasia|loss or impairment of language caused by brain damage]]. Just skim the amber bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "The term itself carries a faint dotted underline and its plain-language meaning follows in amber italics on a peach tint, as in [[lesion|an area of damaged or destroyed brain tissue]]; the two together let you learn the real term and its meaning at once.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail, and note that only a few hook words are bold, never a whole sentence.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word, so you get a visual anchor and a verbal anchor for each answer.",
        "Each topic opens with a short **Answer flow** box that lists, in order, the skeleton to write in the exam \u2014 definition, background, key concepts, evidence and significance \u2014 and the exam points that follow are grouped under small sub-headings in that same order, so the whole topic reads as a ready-to-write answer with a clear beginning, middle and end.",
    ]},
    {"t": "h3", "text": "How citations work"},
    {"t": "ul", "items": [
        "A citation in square brackets names the full book and its author, for example {{c:Kolb}} or {{c:Lezak}}, so you can trace any claim back to the standard text it comes from.",
        "Where a fact was confirmed from an authoritative online reference it is named in full, for example {{c:APA40}}, and the unit's compiled teaching materials are cited by their descriptive title, {{c:HistNP}}.",
    ]},
]

# Honest limits and source conventions.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no authors, citations or page numbers in these notes are invented. The neuropsychological facts stated are standard, textbook-level and verifiable, and the following limits are disclosed so you can check them independently."},
    {"t": "ul", "items": [
        "The characterisations of key ideas are drawn from the standard neuropsychology literature, chiefly Kolb and Whishaw's Fundamentals of Human Neuropsychology, together with Lezak and colleagues' Neuropsychological Assessment and Luria's Higher Cortical Functions in Man; they are given as characterisations, not verbatim quotations, and no page numbers are attached.",
        "The unit's own compiled teaching materials, a history-of-neuropsychology overview, a cerebral-localization presentation, a note on the relation between structure and function, and a brain-imaging-techniques note, were used as inputs and are cited by their descriptive titles, cross-checked against the standard texts; no author is invented for them.",
        "One source supplied with the unit, an anatomy handout on the relation between structure and function of the brain, is an image-only file with no readable text layer, and optical character recognition (OCR) was not available in the environment these notes were prepared in, so its text could not be extracted; moreover that handout is a copyrighted clinical document rather than an openly-licensed one, so its own figures were not embedded either. Its subject matter (the lobes and their functions, the hemispheres, the language areas and the deep structures) is instead covered from the readable note of the same theme and the standard texts, and the figures in these notes are drawn only from verified public-domain sources. This limit is disclosed rather than hidden.",
        "All embedded photographs and plates are genuine public-domain images whose status was verified: a phrenological chart from the People's Cyclopedia (1883), Brodmann's cytoarchitectonic map (1909), Harlow's engraving of Phineas Gage's skull (1868), and two plates from Gray's Anatomy (1918); every image is stored inside the file itself so the notes load no external resource.",
        "Only well-established neuropsychology is stated; no novel claims, study results or quantitative findings beyond standard textbook accounts are presented, and where accounts differ between authors (for example the exact number of Luria's functional units) that difference is noted rather than smoothed over.",
    ]},
]

# Reading list.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the standard neuropsychology books used, the unit's compiled teaching materials, and the authoritative online references consulted for the Unit VIII content."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Bryan Kolb & Ian Q. Whishaw \u2014 Fundamentals of Human Neuropsychology (the primary reference: the development of the discipline, localization and lateralization of function, the hierarchical and functional organization of the brain, split-brain findings, and clinical and cognitive methods).",
        "Muriel D. Lezak, Diane B. Howieson, Erin D. Bigler & Daniel Tranel \u2014 Neuropsychological Assessment (the practice of clinical neuropsychological assessment, test batteries and the interpretation of deficits).",
        "Alexander R. Luria \u2014 Higher Cortical Functions in Man (the functional-systems view of the brain and the three principal functional units or blocks).",
        "Oliver Sacks \u2014 The Man Who Mistook His Wife for a Hat (the single-case clinical tradition that illustrates how focal brain damage fractionates the mind).",
        "Michael S. Gazzaniga \u2014 The Bisected Brain / Cognitive Neuroscience: The Biology of the Mind (split-brain research and the lateralized organization of the two hemispheres).",
        "Barre Vijaya Prasad (2020) \u2014 Examining Biological Foundations of Human Behavior (a cross-check for the brain-structure and history material for the psychology reader).",
    ]},
    {"t": "h3", "text": "Unit teaching materials (cited by descriptive title)"},
    {"t": "ul", "items": [
        "A History of Neuropsychology (unit overview) \u2014 the discipline's definition, its ancient-to-modern history from Imhotep, Aristotle and Hippocrates to Willis, Gall, Bouillaud, Broca and Lashley, and its clinical, experimental and cognitive approaches.",
        "Cerebral Localization (unit presentation) \u2014 the definition of cerebral localization, the history from Gall's phrenology to Brodmann's cytoarchitectonic map, cerebral dominance, and the clinical, electrophysiological, radiological, intra-operative and experimental methods of localization.",
        "The Relation between Structure and Function of the Brain (unit note) \u2014 brain connectivity as structural, functional and effective connectivity, and the large-scale-network view of cognition.",
        "Brain Imaging Techniques (unit note) \u2014 the electrophysiological (EEG, ERP) and haemodynamic (PET, SPECT, fMRI) methods and their logic of signal, averaging and subtraction.",
    ]},
    {"t": "h3", "text": "Online and standard-reference sources"},
    {"t": "ul", "items": [
        "American Psychological Association, Division 40 (Society for Clinical Neuropsychology) \u2014 the professional definition and scope of clinical neuropsychology.",
        "StatPearls (NCBI Bookshelf) \u2014 Neuroanatomy and functional-imaging reference chapters used to confirm standard facts about localization and imaging methods.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. STRUCTURE AND FUNCTION OF THE BRAIN ===========================
_s("t1", "1", "Relationship between structure and function of the brain", [
    {"t": "lead", "text": "Neuropsychology rests on the assumption that the brain's structure and its functions are tightly linked, so that the organisation of neural tissue and its connections gives rise to cognition and behaviour, and damage to a structure predictably changes a function, an idea refined from strict localisation into the modern large-scale-network view."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what the structure\u2013function relationship means and why neuropsychology depends on it.",
        "Background: the core assumption that mind and behaviour emerge from brain tissue and its wiring.",
        "Key concepts: levels of structure, the three kinds of connectivity, and the network view.",
        "Evidence and examples: the lesion method and imaging that reveal structure\u2013function links.",
        "Significance: why the relationship anchors the whole unit and clinical practice.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "structure_function_loop", "caption": "The structure\u2013function loop: neural structure (cells, regions, connections) gives rise to function (cognition, behaviour), while function studied through lesions and imaging lets us infer the structure behind it."},
    {"t": "fig", "key": "pd_gray_lobes", "caption": "The lateral surface of the cerebrum and its four lobes: the physical map onto which functions are plotted."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Standard neuropsychology treats the field as the study of the **relation between brain and behaviour**, asking how the structure and activity of the nervous system produce cognition, emotion and action. {{c:Kolb}}",
        "{{GUIDE}} A central aim is described as the **elucidation of the relationship between structure and function** of large-scale brain systems, that is, how the neural connectional network underlies cognitive functions. {{c:StructFn}}",
        "{{THEORY}} The modern view holds that cognitive functions arise from **interactions within and between a distributed set of brain regions**, so cognition is a global integration of local specialised processing. {{c:StructFn}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and core assumption"},
    {"t": "ul", "items": [
        "The structure\u2013function relationship means that the brain's **physical organisation**, its cells, regions and connections, is the basis of its **functions**, so that mind and behaviour are products of nervous tissue rather than of a separate immaterial agent. {{c:Kolb}}",
        "Neuropsychology depends wholly on this link, because it infers how the brain works by observing how **changes in structure**, from injury or disease, produce **changes in function** such as memory, language or movement. {{c:Kolb}}",
        "The guiding assumption is that the brain shows **functional specialisation**, meaning different structures make different contributions, which is why a focal [[lesion|an area of damaged or destroyed brain tissue]] removes some abilities while sparing others. {{c:HistNP}}",
    ]},
    {"t": "h4", "text": "Levels of structure that carry function"},
    {"t": "ul", "items": [
        "Structure operates at **several levels**, from single [[neuron|a nerve cell that receives, processes and transmits information]] through local circuits and cortical columns to whole regions and systems, and function can be mapped onto any of these levels. {{c:Kolb}}",
        "The **neuron is the functional unit** that lets us receive information, process it and produce actions, while [[glia|support cells that hold neurons together and nourish and insulate them]] maintain the tissue that makes this possible. {{c:Kolb}}",
        "At the largest scale, the **cerebral cortex** carries higher functions such as thought and language, so its surface organisation into lobes and areas is a map of where many functions are chiefly handled. {{c:Prasad}}",
    ]},
    {"t": "h4", "text": "The four lobes and their broad functions"},
    {"t": "ul", "items": [
        "The **frontal lobe** handles movement, speech production, and the executive functions of planning, judgement, problem-solving and the regulation of behaviour and emotion, so it is the brain's control and personality centre. {{c:Prasad}}",
        "The **parietal lobe** integrates touch, pain, temperature and position sense, builds the sense of space and the body schema, and supports reading, arithmetic and the guidance of action in space. {{c:StructFn}}",
        "The **temporal lobe** carries hearing, the recognition of what things are, the comprehension of language, and, through its medial structures, the laying down of new long-term memories. {{c:Kolb}}",
        "The **occipital lobe** is devoted to vision, receiving and interpreting the visual world, so that damage there produces blindness or specific visual deficits with the eyes themselves intact. {{c:Prasad}}",
        "Deeper structures matter too: the **brainstem** sustains the vital functions of breathing and heartbeat and arousal, the **cerebellum** coordinates balance and smooth movement, and the **thalamus** is the relay through which most sensory traffic reaches the cortex. {{c:StructFn}}",
    ]},
    {"t": "h4", "text": "Connectivity: the three kinds"},
    {"t": "ul", "items": [
        "**Brain connectivity** describes how distinct units are linked, and the unit note distinguishes **three kinds**: structural, functional and effective connectivity, each capturing a different sense in which regions are related. {{c:StructFn}}",
        "[[Structural connectivity|the physical white-matter fibre pathways joining regions]] refers to the anatomical links between regions and can be estimated non-invasively with [[diffusion tensor imaging|a scan that traces the direction of water movement along fibre tracts]]. {{c:StructFn}}",
        "[[Functional connectivity|a statistical correlation in activity between regions, without a stated cause]] refers to temporal correlations or statistical dependencies between the activity of remote regions, but it does **not show direction or causation**. {{c:StructFn}}",
        "[[Effective connectivity|the directed causal influence one region exerts over another]] describes the causal influence one neural element exerts over another, uniting the structural wiring with the functional pattern into a directed diagram. {{c:StructFn}}",
    ]},
    {"t": "h4", "text": "The large-scale-network view"},
    {"t": "ul", "items": [
        "A newer paradigm holds that **cognition results from the dynamic interactions** of distributed areas operating in **large-scale networks**, such as the networks for vision, attention, memory and motion. {{c:StructFn}}",
        "On this view the **architecture of the structural connections shapes the functional interactions** among areas, so the fixed wiring constrains which flexible functional patterns can arise. {{c:StructFn}}",
        "This network view softens strict localisation, because a single function is seen to depend on a **distributed circuit**, and damage to a hub region can disturb activity far beyond the site of the lesion. {{c:StructFn}}",
    ]},
    {"t": "h4", "text": "Evidence, significance and a clinical note"},
    {"t": "ul", "items": [
        "The oldest evidence for the link is the **lesion method**, in which a lost function is matched to the damaged structure, as when loss of fluent speech is matched to damage in the left frontal lobe. {{c:Kolb}}",
        "Modern **imaging** supplies converging evidence, because techniques that show activity during a task reveal which structures engage for which functions, confirming the structure\u2013function map in the living brain. {{c:ImgTech}}",
        "The relationship can **change across the lifespan and in disease**, since the coupling between structure and function strengthens with age and is disrupted in conditions such as schizophrenia and epilepsy. {{c:StructFn}}",
        "Clinically the link is the whole justification for **neuropsychological assessment**, because a reliable structure\u2013function map lets a clinician reason from a pattern of deficits back toward the likely site and nature of brain pathology. {{c:Lezak}}",
    ]},
    {"t": "box", "kind": "key", "title": "Structure and function at a glance", "items": [
        "The brain's structure (cells, regions, connections) gives rise to its functions (cognition, behaviour).",
        "Functional specialisation: different structures make different contributions, so focal damage removes specific abilities.",
        "Three connectivities: structural (wiring), functional (correlation), effective (causal direction).",
        "Modern view: cognition emerges from distributed large-scale networks, not isolated spots.",
        "Evidence: the lesion method plus imaging; significance: it justifies neuropsychological assessment.",
    ]},
    {"t": "mn", "word": "SHAPE", "text": "Structure Houses And Produces Experience \u2014 the brain's structure houses and produces behaviour; remember the three connectivities (Structural wiring, Functional correlation, Effective causation) and that cognition emerges from distributed networks, not single spots."},
])

# ========================= 2. THE RISE OF NEUROPSYCHOLOGY ======================
_s("t2", "2", "The rise of neuropsychology as a distinct discipline", [
    {"t": "lead", "text": "Neuropsychology is a relatively young discipline whose roots run from ancient Egypt and the Greeks through Willis, Gall, Bouillaud, Broca, Wernicke, Hughlings-Jackson, Lashley, Hebb and Luria, maturing into a named field only after Broca localised speech and after Hebb's and Kolb and Whishaw's integrative textbooks."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what neuropsychology is and that it is a recent, distinct discipline.",
        "Early history: the ancient brain-versus-heart debate and the mind\u2013body problem.",
        "Nineteenth-century pioneers: Gall, Bouillaud, Broca, Wernicke who localised function.",
        "Twentieth-century shapers: Hughlings-Jackson, Lashley, Hebb and Luria.",
        "Emergence as a field: the naming and the founding textbooks.",
        "Significance and a one-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "np_timeline", "caption": "A timeline of the rise of neuropsychology, from the ancient brain\u2013heart debate through the nineteenth-century localisers to the twentieth-century integrators who named the field."},
    {"t": "fig", "key": "pd_phrenology", "caption": "Gall's phrenology read the skull's bumps as mental faculties: wrong in detail, but the first systematic claim that functions are localised in the brain."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Neuropsychology is defined as the branch of psychology concerned with **how the brain and nervous system influence cognition and behaviour**, and especially how brain injury or illness affects them. {{c:HistNP}}",
        "{{GUIDE}} It is both an **experimental and a clinical field** that seeks to discover how the brain correlates with the mind through the study of neurological patients, bridging neurology and psychology. {{c:HistNP}}",
        "{{SEC}} The first textbook defining the field, **Fundamentals of Human Neuropsychology**, was published by Kolb and Whishaw in 1980, a convenient marker of the discipline's maturity. {{c:HistNP}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the newness of the field"},
    {"t": "ul", "items": [
        "Neuropsychology is the field that studies the **brain\u2013behaviour relationship**, combining neurology's focus on nervous-system pathology with psychology's focus on cognition, which is why it is described as both experimental and clinical. {{c:HistNP}}",
        "It is a **relatively new discipline**, even though its questions are ancient, and it became a recognisable named field only in the twentieth century with integrative theory and textbooks. {{c:HistNP}}",
        "The word itself was popularised through its use as a subtitle of **Donald Hebb's 1949 book** The Organization of Behavior, signalling a deliberately multidisciplinary science of brain and behaviour. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Ancient roots and the mind\u2013body problem"},
    {"t": "ul", "items": [
        "In **ancient Egypt**, medical writings from the time of Imhotep described the brain and head trauma, yet Egyptians still saw the **heart, not the brain**, as the seat of the soul. {{c:HistNP}}",
        "**Aristotle** reinforced this heart-centred view, treating the brain as a radiator that cooled the blood, whereas **Hippocrates** argued the brain held the greatest power and was the seat of the mind. {{c:HistNP}}",
        "The philosopher **Ren\u00e9 Descartes** framed the enduring [[mind\u2013body problem|the question of how a non-physical mind relates to the physical body]], proposing [[dualism|the view that mind and body are distinct but interacting]] with the pineal gland as their meeting point. {{c:HistNP}}",
        "**Galen**, physician to the Roman gladiators, inferred from head injuries that the brain governs sensation and movement, but located mental life in the fluid-filled [[ventricles|the fluid-filled cavities inside the brain]], a 'ventricular doctrine' that dominated for over a thousand years. {{c:HistNP}}",
        "**Thomas Willis** took a physiological view of the brain in the seventeenth century, coined the words 'hemisphere' and 'lobe', and used 'neurology', linking higher structures to complex functions. {{c:HistNP}}",
    ]},
    {"t": "h4", "text": "Nineteenth-century localisers"},
    {"t": "ul", "items": [
        "**Franz Josef Gall** proposed that mental faculties sat in specific brain regions and founded [[phrenology|the discredited claim that skull bumps reveal mental faculties]], which, though wrong, put brain localisation onto the scientific agenda. {{c:Kolb}}",
        "**Pierre Flourens** demolished phrenology with controlled lesion experiments, arguing instead that the cortex worked as a **whole** and could recover function, an early statement of the holistic position. {{c:Kolb}}",
        "**Jean-Baptiste Bouillaud** argued from many clinical cases that **speech is localised in the frontal lobes** and was among the first to use large samples, extending Gall's localisation on firmer ground. {{c:Kolb}}",
        "**Paul Broca** in 1861 examined the patient 'Tan', found a left frontal lesion at autopsy, and so showed that language is **localised and lateralised**, a demonstration widely taken to launch the science. {{c:Kolb}}",
        "**Carl Wernicke** in 1874 showed that a **posterior temporal lesion** disturbs language comprehension, described a connectionist model of language, and predicted a 'conduction' aphasia from a disconnection. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Twentieth-century shapers"},
    {"t": "ul", "items": [
        "**John Hughlings-Jackson** proposed that the nervous system is organised in a **functional hierarchy** of levels, so disease produces 'dissolution', a reversal of evolution toward simpler behaviour. {{c:Kolb}}",
        "**Karl Lashley** sought the memory trace or [[engram|the hypothetical physical location of a stored memory]] in rats and concluded, from his lesion studies, in favour of [[mass action|the idea that the cortex acts as a whole, by amount not location]] and [[equipotentiality|the idea that intact cortex can take over a lost region's role]]. {{c:HistNP}}",
        "**Donald Hebb** gave the field a unifying theory in which **cell assemblies** and strengthened synapses link neural activity to learning, providing a bridge from physiology to psychology. {{c:Kolb}}",
        "**Alexander Luria** built a comprehensive **functional-systems** account of the brain from his work with brain-injured soldiers, shaping both the theory and the clinical assessment of the new discipline. {{c:Luria}}",
    ]},
    {"t": "h4", "text": "Emergence, significance and a note"},
    {"t": "ul", "items": [
        "The field finally **crystallised** when localisation, hierarchy and functional-systems ideas were combined with rigorous patient study and, from 1980, gathered into a defining textbook. {{c:Kolb}}",
        "Its rise matters because it replaced **spirit- and heart-based explanations** of behaviour with a testable science that traces cognition and its disorders to the brain. {{c:HistNP}}",
        "Modern neuropsychology now spans **research, clinical and forensic settings**, showing how far the discipline has travelled from the ancient belief that the brain was a useless organ discarded at burial. {{c:HistNP}}",
    ]},
    {"t": "box", "kind": "key", "title": "The rise of neuropsychology at a glance", "items": [
        "Ancient roots: heart-versus-brain debate (Egypt, Aristotle) \u2192 brain as seat of mind (Hippocrates).",
        "Descartes' mind\u2013body problem and dualism; Willis coins 'hemisphere', 'lobe', 'neurology'.",
        "Localisers: Gall (phrenology), Flourens (holism), Bouillaud, Broca (1861), Wernicke (1874).",
        "Shapers: Hughlings-Jackson (hierarchy), Lashley (mass action, equipotentiality), Hebb, Luria.",
        "Named field: Hebb (1949) popularises the word; Kolb & Whishaw (1980) write the defining textbook.",
    ]},
    {"t": "mn", "word": "GB-WHLH", "text": "Gall, Bouillaud, Wernicke, Broca, Hughlings-Jackson, Lashley, Hebb/Luria \u2014 the chain of pioneers from phrenology through the localisation of speech to the hierarchy, mass action and the functional-systems theories that named the discipline."},
])


# ========================= 3. THE LOGIC OF CEREBRAL ORGANIZATION ===========================
_s("t3", "3", "The logic of cerebral organization", [
    {"t": "lead", "text": "The brain is organised on a logic of hierarchy and functional systems: Hughlings-Jackson's layered levels control behaviour from the spinal cord up to the forebrain, while Luria grouped the cortex into functional units and three processing zones, with redundancy and plasticity allowing parts of a system to compensate for loss."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what is meant by the 'logic' or organising principles of the brain.",
        "Background: Hughlings-Jackson's hierarchy and the evolution\u2013dissolution idea.",
        "Key concepts: Luria's functional units and the primary\u2013secondary\u2013tertiary zones.",
        "Levels and redundancy: parallel organisation, plasticity and compensation.",
        "Evidence and significance: how the logic explains deficits and recovery.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "luria_units", "caption": "The logic of cerebral organization: Hughlings-Jackson's hierarchy of levels and Luria's functional units with their primary, secondary and tertiary zones arranged from simple to complex processing."},
    {"t": "fig", "key": "pd_gray_medial", "caption": "The medial surface of the hemisphere: the vertical stack from brainstem up to forebrain mirrors Hughlings-Jackson's hierarchy of levels."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{THEORY}} Hughlings-Jackson conceived the nervous system as organised in **layers arranged in a functional hierarchy**, each higher level controlling more complex behaviour through the lower levels. {{c:Kolb}}",
        "{{THEORY}} Luria divided the cortex into **functional units**, a posterior sensory unit that receives, processes and stores information and an anterior motor unit that forms intentions and programs action. {{c:Kolb}}",
        "{{GUIDE}} Luria's broader account describes **three principal functional units or blocks**, one for arousal, one for receiving and processing information, and one for planning and regulating activity. {{c:Luria}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and organising principles"},
    {"t": "ul", "items": [
        "The 'logic' of cerebral organization refers to the **general principles** by which the brain is arranged, chiefly hierarchy, functional systems, levels of processing and redundancy, rather than a mere list of parts. {{c:Kolb}}",
        "The brain is not a bag of equal tissue, because it shows an **orderly plan** in which simple functions sit at lower levels and increasingly complex functions are built upon them at higher levels. {{c:Kolb}}",
        "A key principle is that complex behaviours are run by **functional systems**, chains of cooperating regions, so a single ability draws on several structures working together rather than one centre. {{c:Luria}}",
    ]},
    {"t": "h4", "text": "Hierarchical organisation (Hughlings-Jackson)"},
    {"t": "ul", "items": [
        "Hughlings-Jackson arranged the nervous system into a **hierarchy of levels**, often the spinal cord, the brainstem and the forebrain, with each higher level governing more complex aspects of behaviour. {{c:Kolb}}",
        "He borrowed from evolution the idea that the brain was built in **steps**, so that higher levels, added later, control and refine the older, simpler machinery beneath them. {{c:Kolb}}",
        "Damage to the highest levels therefore produces [[dissolution|a reversal of evolution, leaving simpler, more primitive behaviour]], in which the person keeps a repertoire of behaviours but of a simpler, more automatic kind. {{c:Kolb}}",
        "This logic reconciled the old dispute, because experiments showing an animal could still walk after cortical loss simply revealed **lower-level behaviour** surviving once the higher controller was gone. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Luria's functional units and zones"},
    {"t": "ul", "items": [
        "Luria grouped the cortex into a **posterior sensory unit**, which receives and stores information, and an **anterior motor unit**, which forms intentions and executes programs of action. {{c:Kolb}}",
        "His wider model describes **three functional units or blocks**: a brainstem unit for [[arousal|the general level of alertness and readiness of the brain]], a posterior unit for processing information, and a frontal unit for planning. {{c:Luria}}",
        "Within each cortical unit he described **three zones** arranged one above the other, so processing moves from simple to integrated as it passes from zone to zone. {{c:Kolb}}",
        "The **primary zone** handles raw sensation or movement, the [[secondary zone|cortex that organises and combines inputs into patterns]] organises it, and the [[tertiary zone|association cortex that integrates across the senses for the most complex functions]] integrates across systems. {{c:Kolb}}",
        "Note that accounts differ: Kolb and Whishaw summarise Luria as dividing the cortex into **two functional units**, while Luria's own writings set out the **three-block** scheme, a difference worth stating plainly. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Levels, redundancy and plasticity"},
    {"t": "ul", "items": [
        "Because functions are built on **multiple levels**, a behaviour can survive in simplified form when a higher level is lost, which is the hierarchical logic seen from the clinical side. {{c:Kolb}}",
        "The brain shows **redundancy and plasticity**, so within limits intact tissue can take over part of a lost region's role, which Lashley once overstated as equipotentiality. {{c:HistNP}}",
        "Recognising this [[plasticity|the brain's capacity to reorganise and let areas take on new roles]] is one practical pay-off of localization work, because it underlies the possibility of recovery and rehabilitation. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Evidence and significance"},
    {"t": "ul", "items": [
        "The organising logic is supported by the **orderly march of symptoms** in epilepsy and the graded effects of lesions, which fit a layered, systems-based plan rather than a single undifferentiated mass. {{c:Kolb}}",
        "Its significance is that it lets a clinician **predict the pattern of a deficit** from the level and system affected, and understand why some functions recover while others do not. {{c:Lezak}}",
    ]},
    {"t": "box", "kind": "key", "title": "The logic of cerebral organization at a glance", "items": [
        "Organising principles: hierarchy, functional systems, levels of processing, redundancy.",
        "Hughlings-Jackson: layered levels (cord, brainstem, forebrain); damage to the top causes dissolution.",
        "Luria: posterior sensory unit and anterior motor unit; wider three-block (arousal, processing, planning) model.",
        "Zones: primary (raw) \u2192 secondary (organise) \u2192 tertiary (integrate across systems).",
        "Redundancy and plasticity allow partial compensation; the logic predicts deficits and recovery.",
    ]},
    {"t": "mn", "word": "HULiP", "text": "Hierarchy, Units, Levels, Plasticity \u2014 the brain is a hierarchy (Hughlings-Jackson) of functional Units (Luria) built in Levels (primary\u2013secondary\u2013tertiary zones) with Plasticity and redundancy that allow compensation."},
])

# ========================= 4. LOCALIZATION OF FUNCTIONS ===========================
_s("t4", "4", "Localization of functions", [
    {"t": "lead", "text": "Localization of function is the principle that specific mental operations are handled by specific brain regions, demonstrated by Broca's and Wernicke's language areas and the stimulation maps of Fritsch, Hitzig and Ferrier, and set against the opposing equipotentiality or mass-action view, with the modern position a qualified, network-based localisation."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what localization of function means.",
        "Background: the debate between localizationism and holism/equipotentiality.",
        "Key evidence: Broca's and Wernicke's areas and the stimulation maps.",
        "Mapping and modularity: Brodmann's areas and the idea of modules.",
        "The localization debate resolved: a qualified, network-based view.",
        "Significance and a one-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "localization_spectrum", "caption": "The localization debate as a spectrum, from strict localizationism (one function, one spot) through modular and network views to full equipotentiality (the cortex acts as an undifferentiated whole)."},
    {"t": "fig", "key": "pd_brodmann", "caption": "Brodmann's 1909 cytoarchitectonic map numbers the cortical areas by their cell architecture; localisation still pins functions to these numbered areas (44/45 Broca's, 22 Wernicke's, 17 vision)."},
    {"t": "fig", "key": "pd_gage", "caption": "Phineas Gage's skull and tamping iron (Harlow, 1868): the classic lesion case tying the frontal lobes to judgement, planning and social behaviour."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Cerebral localization is defined as the **mapping of the cerebral cortex into areas** and the correlation of those areas with particular cerebral functions. {{c:CerebLoc}}",
        "{{THEORY}} The **localizationist** position holds that different parts of the brain have different functions, so that focal damage removes a specific ability while sparing others. {{c:Kolb}}",
        "{{THEORY}} The opposing view, developed by Flourens and later Lashley, holds that the cortex works as a **whole** through mass action and equipotentiality, de-emphasising fixed centres. {{c:Kolb}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the two positions"},
    {"t": "ul", "items": [
        "Localization of function is the principle that a **specific function is carried out by a specific region**, so that mapping the cortex into areas and matching them to functions is possible and clinically useful. {{c:CerebLoc}}",
        "The clinical form of the idea is the **diagnosis of the site of a lesion** from the patient's signs and symptoms, reasoning backward from the deficit to the damaged structure. {{c:CerebLoc}}",
        "It stands against [[equipotentiality|the idea that any part of a functional area can do the work of the whole]], which with [[mass action|the idea that the cortex acts by the amount of tissue, not its location]] claims the cortex acts as an undifferentiated whole. {{c:HistNP}}",
    ]},
    {"t": "h4", "text": "The founding language evidence"},
    {"t": "ul", "items": [
        "**Paul Broca** localised speech production to the posterior third of the **left inferior frontal gyrus**, now Broca's area, after his patient 'Tan' proved to have a lesion there, giving the first strong evidence for localisation. {{c:Kolb}}",
        "Damage to this region causes [[Broca's aphasia|non-fluent, effortful, telegraphic speech with relatively preserved comprehension]], in which the patient understands well but produces slow, laboured, agrammatic output. {{c:CerebLoc}}",
        "**Carl Wernicke** localised language comprehension to the **posterior superior temporal region**, now Wernicke's area, where a lesion disturbs understanding while speech stays fluent. {{c:CerebLoc}}",
        "Damage there causes [[Wernicke's aphasia|fluent but meaningless speech with poor comprehension]], and a lesion of the connecting fibres produces a [[conduction aphasia|good comprehension and speech but very poor repetition]], exactly as Wernicke predicted. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Stimulation and mapping evidence"},
    {"t": "ul", "items": [
        "**Fritsch and Hitzig** in 1870 showed that electrically stimulating points on the dog's cortex produced **movements of specific muscles**, the first experimental proof that the motor cortex is topographically organised. {{c:CerebLoc}}",
        "**David Ferrier** and others extended stimulation mapping to the **auditory** and other cortices in primates, and **Penfield** later mapped motor and sensory cortex in conscious human surgical patients. {{c:CerebLoc}}",
        "**Penfield and Rasmussen** outlined the motor and sensory [[homunculus|a distorted body map on the cortex, with large areas for the hands and lips]], showing that body parts needing fine control occupy disproportionately large cortical territory. {{c:CerebLoc}}",
        "**Korbinian Brodmann** divided the cortex into roughly **fifty-two numbered areas** on [[cytoarchitectonic|based on the microscopic arrangement and types of cells]] grounds, giving localisation a durable anatomical map still used today. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Modularity and the debate resolved"},
    {"t": "ul", "items": [
        "The idea of **modularity** sharpens localisation, proposing that the mind has specialised, somewhat independent processors, so a specific module can be selectively impaired by focal damage. {{c:Kolb}}",
        "Lashley's **mass action and equipotentiality** were later understood as a misreading of his rat data, because running a maze needs several cortical areas, so large lesions impaired many functions at once. {{c:HistNP}}",
        "The modern position is a **qualified localisation**, in which functions are localised yet distributed across cooperating regions, combining the truth in localizationism with the network view of organisation. {{c:StructFn}}",
    ]},
    {"t": "h4", "text": "Localising signs by lobe"},
    {"t": "ul", "items": [
        "**Frontal-lobe** damage produces impaired memory and planning, poor judgement and insight, disinhibition or apathy, and language problems such as Broca's aphasia, the Phineas Gage case being the classic illustration of a changed personality after frontal injury. {{c:CerebLoc}}",
        "**Parietal-lobe** damage produces somatosensory loss, [[astereognosis|inability to recognise objects by touch alone]], and, on the right, disturbances of the body schema such as neglect, [[anosognosia|denial or unawareness of one's own deficit]], and dressing or constructional apraxia. {{c:CerebLoc}}",
        "**Temporal-lobe** damage produces hearing and auditory-recognition problems, Wernicke's aphasia on the left, memory disturbance, and complex phenomena such as d\u00e9j\u00e0 vu or olfactory hallucinations in temporal-lobe seizures. {{c:CerebLoc}}",
        "**Occipital-lobe** damage produces visual-field loss ([[hemianopia|loss of one half of the visual field]]) and the visual [[agnosias|failures to recognise what is seen, with vision itself intact]], including [[prosopagnosia|the specific inability to recognise faces]]. {{c:CerebLoc}}",
        "A revealing cluster is [[Gerstmann syndrome|agraphia, acalculia, finger agnosia and right\u2013left confusion from the left angular gyrus]], whose four signs together point precisely to the dominant parietal region. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "The aphasias as a localising map"},
    {"t": "ul", "items": [
        "Broca's (**motor, non-fluent**) aphasia follows damage to areas 44/45, leaving good comprehension but effortful, telegraphic, agrammatic speech. {{c:CerebLoc}}",
        "Wernicke's (**sensory, fluent**) aphasia follows damage to area 22, leaving fluent but poorly understood speech with impaired comprehension. {{c:CerebLoc}}",
        "**Conduction aphasia** follows damage to the arcuate fasciculus connecting the two areas, so comprehension and speech are good but repetition is strikingly poor, exactly the disconnection Wernicke predicted. {{c:CerebLoc}}",
        "**Anomic aphasia** shows mainly word-finding difficulty, and a lesion of the angular and supramarginal gyri can yield fluent but unintelligible jargon, so the type of aphasia itself maps the site of damage. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Significance"},
    {"t": "ul", "items": [
        "Localisation is the **working logic of clinical neuropsychology and neurology**, letting a clinician infer where the brain is damaged from the pattern of language, memory or perceptual deficits. {{c:Lezak}}",
        "It also guides **neurosurgery**, where mapping eloquent 'no-go' areas before operating protects speech and movement and increases the precision of safe resection. {{c:CerebLoc}}",
    ]},
    {"t": "box", "kind": "key", "title": "Localization of functions at a glance", "items": [
        "Localization: specific functions in specific regions; the clinical art of locating a lesion from signs.",
        "Opposed by equipotentiality and mass action (Flourens, Lashley): the cortex as a whole.",
        "Language evidence: Broca's area (production) and Wernicke's area (comprehension).",
        "Mapping: Fritsch\u2013Hitzig and Ferrier stimulation, Penfield's homunculus, Brodmann's 52 areas.",
        "Modern view: qualified, network-based localisation; it guides diagnosis and neurosurgery.",
    ]},
    {"t": "mn", "word": "BROCA", "text": "Broca, Regions, Opponents, Cytoarchitecture, Areas \u2014 Broca and Wernicke localised language to Regions; Opponents (Flourens, Lashley) argued equipotentiality; Brodmann's Cytoarchitecture mapped the numbered Areas that localisation still uses."},
])


# ========================= 5. LATERALIZATION OF FUNCTIONS ===========================
_s("t5", "5", "Lateralization of functions", [
    {"t": "lead", "text": "Lateralization is the principle that the two cerebral hemispheres are specialised for different functions: the left usually dominates language, analysis and sequencing, the right spatial, musical and holistic processing, a division revealed by Broca's left-sided speech finding and most dramatically by the split-brain studies of Sperry and Gazzaniga."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what lateralization and hemispheric specialisation mean.",
        "Background: Broca's discovery that language is left-lateralised.",
        "Key concepts: left-hemisphere and right-hemisphere specialisations and dominance.",
        "Evidence: split-brain studies and the crossed visual-field logic.",
        "Handedness and language dominance; cautions against over-simplifying.",
        "Significance and a one-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "hemisphere_split", "caption": "Hemispheric lateralization: the left hemisphere's specialisation for language, analysis and sequencing set beside the right hemisphere's specialisation for spatial, musical and holistic processing, joined by the corpus callosum."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Lateralization refers to the finding that functions can be **localised to one side of the brain**, a property Broca first demonstrated when speech proved to depend on the left hemisphere. {{c:Kolb}}",
        "{{GUIDE}} The **dominant (usually left) hemisphere** is characterised as specialised for language, analytical and mathematical skills and the temporal sequencing of stimuli. {{c:CerebLoc}}",
        "{{GUIDE}} The **non-dominant (usually right) hemisphere** is characterised as specialised for spatial perception, singing and the playing of musical instruments. {{c:CerebLoc}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and discovery"},
    {"t": "ul", "items": [
        "Lateralization, also called [[cerebral dominance|the tendency for one hemisphere to lead a given function]] or asymmetry, is the principle that the two hemispheres are **not identical in function** but specialised for different kinds of processing. {{c:CerebLoc}}",
        "The discovery grew directly from **Broca's** finding, because locating speech in the **left** frontal lobe showed that a function could belong to one side, which he named lateralisation. {{c:Kolb}}",
        "The two hemispheres look nearly symmetrical and are joined by the [[corpus callosum|the great fibre bridge that connects the two hemispheres]], yet they divide cognitive labour between them. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Left- and right-hemisphere specialisations"},
    {"t": "ul", "items": [
        "The **left hemisphere** is usually dominant for **language**, covering speech and writing, and for analytical, mathematical and sequential processing that unfolds step by step in time. {{c:CerebLoc}}",
        "The **right hemisphere** is usually dominant for **spatial perception** of three-dimensional objects, for music such as singing and playing instruments, and for holistic, simultaneous processing. {{c:CerebLoc}}",
        "A useful contrast is that the left tends to work **analytically and verbally** while the right works **holistically and spatially**, though both cooperate in nearly every real task. {{c:Kolb}}",
        "Right-hemisphere damage characteristically produces **spatial and body-schema disturbances** such as neglect of the left side, dressing and constructional difficulties, mirroring the left's language role. {{c:CerebLoc}}",
    ]},
    {"t": "h4", "text": "Split-brain evidence"},
    {"t": "ul", "items": [
        "The clearest evidence came from **split-brain** patients, in whom surgeons cut the corpus callosum to control severe epilepsy, which stops the hemispheres from sharing information directly. {{c:Kolb}}",
        "Work by **Roger Sperry** and **Michael Gazzaniga** used the crossed wiring of vision, in which each visual field projects to the **opposite hemisphere**, to speak to one hemisphere at a time. {{c:Gazzaniga}}",
        "When an object was shown only to the right visual field, the **left (speaking) hemisphere** could name it, but when shown only to the left field, the right hemisphere recognised it yet **could not speak** its name. {{c:Kolb}}",
        "These studies showed that each hemisphere has its **own abilities and awareness**, confirming lateralisation in the most direct way and earning Sperry a share of the Nobel Prize. {{c:Gazzaniga}}",
    ]},
    {"t": "h4", "text": "Handedness, dominance and cautions"},
    {"t": "ul", "items": [
        "Language is **left-lateralised in almost all right-handers** and in most left-handers too, so left-hemisphere dominance for speech is the usual pattern across the population. {{c:Kolb}}",
        "The [[Wada test|an injection that briefly sends one hemisphere to sleep to find the language side]] and modern imaging are used before surgery to confirm which hemisphere holds language in a given patient. {{c:CerebLoc}}",
        "The specialisations are **relative, not absolute**, so popular talk of purely 'left-brained' or 'right-brained' people over-simplifies what is a matter of degree and cooperation. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Significance"},
    {"t": "ul", "items": [
        "Lateralisation lets a clinician **predict which hemisphere** is damaged from the profile of deficits, for example language loss pointing left and spatial neglect pointing right. {{c:Lezak}}",
        "It also protects patients in surgery, because knowing the **language-dominant side** guides the surgeon away from eloquent cortex and informs techniques such as the Wada procedure and cortical mapping. {{c:CerebLoc}}",
    ]},
    {"t": "box", "kind": "key", "title": "Lateralization of functions at a glance", "items": [
        "Lateralization = hemispheric specialisation; the hemispheres differ in function though joined by the corpus callosum.",
        "Left (dominant): language, analysis, mathematics, temporal sequencing.",
        "Right (non-dominant): spatial perception, music, holistic processing; its damage causes neglect.",
        "Split-brain studies (Sperry, Gazzaniga): each hemisphere acts and 'knows' on its own via crossed vision.",
        "Language is left-lateralised in most people; specialisations are relative; the Wada test confirms dominance.",
    ]},
    {"t": "mn", "word": "LARS", "text": "Left = Analysis, Right = Spatial \u2014 the Left hemisphere leads language, Analysis and sequencing; the Right leads Spatial and musical, holistic processing; split-brain work (Sperry, Gazzaniga) proved each side can act alone."},
])

# ========================= 6. METHODS OF CLINICAL NEUROPSYCHOLOGISTS ===========================
_s("t6", "6", "Approaches and methodologies of clinical neuropsychologists", [
    {"t": "lead", "text": "Clinical neuropsychologists assess, diagnose and rehabilitate people with brain injury or disease using standardised test batteries, either fixed batteries such as the Halstead-Reitan and Luria-Nebraska and the Indian NIMHANS battery, or a flexible, hypothesis-driven selection, supported by the case and lesion method, interview and behavioural observation."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what clinical neuropsychology is and what clinicians do.",
        "Background: the case/lesion method and the logic of assessment.",
        "Key approaches: fixed versus flexible batteries, with named examples.",
        "The assessment process: history, testing, norms and interpretation.",
        "Rehabilitation and settings; significance.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "clinical_methods", "caption": "The clinical neuropsychologist's workflow and the fixed-versus-flexible battery choice, from referral and history through standardised testing to diagnosis and rehabilitation."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Clinical neuropsychology is the **application of neuropsychological knowledge to the assessment, management and rehabilitation** of people who have suffered illness or injury affecting the brain. {{c:HistNP}}",
        "{{GUIDE}} Neuropsychological tests are described as **standardised** tasks, administered first to reference groups to yield [[normative data|typical scores from a reference group, used to judge a patient]] against which an individual's performance is compared. {{c:HistNP}}",
        "{{GUIDE}} Clinical assessment is characterised as distinguishing deficits due to **brain pathology** from those due to emotional or other, potentially reversible, causes. {{c:HistNP}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the clinician's task"},
    {"t": "ul", "items": [
        "The clinical neuropsychologist applies knowledge of brain\u2013behaviour relations to **assess, diagnose and rehabilitate** patients whose brain injury or illness has caused cognitive or behavioural problems. {{c:HistNP}}",
        "A core task is to decide whether a difficulty reflects **brain pathology or a reversible cause** such as depression, which refines a vague diagnosis into a specific, useful one. {{c:HistNP}}",
        "The work rests on **standardised tests** whose scores are compared with normative data, so that a patient's performance can be judged against what is typical for their age and background. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "The case and lesion method"},
    {"t": "ul", "items": [
        "The oldest clinical method is the **case and lesion method**, in which a patient's pattern of spared and lost abilities is matched to the site of brain damage to infer that region's normal role. {{c:Kolb}}",
        "This method made the discipline, from **Broca's single case** 'Tan' onward, and it still frames how a clinician reasons from a deficit profile to a likely lesion. {{c:Kolb}}",
        "Careful **interview and behavioural observation** sit alongside testing, because the history, the manner of failure and everyday function all inform the interpretation of scores. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "Fixed versus flexible batteries"},
    {"t": "ul", "items": [
        "A [[fixed battery|a standard set of tests given in full to every patient]] applies the same comprehensive set of tests to everyone, giving broad, comparable coverage and a fixed scoring scheme. {{c:Lezak}}",
        "The **Halstead-Reitan Battery** is the classic fixed battery, a fixed set of tests yielding an impairment index that screens broadly for the presence and lateralisation of brain damage. {{c:Lezak}}",
        "The **Luria-Nebraska Battery** is a fixed battery that systematised Luria's qualitative examination into standardised items grouped into clinical scales. {{c:Lezak}}",
        "The **NIMHANS neuropsychological battery**, developed in India, adapts and norms tests for the Indian population, an example of culturally appropriate standardisation of a battery. {{c:Lezak}}",
        "A [[flexible battery|tests chosen case by case to test specific hypotheses]] instead selects tests to answer the specific referral question, following a hypothesis-testing logic that tailors the examination to the patient. {{c:Lezak}}",
        "The **trade-off** is that fixed batteries give standard, comparable breadth but can be long and insensitive to a specific question, while flexible batteries are efficient and targeted but harder to compare across patients. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "The assessment process and domains"},
    {"t": "ul", "items": [
        "A full assessment samples **several cognitive domains**, including attention, memory, language, [[praxis|the ability to carry out skilled, planned movements]], perception and executive function, so that the whole profile is mapped. {{c:Lezak}}",
        "Example tests named in the literature include the **Wechsler** memory and intelligence scales, the Wisconsin Card Sorting Test and the Boston Naming Test, each linked to particular processes. {{c:HistNP}}",
        "Scores are interpreted against **norms and the patient's own baseline**, looking for the pattern of relative strengths and weaknesses rather than any single number in isolation. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "Rehabilitation, settings and significance"},
    {"t": "ul", "items": [
        "Beyond diagnosis, clinicians plan **cognitive rehabilitation**, using retraining and compensatory strategies that draw on the brain's plasticity to improve everyday functioning after injury. {{c:CerebLoc}}",
        "They work in **hospitals, rehabilitation units, private practice and forensic settings**, often within an interdisciplinary team, and may give expert opinion in medico-legal cases. {{c:HistNP}}",
        "The significance of the clinical approach is that it turns the structure\u2013function map into **practical diagnosis, prognosis and treatment**, directly helping patients recover function. {{c:Lezak}}",
    ]},
    {"t": "box", "kind": "key", "title": "Clinical methods at a glance", "items": [
        "Clinical neuropsychology = assessment, diagnosis and rehabilitation of brain-injured patients.",
        "Foundations: the case/lesion method, interview, behavioural observation and standardised tests with norms.",
        "Fixed batteries: Halstead-Reitan, Luria-Nebraska, NIMHANS (India); broad and comparable.",
        "Flexible batteries: tests chosen to test hypotheses; efficient and targeted.",
        "Domains sampled across attention, memory, language, praxis, perception, executive function; then rehabilitation.",
    ]},
    {"t": "mn", "word": "ADR-FF", "text": "Assess, Diagnose, Rehabilitate \u2014 Fixed or Flexible \u2014 the clinical neuropsychologist Assesses, Diagnoses and Rehabilitates using either Fixed batteries (Halstead-Reitan, Luria-Nebraska, NIMHANS) or a Flexible, hypothesis-driven selection of tests."},
])


# ========================= 7. METHODS OF COGNITIVE NEUROPSYCHOLOGISTS ===========================
_s("t7", "7", "Approaches and methodologies of cognitive neuropsychologists", [
    {"t": "lead", "text": "Cognitive neuropsychology studies brain-damaged people to understand the normal mind, using information-processing models, the single-case method and the logic of double dissociation to argue that the mind is built of separable modules, complemented by modern electrophysiological (EEG, ERP) and haemodynamic (PET, fMRI) imaging and by stimulation methods such as TMS."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what cognitive neuropsychology is and how it differs from the clinical branch.",
        "Background: its roots in experimental and clinical neuropsychology and its assumptions.",
        "Key concepts: information-processing models, single-case method, modularity.",
        "Core logic: single and double dissociation.",
        "Modern methods: EEG/ERP, PET/SPECT, fMRI and TMS.",
        "Significance and a one-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "cognitive_methods", "caption": "The cognitive neuropsychologist's method: an information-processing model whose separable boxes are tested by single cases and double dissociations, now cross-checked by EEG/ERP, PET and fMRI imaging and by TMS."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Cognitive neuropsychology is described as a distillation of **experimental and clinical neuropsychology**, seeking to understand the mind and brain by studying people with brain injury or neurological illness. {{c:HistNP}}",
        "{{GUIDE}} One line of work investigates how the **pattern of errors** produced by brain-damaged individuals can constrain our understanding of mental representations and processes. {{c:HistNP}}",
        "{{THEORY}} The approach assumes **functional modularity**, that the mind is made of specialised, separable components, so that focal damage can impair one component while sparing others. {{c:Kolb}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and distinctiveness"},
    {"t": "ul", "items": [
        "Cognitive neuropsychology uses the **breakdown of cognition after brain damage** to work out the structure of the normal, healthy mind, rather than mainly to diagnose or treat the patient. {{c:HistNP}}",
        "It differs from the clinical branch in its **primary goal**, which is theoretical, using patients as natural experiments that reveal how intact cognition is organised. {{c:Kolb}}",
        "It emerged as a **distillation of experimental and clinical** neuropsychology, borrowing the experimentalist's models and the clinician's patients. {{c:HistNP}}",
    ]},
    {"t": "h4", "text": "Information-processing models and single cases"},
    {"t": "ul", "items": [
        "The approach builds [[information-processing models|box-and-arrow diagrams of the separable stages of a mental task]], which treat a cognitive ability as a flow of distinct stages that can be mapped and then selectively broken. {{c:Kolb}}",
        "It relies on the **single-case method**, studying one informative patient in depth, because averaging across patients with different lesions can blur the very dissociations that matter. {{c:Lezak}}",
        "Rich single cases in the clinical tradition, such as those described by **Oliver Sacks**, show how a focal deficit can fractionate the mind in striking and theoretically telling ways. {{c:Sacks}}",
    ]},
    {"t": "h4", "text": "Modularity and the logic of dissociation"},
    {"t": "ul", "items": [
        "The key assumption is **modularity**, the view that the mind is composed of specialised processors, so that a single process can be knocked out while the rest of cognition continues. {{c:Kolb}}",
        "A [[single dissociation|one patient is impaired on task A but normal on task B]] suggests the two tasks use at least partly different processes, but it can also arise if one task is simply harder. {{c:Kolb}}",
        "A [[double dissociation|one patient fails A but not B, another fails B but not A]] is the strong proof, because it shows the two abilities depend on **separable systems** rather than on shared difficulty. {{c:Kolb}}",
        "The classic example is the double dissociation between **Broca's and Wernicke's aphasias**, where one spares comprehension but ruins production and the other does the reverse, implying separate language components. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Modern electrophysiological methods"},
    {"t": "ul", "items": [
        "[[Electroencephalography|EEG, recording the brain's electrical activity from scalp electrodes]] records ongoing electrical activity with excellent **timing**, and its magnetic counterpart is magneto-encephalography. {{c:ImgTech}}",
        "[[Event-related potentials|ERPs, averaged brain waves time-locked to a stimulus]] are slow waves produced by averaging many trials, which lifts the signal out of background noise and reveals the brain's response to each event. {{c:ImgTech}}",
        "ERP components are named by **polarity and timing**, such as N100 or P300, and the N400 was the first component linked to **language**, growing larger when a word is semantically surprising. {{c:ImgTech}}",
        "Because ERPs lift the signal out of noise by **averaging many trials**, the amplitude of a component such as the N400 tracks processing difficulty, so harder or more surprising items give a larger wave. {{c:ImgTech}}",
    ]},
    {"t": "h4", "text": "Modern haemodynamic and stimulation methods"},
    {"t": "ul", "items": [
        "[[Positron emission tomography|PET, which images blood flow or metabolism using a radioactive tracer]] images function from blood flow and glucose metabolism, and uses **subtraction**, taking a rest image away from a task image to reveal the regions a task engages. {{c:ImgTech}}",
        "[[Functional magnetic resonance imaging|fMRI, which images activity from blood-oxygen changes]] detects the **blood-oxygen-level-dependent** signal that follows neural activity, giving good spatial detail without any radioactive tracer. {{c:ImgTech}}",
        "A general trade-off holds that **EEG and ERP have fine timing but coarse location**, whereas PET and fMRI have fine location but slower timing, so methods are chosen to complement one another. {{c:ImgTech}}",
        "[[Magneto-encephalography|MEG, recording the tiny magnetic fields produced by neural activity]] is the non-invasive magnetic counterpart of EEG, sharing its fine timing while locating sources somewhat more precisely. {{c:CerebLoc}}",
        "[[Diffusion tensor imaging|DTI, which traces white-matter tracts by the direction of water diffusion]] reconstructs the brain's fibre pathways, so structural connectivity itself can now be mapped and related to function in the living brain. {{c:StructFn}}",
        "[[Transcranial magnetic stimulation|TMS, a magnetic pulse that briefly disrupts or excites a cortical area]] can create a brief, reversible 'virtual lesion', letting researchers test a region's causal role without permanent damage. {{c:CerebLoc}}",
    ]},
    {"t": "table", "caption": "Imaging methods compared on what they measure and their resolution.", "head": ["Method", "Signal measured", "Strength"], "rows": [
        ["EEG / MEG", "electrical / magnetic activity", "excellent timing (ms)"],
        ["ERP", "averaged waves time-locked to a stimulus", "times cognitive stages"],
        ["PET", "blood flow / glucose metabolism (tracer)", "function, uses subtraction"],
        ["fMRI", "blood-oxygen (BOLD) signal", "fine spatial location"],
        ["DTI", "water diffusion along fibres", "maps white-matter wiring"],
        ["TMS", "magnetic pulse to the cortex", "tests a causal role ('virtual lesion')"],
    ]},
    {"t": "h4", "text": "Significance and a caution"},
    {"t": "ul", "items": [
        "These approaches are **not mutually exclusive**, and most neuropsychologists combine lesion evidence, cognitive models and imaging to converge on how a function is organised. {{c:HistNP}}",
        "The significance of cognitive neuropsychology is that it turns patients into a **theory of the normal mind**, grounding psychology's models of memory, language and perception in the architecture of the brain. {{c:Kolb}}",
    ]},
    {"t": "box", "kind": "key", "title": "Cognitive methods at a glance", "items": [
        "Cognitive neuropsychology studies patients to model the normal mind (a theoretical, not mainly clinical, goal).",
        "Tools: information-processing models, the single-case method, the modularity assumption.",
        "Core logic: double dissociation proves separable systems (e.g. Broca's vs Wernicke's aphasia).",
        "Electrophysiology: EEG/MEG and ERPs (N100, P300, N400) \u2014 fine timing, coarse location.",
        "Haemodynamic/stimulation: PET (subtraction) and fMRI (BOLD) \u2014 fine location; TMS gives a 'virtual lesion'.",
    ]},
    {"t": "mn", "word": "MIDE-T", "text": "Modularity, Information-processing, Dissociation, ERP/imaging, TMS \u2014 cognitive neuropsychology assumes Modularity, draws Information-processing models, proves them by double Dissociation, and cross-checks with ERP/PET/fMRI imaging and TMS 'virtual lesions'."},
])
