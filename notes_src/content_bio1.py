# -*- coding: utf-8 -*-
"""Single source of truth for Biological Foundations of Behavior, Unit I
(Anatomy of the Brain).

The HTML build (build_bio1.py -> parts) imports SECTIONS and the front-matter
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
  {"t":"ul","items":[str | (str,[str,...]) ]}      # optional one level of sub-bullets
  {"t":"ol","items":[...]}
  {"t":"table","caption":str|None,"head":[...],"rows":[[...],...]}
  {"t":"box","kind":"def|key|law|ex|warn|info","title":str,"items":[...] , "text":str|None}
  {"t":"fig","key":str,"caption":str}
  {"t":"mn","word":str,"text":str}                  # mnemonic
  {"t":"note","text":str}

TITLE and CREDIT are the byte-exact brand strings.
"""

TITLE = "Biological Foundations of Behavior \u2014 Unit I: Anatomy of the Brain"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

SUBTITLE = "M.Phil. Clinical Psychology study notes (bullet style) covering the anatomy of the brain"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens the notes.
INTRO = [
    {"t": "lead", "text": "These are concise, exam-ready revision notes for Unit I of the Biological Foundations of Behavior paper (M.Phil. Clinical Psychology). They cover the anatomy of the brain topic by topic, from its major sub-divisions and the surface of the cortex to its interior structures, connectivity, blood supply, fluid system and microscopic organization, so a whole 5-mark answer can be revised from a single topic."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one Unit I syllabus topic, written as a 5-mark exam answer in full sentences rather than terse fragments, so you can revise straight from the point and reproduce it under exam conditions.",
        "Every topic opens with a one-line summary, carries a flowchart figure that pictures the core idea, gives a short definition box drawn from standard neuroanatomy texts, and closes with a mnemonic that packs the key points into a single memorable word.",
        "A short source note and a references list sit at the end; the references name only the books (full title and author) and the online sources that were actually used.",
    ]},
    {"t": "h3", "text": "How the coloured gloss works (live example)"},
    {"t": "ul", "items": [
        "Hard anatomical terms are kept, but each one is explained right after it in a coloured bracket, like this: [[gyrus|a raised ridge or fold on the brain's surface]]. Just skim the amber bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "The term itself carries a faint dotted underline and its plain-language meaning follows in amber italics on a peach tint, as in [[sulcus|a groove or valley between two folds of the brain]]; the two together let you learn the real term and its meaning at once.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail, and note that only a few hook words are bold, never a whole sentence.",
        "Every topic carries one flowchart figure that pictures the structure and one mnemonic that packs its key points into a single memorable word, so you get a visual anchor and a verbal anchor for each answer.",
    ]},
    {"t": "h3", "text": "How citations work"},
    {"t": "ul", "items": [
        "A citation in square brackets names the full book and its author, for example {{c:Nolte}} or {{c:Snell}}, so you can trace any claim back to the standard text it comes from.",
        "Where a fact was confirmed from an authoritative online reference it is named in full, for example {{c:StatPearls CoW}}, and the primary source text is cited by its real title and editor, {{c:Prasad}}.",
    ]},
]

# Honest limits and source conventions.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no authors, citations or page numbers in these notes are invented. The anatomical facts stated are standard, textbook-level and verifiable, and the following limits are disclosed so you can check them independently."},
    {"t": "ul", "items": [
        "The characterisations of key structures are drawn from standard neuroanatomy wording (Nolte, Snell, Kandel and colleagues, Carlson, FitzGerald) and from the primary source text by Barre Vijaya Prasad; they are given as characterisations, not verbatim quotations, and no page numbers are attached.",
        "Where the primary source text was thin, namely the circle of Willis, the detailed cortical layers and parts of the basal-ganglia and connectivity material, well-known facts were confirmed from authoritative open references such as the NCBI StatPearls neuroanatomy chapters and standard editions of Gray's Anatomy, and these are named in the references.",
        "Only well-established anatomy is stated; no novel claims, study results or quantitative findings beyond standard textbook figures are presented, and where a detail varies between individuals (for example the many normal variants of the circle of Willis) that variability is noted rather than hidden.",
    ]},
]

# Reading list.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the standard neuroanatomy and neuroscience books used, the primary source text, and the authoritative online references consulted for the Unit I anatomy content."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "John Nolte \u2014 The Human Brain: An Introduction to Its Functional Anatomy (functional organisation, sub-divisions, cortex, pathways).",
        "Richard S. Snell \u2014 Clinical Neuroanatomy (gross anatomy of the cortex, sub-cortical structures, meninges, ventricles and blood supply).",
        "Eric R. Kandel, James H. Schwartz & Thomas M. Jessell \u2014 Principles of Neural Science (cortical organisation, connectivity, cytoarchitecture and cortical columns).",
        "Neil R. Carlson \u2014 Physiology of Behavior (the brain's divisions and structures related to behaviour, for the psychology reader).",
        "M. J. T. FitzGerald \u2014 Clinical Neuroanatomy and Neuroscience (regional and systems neuroanatomy with clinical correlation).",
        "Barre Vijaya Prasad (2020) \u2014 Examining Biological Foundations of Human Behavior (the primary source text that anchors the subdivisions, lobes, thalamus, limbic system, ventricles and neuron material).",
    ]},
    {"t": "h3", "text": "Online and standard-reference sources"},
    {"t": "ul", "items": [
        "StatPearls (NCBI Bookshelf) \u2014 Neuroanatomy: Circle of Willis; Neuroanatomy: Cerebral Blood Supply (the anterior and posterior circulations and the communicating arteries).",
        "StatPearls (NCBI Bookshelf) \u2014 Neuroanatomy: Cerebrospinal Fluid (production, the ventricular route, the apertures of Magendie and Luschka, and absorption).",
        "StatPearls (NCBI Bookshelf) \u2014 Neuroanatomy: Basal Ganglia (the striatum, globus pallidus, subthalamic nucleus and substantia nigra).",
        "Gray's Anatomy (standard edition) \u2014 the arteries, meninges and venous sinuses of the brain.",
        "Korbinian Brodmann \u2014 Localisation in the Cerebral Cortex (the cytoarchitectonic map of roughly fifty-two numbered areas).",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. MAJOR ANATOMICAL SUB-DIVISIONS ===========================
_s("t1", "1", "Major anatomical sub-divisions of the human brain", [
    {"t": "lead", "text": "The human brain develops from a simple neural tube into three primary divisions, forebrain, midbrain and hindbrain, which further form five vesicles whose adult derivatives are the cerebrum, diencephalon, midbrain, pons with cerebellum, and medulla, all set within the central nervous system."},
    {"t": "fig", "key": "brain_subdivisions", "caption": "From the neural tube to the five embryonic vesicles and their adult derivatives, oriented within the central and peripheral nervous systems."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Standard neuroanatomy describes the brain as part of the [[central nervous system|the brain and spinal cord together]], which with the peripheral nerves forms the whole nervous system, and divides it embryologically into **forebrain, midbrain and hindbrain**. {{c:Nolte}}",
        "{{GUIDE}} Standard descriptions note that the brain forms from five regions, **telencephalon, diencephalon, mesencephalon, metencephalon and myelencephalon**, arranged from rostral to caudal during development. {{c:Prasad}}",
        "{{GUIDE}} The three primary vesicles are named [[prosencephalon|the forebrain vesicle]], [[mesencephalon|the midbrain vesicle]] and [[rhombencephalon|the hindbrain vesicle]], and each gives rise to defined adult structures. {{c:Snell}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "The brain begins in the embryo as a hollow **neural tube**, whose front end swells into **three primary vesicles**, the forebrain, the midbrain and the hindbrain, from which every adult region is ultimately derived. {{c:Nolte}}",
        "The **forebrain** or [[prosencephalon|the most rostral, front-most brain vesicle]] later splits into the **telencephalon and diencephalon**, making it the source of the cerebrum and of the deep structures around the third ventricle. {{c:Snell}}",
        "The **telencephalon** becomes the [[cerebrum|the large, folded cerebral hemispheres]], the **largest part** of the brain, which carries the cortex, the underlying white matter and the basal ganglia within each hemisphere. {{c:Prasad}}",
        "The **diencephalon** forms the central core of the forebrain and gives rise to the **thalamus, hypothalamus and epithalamus**, the last of which carries the pineal gland and encloses the third ventricle. {{c:Prasad}}",
        "The **midbrain** or [[mesencephalon|the short middle brain segment]] stays **undivided**, becoming a short segment between the diencephalon and the pons that carries reflex and relay centres. {{c:Snell}}",
        "The **hindbrain** or [[rhombencephalon|the most caudal brain vesicle, nearest the spinal cord]] divides into the **metencephalon and myelencephalon**, so it is the source of the pons, the cerebellum and the medulla. {{c:Nolte}}",
        "The **metencephalon** develops into the **pons and the cerebellum**, placing the great coordinating cerebellum and the bridge-like pons together in the upper hindbrain just above the medulla. {{c:Snell}}",
        "The **myelencephalon** becomes the **medulla oblongata**, the lowest part of the brainstem, which continues directly into the spinal cord and houses vital cardiovascular and respiratory centres. {{c:Prasad}}",
        "By function the brain is grouped into the **brainstem** (midbrain, pons and medulla), the **diencephalon**, the **cerebrum** and the **cerebellum**, which is a convenient adult map laid over the embryonic scheme. {{c:Prasad}}",
        "The **brainstem** is the oldest part of the brain and runs **automatic survival functions**, and its net-like [[reticular formation|a diffuse network of neurons through the brainstem core]] filters incoming stimuli and governs arousal. {{c:Prasad}}",
        "The **cerebellum**, the 'little brain', sits behind the brainstem and fine-tunes **balance** and **posture**, smoothing movement, working on the same side of the body that it helps to control. {{c:Carlson}}",
        "Anatomical **orientation** uses paired terms, so **rostral means toward the front** and caudal toward the tail, dorsal means toward the back and ventral toward the belly, letting any structure be located precisely. {{c:Nolte}}",
        "The whole brain is bilaterally symmetrical and weighs roughly **1250 to 1450 grams** in adults, sitting protected inside the bony **cranium** and wrapped in the three meningeal membranes. {{c:Prasad}}",
        "The brain and the spinal cord together make up the **central nervous system**, while the cranial and spinal nerves branching out to the body form the **peripheral nervous system** that connects it to the periphery. {{c:Carlson}}",
        "A deep **interhemispheric fissure** splits the cerebrum into **left and right hemispheres**, and the two are joined underneath by the corpus callosum so that they can share information. {{c:Prasad}}",
        "Historically, **Franz Josef Gall** tried to localise mental faculties to brain regions and **Paul Broca** in the 1860s located a speech region, early steps that founded the discipline of modern neuroanatomy. {{c:Prasad}}",
        "Clinically the scheme matters, because damage at different levels produces different pictures, so a **brainstem lesion** threatens **vital functions** while a cerebral lesion affects thought, movement or sensation instead. {{c:Snell}}",
    ]},
    {"t": "fig", "key": "midsagittal_view", "caption": "Midsagittal (medial) view: the corpus callosum and fornix, the thalamus and hypothalamus around the third ventricle, the midbrain, pons and medulla of the brainstem descending to the spinal cord, the cerebellum, and the ventricular outline."},
    {"t": "box", "kind": "key", "title": "Sub-divisions at a glance", "items": [
        "Neural tube \u2192 three primary vesicles: forebrain, midbrain, hindbrain.",
        "Five secondary vesicles: telencephalon, diencephalon, mesencephalon, metencephalon, myelencephalon.",
        "Adult derivatives: cerebrum; thalamus and hypothalamus; midbrain; pons and cerebellum; medulla.",
        "Functional map: brainstem, diencephalon, cerebrum, cerebellum; CNS = brain + spinal cord.",
        "Orientation: rostral\u2013caudal, dorsal\u2013ventral, medial\u2013lateral locate every structure.",
    ]},
    {"t": "mn", "word": "TDMMM", "text": "Telencephalon, Diencephalon, Mesencephalon, Metencephalon and Myelencephalon \u2014 the five vesicles from front to back; remember they group under forebrain, midbrain and hindbrain, giving cerebrum, diencephalon, midbrain, pons-plus-cerebellum and medulla."},
])

# ========================= 2. SURFACE ANATOMY OF THE CORTEX ======================
_s("t2", "2", "Surface anatomy of cortical regions", [
    {"t": "lead", "text": "The surface of the cerebral cortex is thrown into folds and grooves and divided by the central sulcus and lateral fissure into four lobes, frontal, parietal, temporal and occipital, plus the hidden insula, each carrying primary motor, sensory, visual or auditory areas among wide zones of association cortex."},
    {"t": "fig", "key": "cortical_lobes", "caption": "Surface map of the left hemisphere: the four lobes and the insula, the central sulcus and lateral fissure, and the primary motor, sensory, visual and auditory areas."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} The **cerebral cortex** is characterised as the outermost layer of grey matter, a wrinkled sheet of tightly packed neurons responsible for higher processes such as thought, language, decision-making and problem-solving. {{c:Prasad}}",
        "{{GUIDE}} Standard texts divide the cortex into four lobes, **frontal, parietal, temporal and occipital**, separated by the central sulcus and the lateral fissure, with a fifth lobe, the insula, buried deep within the lateral fissure. {{c:Snell}}",
        "{{THEORY}} **Hitzig and Fritsch** (1870) showed by electrically stimulating the motor cortex of dogs that distinct body muscles map to distinct cortical points, early evidence for the localisation of function. {{c:Prasad}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "The cortex is **folded** to pack a large surface into the skull, so its **raised ridges** are called [[gyri|the raised folds of the cortical surface]] and its grooves are called [[sulci|the valleys between the folds]], and deeper grooves are termed fissures. {{c:Snell}}",
        "Two landmark grooves organise the surface: the **central sulcus** runs down and forward dividing frontal from parietal lobe, while the **lateral (Sylvian) fissure** runs backward separating the temporal lobe below. {{c:Prasad}}",
        "The **frontal lobe** lies in front of the central sulcus and governs **voluntary movement** and **planning**, which is why damage here, as in the classic case of Phineas Gage, changes behaviour. {{c:Prasad}}",
        "The **primary motor cortex** sits on the [[precentral gyrus|the fold just in front of the central sulcus]], and its parts are arranged in **orderly fashion** so that finger, hand and arm controls lie next to one another. {{c:Prasad}}",
        "The **parietal lobe** lies behind the central sulcus and handles **touch, pressure, pain and temperature**, integrating bodily sensation with spatial awareness of where the body is in its surroundings. {{c:Snell}}",
        "The **primary somatosensory cortex** occupies the [[postcentral gyrus|the fold just behind the central sulcus]] and is **highly adaptable**, so skilled users such as Braille readers develop larger fingertip representations here. {{c:Prasad}}",
        "The **temporal lobe** lies below the lateral fissure and processes **hearing, language comprehension and memory**, receiving auditory input mostly from the opposite ear through its primary auditory cortex. {{c:Prasad}}",
        "The **primary auditory cortex** sits on the upper temporal lobe at [[Heschl's gyrus|the transverse temporal gyrus that first receives sound]] and analyses **pitch, loudness and pattern** of incoming acoustic information. {{c:Snell}}",
        "The **occipital lobe** forms the back of the brain and is devoted to **vision**, with its primary visual cortex around the calcarine sulcus processing colour, shape and movement from the opposite visual field. {{c:Prasad}}",
        "Visual input crosses over, so the **right visual field** is processed in the **left hemisphere** and the left field in the right, a crossing pattern that recurs for many sensory and motor pathways. {{c:Prasad}}",
        "The **insula** is a fifth lobe tucked deep inside the lateral fissure and is involved in **taste**, **visceral sensation** and emotional awareness, hidden from the brain's outer surface. {{c:Snell}}",
        "Most of the cortex is not primary but [[association cortex|cortex that integrates and interprets rather than simply receiving]], which **binds and interprets** sensory inputs into **meaningful experience** and action. {{c:Kandel}}",
        "**Localisation of function** means that specific abilities such as language or vision have **specific cortical locations**, a principle supported first by lesion case studies and now by modern neuroimaging. {{c:Prasad}}",
        "**Lateralisation** means the two hemispheres differ in emphasis, so in most people the **left hemisphere leads for language** and the right for spatial and holistic processing, though both cooperate closely. {{c:Carlson}}",
        "Language has two classic surface regions, **Broca's area** in the frontal lobe for speech production and **Wernicke's area** in the temporal lobe for comprehension, linked by an association fibre bundle. {{c:Snell}}",
        "The cortex is only a few millimetres thick but its heavy **folding** greatly **enlarges the surface area**, and small increases in **thickness and folding** are associated with greater processing capacity. {{c:Prasad}}",
        "Clinically the surface map guides diagnosis, because a stroke or tumour in a given lobe produces **predictable deficits**, such as **weakness** from motor-strip damage or visual loss from occipital damage. {{c:Snell}}",
    ]},
    {"t": "fig", "key": "lateral_hemisphere", "caption": "Lateral (side) view of the left hemisphere with the four lobes colour-coded, the insula, the central sulcus and lateral (Sylvian) fissure, the precentral (motor) and postcentral (somatosensory) gyri, and the primary visual and auditory areas with Broca and Wernicke speech areas."},
    {"t": "fig", "key": "homunculus", "caption": "The motor (precentral) and somatosensory (postcentral) homunculus: mirror-image body maps across the central sulcus, with the hand and face magnified because fine control and dense sensation demand more cortex."},
    {"t": "box", "kind": "key", "title": "Cortical surface at a glance", "items": [
        "Folds: gyri (ridges) and sulci (grooves); central sulcus and lateral fissure are the key landmarks.",
        "Four lobes plus the hidden insula: frontal (motor), parietal (sensory), temporal (hearing), occipital (vision).",
        "Primary areas: precentral (motor), postcentral (somatosensory), Heschl's (auditory), calcarine (visual).",
        "Association cortex integrates; Broca (speech) and Wernicke (comprehension) serve language.",
        "Localisation + lateralisation: functions have places, and the hemispheres specialise, especially for language.",
    ]},
    {"t": "mn", "word": "FPTO-I", "text": "Frontal, Parietal, Temporal and Occipital are the four lobes, plus the Insula hidden in the lateral fissure \u2014 pair each with its primary job: Front moves, Parietal feels, Temporal hears, Occipital sees, Insula tastes and senses the viscera."},
])

# ========================= 3. INTERIOR STRUCTURES ===================
_s("t3", "3", "Interior structures of cortical and sub-cortical regions", [
    {"t": "lead", "text": "Beneath the cortex lie the white-matter tracts and a set of deep grey-matter structures, the basal ganglia, the thalamus and hypothalamus, the limbic system with its hippocampus and amygdala, the brainstem and cerebellum, all arranged around the fluid-filled ventricles."},
    {"t": "fig", "key": "subcortical_structures", "caption": "Working inward from the cortex through the white matter to the basal ganglia, diencephalon, limbic system, brainstem, cerebellum and the ventricles."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Standard texts distinguish [[grey matter|tissue rich in neuron cell bodies]], found in the surface cortex and the deep nuclei, from [[white matter|tissue made of myelinated axons]], the connecting fibres that lie beneath. {{c:Snell}}",
        "{{GUIDE}} The **basal ganglia** are characterised as a group of deep forebrain nuclei, the striatum, globus pallidus, subthalamic nucleus and substantia nigra, which tune voluntary movement. {{c:StatPearls BG}}",
        "{{GUIDE}} Standard texts describe the **limbic system** as a doughnut-shaped ring of structures at the border of the brainstem and cerebrum, including the hippocampus, amygdala and hypothalamus, serving emotion and drive. {{c:Prasad}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "Directly under the cortex lies the **cerebral white matter**, a dense mass of **myelinated fibres** that carries signals between cortical areas and between the cortex and the deeper and lower centres of the brain. {{c:Kandel}}",
        "The **basal ganglia** are deep grey-matter nuclei that **fine-tune voluntary movement**, receiving plans from the cortex, processing them and returning instructions to the cortex by way of the thalamus. {{c:StatPearls BG}}",
        "The main basal-ganglia parts are the [[striatum|the caudate nucleus and putamen together]], the **globus pallidus**, the subthalamic nucleus and the **substantia nigra**, whose loss of dopamine cells causes Parkinson's disease. {{c:StatPearls BG}}",
        "The **thalamus** is a pair of egg-shaped grey masses forming the walls of the third ventricle, and it acts as the brain's great **relay station** through which nearly all sensory input passes on its way to the cortex. {{c:Prasad}}",
        "Because of this gateway role the thalamus is central to **sensation**, motor control, arousal and **memory**, so that damage to it can blunt or distort incoming experience. {{c:Prasad}}",
        "The **hypothalamus** sits just below the thalamus and is the master controller of the **autonomic** and **endocrine systems**, regulating hunger, thirst, temperature, sleep and the drives for sex and survival. {{c:Carlson}}",
        "The **limbic system** is the emotional core of the brain, a border ring that links the cortex with deeper structures and underlies **emotion, motivation, memory and drive** for food and sex. {{c:Prasad}}",
        "The [[hippocampus|a curved limbic structure deep in the temporal lobe]] is essential for forming new **long-term memories**, **consolidating** experiences for weeks before they are stored more permanently in the cortex. {{c:Prasad}}",
        "The [[amygdala|an almond-shaped limbic nucleus]] processes **fear** and **emotional significance**, tagging events with feeling and triggering the body's rapid defensive responses to danger. {{c:Carlson}}",
        "The **brainstem** contains the midbrain, pons and medulla, and runs the **automatic survival functions** of heartbeat, breathing and consciousness, acting as the junction through which most ascending and descending tracts pass. {{c:Prasad}}",
        "The **midbrain** carries the superior and inferior colliculi, which serve as **visual and auditory reflex centres**, together with the substantia nigra that links functionally to the basal ganglia. {{c:Prasad}}",
        "The **pons** bridges the two cerebellar hemispheres and helps coordinate **movement, sleep, waking and dreaming**, while the **medulla** below it controls the force and rate of the heartbeat and of breathing. {{c:Prasad}}",
        "The **cerebellum** behind the brainstem has two folded hemispheres and smooths and coordinates **movement, posture and balance**, correcting errors so that actions are accurate and well timed. {{c:Carlson}}",
        "The [[ventricles|fluid-filled cavities deep within the brain]] are an interconnected web of spaces holding cerebrospinal fluid, comprising **two lateral ventricles**, a third and a **fourth** linked by narrow channels. {{c:Prasad}}",
        "Grey and white matter are arranged in a clear pattern, with **grey matter** on the surface and in the **deep nuclei** and the connecting white matter sandwiched between them, a layout repeated through the brain. {{c:Snell}}",
        "The whole interior is **modular yet connected**, so each deep structure does a **specialised job** within defined pathways, and damage to one, such as the basal ganglia or hippocampus, produces a characteristic deficit. {{c:Kandel}}",
        "Clinically these interior structures matter greatly, since **Parkinson's disease** and **movement disorders** each trace to specific deep structures such as the substantia nigra or the hippocampus. {{c:StatPearls BG}}",
    ]},
    {"t": "fig", "key": "coronal_section", "caption": "Coronal section through the basal ganglia: the lateral ventricles, caudate nucleus, putamen and globus pallidus (together the lentiform nucleus), the thalamus, the internal capsule between them, and the corpus callosum over the midline."},
    {"t": "box", "kind": "key", "title": "Interior structures at a glance", "items": [
        "White matter = connecting fibres beneath the cortex; grey matter = cell bodies in cortex and deep nuclei.",
        "Basal ganglia (striatum, globus pallidus, subthalamic nucleus, substantia nigra) tune movement.",
        "Thalamus relays sensation; hypothalamus runs the autonomic and endocrine systems and basic drives.",
        "Limbic system (hippocampus, amygdala) drives memory and emotion; brainstem runs vital functions.",
        "Cerebellum coordinates movement and balance; ventricles hold the cerebrospinal fluid.",
    ]},
    {"t": "mn", "word": "BLT-HeavenC", "text": "Basal ganglia, Limbic system and Thalamus/hypothalamus, with the Hippocampus and amygdala inside the limbic ring, plus the Brainstem and Cerebellum \u2014 the deep structures under the cortex, all set around the fluid-filled ventricles."},
])

# ========================= 4. ANATOMICAL CONNECTIVITY =============================
_s("t4", "4", "Anatomical connectivity among regions", [
    {"t": "lead", "text": "The regions of the brain are wired together by white-matter fibres of three kinds, association fibres that link areas within one hemisphere, commissural fibres such as the corpus callosum that join the two hemispheres, and projection fibres such as the internal capsule that connect the cortex to lower centres."},
    {"t": "fig", "key": "fibre_types", "caption": "The three white-matter fibre classes: association fibres within a hemisphere, commissural fibres (the corpus callosum) between hemispheres, and projection fibres (the internal capsule) to and from lower centres."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Standard texts classify cerebral white-matter tracts into three kinds, **association, commissural and projection** fibres, according to the direction and reach of the connections they make. {{c:Snell}}",
        "{{GUIDE}} The **corpus callosum** is characterised as the largest white-matter structure and the major commissure, carrying on the order of two hundred million axons between the two hemispheres. {{c:Kandel}}",
        "{{GUIDE}} [[Association fibres|fibres linking regions inside one hemisphere]] include the arcuate and superior longitudinal fasciculi, the cingulum and the uncinate fasciculus, each joining defined cortical areas. {{c:Kandel}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "White matter is made of **myelinated axons**, called **tracts** or fasciculi, which carry nerve impulses between grey-matter regions and so coordinate communication across the whole brain. {{c:Kandel}}",
        "**Association fibres** connect different areas **within the same hemisphere**, allowing one lobe to talk to another, for example linking the visual cortex to the areas that interpret what is seen. {{c:Snell}}",
        "**Short** association fibres called [[U-fibres|short arching fibres joining neighbouring gyri]] loop between adjacent gyri, while **long** bundles span whole lobes within a hemisphere. {{c:Kandel}}",
        "The **superior longitudinal (arcuate) fasciculus** is a key long association bundle that links the frontal lobe with the other lobes, famously joining **Broca's and Wernicke's** language areas. {{c:Kandel}}",
        "Other named association bundles include the **cingulum**, which runs within the limbic lobe, and the **uncinate fasciculus**, which hooks the frontal lobe to the anterior temporal lobe. {{c:Snell}}",
        "**Commissural fibres** cross the midline to connect **matching areas** in the **two hemispheres**, so that the left and right sides of the brain can share and integrate their information. {{c:Snell}}",
        "The **corpus callosum** is the great commissure and the **largest white-matter structure** in the brain, a broad arched band of around two hundred million fibres joining the hemispheres front to back. {{c:Kandel}}",
        "A smaller commissure, the **anterior commissure**, links the two temporal lobes, and cutting the callosum, as in **split-brain surgery**, leaves the hemispheres unable to share information directly. {{c:Carlson}}",
        "**Projection fibres** run **vertically** between the cortex and lower centres, carrying motor commands down to the brainstem and spinal cord and sensory signals up from the thalamus to the cortex. {{c:Snell}}",
        "The great **projection gateway** is the [[internal capsule|a compact band of projection fibres between the basal ganglia]], a **dense band** that funnels most ascending and descending fibres past the basal ganglia and thalamus. {{c:Snell}}",
        "Above the internal capsule the fibres **fan out** into the [[corona radiata|the fanned-out crown of projection fibres]], which **spreads the tracts** widely to and from the whole cortical surface. {{c:Kandel}}",
        "The chief descending projection tract is the **corticospinal (pyramidal) tract**, which carries **voluntary motor commands** from the motor cortex and crosses to the opposite side in the medulla's pyramidal decussation. {{c:Prasad}}",
        "Because of this crossing, the **left hemisphere** controls the **right side** of the body and the right hemisphere the left, a decussation that explains why a one-sided stroke weakens the opposite limbs. {{c:Prasad}}",
        "**Sensory projection fibres** carry touch and pain up through the brainstem and thalamus to the cortex, for example the **medial lemniscus** relaying fine touch and position sense. {{c:Prasad}}",
        "Between the forebrain and the cerebellum run **great fibre bundles** called the **cerebellar peduncles**, through which the cerebellum receives and returns the signals that coordinate movement. {{c:Snell}}",
        "Together the three fibre classes make the brain a **single connected network**, not **isolated parts**, so that complex behaviour depends on intact wiring as much as on intact regions. {{c:Kandel}}",
        "Clinically, **broken connections** cause **disconnection syndromes**, so cutting the arcuate fasciculus disrupts repetition of speech and damaging the callosum stops the hemispheres from coordinating. {{c:Carlson}}",
    ]},
    {"t": "box", "kind": "key", "title": "Connectivity at a glance", "items": [
        "Three fibre classes: association (within a hemisphere), commissural (between hemispheres), projection (cortex \u2194 below).",
        "Association: U-fibres, arcuate/superior longitudinal fasciculus, cingulum, uncinate fasciculus.",
        "Commissural: corpus callosum (the largest) and the anterior commissure.",
        "Projection: internal capsule and corona radiata; corticospinal tract crosses in the medulla.",
        "Intact wiring is essential; broken tracts cause disconnection syndromes.",
    ]},
    {"t": "mn", "word": "ACP", "text": "Association (within a hemisphere), Commissural (between hemispheres) and Projection (up and down to lower centres) \u2014 the three fibre types; anchor each to its star tract: Arcuate fasciculus, Corpus callosum and the internal-capsule Projection fibres."},
])

# ========================= 5. BLOOD SUPPLY ===========================
_s("t5", "5", "Blood supply to the brain", [
    {"t": "lead", "text": "The brain is fed by two arterial systems, the internal carotid arteries in front and the vertebrobasilar system behind, which join at the base of the brain in the circle of Willis, a protective ring from which the major cerebral arteries arise, while venous blood drains through the dural sinuses to the internal jugular veins."},
    {"t": "fig", "key": "circle_of_willis", "caption": "The circle of Willis joins the internal carotid (anterior) and vertebrobasilar (posterior) systems through the anterior and posterior communicating arteries, giving collateral flow."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} The [[circle of Willis|an arterial ring at the base of the brain]] is characterised as an anastomotic ring joining the anterior and posterior circulations and providing collateral flow if one vessel is blocked. {{c:StatPearls CoW}}",
        "{{GUIDE}} Standard references describe the brain's supply as two paired systems, the **internal carotid** arteries in front and the **vertebral arteries** behind that fuse into the basilar artery. {{c:Gray}}",
        "{{GUIDE}} Standard physiology notes that the brain uses about **a quarter of the body's glucose and much of its oxygen**, so any fall in blood supply quickly threatens it and can cause a stroke. {{c:Prasad}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "The brain has a very high metabolic demand and receives its blood from **two paired arterial systems**, the **internal carotid** arteries supplying the front and the vertebral arteries supplying the back and base. {{c:Gray}}",
        "The **internal carotid arteries** form the [[anterior circulation|the front blood supply from the carotid arteries]], supplying most of the **cerebral hemispheres** through the anterior and middle cerebral arteries. {{c:StatPearls CoW}}",
        "The two **vertebral arteries** climb through the neck and **fuse** into a single [[basilar artery|the midline artery formed by the two vertebral arteries]], together forming the posterior circulation. {{c:StatPearls CoW}}",
        "The **vertebrobasilar system** supplies the **brainstem**, cerebellum and the back of the cerebrum, including the occipital lobe, through the posterior cerebral arteries that branch from the basilar artery. {{c:Gray}}",
        "At the base of the brain these two systems meet in the **circle of Willis**, a **polygonal arterial ring** that encircles the pituitary stalk and connects the anterior and posterior circulations. {{c:StatPearls CoW}}",
        "The ring is **completed in front** by the two **anterior cerebral arteries** joined by the single [[anterior communicating artery|the short bridge joining the two front arteries]], closing the front of the circle. {{c:StatPearls CoW}}",
        "On each side a [[posterior communicating artery|a side bridge joining front to back circulation]] **links** the carotid to the posterior cerebral artery, joining the two circulations **laterally**. {{c:StatPearls CoW}}",
        "The circle's protective value is its **collateral flow**, so if one vessel is blocked, blood is **rerouted around the ring** to keep the brain tissue perfused and reduce the damage. {{c:StatPearls CoW}}",
        "The three **major cerebral arteries** arise around this system, each watering a **defined territory** of the cortex and the deep structures beneath of the cortex and the deep structures beneath it. {{c:Gray}}",
        "The **middle cerebral artery** is the **largest** and supplies much of the lateral cortex including the language areas, so its blockage causes a severe, common stroke. {{c:Gray}}",
        "The **anterior cerebral artery** supplies the medial frontal and parietal lobes, the region for the leg, so its occlusion typically **weakens the opposite leg**. {{c:Gray}}",
        "The **posterior cerebral artery** supplies the **occipital lobe** and lower temporal lobe, so its occlusion typically causes loss of vision in the opposite field. {{c:Gray}}",
        "The circle of Willis shows **many normal variants**, since in many people one or more segments is **thin or absent**, which can change how well collateral flow protects them. {{c:StatPearls CoW}}",
        "Guarding the supply is the [[blood-brain barrier|a selective filter between blood and brain tissue]], a **tight capillary lining** that lets nutrients pass while **shielding** the brain from blood-borne toxins. {{c:Prasad}}",
        "**Venous blood** is collected into large channels called the **dural venous sinuses**, such as the superior sagittal sinus, which run between the dural layers of the dura mater. {{c:Gray}}",
        "From these sinuses the blood drains mainly into the **internal jugular veins** to **return to the heart**, and the same sinuses also receive the cerebrospinal fluid absorbed from around the brain. {{c:Gray}}",
        "Clinically this anatomy is central, because a **stroke** follows **interrupted blood supply** or bleeding, and the deficit maps onto the artery that failed. {{c:Prasad}}",
    ]},
    {"t": "box", "kind": "key", "title": "Blood supply at a glance", "items": [
        "Two inflow systems: internal carotid (anterior) and vertebral \u2192 basilar (posterior) arteries.",
        "They meet in the circle of Willis, closed by anterior and posterior communicating arteries.",
        "Three cerebral arteries: anterior (medial/leg), middle (lateral/language, largest), posterior (occipital/vision).",
        "The ring gives collateral flow; it has many normal variants; the blood-brain barrier protects the tissue.",
        "Venous drainage: dural venous sinuses \u2192 internal jugular veins.",
    ]},
    {"t": "mn", "word": "CoW-CAMP", "text": "Circle of Willis joins the Carotid and vertebrobasilar systems; its branches are the Anterior, Middle and Posterior cerebral arteries \u2014 remember CoW = CAMP, and that the communicating arteries close the ring to give collateral flow."},
])

# ========================= 6. THE CSF SYSTEM =============================
_s("t6", "6", "The cerebrospinal-fluid (CSF) system", [
    {"t": "lead", "text": "The brain is wrapped in three meninges and bathed in cerebrospinal fluid, which is made by the choroid plexus in the ventricles, flows from the lateral ventricles through the third and fourth ventricles into the subarachnoid space, and is absorbed back into the venous blood at the arachnoid granulations, cushioning and nourishing the brain."},
    {"t": "fig", "key": "csf_flow", "caption": "CSF production at the choroid plexus, circulation through the lateral, third and fourth ventricles into the subarachnoid space, and absorption at the arachnoid granulations into the venous sinuses."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} **Cerebrospinal fluid** is characterised as a clear, colourless, watery fluid similar to blood plasma but with less protein, which cushions the brain, gives it buoyancy and carries away waste. {{c:Prasad}}",
        "{{GUIDE}} The [[meninges|the three protective membranes around the brain]] are the tough **dura mater**, the web-like **arachnoid mater** and the delicate **pia mater** that clings to the brain surface. {{c:Prasad}}",
        "{{GUIDE}} Standard references describe the CSF route from the lateral ventricles through the third and fourth ventricles, out of the apertures of **Magendie and Luschka** into the subarachnoid space. {{c:StatPearls CSF}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "The brain is protected by **three membranes** called the **meninges**, which lie between the skull and the brain, enclose the blood vessels and form the sheath around the venous sinuses. {{c:Prasad}}",
        "The outer **dura mater** is a strong, **leathery, two-layered** membrane whose layers separate to enclose the [[dural venous sinuses|venous channels within the dura]] that drain blood from the brain. {{c:Prasad}}",
        "The middle **arachnoid mater** is a loose, **web-like** membrane, and beneath it the [[subarachnoid space|the fluid-filled gap beneath the arachnoid membrane]] is filled with cerebrospinal fluid and the larger blood vessels. {{c:Prasad}}",
        "The inner **pia mater** is a delicate membrane that **clings tightly** to the brain's folded surface, following every gyrus and sulcus and carrying small blood vessels into the tissue. {{c:Snell}}",
        "**Cerebrospinal fluid** is a **clear fluid**, roughly 80 to 150 millilitres in volume, resembling blood plasma but with less protein, that surrounds the brain and spinal cord and fills the ventricles. {{c:Prasad}}",
        "The fluid is **produced** by the [[choroid plexus|tufts of capillaries that make CSF in the ventricles]], clusters of **specialised capillaries** lined by ependymal cells that hang from the roof of each ventricle. {{c:Prasad}}",
        "The **ventricular system** is an interconnected set of cavities comprising **two lateral ventricles**, a third ventricle between the thalamic halves and a **fourth ventricle** between the brainstem and cerebellum. {{c:Prasad}}",
        "From each lateral ventricle the fluid passes through the [[foramen of Monro|the channel from a lateral ventricle to the third]] into the **midline third ventricle**, beginning its **one-way journey** through the brain. {{c:StatPearls CSF}}",
        "From the third ventricle the fluid flows down the **narrow** [[cerebral aqueduct|the thin channel of Sylvius between third and fourth ventricles]] into the **fourth ventricle** lying behind the pons and medulla. {{c:StatPearls CSF}}",
        "From the fourth ventricle the fluid escapes into the subarachnoid space through **three openings**, the median aperture of **Magendie** and the two lateral apertures of **Luschka**, to bathe the whole brain and cord. {{c:StatPearls CSF}}",
        "Once in the **subarachnoid space** the fluid circulates over the brain and down around the spinal cord, forming a **continuous liquid cushion** between the delicate tissue and the bone. {{c:StatPearls CSF}}",
        "The fluid is finally **reabsorbed** into the venous blood at the [[arachnoid granulations|small projections that return CSF to the venous sinuses]], **tiny tufts** of arachnoid that push into the dural venous sinuses. {{c:StatPearls CSF}}",
        "Production, circulation and absorption are kept in **balance**, so the adult makes **half a litre** a day while holding only a small volume, the surplus being continually absorbed. {{c:StatPearls CSF}}",
        "The first function of the fluid is **mechanical protection**, giving the brain **buoyancy** so that it floats and does not crush under its own weight, and cushioning it against blows. {{c:Prasad}}",
        "Further functions are **nourishment** and **waste removal**, since the fluid carries nutrients to the brain and washes away metabolic waste, helping to keep the brain's environment stable. {{c:Prasad}}",
        "The choroid plexus also helps **regulate the chemical environment**, using **ion pumps** to adjust the composition of the fluid and so of the central nervous system, assisting in cleaning and homeostasis. {{c:Prasad}}",
        "Clinically it matters, because a **blockage** of the narrow channels backs up the fluid and **raises pressure** in a condition called [[hydrocephalus|a build-up of CSF that enlarges the ventricles]], which can damage the brain. {{c:StatPearls CSF}}",
    ]},
    {"t": "fig", "key": "ventricular_system", "caption": "The ventricular system in side view: the paired lateral ventricles, the foramen of Monro, the third ventricle, the cerebral aqueduct and the fourth ventricle, with the apertures of Magendie and Luschka opening into the subarachnoid space."},
    {"t": "box", "kind": "key", "title": "The CSF system at a glance", "items": [
        "Meninges: dura (tough, holds the sinuses), arachnoid (web-like, over the subarachnoid space), pia (clings to brain).",
        "Made by the choroid plexus in the ventricles; ~80\u2013150 ml held, ~500 ml made per day.",
        "Route: lateral ventricles \u2192 foramen of Monro \u2192 third \u2192 aqueduct \u2192 fourth \u2192 Magendie & Luschka \u2192 subarachnoid space.",
        "Absorbed at the arachnoid granulations into the dural venous sinuses.",
        "Functions: cushion and buoyancy, nourishment, waste removal, chemical stability; blockage causes hydrocephalus.",
    ]},
    {"t": "mn", "word": "DAP-MALM", "text": "Dura, Arachnoid and Pia are the meninges; then the flow runs Monro, Aqueduct, Luschka and Magendie out to the subarachnoid space \u2014 DAP then M-A-L-M \u2014 made by choroid plexus, absorbed at the arachnoid granulations."},
])

# ========================= 7. CYTOARCHITECTURE =============================
_s("t7", "7", "Cytoarchitecture and modular organization", [
    {"t": "lead", "text": "At the microscopic level the cortex is built from neurons and glial cells arranged in six horizontal layers of the neocortex, organised into vertical columns or modules, and mapped by Brodmann into about fifty-two numbered areas whose differing layer patterns match differing functions."},
    {"t": "fig", "key": "cortical_layers", "caption": "The six neocortical laminae from the pial surface to the white matter, with the vertical cortical column (module) that Brodmann's cytoarchitectonic map is built upon."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} [[Cytoarchitecture|the microscopic arrangement of cells in the cortex]] is characterised as the pattern of neuronal cell types and layering by which the cortex is divided into structurally distinct areas. {{c:Kandel}}",
        "{{THEORY}} **Korbinian Brodmann** in 1909 mapped the cortex into roughly fifty-two numbered areas based on cell structure, and these **Brodmann areas** are still used to name functional regions. {{c:Brodmann}}",
        "{{GUIDE}} The six neocortical layers are named, from the surface inward, the **molecular, external granular, external pyramidal, internal granular, internal pyramidal and multiform** layers. {{c:Kandel}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "ul", "items": [
        "The cortex is built from **two cell classes**, the information-carrying **neurons** and the supporting [[glia|non-neuronal support cells of the nervous system]], which outnumber neurons and maintain, feed and insulate them. {{c:Prasad}}",
        "The chief excitatory neuron of the cortex is the **pyramidal cell**, a triangular cell whose long axon carries the cortex's main output, while small **stellate (granule) cells** are mostly local interneurons. {{c:Kandel}}",
        "The main glial types are the star-shaped [[astrocytes|star-shaped glia that support neurons and the blood-brain barrier]], the myelin-making **oligodendrocytes**, the immune **microglia** and the ependymal cells lining the ventricles. {{c:Kandel}}",
        "Most of the cortex is **neocortex**, the youngest, **six-layered** type that makes up the great bulk of the human cortex and carries out the highest cognitive functions. {{c:Prasad}}",
        "Older **allocortex**, such as that of the hippocampus, has only **three layers**, and the six-layered neocortex evolved from this simpler cortex. {{c:Prasad}}",
        "Counting from the pial surface, **layer I** is the **molecular layer**, a cell-poor layer of dendrites and axons that allows horizontal connections across the top of the cortex. {{c:Kandel}}",
        "**Layer II**, the **external granular layer**, is packed with small pyramidal and stellate cells, and together with layer III it supplies many of the fibres linking one cortical area to another. {{c:Kandel}}",
        "**Layer III**, the **external pyramidal layer**, holds medium pyramidal cells sending cortico-cortical and commissural output, including fibres that cross in the corpus callosum. {{c:Kandel}}",
        "**Layer IV**, the **internal granular layer**, is rich in stellate cells and is the main input layer that receives sensory fibres arriving from the thalamus, so it is thick in sensory cortex. {{c:Kandel}}",
        "**Layer V**, the **internal pyramidal layer**, contains the largest pyramidal cells, including the giant Betz cells of the motor cortex, and provides the major output to the brainstem and spinal cord. {{c:Kandel}}",
        "**Layer VI, the multiform layer**, holds cells of varied shape and sends much of its output **back to the thalamus**, closing the loop between the cortex and its main relay station. {{c:Kandel}}",
        "The relative thickness of the layers varies with function, so **sensory cortex** has a thick **layer IV** (granular, input-heavy) while **motor cortex** has a thick **layer V** (pyramidal, output-heavy). {{c:Kandel}}",
        "The cortex is organised not only in **layers** but in [[cortical columns|vertical modules of cells running through all six layers]], narrow **vertical modules** running from surface to white matter whose cells share a similar response. {{c:Kandel}}",
        "These columns act as the **basic processing modules** of the cortex, repeated many times across the sheet, so that the cortex works as a **mosaic** of similar computing units. {{c:Kandel}}",
        "**Brodmann's map** divided the cortex into about **fifty-two areas** by their cell structure, and the numbers remain a standard shorthand, so that specific areas correspond to specific functions. {{c:Brodmann}}",
        "Familiar examples link a Brodmann number to a function, for instance **area 4** is the **primary motor cortex**, areas 3, 1 and 2 are the somatosensory cortex, and area 17 is the visual cortex. {{c:Prasad}}",
        "This microscopic organisation underpins the whole brain's working, because the **'minimal wiring' principle** places **co-active neurons** close to one another to save energy and reduce delay. {{c:Prasad}}",
    ]},
    {"t": "box", "kind": "key", "title": "Cytoarchitecture at a glance", "items": [
        "Cells: pyramidal (output) and stellate (local) neurons; astrocytes, oligodendrocytes, microglia, ependyma.",
        "Neocortex has six layers; older allocortex (hippocampus) has three.",
        "Six layers: I molecular, II external granular, III external pyramidal, IV internal granular, V internal pyramidal, VI multiform.",
        "Layer IV = sensory input (thick in sensory cortex); layer V = output with Betz cells (thick in motor cortex).",
        "Vertical columns = processing modules; Brodmann mapped ~52 areas (4 motor, 3-1-2 sensory, 17 visual).",
    ]},
    {"t": "mn", "word": "MEE-IIM", "text": "Molecular, External granular, External pyramidal, Internal granular, Internal pyramidal and Multiform \u2014 the six cortical layers top to bottom (MEE-IIM); remember input lands in IV and output leaves from V, and that vertical columns are the modules Brodmann mapped."},
])
