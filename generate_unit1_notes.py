# -*- coding: utf-8 -*-
"""
Generator for the self-contained offline HTML study-notes file:
"Biological Foundations of Behavior - Unit I: Anatomy of the Brain".

This module defines ONE CONTENT data structure (a list of 7 topic dicts) that is
the single source of truth for the rendered offline HTML deliverable. Keep CONTENT
clean and importable.

Markup conventions used inside point / intro / flow text:
  **word**                 -> a bold memory-hook keyword (a few per sentence only)
  [[plain-language text]]  -> a distinct bracketed plain-language explanation of a
                              difficult term, shown in its own visible style.

Run (from /projects/sandbox/Dango-kiro):
  python3 generate_unit1_notes.py
"""

import base64
import html as _html
import os
import re

TITLE = "Biological Foundations of Behavior \u2014 Unit I: Anatomy of the Brain"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"
SUBTITLE = ("Paper I: Biological Foundations of Behavior \u00b7 Part A (Anatomy, Physiology "
            "and Biochemistry of the CNS) and Part B (Neuropsychology) \u00b7 Exam-ready study notes")
OUTPUT_HTML = "Biological_Foundations_Unit1_Anatomy_of_the_Brain.html"

# ---------------------------------------------------------------------------
# HOW TO READ THESE NOTES (short usage guide; NOT about how the notes were made)
# ---------------------------------------------------------------------------
HOW_TO_READ = [
    "Each topic opens with a one-line **intro** that frames the idea, followed by "
    "at least fifteen numbered points; every point is a full exam sentence, so read it "
    "as a ready-made **5-mark answer** rather than a short note.",
    "The points follow an exam **answer-flow** order, an opening definition first, then the "
    "main labelled points in a logical sequence, then the significance or **clinical link**, "
    "and a closing line, so you can reproduce the whole answer in the same order under exam "
    "conditions.",
    "Hard anatomical terms are kept in full, but each is followed by a plain-language meaning "
    "inside a shaded box like this [[a simple everyday explanation that sits right next to the "
    "difficult word]] so you never lose the technical vocabulary.",
    "Only a **few keywords** in each sentence are shown in bold; treat those bold words as "
    "memory hooks and recall the rest of the sentence around them while you revise.",
    "Every topic carries exactly one **flowchart** of boxes joined by arrows to show the order "
    "of a structure or process, and exactly one **mnemonic** whose letters each map to a point "
    "you must remember.",
    "Most topics carry a **labelled diagram**; use each figure as a visual map, matching its "
    "labels to the numbered points and the flowchart so you can both picture and describe the "
    "structure in the exam.",
    "Revise by first reading the flowchart to fix the sequence, then the mnemonic to lock the "
    "key points, and finally the full sentences in answer-flow order to add exam-ready detail "
    "and wording.",
]

# ---------------------------------------------------------------------------
# CONTENT = single source of truth (7 topics, in the required order)
# Each topic dict:
#   anchor, number, color, title, intro, points[list], definition(optional dict),
#   figures(optional list), flow(dict: title + steps[list]),
#   mnemonic(dict: word, line, items[list of (letter, text)])
#
# NOTE: The points below are PLACEHOLDERS only so the generator runs. The full
# 7-topic exam content is authored in a later feature.
# ---------------------------------------------------------------------------
CONTENT = [
    {
        "anchor": "subdivisions",
        "number": 1,
        "color": "blue",
        "title": "Major Anatomical Sub-divisions of the Human Brain",
        "intro": "The human brain is a single soft organ but is studied as an orderly set of "
                 "sub-divisions that grow out of the embryo and mature into the brainstem, "
                 "diencephalon, cerebrum and cerebellum.",
        "points": [
            "The **brain** [[the master control organ of the nervous system]] is a soft, "
            "bilaterally symmetrical, gelatinous structure weighing about **400 grams** in "
            "infants and **1250 to 1450 grams** in adults, enclosed by the bony cranium.",
            "During **embryogenesis** [[the stage when the embryo is forming]] the brain first "
            "arises from **five regions** named the telencephalon, diencephalon, mesencephalon, "
            "metencephalon and myelencephalon, arranged from the rostral (front) to the caudal "
            "(tail) end.",
            "These five embryonic vesicles are grouped into three primary parts, the "
            "**forebrain** [[prosencephalon]], the **midbrain** [[mesencephalon]] and the "
            "**hindbrain** [[rhombencephalon]], which is the simplest scheme examiners expect "
            "you to open with.",
            "As the brain develops these early regions disappear and new adult structures emerge, "
            "so the **functional adult scheme** divides the brain into the brainstem, the "
            "diencephalon, the cerebrum and the cerebellum for study.",
            "The **brainstem** [[the stalk connecting the brain to the spinal cord]] is the "
            "oldest part of the brain, controls automatic survival functions, and is made of "
            "three parts, the medulla oblongata, the pons and the midbrain.",
            "The **medulla oblongata** [[the lowest brainstem segment just above the spinal "
            "cord]] houses the cardiovascular and respiratory control centres and carries the "
            "pyramids and the decussation where motor fibres cross to the opposite side.",
            "The **pons** [[the rounded bulge above the medulla]] acts as a bridge that helps "
            "coordinate movement and regulates sleeping, waking and dreaming, relaying signals "
            "between the cerebrum and the cerebellum below.",
            "The **midbrain** [[mesencephalon, the short segment linking hindbrain and "
            "forebrain]] contains the cerebral peduncles, the cerebral aqueduct, and colliculi "
            "that serve as visual and auditory reflex centres for the body.",
            "The **diencephalon** [[the central core of the forebrain around the third "
            "ventricle]] is made of three paired structures, the thalamus, the hypothalamus and "
            "the epithalamus, which carries the pineal gland.",
            "The **thalamus** [[the brain's great sensory relay station]] converges and relays "
            "almost all sensory and motor signals heading to the cortex, so damage to it "
            "disrupts sensation, arousal and memory across the whole body.",
            "The **hypothalamus** [[the small control centre below the thalamus]] governs hunger, "
            "thirst, temperature, emotion and sexual behaviour, and regulates hormones by "
            "directing the pituitary gland sitting just beneath it.",
            "The **cerebrum** [[the large wrinkled uppermost part of the brain]] forms about "
            "**83 percent** of total brain volume, is split into two hemispheres, and is the "
            "seat of logic, memory, language and voluntary, complex behaviour.",
            "The **cerebellum** [[the little brain behind the brainstem]] is attached to the "
            "posterior brainstem and coordinates voluntary movement and balance while also "
            "contributing to nonverbal learning and memory.",
            "Think of the sub-divisions as a **building**, the brainstem is the foundation and "
            "wiring that keeps the lights on, the diencephalon is the switchboard, and the "
            "cerebrum is the top floor where conscious thought happens.",
            "Clinically, knowing these sub-divisions lets the neuropsychologist **localise "
            "deficits**, because lower brainstem damage threatens survival reflexes while "
            "cerebral damage alters cognition, personality and voluntary behaviour instead.",
            "In summary, the human brain is studied as a developmental and functional "
            "**hierarchy** running from the life-sustaining brainstem, through the relaying "
            "diencephalon, up to the cerebrum and cerebellum that refine thought and movement.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t1_subdivisions.jpg",
                "alt": "Mid and lower brain showing cerebrum, diencephalon, midbrain, pons, "
                       "medulla and cerebellum labelled",
                "caption": "Figure 1. Major anatomical sub-divisions seen in one view: the "
                           "cerebrum above, the diencephalon (thalamus and hypothalamus) in the "
                           "centre, the brainstem (midbrain, pons and medulla) running down, and "
                           "the cerebellum behind, matching the embryo-to-adult hierarchy in the "
                           "flowchart.",
                "source": "OpenStax Anatomy and Physiology (figure \u201c1311 Brain Stem\u201d)",
                "author": "OpenStax",
                "license": "CC BY 4.0",
            },
        ],
        "definition": None,
        "flow": {
            "title": "Major sub-divisions from embryo to adult brain",
            "steps": [
                ("5 vesicles", "telencephalon to myelencephalon", False),
                ("3 primary parts", "fore / mid / hindbrain", False),
                ("Brainstem", "medulla, pons, midbrain", False),
                ("Diencephalon", "thalamus, hypothalamus, epithalamus", False),
                ("Cerebrum + Cerebellum", "higher cognition + balance", True),
            ],
        },
        "mnemonic": {
            "word": "BDCC",
            "line": "Four adult sub-divisions, bottom to top: Brainstem, Diencephalon, "
                    "Cerebrum, Cerebellum.",
            "items": [
                ("B", "**Brainstem** runs automatic survival functions (medulla, pons, "
                       "midbrain)."),
                ("D", "**Diencephalon** relays and regulates (thalamus, hypothalamus, "
                       "epithalamus)."),
                ("C", "**Cerebrum** handles higher thought, memory and voluntary movement."),
                ("C", "**Cerebellum** coordinates balance and fine, skilled movement."),
            ],
        },
    },
    {
        "anchor": "surface",
        "number": 2,
        "color": "teal",
        "title": "Surface Anatomy of Cortical Regions",
        "intro": "The surface of the cerebrum is a folded sheet of grey matter whose ridges, "
                 "grooves and landmark fissures carve it into four lobes, each carrying motor, "
                 "sensory and association areas with clear behavioural roles.",
        "points": [
            "The **cerebral cortex** [[the thin outer sheet of grey matter covering the "
            "cerebrum]] is only about 0.125 inches thick yet holds roughly **40 percent** of "
            "brain mass and enables sensation, memory, understanding and voluntary movement.",
            "The surface is thrown into folds called **gyri** [[the raised ridges of the brain "
            "surface]] and **sulci** [[the shallow grooves between ridges]], a design that "
            "greatly increases surface area without increasing the brain's overall volume.",
            "Deeper grooves are called **fissures**, and the long **interhemispheric fissure** "
            "[[the deep midline cleft]] splits the cerebrum into a left and a right hemisphere "
            "joined underneath by the corpus callosum.",
            "The **central sulcus** [[the groove separating the frontal and parietal lobes]] "
            "runs down and forward and is the key landmark dividing the motor cortex in front "
            "from the sensory cortex behind it.",
            "The **lateral fissure** [[the Sylvian groove on the side of the brain]] runs "
            "backward and upward and marks off the temporal lobe below it from the frontal and "
            "parietal lobes above.",
            "These landmarks segregate each hemisphere into **four lobes**, the frontal, the "
            "parietal, the temporal and the occipital, which is the framework every exam answer "
            "on surface anatomy should name first.",
            "The **frontal lobe** carries the **primary motor cortex** in the **precentral "
            "gyrus** [[the ridge just in front of the central sulcus]], which uses pyramidal "
            "cells to drive precise, skilled, conscious voluntary movement.",
            "Just ahead lie the **premotor cortex** for learned patterned movements and the "
            "**frontal eye field** [[the area directing voluntary eye movements]], which plan "
            "and sequence actions before the primary motor cortex executes them.",
            "**Broca's area** [[the motor-speech region in the frontal lobe]] lies near the "
            "lateral sulcus, usually on the left, and instructs the tongue and mouth muscles so "
            "a person can produce fluent spoken words.",
            "The **parietal lobe** holds the **primary somatosensory cortex** in the "
            "**postcentral gyrus** [[the ridge just behind the central sulcus]], which receives "
            "touch, pressure and position signals from the skin and skeletal muscles.",
            "The neighbouring **somatosensory association cortex** integrates these raw signals "
            "into a whole perception of an object's size, texture and the spatial relationship "
            "of its parts, turning sensation into meaningful recognition.",
            "The **occipital lobe** contains the **primary visual cortex** at its posterior tip, "
            "which receives signals from the retinas and interprets colour, form and movement to "
            "build conscious vision.",
            "The **temporal lobe** carries the **primary auditory cortex** on its superior "
            "margin for pitch, rhythm and loudness, and houses **Wernicke's area** [[the region "
            "for understanding words]] for comprehending language.",
            "The **prefrontal cortex** [[the front-most association area]] governs intellect, "
            "judgment, reasoning and personality, so its damage, as in the historic case of "
            "**Phineas Gage**, can profoundly change character while sparing basic movement.",
            "Clinically the **lobes map onto lesion syndromes**, Broca's damage causes "
            "non-fluent aphasia, Wernicke's damage causes fluent but meaningless speech, "
            "occipital damage causes cortical blindness and parietal damage causes contralateral "
            "neglect.",
            "Treat the cortical surface like a **city map**, the gyri are blocks, the sulci are "
            "streets, and the central sulcus and lateral fissure are the main avenues you use to "
            "find any functional neighbourhood quickly.",
            "In summary, reading the surface means naming the gyri and sulci, locating the "
            "central and lateral fissures, dividing the **four lobes**, and attaching each lobe's "
            "motor, sensory, association and clinical roles in order.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t2_lobes.png",
                "alt": "Lateral view of the left cerebral hemisphere with the frontal, parietal, "
                       "temporal and occipital lobes shaded in different colours",
                "caption": "Figure 2. Lateral surface of the cerebral hemisphere showing the four "
                           "lobes (frontal, parietal, temporal and occipital), the boundaries used "
                           "to tell them apart being the central sulcus and the lateral fissure "
                           "described in the points and flowchart.",
                "source": "Wikimedia Commons file \u201cLobes of the brain.png\u201d",
                "author": "Monnibear",
                "license": "CC0 1.0 (public-domain dedication)",
            },
        ],
        "definition": None,
        "flow": {
            "title": "Reading the cortical surface",
            "steps": [
                ("Gyri + sulci", "ridges and grooves", False),
                ("Fissures", "interhemispheric, central, lateral", False),
                ("Four lobes", "frontal, parietal, temporal, occipital", False),
                ("Functional areas", "motor, sensory, association", True),
            ],
        },
        "mnemonic": {
            "word": "FPTO",
            "line": "The four cortical lobes, front to back: Frontal, Parietal, Temporal, "
                    "Occipital.",
            "items": [
                ("F", "**Frontal** lobe: motor cortex, premotor, frontal eye field, Broca's "
                       "area, prefrontal personality."),
                ("P", "**Parietal** lobe: primary somatosensory cortex and somatosensory "
                       "association area."),
                ("T", "**Temporal** lobe: primary auditory cortex and Wernicke's comprehension "
                       "area."),
                ("O", "**Occipital** lobe: primary visual cortex interpreting colour, form and "
                       "movement."),
            ],
        },
    },
    {
        "anchor": "interior",
        "number": 3,
        "color": "purple",
        "title": "Interior Structures of Cortical and Sub-cortical Regions",
        "intro": "Beneath the folded cortex lie the deep grey-matter structures that relay "
                 "signals, drive emotion and memory, and smooth movement, forming the "
                 "sub-cortical core on which conscious cortical life depends.",
        "points": [
            "Below the cortical surface the brain holds **sub-cortical structures** [[the deep "
            "grey-matter centres hidden beneath the cortex]] that relay, regulate and modulate "
            "the information the cortex ultimately uses for conscious behaviour.",
            "The two hemispheres are joined internally by the **corpus callosum** [[the thick "
            "bridge of nerve fibres joining the two halves of the brain]], a broad white-matter "
            "bundle that transmits information rapidly from one side to the other.",
            "Deep in the forebrain the **diencephalon** surrounds the third ventricle and "
            "contains the thalamus, the hypothalamus and the epithalamus, the three paired "
            "structures that form the brain's inner control core.",
            "The **thalamus** [[the great relay station for the senses]] is a pair of egg-shaped "
            "masses whose nuclei funnel almost all afferent impulses to the cortex, so it shapes "
            "sensation, arousal, learning and memory.",
            "The **hypothalamus** [[the body's internal regulator below the thalamus]] controls "
            "hunger, thirst, temperature, heartbeat and sexual behaviour, and governs hormones "
            "through the pituitary gland sitting directly beneath it.",
            "The hypothalamus also contains a **pleasure centre** [[a reward region deep in the "
            "brain]], shown when **Olds and Milner (1954)** found rats pressing a lever up to "
            "7000 times an hour for electrical self-stimulation.",
            "The **epithalamus** carries the **pineal gland** [[a tiny endocrine gland "
            "secreting melatonin]], which helps set the sleep and wake rhythm that underlies "
            "mood, alertness and daily behaviour.",
            "The **limbic system** [[the emotional ring of structures around the brainstem]] is "
            "a doughnut-shaped set of centres, including the hippocampus, amygdala and "
            "hypothalamus, driving fear, aggression and the urges for food and sex.",
            "The **hippocampus** [[the seahorse-shaped memory centre]] is essential for forming "
            "new long-term memories, shown by the famous patient **H.M.**, who lost the ability "
            "to make new memories after its surgical removal.",
            "The **amygdala** [[two lima-bean clusters deep in the temporal lobe]] processes "
            "fear and aggression and attaches emotional meaning to memories, so it is central to "
            "anxiety and trauma-related clinical conditions.",
            "The **basal nuclei** [[deep grey masses that help control movement]], also called "
            "the basal ganglia, start and smooth voluntary movements and suppress unwanted ones, "
            "linking cortex to motor output.",
            "The **substantia nigra** [[a dark midbrain nucleus linked to the basal nuclei]] "
            "supplies the dopamine the basal nuclei need, and its degeneration is the direct "
            "cause of the tremor and rigidity of **Parkinson's disease**.",
            "The **reticular formation** [[a net-like web of neurons in the brainstem]] filters "
            "incoming stimuli and passes on only what matters, maintaining consciousness, "
            "arousal and the sleep and wake cycle.",
            "The **cerebral peduncles** [[the stalks carrying motor fibres through the "
            "midbrain]] relay the descending pyramidal motor tracts from the cortex downward, "
            "connecting conscious intention to bodily movement.",
            "Think of the interior as a **control room behind the screen**, the cortex is the "
            "display we watch, while the thalamus, limbic system and basal nuclei quietly route "
            "signals, add emotion and steady every action.",
            "Clinically these structures anchor major disorders, hippocampal damage causes "
            "amnesia, amygdala dysfunction underlies anxiety, and substantia nigra loss produces "
            "Parkinson's, so the neuropsychologist must map behaviour onto depth.",
            "In summary, the interior runs from the connecting **corpus callosum**, through the "
            "relaying thalamus and regulating hypothalamus, to the emotional limbic system and "
            "the movement-smoothing basal nuclei supporting the cortex.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t3_limbic.png",
                "alt": "Medial view of the brain highlighting the hippocampus, amygdala, "
                       "thalamus, hypothalamus and other limbic-system structures",
                "caption": "Figure 3. Interior and sub-cortical structures of the limbic system, "
                           "showing the hippocampus (memory) and amygdala (emotion) sitting deep "
                           "beside the thalamus and hypothalamus, the depth structures named in "
                           "the points and the inward-moving flowchart.",
                "source": "Wikimedia Commons file \u201cBlausen 0614 LimbicSystem.png\u201d",
                "author": "BruceBlaus (Blausen Medical)",
                "license": "CC BY 3.0",
            },
        ],
        "definition": None,
        "flow": {
            "title": "From cortex inward to sub-cortical structures",
            "steps": [
                ("Corpus callosum", "joins the two hemispheres", False),
                ("Thalamus", "relays signals to cortex", False),
                ("Hypothalamus", "regulates body and hormones", False),
                ("Limbic system", "hippocampus + amygdala: emotion, memory", False),
                ("Basal nuclei + substantia nigra", "smooth movement", True),
            ],
        },
        "mnemonic": {
            "word": "THALAMUS",
            "line": "Deep structures, outside inward: Thalamus, Hypothalamus, Amygdala, Limbic "
                    "ring, And basal nuclei, Memory hippocampus, Unseen reticular net, "
                    "Substantia nigra.",
            "items": [
                ("T", "**Thalamus** relays nearly all sensory and motor traffic to the cortex."),
                ("H", "**Hypothalamus** regulates hunger, thirst, temperature and hormones."),
                ("A", "**Amygdala** processes fear, aggression and emotional memory."),
                ("L", "**Limbic** ring binds emotion with drives for food and sex."),
                ("A", "**And** the corpus callosum bridges the two hemispheres."),
                ("M", "**Memory** depends on the hippocampus (patient H.M.)."),
                ("U", "**Unseen** reticular formation keeps us conscious and aroused."),
                ("S", "**Substantia** nigra feeds the basal nuclei; its loss causes "
                       "Parkinson's."),
            ],
        },
    },
    {
        "anchor": "connectivity",
        "number": 4,
        "color": "amber",
        "title": "Anatomical Connectivity Among the Regions",
        "intro": "No brain region works alone; bundles of nerve fibres link hemisphere to "
                 "hemisphere, cortex to spinal cord and relay to cortex, so connectivity is what "
                 "turns separate structures into coordinated behaviour.",
        "points": [
            "The brain is a **connected network**, because no functional area acts alone and "
            "conscious behaviour emerges only when white-matter fibre pathways link the "
            "separate cortical and sub-cortical regions together.",
            "White-matter fibres fall into three families, **commissural fibres** [[fibres "
            "joining the two hemispheres]], **projection fibres** [[fibres running up to or down "
            "from the cortex]] and association fibres linking areas within one hemisphere.",
            "The great commissure is the **corpus callosum** [[the thick bridge of nerve fibres "
            "joining the two halves of the brain]], which transmits information from one "
            "hemisphere to the other so the two sides act as one mind.",
            "When the corpus callosum is surgically cut to control severe epilepsy, the result "
            "is the **split-brain** [[a condition where the two hemispheres can no longer share "
            "information]], revealing that each side can process the world independently.",
            "Each hemisphere works **contralaterally** [[each side of the brain controls the "
            "opposite side of the body]], so the left hemisphere moves and senses the right side "
            "of the body and the right hemisphere the left side.",
            "This crossover is anatomically created at the **decussation of the pyramids** [[the "
            "point in the medulla where motor fibres cross to the opposite side]], where the "
            "descending corticospinal tracts switch sides.",
            "The **corticospinal tracts** [[the main motor highway from cortex to spinal cord]] "
            "begin in the precentral gyrus, descend through the cerebral peduncles and medullary "
            "pyramids, and carry commands for skilled voluntary movement.",
            "Signals flowing the other way, from body to brain, are **ascending sensory tracts** "
            "[[pathways carrying sensation upward]], which climb through the brainstem and "
            "synapse in the thalamus before reaching the cortex.",
            "The **thalamus** acts as the central **relay hub** [[the junction where incoming "
            "signals are switched to the right destination]], because nearly all inputs heading "
            "for the cortex are routed and synapsed through it.",
            "The **hypothalamus** is described as a switchboard that forwards messages onward to "
            "the sensory cortex, the cerebellum and the medulla, linking internal state to "
            "widespread brain regions.",
            "The **reticular formation** forms a diffuse network threading through the brainstem "
            "that connects with ten of the twelve cranial nerves and filters stimuli on their "
            "way to higher centres.",
            "The ventricular channels also link regions, as the **cerebral aqueduct** [[the "
            "narrow tube between the third and fourth ventricles]] connects the fluid cavities "
            "that run through the brain's core.",
            "The **cerebral peduncles** carry the descending pyramidal motor tracts through the "
            "midbrain, acting as the structural cables that connect cortical intention to the "
            "spinal machinery of movement.",
            "Think of connectivity as the brain's **road and cable system**, the corpus callosum "
            "is the bridge between two cities, the corticospinal tracts are the motorways, and "
            "the thalamus is the central switching station.",
            "Clinically, cutting or damaging these links is revealing, split-brain surgery "
            "separates the hemispheres, while a stroke on one side of the motor pathway causes "
            "weakness on the **opposite** side of the body because of decussation.",
            "In summary, connectivity knits the brain together through interhemispheric "
            "commissures, crossed corticospinal tracts, ascending sensory pathways and the "
            "thalamic relay, converting isolated regions into one coordinated system.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t4_corpus_callosum.png",
                "alt": "Midsagittal drawing of the brain showing the corpus callosum arching "
                       "between the two hemispheres",
                "caption": "Figure 4. Midsagittal view showing the corpus callosum, the great "
                           "arched band of fibres that joins the two hemispheres, the main "
                           "interhemispheric bridge named first in the connectivity points and "
                           "flowchart.",
                "source": "Wikimedia Commons file \u201cGray764.png\u201d (Gray\u2019s Anatomy, "
                          "1918 plate)",
                "author": "Henry Vandyke Carter",
                "license": "Public domain",
            },
        ],
        "definition": None,
        "flow": {
            "title": "How regions are wired together",
            "steps": [
                ("Interhemispheric", "corpus callosum joins the sides", False),
                ("Descending motor", "corticospinal tracts to body", False),
                ("Decussation", "fibres cross in the medulla", False),
                ("Ascending sensory", "body signals climb up", False),
                ("Thalamic relay", "switched on to the cortex", True),
            ],
        },
        "mnemonic": {
            "word": "CCDR",
            "line": "Connectivity in order: Callosum, Contralateral crossover, Descending and "
                    "ascending tracts, Relay through thalamus.",
            "items": [
                ("C", "**Corpus callosum** bridges the two hemispheres (cut in split-brain "
                       "surgery)."),
                ("C", "**Contralateral** control: each hemisphere runs the opposite side of the "
                       "body."),
                ("D", "**Decussation** of the pyramids crosses the corticospinal tracts in the "
                       "medulla."),
                ("R", "**Relay** hub: the thalamus switches ascending signals on to the "
                       "cortex."),
            ],
        },
    },
    {
        "anchor": "blood",
        "number": 5,
        "color": "rose",
        "title": "Blood Supply to the Brain",
        "intro": "The brain is a hungry organ that cannot store fuel, so a two-source arterial "
                 "supply, a protective ring of vessels and a selective barrier keep it "
                 "continuously nourished and shielded from harm.",
        "points": [
            "The **brain** is metabolically greedy, using about **25 percent** of the body's "
            "glucose and roughly **75 percent** of its oxygen, so a constant blood supply is "
            "essential to normal brain functioning.",
            "Because the brain cannot store fuel, any **reduction in blood supply** quickly "
            "starves neurons and can cause a **stroke** [[sudden brain damage from loss of "
            "blood flow]], making vascular anatomy clinically central.",
            "Blood reaches the brain through **two arterial sources**, the paired **internal "
            "carotid arteries** [[large vessels at the front of the neck]] in front and the "
            "paired **vertebral arteries** [[vessels running up through the spine]] behind.",
            "The two vertebral arteries join to form the single **basilar artery** [[the main "
            "artery on the underside of the brainstem]], which supplies the brainstem and "
            "cerebellum before feeding the back of the brain.",
            "At the base of the brain these front and back supplies join in the **Circle of "
            "Willis** [[a ring of arteries that links the main vessels together]], described in "
            "standard neuroanatomy texts though not detailed in the primary chapter.",
            "The **Circle of Willis** is protective because it provides **collateral "
            "circulation** [[backup routes for blood if one vessel is blocked]], allowing flow "
            "to reroute and partly preserve supply when a single artery narrows.",
            "From this ring arise three paired territorial vessels, the anterior, middle and "
            "posterior cerebral arteries, each supplying a defined region of the cortex "
            "according to standard neuroanatomy references.",
            "The **anterior cerebral artery** [[the vessel feeding the front and midline "
            "cortex]] supplies the medial frontal and parietal areas, so its blockage typically "
            "weakens the opposite leg more than the arm.",
            "The **middle cerebral artery** [[the largest branch, feeding the side of the "
            "brain]] supplies much of the lateral cortex including the language areas, so its "
            "blockage often causes face and arm weakness with aphasia.",
            "The **posterior cerebral artery** [[the vessel feeding the back of the brain]] "
            "supplies the occipital lobe, so its blockage can cause loss of vision in part of "
            "the visual field on the opposite side.",
            "Protecting this rich supply is the **blood-brain barrier** [[a filter that keeps "
            "harmful substances in the blood out of the brain]], which allows nutrients through "
            "but shields neurons from toxins by differential permeability.",
            "The barrier is **selective**, letting nutrients pass freely but being ineffective "
            "against substances that can diffuse straight through plasma membranes, while "
            "**stress** raises its permeability and lets harmful chemicals slip through.",
            "The **choroid plexuses** [[capillary clusters that filter blood into fluid]] carry "
            "ion pumps that adjust the composition of the central nervous system and help clear "
            "wastes, linking the blood supply to brain housekeeping.",
            "Picture the arterial supply as a **city water system**, two mains enter from front "
            "and back, the Circle of Willis is a ring main that reroutes flow if a pipe bursts, "
            "and the barrier is the filtration plant.",
            "Clinically, each artery defines a **vascular territory**, so a neuropsychologist can "
            "predict the pattern of deficits after a stroke from which cerebral artery was "
            "blocked and which cortical region it fed.",
            "In summary, the brain is fed by paired carotid and vertebral arteries that unite in "
            "the Circle of Willis, branch into territorial cerebral arteries, and are guarded "
            "throughout by the protective blood-brain barrier.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t5_circle_of_willis.png",
                "alt": "Diagram of the arteries at the base of the brain forming the Circle of "
                       "Willis, with the internal carotid, anterior, middle and posterior "
                       "cerebral arteries labelled",
                "caption": "Figure 5. The Circle of Willis at the base of the brain, where the "
                           "paired internal carotid and vertebral-basilar supplies join into a "
                           "ring that feeds the anterior, middle and posterior cerebral arteries "
                           "described in the points and flowchart.",
                "source": "Wikimedia Commons file \u201cCircle of Willis en.svg\u201d",
                "author": "Rhcastilhos (after Gray\u2019s Anatomy)",
                "license": "Public domain",
            },
        ],
        "definition": {
            "label": "Blood-Brain Barrier (characterisation from the source text)",
            "text": "The Blood-Brain Barrier is a protective mechanism that helps in "
                    "maintenance of a stable environment for the brain; harmful substances are "
                    "shielded from the brain by the barrier, which avoids exposure of "
                    "blood-borne substances and other toxic chemicals to the brain by "
                    "differential permeability.",
            "cite": "Examining Biological Foundations of Human Behavior (ed. Barre Vijaya "
                    "Prasad, 2020), chapter \u201cAnatomy of the Brain\u201d.",
        },
        "flow": {
            "title": "Blood reaching and protecting the brain",
            "steps": [
                ("Carotid + vertebral arteries", "two entry sources", False),
                ("Basilar artery", "from fused vertebrals", False),
                ("Circle of Willis", "linking ring at the base", False),
                ("Cerebral arteries", "anterior, middle, posterior", False),
                ("Blood-brain barrier", "selective protection", True),
            ],
        },
        "mnemonic": {
            "word": "AMP",
            "line": "The three territorial cerebral arteries from the Circle of Willis: "
                    "Anterior, Middle, Posterior.",
            "items": [
                ("A", "**Anterior** cerebral artery: medial frontal and parietal cortex "
                       "(opposite leg)."),
                ("M", "**Middle** cerebral artery: lateral cortex and language (face, arm, "
                       "aphasia)."),
                ("P", "**Posterior** cerebral artery: occipital lobe and vision."),
            ],
        },
    },
    {
        "anchor": "csf",
        "number": 6,
        "color": "navy",
        "title": "The Cerebrospinal Fluid (CSF) System",
        "intro": "A thin watery fluid cushions the brain inside its coverings, flows through a "
                 "chain of inner cavities and is finally reabsorbed into the blood, so the CSF "
                 "system both protects and nourishes the central nervous system.",
        "points": [
            "**Cerebrospinal fluid** [[the clear watery fluid that bathes the brain and spinal "
            "cord]] is an 80 to 150 ml solution similar to blood plasma but with less protein "
            "and different ion concentrations.",
            "Its main role is mechanical, because the fluid forms a liquid cushion that gives "
            "**buoyancy** [[the floating effect that reduces effective weight]] so the brain "
            "does not crush under its own weight.",
            "The fluid also protects the central nervous system from blows and trauma, provides "
            "**nourishment** to the brain and carries chemical signals throughout it, so it is "
            "far more than a simple shock absorber.",
            "The brain sits inside three protective membranes called the **meninges** [[the "
            "three coverings wrapping the brain and spinal cord]], named the dura mater, the "
            "arachnoid mater and the pia mater from outside inward.",
            "The **dura mater** [[the tough leathery outer covering]] is a strong two-layered "
            "membrane whose layers separate in places to form the dural sinuses that drain "
            "venous blood from the brain.",
            "The **arachnoid mater** [[the delicate web-like middle covering]] lies beneath the "
            "dura, separated from it by the subdural space, and spans the surface loosely like a "
            "spider's web.",
            "The **subarachnoid space** [[the fluid-filled gap beneath the arachnoid]] lies "
            "between the arachnoid and the pia mater and is filled with cerebrospinal fluid and "
            "large blood vessels that bathe the brain's surface.",
            "The **pia mater** [[the thin inner covering clinging to the brain]] is a delicate "
            "membrane that follows every gyrus and sulcus, hugging the brain surface tightly and "
            "carrying small vessels into the tissue.",
            "The fluid is produced by the **choroid plexuses** [[clusters of capillaries that "
            "filter blood into fluid]], which hang from the roof of each ventricle and use ion "
            "pumps to set the fluid's composition.",
            "The fluid fills the **ventricles** [[a connected web of cavities inside the "
            "brain]], namely the two lateral ventricles, the third ventricle between the "
            "thalamic halves and the fourth ventricle between brainstem and cerebellum.",
            "Fluid flows from the lateral ventricles into the third ventricle through the "
            "**foramen of Monro** [[the channel between the lateral and third ventricles]], "
            "beginning its one-way journey through the brain's core.",
            "From the third ventricle the fluid passes through the **cerebral aqueduct** [[the "
            "narrow tube between the third and fourth ventricles]] into the fourth ventricle at "
            "the back of the brainstem.",
            "It then escapes through the central and lateral **apertures** [[openings letting "
            "fluid leave the fourth ventricle]] into the subarachnoid space, where it circulates "
            "around the brain and spinal cord.",
            "Finally the fluid is reabsorbed into the venous blood through the **arachnoid "
            "villi** [[tiny projections that drain fluid back into the blood]], completing the "
            "cycle of production, circulation and absorption.",
            "Picture the CSF system as a **water cushion in a sealed tank**, fluid is pumped in "
            "at the choroid plexus, flows through linked chambers, bathes the brain and drains "
            "out, keeping pressure and buoyancy balanced.",
            "Clinically, if the flow is blocked anywhere along this chain the fluid backs up and "
            "raises pressure, producing **hydrocephalus** [[a dangerous build-up of fluid and "
            "pressure inside the brain]], a point confirmed in standard neuroanatomy texts.",
            "In summary, the CSF system produces fluid at the choroid plexus, circulates it "
            "through the ventricles and subarachnoid space to cushion and feed the brain, and "
            "reabsorbs it at the arachnoid villi.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t6_csf.jpg",
                "alt": "Diagram of cerebrospinal fluid circulation through the lateral, third "
                       "and fourth ventricles and the subarachnoid space",
                "caption": "Figure 6. Circulation of cerebrospinal fluid from the choroid "
                           "plexuses through the lateral, third and fourth ventricles, out into "
                           "the subarachnoid space and back to the venous blood, the exact "
                           "production-to-absorption path traced in the points and flowchart.",
                "source": "OpenStax Anatomy and Physiology (figure \u201c1317 CSF "
                          "Circulation\u201d)",
                "author": "OpenStax",
                "license": "CC BY 4.0",
            },
        ],
        "definition": None,
        "flow": {
            "title": "CSF production, flow and absorption",
            "steps": [
                ("Choroid plexus", "produces CSF", False),
                ("Lateral ventricles", "via foramen of Monro", False),
                ("Third ventricle", "via cerebral aqueduct", False),
                ("Fourth ventricle", "out through apertures", False),
                ("Arachnoid villi", "reabsorbed into blood", True),
            ],
        },
        "mnemonic": {
            "word": "CLEAR",
            "line": "The CSF pathway in order: Choroid plexus, Lateral ventricles, Each inner "
                    "aqueduct, Apertures out, Reabsorbed at villi.",
            "items": [
                ("C", "**Choroid** plexus produces the cerebrospinal fluid."),
                ("L", "**Lateral** ventricles drain through the foramen of Monro to the third "
                       "ventricle."),
                ("E", "**Each** flow then passes the cerebral aqueduct into the fourth "
                       "ventricle."),
                ("A", "**Apertures** release fluid into the subarachnoid space around the "
                       "brain."),
                ("R", "**Reabsorbed** into venous blood at the arachnoid villi (blockage causes "
                       "hydrocephalus)."),
            ],
        },
    },
    {
        "anchor": "cytoarchitecture",
        "number": 7,
        "color": "blue",
        "title": "Cytoarchitecture and Modular Organization of the Brain",
        "intro": "Studied under the microscope, the cortex is not uniform but built of layers "
                 "and cell patterns that differ region by region, and this cellular architecture "
                 "underlies the modular, map-like way the brain localises its functions.",
        "points": [
            "**Cytoarchitecture** [[the study of how brain cells are arranged under the "
            "microscope]] describes the types, shapes, sizes and layering of neurons in each "
            "region, and it is the foundation on which functional brain maps are built.",
            "To see this architecture, tissue must be **stained**, because untreated nerve "
            "tissue is almost transparent, so staining makes neuronal placement and fine "
            "features visible at the microscopic level.",
            "**Nissl staining** [[a dye method that stains the protein-making machinery of "
            "neurons]] uses basic aniline dyes to mark the acidic polyribosomes in the rough "
            "endoplasmic reticulum, producing dark patches called Nissl bodies.",
            "Nissl staining lets investigators differentiate **cell types** such as neurons and "
            "glia and compare their shapes, sizes and cytoarchitecture, which is exactly how "
            "cortical layers and regions are distinguished.",
            "**Golgi's method** [[a silver-staining technique first described by Camillo Golgi "
            "in 1873]] fills whole neurons and glia with a black-brown silver chromate "
            "precipitate, revealing cell bodies and their branching dendrites.",
            "Because Golgi's method stains only a few cells completely, it shows the full shape "
            "of individual neurons including their **dendrites** [[the branching fibres that "
            "receive signals]], complementing the population view that Nissl staining provides.",
            "**Histochemistry** [[using chemical reactions to reveal brain structure]] adds "
            "another layer of detail by using biochemical reactions to visualise specific "
            "substances in the tissue under the microscope.",
            "At the organ level the brain is **modular**, meaning its regions are organised into "
            "units that each serve a specific role within the defined pathways of the nervous "
            "system rather than acting at random.",
            "This modular, localised organisation is captured by the **Brodmann areas** [[a map "
            "numbering cortical regions by their cell architecture]], which the standard "
            "reference Brodmann (1909) divided into about **52 areas**, as detailed in "
            "neuroanatomy texts.",
            "Each Brodmann area is defined by its distinctive cytoarchitecture and tends to "
            "carry a distinct function, so for example the primary motor and primary visual "
            "cortices correspond to specific numbered areas on the map.",
            "This supports the principle of **functional localisation** [[the idea that "
            "particular tasks are handled by particular brain regions]], first argued by early "
            "workers like Gall and Broca and refined by modern cytoarchitectonic mapping.",
            "Yet localisation is not absolute, because the source stresses that **no functional "
            "area acts alone** and conscious behaviour is the outcome of the whole cortex working "
            "together through its connecting pathways.",
            "Modern **non-invasive imaging** [[scanning methods that view the living brain "
            "without surgery]] such as CT, MRI, PET, fMRI, DTI and SPECT now map structure and "
            "function in living people using radioactive and non-radioactive tracers.",
            "Techniques such as **diffusion tensor imaging** [[a scan that traces white-matter "
            "fibre bundles]] and connectomics extend cytoarchitecture outward to the brain's "
            "wiring, linking cellular modules into large-scale networks.",
            "Think of the cortex as a **patchwork quilt**, each Brodmann patch is stitched from a "
            "distinctive cellular pattern and does a particular job, yet the patches are sewn "
            "together so the whole quilt functions as one.",
            "Clinically, this cytoarchitectonic map lets the neuropsychologist link a specific "
            "deficit to a specific **cortical area**, grounding the localisation of language, "
            "motor and sensory functions that underpins modern assessment.",
            "In summary, staining reveals the brain's cytoarchitecture, that architecture defines "
            "modular regions such as the Brodmann areas, and imaging now maps this localisation "
            "and connectivity in the living brain.",
        ],
        "figures": [
            {
                "file": "assets/unit1_figures/t7_brodmann.png",
                "alt": "Lateral view of the cerebral cortex with Brodmann areas shown as numbered "
                       "coloured patches",
                "caption": "Figure 7. Brodmann's cytoarchitectonic map, dividing the cortex into "
                           "numbered areas that each share a distinctive cell architecture and "
                           "function, the modular organisation that staining reveals in the "
                           "points and flowchart.",
                "source": "Wikimedia Commons file \u201cBrodmann areas 3D.png\u201d",
                "author": "Mark Dow (Brain Development Lab, University of Oregon)",
                "license": "Public domain",
            },
        ],
        "definition": None,
        "flow": {
            "title": "From staining to cortical maps",
            "steps": [
                ("Staining", "Nissl, Golgi, histochemistry", False),
                ("Cytoarchitecture", "cell type, shape, layering", False),
                ("Modular regions", "each area, a role", False),
                ("Brodmann map", "about 52 numbered areas", False),
                ("Imaging + connectomics", "living brain networks", True),
            ],
        },
        "mnemonic": {
            "word": "SCRIB",
            "line": "Cytoarchitecture to maps: Stain, Cell architecture, Regional modules, "
                    "Imaging, Brodmann areas.",
            "items": [
                ("S", "**Stain** the tissue with Nissl, Golgi or histochemistry to see cells."),
                ("C", "**Cell** architecture (type, shape, size, layering) defines each "
                       "region."),
                ("R", "**Regional** modules each serve a specific role in defined pathways."),
                ("I", "**Imaging** (CT, MRI, PET, fMRI, DTI, SPECT) maps living structure and "
                       "function."),
                ("B", "**Brodmann** areas (about 52) number the cortex and localise function."),
            ],
        },
    },
]

REFERENCES = {
    "books": [
        "Examining Biological Foundations of Human Behavior, edited by Barre Vijaya Prasad "
        "(2020), IGI Global \u2014 chapter \u201cAnatomy of the Brain\u201d by Sumit Kumar "
        "(primary source for the general organisation, sub-divisions, surface anatomy, interior "
        "and sub-cortical structures, meninges, ventricles, cerebrospinal fluid, blood-brain "
        "barrier and staining methods).",
    ],
    "online": [
        "Wikipedia, \u201cCircle of Willis\u201d (en.wikipedia.org) \u2014 consulted to confirm "
        "the arterial ring at the base of the brain and the internal carotid, vertebral-basilar, "
        "anterior, middle and posterior cerebral arteries (Topic 5).",
        "Wikipedia, \u201cCerebrospinal fluid\u201d (en.wikipedia.org) \u2014 consulted to "
        "confirm the production of fluid at the choroid plexus and its flow through the lateral, "
        "third and fourth ventricles into the subarachnoid space (Topic 6).",
        "OpenStax, Anatomy and Physiology (openstax.org) \u2014 the open CC BY textbook that is "
        "the source of the labelled brainstem and cerebrospinal-fluid circulation figures used "
        "in Topics 1 and 6.",
    ],
    "flagged": [
        "Clinical Neuroanatomy, by Richard S. Snell \u2014 a standard neuroanatomy text named "
        "here from general knowledge for the arterial territories (Topic 5) and the obstruction "
        "of cerebrospinal-fluid flow causing hydrocephalus (Topic 6); it was NOT opened directly "
        "while preparing these notes, so it is listed as a pointer rather than a consulted "
        "source.",
        "Localisation in the Cerebral Cortex (Vergleichende Lokalisationslehre der "
        "Grosshirnrinde), by Korbinian Brodmann (1909) \u2014 the original cytoarchitectonic map "
        "(Topic 7) named here from general knowledge; the primary 1909 volume was NOT accessed "
        "directly, so it is flagged rather than cited as consulted.",
        "StatPearls (NCBI Bookshelf) articles \u201cNeuroanatomy, Cerebral Hemisphere\u201d and "
        "\u201cNeuroanatomy, Mesencephalon / Midbrain\u201d \u2014 named inside the source "
        "chapter; they were NOT independently accessed while preparing these notes and so are "
        "flagged, not claimed as directly consulted.",
        "Olds, J., & Milner, P. (1954), the reward / pleasure-centre self-stimulation study in "
        "rats \u2014 cited within the source chapter and reported here secondhand; the original "
        "paper was not accessed directly.",
        "Journal articles cited inside the source chapter (for example Barthas & Kwan 2017, "
        "Binder 2017, Friederici 2018, Gittins & Harrison 2004 and Golgi 1873) are referenced "
        "only as they appear in the chapter and were not accessed directly.",
    ],
}


def figure_credits():
    """Build the figure-image source list from the figures declared in CONTENT.

    Each entry names the figure, its title, author and license honestly; kept
    in sync with the per-figure caption credits. Works hosted on Wikimedia
    Commons carry a "Via Wikimedia Commons" tag, while a figure whose source
    already states its primary origin (and any access route) is left as named.
    """
    credits = []
    for topic in CONTENT:
        for fig in topic.get("figures", []):
            caption = fig["caption"]
            fig_label = caption.split(".", 1)[0].strip() if "." in caption else caption
            suffix = "" if fig["source"].startswith("NIH BioArt") else " Via Wikimedia Commons."
            credits.append(
                "%s \u2014 %s by %s. License: %s.%s"
                % (fig_label, fig["source"], fig["author"], fig["license"], suffix))
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
    # Tokenise on the two markers while escaping everything else.
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
    """Plain text version (used by a companion generator if it imports this)."""
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
  font-size: 17px; color: rgba(255,255,255,0.85); max-width: 640px; margin-bottom: 38px;
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
.intro { font-style: italic; color: #40506a; margin: 10px 0 20px; font-size: 16.5px; }
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
.fig img {
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


def render_figures(topic):
    """Render each figure as an inline base64 <figure> (no external resources)."""
    figs = topic.get("figures")
    if not figs:
        return ""
    parts = []
    for fig in figs:
        credit = "Source: %s by %s. License: %s." % (fig["source"], fig["author"], fig["license"])
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
    parts.append('<ol class="points">')
    for p in t["points"]:
        parts.append("<li>" + render_inline_html(p) + "</li>")
    parts.append("</ol>")
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
    <h3>Figure image sources (online)</h3>
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
