# -*- coding: utf-8 -*-
"""
Generator for the self-contained offline HTML study-notes file:
"Biological Foundations of Behavior - Unit II: Structure and Functions of Cells".

This module defines ONE CONTENT data structure (a list of 7 topic dicts) that is
the single source of truth for both the HTML (rendered here) and the Word (.docx)
deliverable produced by the companion generator. Keep CONTENT clean and importable.

Markup conventions used inside point / intro / flow text:
  **word**                 -> a bold memory-hook keyword (a few per sentence only)
  [[plain-language text]]  -> a distinct bracketed plain-language explanation of a
                              difficult term, shown in its own visible style.

Run (from /projects/sandbox/Dango-kiro):
  PYTHONPATH=/projects/sandbox/.pylibs python3 generate_unit2_notes.py
"""

import html as _html
import os
import re

TITLE = "Biological Foundations of Behavior \u2014 Unit II: Structure and Functions of Cells"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"
SUBTITLE = ("Paper I: Biological Foundations of Behavior \u00b7 Part A (Anatomy, Physiology "
            "and Biochemistry of the CNS) \u00b7 Exam-ready study notes")
OUTPUT_HTML = "Biological_Foundations_Unit2_Structure_and_Functions_of_Cells.html"

# ---------------------------------------------------------------------------
# HOW TO READ THESE NOTES (short usage guide; NOT about how the notes were made)
# ---------------------------------------------------------------------------
HOW_TO_READ = [
    "Each topic opens with a one-line **intro** that frames the idea, followed by "
    "at least fifteen numbered points; every point is a full exam sentence, so read it "
    "as a ready-made **5-mark answer** rather than a short note.",
    "Hard anatomical or physiological terms are kept in full, but each is followed by a "
    "plain-language meaning inside a shaded box like this [[a simple everyday explanation "
    "that sits right next to the difficult word]] so you never lose the technical vocabulary.",
    "Only a **few keywords** in each sentence are shown in bold; treat those bold words as "
    "memory hooks and recall the rest of the sentence around them while you revise.",
    "Every topic carries exactly one **flowchart** of boxes joined by arrows to show the order "
    "of a process, and exactly one **mnemonic** whose letters each map to a point you must remember.",
    "Revise by first reading the flowchart to fix the sequence, then the mnemonic to lock the "
    "key points, and finally the full sentences to add exam-ready detail and wording.",
]

# ---------------------------------------------------------------------------
# CONTENT = single source of truth (7 topics, in the required order)
# Each topic dict:
#   anchor, number, color, title, intro, points[list], definition(optional dict),
#   flow(dict: title + steps[list]), mnemonic(dict: word, line, items[list of (letter, text)])
# ---------------------------------------------------------------------------
CONTENT = [
    {
        "anchor": "neurons",
        "number": 1,
        "color": "blue",
        "title": "Cells of the Nervous System \u2014 Neurons",
        "intro": "Neurons are the signalling cells that receive, process and pass on information across "
                 "the whole nervous system, and because every thought, feeling and behaviour ultimately "
                 "rests on how these cells fire and connect, they are the natural starting point for a "
                 "clinical psychologist studying the biology of behaviour.",
        "points": [
            "**Neurons** are the elementary information-processing cells of the nervous system, "
            "electrically **excitable** [[able to produce tiny electrical signals]] cells that receive "
            "and rapidly transmit chemical or electrical messages across the body.",
            "A typical neuron has three core parts, the **cell body** (soma), the branching **dendrites** "
            "and a single long **axon**, and this layout lets one cell gather many inputs yet send one "
            "focused output signal.",
            "The **soma** [[the main body of the cell that holds the nucleus]] contains the nucleus, "
            "endoplasmic reticulum, Golgi apparatus, lysosomes and mitochondria, so most of the cell's "
            "housekeeping and manufacturing work happens here.",
            "The **nucleus** stores genetic material inside chromosomes and transcribes **DNA into RNA** "
            "[[copying the gene instructions into a message]], which then leaves through nuclear pores to "
            "guide protein synthesis in the cytoplasm.",
            "**Dendrites** are fine branching filaments that carry information toward the cell body, and "
            "their many **dendritic spines** [[tiny bumps on the branches]] greatly increase the surface "
            "area available for synaptic connections.",
            "The **axon** behaves like the insulated wire of an electrical cable, starting at the "
            "**axon hillock** [[the cone-shaped base where the axon leaves the soma, which acts as the "
            "cell's decision point for whether to fire]] and carrying the signal away toward another "
            "neuron, a muscle or a gland.",
            "A fatty **myelin sheath** [[an insulating fatty wrapping around the axon, much like the "
            "plastic coating on an electrical wire]] encloses many axons in segments, and this insulation "
            "sharply increases the speed at which nerve impulses travel along the fibre.",
            "Gaps in the myelin called **nodes of Ranvier** [[bare spaces between the insulating segments]] "
            "let the signal be recharged as it jumps along the axon, keeping the impulse strong over long "
            "distances.",
            "Inside the soma, **Nissl granules** [[clusters of rough endoplasmic reticulum]] and free "
            "ribosomes build proteins, while **mitochondria** act as the powerhouse supplying the energy "
            "the busy neuron constantly needs.",
            "The **Golgi apparatus** [[the cell's packaging and dispatch centre]] processes peptides, "
            "proteins and neurotransmitters into vesicles, and microtubules then transport these packages "
            "to where the neuron needs them.",
            "By **function**, neurons are grouped as sensory (afferent) neurons that carry input to the "
            "CNS, motor (efferent) neurons that drive muscles and glands, and interneurons that connect "
            "the two inside the central nervous system.",
            "**Sensory (afferent) neurons** convert a stimulus from the outside world into an internal "
            "electrical signal and then carry that signal inward toward the spinal cord or brain for "
            "processing.",
            "**Motor (efferent) neurons** sit in the CNS and control muscles through the **neuromuscular "
            "junction** [[the contact point where a nerve meets a muscle]], and they are multipolar with "
            "one axon and many dendrites.",
            "**Interneurons** act as a local broker between sensory and motor neurons, handling signals "
            "nearby so that quick reflex-like responses do not always have to travel all the way back to "
            "the spinal cord.",
            "By **shape and number of processes**, neurons are classed as **unipolar**, **bipolar** or "
            "**multipolar** [[one, two, or many extensions leaving the cell body]], a structural scheme "
            "that reflects how each cell is wired into its circuit.",
            "Neurons can also be named by the **neurotransmitter** they release, such as cholinergic or "
            "glutamatergic cells, which links a cell's chemical identity directly to the kind of message "
            "it delivers to its partners.",
            "Neurons are enclosed by a **plasma membrane** [[the thin oily outer skin of the cell that "
            "controls what passes in and out]] and, because they rarely divide once mature, damage to "
            "them tends to be long-lasting, which is why neuronal loss in injury or dementia produces "
            "such enduring changes in behaviour.",
            "Taken together, a neuron is best understood as an **input-integrate-output device**: "
            "dendrites gather signals, the soma and axon hillock weigh them up, and the axon delivers a "
            "single decision onward, so disruption at any one stage can reshape how a person senses, "
            "moves, thinks or feels.",
        ],
        "definition": None,
        "flow": {
            "title": "Direction of information flow through a single neuron",
            "steps": [
                ("Dendrites", "receive input", False),
                ("Cell body (soma)", "integrates", False),
                ("Axon hillock", "trigger zone", False),
                ("Axon + myelin", "conducts", False),
                ("Axon terminal", "output / synapse", True),
            ],
        },
        "mnemonic": {
            "word": "DRAMA",
            "line": "Use DRAMA to recall the neuron's parts in signal order.",
            "items": [
                ("D", "**Dendrites** receive incoming signals."),
                ("R", "**Receiving** done, the cell body integrates them."),
                ("A", "**Axon hillock** decides whether to fire."),
                ("M", "**Myelin** on the axon speeds the impulse."),
                ("A", "**Axon terminal** passes the message on at the synapse."),
            ],
        },
    },
    {
        "anchor": "glia",
        "number": 2,
        "color": "teal",
        "title": "Supporting Cells (Glia)",
        "intro": "Glial cells are the support cells that feed, insulate, defend and hold the nervous "
                 "system together, and far from being mere packing, they shape how well neurons signal, "
                 "which is why glial dysfunction now sits at the centre of our understanding of "
                 "demyelinating and neurodegenerative disorders a clinical psychologist will meet.",
        "points": [
            "**Glial (neuroglial) cells** are the support cells that hold the nervous system together, "
            "providing physical integrity, nutrition and a stable **homeostasis** [[a steady internal "
            "environment]] around the neurons.",
            "Glia outnumber neurons by roughly a **three-to-one ratio**, which underlines that healthy "
            "brain function depends not only on neurons but heavily on the supporting cells that surround "
            "and service them.",
            "Beyond simple support, glia help regulate the chemical environment around neurons and take "
            "part in **cortical development**, axon and dendrite growth, making them active partners in "
            "building the nervous system.",
            "**Astrocytes** are star-shaped cells of the brain and spinal cord that keep the chemical "
            "environment stable by clearing surplus ions such as potassium and chloride from around the "
            "nerve cells.",
            "Astrocytes also **recycle neurotransmitters** [[reclaim and reuse chemical messengers]] that "
            "were released during synaptic transmission, so signalling chemicals are not wasted and the "
            "synapse is reset for the next message.",
            "Astrocyte **end-feet** [[swollen tips of the astrocyte]] wrap around and support the "
            "endothelial cells of the blood-brain barrier, helping shield the brain from harmful "
            "blood-borne substances while normal metabolism continues.",
            "The brain holds two kinds of astrocyte, **protoplasmic astrocytes** in grey matter and "
            "**fibrous astrocytes** in white matter, and together they are the most numerous glial cells "
            "of the central nervous system.",
            "**Oligodendrocytes** in the CNS form the **myelin sheath** [[fatty insulation]] around axons, "
            "and because one oligodendrocyte can coat up to about fifty axons they are highly efficient "
            "insulating cells.",
            "Myelin from oligodendrocytes lets the action potential jump between **nodes of Ranvier** "
            "[[gaps in the insulation]], so myelinated axons conduct far faster than unmyelinated ones, "
            "and when this insulation is stripped away, as in **demyelination** [[loss of the myelin "
            "coating, as seen in multiple sclerosis]], reasoning, speed of processing and motor "
            "performance all decline.",
            "**Microglia** are the resident immune cells of the CNS and act as specialised **macrophages** "
            "[[cells that engulf and digest debris]], guarding and supporting neurons throughout the brain "
            "and spinal cord.",
            "Microglia carry out **phagocytosis** [[swallowing and clearing waste]], make up roughly "
            "fifteen to seventeen percent of CNS cells, and multiply in response to brain injury to contain "
            "the damage.",
            "Microglia drive the **inflammatory response** [[the brain's reaction of swelling and "
            "immune activity to injury]] to brain trauma and are implicated in neurodegenerative "
            "disorders such as multiple sclerosis, motor neurone disease and Alzheimer's disease, and "
            "are key targets of HIV-1 in the CNS, which is why chronic brain inflammation is now linked "
            "to cognitive and mood changes.",
            "**Radial glial cells** are a temporary population in the developing CNS that act both as "
            "precursor cells and as a **scaffold** [[a guiding framework]] along which young neurons migrate "
            "to their correct positions.",
            "**Schwann cells** are the PNS counterpart of oligodendrocytes, and each Schwann cell insulates "
            "one single stretch of a peripheral axon with myelin to speed conduction outside the brain and "
            "spinal cord.",
            "**Non-myelinating Schwann cells** are important for the survival and maintenance of peripheral "
            "axons, showing that even glia which do not insulate still keep neurons alive and functioning "
            "healthily.",
            "A newly recognised glial type, the **oligodendrocyte precursor cell**, takes part in "
            "synaptic signalling in the **hippocampus** [[a brain region central to learning and "
            "memory]], revising the older idea that glia only ever served passive supporting roles.",
            "For the clinician the overall lesson is that glia are **active partners in behaviour**, "
            "since healthy astrocytes, intact myelin and well-regulated microglia are each needed for "
            "clear thinking, and their failure underlies conditions from multiple sclerosis to the "
            "neuroinflammation seen in dementia.",
        ],
        "definition": None,
        "flow": {
            "title": "Matching each glial cell to its main job",
            "steps": [
                ("Astrocyte", "feeds + balances ions", False),
                ("Oligodendrocyte", "myelin in CNS", False),
                ("Schwann cell", "myelin in PNS", False),
                ("Microglia", "immune defence", False),
                ("Radial glia", "migration scaffold", True),
            ],
        },
        "mnemonic": {
            "word": "AOMRS",
            "line": "Say \"A Old Monk Rarely Sleeps\" (A-O-M-R-S) for the five glial cells.",
            "items": [
                ("A", "**Astrocytes** nourish neurons and balance ions."),
                ("O", "**Oligodendrocytes** make myelin in the CNS."),
                ("M", "**Microglia** are the immune defenders."),
                ("R", "**Radial glia** guide neuron migration."),
                ("S", "**Schwann cells** make myelin in the PNS."),
            ],
        },
    },
    {
        "anchor": "bbb",
        "number": 3,
        "color": "purple",
        "title": "The Blood\u2013Brain Barrier",
        "intro": "The blood-brain barrier is the selective gate that protects the brain while still "
                 "letting essential nutrients pass, and understanding it explains one of the central "
                 "puzzles of clinical neuroscience, namely why some psychiatric and neurological drugs "
                 "reach the brain and work while others are turned away at the door.",
        "points": [
            "The **blood-brain barrier** is a protective mechanism that keeps a stable environment for "
            "the brain by shielding it from harmful substances carried in the circulating blood.",
            "It works by **differential permeability** [[letting some things through but not others]], so "
            "blood-borne toxins and many drugs are kept out while the brain's normal metabolism continues "
            "undisturbed.",
            "The barrier is built mainly from the tightly joined **endothelial cells** [[the cells lining "
            "the blood vessels]] of brain capillaries, whose close seals stop substances from leaking "
            "freely between the blood and brain tissue.",
            "Astrocyte **end-feet** wrap around these capillary walls and help form and support the "
            "barrier, which is why the barrier is described as a partnership between blood vessels and "
            "surrounding glial cells.",
            "The barrier allows **passive diffusion** [[free drifting across the membrane]] of small gases "
            "such as oxygen and carbon dioxide, so the brain can breathe at the cellular level without any "
            "special transport system.",
            "Essential molecules that cannot simply diffuse, such as **glucose** and certain amino acids, "
            "are moved across by **selective transport** [[special carrier proteins that ferry them "
            "through]] to feed the demanding brain tissue.",
            "The brain is metabolically expensive, using about **twenty-five percent of the body's glucose** "
            "and roughly seventy-five percent of its oxygen, so a steady, well-guarded blood supply is "
            "vital to normal brain function.",
            "A serious fall in blood supply to the brain can cause a **stroke** [[brain damage from loss of "
            "blood flow]], showing how dependent the brain is on continuous, properly regulated "
            "circulation.",
            "**Lipid-soluble** [[fat-dissolving]] substances can slip through the plasma membranes of the "
            "barrier more easily than water-soluble ones, which is why some drugs reach the brain readily "
            "while others are largely blocked.",
            "**Stress** increases the permeability of the barrier, so under strain even some harmful "
            "chemicals that are normally excluded may cross into the brain, a point of real clinical "
            "importance.",
            "The barrier's selective nature explains why treating brain disease is difficult, because many "
            "useful drugs are **rejected at the barrier** and never reach the neurons that need them.",
            "A clear clinical illustration is **L-DOPA (levodopa)**, the **precursor** [[the chemical "
            "building block the body converts into the active substance]] of dopamine, which is given in "
            "Parkinson's disease precisely because it can cross the barrier whereas dopamine itself "
            "cannot, so the brain is handed the raw material and makes the dopamine on the inside.",
            "The barrier is described as a **semipermeable** [[only partly letting things through]] "
            "structure formed with extracellular fluid that separates the circulating blood from the brain "
            "and the extracellular fluid of the central nervous system.",
            "By keeping ions and messenger molecules at controlled levels, the barrier protects the "
            "delicate **electrical environment** neurons need, so signalling across the brain stays "
            "reliable and is not disturbed by swings in blood chemistry.",
            "In short, the blood-brain barrier is a finely balanced **gatekeeper** that blocks threats, "
            "admits fuel and protects the brain's internal setting, making it central to both brain health "
            "and drug design.",
            "Because the barrier guards brain chemistry so strictly, damage or leakage of the barrier is "
            "linked to swelling, infection and several neurological conditions, underlining its role as a "
            "frontline protector of the central nervous system.",
            "For a clinical psychologist the key takeaway is that the barrier is both a **shield and a "
            "hurdle**: it keeps the brain's chemistry stable enough for reliable mood and cognition, yet "
            "the same selectivity is why delivering medicines to the brain, and explaining why a drug "
            "does or does not help, is so often a question of whether it can cross this gate.",
        ],
        "definition": {
            "label": "Sourced definition \u2014 Blood Brain Barrier (book's Key Terms and Definitions)",
            "text": "\u201cBlood brain barrier is a protective element with semipermeable architecture "
                    "formed by extracellular fluid that separates the circulating blood from the brain and "
                    "extracellular fluid in the central nervous system. Blood brain barrier system allows "
                    "passage of small molecules including gases such as oxygen and carbon dioxide by passive "
                    "diffusion; as well also promote the selective transportation of essential nutrients or "
                    "molecules such as, water, glucose and amino acids, which cannot diffuse due to absence "
                    "of required polarity.\u201d",
            "cite": "Examining Biological Foundations of Human Behavior (ed. Barre Vijaya Prasad, 2020), "
                    "Key Terms and Definitions.",
        },
        "flow": {
            "title": "What the barrier lets in and keeps out",
            "steps": [
                ("Blood vessel", "carries everything", False),
                ("Tight endothelium + astrocyte end-feet", "the gate", False),
                ("IN: O2, CO2, glucose, amino acids", "needed fuel", False),
                ("OUT: toxins, most drugs, pathogens", "blocked", True),
            ],
        },
        "mnemonic": {
            "word": "GUARD",
            "line": "The barrier acts as a GUARD for the brain.",
            "items": [
                ("G", "**Gases** oxygen and carbon dioxide diffuse in freely."),
                ("U", "**Useful** fuel like glucose and amino acids is carried in."),
                ("A", "**Astrocyte** end-feet help build and support the barrier."),
                ("R", "**Rejects** toxins, pathogens and most drugs."),
                ("D", "**Differential** permeability, loosened by stress, controls entry."),
            ],
        },
    },
    {
        "anchor": "resting",
        "number": 4,
        "color": "amber",
        "title": "Communication Within a Neuron \u2014 Membrane (Resting) Potential",
        "intro": "At rest a neuron holds a steady negative voltage inside, like a charged battery waiting "
                 "to be used, and this baseline is worth mastering because psychoactive drugs, "
                 "anaesthetics and ion imbalances all change behaviour by nudging this very voltage.",
        "points": [
            "The **resting membrane potential** is the steady voltage across a neuron's membrane while the "
            "cell is at rest and not sending a signal, and this chapter gives its value as about "
            "**\u221270 mV** inside relative to outside.",
            "This resting state is called **polarised** [[charged, with a difference between inside and "
            "outside]] because the inside of the neuron is negative compared with the fluid surrounding "
            "the cell.",
            "The negativity arises from an uneven spread of **ions** [[charged particles such as sodium, "
            "potassium and chloride]] across the membrane, so charge is not balanced between the inside and "
            "the outside.",
            "The neuron's cytoplasm is a watery, salty, **potassium-rich** solution, while sodium ions are "
            "more concentrated outside, and this mirror-image arrangement sets up the resting voltage the "
            "cell holds.",
            "The **plasma membrane** [[the cell's outer skin]] is selectively permeable, meaning it lets "
            "some ions cross more easily than others, which is essential for building and keeping the "
            "resting potential.",
            "At rest the membrane is far more permeable to **potassium (K+)** than to sodium, so potassium "
            "tends to leak outward down its gradient and leaves the inside of the cell more negative.",
            "A separate **sodium-potassium pump** [[a protein that actively swaps the two ions using "
            "energy]] pushes sodium out and potassium in, maintaining the concentration differences that "
            "underlie the resting potential. (Standard physiology texts.)",
            "This pumping requires **energy from ATP** supplied by mitochondria, which is why the resting "
            "potential is an active, energy-dependent state rather than something the cell gets for free.",
            "The resting potential represents a form of **stored electrical energy**, like a charged "
            "battery holding a voltage until a circuit is switched on, that the neuron can release "
            "quickly when it is stimulated strongly enough to fire.",
            "A stimulus that makes the inside **less negative** moves the voltage toward zero, a change "
            "called depolarisation, whereas a stimulus that makes it more negative is called "
            "**hyperpolarisation** [[becoming even more negative than rest]].",
            "The special voltage that must be reached to trigger a signal is the **threshold**, given in "
            "this chapter as about **\u221255 mV**, the tipping point beyond which the neuron is committed "
            "to firing.",
            "Below threshold the neuron stays quiet, so the resting potential also acts as a **filter** "
            "that prevents weak or random inputs from accidentally setting off a full nerve impulse.",
            "The resting potential is not fixed forever but is **actively maintained**, constantly "
            "rebalanced by ion channels and pumps so the cell is always poised and ready to respond to the "
            "next stimulus.",
            "Because the voltage sits between the negative resting level and the threshold, the neuron "
            "behaves like a loaded trigger, holding its charge until an input is **strong enough** to push "
            "it past the tipping point.",
            "Understanding the resting potential matters clinically because drugs, toxins and ion "
            "imbalances that alter membrane charge can change how easily neurons fire and so directly "
            "affect mood, movement and behaviour.",
            "For example, many **mood-stabilising and anti-seizure drugs** work by acting on ion channels "
            "to make neurons harder to depolarise, which stabilises the resting state and calms "
            "over-excitable circuits, a direct bridge from this basic physiology to psychiatric "
            "treatment.",
            "Overall, the resting potential is the stable baseline, set by ion gradients and the "
            "**sodium-potassium pump**, from which every action potential begins and to which the neuron "
            "always returns after firing.",
        ],
        "definition": None,
        "flow": {
            "title": "How the resting potential is set up and held",
            "steps": [
                ("Na+ pumped out, K+ pumped in", "Na/K pump", False),
                ("K+ leaks out through channels", "membrane", False),
                ("Inside becomes negative", "\u221270 mV", False),
                ("Polarised cell ready to fire", "stable baseline", True),
            ],
        },
        "mnemonic": {
            "word": "PRIME",
            "line": "A resting neuron is PRIMEd and waiting.",
            "items": [
                ("P", "**Polarised**, inside is negative at about \u221270 mV."),
                ("R", "**Resting**, no signal is being sent."),
                ("I", "**Ions** sodium and potassium are unevenly spread."),
                ("M", "**Maintained** by the sodium-potassium pump using ATP."),
                ("E", "**Excitable**, ready to fire once threshold (\u221255 mV) is reached."),
            ],
        },
    },
    {
        "anchor": "action",
        "number": 5,
        "color": "rose",
        "title": "Communication Within a Neuron \u2014 The Action Potential",
        "intro": "The action potential is the brief, all-or-none electrical spike that carries the nerve "
                 "impulse along the axon, and following its exact sequence of ion movements shows a "
                 "clinical psychologist precisely where local anaesthetics, toxins and some drugs act to "
                 "switch signalling on or off.",
        "points": [
            "When the nervous system transmits information, a resting neuron produces an **action "
            "potential**, a short-lived event in which the membrane voltage rapidly rises and then falls, "
            "also called the **nerve impulse** [[the travelling electrical signal]].",
            "An action potential only fires once the voltage reaches the **threshold** of about "
            "**\u221255 mV**, the tipping point; once this value is reached, the impulse is generated in a "
            "committed, **all-or-none** [[either a full signal fires or none at all, like a trigger that "
            "is pulled or not]] way.",
            "**Depolarisation** [[the inside becoming less negative]] begins when a stimulus, for example "
            "at a touch receptor, opens **sodium (Na+) channels** and positive ions rush into the "
            "negatively charged neuron.",
            "As **sodium floods in**, the inside of the neuron becomes positive and the membrane potential "
            "climbs steeply, which is the sharp upstroke seen on a recording of the action potential.",
            "Next, **potassium (K+) channels** open and potassium ions move out of the cell, a flow that "
            "reverses the rising charge and begins to bring the membrane voltage back down.",
            "The brief peak created as sodium enters and then potassium leaves is called the **spike "
            "potential** [[the tall, narrow burst of voltage]], the characteristic shape of a single nerve "
            "impulse.",
            "**Repolarisation** [[returning toward the negative resting value]] happens as sodium channels "
            "close back to their resting state, driving the membrane potential back toward about "
            "\u221270 mV.",
            "Because the potassium channels stay open slightly too long, the voltage briefly overshoots "
            "past \u221270 mV to nearly **\u221280 mV**, a dip called **hyperpolarisation** [[more negative "
            "than the normal resting level]].",
            "Ion concentrations are then restored and the cell settles back to its **\u221270 mV** resting "
            "membrane potential, ready to respond again once a new stimulus arrives.",
            "Immediately after firing there is an **absolute refractory period** [[a short time when no new "
            "impulse is possible]], during which the neuron cannot generate another action potential no "
            "matter how strong the stimulus.",
            "This is followed by a **relative refractory period** during which a new action potential can "
            "occur, but only if the stimulus is **stronger than usual** to overcome the lingering "
            "hyperpolarisation.",
            "In myelinated axons the impulse jumps between **nodes of Ranvier** [[bare gaps in the "
            "insulation]], a leaping conduction that makes signalling much faster than in unmyelinated "
            "fibres.",
            "The action potential is essentially **self-regenerating**, being recharged at each node as it "
            "travels, so the signal keeps its full strength over long distances along the axon.",
            "Because of the all-or-none rule, information is coded not by the size of each spike but by the "
            "**frequency** of impulses, so a stronger stimulus produces more action potentials per second "
            "rather than bigger ones.",
            "The refractory periods also ensure the impulse travels in **one direction** only, from the "
            "cell body toward the axon terminal, preventing the signal from running backward along the "
            "fibre.",
            "Clinically, anything that disturbs sodium or potassium channels, such as **local "
            "anaesthetics** [[numbing drugs that block pain signals]] or certain toxins, can **block the "
            "action potential**, which is how nerve conduction and sensation can be switched off for "
            "medical purposes.",
            "This same logic explains everyday clinical facts: an anaesthetic injected near a nerve stops "
            "its sodium channels so the signal never reaches the brain and no pain is felt, which makes "
            "the action potential not just a textbook spike but the exact event therapy can **start or "
            "silence**.",
        ],
        "definition": None,
        "flow": {
            "title": "The sequence of one action potential",
            "steps": [
                ("Resting \u221270 mV", "polarised", False),
                ("Threshold \u221255 mV", "Na+ channels open", False),
                ("Depolarisation", "Na+ in, +spike", False),
                ("Repolarisation", "K+ out", False),
                ("Hyperpolarisation \u221280 mV", "overshoot", False),
                ("Refractory + reset", "back to \u221270 mV", True),
            ],
        },
        "mnemonic": {
            "word": "DRRR",
            "line": "Remember DRRR (\"Depolarise, Repolarise, Refract, Reset\") for the spike's phases.",
            "items": [
                ("D", "**Depolarisation**, sodium rushes in and voltage rises."),
                ("R", "**Repolarisation**, potassium leaves and voltage falls."),
                ("R", "**Refractory + hyperpolarisation**, a brief overshoot and recovery gap."),
                ("R", "**Reset**, ions are restored to the \u221270 mV resting level."),
            ],
        },
    },
    {
        "anchor": "neurotransmitters",
        "number": 6,
        "color": "navy",
        "title": "Communication Between Neurons \u2014 Neurotransmitters",
        "intro": "At the synapse, chemical messengers carry the signal from one neuron across a tiny gap "
                 "to the next, and because almost every psychiatric medication works here, the synapse is "
                 "the single most important meeting point between the biology of behaviour and clinical "
                 "practice.",
        "points": [
            "A **synapse** is the junction where one neuron communicates with the next, and because it is "
            "a thin gap neurons do not physically merge but instead signal across it chemically.",
            "The synapse has three parts, the **presynaptic ending** [[the sending side]] that releases "
            "the message, the **postsynaptic ending** that receives it, and the narrow **synaptic cleft** "
            "[[the gap between them, about 20 nm wide]] in between.",
            "**Neurotransmitters** are the chemical messengers synthesised by neurons and released across "
            "the synapse, and they are stored in tiny sacs called **vesicles** [[membrane bubbles holding "
            "the chemical]] in the terminal.",
            "When the impulse arrives, **calcium (Ca2+) channels** open and calcium entering the terminal "
            "drives the vesicles to fuse with the membrane in a process called **exocytosis** [[releasing "
            "contents by the sac merging with the cell wall]].",
            "The released neurotransmitter crosses the cleft by **diffusion** [[drifting across the gap]] "
            "and binds to receptors on the postsynaptic membrane, opening ion channels and changing that "
            "cell's local voltage.",
            "If the binding causes **depolarisation** the result is an excitatory response, whereas if it "
            "causes **hyperpolarisation** [[making the cell more negative]] the result is an inhibitory "
            "response that quietens the next neuron.",
            "**Ionotropic (channel-linked) receptors** give fast, direct responses because the "
            "neurotransmitter itself opens an ion channel, while **metabotropic (G-protein-linked) "
            "receptors** act indirectly through **second messengers** [[internal relay chemicals]] for "
            "slower, longer effects.",
            "**GABA** is the major **inhibitory** [[calming, signal-dampening]] neurotransmitter of the "
            "CNS, found at thirty to forty percent of all synapses, and because its shortage is linked "
            "to seizures and anxiety disorders, many anti-anxiety and anti-seizure drugs work by boosting "
            "GABA's quietening effect.",
            "**Glycine** is a simple amino-acid **inhibitory** neurotransmitter of the spinal cord, brain "
            "stem and retina, and the poison strychnine causes convulsions by blocking glycine's calming "
            "action.",
            "**Glutamate** is the main **excitatory** neurotransmitter and is central to learning and "
            "memory, while **aspartate** is another excitatory transmitter that activates **NMDA "
            "receptors** [[a specific glutamate-type receptor]].",
            "**Serotonin (5-HT)** helps regulate mood, anger, aggression, sleep, appetite and body "
            "temperature, which is why **SSRIs** [[selective serotonin reuptake inhibitors, a class of "
            "antidepressant that keeps more serotonin in the synapse by slowing its reabsorption]] are "
            "used to treat depression.",
            "**Dopamine** works along four brain pathways, and its shortage causes **Parkinson's "
            "disease** [[a movement disorder with tremor and stiffness]] while its excess is linked to "
            "**schizophrenia**, so this one transmitter connects both a neurological and a psychiatric "
            "disorder a clinical psychologist must understand.",
            "**Noradrenaline** is an excitatory monoamine from the locus coeruleus that plays a key role in "
            "**attention and arousal**, as well as heart rate, blood pressure and digestive processes.",
            "**Acetylcholine (ACh)**, the first neurotransmitter ever identified, acts in both the PNS "
            "and CNS, and its deficiency is one of the causes of **Alzheimer's disease** with its memory "
            "loss, which is why several dementia drugs aim to preserve acetylcholine at the synapse.",
            "The synapse is the main site of action of most **psychoactive drugs** [[substances that "
            "change mood, perception or behaviour]], for example cocaine blocks **dopamine reuptake** "
            "[[reabsorption of the transmitter back into the sending neuron]] so its effect is "
            "intensified and prolonged in the synaptic cleft.",
            "Signalling is kept precise by stopping it promptly, as transmitters are cleared by "
            "**reuptake** [[reabsorption back into the sending neuron]] or by enzymes, and "
            "**autoreceptors** [[self-sensing receptors on the sending neuron that monitor how much it "
            "has released]] give feedback that controls further release.",
            "The clinical thread running through this whole topic is that **drugs act at the synapse**: "
            "by changing how much transmitter is released, how long it lingers, or which receptors it "
            "reaches, antidepressants, antipsychotics and anti-anxiety agents all exert their effects "
            "exactly here, which is why mastering synaptic transmission underpins understanding "
            "psychopharmacology.",
        ],
        "definition": None,
        "flow": {
            "title": "How a chemical message crosses the synapse",
            "steps": [
                ("Impulse reaches terminal", "arrives", False),
                ("Ca2+ enters", "channels open", False),
                ("Vesicles fuse (exocytosis)", "release", False),
                ("Transmitter crosses cleft", "diffusion", False),
                ("Binds postsynaptic receptors", "excite / inhibit", False),
                ("Reuptake or enzyme clears it", "signal ends", True),
            ],
        },
        "mnemonic": {
            "word": "GGGSDNA",
            "line": "\"Good Girls Give Sweets, Dad Needs Acetylcholine\" lists the main transmitters "
                    "(G-G-G-S-D-N-A).",
            "items": [
                ("G", "**GABA**, the major inhibitory transmitter."),
                ("G", "**Glycine**, inhibitory in the spinal cord."),
                ("G", "**Glutamate**, the main excitatory transmitter."),
                ("S", "**Serotonin**, mood, sleep and appetite."),
                ("D", "**Dopamine**, movement, reward and psychosis."),
                ("N", "**Noradrenaline**, attention and arousal."),
                ("A", "**Acetylcholine**, the first transmitter, tied to memory."),
            ],
        },
    },
    {
        "anchor": "modulators",
        "number": 7,
        "color": "teal",
        "title": "Neuromodulators and Hormones",
        "intro": "Beyond fast synapses, neuromodulators and hormones tune behaviour more slowly and over "
                 "wider reaches of the body, and this broader, slower chemistry is what links the "
                 "nervous and endocrine systems to lasting states such as stress, mood, sleep and "
                 "arousal that are central to clinical work.",
        "points": [
            "**Neuromodulators** are a slower mode of chemical communication, released by neurons but "
            "travelling **larger distances** and spreading more widely through the nervous system than "
            "ordinary point-to-point transmitters.",
            "Many neuromodulators are **neuropeptides** [[short chains of amino acids linked by peptide "
            "bonds]], protein-like molecules that act alongside classical neurotransmitters rather than "
            "replacing them.",
            "Rather than simply exciting or inhibiting one cell, neuromodulators **adjust the general "
            "state** of many neurons, shaping broad conditions such as alertness, attentiveness and "
            "anxiety.",
            "Neuropeptides typically work by **raising or lowering the release** of other neurotransmitters "
            "or by altering how strongly postsynaptic cells respond, so they fine-tune signalling rather "
            "than carry it alone.",
            "**Hormones** are a third mode of communication in the body, chemical substances secreted by "
            "the **endocrine glands** [[ductless glands that release into the blood]] or by cells in "
            "organs such as the stomach and kidneys.",
            "Hormones fall into three chemical classes, **amino-acid derivatives**, **peptides and "
            "proteins**, and **steroids** [[fat-based hormones made from cholesterol]], and this chemistry "
            "shapes how each hormone acts on its targets.",
            "Hormones are released into the **extracellular fluid** and carried in the **bloodstream** to "
            "distant target cells, so unlike synaptic signals they can influence the whole body at once.",
            "Because many neurons carry **hormone receptors**, hormones can directly change behaviour; for "
            "example **testosterone**, a sex hormone, increases aggressiveness in male animals.",
            "The **adrenal glands** release hormones for the **stress response** [[the body's "
            "fight-or-flight reaction to threat]], preparing the body to react quickly during demanding "
            "or frightening situations, a system directly relevant to understanding anxiety and chronic "
            "stress in the clinic.",
            "The **hypothalamus** controls sleep, hunger and appetite, acting as a key link between the "
            "nervous system and the endocrine system that keeps basic drives regulated.",
            "The **ovaries and testicles (gonads)** control female and male characteristics by secreting "
            "the sex hormones that govern reproduction and the development of secondary sexual traits.",
            "The **pancreas** releases **insulin** [[the hormone that lowers blood sugar]] and so controls "
            "the level of glucose in the body, a balance that is essential for steady energy supply to the "
            "brain.",
            "The **parathyroid glands** retain and regulate the body's **calcium** levels, while the "
            "**pineal body** secretes melatonin to maintain sleep cycles and other daily biological "
            "rhythms.",
            "The **pituitary gland** is often called the master gland because it helps **maintain the "
            "other endocrine glands**, while the **thymus** supports the body's immune system.",
            "The **thyroid gland** maintains the body's **metabolism** [[the rate at which the body uses "
            "energy]], including bone growth and heat production, so it influences overall energy and "
            "activity levels.",
            "**Melatonin**, derived from serotonin in the pineal gland, governs the **sleep-wake cycle** "
            "[[the daily rhythm of sleeping and waking, also called the body clock]], being made in "
            "darkness and kept low in daylight, which shows how a single molecule can act as both a "
            "transmitter-related signal and a hormone.",
            "For the clinician the unifying idea is one of **widening scope and slowing speed**: "
            "neurotransmitters act in milliseconds at a single synapse, neuromodulators shade the mood "
            "of whole circuits, and hormones act over minutes to days across the body, so disorders of "
            "stress, sleep, appetite and sexual function often trace back to this slower chemistry "
            "rather than to one fast synapse.",
        ],
        "definition": None,
        "flow": {
            "title": "Three widening modes of chemical communication",
            "steps": [
                ("Neurotransmitter", "fast, one synapse", False),
                ("Neuromodulator", "slower, wider in CNS", False),
                ("Hormone", "slowest, whole body via blood", True),
            ],
        },
        "mnemonic": {
            "word": "HAHA-POP-TT",
            "line": "\"Hypothalamus, Adrenals, Hormones, Ability... POP, TT\" \u2014 use the letters to walk "
                    "the endocrine glands.",
            "items": [
                ("H", "**Hypothalamus** controls sleep, hunger and appetite."),
                ("A", "**Adrenal** glands release stress hormones."),
                ("P", "**Pituitary**, the master gland, maintains the others."),
                ("P", "**Pancreas** releases insulin to control glucose."),
                ("T", "**Thyroid** sets metabolism; **Thymus** aids immunity."),
                ("P", "**Pineal** and parathyroid handle rhythms and calcium."),
            ],
        },
    },
]

REFERENCES = {
    "books": [
        "Examining Biological Foundations of Human Behavior, edited by Barre Vijaya Prasad "
        "(2020), IGI Global. Chapter \u201cStructure and Function of Cells\u201d by Shazia Ali, "
        "Sumit Kumar, Asma Ali and Syed Irtiza; and the chapter \u201cAnatomy of the Brain\u201d "
        "(Key Terms and Definitions, Blood Brain Barrier).",
        "Principles of Neural Science, by Eric R. Kandel, James H. Schwartz, Thomas M. Jessell, "
        "Steven A. Siegelbaum and A. J. Hudspeth (standard neuroscience text, used to cross-check "
        "the resting-potential value, the sodium-potassium pump and the ionic basis of the action "
        "potential).",
        "Guyton and Hall Textbook of Medical Physiology, by John E. Hall (standard physiology text, "
        "used to cross-check membrane potential, ion gradients and nerve-impulse conduction).",
    ],
    "online": [
        "StatPearls (NCBI Bookshelf), Treasure Island (FL): StatPearls Publishing \u2014 physiology "
        "articles on the resting membrane potential and the action potential, used to confirm "
        "standard ionic values and mechanisms.",
    ],
    "flagged": [
        "Primary journal articles cited inside the source book (for example Folk & Long, 1988 on "
        "serotonin; Hernandes & Troncone, 2009 on glycine; Stone, Morrison & Pilowsky, 2007 on "
        "dopamine) were referenced only as they appear within the book above and were not accessed "
        "directly for these notes.",
    ],
}

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
    """Plain text version (used by the .docx generator if it imports this)."""
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


def render_topic(t):
    parts = ['<section class="card %s" id="%s">' % (t["color"], t["anchor"]),
             '<div class="card-head"><div class="sec-num">%d</div><h2>%s</h2></div>'
             % (t["number"], _html.escape(t["title"])),
             '<p class="intro">%s</p>' % render_inline_html(t["intro"])]
    parts.append('<ol class="points">')
    for p in t["points"]:
        parts.append("<li>" + render_inline_html(p) + "</li>")
    parts.append("</ol>")
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
           ref_flagged=ref_flagged)
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
