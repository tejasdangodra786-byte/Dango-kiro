# -*- coding: utf-8 -*-
"""Single source of truth for Biological Foundations of Behavior, Unit IX
(Frontal Lobe Syndrome).

The HTML build (build_bio9.py -> parts) imports SECTIONS and the front-matter
from this module. A byte-identical copy lives in the repository's notes_src/
directory so both the builder and the repository carry the same source; the two
copies are kept in sync (diff -q clean).

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

TITLE = "Biological Foundations of Behavior \u2014 Unit IX: Frontal Lobe Syndrome"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

SUBTITLE = "M.Phil. Clinical Psychology study notes (bullet style) on frontal lobe syndrome"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens the notes.
INTRO = [
    {"t": "lead", "text": "These are concise, exam-ready revision notes for Unit IX of the Biological Foundations of Behavior paper (M.Phil. Clinical Psychology). They cover frontal lobe syndrome topic by topic, from the prefrontal anatomy and its three classic syndrome pictures, through the disturbances of the regulatory (executive) functions, attention, emotion and personality, memory and intellectual activity, to the language and motor disturbances, so a whole 5-mark answer can be revised from a single topic."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one part of the Unit IX syllabus line on frontal lobe syndrome, written as a 5-mark exam answer in full sentences rather than terse fragments, so you can revise straight from the point and reproduce it under exam conditions.",
        "Every topic opens with an Answer-flow box that gives the ordered skeleton to write, carries a flowchart figure that pictures the core idea, gives a short definition box drawn from standard neuropsychology texts, and closes with a mnemonic that packs the key points into a single memorable word.",
        "A short source note and a references list sit at the end; the references name only the books (full title and author), the unit's frontal-lobe teaching materials, and the authoritative sources that were actually used.",
    ]},
    {"t": "h3", "text": "How the coloured gloss works (live example)"},
    {"t": "ul", "items": [
        "Hard neuropsychology terms are kept, but each one is explained right after it in a coloured bracket, like this: [[perseveration|getting stuck and repeating the same response when a new one is needed]]. Just skim the amber bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "The term itself carries a faint dotted underline and its plain-language meaning follows in amber italics on a peach tint, as in [[dysexecutive syndrome|the pattern of poor planning, control and flexibility after frontal damage]]; the two together let you learn the real term and its meaning at once.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail, and note that only a few hook words are bold, never a whole sentence.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word, so you get a visual anchor and a verbal anchor for each answer.",
        "Each topic opens with a short **Answer flow** box that lists, in order, the skeleton to write in the exam \u2014 definition, anatomy or background, key deficits, evidence and examples, assessment and significance \u2014 and the exam points that follow are grouped under small sub-headings in that same order, so the whole topic reads as a ready-to-write answer with a clear beginning, middle and end.",
    ]},
    {"t": "h3", "text": "How citations work"},
    {"t": "ul", "items": [
        "A citation in square brackets names the full book and its author, for example {{c:Kolb}} or {{c:Stuss}}, so you can trace any claim back to the standard text it comes from.",
        "The unit's own frontal-lobe teaching materials are cited by their descriptive title, for example {{c:Seminar}} or {{c:Regions}}, with no author invented beyond what those files themselves state.",
    ]},
]

# Honest limits and source conventions.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no authors, citations or page numbers in these notes are invented. The neuropsychological facts stated are standard, textbook-level and verifiable, and the following limits are disclosed so you can check them independently."},
    {"t": "ul", "items": [
        "The characterisations of frontal lobe syndrome are drawn from the standard neuropsychology literature, chiefly Kolb and Whishaw's Fundamentals of Human Neuropsychology, together with Stuss and Benson's The Frontal Lobes, Luria's Higher Cortical Functions in Man, Lezak and colleagues' Neuropsychological Assessment and Fuster's The Prefrontal Cortex; they are given as characterisations, not verbatim quotations, and no page numbers are attached.",
        "Two of the unit's own frontal-lobe teaching materials were readable and were mined directly: a note on the regions of the frontal lobe and their functions, and a seminar titled 'Frontal Lobe Syndromes: Clinical Features and Assessment'. They are cited by those descriptive titles and cross-checked against the standard texts; no author is invented for them beyond the names those files themselves cite.",
        "One source supplied with the unit, a scanned handout on frontal-lobe dysfunction and its associated disorders drawn from Kolb, is an image-only file with no readable text layer, and optical character recognition (OCR) was not available in the environment these notes were prepared in, so its text could not be extracted; moreover it is a copyrighted excerpt rather than an openly-licensed one, so neither its wording nor its figures were reproduced. Its subject matter (the regions of the frontal lobe, the assessment tests and the disorders associated with frontal dysfunction) is instead covered from the readable frontal-lobe note, the seminar and the standard Kolb and Whishaw text, and the figures in these notes are either original schematic diagrams or verified public-domain plates. This limit is disclosed rather than hidden.",
        "The two embedded photographs are genuine public-domain images whose status was verified: Harlow's 1868 engraving of Phineas Gage's skull and tamping iron, and a plate of the lateral surface of the cerebrum from Gray's Anatomy (1918); both are stored inside the file itself so the notes load no external resource, and both appear in the earlier units of this paper.",
        "Only well-established neuropsychology is stated; no novel claims, study results or quantitative findings beyond standard textbook accounts are presented, and where accounts differ between authors (for example the exact Brodmann areas assigned to the prefrontal subdivisions) that difference is noted rather than smoothed over.",
    ]},
]

# Reading list.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the standard neuropsychology books used, the unit's frontal-lobe teaching materials, and the authoritative references consulted for the Unit IX content."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Bryan Kolb & Ian Q. Whishaw \u2014 Fundamentals of Human Neuropsychology (the primary reference: the anatomy and connections of the frontal and prefrontal cortex, the theory of frontal-lobe function as the temporal organisation of behaviour, and the symptoms of frontal-lobe damage across motor, cognitive, memory, language and personality domains).",
        "Donald T. Stuss & D. Frank Benson \u2014 The Frontal Lobes (the clinical syndromes of frontal damage, the amorphous and heterogeneous nature of frontal lobe syndrome, and the regional personality changes).",
        "Alexander R. Luria \u2014 Higher Cortical Functions in Man (the regulatory role of the frontal lobes and of language over behaviour, perseveration, dynamic aphasia and the dissociation between knowing and doing).",
        "Muriel D. Lezak, Diane B. Howieson, Erin D. Bigler & Daniel Tranel \u2014 Neuropsychological Assessment (the clinical assessment of executive function, attention, memory and the tests used to detect frontal dysfunction).",
        "Joaquin M. Fuster \u2014 The Prefrontal Cortex (the prefrontal cortex as the organ of executive function, working memory and the temporal organisation of goal-directed action).",
        "G. Neil Martin \u2014 Human Neuropsychology (a cross-check for the regional functions of the frontal lobe and the sensory-perceptual role of the orbitofrontal cortex).",
        "Barre Vijaya Prasad (2020) \u2014 Examining Biological Foundations of Human Behavior (a cross-check for the lobes-and-functions and brain-structure material for the psychology reader).",
    ]},
    {"t": "h3", "text": "Unit teaching materials (cited by descriptive title)"},
    {"t": "ul", "items": [
        "Regions of the Frontal Lobes (unit note) \u2014 the primary motor, premotor, frontal-eye-field, dorsolateral prefrontal, orbitofrontal and cingulate / supplementary-motor regions, their Brodmann areas, connections and functions, the frontal\u2013subcortical circuits, the tests used to assess each region, and the disorders associated with frontal dysfunction.",
        "Frontal Lobe Syndromes: Clinical Features and Assessment (seminar materials) \u2014 a seminar on the cerebral cortex and the frontal lobe's anatomy and functions; the definition, etiology and clinical features of frontal lobe syndrome across motor, sensory-perceptual, speech and language, cognitive and personality / behavioural domains; the assessment tests for each; the two 'frontal lobe personalities'; and the limitations and prognosis of the syndrome.",
    ]},
    {"t": "h3", "text": "Standard-reference sources"},
    {"t": "ul", "items": [
        "Classic cases and tests that are standard knowledge and are named by their source: Phineas Gage (Harlow, 1848 and 1868); the Wisconsin Card Sorting Test (perseverative errors); the Stroop Test; verbal fluency (FAS and category fluency); the Tower of London / Hanoi; and the go/no-go task.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. OVERVIEW + PREFRONTAL ANATOMY =========================
_s("t1", "1", "Frontal lobe syndrome and prefrontal anatomy", [
    {"t": "lead", "text": "Frontal lobe syndrome is the cluster of behavioural changes that follows damage to the frontal cortex, especially the prefrontal cortex, which divides into dorsolateral, orbitofrontal (ventromedial) and medial (anterior cingulate) regions, each tied through frontal\u2013subcortical circuits to a distinct clinical picture: dysexecutive, disinhibited and apathetic."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what frontal lobe syndrome is and why the deficits are grouped as a syndrome.",
        "Anatomy: the frontal lobe's size and gross divisions, then the three prefrontal subdivisions.",
        "Key features: the three classic syndrome pictures and the frontal\u2013subcortical circuits.",
        "Evidence and examples: the heterogeneity of the syndrome and the common etiologies.",
        "Significance: why the frontal lobe is the brain's control and personality centre.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "prefrontal_subdivisions", "caption": "The three prefrontal subdivisions \u2014 dorsolateral, orbitofrontal/ventromedial and medial/anterior cingulate \u2014 each with its function and the dysexecutive, disinhibited or apathetic picture it yields when damaged, all routed through frontal\u2013subcortical circuits."},
    {"t": "fig", "key": "pd_gray_lobes", "caption": "The lateral surface of the cerebrum: the frontal lobe lies in front of the central sulcus and makes up about a third of the cortex."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Frontal lobe syndrome is described as a **cluster of behavioural symptoms** that reliably appear after damage to the frontal cortex, impairing higher functions such as motivation, planning, social behaviour and speech. {{c:Seminar}}",
        "{{GUIDE}} Because the deficits **cannot be neatly subclassified**, the term is used for an amorphous, varied group of deficits that arise from diverse etiologies, different locations and variable extents of frontal abnormality. {{c:Stuss}}",
        "{{THEORY}} The general function of the frontal lobe is the **temporal organisation of behaviour**: it holds control systems that implement behavioural strategies in response to internal and external cues, now often called executive functions. {{c:Kolb}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the idea of a syndrome"},
    {"t": "ul", "items": [
        "Frontal lobe syndrome is the **clinical picture that follows damage to the frontal cortex**, producing impairments of motivation, planning, social behaviour and language that cluster together often enough to be named as a syndrome. {{c:Seminar}}",
        "It is deliberately called an **amorphous syndrome** because the deficits cannot be sharply subclassified and arise from diverse causes, locations and extents of damage, so no two frontal patients look exactly alike. {{c:Stuss}}",
        "A key teaching point is that **not all frontal patients show all symptoms**, and severity varies with the specific region and hemisphere injured, which is why the syndrome is defined loosely rather than as a fixed list. {{c:Seminar}}",
        "The frontal lobe is the **most recently evolved lobe** and makes up about one-third of the cerebral cortex, so it is regarded as the cortical locus of higher learning and the structure that defines the self. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Gross divisions of the frontal lobe"},
    {"t": "ul", "items": [
        "The frontal lobe is commonly divided into a **precentral (motor) area**, a **premotor area** and a large **prefrontal area**, moving from behind forward, so posterior regions execute movement and anterior regions control and plan it. {{c:Seminar}}",
        "The **primary motor cortex** sits on the precentral gyrus just in front of the central sulcus and drives voluntary movement, while the **premotor and supplementary motor areas** select and programme those movements. {{c:Regions}}",
        "The **frontal eye fields** control voluntary eye movements for visual search, and **Broca's area** in the inferior frontal gyrus controls the motor production of speech. {{c:Regions}}",
        "The **prefrontal cortex** is defined by its input from the [[dorsomedial nucleus|a thalamic relay nucleus whose projection marks out the prefrontal cortex]] of the thalamus, and it controls the cognitive processes that select the right action at the right time and place. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "The three prefrontal subdivisions"},
    {"t": "ul", "items": [
        "The [[dorsolateral prefrontal cortex|the upper, outer prefrontal region handling planning and working memory]] (around Brodmann areas 9 and 46) is the largest prefrontal region and carries the **executive functions** that use multimodal input to generate appropriate responses. {{c:Regions}}",
        "The [[orbitofrontal cortex|the prefrontal region above the eye sockets handling emotion and social conduct]] (roughly areas 10\u201314), also called ventromedial, modulates **affective and social behaviour**, reward, working memory for features, and the discrimination of smell and flavour. {{c:Regions}}",
        "The **medial frontal cortex and anterior cingulate** (areas 24 and 32), together with the supplementary motor area, govern **drive, motivation and the exploration** of the environment, drawing on deep limbic connections. {{c:Regions}}",
        "The orbital cortex **gains input from all sensory modalities** and projects to the amygdala and hypothalamus, giving it a route to influence the autonomic changes that accompany emotion. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "The three classic syndrome pictures"},
    {"t": "ul", "items": [
        "A **dorsolateral lesion** typically yields the [[dysexecutive syndrome|poor planning, control and flexibility after dorsolateral frontal damage]], with poor planning, perseveration and apathy, a picture also labelled pseudodepressed. {{c:Seminar}}",
        "An **orbitofrontal lesion** typically yields the [[disinhibited syndrome|impulsive, tactless, socially unrestrained behaviour after orbitofrontal damage]] or pseudopsychopathic picture, with impulsivity, tactlessness, emotional lability and poor judgement. {{c:Stuss}}",
        "A **medial / anterior cingulate lesion** typically yields the **apathetic or akinetic picture**, with loss of drive, abulia and, in severe cases, akinetic mutism, reflecting a collapse of motivation. {{c:Regions}}",
        "These are **tendencies, not rigid rules**, because real lesions cross boundaries and the same patient may show a blend, so the three pictures are best used as organising prototypes. {{c:Stuss}}",
    ]},
    {"t": "h4", "text": "Circuits, etiology and significance"},
    {"t": "ul", "items": [
        "Each prefrontal region sits in a **frontal\u2013subcortical circuit** running cortex \u2192 striatum \u2192 globus pallidus / substantia nigra \u2192 thalamus \u2192 cortex, so a **small subcortical lesion can mimic a large cortical one**. {{c:Regions}}",
        "The circuits act as a **filter** that modulates the output of the frontal cortex to the motor systems, which is why basal-ganglia disorders such as Huntington's and Parkinson's disease can produce frontal-type deficits. {{c:Regions}}",
        "Common **etiologies** include closed-head trauma, neurodegenerative dementias such as frontotemporal degeneration, tumours compressing the frontal lobe, stroke, epilepsy with frontal foci, HIV and multiple sclerosis. {{c:Seminar}}",
        "The syndrome matters clinically because the frontal lobe is the brain's **'orchestra leader'**, directing and coordinating the sensory, motor, cognitive and limbic systems, so its damage disrupts the orchestration of behaviour as a whole. {{c:Seminar}}",
        "Its significance is heightened by **poor insight**: patients often fail to appreciate their deficits, which makes the syndrome disabling and creates a substantial, lifelong need for care and supervision. {{c:Seminar}}",
    ]},
    {"t": "box", "kind": "key", "title": "Frontal lobe syndrome at a glance", "items": [
        "A cluster of behavioural changes after frontal damage; amorphous and heterogeneous, so not all patients show all symptoms.",
        "Frontal lobe = about one-third of the cortex; divided into precentral (motor), premotor and prefrontal areas.",
        "Three prefrontal subdivisions: dorsolateral (executive), orbitofrontal/ventromedial (emotion, social), medial/cingulate (drive).",
        "Three syndrome pictures: dysexecutive, disinhibited (pseudopsychopathic), apathetic (pseudodepressed).",
        "Frontal\u2013subcortical circuits mean subcortical lesions mimic cortical ones; many etiologies; poor insight disables.",
    ]},
    {"t": "mn", "word": "DOM", "text": "Dorsolateral (executive / dysexecutive), Orbitofrontal (emotion & social / disinhibited), Medial-cingulate (drive / apathetic) \u2014 the three prefrontal subdivisions and the three syndrome pictures; all loop through the frontal\u2013subcortical circuits, and the frontal lobe is the brain's orchestra leader."},
])


# ========================= 2. REGULATORY (EXECUTIVE) FUNCTIONS =========================
_s("t2", "2", "Disturbances of regulatory (executive) functions", [
    {"t": "lead", "text": "The frontal lobes regulate goal-directed behaviour through the executive functions \u2014 planning, initiation, working memory, set-shifting, inhibition and abstraction \u2014 and dorsolateral damage produces a dysexecutive syndrome marked by perseveration, poor strategy and a striking dissociation between knowing a rule and being able to use it."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: what the regulatory or executive functions are and where they sit.",
        "Background: the frontal lobe as the central executive / supervisory system.",
        "Key deficits: planning, initiation, working memory, set-shifting, inhibition, perseveration, abstraction.",
        "Evidence and examples: Wisconsin Card Sorting, the knowing\u2013doing dissociation, real-life tasks.",
        "Assessment and significance: the tests and why executive loss is so disabling.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "dysexecutive_fan", "caption": "The dysexecutive fan: the dorsolateral central executive fans out to planning, initiation, working memory, set-shifting and inhibition, with perseveration and lost abstraction as hallmark errors and a core knowing\u2013doing dissociation."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} The **executive functions** are the higher-order processes essential for planning, setting goals, and regulating and verifying behaviour, including organisation, initiation, reasoning, judgement, abstraction and cognitive flexibility. {{c:Seminar}}",
        "{{THEORY}} Frontal (specifically prefrontal) dysfunction has been characterised as a **dysexecutive syndrome** involving disruption of the central executive of working memory, the component that controls resources and monitors information processing. {{c:Seminar}}",
        "{{THEORY}} Luria described a **dissociation between knowing and doing** after frontal damage, such that verbalisation no longer controls active behaviour \u2014 the patient can state the rule but cannot use it to steer action. {{c:Luria}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the central executive"},
    {"t": "ul", "items": [
        "The **regulatory or executive functions** are the control processes that organise, direct and verify goal-directed behaviour, and they are the special province of the prefrontal cortex. {{c:Fuster}}",
        "They can be thought of as a **central executive** that allocates attentional resources and monitors ongoing processing, so frontal damage is aptly named a dysexecutive syndrome. {{c:Seminar}}",
        "Because these functions **supervise other systems rather than carrying out one task**, their loss leaves basic perception and movement intact while wrecking the orchestration of behaviour over time. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Planning, initiation and goal-directed behaviour"},
    {"t": "ul", "items": [
        "Frontal patients are **especially impaired at devising novel plans or strategies**, and sequencing a chain of actions toward a goal is perhaps the behaviour most clearly disrupted in them. {{c:Seminar}}",
        "Many show a **loss of initiation and spontaneity**, speaking little, producing few responses, and having difficulty getting out of bed, dressing or starting the day's activities. {{c:Seminar}}",
        "The temporal organisation of a real errand \u2014 **planning ahead, selecting from options, and persisting** \u2014 is exactly what frontal patients cannot manage, as in Kolb and Whishaw's shopping-trip example. {{c:Kolb}}",
        "They tend to **conceive activities impulsively but rarely complete them**, beginning chores at once without first planning how to carry them out. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Working memory, set-shifting and inhibition"},
    {"t": "ul", "items": [
        "[[Working memory|holding and manipulating information in mind for a few seconds]] is disrupted, so frontal patients fail delayed-response and delayed-alternation tasks that require them to hold information across a gap. {{c:Regions}}",
        "They show poor [[set-shifting|switching flexibly from one rule or category to another]], an **inflexibility in switching cognitive sets** that leaves them locked into an outdated response. {{c:Seminar}}",
        "**Response inhibition fails**, so the patient cannot suppress an automatic or prepotent response, which is seen in poor performance on go/no-go and Stroop tasks. {{c:Kolb}}",
        "Impaired inhibition can extend to [[utilisation behaviour|automatically using an object just because it is in reach]], in which the sight of an object becomes an order to use it. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Perseveration and abstraction"},
    {"t": "ul", "items": [
        "[[Perseveration|getting stuck and repeating the same response when a new one is needed]] is a hallmark error, seen when a patient keeps drawing the same figure, repeats a rule, or cannot abandon a sorting category. {{c:Regions}}",
        "**Abstract thought is impaired**, and loss of the abstract attitude was regarded as an early sign of frontal damage, so proverbs are interpreted concretely and concepts are hard to form. {{c:Seminar}}",
        "Patients also struggle with **rule deduction and rule following**, failing to identify the rule, to use feedback, or to switch set, which is captured in their poor Wisconsin Card Sorting performance. {{c:Seminar}}",
        "Impaired **associative learning** means they cannot reliably learn to pair a stimulus with a response or learn from experience, as in Luria's patients who could not learn to respond to coloured lights. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Evidence, assessment and significance"},
    {"t": "ul", "items": [
        "The clearest evidence is the **knowing\u2013doing dissociation**: patients can verbalise a task and know their errors, yet cannot use that knowledge to modify their behaviour. {{c:Luria}}",
        "In Shallice and Burgess's errands study, frontal patients were **inefficient and broke rules** (such as entering unnecessary shops) even though, when quizzed, they understood the task and tried to comply. {{c:Seminar}}",
        "Executive function is assessed with the **Wisconsin Card Sorting Test** (set-shifting and perseveration), Tower tests and maze tests (planning), verbal fluency (initiation) and the Stroop test (inhibition). {{c:Lezak}}",
        "Significance is underlined by a **paradox of everyday life**: many patients score near-normally on standard tests yet cannot plan and organise daily activities, so the deficit is often missed by conventional testing. {{c:Seminar}}",
    ]},
    {"t": "box", "kind": "key", "title": "Executive disturbance at a glance", "items": [
        "Executive functions = planning, initiation, working memory, set-shifting, inhibition, abstraction; the prefrontal 'central executive'.",
        "Dorsolateral damage \u2192 dysexecutive syndrome: poor novel plans, lost initiation, impulsive uncompleted actions.",
        "Hallmark errors: perseveration, inflexible set, failed inhibition, concrete thinking, poor rule use.",
        "Core sign: a dissociation between knowing and doing \u2014 can state the rule, cannot use it to steer behaviour.",
        "Tests: Wisconsin Card Sorting, Tower of London, mazes, verbal fluency, Stroop; often missed on routine testing.",
    ]},
    {"t": "mn", "word": "PISA-IP", "text": "Planning, Initiation, Set-shifting, Abstraction, Inhibition, Perseveration \u2014 the executive functions the dorsolateral frontal lobe regulates and the errors that appear when it is damaged; remember the Wisconsin Card Sort and Luria's knowing\u2013doing dissociation as the signature evidence."},
])


# ========================= 3. ATTENTIONAL PROCESSES =========================
_s("t3", "3", "Disturbances of attentional processes", [
    {"t": "lead", "text": "The frontal systems provide the top-down, supervisory control of attention, so frontal damage leaves the patient with impaired attentional capacity, poor concentration and high distractibility, affecting sustained, selective, divided and switching attention and the ability to resolve response conflict."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: attention and the frontal lobe's role in its top-down control.",
        "Background: the supervisory attentional system and its link to executive function.",
        "Key deficits: sustained, selective, divided and switching attention; distractibility.",
        "Evidence and examples: distractibility, Stroop interference, released parietal activity.",
        "Assessment and significance: the attention tests and why attention underpins the rest.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "attention_control", "caption": "A top-down frontal supervisory system controls sustained, selective, divided and switching attention; frontal damage releases this control, giving distractibility, poor concentration and Stroop-type conflict failures."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Frontal lobe patients appear to have **impaired attentional capacity**: they show poor concentration and are highly distractible, so attention is counted among the cognitive symptoms of the syndrome. {{c:Seminar}}",
        "{{THEORY}} The frontal lobes exert a **supervisory, top-down control** over attention, selecting what to attend to and holding focus, rather than performing the basic registration of stimuli themselves. {{c:Lezak}}",
        "{{THEORY}} Utilisation behaviour is read as a **release of parietally-driven responses** from frontal inhibition, so the patient becomes captured by whatever external stimulus is present. {{c:Seminar}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the frontal role"},
    {"t": "ul", "items": [
        "**Attention** is the selection and maintenance of processing on relevant information, and the frontal lobes supply its top-down, effortful control rather than its automatic capture. {{c:Lezak}}",
        "Attention and concentration are listed among the **cognitive and executive functions** of the prefrontal cortex, so they break down together with planning and working memory. {{c:Seminar}}",
        "Because frontal attention is **supervisory**, its loss does not blind or deafen the patient but leaves them unable to steer and sustain the spotlight of attention where it is needed. {{c:Fuster}}",
    ]},
    {"t": "h4", "text": "The kinds of attention affected"},
    {"t": "ul", "items": [
        "[[Sustained attention|keeping focus on a task over a stretch of time]], or vigilance, is impaired, so the patient cannot hold concentration on a continuous task and lapses as it goes on. {{c:Lezak}}",
        "[[Selective attention|focusing on a target while ignoring competing distractors]] is disturbed, and frontal patients are described as showing difficulty with selective attention and being highly distractible. {{c:Seminar}}",
        "[[Divided attention|sharing attention between two tasks at once]] suffers, so performance collapses when the patient must track two streams or do two things together. {{c:Lezak}}",
        "**Attentional switching** between tasks or sets is impaired, linking the attention deficit directly to the set-shifting failure of the dysexecutive syndrome. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "Distractibility and conflict"},
    {"t": "ul", "items": [
        "The most visible everyday sign is **high distractibility**: the patient's attention is pulled to irrelevant stimuli, so they cannot persist in the task at hand. {{c:Seminar}}",
        "They show difficulty in **resolving response conflict**, which is why they perform worse than patients with posterior lesions on the Stroop test, where a prepotent reading response must be suppressed. {{c:Seminar}}",
        "Attentional capture can shade into **utilisation and imitation behaviour**, in which the sight of an object or gesture drives an automatic response because frontal control over parietal activity is released. {{c:Seminar}}",
        "There may also be a **delayed habituation** to external stimuli after cingulate damage, so the patient keeps responding to a stimulus that should have faded from attention. {{c:Regions}}",
    ]},
    {"t": "h4", "text": "Assessment and significance"},
    {"t": "ul", "items": [
        "Sustained attention is assessed with the **Continuous Performance Test and Digit Span**, which probe vigilance and the simple holding of information. {{c:Seminar}}",
        "Selective, divided and switching attention are probed by the **Test of Everyday Attention and the Trail Making Test**, the latter also tapping visual scanning and executive control. {{c:Seminar}}",
        "The **Stroop Test** measures freedom from distractibility, selective attention and the ability to resolve response conflict, and it is a classic marker of frontal, as opposed to posterior, damage. {{c:Seminar}}",
        "Attention matters because it is the **gateway to the other functions**: poor sustained and selective attention undermines memory encoding, learning and the execution of any multi-step plan. {{c:Lezak}}",
        "The deficit also explains part of the **real-life disability**, since a highly distractible patient cannot see a daily task through from start to finish even when they understand it. {{c:Seminar}}",
    ]},
    {"t": "box", "kind": "key", "title": "Attentional disturbance at a glance", "items": [
        "The frontal lobe supplies top-down, supervisory control of attention, not basic sensory registration.",
        "Impaired across sustained (vigilance), selective, divided and switching attention.",
        "Signature signs: poor concentration, high distractibility, failure to resolve conflict (Stroop).",
        "Links to utilisation / imitation behaviour \u2014 attention captured by external stimuli when frontal control is released.",
        "Tests: CPT, Digit Span, Trail Making, Test of Everyday Attention, Stroop; attention gates memory and executive tasks.",
    ]},
    {"t": "mn", "word": "SSDS", "text": "Sustained, Selective, Divided, Switching \u2014 the four kinds of attention the frontal supervisory system controls; when it fails the patient becomes DISTRACTIBLE and cannot resolve conflict (think Stroop), which is why attention is the gateway deficit of the syndrome."},
])


# ========================= 4. EMOTIONS, PERSONALITY, SOCIAL CONDUCT =========================
_s("t4", "4", "Disturbances of emotion, personality and social conduct", [
    {"t": "lead", "text": "Frontal damage, especially of the orbitofrontal and ventromedial cortex, changes emotion, personality and social conduct, producing either a disinhibited pseudopsychopathic picture or an apathetic pseudodepressed one; the case of Phineas Gage and the somatic-marker idea capture how lost emotional signals wreck judgement though intellect is spared."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: the frontal lobe's role in emotion, personality and social behaviour.",
        "Anatomy: the orbitofrontal / ventromedial and dorsolateral contributions.",
        "Key features: the two frontal personalities \u2014 disinhibited versus apathetic.",
        "Evidence and examples: Phineas Gage, impaired empathy and theory of mind, the somatic-marker idea.",
        "Assessment and significance: the tests and the social disability produced.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "two_personalities", "caption": "Two opposite frontal personalities: the orbitofrontal disinhibited (pseudopsychopathic) picture versus the dorsolateral/medial apathetic (pseudodepressed) picture, with the ventromedial somatic-marker account between them."},
    {"t": "fig", "key": "pd_gage", "caption": "The skull of Phineas Gage and the tamping iron (Harlow, 1868): the first report of personality change after frontal damage \u2014 intellect intact, but 'no longer Gage'."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Blumer and Benson distinguished **two frontal-lobe personalities**: an orbitofrontal pseudopsychopathic picture of impulsivity and disinhibition, and a dorsolateral pseudodepressed picture of apathy and inertia. {{c:Stuss}}",
        "{{THEORY}} The **case of Phineas Gage** (Harlow, 1848 and 1868) was the first report of personality change following frontal damage: after a tamping iron passed through his frontal lobes he recovered physically but became profane, capricious and 'no longer Gage'. {{c:Seminar}}",
        "{{THEORY}} Orbitofrontal patients are notorious for **social gaffes**: they struggle to choose context-appropriate behaviour and to use feedback from the social environment to adjust it. {{c:Seminar}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and anatomy"},
    {"t": "ul", "items": [
        "The frontal lobe governs **personality, emotional regulation, impulse control and social behaviour**, so it is often described as the structure that defines the self. {{c:Seminar}}",
        "The [[orbitofrontal cortex|the prefrontal region above the eye sockets handling emotion and social conduct]] and the adjacent ventromedial cortex modulate affect, reward and social conduct through their links to the amygdala and hypothalamus. {{c:Kolb}}",
        "Because the orbital cortex projects to the **autonomic system**, it couples the bodily changes of emotion to events, giving it a central role in emotional responding. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "The disinhibited (pseudopsychopathic) picture"},
    {"t": "ul", "items": [
        "An **orbitofrontal lesion** tends to produce impulsivity, disinhibition and a lack of tact and restraint, a presentation labelled pseudopsychopathic. {{c:Stuss}}",
        "Features include **emotional lability, irritability and aggression**, increased motor activity, poor judgement, risk-taking and rule-breaking, and coarse or inappropriate language. {{c:Seminar}}",
        "Patients show **decreased concern with propriety**, facetiousness, egocentricity and inappropriate social comments, jokes and sexual behaviour such as reduced inhibition. {{c:Seminar}}",
        "This disinhibition is a **release phenomenon**, a loss of the frontal brake over behaviour, rather than a gain of any new drive. {{c:Stuss}}",
    ]},
    {"t": "h4", "text": "The apathetic (pseudodepressed) picture"},
    {"t": "ul", "items": [
        "A **dorsolateral or medial lesion** tends to produce apathy, indifference and inertia, a presentation labelled pseudodepressed because it mimics depression without the sustained low mood. {{c:Seminar}}",
        "Features include **reduced motivation, loss of initiative, aspontaneity and little overt emotion**, with slowed thinking and reduced verbal output. {{c:Seminar}}",
        "At its extreme the medial picture becomes [[akinetic mutism|a state of being awake but neither moving nor speaking spontaneously]], reflecting a profound collapse of drive. {{c:Regions}}",
        "The two pictures are **mirror images**, so clinicians use the lesion site to predict whether disinhibition or apathy will dominate, while noting that mixed pictures are common. {{c:Stuss}}",
    ]},
    {"t": "h4", "text": "Empathy, social judgement and the somatic-marker idea"},
    {"t": "ul", "items": [
        "Frontal patients show impaired [[theory of mind|the ability to infer other people's thoughts, feelings and intentions]], being poorer than other patients at inferring others' viewpoints and at detecting deception. {{c:Seminar}}",
        "They lose [[autonoetic awareness|the conscious sense of oneself as a continuous person over time]], so orbital-frontal injury impairs self-knowledge and the self-regulation of behaviour. {{c:Seminar}}",
        "**Emotional regulation fails**, with lability of mood and affect \u2014 exhilaration, depression, anxiety and irritability \u2014 linked to ventromedial damage and difficulty grasping the emotional consequences of one's behaviour. {{c:Seminar}}",
        "The [[somatic-marker idea|the view that bodily 'gut' feelings normally flag the value of choices]] holds that emotional signals normally guide decision-making, so ventromedial damage leaves judgement and social decisions poor despite preserved intellect. {{c:Stuss}}",
    ]},
    {"t": "h4", "text": "Evidence, assessment and significance"},
    {"t": "ul", "items": [
        "**Phineas Gage** remains the classic evidence, and drastic personality change of this kind generally follows large and usually bilateral frontal injuries. {{c:Seminar}}",
        "A modern parallel is the salesman **M. L.**, whose closed-head orbital injury left him unable to relate to his wife, illustrating how the deficit corrodes everyday relationships. {{c:Seminar}}",
        "Personality and social conduct are assessed with the **Frontal Lobe Personality Scale** (premorbid versus current personality), the Hare Psychopathy Checklist and the Faux Pas Test of theory of mind. {{c:Seminar}}",
        "The significance is that **one of the possible consequences of severe frontal derangement is social inadequacy**, so the patient may keep their intellect yet lose the capacity to live and work with others. {{c:Seminar}}",
    ]},
    {"t": "box", "kind": "key", "title": "Emotion and personality at a glance", "items": [
        "Orbitofrontal / ventromedial cortex modulates affect, impulse control and social conduct via amygdala and hypothalamus.",
        "Two frontal personalities: orbitofrontal disinhibited (pseudopsychopathic) vs dorsolateral/medial apathetic (pseudodepressed).",
        "Impaired empathy, theory of mind, autonoetic awareness and emotional regulation; social gaffes and poor judgement.",
        "Phineas Gage is the classic case; the somatic-marker idea explains lost emotional guidance of decisions.",
        "Tests: Frontal Lobe Personality Scale, Hare Psychopathy Checklist, Faux Pas Test; outcome often social inadequacy.",
    ]},
    {"t": "mn", "word": "GAGE", "text": "Gaffes & disinhibition (orbitofrontal), Apathy & inertia (dorsolateral/medial), Gut-feeling loss (somatic marker / ventromedial), Empathy & theory-of-mind failure \u2014 the frontal changes in emotion, personality and social conduct, anchored by the case of Phineas Gage."},
])


# ========================= 5. MEMORY =========================
_s("t5", "5", "Disturbances of memory", [
    {"t": "lead", "text": "Frontal memory disturbance is not a primary amnesia but a failure of the organisation, retrieval and monitoring of memory: working memory, prospective memory, temporal-order and source memory and metamemory are impaired, and confabulation can appear, yet recognition is relatively spared because storage itself is intact."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: frontal memory as a retrieval / organisation problem, not a storage amnesia.",
        "Background: the prefrontal role in working memory and the central executive.",
        "Key deficits: working, prospective, temporal-order / source memory, metamemory, confabulation.",
        "Evidence and examples: spared recognition, the Tulving encoding/retrieval asymmetry.",
        "Assessment and significance: the memory tests and the contrast with hippocampal amnesia.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "frontal_memory", "caption": "The frontal memory profile \u2014 working, prospective, temporal-order/source memory and metamemory, with confabulation \u2014 reflects a retrieval-and-monitoring failure, not a storage amnesia, which is why recognition is relatively spared."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{THEORY}} Frontal dysfunction has been characterised as a **dysexecutive syndrome** involving disruption of working memory, specifically a malfunction of the central executive that controls resources and monitors processing. {{c:Seminar}}",
        "{{GUIDE}} Frontal patients show **poor temporal memory**, especially poor working memory and poor delayed response, the deficits most associated with dorsolateral lesions. {{c:Kolb}}",
        "{{THEORY}} Tulving and colleagues proposed a **hemispheric asymmetry** in which the left prefrontal cortex is more engaged in encoding information and the right prefrontal cortex more in its retrieval. {{c:Kolb}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition: not a primary amnesia"},
    {"t": "ul", "items": [
        "Frontal memory disturbance is best defined as a **failure to organise, retrieve and monitor memory**, rather than the loss of the ability to store new information that marks hippocampal amnesia. {{c:Kolb}}",
        "The telling sign is that **recall is poor while recognition is relatively spared**, so the information has been stored but the patient cannot strategically retrieve it unaided. {{c:Lezak}}",
        "Frontal patients therefore **exhibit memory deficits on free recall, interference and working-memory tasks**, the tasks that make the heaviest demands on self-organised retrieval. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Working memory and prospective memory"},
    {"t": "ul", "items": [
        "[[Working memory|holding and manipulating information in mind for a few seconds]] is central to the frontal deficit, and its disruption is the core of Baddeley's dysexecutive account of frontal patients. {{c:Seminar}}",
        "Working-memory failure is shown on **spatial and feature delayed-response and delayed-alternation tasks**, where the patient must hold a position or a figure across a delay. {{c:Regions}}",
        "[[Prospective memory|remembering to carry out an intended action in the future]] is impaired, so the patient forgets to perform a planned activity at the right moment. {{c:Seminar}}",
        "Prospective-memory failure is especially disabling in daily life, because it undermines **keeping appointments, taking medication and following through on intentions**. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "Temporal-order, source memory and metamemory"},
    {"t": "ul", "items": [
        "[[Temporal-order memory|remembering the sequence in which events happened]] is impaired, so a patient may recall items yet misjudge their recency or order, a deficit tied to dorsolateral damage. {{c:Kolb}}",
        "[[Source memory|remembering where or how a piece of information was learned]] is disturbed, so the patient knows a fact but cannot say when or from whom they learned it. {{c:Lezak}}",
        "[[Metamemory|awareness and monitoring of one's own memory]] is poor, so the patient misjudges what they do and do not know, as in an unreliable feeling of knowing. {{c:Lezak}}",
        "There is also an impaired ability to **retrieve episodic (autobiographical) memories despite intact semantic retrieval**, so the patient claims to know the information but cannot consciously recollect the event. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Confabulation and evidence"},
    {"t": "ul", "items": [
        "[[Confabulation|producing false memories without any intent to deceive]] can appear, as the patient unwittingly fills memory gaps with plausible but untrue material because retrieval is not monitored or verified. {{c:Stuss}}",
        "Confabulation is most associated with **orbitomedial and ventromedial damage**, linking it to the same monitoring failure that underlies the social and personality changes. {{c:Stuss}}",
        "The relative sparing of recognition and of stored knowledge is the **key dissociation** that separates frontal memory disturbance from the storage amnesia of medial-temporal damage. {{c:Kolb}}",
        "The proposed **left-encode / right-retrieve asymmetry** remains debated, and it is given as a hypothesis rather than a settled fact, so it is noted with its controversy. {{c:Kolb}}",
    ]},
    {"t": "h4", "text": "Assessment and significance"},
    {"t": "ul", "items": [
        "Memory is assessed with the **Wechsler Memory Scale** and digit span, and with list-learning tests such as the California Verbal Learning Test and the Rey Auditory Verbal Learning Test. {{c:Seminar}}",
        "Everyday and prospective memory is probed by the **Rivermead Behavioural Memory Test**, which samples remembering an appointment, a name or a message. {{c:Seminar}}",
        "The frontal pattern of **poor recall but preserved recognition, with poor use of strategy**, is what distinguishes it on these tests from an amnesic storage deficit. {{c:Lezak}}",
        "The significance is that the patient's **knowledge is intact but inaccessible when needed**, so they fail real tasks that depend on self-cued retrieval even though they can recognise the right answer when shown it. {{c:Kolb}}",
    ]},
    {"t": "box", "kind": "key", "title": "Memory disturbance at a glance", "items": [
        "Frontal memory = a retrieval, organisation and monitoring failure, not a storage amnesia.",
        "Impaired working, prospective, temporal-order and source memory, and metamemory; episodic retrieval poor, semantic spared.",
        "Confabulation (unintended false memories) can appear, most after orbitomedial / ventromedial damage.",
        "Key dissociation: poor recall but relatively spared recognition \u2014 unlike hippocampal amnesia.",
        "Tests: WMS, digit span, California / Rey verbal learning, Rivermead; knowledge intact but inaccessible when needed.",
    ]},
    {"t": "mn", "word": "WPM-SC", "text": "Working memory, Prospective memory, Metamemory, Source/temporal-order memory, Confabulation \u2014 the frontal memory failures; remember the signature dissociation (poor RECALL, spared RECOGNITION) that proves it is a retrieval problem, not a storage amnesia."},
])


# ========================= 6. INTELLECTUAL ACTIVITY =========================
_s("t6", "6", "Disturbances of intellectual activity", [
    {"t": "lead", "text": "Frontal damage impairs the active side of intelligence \u2014 abstract reasoning, concept formation, problem-solving and cognitive flexibility \u2014 while often leaving measured IQ near normal, producing the paradox of a patient whose stored knowledge is intact yet who cannot apply it to a novel, open-ended problem in real life."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: intellectual activity and the frontal lobe's role in its active, fluid side.",
        "Background: the problem-solving loop and the prefrontal cortex.",
        "Key deficits: abstraction, concept formation, problem-solving, cognitive flexibility, decision-making.",
        "Evidence and examples: Wisconsin Card Sorting, Tower and maze tests, the IQ paradox.",
        "Assessment and significance: the reasoning tests and why near-normal IQ misleads.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "intellect_flow", "caption": "The problem-solving loop \u2014 abstract, form concepts, plan, act on feedback and shift set \u2014 is broken at every stage in frontal injury, so novel problem-solving collapses even though stored knowledge and measured IQ may be near normal."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{THEORY}} Loss of **abstract thought** was considered an early sign of frontal-lobe damage, so the patient falls back on a concrete attitude and cannot adopt the abstract one. {{c:Seminar}}",
        "{{GUIDE}} Frontal patients are **especially impaired at developing novel cognitive plans or strategies** for solving problems, and sequencing is perhaps the behaviour most clearly affected. {{c:Seminar}}",
        "{{GUIDE}} Despite scoring **normally on conventional tests of general cognitive ability**, patients may require constant care because they cannot apply their retained cognitive skills in everyday life. {{c:Seminar}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and the active side of intellect"},
    {"t": "ul", "items": [
        "Intellectual activity here means the **active, fluid use of reasoning** to solve new problems, as distinct from the crystallised store of facts and vocabulary that frontal damage largely spares. {{c:Fuster}}",
        "The prefrontal cortex supports **intelligence, creativity and higher-order reasoning**, so its damage strikes the generative side of thought rather than knowledge already laid down. {{c:Seminar}}",
        "This is why the deficit is best framed as a **loss of fluid rather than crystallised intelligence**, a distinction that explains the near-normal IQ scores. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "Abstraction and concept formation"},
    {"t": "ul", "items": [
        "**Abstract reasoning is impaired**, so proverbs are interpreted concretely and the patient cannot rise above the literal surface of a problem. {{c:Seminar}}",
        "[[Concept formation|grouping things by an underlying rule or category]] is poor, so the patient struggles to sort by an abstract principle and to shift that principle when it changes. {{c:Lezak}}",
        "**Rule deduction and rule following** break down, with the patient failing to identify the rule, to use feedback, or to switch set, which is the core of Wisconsin Card Sorting failure. {{c:Seminar}}",
        "Loss of the **abstract attitude** was treated historically as an early, sensitive marker of frontal involvement, underlining how central abstraction is to frontal intellect. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Problem-solving and flexibility"},
    {"t": "ul", "items": [
        "Frontal patients are **especially poor at generating novel strategies**, so open-ended problems with no ready-made solution expose the deficit most clearly. {{c:Seminar}}",
        "[[Cognitive flexibility|switching approach when the situation or feedback demands it]] is reduced, leaving the patient rigid and unable to abandon a failing strategy. {{c:Seminar}}",
        "**Decision-making is impaired**, with studies suggesting that frontal patients have great difficulty making correct decisions, particularly where outcomes are uncertain. {{c:Seminar}}",
        "**Sequencing and planning** of the steps of a solution are disrupted, so even when the patient knows the goal they cannot build the ordered path toward it. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Evidence, assessment and the IQ paradox"},
    {"t": "ul", "items": [
        "The **Wisconsin Card Sorting Test** is the classic demonstration, with frontal patients making more perseverative errors and failing to complete category sorts. {{c:Seminar}}",
        "**Tower tests (Hanoi, London, Toronto) and maze tests** (Porteus) show impaired planning and rule-adherence, as the patient cannot build a sequence of moves to a goal and ignores the rules. {{c:Seminar}}",
        "Abstraction itself is assessed with the **Category Test and Proverb Test**, which probe concept formation and abstract thinking. {{c:Seminar}}",
        "The central finding is the **IQ paradox**: patients can score normally on general-ability tests yet need round-the-clock care because they cannot apply those skills to real life. {{c:Seminar}}",
        "The significance is that **measured IQ can badly understate the disability**, so a frontal patient must be assessed with executive and real-life tasks, not a standard intelligence test alone. {{c:Lezak}}",
    ]},
    {"t": "box", "kind": "key", "title": "Intellectual disturbance at a glance", "items": [
        "Frontal intellect = the active, fluid use of reasoning on novel problems; crystallised knowledge is largely spared.",
        "Impaired abstraction (concrete proverbs), concept formation, rule deduction, decision-making and sequencing.",
        "Reduced cognitive flexibility \u2192 rigid, cannot drop a failing strategy (Wisconsin Card Sorting).",
        "Tests: Wisconsin Card Sorting, Tower of London / Hanoi, Porteus mazes, Category Test, Proverb Test.",
        "The IQ paradox: near-normal IQ yet unable to apply skills in daily life \u2014 so IQ understates the disability.",
    ]},
    {"t": "mn", "word": "ACPF", "text": "Abstraction, Concept formation, Problem-solving, Flexibility \u2014 the four pillars of frontal intellect that break down; remember the IQ PARADOX (normal score, real-life failure) and the Wisconsin Card Sort as the signature test."},
])


# ========================= 7. LANGUAGE FUNCTIONS =========================
_s("t7", "7", "Disturbances of language functions", [
    {"t": "lead", "text": "Frontal damage disturbs the output and regulatory sides of language rather than comprehension, giving non-fluent Broca-type speech, reduced verbal fluency, transcortical motor and dynamic aphasia, agrammatism and impaired discourse, and, at its deepest, the loss of the regulatory power of language over behaviour."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: the frontal language profile as a disturbance of output, not comprehension.",
        "Anatomy: Broca's area, the supplementary motor area and the left frontal cortex.",
        "Key deficits: Broca's aphasia, transcortical motor and dynamic aphasia, reduced fluency, discourse.",
        "Evidence and examples: agrammatism, sequence comprehension, the regulatory power of language.",
        "Assessment and significance: the language tests and the link to executive control.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "frontal_language", "caption": "Frontal language runs from non-fluent Broca's aphasia and transcortical motor aphasia through Luria's dynamic aphasia to reduced verbal fluency, and, at its deepest, the loss of the regulatory power of language over behaviour."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} **Broca's aphasia**, linked to lesions in Broca's area, is non-fluent \u2014 sparse, effortful, dysarthric, dysprosodic and agrammatic \u2014 with poor repetition but relatively good comprehension. {{c:Seminar}}",
        "{{GUIDE}} **Transcortical motor aphasia**, linked to left medial frontal (supplementary motor) lesions, shows an inability to initiate spontaneous speech, yet repetition and comprehension remain comparatively intact. {{c:Seminar}}",
        "{{THEORY}} Luria described a **loss of the regulatory power of language**, so frontal patients may verbalise a task correctly yet cannot use that verbalisation to direct their behaviour. {{c:Luria}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and anatomy"},
    {"t": "ul", "items": [
        "The frontal language profile is a disturbance of **speech output and the control of language**, not of the comprehension that depends on the posterior temporal cortex. {{c:Seminar}}",
        "**Broca's area** in the inferior frontal gyrus (around area 44) controls the motor production of voluntary speech and written expression. {{c:Seminar}}",
        "The **supplementary motor area and left medial frontal cortex** supply the drive to initiate speech, which is why their damage silences spontaneous output while sparing repetition. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "The frontal aphasias"},
    {"t": "ul", "items": [
        "[[Broca's aphasia|non-fluent, effortful, agrammatic speech with spared comprehension]] gives halting, telegraphic output with difficulty in word-finding, reading and writing but relatively preserved understanding. {{c:Seminar}}",
        "[[Transcortical motor aphasia|loss of spontaneous speech with repetition and comprehension intact]] leaves the patient unable to start speaking of their own accord even though they can still repeat what they hear. {{c:Seminar}}",
        "[[Dynamic aphasia|fluent words but an inability to build connected, propositional discourse]], described by Luria, spares single words yet blocks the construction of a connected narrative or proposition. {{c:Luria}}",
        "[[Aphemia|an apraxia of speech with intact comprehension and writing]] is an isolated breakdown of articulatory coordination, a complete articulatory failure with comprehension and writing preserved. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Fluency, grammar and discourse"},
    {"t": "ul", "items": [
        "**Reduced verbal fluency** is characteristic, so on letter (FAS) and category fluency tasks the patient names far fewer words than controls or patients with non-frontal lesions. {{c:Seminar}}",
        "[[Agrammatism|difficulty using verbs and correct grammar]] appears, with impoverished use of verbs and grammatical structure giving a telegraphic quality to speech. {{c:Seminar}}",
        "Because frontal patients cannot **maintain a verbal sequence**, their comprehension of long word-strings also suffers: they understand single words but grasp only part of a sequence, often the first or last items. {{c:Seminar}}",
        "At the discourse level, **pragmatics and connected narrative** are impoverished, so the patient's talk is thin, poorly organised and hard to sustain. {{c:Lezak}}",
    ]},
    {"t": "h4", "text": "The regulatory power of language, assessment and significance"},
    {"t": "ul", "items": [
        "The deepest frontal language sign is Luria's **loss of the regulatory power of language**: the patient can state a task aloud, yet the words no longer steer action, mirroring the knowing\u2013doing dissociation. {{c:Luria}}",
        "This loss ties language directly to **executive control**, because speech normally guides and plans behaviour, and frontal damage severs that guiding link. {{c:Fuster}}",
        "Language is assessed with the **Boston Naming Test** (word retrieval), the Token Test (verbal comprehension of sequences) and the Controlled Oral Word Association test of verbal fluency. {{c:Seminar}}",
        "Comprehensive aphasia batteries such as the **Western Aphasia Battery and the Boston Diagnostic Aphasia Battery** are used to classify the type and severity of the aphasia. {{c:Seminar}}",
        "The significance is that frontal language disturbance is primarily a **failure of production and self-regulation**, so the patient may understand yet cannot generate, organise or use speech to control what they do. {{c:Luria}}",
    ]},
    {"t": "box", "kind": "key", "title": "Language disturbance at a glance", "items": [
        "Frontal language = a disturbance of output and control, with comprehension largely spared.",
        "Broca's aphasia (non-fluent, agrammatic), transcortical motor aphasia (no spontaneous speech, repetition intact), Luria's dynamic aphasia.",
        "Reduced verbal fluency (FAS / category), agrammatism, and impaired comprehension of long word-sequences.",
        "Deepest sign: loss of the regulatory power of language \u2014 words no longer steer behaviour.",
        "Tests: Boston Naming, Token Test, Controlled Oral Word Association, Western / Boston aphasia batteries.",
    ]},
    {"t": "mn", "word": "BTD-FR", "text": "Broca's (non-fluent), Transcortical motor (no spontaneous speech), Dynamic (no connected discourse), reduced Fluency, and loss of the Regulatory power of language \u2014 the frontal language disturbances, all sparing comprehension while wrecking output and control."},
])


# ========================= 8. MOTOR FUNCTIONS =========================
_s("t8", "8", "Disturbances of motor functions", [
    {"t": "lead", "text": "The frontal lobe organises movement along a rostral-to-caudal hierarchy \u2014 prefrontal planning, premotor and supplementary-motor programming and selection, primary motor execution \u2014 so its damage gives contralateral weakness, loss of fine movement, apraxia, poor movement programming, akinesia, gait disturbance and the re-emergence of primitive release reflexes."},
    {"t": "box", "kind": "info", "title": "Answer flow \u2014 how to write this topic", "text": "Write the answer in this order, and the points below follow the same skeleton:", "items": [
        "Definition: the frontal lobe's role in the organisation and execution of movement.",
        "Anatomy: the primary motor, premotor / supplementary-motor and prefrontal levels.",
        "Key deficits: weakness, fine-movement loss, apraxia, poor programming, akinesia, perseveration.",
        "Evidence and examples: corollary discharge, the grasp reflex, gait and release signs.",
        "Assessment and significance: the motor tests and the link to the wider syndrome.",
        "One-line summary (use the mnemonic).",
    ]},
    {"t": "fig", "key": "motor_hierarchy", "caption": "The frontal motor hierarchy runs prefrontal (plan) \u2192 premotor/SMA (programme and select) \u2192 primary motor (execute); damage at each level gives a distinct deficit, and loss of frontal inhibition releases the primitive reflexes."},
    {"t": "box", "kind": "def", "title": "Characterisations by standard texts", "items": [
        "{{GUIDE}} Damage to the **primary motor cortex (M1, area 4)** causes contralateral motor deficits \u2014 initially a flaccid and later a spastic hemiplegia or hemiparesis. {{c:Regions}}",
        "{{GUIDE}} The motor cortex **executes** individual movements while the premotor cortex **selects** the movements to be executed, so the two contribute to movement in different ways. {{c:Kolb}}",
        "{{GUIDE}} Frontal patients may show **release signs** \u2014 reflexes such as the snout, suck, glabellar and grasp reflexes that are present in infancy but normally disappear with development. {{c:Regions}}",
    ]},
    {"t": "h3", "text": "Exam points"},
    {"t": "h4", "text": "Definition and anatomy"},
    {"t": "ul", "items": [
        "The frontal lobe organises movement along a **rostral-to-caudal hierarchy**, from prefrontal intention through premotor programming to primary-motor execution. {{c:Regions}}",
        "The **primary motor cortex** on the precentral gyrus projects via the corticospinal and corticobulbar tracts to drive voluntary limb, hand, digit and facial movement. {{c:Regions}}",
        "The **premotor and supplementary motor areas (area 6)** integrate sensory and motor information to select and programme actions, projecting onward to the primary motor cortex. {{c:Regions}}",
        "The **frontal eye fields (area 8)** control voluntary eye movements for visual search, so their damage impairs active scanning while sparing the passive following of a moving object. {{c:Regions}}",
    ]},
    {"t": "h4", "text": "Weakness, fine movement and apraxia"},
    {"t": "ul", "items": [
        "Damage to the primary motor cortex causes **contralateral weakness**, a flaccid hemiparesis at first that later becomes spastic, with the distribution depending on how much cortex is involved. {{c:Regions}}",
        "There is a chronic **loss of fine, independent movements** of the hands, fingers and face, together with a reduction in fine motor speed and strength. {{c:Seminar}}",
        "[[Apraxia|loss of skilled, learned movement that is not due to weakness]] can follow premotor damage, so the patient cannot carry out skilled actions they once performed even though they are not paralysed. {{c:Regions}}",
        "**Buccofacial apraxia** is seen as difficulty copying simple facial gestures, and frontal patients are specifically poor at copying sequences of facial movements. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Movement programming, corollary discharge and perseveration"},
    {"t": "ul", "items": [
        "**Movement programming is impaired**, so the patient has difficulty planning, preparing and selecting movements and serially ordering a complex chain of actions. {{c:Seminar}}",
        "Luria described a **disintegration of learned sequential actions** into a 'series of isolated fragments', so smooth skilled sequences fall apart into disconnected pieces. {{c:Seminar}}",
        "[[Corollary discharge|the brain's internal copy of a movement command, telling the rest of the brain a movement is coming]] is impaired, so a frontal lesion both disturbs movement and interferes with the signal that a movement is taking place. {{c:Seminar}}",
        "**Motor perseveration** appears, with compulsive repetition of an initiated action in movement, speech, drawing or writing, especially when the lesion extends to the basal ganglia. {{c:Seminar}}",
    ]},
    {"t": "h4", "text": "Akinesia, release signs and gait"},
    {"t": "ul", "items": [
        "[[Akinesia|poverty or absence of spontaneous, self-initiated movement]] and reduced spontaneous facial expression occur, reflecting the prefrontal and medial contribution to the initiation of movement. {{c:Kolb}}",
        "**Frontal release signs** re-emerge, including the grasp, snout, suck, palmomental and glabellar reflexes, as the loss of frontal inhibition frees these infantile reflexes. {{c:Regions}}",
        "The **grasp reflex** is a telling example, an involuntary grasp of an object placed in the hand that persists despite the patient's conscious attempts to let go. {{c:Seminar}}",
        "**Dystonic signs** such as gegenhalten \u2014 a paratonic resistance to passive movement \u2014 appear, distinct from the cogwheel rigidity of Parkinson's disease. {{c:Regions}}",
        "**Gait is disturbed**, classically a magnetic or apraxic gait in which the patient has difficulty lifting the feet from the ground when walking. {{c:Regions}}",
    ]},
    {"t": "h4", "text": "Assessment and significance"},
    {"t": "ul", "items": [
        "Motor strength and speed are assessed with **hand dynamometry (grip strength) and finger tapping**, and motor sequencing with the Facial Sequence Test and Luria's alternating-movement tasks. {{c:Seminar}}",
        "Visuomotor integration and apraxia are probed with the **Bender Visual Motor Gestalt Test** and simple gesture or object-use tasks, and voluntary gaze with a finger-gaze test. {{c:Seminar}}",
        "The significance is that these motor signs are **objective, examinable markers** of frontal dysfunction, anchoring an otherwise amorphous syndrome in observable neurological findings. {{c:Regions}}",
        "They also show the **continuity of the frontal lobe's work**, from executing a single movement to programming sequences and initiating action, so motor and executive deficits are two ends of one organising function. {{c:Kolb}}",
    ]},
    {"t": "box", "kind": "key", "title": "Motor disturbance at a glance", "items": [
        "Motor hierarchy: prefrontal (plan) \u2192 premotor / SMA (select & programme) \u2192 primary motor (execute).",
        "Primary-motor damage \u2192 contralateral weakness and loss of fine movement; premotor damage \u2192 apraxia and poor programming.",
        "Impaired corollary discharge and serial ordering; motor perseveration; akinesia and reduced facial expression.",
        "Frontal release signs (grasp, snout, suck, glabellar), gegenhalten, and a magnetic / apraxic gait.",
        "Tests: grip dynamometry, finger tapping, facial-sequence and alternating-movement tasks, reflex testing.",
    ]},
    {"t": "mn", "word": "WAP-GR", "text": "Weakness & fine-movement loss, Apraxia, Programming failure (+ perseveration), Gait disturbance & akinesia, Release signs (grasp, snout, suck) \u2014 the frontal motor disturbances along the plan\u2192programme\u2192execute hierarchy."},
])
