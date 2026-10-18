# -*- coding: utf-8 -*-
"""Single source of truth for Psychotherapy Unit I notes.

Both the HTML build (build.py -> parts) and the DOCX build (FEAT-002) import
SECTIONS from this module so the two deliverables carry identical content.

Inline markup tokens kept in text strings (renderers interpret them):
  {{LAW}} {{CODE}} {{GUIDE}} {{THEORY}} {{EG}} {{SEC}}  -> evidence labels
  {{c:TEXT}}                                            -> citation [TEXT]
  [[term|gloss]]                                        -> glossary term (gloss)
  {{fig:NAME}}                                          -> inline SVG figure (HTML only)

Block schema (dicts):
  {"t":"lead","text":str}
  {"t":"h3","text":str}
  {"t":"ul","items":[str | (str,[str,...]) ]}      # optional one level of sub-bullets
  {"t":"ol","items":[...]}                          # ordered, same item shape
  {"t":"table","caption":str|None,"head":[...],"rows":[[...],...]}
  {"t":"box","kind":"def|key|law|ex|warn|info","title":str,"items":[...] , "text":str|None}
  {"t":"fig","key":str,"caption":str}
  {"t":"mn","word":str,"text":str}                  # mnemonic
  {"t":"note","text":str}                           # small source/limitation note

TITLE and CREDIT are the byte-exact brand strings (shared by both deliverables).
"""

TITLE = "Psychotherapy \u2014 Unit I: Introduction to Psychotherapy"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

# Cover / title-page sub-heading used by both deliverables.
SUBTITLE = "M.Phil. Clinical Psychology \u2014 short-form study notes (bullet style)"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens both deliverables.
INTRO = [
    {"t": "lead", "text": "These are concise, bullet-style revision notes for Unit I of the Psychotherapy paper (M.Phil. Clinical Psychology). They cover all ten syllabus topics and fold in the relevant regulatory and professional bodies (RCI, IACP, WHO and APA) on the topics where each is examinable."},
    {"t": "h3", "text": "Evidence labels used throughout"},
    {"t": "table", "caption": "Each claim carries a label so that opinion is never presented as law.",
     "head": ["Label", "Meaning"],
     "rows": [
        ["{{LAW}}", "A statute or court order that binds by force of law."],
        ["{{CODE}}", "An enforceable professional or ethics code of a regulator or association."],
        ["{{GUIDE}}", "Aspirational or best-practice guidance; recommended, not mandatory."],
        ["{{THEORY}}", "A theoretical or author's position; schools differ."],
        ["{{EG}}", "An invented teaching example; names and details are fictional."],
        ["{{SEC}}", "Information available only through a secondary summary, not the primary document."],
     ]},
    {"t": "h3", "text": "How citations work"},
    {"t": "ul", "items": [
        "A citation in square brackets names the full book and its author, for example {{c:Wolberg}} or {{c:Corey}}, so you can trace any claim back to the work it comes from.",
        "Where a point rests on a statute or a professional body rather than a book, the citation names the document or body and its year, for example {{c:APA 2017}} or {{c:MHCA 2017, s.23}}, keeping any section number.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail.",
        "A term followed by text in rounded brackets is a plain-language gloss, for example alliance (the bond and agreement on goals and tasks), so hard clinical terms stay but are explained simply.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no laws, citations, clause numbers or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "{{SEC}} IACP (Indian Association of Clinical Psychologists): its members' code of conduct is represented only from secondary summaries, so it is described in general terms as broadly aligning with APA and RCI duties, and no IACP clause numbers are quoted.",
        "Primary texts relied on directly: the RCI Act 1992, the Mental Healthcare Act 2017, the POCSO Act 2012, the ICMR 2017 ethical guidelines, and the APA Ethics Code 2017 together with the APA Record Keeping and Telepsychology Guidelines.",
        "Statements about Indian law are for study; in practice read the current text and seek advice, as laws change.",
    ]},
]

# Short reading list keyed to the topics. Full source rows live in the HTML
# intro part; this is the compact version shared by both deliverables.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the books used and the primary legal, ethical and online sources consulted."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Lewis R. Wolberg \u2014 The Technique of Psychotherapy (definitions and aims of psychotherapy).",
        "Gerald Corey \u2014 Theory and Practice of Counseling and Psychotherapy (ethics and evidence-based practice).",
        "Judith S. Beck \u2014 Cognitive Behavior Therapy: Basics and Beyond (session structure and planning).",
        "Arthur E. Jongsma et al. \u2014 The Complete Adult Psychotherapy Treatment Planner (the six-step planning sequence).",
        "Glen O. Gabbard \u2014 Long-Term Psychodynamic Psychotherapy: A Basic Text (alliance, working through and supervision).",
        "Christine Lister-Ford \u2014 A Short Introduction to Psychotherapy (process, stages and evidence-based practice).",
        "Raymond J. Corsini & Danny Wedding \u2014 Current Psychotherapies (common factors and outcome research).",
        "Richard S. Sharf \u2014 Theories of Psychotherapy and Counseling (comparison of approaches).",
    ]},
    {"t": "h3", "text": "Primary legal, ethical and online sources"},
    {"t": "ul", "items": [
        "APA Ethics Code (2017); APA Dictionary of Psychology; APA EBPP policy (2006); APA Record Keeping Guidelines; APA Telepsychology Guidelines.",
        "WHO mhGAP (Mental Health Gap Action Programme) and WHO guidance on psychological interventions.",
        "Mental Healthcare Act 2017; POCSO Act 2012; ICMR 2017 ethical guidelines; Rehabilitation Council of India (RCI) Act 1992.",
        "Clinical rating-scale reference for the PHQ-9 and GAD-7 severity bands.",
        "{{SEC}} IACP code of conduct \u2014 secondary summaries only; primary text not accessed, so no clause numbers are quoted.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. DEFINITIONS =================================
_s("t1", "1", "Definitions of psychotherapy", [
    {"t": "lead", "text": "Psychotherapy = a planned, systematic, psychological (non-medical) treatment in which a trained professional uses a confiding relationship and theory-based methods to help a help-seeking client change feelings, thoughts and behaviour."},
    {"t": "fig", "key": "def_map", "caption": "Concept map: the five elements every modern definition of psychotherapy shares (trained person, relationship, approach, purpose, shared direction)."},
    {"t": "h3", "text": "Classic definitions"},
    {"t": "ul", "items": [
        "{{THEORY}} Wolberg (1954) defines psychotherapy as the treatment of problems of an emotional nature in which a trained person deliberately establishes a professional relationship to help a client change. {{c:Wolberg}}",
        "Wolberg's definition sets three explicit aims, namely to remove, modify or retard existing symptoms, to mediate [[disturbed patterns of behaviour|long-standing unhelpful ways of acting and relating]], and to promote positive personality growth.",
        "{{THEORY}} Jerome Frank (1961) describes therapy as a planned, [[emotionally charged|involving strong, genuine feeling]], confiding interaction between a trained, socially sanctioned healer and a sufferer that aims to change the sufferer's feelings and attitudes.",
        "{{THEORY}} Norcross (1990) calls it the informed and intentional application of clinical methods and [[interpersonal stances|the therapist's deliberate way of relating to the client]] drawn from established psychological principles to help people change.",
        "{{THEORY}} Strupp views psychotherapy as a [[learning process|change achieved through new experience and understanding, not through drugs]] in which a trained person uses the relationship to help the patient overcome distress and function better.",
        "Across these authorities the emphasis shifts but the core is constant, because each insists on a trained helper, a client in distress, a psychological method and a goal of desirable, agreed change.",
    ]},
    {"t": "h3", "text": "Definitions by professional bodies"},
    {"t": "ul", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises psychotherapy as any **psychological service** provided by a trained professional that uses **theory-based methods** and the **therapeutic relationship** to help a person reduce distress and change thoughts, feelings and behaviour. {{c:APA Dictionary of Psychology}}",
        "The APA framing deliberately keeps the term **broad**, so it spans many [[modalities|recognised forms or schools of therapy, such as psychodynamic, cognitive-behavioural or humanistic]] rather than tying psychotherapy to any single technique or school.",
        "{{GUIDE}} The **World Health Organization**, through its **mhGAP** programme, positions psychotherapy within the wider family of [[psychological interventions|structured, talking-based treatments delivered by trained providers]], recommending brief, structured talking treatments for common mental disorders. {{c:WHO mhGAP}}",
        "This is a **characterisation**, not a fixed verbatim WHO definition, because the WHO emphasises making **evidence-based** psychological care widely **accessible**, including through trained non-specialist providers in low-resource settings.",
        "{{GUIDE}} The **UK NICE guidance** reinforces the same view by recommending **specific psychological therapies**, such as cognitive-behavioural therapy, as **first-line** treatments for conditions like depression and anxiety disorders.",
        "Read together, these bodies agree that psychotherapy is a **professional, psychological and evidence-informed** treatment, differing mainly in whether they stress the **method**, its **accessibility** or its place in formal **clinical guidelines**.",
    ]},
    {"t": "box", "kind": "key", "title": "Elements common to all definitions", "items": [
        "A trained, professional therapist (socially sanctioned).",
        "A help-seeking client in distress.",
        "A psychological (not pharmacological / surgical) method.",
        "A planned, systematic, theory-based process.",
        "A confiding, therapeutic relationship (alliance).",
        "The aim of change: symptom relief, behaviour change and growth.",
    ]},
    {"t": "h3", "text": "Common elements in sentences (exam form)"},
    {"t": "ul", "items": [
        "Every definition requires a **trained, socially sanctioned therapist**, meaning a professional whose education and registration give them both the competence and the public authority to treat.",
        "Every definition assumes a **help-seeking client in distress**, because psychotherapy is a **voluntary** treatment entered into by someone who recognises a problem and wants help with it.",
        "The method is **psychological** rather than [[pharmacological or surgical|using medication or physical operations]], so change is pursued through **talking, relationship and new learning** instead of through the body.",
        "The process is **planned, systematic and theory-based**, which distinguishes professional therapy from ordinary **advice, friendship** or informal emotional support.",
        "A confiding **therapeutic relationship**, often called the [[alliance|the bond and agreement on goals and tasks between therapist and client]], is the **vehicle** through which techniques take effect.",
        "The shared aim is **change the client endorses**, whether that is symptom relief, altered behaviour or personal growth, so direction is **negotiated** rather than imposed.",
    ]},
    {"t": "h3", "text": "Wolberg's three aims"},
    {"t": "ul", "items": [
        "Wolberg's **first aim** is to **remove, modify or retard** existing symptoms, that is to reduce the client's current suffering such as anxiety, low mood or compulsive behaviour.",
        "His **second aim** is to **mediate**, meaning to correct, **disturbed patterns of behaviour** that keep the client stuck in unhelpful ways of acting and relating.",
        "His **third aim** is to promote **positive personality growth** and development, so that therapy builds **strengths and maturity** rather than only subtracting symptoms. {{c:Wolberg}}",
    ]},
    {"t": "h3", "text": "Psychotherapy vs counselling vs medical treatment"},
    {"t": "table", "caption": "Overlapping but distinguishable helping activities (difference is mostly of degree).",
     "head": ["Feature", "Counselling", "Psychotherapy", "Medical treatment"],
     "rows": [
        ["Focus", "Normal developmental, situational, adjustment problems", "Deeper, enduring or pathological problems; personality reconstruction", "Diagnosable illness via biological route"],
        ["Time frame", "Present- and problem-focused, shorter", "Often longer, may explore past", "Variable"],
        ["Method", "Guidance, support, problem solving", "Relationship + theory-based technique", "Drugs, procedures"],
        ["Overlap", "Large overlap with psychotherapy", "Often combined with medication", "Often combined with psychotherapy (e.g. depression, psychosis)"],
     ]},
    {"t": "ul", "items": [
        "**Counselling** usually deals with normal developmental, situational and [[adjustment problems|difficulty coping with an identifiable life change or stress]], is more **present- and problem-focused**, and tends to be **shorter** in duration.",
        "**Psychotherapy** addresses **deeper, more enduring** or pathological problems and [[personality reconstruction|reshaping long-standing traits and ways of relating]], usually runs **longer**, and may explore the client's past as well as the present.",
        "The counselling-psychotherapy boundary is one of **degree** rather than a sharp line, because the methods **overlap heavily** and many practitioners use the two terms almost interchangeably in practice.",
        "Psychotherapy works through **communication, relationship and learning** rather than drugs or surgery, which distinguishes it from the [[biological route|treatment by medication, physical procedures or other bodily means]] of **medical care**.",
        "Psychological and medical treatments are frequently **combined**, for example pairing antidepressant medication with CBT for depression or antipsychotics with psychosocial therapy for psychosis.",
        "{{THEORY}} Because the distinction is mostly of **degree**, examiners expect you to **name the overlap first**, then separate the activities by **focus, time frame and method** rather than by a rigid wall.",
    ]},
    {"t": "mn", "word": "TRAPS", "text": "Trained person, Relationship (confiding, professional), Approach (psychological, theory-based), Purpose (symptom relief and behaviour change), Shared direction (change the client endorses) \u2014 the five elements common to all definitions."},
    {"t": "box", "kind": "info", "title": "Exam pointer", "items": [
        "If asked to \u201cdefine psychotherapy\u201d, give one classic definition (Wolberg or Frank), then list the five common elements (TRAPS).",
        "If asked to \u201cdistinguish\u201d, use the table rows: focus, time frame, method and overlap.",
    ]},
])

# ========================= 2. OBJECTIVES ==================================
_s("t2", "2", "Objectives (goals) of psychotherapy", [
    {"t": "lead", "text": "Objectives run from immediate symptom relief to long-term growth; most therapies pursue several at once, weighted by the model and the client's needs."},
    {"t": "fig", "key": "aims_tree", "caption": "Hierarchy of aims: from the overall goal of less distress down through Wolberg's three aims to specific, measurable objectives."},
    {"t": "ul", "items": [
        "The first objective is to **relieve symptoms and distress**, such as anxiety, low mood, [[obsessions|unwanted, intrusive and repetitive thoughts]] or sleep disturbance, so that daily suffering is reduced.",
        "A second objective is to modify [[maladaptive|unhelpful and self-defeating]] **behaviour, thoughts and emotions**, replacing patterns that maintain the problem with more **adaptive** responses.",
        "Many therapies aim to help the client gain [[insight|an accurate understanding of the origins and meaning of one's difficulties]], which is especially central to **psychodynamic** work and supports **lasting change**.",
        "Therapy frequently seeks to improve **interpersonal relationships** and social functioning, because distress is often created and maintained within the client's **close relationships** and wider social roles.",
        "A further objective is to strengthen **coping and problem-solving** and [[adaptive defences|healthy ways of managing stress and difficult feelings]], while raising self-esteem and [[self-efficacy|the belief that one can succeed at a task]].",
        "Some objectives focus on resolving **inner conflicts and past trauma** and on fostering [[emotional regulation|the capacity to notice, tolerate and manage one's feelings]] rather than being overwhelmed by them.",
        "At the growth end, therapy promotes **personal development**, [[self-actualisation|realising one's fuller potential and becoming more fully oneself]] and a renewed sense of **meaning and purpose** in living.",
        "A long-term objective is to **prevent relapse** and improve quality of life, and in severe illness to **maintain functioning** in the community and prevent avoidable hospitalisation.",
        "Objectives are always **negotiated** with the client rather than imposed, so that the ends pursued are ones the client genuinely **values** and is motivated to work towards.",
        "Because objectives compete for limited time, the therapist **orders** them, placing **safety and symptom stabilisation** before the deeper work of insight, relationships and personal growth.",
    ]},
    {"t": "box", "kind": "info", "title": "Outcome vs process goals", "items": [
        "Outcome goals = the end the client wants (less panic, better sleep).",
        "Process goals = the in-therapy means (insight, alliance, skills practice).",
        "Both are negotiated and revisited as therapy proceeds.",
    ]},
    {"t": "h3", "text": "How objectives differ by model"},
    {"t": "ul", "items": [
        "In **psychodynamic** therapy the central objective is [[insight|understanding how past relationships and unconscious conflict shape present difficulties]], together with **working through** and more mature defences.",
        "In **cognitive-behavioural** therapy the objectives are **symptom reduction**, changing unhelpful thoughts and behaviours, learning new **coping skills**, and preventing relapse.",
        "In **humanistic** therapy the objectives are [[self-actualisation|realising one's fuller potential]], [[congruence|being genuine and consistent between inner experience and outward expression]], **self-acceptance** and autonomy.",
        "In **systemic or family** therapy the objectives shift from the individual to the **system**, aiming to change repeating **interaction patterns**, clarify boundaries and improve relationships.",
        "Because models **weight objectives differently**, the therapist chooses and **sequences** goals in line with the client's problem, preferences and the model used rather than applying one fixed list.",
    ]},
    {"t": "box", "kind": "ex", "title": "Illustrative example (fictional)", "text":
        "{{EG}} A client with panic names her outcome goal as travelling to work without a panic attack. The therapist adds process goals of building the alliance, practising interoceptive exposure and learning to tolerate anxiety, and together they order safety and symptom relief before deeper relationship work. Details invented."},
    {"t": "box", "kind": "info", "title": "Ordering objectives in practice", "items": [
        "Safety first: manage risk (suicidality, harm, abuse) before other goals.",
        "Then stabilise symptoms and distress, and build the alliance and coping skills.",
        "Then address underlying patterns, relationships and growth.",
        "Finally consolidate gains and plan relapse prevention and maintenance.",
    ]},
    {"t": "mn", "word": "RICHE-S", "text": "Relieve symptoms, Insight, Change maladaptive patterns, Healthy relationships, Enhance coping and self-esteem, Self-actualisation and growth \u2014 the ladder of objectives from relief to growth."},
    {"t": "note", "text": "{{GUIDE}} WHO frames mental-health care around functioning, participation and rights, not symptom removal alone \u2014 consistent with objectives such as social functioning and quality of life. {{c:WHO}}"},
])

# ========================= 3. PROCESS =====================================
_s("t3", "3", "The process of psychotherapy at a glance", [
    {"t": "lead", "text": "Therapy moves through recognisable stages; the loop of assess \u2192 formulate \u2192 plan \u2192 intervene \u2192 review repeats until goals are met."},
    {"t": "fig", "key": "planning_cycle", "caption": "The recurring clinical cycle of psychotherapy."},
    {"t": "ol", "items": [
        "Referral / help-seeking and first contact.",
        "Intake, assessment and risk screening.",
        "Case formulation (a working hypothesis of the problem).",
        "Contracting, goal-setting and treatment planning.",
        "Active intervention and working through.",
        "Monitoring progress with repeated measures.",
        "Consolidation, relapse prevention and termination.",
        "Follow-up / review.",
    ]},
    {"t": "box", "kind": "info", "title": "Reading the map", "items": [
        "The early steps (intake, formulation, contracting) belong to the beginning phase.",
        "Intervention and progress-monitoring repeat through the middle phase.",
        "Consolidation, relapse prevention and termination close the end phase; follow-up checks maintenance.",
        "The cycle is iterative: new information re-opens formulation and re-sets goals.",
    ]},
    {"t": "h3", "text": "The process in sentences (exam form)"},
    {"t": "ul", "items": [
        "The process usually begins with a **referral or self-referral**, in which a client, a doctor or a family member first makes contact and a brief screening decides whether therapy is appropriate and urgent.",
        "The **first contact** sets the tone, so the therapist explains the service, clarifies expectations and begins building trust even before any formal assessment work has started.",
        "**Engagement and rapport** come next, because a warm, respectful and collaborative stance helps the anxious or reluctant client feel safe enough to disclose and to commit to the work. {{c:Lister-Ford}}",
        "A structured **intake and assessment** then gathers the presenting complaint, history, mental-state findings and any test results, giving a rounded picture of the person and their difficulty.",
        "**Risk screening** is woven into assessment from the outset, so that suicidality, risk of harm to others and safeguarding concerns are identified early and managed before other work proceeds.",
        "The findings are drawn together into a **case formulation**, a working hypothesis that explains why the problem began and why it persists and that guides every later decision. {{c:Beck}}",
        "**Contracting** follows, agreeing the frame of time, place, fees, confidentiality limits and the respective responsibilities of client and therapist so expectations are explicit and shared.",
        "**Goal-setting and treatment planning** convert the formulation into specific, measurable aims and a chosen approach, a step set out in detail as a staged planning sequence. {{c:Jongsma}}",
        "A course of therapy is organised into a **beginning, middle and end phase**, each with its own tasks, so that the work has a recognisable shape rather than drifting without direction.",
        "The **beginning phase** concentrates on assessment, formulation, contracting and the forming of a strong [[working alliance|the collaborative bond and agreement on goals and tasks]] that predicts later outcome. {{c:Gabbard}}",
        "The **middle phase** is the heart of therapy, where **active intervention** applies the techniques of the chosen model to change thoughts, feelings, behaviour or relationships.",
        "**Working through** happens in this phase too, as the client repeatedly faces and resolves the same difficulty in different situations until the change becomes stable and durable. {{c:Gabbard}}",
        "**Homework and between-session work** extend therapy into daily life, so that skills practised in the room are rehearsed where the problems actually occur and gains generalise faster. {{c:Beck}}",
        "**Routine progress monitoring** with brief repeated measures tracks change session by session, letting therapist and client see objectively whether the plan is working or needs adjustment.",
        "When the bond is strained, the therapist attends to **alliance ruptures** and repairs them openly, treating the rupture itself as valuable therapeutic material rather than a failure. {{c:Gabbard}}",
        "**Review and plan revision** form an iterative loop, because new information or a plateau in progress re-opens the formulation and re-sets the goals rather than leaving the plan fixed.",
        "As goals are met the work turns to **consolidation and relapse prevention**, helping the client recognise early warning signs and rehearse a plan for managing future setbacks on their own.",
        "**Termination** is then planned rather than abrupt, with gains summarised, feelings about ending acknowledged, and sessions often spaced out so the client leaves feeling prepared and autonomous. {{c:Gabbard}}",
        "A **follow-up** contact some weeks or months later checks that improvements have been **maintained**, offering a booster session or re-referral if the client has begun to slip back.",
        "Across every stage the process stays **collaborative and ethical**, with consent revisited, records kept and decisions shared, so the client remains an active partner from first contact to follow-up.",
    ]},
    {"t": "mn", "word": "REAF-CIRT-F", "text": "Refer, Engage, Assess (with risk), Formulate, Contract and set goals, Intervene, Review, Terminate, Follow-up \u2014 the ordered stages of the therapy process from first contact to maintenance check."},
])

# ========================= 4. TRAINING ====================================
_s("t4", "4", "Issues related to the training of professional therapists", [
    {"t": "lead", "text": "Training defines who may practise, what competencies are required, and how the public is protected from unqualified practice."},
    {"t": "h3", "text": "Who may practise (India)"},
    {"t": "ul", "items": [
        "{{LAW}} The Rehabilitation Council of India (RCI) Act, 1992 (amended 2000) regulates training, registration and conduct of rehabilitation professionals, including clinical psychologists. {{c:RCI Act 1992}}",
        "{{LAW}} Only persons enrolled on the Central Rehabilitation Register (CRR) may practise; practising without registration is an offence. {{c:RCI Act 1992}}",
        "The RCI-recognised M.Phil. in Clinical Psychology (two years, with supervised clinical practice) is the entry qualification.",
        "{{LAW}} The Mental Healthcare Act (MHCA), 2017 recognises the clinical psychologist (RCI-recognised qualification) as a mental health professional. {{c:MHCA 2017}}",
        "{{SEC}} The IACP (Indian Association of Clinical Psychologists) is a professional association that advocates on training and professional identity; its code of conduct could not be primary-sourced here, so no clause numbers are quoted.",
    ]},
    {"t": "fig", "key": "training_pipeline", "caption": "From degree to registered independent practice."},
    {"t": "h3", "text": "Components of training"},
    {"t": "ul", "items": [
        ("Theoretical knowledge:", [
            "Psychopathology and classification; theories of psychotherapy; principles of assessment and psychometrics.",
        ]),
        ("Supervised clinical practice:", [
            "Direct observation, audio/video review, case conferences and live supervision under qualified supervisors.",
            "Graded responsibility, from observing to co-therapy to supervised independent work.",
        ]),
        "Personal therapy or self-awareness work (to manage countertransference and blind spots).",
        "Learning ethics and law and how to apply them in practice.",
        "Research literacy and the skills to practise evidence-based care.",
        "Cultural competence and working with diversity.",
        "Lifelong learning and continuing professional education.",
    ]},
    {"t": "h3", "text": "Competencies expected"},
    {"t": "ul", "items": [
        "Relationship (alliance) skills.",
        "Assessment and case formulation.",
        "Intervention skills in at least one major model.",
        "Ethical decision making and record keeping.",
        "Working with diversity; use of supervision; reflective practice.",
    ]},
    {"t": "box", "kind": "info", "title": "APA standards on competence & training", "items": [
        "{{CODE}} APA Ethics Code Std 2.01 \u2014 practise only within boundaries of competence. {{c:APA 2017, Std 2.01}}",
        "{{CODE}} Std 2.03 \u2014 maintain competence; Std 2.06 \u2014 manage personal problems that could impair work. {{c:APA 2017}}",
        "{{CODE}} Std 10.01(c) \u2014 when a trainee provides therapy, the client is told the trainee is supervised and given the supervisor's name. {{c:APA 2017, Std 10.01(c)}}",
    ]},
    {"t": "h3", "text": "Issues and challenges"},
    {"t": "ul", "items": [
        "There is wide **variation in quality, duration and supervision** of training across institutions, and a **shortage of trained supervisors** limits how many competent therapists can be produced each year.",
        "Indian training relies heavily on **Western models** that may not fit local realities, so there is a pressing need for **culturally adapted** and [[indigenous|developed from and suited to the local culture]] approaches.",
        "It is debated whether **personal therapy** should be **compulsory** for trainees, the argument being that it helps them manage [[countertransference|the therapist's own feelings stirred up by the client]] and their blind spots.",
        "Programmes must act as **gatekeepers**, judging suitability on **empathy, maturity and integrity**, and must be willing to remediate or, in the last resort, dismiss trainees who are unfit to practise.",
        "A **training-versus-licensing gap** allows untrained persons to practise, so protecting the public from [[quacks|unqualified practitioners posing as therapists]] through **registration** is an ongoing concern.",
        "Therapist **stress, burnout**, [[vicarious traumatisation|distress absorbed from repeatedly hearing clients' trauma]] and compassion fatigue are real risks that make **self-care** part of competent training.",
        "**Boundaries** and [[dual relationships|having more than one kind of relationship with the same person]] are especially tricky in **trainee-supervisor** pairings and in small communities where people already know one another.",
        "Training must now include **tele-therapy** and its particular ethics, since remote delivery raises fresh questions about **privacy, consent, jurisdiction** and crisis management.",
        "Throughout their careers therapists must practise within the **limits of their competence** and recognise when a case should be **referred** to a more suitable or more senior colleague.",
        "Staying **evidence-based** requires therapists to integrate **new research** into their continuing professional development rather than relying indefinitely on what they first learned.",
    ]},
    {"t": "mn", "word": "TRAIN", "text": "Theory, Relationship skills, Assessment, Intervention, iNtegrity/ethics \u2014 the core domains a programme must build."},
    {"t": "box", "kind": "info", "title": "WHO: task-sharing and scaling training", "text":
        "{{GUIDE}} The WHO mhGAP Intervention Guide and Problem Management Plus (PM+) support training non-specialist and lay providers to deliver brief psychological care in low-resource settings \u2014 relevant to the training-supply gap in India. {{c:WHO mhGAP}}"},
    {"t": "h3", "text": "Supervision and self-care"},
    {"t": "ul", "items": [
        "{{THEORY}} Supervision protects the client and develops the trainee: direct observation, audio/video review, case discussion and reflective feedback. {{c:Gabbard}}",
        "Supervisees must disclose difficulties honestly; supervisors hold responsibility for the client's care during training.",
        "Self-care (peer support, personal therapy, workload limits) guards against burnout, vicarious trauma and compassion fatigue.",
    ]},
    {"t": "note", "text": "{{GUIDE}} APA Telepsychology Guidelines ask therapists to be competent in the technology they use and to manage the extra confidentiality and jurisdiction issues of tele-therapy \u2014 now a training need. {{c:APA Telepsychology}}"},
])

# ========================= 5. ETHICS ======================================
_s("t5", "5", "Ethical issues in therapy work", [
    {"t": "lead", "text": "Ethics sets the higher, aspirational standard of conduct (law sets the enforceable minimum). The APA's five general principles are the most widely taught framework."},
    {"t": "h3", "text": "APA five general principles (2017)"},
    {"t": "fig", "key": "ethics_flow", "caption": "Ethical decision-making pathway."},
    {"t": "table", "caption": "APA (2017) General Principles A\u2013E. {{c:APA 2017}}",
     "head": ["Principle", "Meaning"],
     "rows": [
        ["A. Beneficence & non-maleficence", "Do good; avoid harm; weigh benefits against risks."],
        ["B. Fidelity & responsibility", "Trust, professional and scientific responsibility, accountability."],
        ["C. Integrity", "Honesty, accuracy, truthfulness; no deception or exaggerated claims."],
        ["D. Justice", "Fairness, equal access to quality services; no discrimination."],
        ["E. Respect for rights & dignity", "Privacy, confidentiality, autonomy, self-determination."],
     ]},
    {"t": "note", "text": "{{CODE}} These five principles are aspirational; APA's numbered Standards (e.g. 2.01, 3.04, 3.05, 3.10, 4.01\u20134.07) are enforceable for members. The RCI Code of Conduct reflects similar duties. {{c:APA 2017}}"},
    {"t": "h3", "text": "Major ethical issues"},
    {"t": "ul", "items": [
        "Competence requires the therapist to practise only within the limits of their training and experience, to refer problems that fall outside that scope, and to keep knowledge and skills current. {{c:APA 2017, Std 2.01}}",
        "Informed consent must be obtained before any assessment or treatment and treated as an ongoing process rather than a single signature, so the client agrees freely and understands what is proposed (see Section 8).",
        "Confidentiality means private information is shared only with the client's consent, and the therapist must explain its limits at the outset so trust is not later broken by surprise. {{c:APA 2017, Std 4.01-4.02}}",
        "Confidentiality has recognised exceptions, namely serious risk to the client or an identifiable other, suspected abuse of a child, and a lawful [[court order|a direction from a court that legally compels disclosure]].",
        "Avoiding harmful [[dual (multiple) relationships|having a second role with a client, such as friend, employer or lover, alongside being their therapist]] is essential, so social, business, financial and sexual ties with clients are prohibited. {{c:APA 2017, Std 3.05-3.08}}",
        "Sexual intimacy with a current client is absolutely forbidden, and with a former client it is prohibited for a specified period, because the power imbalance makes genuine consent doubtful and the risk of harm high.",
        "Boundary issues such as gifts, [[self-disclosure|the therapist revealing personal information]], physical touch and social-media contact must be managed thoughtfully so that the work stays focused on the client's needs.",
        "Fees and contracts must be clear, honest and agreed in advance, and the therapist must never exploit the client's dependence for financial, personal or other advantage.",
        "Record keeping and data security are ethical duties, because inaccurate, careless or insecurely stored records can harm the client and breach confidentiality.",
        "Termination and referral must be handled so that clients are not abandoned, which means giving notice, summarising progress and arranging a proper handover to continuing care.",
        "Cultural sensitivity requires the therapist to avoid bias linked to gender, caste, religion, sexuality, disability or socio-economic status and to adapt methods to the client's context.",
        "Seeking supervision and consultation when facing an ethical dilemma is itself an ethical duty, since a second professional view guards against blind spots and self-serving reasoning.",
        "Advertising and public claims must be truthful, so the therapist makes no false or [[unsubstantiated|not backed by evidence]] promises about cure rates or guaranteed outcomes.",
        "Research with clients carries its own duties of informed consent, avoidance of harm, confidentiality and prior approval by an ethics committee before the study begins.",
    ]},
    {"t": "box", "kind": "key", "title": "Ethical decision-making steps", "items": [
        "Identify the problem and the people affected.",
        "Consult codes and law.",
        "Consult supervisor / colleagues.",
        "Consider options and consequences.",
        "Decide and act.",
        "Document, then evaluate the outcome.",
    ]},
    {"t": "h3", "text": "Limits of confidentiality"},
    {"t": "fig", "key": "conf_tree", "caption": "Decision tree for confidentiality: screen for a legal duty, then serious risk to life or others, then consent, before deciding to disclose the minimum necessary or maintain confidentiality."},
    {"t": "box", "kind": "info", "title": "Confidentiality decision as steps (DOCX view)", "items": [
        "Step 1: If a legal duty applies (for example POCSO s.19, or a court or Board order), comply, share only the minimum needed, and document.",
        "Step 2: Else if there is a serious, real risk to life or of harm to another that disclosure could prevent, release only what is needed to prevent the harm.",
        "Step 3: Else seek the client's consent for any disclosure that would help them, such as involving family or a psychiatrist.",
        "Step 4: Otherwise maintain confidentiality, keep a safety plan, monitor, and review the situation in supervision.",
    ]},
    {"t": "ul", "items": [
        "{{LAW}} Tarasoff v. Regents of the University of California (1976) established a duty to warn and protect identifiable third parties from a client's serious threats of violence. {{c:Tarasoff 1976}}",
        "{{LAW}} MHCA 2017 protects confidentiality of information about a person's mental illness and treatment; disclosure is allowed with consent, for the treatment, for safety, or by court order. {{c:MHCA 2017, s.23}}",
        "{{LAW}} POCSO Act, 2012 makes reporting of known or suspected child sexual abuse mandatory. {{c:POCSO 2012, s.19}}",
        "{{LAW}} The right to privacy is a fundamental right (Puttaswamy v. Union of India, 2017). {{c:Puttaswamy 2017}}",
    ]},
    {"t": "box", "kind": "ex", "title": "Illustrative example (fictional)", "text":
        "{{EG}} A client states a specific, credible plan to seriously harm a named person. The therapist consults the supervisor, documents the reasoning, warns the identifiable third party and informs authorities, and limits disclosure to what safety requires. Names and details here are invented."},
    {"t": "note", "text": "{{SEC}} The IACP code of conduct for its members is reported to stress competence, responsibility, client welfare and integrity \u2014 broadly matching the APA principles above. The primary code could not be read here, so no clause numbers are given."},
    {"t": "h3", "text": "How ethics bodies map onto this topic"},
    {"t": "ul", "items": [
        "{{CODE}} APA Ethics Code (2017): five general principles plus enforceable Standards (competence 2.01, avoiding harm 3.04, multiple relationships 3.05, informed consent 3.10, confidentiality 4.01\u20134.07). {{c:APA 2017}}",
        "{{LAW}} RCI Act s.21 empowers the Council to prescribe standards of professional conduct and a code of ethics; breach can lead to removal from the register. {{c:RCI Act 1992, s.21}}",
        "{{SEC}} IACP: members' code of conduct (secondary summary only).",
        "{{GUIDE}} BPS (UK) states four principles: respect, competence, responsibility, integrity \u2014 a useful cross-check.",
    ]},
    {"t": "mn", "word": "FIBRE (A-E)", "text": "Fidelity and responsibility (B), Integrity (C), Beneficence and non-maleficence (A), Respect for rights and dignity (E), and justicE / Equity (D) \u2014 one word recalling the APA's five general principles A to E."},
])

# ========================= 6. LEGAL (INDIA) ================================
_s("t6", "6", "Legal issues in therapy work (India)", [
    {"t": "lead", "text": "Indian psychotherapy practice is governed by a statutory registration law, a rights-based mental-health law, child-protection law, disability law, consumer law and tele-practice guidelines."},
    {"t": "fig", "key": "authority_stack", "caption": "Hierarchy of authorities that bind a therapist, from binding statutes at the top down to personal values at the base; higher legal layers prevail where layers conflict."},
    {"t": "box", "kind": "info", "title": "Reading the authority stack (DOCX step view)", "items": [
        "Step 1: Binding law \u2014 RCI Act 1992, MHCA 2017, POCSO 2012 and court orders sit at the top and must always be obeyed.",
        "Step 2: Statutory regulations \u2014 the RCI's conduct and ethics regulations made under s.21 apply next.",
        "Step 3: Professional codes \u2014 APA Ethics Code, the IACP code and institutional rules bind members.",
        "Step 4: Guidelines \u2014 aspirational best practice such as APA record-keeping and telepsychology guidance.",
        "Step 5: Personal values and clinical judgement operate only inside all the layers above; where they conflict, the higher legal layer prevails.",
    ]},
    {"t": "ul", "items": [
        "{{LAW}} The RCI Act 1992 regulates the training, registration and conduct of rehabilitation professionals including clinical psychologists, and practising without registration on the Central Rehabilitation Register is a punishable offence. {{c:RCI Act 1992}}",
        "{{LAW}} The Mental Healthcare Act 2017 is a rights-based law that reframes mental-health care around the dignity, autonomy and legal entitlements of the person rather than around institutional control. {{c:MHCA 2017}}",
        "Under the MHCA every person with mental illness has rights to affordable quality care, community living, confidentiality, access to their own records, equality and freedom from cruel or [[degrading treatment|treatment that humiliates or demeans the person]].",
        "The MHCA lets a person make an [[advance directive|a statement, made while well, of how one wishes to be treated during a future crisis]] and name a [[nominated representative|a trusted person chosen to take treatment decisions on one's behalf]].",
        "The Act presumes that every adult has the capacity to make their own treatment decisions unless the contrary is properly established, protecting clients from being overridden without good reason.",
        "{{LAW}} Section 115 presumes that a person who attempts suicide is under severe stress and is not to be punished, and it places a duty on government to provide care and rehabilitation. {{c:MHCA 2017, s.115}}",
        "The MHCA restricts [[ECT|electroconvulsive therapy, a medical treatment passing a brief electric current through the brain]], barring it without anaesthesia and muscle relaxant and barring it for minors, and limits the use of restraints and sterilisation.",
        "{{LAW}} The Rights of Persons with Disabilities Act 2016 covers mental illness and intellectual disability and underpins non-discrimination, reasonable accommodation and access to services.",
        "{{LAW}} The POCSO Act 2012 makes reporting of known or suspected child sexual abuse mandatory, and the Juvenile Justice Act governs the care and protection of children. {{c:POCSO 2012, s.19}}",
        "{{LAW}} The Consumer Protection Act and the civil and criminal law on [[negligence|failing to take the care a reasonably competent professional would take, causing harm]] expose therapists to malpractice claims, and courts may require records or testimony.",
        "{{GUIDE}} For tele-mental health the Telemedicine Practice Guidelines 2020 and guidance on tele-psychiatry apply, adding duties around identity verification, consent and secure technology.",
        "{{LAW}} The right to privacy is a fundamental right after Puttaswamy (2017), which strengthens the legal basis for confidentiality and careful handling of client data. {{c:Puttaswamy 2017}}",
    ]},
    {"t": "box", "kind": "key", "title": "Legal vs ethical", "items": [
        "Law sets the minimum enforceable standard; breach can mean penalty or loss of registration.",
        "Ethics sets the higher aspirational standard of conduct.",
        "Therapist's legal duties: valid consent, confidentiality, competent practice, accurate records, reporting obligations, no exploitation, referral when needed.",
    ]},
    {"t": "ul", "items": [
        "Law and ethics overlap but are not identical, because the law states the enforceable floor below which conduct becomes punishable while ethics describes the higher standard a good therapist aspires to.",
        "Where a legal duty and an ethical value appear to clash, the therapist clarifies the true legal obligation, seeks consultation, and documents the reasoning rather than acting on assumption.",
        "A therapist's core legal duties include obtaining valid consent, protecting confidentiality within its limits, practising competently, keeping accurate records and meeting mandatory-reporting obligations.",
    ]},
    {"t": "mn", "word": "RM-PCDT", "text": "RCI Act, MHCA 2017, POCSO 2012, Consumer/negligence law, Disability (RPwD 2016) Act, Telemedicine guidelines \u2014 the chain of Indian legal sources a therapist must respect."},
])

# ========================= 7. RIGHTS & RESPONSIBILITIES ====================
_s("t7", "7", "Rights and responsibilities in psychotherapy", [
    {"t": "lead", "text": "Therapy is a contract: both parties hold rights and duties. Clarifying them at the start supports the alliance and informed consent."},
    {"t": "fig", "key": "rights_resp", "caption": "Each client right (PRIDE) maps to a matching therapist duty: Privacy, Respect, Information, Decide (consent), Exit/Escalate."},
    {"t": "box", "kind": "info", "title": "Rights map to duties (DOCX step view)", "items": [
        "Privacy and confidentiality (client right) -> therapist keeps information secure and discloses only the minimum necessary.",
        "Respect and non-discrimination -> therapist practises in a culturally competent, non-exploitative way with no boundary violations.",
        "Information about treatment -> therapist describes approach, risks, benefits, alternatives, fees and qualifications.",
        "Decide (consent, refuse, withdraw) -> therapist obtains valid consent and assent and never coerces the client.",
        "Exit and Escalate (end therapy, complain) -> therapist terminates ethically, refers, and tells clients how to raise concerns.",
    ]},
    {"t": "mn", "word": "PRIDE", "text": "Privacy/confidentiality, Respect and dignity, Information about treatment, Decide (consent, refuse, withdraw), Exit and Escalate (end therapy, complain) \u2014 the client's core rights, each mirrored by a therapist duty."},
    {"t": "table", "caption": "Client rights and client responsibilities.",
     "head": ["Client rights", "Client responsibilities"],
     "rows": [
        ["Informed consent; know the therapist's training, approach and fees", "Attend sessions regularly and on time; give notice of cancellation"],
        ["Confidentiality (with known limits)", "Be honest and open about problems, history and progress"],
        ["Be treated with respect and dignity, free from discrimination", "Actively participate and complete agreed homework"],
        ["Know the nature, expected duration and alternatives of treatment", "Pay agreed fees; follow the agreed contract"],
        ["Refuse or withdraw from therapy at any time", "Respect the therapist's boundaries"],
        ["Access own records; a second opinion; make complaints", "Inform the therapist about other treatments, medicines, risks"],
     ]},
    {"t": "table", "caption": "Therapist responsibilities.",
     "head": ["Therapist responsibilities"],
     "rows": [
        ["Competence; keep within limits of training; refer when appropriate"],
        ["Obtain informed consent; maintain confidentiality"],
        ["Maintain boundaries; avoid dual and exploitative relationships"],
        ["Keep accurate records; protect them; follow reporting laws"],
        ["Respect diversity; use culturally appropriate methods"],
        ["Continue supervision and learning; practise self-care"],
        ["Terminate appropriately; ensure continuity of care"],
     ]},
    {"t": "h3", "text": "Expanded in sentences (exam form)"},
    {"t": "ul", "items": [
        "Psychotherapy is best understood as a [[therapeutic contract|an explicit mutual agreement about aims, methods, fees and boundaries]] in which both parties **knowingly accept rights** and matching responsibilities.",
        "The client has a right to **informed consent**, which means knowing the therapist's **training, approach**, likely duration, fees and the alternatives before agreeing to proceed.",
        "The client has a right to **confidentiality** subject to its **known limits**, so disclosures made in therapy are protected except where law or serious risk requires a narrow breach.",
        "The client has a right to be treated with **respect and dignity** and to be **free from discrimination** based on gender, caste, religion, disability, sexuality or socio-economic status.",
        "The client retains the right to **refuse or withdraw** from therapy at any time without penalty, which keeps participation genuinely **voluntary** throughout the work.",
        "The client has a right to **access their own records**, to seek a **second opinion** and to make a complaint through a recognised grievance route if dissatisfied.",
        "In return the client is responsible for **attending regularly**, giving notice of cancellation, being **honest** about history and progress, and completing agreed [[homework|tasks practised between sessions to extend therapy into daily life]].",
        "The therapist is responsible for working only within their **competence**, obtaining **valid consent**, protecting confidentiality and maintaining clear **professional boundaries** at all times.",
        "The therapist is further responsible for keeping **accurate records**, respecting diversity, using **supervision**, practising self-care, and ensuring **continuity of care** when therapy ends.",
        "The therapist also holds **rights**, including the right to set a **safe working frame**, to be free from abuse or harassment, and to decline work that cannot be conducted ethically.",
        "The therapist has a right to **fair remuneration** under the agreed contract and to expect the client to respect the agreed **time, fee and boundary** arrangements of the work.",
        "Rights and responsibilities are **reciprocal**, so each client right is **mirrored** by a therapist duty and each therapist right is matched by a client responsibility, keeping the relationship balanced.",
        "Clarifying rights and duties in the **first session** operationalises the [[informed consent|voluntary agreement after understanding adequate information]] process and strengthens the [[therapeutic alliance|the collaborative, trusting bond between client and therapist]].",
        "{{LAW}} Many of these client rights are not merely ethical but **statutory**, because the **Mental Healthcare Act 2017** gives legal force to information, confidentiality, access to records and complaint rights.",
        "{{GUIDE}} The **rights-based** framing reflects **WHO** guidance, which urges that mental-health care protect **dignity, autonomy and participation** rather than focus on control or symptom removal alone.",
        "When rights and responsibilities are made **explicit** and revisited, **boundary problems**, missed-session disputes and premature drop-out all become markedly less likely over the course of therapy.",
    ]},
    {"t": "note", "text": "{{LAW}} Many client rights are also statutory under MHCA 2017 ss.18\u201325 (information, confidentiality, access to records, complaints). {{c:MHCA 2017}}"},
    {"t": "box", "kind": "key", "title": "Why this matters", "items": [
        "Spelling out rights and duties at the first session operationalises informed consent and the therapeutic contract.",
        "{{GUIDE}} It also reflects a rights-based model of care consistent with the MHCA 2017 and WHO guidance on dignity and participation.",
        "Clear mutual expectations reduce boundary problems, missed-session disputes and premature drop-out.",
    ]},
    {"t": "h3", "text": "Therapist rights too"},
    {"t": "ul", "items": [
        "To set and maintain professional boundaries and a safe working frame.",
        "To be treated with respect and to work free from abuse, threats or harassment.",
        "To decline or terminate work that is outside competence or that cannot be conducted ethically, with appropriate referral.",
        "To fair remuneration under the agreed contract.",
    ]},
])

# ========================= 8. CONSENT & ASSENT =============================
_s("t8", "8", "Consent and assent (including minors)", [
    {"t": "lead", "text": "Informed consent is the client's voluntary agreement to assessment or therapy after receiving and understanding adequate information. It is a process, not a one-time signature."},
    {"t": "h3", "text": "Elements of valid consent"},
    {"t": "ul", "items": [
        "Capacity (competence): ability to understand, appreciate, reason and express a choice.",
        ("Disclosure of information:", [
            "Purpose and nature of therapy, methods, expected benefits.",
            "Possible risks and discomfort; alternatives (including no treatment).",
            "Duration, fees, confidentiality and its limits.",
            "Recording or observation, supervision, therapist's qualifications.",
            "The right to withdraw without penalty.",
        ]),
        "Comprehension: simple, understandable language; check understanding.",
        "Voluntariness: free from coercion or undue pressure.",
        "Documentation: written (or recorded) consent, revisited when the plan changes.",
    ]},
    {"t": "note", "text": "{{CODE}} APA Std 3.10 (informed consent) and 10.01 (consent to therapy) set out these duties; Std 3.10(b) lists four duties when a person cannot consent. {{c:APA 2017, Std 3.10 & 10.01}}"},
    {"t": "h3", "text": "Elements in sentences (exam form)"},
    {"t": "ul", "items": [
        "Valid consent first requires [[capacity|the mental ability to understand information, weigh it and communicate a choice]], which is **presumed in adults** unless a specific impairment is shown.",
        "The therapist must **disclose adequate information**, covering the **nature and purpose** of therapy, its expected benefits, foreseeable risks and the realistic **alternatives** including no treatment.",
        "Disclosure must also state the likely **duration**, the **fees**, the **limits of confidentiality**, any recording or supervision, and the client's right to **withdraw** without penalty.",
        "**Comprehension** matters as much as disclosure, so the therapist uses **plain language**, avoids jargon, and **checks understanding** rather than assuming the client has grasped the information.",
        "Consent must be **voluntary**, meaning it is given **freely** and without **coercion**, undue pressure or inducement from the therapist, family members or an institution.",
        "Consent is an **ongoing process** rather than a **one-time signature**, so it is **revisited** whenever the treatment plan, the risks or the therapeutic methods change materially.",
        "Consent should normally be **documented** in writing, and the record should note what was explained, what the client **agreed to**, and that their questions were answered.",
        "For a minor, **legal consent** is given by the **parent or guardian** while the child's own [[assent|the child's affirmative agreement, sought to the extent of their developing understanding]] is sought alongside it.",
        "Seeking **assent** respects the child as a **developing person**, so the therapist explains therapy in **age-appropriate words**, invites questions and gives genuine weight to a mature adolescent's views.",
        "{{GUIDE}} **ICMR 2017** guidance generally expects **assent** from children of roughly **seven to eighteen** years, using simpler language for younger children and more weight for mature adolescents. {{c:ICMR 2017}}",
        "A clear [[confidentiality contract|an up-front agreement on what stays private and what parents are told]] is agreed with the minor and guardian, naming the **safety situations** in which privacy is overridden.",
        "When a guardian and a minor **disagree**, or the minor declines, the therapist **explores the reasons** without coercion, consults a supervisor, and considers the appropriate legal route before proceeding.",
        "{{CODE}} Where a person **cannot give consent**, APA Standard 3.10(b) still requires the therapist to give an **explanation**, **seek assent**, consider preferences and obtain **substitute authorisation**. {{c:APA 2017, Std 3.10}}",
        "{{LAW}} If a child discloses **sexual abuse** during therapy, the **POCSO Act 2012** makes reporting **mandatory**, so the duty to report **overrides** the ordinary promise of confidentiality. {{c:POCSO 2012, s.19}}",
        "Consent for clients with **impaired capacity**, such as severe psychosis or intellectual disability, combines **substitute decision-making** with the person's own **assent** wherever possible.",
    ]},
    {"t": "h3", "text": "Special issues"},
    {"t": "fig", "key": "consent_flow", "caption": "Consent-and-assent flow for a minor: check age/capacity, take guardian consent, seek the minor's assent, agree confidentiality limits, handle disagreement and mandatory reporting."},
    {"t": "box", "kind": "info", "title": "Consent-assent flow as steps (DOCX view)", "items": [
        "Step 1: Check whether the client is under 18 (a minor under MHCA and a child under POCSO); if an adult, use the normal adult-capacity consent process.",
        "Step 2: Explain aims, methods, risks, alternatives, fees and confidentiality limits to the parent or legal guardian and obtain informed consent.",
        "Step 3: Explain therapy to the minor in age-appropriate words, invite questions and seek assent, weighing the minor's preferences and best interests.",
        "Step 4: Agree a confidentiality contract with both about what stays private, what parents are told, and when safety overrides privacy.",
        "Step 5: If guardian and minor disagree or the minor declines, explore reasons without coercion, consult a supervisor, and consider the legal route.",
        "Step 6: Begin therapy, documenting both consent and assent; if abuse is disclosed, follow the POCSO s.19 reporting duty.",
    ]},
    {"t": "mn", "word": "CARVED", "text": "Capacity, Adequate disclosure, Reasoning/comprehension, Voluntariness, Enduring process (not one-off), Documentation \u2014 the elements that make consent valid; for minors add Assent on top of guardian consent."},
    {"t": "ul", "items": [
        ("Minors (<18 years in India):", [
            "Legal consent comes from the parent or legal guardian.",
            "The child's assent (affirmative agreement to the extent of developing capacity) is also sought.",
            "{{GUIDE}} ICMR 2017 guidelines generally require assent for children roughly 7\u201318 years, with simpler language for younger children and more weight to mature adolescents. {{c:ICMR 2017}}",
            "If a child refuses, explore concerns and, unless safety requires otherwise, respect it.",
        ]),
        ("Adults lacking capacity:", [
            "{{LAW}} Involve the nominated representative or legal guardian under MHCA 2017; still seek the person's views and use the least restrictive approach. {{c:MHCA 2017}}",
        ]),
        "Consent for recording, observation, supervision, research or publication is separate and specific; case material is anonymised.",
        "Emergency (immediate risk to life) can justify action without full consent, with documentation.",
        "Cultural context: in collectivist settings family involvement is common; still ensure the client's own consent and confidentiality.",
    ]},
    {"t": "box", "kind": "ex", "title": "Illustrative example (fictional)", "text":
        "{{EG}} A 14-year-old is brought by a parent. The therapist takes the parent's legal consent, then explains therapy to the adolescent in simple words and obtains assent, agreeing what will and will not be shared with the parent. All details invented."},
    {"t": "box", "kind": "key", "title": "Consent vs assent", "items": [
        "Consent = the legally valid permission (the competent adult, or a minor's parent/guardian).",
        "Assent = the child's own affirmative agreement, sought to the extent of their developing capacity.",
        "A minor's refusal (dissent) is explored and, unless safety requires otherwise, respected.",
    ]},
    {"t": "note", "text": "{{LAW}} POCSO 2012 defines a \u201cchild\u201d as a person below 18 and makes reporting of known or suspected child sexual abuse mandatory; this can override a minor's or family's wish for confidentiality. {{c:POCSO 2012, s.19}}"},
])

# ========================= 9. PLANNING & RECORDING =========================
_s("t9", "9", "Planning and recording of therapy sessions", [
    {"t": "lead", "text": "Good planning turns assessment into a formulation and a concrete treatment plan; good records support continuity, supervision and accountability."},
    {"t": "h3", "text": "Planning"},
    {"t": "ul", "items": [
        "Intake and assessment: presenting complaints, history, risk assessment, mental status, test results and diagnosis.",
        ("Case formulation \u2014 an integrated hypothesis that explains the problem and guides the plan:", [
            "Predisposing factors (vulnerabilities), precipitating factors (triggers).",
            "Perpetuating factors (what maintains the problem), protective factors (strengths and resources).",
        ]),
        ("Treatment plan:", [
            "Goals and the chosen therapy; frequency (usually weekly) and session length (usually 45\u201350 minutes).",
            "Expected duration, measures of progress, a plan for emergencies, and criteria for termination.",
        ]),
        "Therapeutic frame: time, place, fee, boundaries, cancellation policy and contact between sessions.",
        "Session planning: prepare using previous notes; plan the aims for each session but stay responsive to the client's needs.",
    ]},
    {"t": "note", "text": "{{THEORY}} Jongsma's Treatment Planner sets out a 6-step planning sequence (problem selection \u2192 definition \u2192 goals \u2192 objectives \u2192 interventions \u2192 diagnosis). {{c:Jongsma}}"},
    {"t": "h3", "text": "Planning in sentences (exam form)"},
    {"t": "ul", "items": [
        "Planning begins with a thorough **intake and assessment** in which presenting complaints, history, risk, mental state and any test results are gathered to **understand the problem** fully.",
        "The assessment is then organised into a [[case formulation|an integrated working hypothesis explaining why the problem began and why it persists]], which links the findings into a **coherent story**.",
        "A useful formulation weighs the [[four Ps|predisposing, precipitating, perpetuating and protective factors]], distinguishing **vulnerabilities and triggers** from what **maintains** the problem and what **protects** the client.",
        "From the formulation flows a **treatment plan** stating the **goals**, the chosen therapy, the frequency and length of sessions, the measures of progress and the **criteria for ending** therapy.",
        "The plan also sets the **therapeutic frame** of time, place, fee and boundaries, and it includes a **plan for emergencies** so that risk is managed rather than improvised.",
    ]},
    {"t": "box", "kind": "ex", "title": "Illustrative example (fictional)", "text":
        "{{EG}} For a client with low mood, the formulation notes a family history (predisposing), a recent job loss (precipitating), social withdrawal (perpetuating) and a supportive partner (protective); the plan is weekly behavioural activation for twelve sessions with PHQ-9 monitoring. Details invented."},
    {"t": "h3", "text": "Structure of a typical CBT session"},
    {"t": "ol", "items": [
        "Brief check-in and mood check.",
        "Bridge from the previous session.",
        "Collaborative agenda setting.",
        "Review of homework.",
        "Work on agenda items (main therapeutic work).",
        "Capsule summaries during the session.",
        "Assign new homework.",
        "Client feedback on the session.",
    ]},
    {"t": "note", "text": "{{THEORY}} J. Beck describes this session structure. {{c:Beck}}"},
    {"t": "h3", "text": "Recording"},
    {"t": "ul", "items": [
        ("Types of records:", [
            "Intake / case-history form; assessment reports and test data; the treatment plan; consent forms.",
            "Progress notes for each session; process notes (the therapist's private reflections, kept separate).",
            "Risk assessments; referral and termination summaries; correspondence.",
        ]),
        "Formats of progress notes: SOAP (Subjective, Objective, Assessment, Plan), DAP (Data, Assessment, Plan), BIRP (Behaviour, Intervention, Response, Plan).",
        "Good records are timely, accurate, objective, factual, brief but complete, free of judgemental language, and signed and dated.",
        "Purposes: continuity of care, progress monitoring, supervision, communication with other professionals, legal/ethical evidence, research and quality improvement.",
        ("Security and confidentiality:", [
            "Secure storage (locked cabinets, password-protected files); restricted access; coded identifiers.",
            "Retention for the period required by law or institutional policy; safe disposal afterwards.",
        ]),
        "Audio/video recording needs explicit written consent; recordings are used only for the agreed purpose (such as supervision) and destroyed as agreed.",
    ]},
    {"t": "h3", "text": "Recording in sentences (exam form)"},
    {"t": "ul", "items": [
        "Clinical records exist to support **continuity of care**, so that another professional, or the same therapist after a gap, can understand what happened and continue the work safely.",
        "Records serve **accountability**, because accurate, **dated and signed** notes provide evidence of what was done and why if the work is later questioned clinically, ethically or legally.",
        "Good [[progress notes|the brief factual record written after each session]] are **timely, accurate, objective** and free of judgemental language, recording what was said, observed, judged and planned.",
        "Several **structured formats** exist, including **SOAP, DAP** and [[BIRP|Behaviour, Intervention, Response, Plan]], each giving a consistent skeleton so that important information is not omitted.",
        "[[Process notes|the therapist's private reflections, hunches and reactions]] are kept **separate** from the formal clinical record and generally receive **greater protection** from disclosure.",
        "Records must be **stored securely**, for example in locked cabinets or password-protected files with **restricted access**, because a confidentiality breach can seriously harm the client.",
        "Records are **retained** for the period required by law or policy and then **disposed of safely**, so that out-of-date files do not become a lingering privacy risk.",
        "{{LAW}} The client generally has a right to **access their own records**, a right made **statutory** in India by the **Mental Healthcare Act 2017** s.25. {{c:MHCA 2017, s.25}}",
    ]},
    {"t": "fig", "key": "soap_dap", "caption": "SOAP (four parts) versus DAP (three parts): how the Subjective and Objective of SOAP combine into the single Data section of DAP."},
    {"t": "box", "kind": "info", "title": "APA guidance on records", "items": [
        "{{CODE}} APA Std 6.01\u20136.02 \u2014 create, maintain and dispose of records to protect confidentiality. {{c:APA 2017}}",
        "{{GUIDE}} APA Record Keeping Guidelines (2007) \u2014 content, retention and security best practice. {{c:APA Record Keeping}}",
        "{{LAW}} MHCA 2017 s.25 gives the person a right to access their own records. {{c:MHCA 2017, s.25}}",
    ]},
    {"t": "box", "kind": "ex", "title": "Sample SOAP note (fictional)", "text":
        "{{EG}} S: \u201cFelt panicky on the bus twice.\u201d O: Engaged, completed thought record. A: Reduced avoidance; PHQ-9 = 9 (down from 14). P: Continue interoceptive exposure; homework = two bus trips. Illustrative only."},
    {"t": "mn", "word": "SOAP", "text": "Subjective (what the client reports), Objective (what you observe/measure), Assessment (your clinical judgement), Plan (next steps)."},
    {"t": "note", "text": "Process notes (the therapist's private reflections) are kept separate from the formal record and are generally given more protection; progress notes form the official clinical record."},
])

# ========================= 10. STRUCTURE & GOALS ===========================
_s("t10", "10", "Structuring therapy and setting goals", [
    {"t": "lead", "text": "Structure gives therapy direction, safety and efficiency; collaborative goals increase motivation and show when therapy is complete."},
    {"t": "ul", "items": [
        "Structure begins with a clear [[frame|the agreed boundaries of therapy: time, place, fee, length, confidentiality and contact rules]], which creates the **predictable, safe container** within which change can take place.",
        "Each session is itself structured with an **agenda**, so that limited time is spent on the client's **priorities** rather than drifting without direction or purpose.",
        "A course of therapy is structured into **three broad phases**, namely a **beginning, a middle and an end**, each with its own characteristic tasks and emphasis.",
        "The **beginning phase** covers engagement, assessment, [[case formulation|a working explanation of the client's problem that guides the plan]], contracting and the setting of **collaborative goals**.",
        "The **middle phase** is where **active intervention** and [[working through|repeatedly facing and resolving a problem until change becomes stable]] occur, using the techniques of the chosen model.",
        "The **end phase** covers **consolidation** of gains, [[relapse prevention|planning ahead to recognise and manage early warning signs of a setback]], termination and arrangements for **follow-up**.",
        "**Regular review** of progress is built into the structure, so that goals and methods are **revised** in light of outcome data rather than left unchanged out of habit.",
    ]},
    {"t": "fig", "key": "therapy_phases", "caption": "The three phases of a course of therapy."},
    {"t": "h3", "text": "Goal setting"},
    {"t": "ul", "items": [
        "Goals should be **collaborative**, meaning they are **jointly negotiated** with the client rather than dictated, because goals the client **owns** carry far more motivation and commitment.",
        "Goals are **derived from the formulation** and expressed where possible in the **client's own words**, so they stay relevant to the problem and meaningful to the person.",
        "Well-formed goals are **SMART**, that is Specific, Measurable, Achievable, Realistic or Relevant, and Time-bound, which turns a **vague wish** into a **concrete target** progress can be checked against.",
        "Goals are also best stated **positively** as something to move towards, kept under the **client's own control**, and **ordered by priority** so the most pressing targets come first.",
        ("Good goals are SMART:", [
            "Specific, Measurable, Achievable, Realistic/Relevant, Time-bound.",
            "Also stated positively, under the client's control, and ordered by priority.",
        ]),
        "Types: ultimate vs intermediate vs immediate (session) goals; outcome vs process goals; short- vs long-term.",
    ]},
    {"t": "ul", "items": [
        "{{THEORY}} Goal Attainment Scaling (Kiresuk & Sherman, 1968): each goal is scaled from \u22122 (much worse) to +2 (much better than expected) for individualised measurement.",
        "{{THEORY}} Goal consensus is a component of the therapeutic alliance (Bordin) and predicts outcome. {{c:Gabbard}}",
        "Revisiting: goals are revised as new information emerges or progress occurs.",
    ]},
    {"t": "mn", "word": "SMART", "text": "Specific, Measurable, Achievable, Realistic/Relevant, Time-bound \u2014 the test of a well-formed therapy goal."},
    {"t": "box", "kind": "info", "title": "Why structure and goals matter", "items": [
        "Structure gives direction, safety and efficiency and keeps therapy within its agreed frame.",
        "Clear goals increase motivation and let both parties recognise when therapy is complete.",
        "Collaborative goal-setting is itself a relationship-building, alliance-strengthening act.",
    ]},
    {"t": "h3", "text": "Good goals in practice"},
    {"t": "ul", "items": [
        "Stated positively (what the client will do or have, not only what to stop).",
        "Under the client's own control rather than dependent on others.",
        "Ordered by priority, so the most pressing or achievable goals come first.",
        "Broken into ultimate, intermediate and immediate (session) targets.",
        "Reviewed regularly and revised as new information emerges or progress occurs.",
    ]},
])

# ========================= 11. PRE/POST ASSESSMENT =========================
_s("t11", "11", "Pre- and post-assessment", [
    {"t": "lead", "text": "A baseline (pre) gives diagnosis, severity, risk and a benchmark; the post-assessment evaluates change; follow-up checks maintenance. Repeated monitoring lets therapy be adjusted (feedback-informed treatment)."},
    {"t": "fig", "key": "assess_chart", "caption": "Illustrative PHQ-9 trajectory across baseline, weekly monitoring, post-treatment and follow-up, with the moderate-severity threshold marked."},
    {"t": "h3", "text": "Why assess before, during and after"},
    {"t": "ul", "items": [
        "**Pre-assessment** establishes a [[baseline|the client's starting level before therapy, used as the benchmark for later comparison]], giving diagnosis, severity, risk and strengths against which all change is judged.",
        "**Post-assessment**, conducted at the **end** of therapy, measures how far the client has **changed from baseline** and whether the agreed goals have actually been met.",
        "**Follow-up** assessment some weeks or months later checks whether gains have been **maintained** or whether the client has begun to **relapse** and needs further support.",
        "Ongoing **session-by-session monitoring** with brief measures, such as the OQ-45, CORE-OM or the Outcome and Session Rating Scales, feeds results back so therapy can be **adjusted in real time**.",
        "This continuous feedback, known as [[feedback-informed treatment|routinely using the client's own outcome scores to guide and adjust therapy]], **reduces deterioration**, especially for clients flagged as **not on track**.",
        "Repeated measurement also supports **accountability** and supervision, because **objective scores** give a shared, less biased picture of progress than clinical impression alone.",
    ]},
    {"t": "h3", "text": "Methods and measures"},
    {"t": "ul", "items": [
        "Methods: structured/unstructured interviews, behavioural observation, self-report and clinician-rated scales, projective tests when relevant, collateral information, physiological measures, client self-monitoring.",
        ("Commonly used measures:", [
            "Depression: BDI-II, PHQ-9, HAM-D.",
            "Anxiety: BAI, GAD-7, HAM-A; OCD: Y-BOCS; PTSD: PCL-5 / CAPS-5.",
            "General distress: SCL-90-R, GHQ; well-being/QoL: WHO-5, WHOQOL-BREF.",
            "Functioning: GAF, WHODAS 2.0; global impression: CGI; alliance: WAI.",
        ]),
    ]},
    {"t": "h3", "text": "Methods in sentences (exam form)"},
    {"t": "ul", "items": [
        "Assessment draws on **multiple methods**, including clinical interviews, [[behavioural observation|systematically watching and recording what the client actually does]], self-report questionnaires and clinician-rated scales.",
        "[[Self-report scales|questionnaires the client completes about their own symptoms]] such as the **PHQ-9 and GAD-7** are quick, standardised and **easy to repeat**, making them ideal for tracking change over time.",
        "[[Clinician-rated scales|measures scored by the therapist from interview and observation]] such as the Hamilton scales add an **expert perspective** that can complement the client's own account.",
        "[[Collateral information|accounts from family, teachers or other professionals who know the client]] can be sought **with consent** to give a **fuller, balanced** picture of functioning.",
        "The **same measures** should be used **before and after** therapy so that pre and post scores are **directly comparable** and genuine change can be distinguished from noise.",
        "Measures must be chosen for their **reliability and validity** and for their fit with the client's **language, culture and literacy**, because an unsuitable test yields misleading numbers.",
    ]},
    {"t": "note", "text": "{{GUIDE}} WHO-5 and WHODAS 2.0 are WHO instruments for well-being and functioning. {{c:WHO}}"},
    {"t": "table", "caption": "Two widely used screeners and their severity bands.",
     "head": ["Scale", "Construct", "Severity bands"],
     "rows": [
        ["PHQ-9", "Depression (9 items)", "0\u20134 none; 5\u20139 mild; 10\u201314 moderate; 15\u201319 moderately severe; 20\u201327 severe {{c:cheat sheet}}"],
        ["GAD-7", "Generalised anxiety (7 items)", "0\u20134 minimal; 5\u20139 mild; 10\u201314 moderate; 15\u201321 severe {{c:cheat sheet}}"],
     ]},
    {"t": "h3", "text": "Evaluating change"},
    {"t": "ul", "items": [
        "[[Statistical significance|a result unlikely to be due to chance across a group]] answers whether a treatment works **on average**, but it does not tell you whether an **individual client** meaningfully improved.",
        "[[Clinical significance|a change large enough to matter in the client's real life]] asks whether this particular client crossed from the **dysfunctional** into the **functional** range on the measure used.",
        "{{THEORY}} The **Reliable Change Index** (Jacobson & Truax, 1991) tests whether an individual's change **exceeds measurement error**, with a value above **1.96** taken as reliable rather than random fluctuation. {{c:Jacobson & Truax 1991}}",
        "[[Effect size|a standardised number showing how large a change is, such as Cohen's d]] and simple **percentage improvement** give a sense of the **magnitude** of change beyond mere statistical testing.",
        "[[Single-case designs|research designs that track one client repeatedly, such as AB, ABAB or multiple-baseline]] let the clinician evaluate change for an **individual** rather than only for a group.",
        "Change should be judged on **several measures** and on the client's own account of **functioning**, because any one score can move for reasons unrelated to genuine therapeutic gain.",
    ]},
    {"t": "mn", "word": "BE-MAF", "text": "Baseline (pre), Evaluate at end (post), Monitor each session, Appraise change (clinical vs statistical, RCI), Follow-up to check maintenance \u2014 the measurement points and tests of change."},
    {"t": "box", "kind": "warn", "title": "Interpret with caution", "items": [
        "Placebo, regression to the mean, practice effects, response bias.",
        "Use reliable, valid measures suited to the client's language, culture and literacy.",
        "Use the same measures before and after; share results with the client.",
    ]},
])

# ========================= 12. EVIDENCE-BASED PRACTICE =====================
_s("t12", "12", "Evidence-based practice (EBP) of therapies", [
    {"t": "lead", "text": "EBP integrates the best research with clinical expertise in the context of the patient's characteristics, culture and preferences."},
    {"t": "h3", "text": "Core ideas in sentences (exam form)"},
    {"t": "ul", "items": [
        "{{GUIDE}} The **APA (2006)** defines evidence-based practice as the **integration** of the **best available research** with **clinical expertise** in the context of the patient's characteristics, culture and preferences. {{c:APA EBPP 2006}}",
        "The idea was borrowed from [[evidence-based medicine|Sackett's approach of using the best current research to inform clinical decisions]] and adapted so psychology weighs **research alongside judgement** and client values.",
        "[[Empirically supported treatments|specific therapies shown effective for specific disorders in controlled trials]] form **one pillar**, judged by criteria such as those of **APA Division 12** (Chambless and Hollon, 1998).",
        "A treatment is called **well-established** when at least **two good independent trials** show it superior to placebo or another treatment, using **manuals** and clearly defined client samples.",
        "The [[hierarchy of evidence|a ranking of study designs by how far they guard against bias]] places **systematic reviews and randomised trials** above cohort studies, single-case designs and expert opinion.",
        "A key distinction is between [[efficacy|how well a treatment works under controlled trial conditions]] and [[effectiveness|how well it works in ordinary, messy clinical practice]], since **trial results do not always transfer**.",
        "EBP is carried out as a **repeating cycle**: **asking** a clear question, **searching** for the best evidence, **appraising** it, **integrating** it with expertise and client values, **applying** it and **evaluating** the outcome.",
    ]},
    {"t": "ul", "items": [
        "{{GUIDE}} Definition (APA Presidential Task Force, 2006): the integration of the best available research with clinical expertise in the context of patient characteristics, culture and preferences. The concept came from evidence-based medicine (Sackett). {{c:APA EBPP 2006}}",
        ("{{THEORY}} Empirically supported treatments (ESTs):", [
            "APA Division 12 (Society of Clinical Psychology) criteria, Chambless & Hollon (1998).",
            "Well-established: \u2265 two good between-group studies by independent teams (or a large single-case series) showing superiority to placebo or another treatment, with manuals and clear samples.",
            "Probably efficacious: a lower tier of evidence.",
        ]),
    ]},
    {"t": "ul", "items": [
        "Hierarchy of evidence: systematic reviews/meta-analyses of RCTs > RCTs > controlled non-randomised > cohort/case-control > single-case > expert opinion.",
        "Efficacy (works in controlled trials) vs effectiveness (works in real practice) vs clinical utility (feasible, acceptable, cost-effective).",
        ("Examples of well-supported treatments:", [
            "CBT for depression, panic, social anxiety, GAD; ERP for OCD.",
            "Prolonged exposure, CPT and EMDR for PTSD; behavioural activation and IPT for depression.",
            "DBT for borderline personality disorder; family-based treatment for adolescent anorexia.",
            "Parent management training for child conduct problems; CBT-I for insomnia.",
            "Social skills training, family psychoeducation and cognitive remediation for schizophrenia.",
        ]),
    ]},
    {"t": "ul", "items": [
        "{{GUIDE}} Guidelines: NICE (UK), APA, and WHO mhGAP recommend specific treatments. {{c:WHO mhGAP}}",
        "{{GUIDE}} WHO mhGAP targets non-specialist settings so that evidence-based treatments reach low-resource populations; task-sharing lets trained lay providers deliver brief, manualised care under supervision.",
        "{{GUIDE}} India: culturally adapted low-intensity care by lay counsellors \u2014 the Healthy Activity Programme (PREMIUM trials, Goa) for depression and the Thinking Healthy Programme for perinatal depression \u2014 shows evidence-based care can be scaled up.",
        "Steps of EBP: ask a clear clinical question \u2192 search best evidence \u2192 appraise \u2192 integrate with expertise and client values \u2192 apply \u2192 evaluate.",
    ]},
    {"t": "h3", "text": "The outcome debate and common factors"},
    {"t": "ul", "items": [
        "{{THEORY}} **Eysenck (1952)** provocatively claimed that psychotherapy was **no more effective** than [[spontaneous remission|natural recovery that happens over time without any treatment]], which stung the field into gathering better evidence.",
        "{{THEORY}} **Smith and Glass (1977)** answered with a [[meta-analysis|a statistical pooling of many separate studies into one overall estimate]] of **375 studies**, finding a mean **effect size of about 0.68** favouring therapy.",
        "That effect size means the **average treated person** ended up better off than roughly **seventy-five per cent** of untreated people, providing strong aggregate evidence that psychotherapy **works**.",
        "{{THEORY}} The so-called **Dodo bird verdict** (Rosenzweig; Luborsky) holds that all [[bona fide|genuine, properly delivered]] therapies produce about **equal benefit**, pointing towards shared **common factors** rather than unique techniques.",
        "This debate matters because it warns against assuming that a favoured **technique** is the main active ingredient when the **relationship and the client** may matter more.",
    ]},
    {"t": "fig", "key": "variance_bars", "caption": "Lambert's approximate outcome-variance estimates."},
    {"t": "ul", "items": [
        ("{{THEORY}} Common factors (Lambert):", [
            "Therapeutic relationship (alliance, empathy); client expectations and hope (placebo).",
            "A healing setting and rationale; a credible ritual; new learning and mastery.",
            "Lambert (1992) estimates: extratherapeutic/client 40%, relationship 30%, expectancy 15%, technique 15% (approximate and debated). {{c:Corsini}}",
        ]),
        "{{THEORY}} Wampold's contextual model argues these factors, more than specific techniques, explain most benefit.",
        "Balance: evidence-based techniques work best delivered within a strong relationship.",
    ]},
    {"t": "note", "text": "{{THEORY}} Lister-Ford discusses EBP and routine outcome monitoring as complements to RCTs. {{c:Lister-Ford}}"},
    {"t": "note", "text": "{{GUIDE}} Corey summarises the APA evidence-based-practice policy and its three pillars (research, expertise, client values). {{c:Corey}}"},
    {"t": "mn", "word": "EBP = 3 circles", "text": "Best research evidence + clinical expertise + patient characteristics, culture and preferences."},
    {"t": "box", "kind": "warn", "title": "Criticisms and challenges of EBP", "items": [
        "RCT samples often unrepresentative (strict exclusion, no comorbidity); limited generalisability to diverse and Indian populations.",
        "Manuals can reduce flexibility; researcher-allegiance effects; publication bias.",
        "Under-emphasis of common factors and the therapist's contribution.",
        "Outcome focus on symptoms rather than well-being or meaning.",
        "Practice-based evidence, routine outcome monitoring, cultural adaptation and task-shifting are suggested remedies.",
    ]},
])

# ---------------------------------------------------------------------------
# Per-topic exam answer skeletons and practice questions. These are study
# aids (how to structure an answer and what to revise), not factual claims,
# so they add no citations. Appended to each section so revision cues travel
# with the content in both deliverables.
_SKELETONS = {
    "t1": (["Open with one classic definition (Wolberg or Frank).",
            "List the six elements common to all definitions.",
            "State Wolberg's three aims.",
            "Distinguish psychotherapy from counselling and medical treatment (table rows).",
            "Close: the distinction is one of degree; methods often combine."],
           ["Define psychotherapy and bring out its essential features.",
            "Differentiate psychotherapy from counselling.",
            "Discuss Wolberg's three aims of psychotherapy."]),
    "t2": (["Group objectives as symptom relief, change, insight, relationships, growth, prevention.",
            "Distinguish outcome from process goals.",
            "Note how objectives vary by model (psychodynamic, CBT, humanistic, systemic)."],
           ["Describe the objectives of psychotherapy.",
            "How do the goals of psychotherapy differ across the major schools?"]),
    "t4": (["Who may practise in India: RCI Act, CRR, M.Phil.; MHCA recognition.",
            "Components of training and expected competencies.",
            "Issues: quality variation, cultural adaptation, gate-keeping, self-care.",
            "Role of supervision; WHO task-sharing for the supply gap."],
           ["Discuss the issues involved in training professional therapists.",
            "What competencies should the training of a clinical psychologist develop?"]),
    "t5": (["Define ethics vs law (aspirational vs enforceable minimum).",
            "State APA's five general principles.",
            "List the major ethical issues (competence, consent, confidentiality, boundaries).",
            "Give the ethical decision-making steps; add the Tarasoff/POCSO limits of confidentiality."],
           ["Discuss the ethical issues involved in psychotherapy.",
            "Explain the limits of confidentiality with reference to law.",
            "Describe a model of ethical decision-making."]),
    "t6": (["List the main Indian statutes (RCI, MHCA, RPwD, POCSO, consumer/negligence law).",
            "Expand MHCA 2017 rights and Section 115.",
            "State the therapist's legal duties; contrast legal with ethical standards."],
           ["Describe the legal issues involved in therapy work in India.",
            "Outline the key provisions of the Mental Healthcare Act, 2017 relevant to a therapist."]),
    "t7": (["Introduce therapy as a contract of mutual rights and duties.",
            "Give client rights and responsibilities (table).",
            "Give therapist responsibilities.",
            "Link rights to MHCA 2017 ss.18\u201325 and a rights-based model of care."],
           ["Discuss the rights and responsibilities of the client and therapist in psychotherapy."]),
    "t8": (["Define informed consent as a process, not a signature.",
            "List the elements (capacity, disclosure, comprehension, voluntariness, documentation).",
            "Explain consent vs assent and the handling of minors (ICMR 7\u201318).",
            "Note adults lacking capacity (MHCA), recording consent and emergencies."],
           ["Explain informed consent and its essential elements.",
            "Discuss consent and assent in work with minors."]),
    "t9": (["Planning: intake/assessment, formulation (4 Ps), treatment plan, frame.",
            "Describe a typical CBT session structure.",
            "Recording: types of records, SOAP/DAP/BIRP, qualities of good records, security.",
            "Add APA record-keeping guidance and MHCA s.25 access right."],
           ["Describe the planning and recording of therapy sessions.",
            "What are the characteristics of good clinical records?"]),
    "t10": (["Explain why structure matters (direction, safety, efficiency).",
             "Describe the beginning, middle and end phases.",
             "State how goals are set and the SMART criteria.",
             "Add Goal Attainment Scaling and goal consensus as an alliance factor."],
            ["Discuss structuring of therapy and setting of goals.",
             "Explain SMART goals with reference to psychotherapy."]),
    "t11": (["State the purpose of pre, post and follow-up assessment and ongoing monitoring.",
             "List methods and commonly used measures (with PHQ-9 and GAD-7 bands).",
             "Explain statistical vs clinical significance and the Reliable Change Index.",
             "Note cautions in interpreting change."],
            ["Discuss pre- and post-assessment in psychotherapy.",
             "Explain how change in therapy is evaluated (RCI and clinical significance)."]),
    "t12": (["Define EBP (APA 2006) and its three pillars.",
             "Explain ESTs and the hierarchy of evidence; give examples.",
             "Summarise the outcome debate (Eysenck, Smith & Glass, Dodo bird).",
             "Explain common vs specific factors (Lambert) and the criticisms of EBP."],
            ["Explain evidence-based practice in psychotherapy.",
             "Discuss the common-factors vs specific-factors debate.",
             "Critically evaluate the evidence-based-practice movement."]),
}

for _sec in SECTIONS:
    _sk = _SKELETONS.get(_sec["id"])
    if not _sk:
        continue
    steps, qs = _sk
    _sec["blocks"].append({"t": "box", "kind": "key", "title": "Answer skeleton", "items": list(steps)})
    _sec["blocks"].append({"t": "box", "kind": "info", "title": "Practice questions", "items": list(qs)})
