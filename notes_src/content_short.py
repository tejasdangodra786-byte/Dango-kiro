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
CREDIT = "Made by Tejas Dangodra | M.Phil. Clinical Psychology"

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
    {"t": "h3", "text": "Citation convention"},
    {"t": "ul", "items": [
        "Citations look like {{c:Corey, PDF p.74-77}}. \u201cPDF p.\u201d is the page index inside the source file in the repository, which usually differs from the printed page number (printed pages change between editions and are missing from reflowed e-books; the PDF index is exactly reproducible).",
        "Everything quotable is paraphrased; only verified page indices are cited. Where a page index was not verified, no page is given.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no laws, citations, clause numbers or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "{{SEC}} IACP (Indian Association of Clinical Psychologists): the association's code of conduct could not be read from a primary source (iacp.in was not reachable). IACP is mentioned only in general terms (a members' code exists and broadly aligns with APA/RCI duties). No IACP clause numbers are quoted.",
        "Page citations use the PDF page index inside each repository file, not the printed page number; this convention is stated so every citation is reproducible.",
        "Primary texts read in full and relied on here: RCI Act 1992, Mental Healthcare Act 2017, POCSO Act 2012, ICMR 2017 ethical guidelines (research), and the APA Ethics Code 2017, Record Keeping Guidelines and Telepsychology Guidelines.",
        "Several repository files (for example scanned test manuals) are image-only PDFs with no text layer; they were not OCR'd and are not cited.",
        "Statements about Indian law are for study; in practice read the current text and seek advice, as laws change.",
    ]},
]

# Short reading list keyed to the topics. Full source rows live in the HTML
# intro part; this is the compact version shared by both deliverables.
REFERENCES = [
    {"t": "lead", "text": "Key sources (repository files and primary legal/ethical texts)."},
    {"t": "ul", "items": [
        "Wolberg, L. R. \u2014 The Technique of Psychotherapy (definitions; {{c:Wolberg, PDF p.29}}).",
        "Corey, G. \u2014 Theory and Practice of Counseling and Psychotherapy (ethics {{c:Corey, PDF p.74-77}}; EBPP {{c:Corey, PDF p.82}}).",
        "Beck, J. S. \u2014 Cognitive Behavior Therapy: Basics and Beyond (session structure {{c:Beck, PDF p.82 & 122-123}}).",
        "Jongsma, A. E. et al. \u2014 The Complete Adult Psychotherapy Treatment Planner (6-step planning {{c:Jongsma, PDF p.24-26}}).",
        "Gabbard, G. O. \u2014 Long-Term Psychodynamic Psychotherapy (alliance {{c:Gabbard, PDF p.58}}; supervision {{c:Gabbard, PDF p.192-193}}).",
        "Lister-Ford, C. \u2014 A Short Introduction to Psychotherapy (EBP {{c:Lister-Ford, PDF p.114-118}}).",
        "Corsini & Wedding \u2014 Current Psychotherapies (common factors {{c:Corsini, PDF p.194}}; ethics {{c:Corsini, PDF p.594-605}}).",
        "Psychological rating scales cheat sheet (PHQ-9 bands {{c:cheat sheet, PDF p.78}}; GAD-7 {{c:cheat sheet, PDF p.80}}).",
        "Primary legal/ethical texts: RCI Act 1992; Mental Healthcare Act 2017; POCSO Act 2012; ICMR 2017; APA Ethics Code 2017; APA Record Keeping & Telepsychology Guidelines; APA EBPP policy (2006).",
        "{{SEC}} IACP code of conduct \u2014 secondary summaries only; primary text not accessed.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. DEFINITIONS =================================
_s("t1", "1", "Definitions of psychotherapy", [
    {"t": "lead", "text": "Psychotherapy = a planned, systematic, psychological (non-medical) treatment in which a trained professional uses a confiding relationship and theory-based methods to help a help-seeking client change feelings, thoughts and behaviour."},
    {"t": "h3", "text": "Classic definitions"},
    {"t": "ul", "items": [
        "{{THEORY}} Wolberg (1954): treatment of problems of an emotional nature in which a trained person deliberately establishes a professional relationship to (1) remove, modify or retard existing symptoms, (2) mediate disturbed patterns of behaviour, and (3) promote positive personality growth. {{c:Wolberg, PDF p.29}}",
        "{{THEORY}} Jerome Frank (1961, Persuasion and Healing): a planned, emotionally charged, confiding interaction between a trained, socially sanctioned healer and a sufferer, aiming to change the sufferer's feelings, attitudes and behaviour.",
        "{{THEORY}} Norcross (1990): the informed and intentional application of clinical methods and interpersonal stances derived from established psychological principles, to help people modify behaviours, cognitions, emotions and/or personal characteristics in directions the participants deem desirable.",
        "{{THEORY}} Strupp: a learning process in which a trained person uses the relationship to help the patient overcome distress and develop more adaptive functioning.",
    ]},
    {"t": "box", "kind": "key", "title": "Elements common to all definitions", "items": [
        "A trained, professional therapist (socially sanctioned).",
        "A help-seeking client in distress.",
        "A psychological (not pharmacological / surgical) method.",
        "A planned, systematic, theory-based process.",
        "A confiding, therapeutic relationship (alliance).",
        "The aim of change: symptom relief, behaviour change and growth.",
    ]},
    {"t": "h3", "text": "Wolberg's three aims"},
    {"t": "ul", "items": [
        "Remove, modify or retard existing symptoms.",
        "Mediate (correct) disturbed patterns of behaviour.",
        "Promote positive personality growth and development. {{c:Wolberg, PDF p.29}}",
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
        "Counselling usually deals with normal developmental, situational and adjustment problems, is more present- and problem-focused, and is shorter.",
        "Psychotherapy deals with deeper, more enduring or pathological problems and personality reconstruction, and is usually longer; the overlap is large and the distinction is mostly one of degree.",
        "Psychotherapy works through communication, relationship and learning rather than drugs; both are often combined (for example for depression or psychosis).",
    ]},
    {"t": "box", "kind": "info", "title": "Exam pointer", "items": [
        "If asked to \u201cdefine psychotherapy\u201d, give one classic definition (Wolberg or Frank), then list the six common elements.",
        "If asked to \u201cdistinguish\u201d, use the table rows: focus, time frame, method and overlap.",
    ]},
])

# ========================= 2. OBJECTIVES ==================================
_s("t2", "2", "Objectives (goals) of psychotherapy", [
    {"t": "lead", "text": "Objectives run from immediate symptom relief to long-term growth; most therapies pursue several at once, weighted by the model and the client's needs."},
    {"t": "ul", "items": [
        "Relieve symptoms and distress (anxiety, depression, obsessions, etc.).",
        "Modify maladaptive behaviour, thoughts and emotions (change).",
        "Gain insight into the origin and meaning of problems (especially psychodynamic work).",
        "Improve interpersonal relationships and social functioning.",
        "Strengthen coping, problem solving and adaptive defences; improve self-esteem and self-efficacy.",
        "Resolve inner conflicts and past trauma; foster emotional expression and regulation.",
        "Promote personal growth, self-actualisation and a sense of meaning.",
        "Prevent relapse and improve quality of life; in severe illness, maintain functioning and prevent hospitalisation.",
    ]},
    {"t": "box", "kind": "info", "title": "Outcome vs process goals", "items": [
        "Outcome goals = the end the client wants (less panic, better sleep).",
        "Process goals = the in-therapy means (insight, alliance, skills practice).",
        "Both are negotiated and revisited as therapy proceeds.",
    ]},
    {"t": "h3", "text": "How objectives differ by model"},
    {"t": "ul", "items": [
        "Psychodynamic: insight into unconscious conflict, working through, more adaptive defences and object relations.",
        "Cognitive-behavioural: symptom reduction, changed thoughts and behaviour, new coping skills and relapse prevention.",
        "Humanistic: self-actualisation, congruence, self-acceptance, autonomy and a sense of meaning.",
        "Systemic/family: changed interaction patterns, clearer boundaries and improved relationships.",
    ]},
    {"t": "box", "kind": "info", "title": "Ordering objectives in practice", "items": [
        "Safety first: manage risk (suicidality, harm, abuse) before other goals.",
        "Then stabilise symptoms and distress, and build the alliance and coping skills.",
        "Then address underlying patterns, relationships and growth.",
        "Finally consolidate gains and plan relapse prevention and maintenance.",
    ]},
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
        "Wide variation in quality, duration and supervision of training; shortage of trained supervisors.",
        "Over-reliance on Western models with limited cultural adaptation; need for indigenous, culturally sensitive approaches.",
        "Whether personal therapy for trainees should be mandatory (to manage countertransference and blind spots).",
        "Gate-keeping: judging suitability (empathy, maturity, integrity) and remediating or dismissing unfit trainees.",
        "Training vs licensing gap: unregulated practice by untrained persons (\u201cquacks\u201d); need to protect the public.",
        "Therapist stress, burnout, vicarious traumatisation and compassion fatigue; importance of self-care and peer support.",
        "Boundaries and dual relationships in trainee-supervisor and small-community settings.",
        "Training in tele-therapy and its ethics.",
        "Practising within the limits of one's competence; knowing when to refer.",
        "Keeping evidence-based: integrating new research into continuing professional development.",
    ]},
    {"t": "mn", "word": "TRAIN", "text": "Theory, Relationship skills, Assessment, Intervention, iNtegrity/ethics \u2014 the core domains a programme must build."},
    {"t": "box", "kind": "info", "title": "WHO: task-sharing and scaling training", "text":
        "{{GUIDE}} The WHO mhGAP Intervention Guide and Problem Management Plus (PM+) support training non-specialist and lay providers to deliver brief psychological care in low-resource settings \u2014 relevant to the training-supply gap in India. {{c:WHO mhGAP}}"},
    {"t": "h3", "text": "Supervision and self-care"},
    {"t": "ul", "items": [
        "{{THEORY}} Supervision protects the client and develops the trainee: direct observation, audio/video review, case discussion and reflective feedback. {{c:Gabbard, PDF p.192-193}}",
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
        "Competence: practise only within the limits of one's training and experience; refer when the problem is outside one's competence; keep knowledge and skills updated. {{c:APA 2017, Std 2.01}}",
        "Informed consent must be obtained before assessment and treatment and treated as an ongoing process (see Section 8).",
        ("Confidentiality and its limits:", [
            "{{CODE}} Private information shared only with consent (APA Std 4.01\u20134.02).",
            "Exceptions: serious risk to self/others, abuse of a child, court order.",
        ]),
        ("Dual (multiple) relationships and boundaries:", [
            "{{CODE}} Avoid social, business, financial or sexual relationships with clients (APA Std 3.05\u20133.08).",
            "Sexual intimacy with a current client, and with a former client for a specified period, is prohibited.",
            "Manage gifts, self-disclosure, touch and social-media contact carefully.",
        ]),
        "Fees and contracts: be clear and honest; do not exploit.",
        "Record keeping and data security.",
        "Termination and referral: do not abandon clients; provide handover.",
        "Cultural sensitivity: avoid bias by gender, caste, religion, sexuality, disability, socio-economic status.",
        "Supervision and consultation in ethical dilemmas.",
        "Advertising and claims: no false or unsubstantiated outcome claims.",
        "Research ethics: consent, no harm, confidentiality, ethics-committee approval.",
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
    {"t": "fig", "key": "conf_tree", "caption": "When confidentiality may or must be broken."},
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
])

# ========================= 6. LEGAL (INDIA) ================================
_s("t6", "6", "Legal issues in therapy work (India)", [
    {"t": "lead", "text": "Indian psychotherapy practice is governed by a statutory registration law, a rights-based mental-health law, child-protection law, disability law, consumer law and tele-practice guidelines."},
    {"t": "ul", "items": [
        "{{LAW}} RCI Act, 1992: regulates training, registration and conduct of rehabilitation professionals including clinical psychologists; practising without registration is an offence. {{c:RCI Act 1992}}",
        ("{{LAW}} Mental Healthcare Act, 2017 \u2014 rights-based law: {{c:MHCA 2017}}", [
            "Rights of persons with mental illness: access to affordable quality care, community living, protection from cruel or degrading treatment, equality and non-discrimination, information, confidentiality, access to records, personal contacts, legal aid, and to make complaints.",
            "Advance directive: a person may state how they wish to be treated and name a nominated representative.",
            "Capacity to make treatment decisions is presumed unless shown otherwise.",
            "{{LAW}} Section 115: a person who attempts suicide is presumed to be under severe stress and is not punishable; the government must provide care and rehabilitation. {{c:MHCA 2017, s.115}}",
            "Restrictions on ECT (no ECT without anaesthesia and muscle relaxant; not for minors), prohibition of sterilisation as treatment, and limited use of restraints.",
        ]),
        "{{LAW}} Rights of Persons with Disabilities Act, 2016 (includes mental illness and intellectual disability).",
        "{{LAW}} POCSO Act, 2012 (mandatory reporting) and the Juvenile Justice Act (care of children).",
        "{{LAW}} Consumer Protection Act and civil/criminal law on negligence (professional malpractice); courts may seek records or testimony.",
        "{{GUIDE}} Tele-mental health: the Telemedicine Practice Guidelines (2020) and guidance on tele-psychiatry apply.",
    ]},
    {"t": "box", "kind": "key", "title": "Legal vs ethical", "items": [
        "Law sets the minimum enforceable standard; breach can mean penalty or loss of registration.",
        "Ethics sets the higher aspirational standard of conduct.",
        "Therapist's legal duties: valid consent, confidentiality, competent practice, accurate records, reporting obligations, no exploitation, referral when needed.",
    ]},
])

# ========================= 7. RIGHTS & RESPONSIBILITIES ====================
_s("t7", "7", "Rights and responsibilities in psychotherapy", [
    {"t": "lead", "text": "Therapy is a contract: both parties hold rights and duties. Clarifying them at the start supports the alliance and informed consent."},
    {"t": "fig", "key": "rights_resp", "caption": "Balance of client and therapist rights and responsibilities."},
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
    {"t": "h3", "text": "Special issues"},
    {"t": "fig", "key": "consent_flow", "caption": "Consent and assent decision flow for minors and adults lacking capacity."},
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
    {"t": "note", "text": "{{THEORY}} Jongsma's Treatment Planner sets out a 6-step planning sequence (problem selection \u2192 definition \u2192 goals \u2192 objectives \u2192 interventions \u2192 diagnosis). {{c:Jongsma, PDF p.24-26}}"},
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
    {"t": "note", "text": "{{THEORY}} J. Beck describes this session structure. {{c:Beck, PDF p.82 & 122-123}}"},
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
    {"t": "fig", "key": "soap_dap", "caption": "SOAP and DAP progress-note formats compared."},
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
        "A clear frame (contract), a clear agenda each session, defined phases, and regular review of progress.",
        "Phases: (1) Beginning \u2014 engagement, assessment, formulation, contracting, goal setting; (2) Middle \u2014 active intervention and working through; (3) End \u2014 consolidation, relapse prevention, termination, follow-up.",
    ]},
    {"t": "fig", "key": "therapy_phases", "caption": "The three phases of a course of therapy."},
    {"t": "h3", "text": "Goal setting"},
    {"t": "ul", "items": [
        "Goals are collaborative (jointly negotiated), derived from the formulation, and in the client's own words.",
        ("Good goals are SMART:", [
            "Specific, Measurable, Achievable, Realistic/Relevant, Time-bound.",
            "Also stated positively, under the client's control, and ordered by priority.",
        ]),
        "Types: ultimate vs intermediate vs immediate (session) goals; outcome vs process goals; short- vs long-term.",
    ]},
    {"t": "ul", "items": [
        "{{THEORY}} Goal Attainment Scaling (Kiresuk & Sherman, 1968): each goal is scaled from \u22122 (much worse) to +2 (much better than expected) for individualised measurement.",
        "{{THEORY}} Goal consensus is a component of the therapeutic alliance (Bordin) and predicts outcome. {{c:Gabbard, PDF p.58}}",
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
    {"t": "fig", "key": "assess_chart", "caption": "Pre, during, post and follow-up measurement points."},
    {"t": "h3", "text": "Why assess before, during and after"},
    {"t": "ul", "items": [
        "Pre-assessment (baseline) provides diagnosis, severity, risk, strengths and a benchmark against which change is judged.",
        "Post-assessment evaluates change at the end of therapy; follow-up assessment checks whether gains are maintained.",
        "Ongoing session-by-session monitoring (for example OQ-45, CORE-OM, Outcome Rating Scale / Session Rating Scale) gives feedback so therapy can be adjusted (feedback-informed treatment).",
        "Tracking progress reduces deterioration, especially for clients who are flagged as \u201cnot on track\u201d.",
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
    {"t": "note", "text": "{{GUIDE}} WHO-5 and WHODAS 2.0 are WHO instruments for well-being and functioning. {{c:WHO}}"},
    {"t": "table", "caption": "Two widely used screeners and their severity bands.",
     "head": ["Scale", "Construct", "Severity bands"],
     "rows": [
        ["PHQ-9", "Depression (9 items)", "0\u20134 none; 5\u20139 mild; 10\u201314 moderate; 15\u201319 moderately severe; 20\u201327 severe {{c:cheat sheet, PDF p.78}}"],
        ["GAD-7", "Generalised anxiety (7 items)", "0\u20134 minimal; 5\u20139 mild; 10\u201314 moderate; 15\u201321 severe {{c:cheat sheet, PDF p.80}}"],
     ]},
    {"t": "h3", "text": "Evaluating change"},
    {"t": "ul", "items": [
        "Statistical significance (group level) vs clinical significance (individual level).",
        "{{THEORY}} Reliable Change Index (RCI; Jacobson & Truax, 1991): change must exceed measurement error (RCI > 1.96). {{c:Jacobson & Truax 1991}}",
        "Clinically significant change: the client crosses a cutoff from the dysfunctional into the functional range.",
        "Effect size (Cohen's d) and percentage improvement.",
        "Single-case designs (AB, ABAB, multiple baseline) for individual evaluation.",
    ]},
    {"t": "box", "kind": "warn", "title": "Interpret with caution", "items": [
        "Placebo, regression to the mean, practice effects, response bias.",
        "Use reliable, valid measures suited to the client's language, culture and literacy.",
        "Use the same measures before and after; share results with the client.",
    ]},
])

# ========================= 12. EVIDENCE-BASED PRACTICE =====================
_s("t12", "12", "Evidence-based practice (EBP) of therapies", [
    {"t": "lead", "text": "EBP integrates the best research with clinical expertise in the context of the patient's characteristics, culture and preferences."},
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
        "{{THEORY}} Eysenck (1952) claimed psychotherapy was no better than spontaneous remission.",
        "{{THEORY}} Smith & Glass (1977) meta-analysed 375 studies, finding a mean effect size of about 0.68 (the average treated person better off than ~75% of untreated persons).",
        "{{THEORY}} \u201cDodo bird verdict\u201d (Rosenzweig; Luborsky): all bona fide therapies work about equally \u2014 pointing to common factors.",
    ]},
    {"t": "fig", "key": "variance_bars", "caption": "Lambert's approximate outcome-variance estimates."},
    {"t": "ul", "items": [
        ("{{THEORY}} Common factors (Lambert):", [
            "Therapeutic relationship (alliance, empathy); client expectations and hope (placebo).",
            "A healing setting and rationale; a credible ritual; new learning and mastery.",
            "Lambert (1992) estimates: extratherapeutic/client 40%, relationship 30%, expectancy 15%, technique 15% (approximate and debated). {{c:Corsini, PDF p.194}}",
        ]),
        "{{THEORY}} Wampold's contextual model argues these factors, more than specific techniques, explain most benefit.",
        "Balance: evidence-based techniques work best delivered within a strong relationship.",
    ]},
    {"t": "note", "text": "{{THEORY}} Lister-Ford discusses EBP and routine outcome monitoring as complements to RCTs. {{c:Lister-Ford, PDF p.114-118}}"},
    {"t": "note", "text": "{{GUIDE}} Corey summarises the APA evidence-based-practice policy and its three pillars (research, expertise, client values). {{c:Corey, PDF p.82}}"},
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
