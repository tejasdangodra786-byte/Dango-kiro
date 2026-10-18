# -*- coding: utf-8 -*-
"""Single source of truth for Psychotherapy Unit III notes (Interviewing).

Both the HTML build (build_unit3.py -> parts) and the DOCX build
(build_docx_unit3.py) import SECTIONS from this module so the two deliverables
carry identical content.

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

TITLE = "Psychotherapy \u2014 Unit III: Interviewing"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

# Cover / title-page sub-heading used by both deliverables.
SUBTITLE = "M.Phil. Clinical Psychology \u2014 short-form study notes (bullet style)"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens both deliverables.
INTRO = [
    {"t": "lead", "text": "These are concise, exam-ready revision notes for Unit III of the Psychotherapy paper (M.Phil. Clinical Psychology). They cover the clinical interview in full, from its objectives and techniques to its types, structured and unstructured forms, micro-skills, the core responding skills, the use of silence and the verbal and non-verbal components, and they fold in the relevant professional bodies (RCI, IACP, WHO and APA) wherever each is examinable."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one syllabus topic, written as a 5-mark exam answer in full sentences rather than terse fragments, so you can revise straight from the point and reproduce it under exam conditions.",
        "Every topic opens with a one-line summary, carries a flowchart figure that pictures the idea, and closes with a mnemonic that packs the key points into a single memorable word.",
        "A short source note and a references list sit at the end; the references name only the books (full title and author) and the online or primary sources that were actually used.",
    ]},
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
        "A citation in square brackets names the full book and its author, for example {{c:Wolberg}} or {{c:Ivey}}, so you can trace any claim back to the work it comes from.",
        "Where a point rests on a professional body rather than a book, the citation names the document or body and its year, for example {{c:APA 2017}}, keeping any section number.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail.",
        "Hard clinical terms are kept, but each one is explained right after it in a coloured bracket, like this: [[rapport|the warm, trusting working connection between interviewer and client]]. Just skim the coloured bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no laws, citations, clause numbers or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "{{SEC}} IACP (Indian Association of Clinical Psychologists): its members' code of conduct is represented only from secondary summaries, so it is described in general terms as broadly aligning with APA and RCI duties, and no IACP clause numbers are quoted.",
        "Professional-body characterisations of the interview (APA, WHO) are given as characterisations drawn from standard reference and guidance wording, not as verbatim quotations, and no page numbers are attached.",
        "Primary texts relied on directly for ethics include the APA Ethics Code 2017 and the Rehabilitation Council of India (RCI) Act 1992; statements about bodies are for study, so read the current text before relying on them in practice.",
    ]},
]

# Short reading list keyed to the topics.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the books used and the primary ethical and online sources consulted for the interviewing content."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Lewis R. Wolberg \u2014 The Technique of Psychotherapy (the interview, interviewing technique, rapport and silence).",
        "Allen E. Ivey, Mary Bradford Ivey & Carlos P. Zalaquett \u2014 Intentional Interviewing and Counseling (micro-skills hierarchy and the basic listening sequence).",
        "Gerard Egan \u2014 The Skilled Helper (SOLER attending and active listening).",
        "William R. Miller & Stephen Rollnick \u2014 Motivational Interviewing (open questions, affirmation, reflection, summary \u2014 OARS).",
        "Gerald Corey \u2014 Theory and Practice of Counseling and Psychotherapy (listening, reflection and confrontation).",
        "Richard S. Sharf \u2014 Theories of Psychotherapy and Counseling (structured and unstructured interviews, reflection).",
        "Glen O. Gabbard \u2014 Long-Term Psychodynamic Psychotherapy: A Basic Text (the clinical interview, silence and the therapeutic frame).",
        "Nitin Sethi \u2014 Psychiatric History Taking (history, mental status examination and non-verbal observation in the interview).",
    ]},
    {"t": "h3", "text": "Primary ethical and online sources"},
    {"t": "ul", "items": [
        "American Psychological Association \u2014 Ethics Code (2017); APA Dictionary of Psychology (characterisation of 'interview' and 'clinical interview').",
        "World Health Organization \u2014 mhGAP Intervention Guide (respectful, collaborative communication with the person).",
        "Rehabilitation Council of India (RCI) Act 1992 (standards of professional conduct for registered professionals).",
        "Edward T. Hall \u2014 proxemics framework for interpersonal distance (zones used in the non-verbal components).",
        "{{SEC}} IACP code of conduct \u2014 secondary summaries only; primary text not accessed, so no clause numbers are quoted.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. OBJECTIVES OF THE INTERVIEW ===========================
_s("t1", "1", "Objectives of the clinical interview", [
    {"t": "lead", "text": "The clinical interview is a purposeful, goal-directed professional conversation whose objectives are to build rapport, gather history, assess and diagnose, understand the person, plan and deliver help, manage risk, and meet ethical duties such as consent and confidentiality."},
    {"t": "fig", "key": "interview_objectives", "caption": "How the clinical interview fans out from rapport into assessment, understanding, planning and risk management."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises an [[interview|a directed conversation held to gather information, assess or help]] as a **directed conversation** conducted to **obtain information or assess** a person, and a [[clinical interview|an interview used to screen, diagnose and plan treatment]] as one carried out to **screen, diagnose, formulate and plan treatment**. {{c:APA Dictionary of Psychology}}",
        "{{GUIDE}} The **World Health Organization**, in its mhGAP and related guidance, characterises good assessment as being conducted through **respectful, empathic and collaborative communication**, so that gathering clinical information and engaging the person are treated as a **single task** rather than two separate ones. {{c:WHO mhGAP}}",
        "{{THEORY}} **Wolberg** treats the interview as the **principal instrument** of psychotherapy, through which the therapist establishes a working relationship, understands the patient's dynamics and brings about change, so **interviewing skill underpins the whole treatment**. {{c:Wolberg}}",
        "{{THEORY}} Standard teaching frames the interview as differing from **ordinary conversation** by having a **purpose, structure, asymmetrical roles and professional ethics**, which is why it is planned and goal-directed rather than casual. {{c:Sethi}}",
    ]},
    {"t": "h3", "text": "The main objectives (exam form)"},
    {"t": "ul", "items": [
        "The first objective is to **establish rapport** and a working relationship, because a [[therapeutic alliance|the trusting, collaborative bond needed for honest disclosure]] makes the person feel safe enough to speak openly, and without it even expert technique yields little. {{c:Wolberg}}",
        "A second objective is to **collect information** on the presenting complaints and on the personal, family, medical, psychiatric and social history, so the clinician builds a full picture of the person and their current functioning rather than isolated symptoms. {{c:Sethi}}",
        "A third objective is to **assess and diagnose**, using the gathered data against **DSM or ICD** criteria, which lets the clinician name the problem, compare it with known conditions and choose an evidence-based plan. {{c:Sharf}}",
        "A fourth objective is to conduct the [[mental status examination|a structured snapshot of current mental functioning at interview]], so that appearance, speech, mood, thought, perception, cognition, insight and judgement are **observed and recorded systematically**. {{c:Sethi}}",
        "A fifth objective is to **understand the person's perspective**, including their feelings, beliefs, explanatory model and expectations, because treatment that ignores how the client sees their own problem usually fails to engage them. {{c:Corey}}",
        "A sixth objective is to **assess risk**, chiefly of suicide, self-harm and harm to others, since every clinical interview must screen for safety and may need to shift immediately into crisis management when danger emerges. {{c:Sethi}}",
        "A seventh objective is to **plan treatment**, which means choosing the intervention, agreeing goals, and giving the person clear information so they can make an **informed decision** about the care offered. {{c:Wolberg}}",
        "An eighth objective is to **deliver help within the interview itself**, because in a [[therapeutic interview|an interview that is itself a vehicle of change]] the skilled conversation, support and motivation are **part of the treatment**, not merely preliminaries to it. {{c:Corey}}",
        "A ninth objective is to **motivate and support** the person, and approaches such as **motivational interviewing** use the conversation itself to strengthen the person's own reasons for change rather than lecturing them into it. {{c:MET}}",
        "A tenth objective is to **evaluate progress and outcome** across later interviews, so the clinician checks whether the agreed goals are being met and adjusts the plan in the light of the person's response. {{c:Wolberg}}",
        "An objective in some settings is to **gather data for research, forensic or administrative purposes**, in which case the interviewer must be especially clear about the altered purpose, the limits of confidentiality and the use of the information. {{c:Sethi}}",
        "Running through all of these is the ethical objective of **obtaining informed consent** before the interview proceeds, and the {{c:APA 2017, Std 3.10}} **APA Ethics Code 2017** requires that consent be sought in understandable language at the outset. {{c:APA 2017, Std 3.10}}",
        "{{CODE}} A further ethical objective is to **protect confidentiality** and to explain its **limits**, because the APA Ethics Code 2017 treats safeguarding information, and disclosing only where law or safety demands, as a core duty of the interview. {{c:APA 2017, Std 4.01}}",
        "{{GUIDE}} The interview also aims to **respect the person's dignity and autonomy**, which the WHO frames as delivering assessment and care within a relationship built on **trust, respect and shared decision-making**. {{c:WHO mhGAP}}",
        "The overarching objective is that **assessment and engagement happen together**, because the same skilled, respectful conversation both gathers reliable information and builds the relationship through which any later help will be delivered. {{c:Wolberg}}",
    ]},
    {"t": "box", "kind": "key", "title": "Objectives at a glance", "items": [
        "Relationship: build rapport and a working alliance so the person speaks freely.",
        "Information: gather history and the mental status examination systematically.",
        "Assessment: diagnose, formulate and assess risk (suicide, harm).",
        "Action: plan and deliver help, motivate, and review outcome over time.",
        "Ethics: informed consent and confidentiality run through every interview.",
    ]},
    {"t": "mn", "word": "RICHER", "text": "Rapport, Information (history and MSE), Classify and diagnose, Help and plan (with motivation), Ethics (consent and confidentiality), Risk and Review \u2014 the objectives of the clinical interview."},
])

# ========================= 2. INTERVIEWING TECHNIQUES ======================
_s("t2", "2", "Interviewing techniques and the stages of an interview", [
    {"t": "lead", "text": "Interviewing technique is the deliberate way the clinician prepares, opens, conducts and closes the conversation, using a comfortable setting, a clear structure, good questioning and listening, and careful note-taking so that reliable information and a strong relationship both result."},
    {"t": "fig", "key": "interview_stages", "caption": "The stages of an interview: preparation, opening, body, closing and after the interview."},
    {"t": "ul", "items": [
        "The first technique is **preparation**, in which the clinician reviews the referral and records, arranges **privacy and comfortable seating**, protects the time from interruptions, and clarifies their own aims before the person enters. {{c:Sethi}}",
        "Seating is arranged at a **slight angle rather than directly behind a desk**, because an open, non-confrontational arrangement lowers the power distance and signals that the interview is a shared, collaborative task. {{c:Sethi}}",
        "The **opening** technique is to greet the person by name, introduce oneself and one's role, explain the **purpose, time and confidentiality**, obtain consent, and then begin with an open invitation such as asking what brings them in. {{c:Sethi}}",
        "A key early technique is **putting the person at ease**, since an anxious or guarded client discloses little, so warmth, unhurried pace and simple language in the first minutes pay off throughout the interview. {{c:Wolberg}}",
        "In the **body** of the interview the clinician explores the presenting problem for **onset, course, severity, precipitants and impact**, then moves into relevant history, keeping the person's own account at the centre. {{c:Sethi}}",
        "The main structuring technique is [[funnelling|opening with broad questions and then narrowing to specifics]], which begins with **open questions** to let the person tell their story, then narrows with focused questions to pin down detail. {{c:Sethi}}",
        "An **inverted funnel**, moving from specific closed questions toward broader ones, is used for **sensitive or shameful topics**, because starting with concrete, answerable questions eases the person into harder material. {{c:Sethi}}",
        "Throughout the body the clinician blends the **core responding skills** of clarification, paraphrasing, reflection of feeling, summarising and facilitation, which keep the person talking and confirm that they have been understood. {{c:Ivey}}",
        "A central technique is **active listening**, meaning the clinician attends fully, follows the person's topic, and resists the urge to plan the next question while the person is still speaking. {{c:Egan}}",
        "{{THEORY}} **Wolberg** stresses that the therapist's **manner**, conveyed through facial expression, gesture and tone, is itself a technique, because an attentive and accepting manner encourages disclosure as much as the words used. {{c:Wolberg}}",
        "**Note-taking** is handled so that it does not break the relationship, so the clinician records enough to be accurate while keeping **eye contact and engagement**, and explains to the person why notes are being taken. {{c:Sethi}}",
        "The **closing** technique is to **summarise, check accuracy and invite additions**, typically asking whether there is anything that has not been asked, which often surfaces important material held back until the end. {{c:Sethi}}",
        "At closing the clinician also gives **feedback or information, plans next steps and fixes the next appointment**, ending on time and on a supportive note so the person leaves feeling contained rather than abruptly dismissed. {{c:Sethi}}",
        "The final stage, **after the interview**, is to **record notes promptly, formulate, plan and reflect** on one's own reactions, because memory fades quickly and the clinician's responses can themselves be useful clinical data. {{c:Gabbard}}",
        "{{GUIDE}} Across every stage the guiding rule is that **structure serves the relationship**, so the clinician flexibly adapts pace and depth to the person while still covering the ground the interview is meant to cover. {{c:Wolberg}}",
    ]},
    {"t": "box", "kind": "key", "title": "The five stages", "items": [
        "Preparation: review records, arrange privacy and comfortable seating, clarify aims.",
        "Opening: greet, introduce, explain purpose and confidentiality, gain consent, open invitation.",
        "Body: explore problem then history; funnel from open to specific; listen and respond.",
        "Closing: summarise, check, invite additions, give information, plan, end on time.",
        "After: record promptly, formulate, plan, and reflect on your own reactions.",
    ]},
    {"t": "mn", "word": "POBCA", "text": "Preparation, Opening, Body, Closing, After \u2014 the five stages that organise interviewing technique; remember them in order to never skip consent at the start or risk-check and summary at the end."},
])

# ========================= 3. TYPES OF INTERVIEW ===================
_s("t3", "3", "Types of interview", [
    {"t": "lead", "text": "Interviews are classified by purpose, such as intake, diagnostic, case-history, mental status, crisis, therapeutic, research and forensic, and by structure and direction, ranging from structured to unstructured and from directive to non-directive, with each type suited to a particular clinical aim."},
    {"t": "fig", "key": "interview_types", "caption": "Types of interview grouped by purpose and by structure/direction."},
    {"t": "h3", "text": "Types by purpose"},
    {"t": "ul", "items": [
        "The **intake or initial interview** is the first contact, and it establishes the **reason for referral, the presenting problem, expectations, basic screening and risk**, setting the agenda for everything that follows. {{c:Sethi}}",
        "The **diagnostic interview** is aimed at arriving at a **diagnosis** by systematically checking the person's symptoms against **DSM or ICD** criteria, and it may draw on standardised instruments to do so reliably. {{c:Sharf}}",
        "The **case-history or [[anamnesis|the systematic account of the person's past and development]] interview** gathers a detailed developmental, family, educational, occupational, medical and social history, often drawing on **informants** such as parents or spouse. {{c:Sethi}}",
        "The **mental status examination** is an interview that produces a structured **snapshot of current mental functioning**, covering appearance, speech, mood and affect, thought, perception, cognition, insight and judgement. {{c:Sethi}}",
        "The **crisis interview** is **brief, focused and directive**, and its priorities are to **assess safety, stabilise the person and link them to resources**, deferring deeper exploration until the emergency has passed. {{c:Corey}}",
        "The **therapeutic interview** treats the conversation itself as a **means of change**, as in counselling and motivational interviewing, so that the skilled dialogue delivers help rather than merely collecting data. {{c:MET}}",
        "**Research interviews** apply a fixed protocol so results can be compared across people, while **forensic interviews** serve legal questions and demand special care about the altered purpose and the limits of confidentiality. {{c:Sethi}}",
        "Other recognised types include **employment or selection interviews, informant interviews** with relatives, **stress interviews** and **group interviews**, each shaped by who is present and what decision the interview must inform. {{c:Sethi}}",
    ]},
    {"t": "h3", "text": "Types by structure and direction"},
    {"t": "ul", "items": [
        "By structure, interviews run from **structured**, with fixed questions in a fixed order, through **semi-structured**, with core questions the interviewer may probe and rephrase, to **unstructured**, which is open and led by the interviewee. {{c:Sharf}}",
        "{{THEORY}} **Structured diagnostic interviews** such as the [[SCID|Structured Clinical Interview for DSM disorders]] and the brief **MINI** standardise both the questions and the scoring, which makes diagnosis far more **reliable and comparable** across clinicians. {{c:Sharf}}",
        "By direction, interviews are **directive** when the interviewer controls the topic and pace, and **non-directive** when the client leads, with the person-centred tradition after **Rogers** favouring a non-directive, following stance. {{c:Corey}}",
        "A **directive** style suits crisis work, risk assessment and structured diagnosis, whereas a **non-directive** style suits exploratory, insight-oriented and relationship-building work where the person's own narrative matters most. {{c:Corey}}",
        "The clinician usually **blends types within a single meeting**, for example opening non-directively to build rapport, then moving to a more structured, directive risk assessment when safety must be established. {{c:Sethi}}",
        "The choice of type follows the **purpose and setting**, so an epidemiological survey calls for a fully structured interview while a first psychotherapy session calls for a flexible, largely unstructured and relationship-focused one. {{c:Sharf}}",
        "Across settings the **semi-structured** interview is the most widely used compromise, because it preserves **comparability and completeness** while leaving room for rapport, probing and the person's unique concerns. {{c:Sharf}}",
    ]},
    {"t": "mn", "word": "I-DC-MART", "text": "Intake, Diagnostic, Case-history, Mental status, crisis, therApeutic, Research and forensic, plus Types by structure/direction \u2014 the families of interview grouped by purpose and by how structured and directive they are."},
])

# ========================= 4. STRUCTURED VS UNSTRUCTURED =============================
_s("t4", "4", "Characteristics of structured, semi-structured and unstructured interviews", [
    {"t": "lead", "text": "Structured interviews use fixed questions scored by rules and so are highly reliable but less flexible, unstructured interviews are open and rich but less reliable and more prone to bias, and semi-structured interviews combine the strengths of both and are preferred for most clinical and research work."},
    {"t": "fig", "key": "structure_spectrum", "caption": "The structure spectrum: reliability rises and flexibility falls as interviews become more structured."},
    {"t": "table", "caption": "Structured, semi-structured and unstructured interviews compared on the examinable features.",
     "head": ["Feature", "Structured", "Semi-structured", "Unstructured"],
     "rows": [
        ["Questions", "Fixed wording and order", "Core questions, probing allowed", "Open, led by the client"],
        ["Flexibility", "Very low", "Moderate", "Very high"],
        ["Reliability", "High (standardised, scored)", "Good with trained interviewers", "Low; depends on interviewer"],
        ["Rapport", "Lower, can feel mechanical", "Good", "High"],
        ["Examples", "DIS, CIDI, MINI", "SCID-5, SCAN, K-SADS, CAPS-5", "Clinical intake, therapy interview"],
     ]},
    {"t": "ul", "items": [
        "A **structured interview** uses **fixed question wording and order** with decision rules for scoring, which removes most interviewer variation and makes the method standardised enough for lay interviewers in large surveys. {{c:Sharf}}",
        "The great strength of structured interviews is **high reliability and comparability**, because asking everyone the same questions in the same way means differences in the answers reflect the person, not the interviewer. {{c:Sharf}}",
        "Structured interviews also give **completeness and reduced bias**, since every criterion is covered in order, so the clinician cannot accidentally skip an area or let a first impression steer the questioning. {{c:Sharf}}",
        "Their limitations are **rigidity and weaker rapport**, because the fixed script can feel **mechanical and interrogative**, may miss the person's unique concerns, and often needs cultural adaptation to fit the setting. {{c:Sharf}}",
        "Fully or highly structured instruments such as the **DIS, CIDI and MINI** are typical examples, and they are used chiefly for **research, epidemiology and standardised diagnosis** where comparability matters most. {{c:Sharf}}",
        "An **unstructured interview** is **open and flexible**, with the general topics set but the specific course determined by the interviewee, so the conversation can follow whatever the person brings as most important. {{c:Sharf}}",
        "The strengths of the unstructured interview are **flexibility, depth and rapport**, because following the person's lead allows rich, detailed material and the discovery of unexpected issues a script would never reach. {{c:Sharf}}",
        "Its weaknesses are **low reliability and vulnerability to bias**, including the [[halo effect|letting one strong impression colour the whole judgement]], first-impression error, confirmation bias and anchoring, which distort what the interviewer hears. {{c:Sethi}}",
        "Unstructured interviews are also **hard to compare across people** and can **miss symptoms** that were simply never raised, which is why they are weaker for formal diagnosis despite their clinical richness. {{c:Sharf}}",
        "The **semi-structured interview** keeps a set of **core questions** but lets the trained interviewer **probe, rephrase and follow up**, so it combines the coverage of a structure with the responsiveness of a conversation. {{c:Sharf}}",
        "Semi-structured instruments such as the **SCID-5, SCAN, K-SADS, ADIS-5 and CAPS-5** achieve **good reliability with trained clinicians** while still allowing the clinical judgement that pure structure removes. {{c:Sharf}}",
        "Because it balances **reliability against rapport and depth**, the semi-structured interview is **widely preferred** for clinical diagnosis and for most clinical research, offering the best of both extremes. {{c:Sharf}}",
        "A practical point is that **training demand rises** from the structured end, where little training is needed, through the semi-structured middle, to the unstructured end, which needs considerable **clinical skill** to do well. {{c:Sharf}}",
        "The choice therefore turns on **purpose**, with structured interviews chosen when reliable, comparable diagnosis is paramount and unstructured interviews chosen when exploration, rapport and the person's own story come first. {{c:Sharf}}",
        "{{GUIDE}} Whatever the structure, the **ethical frame is identical**, so informed consent, confidentiality and respect, as set out by the APA Ethics Code 2017, apply equally to a rigid survey and to an open therapeutic conversation. {{c:APA 2017, Std 3.10}}",
    ]},
    {"t": "mn", "word": "FRRRT", "text": "Flexibility, Reliability, Rapport, Richness and Training \u2014 the five features on which structured, semi-structured and unstructured interviews differ; reliability rises and flexibility falls as you move toward structure."},
])

# ========================= 5. MICRO-SKILLS =============================
_s("t5", "5", "Interviewing skills and the micro-skills hierarchy", [
    {"t": "lead", "text": "Interviewing skills are the small, specific, trainable behaviours, organised by Ivey as a micro-skills hierarchy built on attending and the basic listening sequence, through which the clinician shows attention, draws out the story, and communicates accurate understanding."},
    {"t": "fig", "key": "microskills_pyramid", "caption": "Ivey's micro-skills hierarchy, rising from attending through the basic listening sequence to influencing skills."},
    {"t": "ul", "items": [
        "{{THEORY}} **Ivey's micro-skills** are **small, specific communication skills** that, trained one at a time and then combined, make up competent interviewing, giving teaching and supervision a concrete vocabulary. {{c:Ivey}}",
        "The hierarchy rests on a **foundation of attending behaviour**, because unless the clinician first shows the person that they are being listened to, no later skill such as questioning or reflection will work well. {{c:Ivey}}",
        "{{THEORY}} **Egan's SOLER** captures attending as **S**quarely facing the person, an **O**pen posture, **L**eaning slightly toward them, appropriate **E**ye contact, and a **R**elaxed, natural manner. {{c:Egan}}",
        "A core listening skill is the **open-ended question**, which begins with what, how or tell me about and **invites elaboration**, giving the person control and encouraging a fuller, more spontaneous account. {{c:Ivey}}",
        "**Closed questions**, answerable with a yes, no or short fact, are used to pin down **specific detail, diagnosis and risk**, but too many in a row turn the interview into an **interrogation** that shuts the person down. {{c:Sethi}}",
        "**Encouragers**, or minimal prompts such as a nod, a mm-hmm or repeating a key word, are small facilitation skills that **keep the person talking** without diverting their flow or imposing the clinician's agenda. {{c:Ivey}}",
        "**Paraphrasing** restates the **content** of what was said in the clinician's own words, which checks accuracy and shows the person that the facts of their story have been **heard correctly**. {{c:Ivey}}",
        "**Reflection of feeling** names the **emotion beneath the content**, which demonstrates empathy, helps the person label and own their feelings, and often deepens the disclosure that follows. {{c:Corey}}",
        "**Summarising** pulls together the main themes at **transitions and at the end**, which organises the material, checks the clinician's understanding, and gives the person a chance to correct or add. {{c:Ivey}}",
        "{{THEORY}} Ivey groups the first cluster of these skills as the **basic listening sequence**, namely open and closed questions, encouragers, paraphrase, reflection of feeling and summary, which together draw out the person's story. {{c:Ivey}}",
        "Above the listening skills sit the **influencing skills**, including interpretation, feedback, self-disclosure, logical consequences and directives, which are used **later and more sparingly** once the story is understood. {{c:Ivey}}",
        "{{THEORY}} In motivational interviewing the core skills are summarised as **OARS**, namely **O**pen questions, **A**ffirmations, **R**eflections and **S**ummaries, which keep the conversation person-centred and draw out the person's own motivation. {{c:MET}}",
        "**Observation skills** sit alongside the listening skills, because the clinician reads the person's **non-verbal cues and incongruities** at the same time as hearing their words, and adjusts the next response accordingly. {{c:Ivey}}",
        "A unifying idea is **intentionality**, meaning the skilled interviewer can **choose flexibly** among many responses and switch when one is not working, rather than relying on a single habitual style. {{c:Ivey}}",
        "The overarching skill is to **sequence the micro-skills well**, so that attending and listening come first to build the story and trust, and influencing skills follow only once enough understanding and rapport are in place. {{c:Ivey}}",
    ]},
    {"t": "box", "kind": "key", "title": "The micro-skills hierarchy", "items": [
        "Foundation: attending behaviour (Egan's SOLER) and observation.",
        "Basic listening sequence: open and closed questions, encouragers, paraphrase, reflection of feeling, summary.",
        "Influencing skills (used later): interpretation, feedback, self-disclosure, directives.",
        "Motivational interviewing core: OARS (Open questions, Affirmations, Reflections, Summaries).",
        "Intentionality: choose and switch skills flexibly to fit the person.",
    ]},
    {"t": "mn", "word": "SOLER + OARS", "text": "SOLER (Squarely face, Open posture, Lean in, Eye contact, Relaxed) for attending, and OARS (Open questions, Affirmations, Reflections, Summaries) for responding \u2014 the two acronyms that carry the interviewing micro-skills."},
])

# ========================= 6. CORE RESPONDING SKILLS ===========================
_s("t6", "6", "Open-ended questions, clarification, reflection, facilitation and confrontation", [
    {"t": "lead", "text": "These are the core responding skills of the interview: open-ended questions that invite the story, clarification that makes vague statements specific, reflection that mirrors content and feeling, facilitation that keeps the person talking, and confrontation that gently points out discrepancies once rapport can bear it."},
    {"t": "fig", "key": "responding_skills", "caption": "The core responding skills, from open questions through clarification, reflection and facilitation to carefully timed confrontation."},
    {"t": "h3", "text": "Open-ended questions"},
    {"t": "ul", "items": [
        "An **open-ended question** begins with **what, how, could or tell me about** and **invites elaboration**, so instead of a yes or no the person gives a fuller, more spontaneous account in their own words. {{c:Sethi}}",
        "Open questions **encourage exploration, give control to the person and build rapport**, which is why skilled interviewers open each new area with one before narrowing down with more specific questions. {{c:Ivey}}",
        "The clinician **avoids leading, double-barrelled and stacked questions**, and is wary of **why questions**, which can sound judgemental and provoke defensiveness or intellectualisation rather than honest reflection. {{c:Sethi}}",
    ]},
    {"t": "h3", "text": "Clarification"},
    {"t": "ul", "items": [
        "**Clarification** seeks to make a **vague or ambiguous statement specific**, for example asking whether 'always' really means every day, so the clinician does not quietly assume a meaning the person never intended. {{c:Sethi}}",
        "Clarification **checks the interviewer's understanding** and signals careful listening, and it is especially important around **risk, timelines and symptoms**, where a loose word could change the whole clinical picture. {{c:Ivey}}",
    ]},
    {"t": "h3", "text": "Reflection (content and feeling)"},
    {"t": "ul", "items": [
        "**Paraphrasing, or reflection of content**, restates what the person said in the clinician's **own words**, which confirms accurate hearing and gently invites the person to continue or correct the account. {{c:Ivey}}",
        "**Reflection of feeling** names the **emotion beneath the words**, such as noting that the person sounds humiliated and angry, which demonstrates empathy and helps the person recognise and own the feeling. {{c:Corey}}",
        "{{THEORY}} **Reflection of meaning** goes a step further to mirror the person's **values and beliefs**, and in motivational interviewing well-chosen reflections are used deliberately to draw out the person's own change talk. {{c:MET}}",
        "**Summarising** is an extended reflection that **gathers several statements together** at a transition or at the end, which organises the material, checks understanding, and shows the person that their whole account has been followed. {{c:Ivey}}",
    ]},
    {"t": "h3", "text": "Facilitation (encouragers)"},
    {"t": "ul", "items": [
        "**Facilitation** uses **minimal verbal and non-verbal prompts**, such as mm-hmm, go on, a nod, an expectant pause or echoing a key word, to **keep the person talking** without steering the content. {{c:Sethi}}",
        "Facilitation works because it **rewards disclosure and shows interest**, so the person feels accompanied rather than interrogated, and it costs the clinician almost nothing in words while adding a great deal in flow. {{c:Wolberg}}",
    ]},
    {"t": "h3", "text": "Confrontation"},
    {"t": "ul", "items": [
        "**Confrontation** means **pointing out a discrepancy**, for instance between words and actions, between two statements, or between stated feeling and visible behaviour, so that a contradiction the person has avoided is gently named. {{c:Corey}}",
        "Confrontation is **not an attack**; it is a **caring, tentative, specific and gentle challenge**, offered as an observation with an invitation to explore, such as noting that the person says they are fine yet looks close to tears. {{c:Corey}}",
        "Crucially, confrontation is used **only after rapport is established**, because a premature or harsh challenge, delivered before trust can bear it, breeds **defensiveness** and can rupture the alliance. {{c:Gabbard}}",
        "Well-timed confrontation can be powerful, since it helps the person **see a blind spot** and take responsibility, but it must always leave them room to respond rather than cornering them into agreement. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Five core responding skills", "items": [
        "Open-ended questions: invite elaboration; avoid leading and 'why' questions.",
        "Clarification: make vague statements specific; check understanding.",
        "Reflection: mirror content (paraphrase), feeling and meaning to show empathy.",
        "Facilitation: minimal prompts (mm-hmm, nod, go on) to keep the person talking.",
        "Confrontation: gentle, caring challenge of discrepancies, only after rapport.",
    ]},
    {"t": "mn", "word": "OCR-FC", "text": "Open questions, Clarification, Reflection, Facilitation, Confrontation \u2014 the five core responding skills; the order also hints at timing, since open questions and gentle skills come first and confrontation comes last, only once rapport can bear it."},
])

# ========================= 7. SILENCES =========================
_s("t7", "7", "Silences in the interview", [
    {"t": "lead", "text": "Silence is an active interviewing tool, not an empty gap: different silences carry different meanings, from reflective and emotional to resistant or anxious, and the skilled clinician reads each one, waits the right amount, and responds rather than rushing to fill it."},
    {"t": "fig", "key": "silence_types", "caption": "Reading and responding to the main types of silence in the interview."},
    {"t": "ul", "items": [
        "Silence is best understood as an **active and meaningful part** of the interview, because the pause itself communicates, and a clinician who rushes to fill every gap can cut off the person's most important disclosures. {{c:Wolberg}}",
        "A **thoughtful or reflective silence** occurs when the person is **thinking or working something through**, and here the clinician's task is simply to **wait and allow it**, protecting the space for the person to continue. {{c:Sethi}}",
        "An **emotional silence** arises when the person is **overwhelmed, tearful or moved**, and the clinician responds with a warm, unhurried presence and perhaps a gentle reflection rather than a new question. {{c:Gabbard}}",
        "A **resistant silence** signals unwillingness, hostility or [[resistance|conscious or unconscious reluctance to engage with the process]], and it is explored with curiosity rather than forced, since pushing usually deepens the standoff. {{c:Gabbard}}",
        "An **anxious or awkward silence** reflects discomfort or not knowing what to say, where the clinician offers **reassurance and a gentle open invitation** to ease the person back into speaking. {{c:Sethi}}",
        "A **confused silence** means the person **does not understand or does not know how to answer**, so the clinician rephrases the question more simply rather than repeating it or letting the gap grow. {{c:Sethi}}",
        "The clinician may also use a **strategic silence of their own**, a deliberate pause that gives the person room to go deeper, because people often fill a respectful silence with the very material that matters most. {{c:Wolberg}}",
        "A practical guide is to **allow a pause of roughly five to ten seconds** before intervening, long enough to let the person think but not so long that the silence tips into feeling like rejection. {{c:Sethi}}",
        "While waiting, the clinician **observes the non-verbal signals**, since posture, facial expression, tearfulness or agitation usually reveal which kind of silence this is and therefore how to respond. {{c:Sethi}}",
        "When a response is needed, the clinician uses **encouragers, a gentle reflection such as noting that it seems hard to say, or an open invitation**, rather than firing in a fresh question that changes the subject. {{c:Sethi}}",
        "{{THEORY}} **Gabbard** notes that in psychodynamic work silence is often **rich material**, a form of resistance or communication to be understood, so the therapist tolerates it and reflects on what it may be expressing. {{c:Gabbard}}",
        "A benefit of allowing silence is that it **gives the person time to think and signals acceptance**, which frequently leads to **deeper disclosure** than a stream of questions ever would. {{c:Wolberg}}",
        "The main risk is that a **prolonged, unaddressed silence** may be experienced as **coldness or rejection**, so the clinician stays warmly present and does not simply let the person flounder. {{c:Gabbard}}",
        "Silence carries strong **cultural meaning**, so the clinician respects that comfort with pauses varies across cultures and does not read a culturally normal silence as resistance or incomprehension. {{c:Corey}}",
        "The overarching principle is to treat silence as **information to be read and used**, timing the response to the kind of silence rather than reflexively filling every gap with talk. {{c:Wolberg}}",
    ]},
    {"t": "box", "kind": "key", "title": "Reading silences", "items": [
        "Reflective: the person is thinking; wait and protect the space.",
        "Emotional: the person is overwhelmed; stay warmly present, gentle reflection.",
        "Resistant: explore with curiosity, do not force.",
        "Anxious / confused: reassure or rephrase simply; offer an open invitation.",
        "Rule of thumb: allow about 5 to 10 seconds; read non-verbals; respect culture.",
    ]},
    {"t": "mn", "word": "TEARS", "text": "Thoughtful, Emotional, Anxious, Resistant and Strategic \u2014 the main kinds of silence; read which one it is from the non-verbal signals, wait about 5 to 10 seconds, then respond rather than rushing to fill the gap."},
])

# ========================= 8. VERBAL & NON-VERBAL =========================
_s("t8", "8", "Verbal and non-verbal components of interviewing", [
    {"t": "lead", "text": "Communication in the interview runs on two channels: verbal components such as content, language and paralanguage, and non-verbal components such as facial expression, posture, interpersonal distance and touch, and the skilled clinician reads both, watches for congruence between them, and interprets patterns rather than isolated cues."},
    {"t": "fig", "key": "verbal_nonverbal", "caption": "The verbal and non-verbal channels of the interview and the importance of congruence between them."},
    {"t": "h3", "text": "Verbal components"},
    {"t": "ul", "items": [
        "The **content** is **what is actually said**, including the facts, themes, beliefs and concerns the person reports, and it is the explicit layer the clinician records and checks for accuracy. {{c:Sethi}}",
        "**Language** matters because the clinician uses **simple, clear words and avoids jargon**, borrowing the person's own vocabulary and metaphors so that explanations land and the person feels understood. {{c:Wolberg}}",
        "**Speech characteristics** such as **rate, tone, pitch, volume and fluency** are themselves clinical data, since pressured, slowed, hesitant or monotonous speech can point to mood, anxiety or thought disorder. {{c:Sethi}}",
        "[[Paralanguage|the vocal features around words, such as sighs, pauses and tone]] carries meaning beyond the words, so sighs, pauses, laughter or a tremor in the voice often reveal feeling the content alone conceals. {{c:Sethi}}",
        "The **style of questions and responses** is a verbal component too, because the balance of open and closed questions and the warmth of the clinician's wording shape how much the person discloses. {{c:Ivey}}",
    ]},
    {"t": "h3", "text": "Non-verbal components"},
    {"t": "ul", "items": [
        "[[Kinesics|communication through body movement and expression]] covers **facial expression, eye contact, gestures, posture and head nods**, together with fidgeting, tension or psychomotor change, all of which the clinician observes continuously. {{c:Sethi}}",
        "**Eye contact** is a key non-verbal cue, useful for judging **rapport, engagement and affect**, since lowered eyes may accompany depression and a shifting gaze may accompany anxiety, though its meaning varies by culture. {{c:Sethi}}",
        "{{THEORY}} [[Proxemics|the use of interpersonal distance]], described by **Edward T. Hall**, distinguishes intimate, **personal, social and public zones**, with the personal zone of roughly one and a half to four feet suiting most interviews. {{c:Hall}}",
        "[[Haptics|communication through touch]] must be handled **cautiously and only with consent**, because although a touch can convey support, it can also be intrusive or be misread, especially across gender and culture. {{c:Sethi}}",
        "**Appearance and the environment** send non-verbal messages too, so the person's grooming, and the clinician's seating, lighting, privacy and even a visible clock, all shape how safe and respected the person feels. {{c:Sethi}}",
        "[[Chronemics|the use and meaning of time]] covers **punctuality, pace and how time is used**, and starting on time, not rushing and ending on time all communicate respect and reliability to the person. {{c:Sethi}}",
        "The most important non-verbal principle is **congruence**, because non-verbal behaviour often reveals feelings the words conceal, as when a person **smiles while describing a tragedy**, flagging a mismatch worth gently exploring. {{c:Corey}}",
        "The clinician interprets **patterns, not isolated cues**, and **avoids over-reading** a single gesture, since a crossed arm or a pause means little alone but much when it fits a consistent pattern. {{c:Sethi}}",
        "**Culture strongly shapes non-verbal meaning**, so in some Indian and other contexts direct eye contact may read as disrespectful, a head wobble may signal assent, and expected interpersonal distance differs, all of which the clinician respects. {{c:Corey}}",
        "{{THEORY}} **Wolberg** underlines that the clinician's **own non-verbal behaviour** is a technique, since an attentive facial expression, a nod or a well-judged gesture conveys acceptance and encourages the person to continue. {{c:Wolberg}}",
    ]},
    {"t": "box", "kind": "key", "title": "Two channels, one message", "items": [
        "Verbal: content, language (simple, no jargon), speech characteristics, paralanguage, style.",
        "Non-verbal: kinesics, eye contact, proxemics (Hall), haptics, appearance, chronemics.",
        "Congruence: watch for words and body not matching (smiling while sad).",
        "Interpret patterns, not single cues; avoid over-reading.",
        "Always read non-verbal cues through the person's culture.",
    ]},
    {"t": "mn", "word": "CLiPS + KEPHAC", "text": "Verbal = CLiPS (Content, Language, Paralanguage, Speech characteristics, Style); non-verbal = KEPHAC (Kinesics, Eye contact, Proxemics, Haptics, Appearance, Chronemics) \u2014 read both channels and check they are congruent."},
])
