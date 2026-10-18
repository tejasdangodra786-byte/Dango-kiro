#!/usr/bin/env python3
"""Generator for the "Psychotherapy - Unit I: Introduction to Psychotherapy" study notes.

This module defines ONE in-memory content model (CONTENT) that is the single
source of truth for the study notes, plus an HTML renderer that turns that model
into a self-contained offline .html file.

FEAT-002 will add a DOCX renderer that reuses the SAME content model, so the two
deliverables stay identical. For that reason the content lives as plain Python
data (lists / dicts), never only inside an HTML string.

Integrity rules honoured here:
  * No fabricated laws, citations, clause numbers, page numbers or statistics.
  * The IACP (Indian Association of Clinical Psychologists) code of ethics was
    NOT accessible and is absent from both extracted sources, so every IACP
    mention is explicitly labelled SECONDARY / paraphrased, with no invented
    clause numbers, quotes or dates.
  * RCI (Act 1992), APA (Ethical Principles; 2006 EBP definition) and WHO
    (mhGAP, task-shifting) are supported by the backbone and stated as such.

Run with:
    PYTHONPATH=/projects/sandbox/.pylibs \
        /root/.pyenv/versions/3.11.15/bin/python3 build_unit1_notes.py
"""

from __future__ import annotations

import html
import os
from typing import Dict, List, Union

# ---------------------------------------------------------------------------
# Exact strings mandated by the task (do not alter).
# ---------------------------------------------------------------------------
TITLE = "Psychotherapy \u2014 Unit I: Introduction to Psychotherapy"  # em dash
SUBTITLE = "M.Phil. Clinical Psychology"
CREDIT = "Made by Tejas Dangodra | M.Phil. Clinical Psychology"

HERE = os.path.dirname(os.path.abspath(__file__))
HTML_OUT = os.path.join(HERE, "Psychotherapy_Unit_I_Study_Notes.html")
DOCX_OUT = os.path.join(HERE, "Psychotherapy_Unit_I_Study_Notes.docx")


# ===========================================================================
# CONTENT MODEL  (single source of truth)
# ===========================================================================
# A "block" is a dict with a "type". Supported types:
#   {"type": "bullets", "items": [str | {"text": str, "sub": [str, ...]}]}
#   {"type": "subhead", "text": str}
#   {"type": "para", "text": str}
#   {"type": "note", "title": str, "items": [str, ...]}        -> source / caution box
#   {"type": "mnemonic", "acronym": str, "expansion": str,
#        "lines": [[letter, word, meaning], ...]}
#   {"type": "flow", "title": str, "nodes": [str, ...],
#        "branches": [{"label": str, "node": str}, ...]?}      -> vertical flow chart
#
# A "section" is a dict: {"num": int, "title": str, "blocks": [block, ...]}
# ---------------------------------------------------------------------------

Block = Dict[str, Union[str, list]]
Section = Dict[str, Union[int, str, List[Block]]]


SOURCE_NOTE: Block = {
    "type": "note",
    "title": "A note on sources (read first)",
    "items": [
        "PRIMARY (used directly): the course backbone \"Psychotherapy Exam Notes\" "
        "(Unit I), which is the organised spine of these notes, plus a secondary "
        "set of typed psychotherapy notes used for supporting definitions.",
        "SUPPORTED and stated as such: RCI (Rehabilitation Council of India Act, "
        "1992), the APA Ethical Principles and the APA (2006) definition of "
        "evidence-based practice, and WHO mhGAP / task-shifting. These appear in "
        "the backbone and are presented directly.",
        "SECONDARY / paraphrased: IACP (Indian Association of Clinical "
        "Psychologists) code of ethics was NOT accessible and is absent from the "
        "extracted sources. Every IACP mention below is flagged SECONDARY and "
        "paraphrased from general knowledge only \u2014 no clause numbers, no quoted "
        "text, no dates are asserted for it.",
        "Integrity: no laws, citations, clause numbers, page numbers or statistics "
        "have been invented. Where the backbone gives figures (Lambert's "
        "percentages; Eysenck 1952; Smith & Glass 1977) they are reproduced and "
        "flagged as approximate / debated, exactly as the backbone frames them.",
    ],
}


SECTIONS: List[Section] = [
    # -------------------------------------------------------------------
    {
        "num": 1,
        "title": "Definitions of Psychotherapy",
        "blocks": [
            {"type": "subhead", "text": "Meaning and origin of the word"},
            {"type": "bullets", "items": [
                {"text": "Etymology (Kleinke, 1994): \"psychotherapy\" literally means "
                 "caring for another person's soul or being, from the Greek:", "sub": [
                    "psyche \u2014 \"soul\" or \"being\";",
                    "therapeutikos \u2014 \"caring for another\".",
                 ]},
                "Umbrella term: it covers an enormous range of interventions, "
                "modalities and integrative strategies used to improve quality of life "
                "and health.",
                "Historical root \u2014 the \"talking cure\": Anna O., an early "
                "psychoanalytic patient of Breuer, called her treatment the \"talking "
                "cure\"; an elegant if vague description \u2014 talking, expressing and "
                "sharing one's pain can itself be healing. It is a foundational frame, "
                "though too loose for contemporary researchers.",
            ]},
            {"type": "subhead", "text": "Some shorter \"favourite\" definitions "
                                        "(secondary notes; imprecise but illustrative)"},
            {"type": "bullets", "items": [
                "Korchin (1976): \"a conversation with a therapeutic purpose.\"",
                "Schofield (1964): \"the purchase of friendship.\"",
                "Bruch (1981): a situation in which two people interact and try to "
                "understand one another, with the specific goal of accomplishing "
                "something beneficial for the complaining person.",
                "J. Watkins (1983): \"when one person with an emotional disorder gets "
                "help from another person who has less of an emotional disorder.\"",
            ]},
            {"type": "subhead", "text": "Classic definitions"},
            {"type": "bullets", "items": [
                {"text": "Wolberg (1954): a form of treatment for problems of an "
                 "emotional nature, in which a trained person deliberately "
                 "establishes a professional relationship with the patient in order to:",
                 "sub": [
                    "remove, modify or retard existing symptoms;",
                    "mediate (change) disturbed patterns of behaviour;",
                    "promote positive personality growth and development.",
                 ]},
                {"text": "Jerome Frank (1961, Persuasion and Healing): a planned, "
                 "emotionally charged, confiding interaction between a trained, "
                 "socially sanctioned healer and a sufferer; aims to change the "
                 "sufferer's feelings, attitudes and behaviour."},
                {"text": "Norcross (1990): the informed and intentional application of "
                 "clinical methods and interpersonal stances derived from established "
                 "psychological principles, to help people modify behaviours, "
                 "cognitions, emotions and other characteristics in directions the "
                 "participants deem desirable."},
                {"text": "Strupp: a learning process in which a trained person uses the "
                 "relationship to help the patient overcome distress and develop more "
                 "adaptive functioning."},
            ]},
            {"type": "subhead", "text": "Key elements common to all definitions"},
            {"type": "bullets", "items": [
                "A trained, professional therapist.",
                "A help-seeking client.",
                "A psychological (non-medical) method.",
                "A planned, systematic, theory-based process.",
                "A confiding, therapeutic relationship.",
                "The aim of change (symptom relief, behaviour change, growth).",
            ]},
            {"type": "subhead", "text": "Psychotherapy vs counselling vs medical treatment"},
            {"type": "bullets", "items": [
                {"text": "Psychotherapy vs counselling:", "sub": [
                    "Counselling: usually normal developmental, situational and "
                    "adjustment problems; more present- and problem-focused; shorter.",
                    "Psychotherapy: deeper, more enduring or pathological problems and "
                    "personality reconstruction; usually longer.",
                    "The overlap is large; the distinction is mostly one of degree.",
                ]},
                {"text": "Psychotherapy vs medical treatment:", "sub": [
                    "Psychotherapy works through communication, relationship and "
                    "learning rather than drugs.",
                    "Both are often combined (e.g. for depression or psychosis).",
                ]},
                {"text": "Psychotherapy vs talking to a friend (secondary notes):", "sub": [
                    "It is a unilateral (\"one-way\") professional relationship \u2014 the "
                    "client's problems are discussed, not the therapist's needs.",
                    "It is circumscribed by limits on time, frequency, content and "
                    "intimacy.",
                    "Confidentiality is assured (within ethical and legal limits), so "
                    "the client can speak freely.",
                    "It must not be compromised by unethical dual relationships (friend, "
                    "colleague, lover).",
                ]},
            ]},
            {"type": "subhead", "text": "Four major factors psychotherapy shares "
                                        "(secondary notes)"},
            {"type": "bullets", "items": [
                "A relationship in which the client trusts that the therapist is "
                "competent and cares about their welfare.",
                "The expectancy / hope, shared by therapist and client, that benefit "
                "will result.",
                "An underlying conceptual framework giving a rationale and procedures, "
                "requiring active participation of both.",
                "A practice setting socially defined and accepted as a place of "
                "healing.",
            ]},
            {"type": "flow", "title": "Flow: from help-seeking to change", "nodes": [
                "Help-seeking client in distress",
                "Trained, socially sanctioned therapist",
                "Professional relationship established (confiding, planned)",
                "Psychological, theory-based method applied",
                "Symptom relief \u2022 behaviour change \u2022 growth",
            ]},
            {"type": "mnemonic", "acronym": "PROMPT",
             "expansion": "the six common elements of every definition",
             "lines": [
                 ["P", "Professional", "a trained professional therapist"],
                 ["R", "Relationship", "a confiding therapeutic relationship"],
                 ["O", "One seeking help", "a help-seeking client"],
                 ["M", "Method (psychological)", "a non-medical, psychological method"],
                 ["P", "Planned", "a planned, systematic, theory-based process"],
                 ["T", "Transformation", "the aim of change / growth"],
             ]},
            {"type": "flow", "title": "Flow: the four shared factors of therapy "
                                      "(secondary notes)", "nodes": [
                "A caring, competent relationship (trust)",
                "Shared expectancy / hope of benefit",
                "A conceptual framework + rationale and procedures",
                "A socially sanctioned healing setting",
                "\u2192 decrease in distress, increase in adaptive functioning",
            ]},
            {"type": "mnemonic", "acronym": "REST",
             "expansion": "the four factors common to psychotherapies (secondary notes)",
             "lines": [
                 ["R", "Relationship", "trust that the therapist is competent and cares"],
                 ["E", "Expectancy", "shared hope that benefit will result"],
                 ["S", "System of ideas", "a conceptual framework and procedures"],
                 ["T", "Therapeutic setting", "a socially sanctioned place of healing"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 2,
        "title": "Objectives (Goals) of Psychotherapy",
        "blocks": [
            {"type": "subhead", "text": "Kleinke (1994): six fundamental therapeutic "
                                        "goals (secondary notes)"},
            {"type": "bullets", "items": [
                "Overcoming demoralisation and gaining hope.",
                "Enhancing mastery and self-efficacy.",
                "Overcoming avoidance.",
                "Becoming aware of one's misconceptions.",
                "Accepting life's realities.",
                "Achieving insight.",
            ]},
            {"type": "subhead", "text": "The broader list of goals"},
            {"type": "bullets", "items": [
                "Relieve symptoms and distress (anxiety, depression, obsessions, etc.).",
                "Modify maladaptive behaviour, thoughts and emotions (change).",
                "Gain insight into the origin and meaning of problems (esp. "
                "psychodynamic work).",
                "Improve interpersonal relationships and social functioning.",
                "Strengthen coping, problem solving and adaptive defences; improve "
                "self-esteem and self-efficacy.",
                "Resolve inner conflicts and past trauma; foster emotional expression "
                "and regulation.",
                "Promote personal growth, self-actualisation and a sense of meaning.",
                "Prevent relapse and improve quality of life; in severe illness, "
                "maintain functioning and prevent hospitalisation.",
                "Concrete life outcomes (e.g. finding a job or a partner, staying out "
                "of hospital, staying alive) \u2014 the list of possible goals is "
                "effectively endless (secondary notes).",
            ]},
            {"type": "subhead", "text": "Whose goals? (goals can differ \u2014 secondary notes)"},
            {"type": "bullets", "items": [
                "Goals may differ for the patient, the significant others in the "
                "patient's life, and the psychologist (e.g. the patient wants reduced "
                "anxiety, the therapist wants increased insight, the spouse wants "
                "greater independence).",
                "Grounded in beneficence and non-maleficence: help with the problems, "
                "help the person feel better, and do no harm.",
                "Short-term symptom relief (often CBT) and deeper work (often "
                "psychodynamic) may pull in different directions; CBT may seek to "
                "reduce excess emotion while the patient wants to remove all negative "
                "emotion.",
                "Because goals and who sets them can be uncertain or in conflict, "
                "goal-setting makes informed consent in therapy genuinely complex "
                "(see Topic 6).",
            ]},
            {"type": "note", "title": "Where WHO fits (supported)", "items": [
                "WHO's rights-based, recovery-oriented view frames broader objectives: "
                "restoring functioning, participation and quality of life, not only "
                "symptom removal (consistent with mhGAP and task-shifting goals "
                "discussed under evidence-based practice).",
            ]},
            {"type": "flow", "title": "Flow: goals cascade", "nodes": [
                "Presenting distress / impairment",
                "Immediate: safety, symptom relief, engagement",
                "Intermediate: change in behaviour, cognition, emotion; insight",
                "Ultimate: growth, functioning, quality of life",
                "Maintenance: relapse prevention, follow-up",
            ]},
            {"type": "mnemonic", "acronym": "RICHER",
             "expansion": "the aims of therapy",
             "lines": [
                 ["R", "Relieve", "relieve symptoms and distress"],
                 ["I", "Insight", "insight into origin and meaning of problems"],
                 ["C", "Change", "change maladaptive behaviour, thoughts, emotions"],
                 ["H", "Harmony", "improve relationships and social functioning"],
                 ["E", "Empower", "strengthen coping, self-esteem, self-efficacy"],
                 ["R", "Relapse prevention", "prevent relapse; protect quality of life"],
             ]},
            {"type": "flow", "title": "Flow: whose goals? reconciling the parties "
                                      "(secondary notes)", "nodes": [
                "Patient's goal (e.g. reduce anxiety)",
                "Therapist's goal (e.g. insight, self-reflection)",
                "Significant others' goal (e.g. independence)",
                "Negotiate a shared, beneficence-based plan (do no harm)",
                "Clarify goals \u2192 informed consent",
            ]},
            {"type": "mnemonic", "acronym": "HE-MAMA",
             "expansion": "Kleinke's six fundamental therapeutic goals (secondary notes)",
             "lines": [
                 ["H", "Hope", "overcome demoralisation and gain hope"],
                 ["E", "Efficacy", "enhance mastery and self-efficacy"],
                 ["M", "Meet avoidance", "overcome avoidance"],
                 ["A", "Awareness", "become aware of one's misconceptions"],
                 ["M", "Make peace", "accept life's realities"],
                 ["A", "Achieve insight", "achieve insight"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 3,
        "title": "Issues in Training Professional Therapists",
        "blocks": [
            {"type": "subhead", "text": "Who may practise (India)"},
            {"type": "bullets", "items": [
                {"text": "Rehabilitation Council of India (RCI) Act, 1992 (amended "
                 "2000) regulates training and registration of clinical psychologists "
                 "(supported).", "sub": [
                    "The RCI-recognised M.Phil in Clinical Psychology (two years, with "
                    "supervised clinical practice) is the entry qualification.",
                    "Practitioners must hold a Central Rehabilitation Register (CRR) "
                    "number.",
                 ]},
                "Mental Healthcare Act (MHCA), 2017 recognises the clinical "
                "psychologist as a mental health professional.",
            ]},
            {"type": "subhead", "text": "Components of training"},
            {"type": "bullets", "items": [
                "Theoretical knowledge (psychopathology, theories of therapy, "
                "assessment).",
                "Supervised clinical practice (direct observation, audio/video "
                "review, case conferences).",
                "Personal therapy / self-awareness work (manage countertransference "
                "and blind spots).",
                "Learning ethics and law.",
                "Research literacy and evidence-based practice.",
                "Cultural competence.",
                "Lifelong learning and continuing education.",
            ]},
            {"type": "subhead", "text": "Competencies expected"},
            {"type": "bullets", "items": [
                "Relationship (alliance) skills; assessment and case formulation.",
                "Intervention skills in at least one major model.",
                "Ethical decision making; record keeping.",
                "Working with diversity; use of supervision; reflective practice.",
            ]},
            {"type": "subhead", "text": "Issues and challenges"},
            {"type": "bullets", "items": [
                "Wide variation in quality, duration and supervision of training; "
                "shortage of trained supervisors.",
                "Over-reliance on Western models with limited cultural adaptation.",
                "Whether personal therapy for trainees should be mandatory.",
                "Gate-keeping: judging suitability (empathy, maturity, integrity) and "
                "remediating or dismissing unfit trainees.",
                "Training vs licensing gap: unregulated practice by untrained persons; "
                "need to protect the public.",
                "Therapist stress, burnout, vicarious traumatisation, compassion "
                "fatigue; self-care and peer support.",
                "Boundaries and dual relationships in trainee-supervisor and "
                "small-community settings.",
                "Training in tele-therapy and its ethics.",
                "Practising within the limits of one's competence; knowing when to refer.",
            ]},
            {"type": "flow", "title": "Flow: the training pipeline", "nodes": [
                "Selection / gate-keeping (suitability, personal qualities)",
                "Theory + ethics + research literacy",
                "Supervised practice (observation, audio/video, case conferences)",
                "Personal therapy / self-awareness; cultural competence",
                "RCI registration (CRR number) \u2192 competent practice",
                "Continuing education, supervision, self-care",
            ]},
            {"type": "mnemonic", "acronym": "SCRIPT",
             "expansion": "the pillars of therapist training",
             "lines": [
                 ["S", "Supervision", "supervised clinical practice"],
                 ["C", "Competence & culture", "competencies and cultural competence"],
                 ["R", "Research & RCI", "research literacy; RCI registration"],
                 ["I", "Insight", "personal therapy / self-awareness"],
                 ["P", "Principles", "ethics and law"],
                 ["T", "Theory", "theoretical knowledge base"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 4,
        "title": "Ethical and Legal Issues in Therapy Work",
        "blocks": [
            {"type": "subhead", "text": "Core ethical principles (APA Ethical Principles)"},
            {"type": "bullets", "items": [
                "Beneficence and non-maleficence: do good; avoid harm.",
                "Fidelity and responsibility: trust, professional and scientific "
                "responsibility, accountability.",
                "Integrity: honesty, accuracy, truthfulness; no deception or "
                "exaggerated claims.",
                "Justice: fairness, equal access to quality services; no "
                "discrimination.",
                "Respect for people's rights and dignity: privacy, confidentiality, "
                "autonomy, self-determination.",
            ]},
            {"type": "note", "title": "IACP \u2014 SECONDARY / paraphrased", "items": [
                "The IACP (Indian Association of Clinical Psychologists) is the "
                "national professional body of clinical psychologists in India and is "
                "generally understood to publish a code of ethics for its members.",
                "Its code of ethics was NOT accessible and is absent from the "
                "extracted sources, so it is cited here only as SECONDARY / "
                "paraphrased general knowledge. No clause numbers, quotations or dates "
                "are asserted. Its broad themes (competence, confidentiality, consent, "
                "avoiding exploitation, professional conduct) overlap with the APA "
                "principles and the RCI code of conduct and should be read as such.",
            ]},
            {"type": "subhead", "text": "Major ethical issues"},
            {"type": "bullets", "items": [
                "Competence: practise only within the limits of training; refer when "
                "needed; keep updated.",
                "Informed consent: obtained before assessment and treatment (see "
                "Section 6).",
                "Confidentiality and its limits: private information shared only with "
                "consent; exceptions below.",
                {"text": "Dual (multiple) relationships and boundaries:", "sub": [
                    "Avoid social, business, financial or sexual relationships with "
                    "clients.",
                    "Sexual intimacy with a client (or former client, at least for a "
                    "specified period) is unethical and prohibited.",
                    "Avoid exploitation and misuse of power; manage gifts, "
                    "self-disclosure, touch and social media contact carefully.",
                ]},
                "Fees and contracts: clear and honest about fees, missed sessions and "
                "payment; do not exploit.",
                "Record keeping and data security.",
                "Termination and referral: do not abandon clients; terminate when the "
                "client no longer needs or benefits; provide referrals and handover.",
                "Cultural sensitivity and respect for diversity (gender, caste, "
                "religion, sexuality, disability, socio-economic status).",
                "Supervision and consultation: use regularly; seek help in ethical "
                "dilemmas.",
                "Advertising and claims: avoid false or unsubstantiated claims about "
                "outcomes.",
                "Research ethics: consent, no harm, confidentiality, ethics committee "
                "approval.",
            ]},
            {"type": "subhead", "text": "Ethical decision-making steps"},
            {"type": "bullets", "items": [
                "Identify the problem and the people affected.",
                "Consult codes and law; consult supervisor / colleagues.",
                "Consider options and consequences.",
                "Decide and act; document; evaluate.",
            ]},
            {"type": "note", "title": "Where ethical guidance comes from (secondary notes)",
             "items": [
                "Professional bodies' codes and guidelines define professional "
                "identity and duties.",
                "Legal statutes and case law show how courts have resolved dilemmas "
                "(illustrative, not legal advice; the law is not always an ethically "
                "sufficient guide).",
                "The therapist must still reason for themselves; the overarching duty "
                "is to make the best-quality decision that can be made.",
                "Tjelveit (1999) notes many different styles of ethical reasoning "
                "(e.g. casuistry, virtue ethics, feminist ethics, communitarianism, "
                "pragmatic, narrative) that a therapist may draw on.",
                "Classic dilemmas cluster in three domains: the goals/objectives of "
                "therapy; the boundaries between the therapist's different identities; "
                "and the social/political frameworks of practice.",
            ]},
            {"type": "subhead", "text": "The four sources of the duty to safeguard "
                                        "confidentiality (secondary notes)"},
            {"type": "bullets", "items": [
                "The ethical codes of the mental health professions.",
                "Professional licensure regulations / confidentiality and privilege "
                "statutes.",
                "The common law (case law has carved out protection for "
                "psychotherapists, though less stringent than attorney-client "
                "privilege).",
                "Constitutional guarantees (a right of privacy is inferred rather than "
                "explicit).",
                "Framed here as general / secondary professional and legal background "
                "(US-oriented), not Indian statute; for India see MHCA 2017 and "
                "Puttaswamy (2017) below. Jaffe v. Redmond (1996): the US Supreme Court "
                "recognised psychotherapist-patient communications as confidential in "
                "federal trials.",
            ]},
            {"type": "subhead", "text": "Why informed consent matters even in a "
                                        "\"talking\" treatment (secondary notes)"},
            {"type": "bullets", "items": [
                "All psychotherapies have risks and benefits the client needs to "
                "understand; benefits are often emphasised while risks are "
                "under-stated.",
                "Possible risks to disclose include: distressing transference "
                "reactions, regressive dependency, and temporary worsening ("
                "\"feeling worse before feeling better\", esp. in insight work).",
                "Greater client participation in treatment decisions lowers the risk "
                "of harmful treatment; in some jurisdictions failure to obtain consent "
                "is treated as negligence.",
            ]},
            {"type": "subhead", "text": "Boundary crossings vs boundary violations "
                                        "(secondary notes)"},
            {"type": "bullets", "items": [
                {"text": "Crossings (may be unintended, by either party):", "sub": [
                    "lateness/earliness, altered or cancelled sessions without notice;",
                    "self-disclosure (verbal or non-verbal);",
                    "discussing patient material with others (even with consent);",
                    "jokes or arguments; accidental contact outside sessions; any "
                    "physical contact.",
                ]},
                "Violations: intended and initiated by the therapist; serious, "
                "exploitative, and unethical (e.g. sexual contact).",
                "The prohibition on sexual contact traces back to the Hippocratic "
                "Oath: the professional working with a patient cannot also be that "
                "patient's lover.",
            ]},
            {"type": "subhead", "text": "Limits of confidentiality"},
            {"type": "bullets", "items": [
                "Tarasoff v. Regents of the University of California (1976) established "
                "a duty to warn and protect identifiable third parties from serious "
                "threats of violence by a client.",
                {"text": "In India:", "sub": [
                    "MHCA 2017 protects confidentiality of information about mental "
                    "illness and treatment (disclosure allowed with consent, for the "
                    "treatment, for safety, or by court order).",
                    "POCSO Act, 2012 makes reporting of known or suspected child sexual "
                    "abuse mandatory.",
                    "Right to Privacy recognised as a fundamental right by the Supreme "
                    "Court in Puttaswamy (2017).",
                ]},
            ]},
            {"type": "subhead", "text": "Legal issues in therapy work (India)"},
            {"type": "bullets", "items": [
                "RCI Act, 1992: regulates training, registration and conduct; "
                "practising without registration is an offence.",
                {"text": "Mental Healthcare Act, 2017 (rights-based law):", "sub": [
                    "Rights of persons with mental illness: access to mental "
                    "healthcare; community living; protection from cruel, inhuman or "
                    "degrading treatment; equality and non-discrimination; information; "
                    "confidentiality; access to medical records; personal contacts and "
                    "communication; legal aid; and complaints about services.",
                    "Advance directive and nominated representative.",
                    "Capacity to make treatment decisions is presumed unless shown "
                    "otherwise.",
                    "Suicide attempt presumed under severe stress and not punishable "
                    "(Section 115); the government must provide care and rehabilitation.",
                    "Restrictions on ECT (none without anaesthesia and muscle relaxant; "
                    "not for minors); prohibition of sterilisation as treatment; "
                    "limited use of restraints.",
                ]},
                "Rights of Persons with Disabilities Act, 2016 (includes mental illness "
                "and intellectual disability).",
                "POCSO Act, 2012 (mandatory reporting) and the Juvenile Justice Act "
                "(care of children).",
                "Consumer Protection Act and civil/criminal law on negligence "
                "(professional malpractice); courts may seek records or testimony.",
                "Tele-mental health: Telemedicine Practice Guidelines (2020) and "
                "tele-psychiatry guidance apply.",
                "Legal vs ethical: law sets the minimum enforceable standard; ethics "
                "sets the higher aspirational standard.",
            ]},
            {"type": "flow", "title": "Confidentiality decision flow", "nodes": [
                "Client discloses information",
                "Default: keep confidential",
            ], "branches": [
                {"label": "Client gives valid consent to share", "node": "Disclose only what is agreed"},
                {"label": "Serious, identifiable threat to a third party (Tarasoff)", "node": "Duty to warn / protect"},
                {"label": "Risk to self / safety", "node": "Act to protect; least disclosure needed"},
                {"label": "Child abuse known/suspected (POCSO)", "node": "Mandatory report"},
                {"label": "Court order", "node": "Disclose as legally required"},
            ]},
            {"type": "mnemonic", "acronym": "FIRBJ",
             "expansion": "the APA ethical principles",
             "lines": [
                 ["F", "Fidelity", "fidelity and responsibility"],
                 ["I", "Integrity", "honesty, accuracy, no deception"],
                 ["R", "Respect", "respect for rights and dignity"],
                 ["B", "Beneficence", "beneficence and non-maleficence"],
                 ["J", "Justice", "fairness and equal access"],
             ]},
            {"type": "mnemonic", "acronym": "CCS-PC",
             "expansion": "when confidentiality may be breached",
             "lines": [
                 ["C", "Consent", "client gives valid consent"],
                 ["C", "Child abuse", "mandatory reporting under POCSO"],
                 ["S", "Safety / self-harm", "risk to self or others"],
                 ["P", "Protect third party", "Tarasoff duty to warn / protect"],
                 ["C", "Court order", "legally compelled disclosure"],
             ]},
            {"type": "mnemonic", "acronym": "ELCC",
             "expansion": "the four sources of the duty to safeguard confidentiality "
                          "(secondary notes)",
             "lines": [
                 ["E", "Ethical codes", "the mental health professions' ethics codes"],
                 ["L", "Licensure / law", "licensure regulations and privilege statutes"],
                 ["C", "Common law", "case law protecting therapist-patient communication"],
                 ["C", "Constitution", "an inferred constitutional right of privacy"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 5,
        "title": "Rights and Responsibilities in Psychotherapy",
        "blocks": [
            {"type": "subhead", "text": "Client rights"},
            {"type": "bullets", "items": [
                "Right to competent, respectful, non-discriminatory treatment.",
                "Right to informed consent and to adequate information about therapy.",
                "Right to confidentiality (within its legal limits).",
                "Right to autonomy and self-determination, including the right to "
                "withdraw or refuse treatment without penalty.",
                "Right to privacy and dignity; right to safety (no exploitation or "
                "harm).",
                "Right to access one's records (consistent with MHCA 2017) and to make "
                "complaints about services.",
                "Right to a clear contract (fees, frequency, duration, cancellation).",
            ]},
            {"type": "subhead", "text": "Client responsibilities"},
            {"type": "bullets", "items": [
                "Attend sessions, be punctual, give notice of cancellation.",
                "Participate actively and honestly; complete agreed homework / tasks.",
                "Pay agreed fees; respect the therapeutic frame and boundaries.",
            ]},
            {"type": "subhead", "text": "Therapist responsibilities"},
            {"type": "bullets", "items": [
                "Practise competently and within the limits of training; refer when "
                "needed.",
                "Obtain valid informed consent; protect confidentiality and its "
                "limits.",
                "Maintain boundaries; avoid dual relationships and exploitation.",
                "Keep accurate, secure records; honour reporting obligations.",
                "Provide continuity of care; avoid abandonment; arrange referral and "
                "handover.",
                "Use supervision; engage in self-care; keep up to date (continuing "
                "education).",
            ]},
            {"type": "subhead", "text": "Confidentiality as a right (secondary notes)"},
            {"type": "bullets", "items": [
                "Confidentiality may be understood as the client's right of \"informed "
                "consent to disclosure\": the therapist has an ethical duty to obtain "
                "the client's consent before disclosing treatment details to anyone, "
                "including those close to the client.",
                "The right can create dilemmas when the client discloses that a third "
                "party is at risk (see Topic 4 limits), or when family/employers seek "
                "information or ask the therapist not to tell the client about their "
                "contact.",
                {"text": "Use of case material for teaching / research is a separate, "
                 "specific consent issue:", "sub": [
                    "rising emphasis on client autonomy and ownership of personal "
                    "material;",
                    "some journals require the client to have read and consented to "
                    "the write-up;",
                    "consent sought after therapy ends can itself be intrusive.",
                 ]},
            ]},
            {"type": "note", "title": "Codes behind these rights/duties", "items": [
                "APA Ethical Principles and the RCI code of conduct underpin these "
                "rights and responsibilities (supported).",
                "IACP themes are broadly consistent (SECONDARY / paraphrased \u2014 no "
                "clause numbers asserted).",
            ]},
            {"type": "flow", "title": "Flow: rights and duties in balance", "nodes": [
                "Contract sets the frame (fees, time, boundaries)",
                "Client rights <-> Therapist duties (consent, confidentiality, competence)",
                "Shared responsibilities (attendance, honesty, participation)",
                "Breach? \u2192 repair, refer, or complaint / review mechanism",
            ]},
            {"type": "flow", "title": "Flow: a request to disclose client information "
                                      "(secondary notes)", "nodes": [
                "Someone requests information about the client / therapy",
                "Default duty: do not disclose without consent",
            ], "branches": [
                {"label": "Client consents to share agreed material", "node": "Disclose only what is agreed"},
                {"label": "Third party at risk of serious harm", "node": "Weigh duty to protect vs confidentiality"},
                {"label": "Teaching / research / publication", "node": "Separate specific consent; anonymise"},
                {"label": "No consent and no overriding duty", "node": "Maintain confidentiality"},
            ]},
            {"type": "mnemonic", "acronym": "CARED",
             "expansion": "what the client is owed",
             "lines": [
                 ["C", "Consent & confidentiality", "informed consent; confidentiality"],
                 ["A", "Autonomy", "self-determination; right to withdraw"],
                 ["R", "Respect & records", "dignity; access to own records"],
                 ["E", "Equality", "non-discriminatory, competent care"],
                 ["D", "Dignity & safety", "privacy; no exploitation or harm"],
             ]},
            {"type": "mnemonic", "acronym": "PROTECT",
             "expansion": "what the therapist is responsible for",
             "lines": [
                 ["P", "Practise competently", "within limits of training; refer when needed"],
                 ["R", "Records", "accurate, secure record keeping"],
                 ["O", "Obtain consent", "valid informed consent"],
                 ["T", "Trust / confidentiality", "protect confidentiality and its limits"],
                 ["E", "Ethical boundaries", "avoid dual relationships and exploitation"],
                 ["C", "Continuity", "no abandonment; arrange referral / handover"],
                 ["T", "Training kept up", "supervision, self-care, continuing education"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 6,
        "title": "Consent and Assent (Including Minors)",
        "blocks": [
            {"type": "para", "text": "Informed consent is the client's voluntary "
             "agreement to assessment or therapy after receiving and understanding "
             "adequate information. Consent is a process, not a one-time signature."},
            {"type": "subhead", "text": "The principle behind consent (secondary notes)"},
            {"type": "bullets", "items": [
                "The underpinning ethical principle is respect for autonomy: clients "
                "should be free to choose or refuse treatment for themselves.",
                {"text": "For consent to be legally valid it must be:", "sub": [
                    "voluntary;",
                    "given by a client competent to make that decision;",
                    "adequately informed, including about possible negative "
                    "side-effects.",
                ]},
                "A specific difficulty in therapy: the therapist often cannot predict "
                "exactly what the client will experience or what the outcome will be, "
                "so it is good practice to warn that the process is not always "
                "comfortable and that many clients feel worse before they feel better "
                "(less so in CBT).",
            ]},
            {"type": "subhead", "text": "Elements of informed consent"},
            {"type": "bullets", "items": [
                "Capacity (competence): ability to understand, appreciate, reason and "
                "express a choice.",
                {"text": "Disclosure of information:", "sub": [
                    "purpose and nature of therapy; methods; expected benefits;",
                    "possible risks and discomfort; alternatives (including no "
                    "treatment);",
                    "duration; fees; confidentiality and its limits;",
                    "recording / observation; supervision; therapist's qualifications;",
                    "the right to withdraw without penalty.",
                ]},
                "Comprehension: simple, understandable language; check understanding.",
                "Voluntariness: free from coercion or undue pressure.",
                "Documentation: written (or recorded) consent, revisited when the plan "
                "changes.",
            ]},
            {"type": "subhead", "text": "Special issues"},
            {"type": "bullets", "items": [
                {"text": "Minors (<18 years in India):", "sub": [
                    "Legal consent comes from the parent or legal guardian.",
                    "The child's assent (affirmative, informed agreement to the extent "
                    "of the child's developing capacity) is also sought.",
                    "Assent respects autonomy, increases cooperation and strengthens "
                    "the alliance.",
                    "Guidelines (for example ICMR 2017) generally require assent for "
                    "children roughly 7-18 years, with simpler language for younger "
                    "children and more weight to the wishes of older, mature "
                    "adolescents.",
                    "If a child refuses, explore concerns and, unless safety requires "
                    "otherwise, respect it.",
                ]},
                "Adults lacking capacity: involve the nominated representative / legal "
                "guardian under MHCA 2017; still seek the person's views; use the least "
                "restrictive approach.",
                "Separate, specific consent for recording, observation, supervision, "
                "research or publication; case material is anonymised.",
                "Emergency (immediate risk to life) can justify action without full "
                "consent, with documentation.",
                "Cultural context: family involvement is common in collectivist "
                "settings; still ensure the client's own consent and confidentiality.",
            ]},
            {"type": "subhead", "text": "Capacity and voluntariness can fluctuate "
                                        "(secondary notes)"},
            {"type": "bullets", "items": [
                "The problem is usually not a global loss of competence but fluctuating "
                "capacity or rapidly changing decisions (e.g. recent trauma or "
                "bereavement, psychotic symptoms).",
                "A child who agreed to therapy one day and refuses the next may be "
                "making an informed refusal \u2014 or showing that the work has begun; "
                "this needs careful judgement.",
                {"text": "Voluntariness is strained when clients are pressured or "
                 "\"coerced\":", "sub": [
                    "family-pushed children or family members who agree only under "
                    "pressure;",
                    "mandated / forensic clients (therapy as an alternative to prison "
                    "or part of detention) \u2014 handled openly as part of the work "
                    "from the start.",
                ]},
            ]},
            {"type": "flow", "title": "Consent and assent decision flow", "nodes": [
                "Is the client an adult with capacity?",
            ], "branches": [
                {"label": "Adult, has capacity", "node": "Informed consent from the client"},
                {"label": "Minor (<18)", "node": "Guardian consent + child's assent"},
                {"label": "Adult lacking capacity", "node": "Nominated representative (MHCA) + seek the person's views"},
                {"label": "Emergency, risk to life", "node": "Act in best interests; document; obtain consent as soon as possible"},
                {"label": "Refusal by a minor", "node": "Explore concerns; respect unless safety requires otherwise"},
            ]},
            {"type": "mnemonic", "acronym": "CIVIC",
             "expansion": "the elements of valid consent",
             "lines": [
                 ["C", "Capacity", "capacity / competence to decide"],
                 ["I", "Information", "full disclosure of relevant information"],
                 ["V", "Voluntariness", "free from coercion or undue pressure"],
                 ["I", "Insight / comprehension", "the client understands the information"],
                 ["C", "Consent documented", "documented and revisited as a process"],
             ]},
            {"type": "mnemonic", "acronym": "AVID",
             "expansion": "the test for valid consent (secondary notes)",
             "lines": [
                 ["A", "Autonomy", "respects the client's right to choose or refuse"],
                 ["V", "Voluntary", "free from coercion or undue pressure"],
                 ["I", "Informed", "adequately informed, including of side-effects"],
                 ["D", "Decision-capable", "given by a competent client"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 7,
        "title": "Planning and Recording of Therapy Sessions",
        "blocks": [
            {"type": "subhead", "text": "Planning"},
            {"type": "bullets", "items": [
                "Intake and assessment: presenting complaints, history, risk "
                "assessment, mental status, test results, diagnosis.",
                {"text": "Case formulation: an integrated hypothesis explaining the "
                 "problem using the \"4 Ps\":", "sub": [
                    "Predisposing factors;",
                    "Precipitating factors;",
                    "Perpetuating factors;",
                    "Protective factors.",
                 ]},
                "Treatment plan: goals, chosen therapy, frequency (usually weekly), "
                "session length (usually 45-50 minutes), expected duration, measures of "
                "progress, emergency plan, criteria for termination.",
                "Therapeutic frame: time, place, fee, boundaries, cancellation policy, "
                "contact between sessions.",
                "Session planning: prepare using previous notes; plan aims per session "
                "but stay responsive to the client's needs.",
            ]},
            {"type": "subhead", "text": "Structure of a typical session (e.g. CBT)"},
            {"type": "bullets", "items": [
                "Brief check-in and mood check.",
                "Bridge from the previous session.",
                "Agenda setting (collaboratively).",
                "Review of homework.",
                "Work on agenda items (main therapeutic work).",
                "Capsule summaries during the session.",
                "Assign new homework.",
                "Feedback from the client on the session.",
            ]},
            {"type": "subhead", "text": "Recording"},
            {"type": "bullets", "items": [
                "Types of records: intake / case history; assessment reports and test "
                "data; treatment plan; consent forms; progress notes; process notes "
                "(private reflections, kept separate); risk assessments; referral and "
                "termination summaries; correspondence.",
                {"text": "Formats of progress notes:", "sub": [
                    "SOAP: Subjective, Objective, Assessment, Plan;",
                    "DAP: Data, Assessment, Plan;",
                    "BIRP: Behaviour, Intervention, Response, Plan.",
                ]},
                "Good records are timely, accurate, objective, factual, brief but "
                "complete, free of judgemental language, signed and dated.",
                "Purposes: continuity of care, monitoring progress, supervision, "
                "communication with other professionals, evidence in legal/ethical "
                "review, research and quality improvement.",
                "Security and confidentiality: secure storage (locked cabinets, "
                "password-protected files), restricted access, coded identifiers; "
                "retention per law or policy; safe disposal.",
                "Audio/video recording needs explicit written consent; used only for "
                "the agreed purpose and destroyed as agreed.",
            ]},
            {"type": "subhead", "text": "Progress notes vs process notes (secondary notes)"},
            {"type": "bullets", "items": [
                "A common practice: note briefly in the medical record that the "
                "session took place, and keep short notes of the session's main themes "
                "separately.",
                "Process notes (the therapist's private reflections) are generally not "
                "shared with the client; reading them is unlikely to help the client.",
                "The therapist does not \"own\" the notes outright: in some "
                "jurisdictions the client may have legal access to them unless access "
                "would harm the client or another person.",
                "In CBT, therapist and client often share records of thoughts, feelings "
                "and behaviours (as with homework); caution where material could be "
                "misread out of context.",
            ]},
            {"type": "flow", "title": "Flow: case formulation \u2014 the 4 Ps", "nodes": [
                "Predisposing factors (vulnerability)",
                "Precipitating factors (triggers)",
                "Perpetuating factors (what maintains it)",
                "Protective factors (strengths / supports)",
                "\u2192 integrated formulation drives the treatment plan",
            ]},
            {"type": "flow", "title": "Flow: a typical session", "nodes": [
                "Check-in + mood check",
                "Bridge from previous session",
                "Collaborative agenda setting",
                "Review homework",
                "Work on agenda items",
                "Summaries + new homework",
                "Client feedback \u2192 write progress note (SOAP/DAP/BIRP)",
            ]},
            {"type": "mnemonic", "acronym": "SOAP",
             "expansion": "the classic progress-note format",
             "lines": [
                 ["S", "Subjective", "the client's reported experience"],
                 ["O", "Objective", "observed, measurable data"],
                 ["A", "Assessment", "clinical interpretation / progress"],
                 ["P", "Plan", "next steps, homework, plan for next session"],
             ]},
            {"type": "mnemonic", "acronym": "4 Ps",
             "expansion": "the pillars of a case formulation",
             "lines": [
                 ["P", "Predisposing", "long-standing vulnerability factors"],
                 ["P", "Precipitating", "triggers that set off the problem"],
                 ["P", "Perpetuating", "factors that maintain the problem"],
                 ["P", "Protective", "strengths and supports to build on"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 8,
        "title": "Structuring Therapy and Setting Goals",
        "blocks": [
            {"type": "bullets", "items": [
                "Structure gives therapy direction, safety and efficiency: a clear "
                "frame (contract), a clear agenda each session, phases, and regular "
                "review of progress.",
                {"text": "Phases of therapy:", "sub": [
                    "Beginning: engagement, assessment, formulation, contracting, goal "
                    "setting.",
                    "Middle: active intervention and working through.",
                    "End: consolidation, relapse prevention, termination, follow-up.",
                ]},
                "Goal setting: collaborative (jointly negotiated), derived from the "
                "case formulation, expressed in the client's own words.",
                "Because patient, therapist and significant others may hold different "
                "goals (see Topic 2), negotiating and clarifying goals is part of "
                "structuring therapy and of valid informed consent.",
            ]},
            {"type": "subhead", "text": "Characteristics of good goals (SMART)"},
            {"type": "bullets", "items": [
                "Specific; Measurable; Achievable; Realistic/Relevant; Time-bound.",
                "Also stated positively, under the client's control, and ordered by "
                "priority.",
            ]},
            {"type": "subhead", "text": "Types of goals and measurement"},
            {"type": "bullets", "items": [
                "Ultimate vs intermediate vs immediate (session) goals.",
                "Outcome goals (symptom relief, behaviour change) vs process goals "
                "(insight, alliance, skills).",
                "Short-term vs long-term.",
                "Goal Attainment Scaling (Kiresuk and Sherman, 1968): each goal scaled "
                "from -2 (much worse) to +2 (much better than expected) for "
                "individualised measurement.",
            ]},
            {"type": "subhead", "text": "Why structure and goals matter"},
            {"type": "bullets", "items": [
                "Goal consensus is a component of the therapeutic alliance (Bordin) "
                "and predicts outcome.",
                "Clear goals increase motivation and signal when therapy is complete.",
                "Goals are revisited and revised as new information or progress "
                "emerges.",
            ]},
            {"type": "flow", "title": "Flow: structuring the course of therapy", "nodes": [
                "Beginning: engage, assess, formulate, contract",
                "Set collaborative SMART goals (from formulation)",
                "Middle: intervene, work through, scale progress (GAS)",
                "Review and revise goals",
                "End: consolidate, relapse prevention, terminate, follow-up",
            ]},
            {"type": "mnemonic", "acronym": "SMART",
             "expansion": "the characteristics of good goals",
             "lines": [
                 ["S", "Specific", "clearly defined"],
                 ["M", "Measurable", "progress can be measured"],
                 ["A", "Achievable", "attainable for the client"],
                 ["R", "Realistic / Relevant", "meaningful and feasible"],
                 ["T", "Time-bound", "has a time frame"],
             ]},
            {"type": "flow", "title": "Flow: the three phases of therapy", "nodes": [
                "Beginning: engage, assess, formulate, contract, set goals",
                "Middle: active intervention and working through",
                "End: consolidate, relapse prevention, terminate, follow-up",
            ]},
            {"type": "mnemonic", "acronym": "OPIS",
             "expansion": "the ways goals are classified",
             "lines": [
                 ["O", "Outcome vs process", "symptom/behaviour change vs insight/alliance/skills"],
                 ["P", "Priority-ordered", "ranked by importance"],
                 ["I", "Immediate/intermediate/ultimate", "session, mid-course and final goals"],
                 ["S", "Short vs long term", "time horizon of the goal"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 9,
        "title": "Pre- and Post-Assessment",
        "blocks": [
            {"type": "subhead", "text": "Purpose"},
            {"type": "bullets", "items": [
                "Pre-assessment (baseline): diagnosis, severity, risk, strengths, and "
                "a benchmark.",
                "Post-assessment: evaluates change after therapy; follow-up checks "
                "maintenance.",
                "Ongoing session-by-session monitoring (e.g. OQ-45, CORE-OM, Outcome "
                "Rating Scale / Session Rating Scale) gives feedback so therapy can be "
                "adjusted (feedback-informed treatment); reduces deterioration, "
                "especially in clients \"not on track\".",
            ]},
            {"type": "subhead", "text": "Methods"},
            {"type": "bullets", "items": [
                "Structured and unstructured interviews; behavioural observation.",
                "Self-report scales; clinician-rated scales; projective or other tests "
                "when relevant.",
                "Collateral information; physiological measures; client diaries / "
                "self-monitoring.",
            ]},
            {"type": "subhead", "text": "Commonly used measures (from the backbone)"},
            {"type": "bullets", "items": [
                "Depression: BDI-II, PHQ-9, HAM-D.",
                "Anxiety: BAI, GAD-7, HAM-A.",
                "OCD: Y-BOCS. PTSD: PCL-5 / CAPS-5.",
                "General distress: SCL-90-R, GHQ.",
                "Well-being / quality of life: WHO-5, WHOQOL-BREF.",
                "Functioning: GAF / WHODAS 2.0. Global: CGI.",
                "Target-complaint ratings; therapeutic alliance (WAI).",
            ]},
            {"type": "subhead", "text": "Evaluating change"},
            {"type": "bullets", "items": [
                "Statistical significance (group level) vs clinical significance "
                "(individual level).",
                "Reliable Change Index (RCI; Jacobson and Truax, 1991): the change must "
                "exceed measurement error (RCI > 1.96).",
                "Clinically significant change: the client crosses a cutoff between the "
                "dysfunctional and functional populations.",
                "Effect size (Cohen's d) and percentage improvement.",
                "Single-case designs (AB, ABAB, multiple baseline) for individual "
                "evaluation.",
            ]},
            {"type": "note", "title": "Two different RCIs \u2014 do not confuse them", "items": [
                "RCI (Act): Rehabilitation Council of India Act, 1992 \u2014 the regulator "
                "(Sections 3-4-10).",
                "RCI (statistic): Reliable Change Index (Jacobson & Truax, 1991) \u2014 "
                "used here to judge whether an individual's change is reliable (> 1.96).",
            ]},
            {"type": "subhead", "text": "Good practice"},
            {"type": "bullets", "items": [
                "Use reliable and valid measures suited to the client's language, "
                "culture and literacy.",
                "Use the same measures before and after.",
                "Interpret with caution (placebo, regression to the mean, practice "
                "effects, response bias).",
                "Share results with the client and use them to plan further care.",
            ]},
            {"type": "flow", "title": "Flow: measuring change across therapy", "nodes": [
                "Pre-assessment (baseline): diagnosis, severity, risk, strengths",
                "Session-by-session monitoring (feedback-informed)",
                "Post-assessment with the SAME measures",
                "Reliable change? RCI > 1.96",
            ], "branches": [
                {"label": "Yes + crosses clinical cutoff", "node": "Clinically significant change"},
                {"label": "No / partial", "node": "Revise plan; continue or refer"},
            ]},
            {"type": "mnemonic", "acronym": "BASE",
             "expansion": "the logic of pre/post assessment",
             "lines": [
                 ["B", "Baseline", "pre-assessment benchmark"],
                 ["A", "Assess ongoing", "session-by-session monitoring"],
                 ["S", "Same measures", "same valid tools before and after"],
                 ["E", "Evaluate change", "RCI, clinical significance, effect size"],
             ]},
            {"type": "flow", "title": "Flow: judging whether change is real", "nodes": [
                "Pre vs post scores on the same measure",
                "Statistical significance (group level)?",
                "Reliable Change Index > 1.96 (individual level)?",
                "Crosses the clinical cutoff (dysfunctional \u2192 functional)?",
                "\u2192 reliable AND clinically significant change",
            ]},
            {"type": "mnemonic", "acronym": "SCORE",
             "expansion": "sound assessment methods",
             "lines": [
                 ["S", "Self-report", "self-report scales and diaries"],
                 ["C", "Clinician-rated", "clinician-administered scales"],
                 ["O", "Observation", "behavioural observation"],
                 ["R", "Reports / collateral", "collateral and physiological information"],
                 ["E", "Examine / interview", "structured and unstructured interviews"],
             ]},
        ],
    },
    # -------------------------------------------------------------------
    {
        "num": 10,
        "title": "Practice of Evidence-Based Therapies (EBP)",
        "blocks": [
            {"type": "subhead", "text": "Definition and origins"},
            {"type": "bullets", "items": [
                "APA Presidential Task Force (2006): EBP is the integration of the best "
                "available research with clinical expertise in the context of patient "
                "characteristics, culture and preferences (supported).",
                "The concept came from evidence-based medicine (Sackett).",
            ]},
            {"type": "subhead", "text": "Empirically supported treatments (ESTs)"},
            {"type": "bullets", "items": [
                {"text": "Division 12 (Society of Clinical Psychology, APA); Chambless "
                 "and Hollon (1998) criteria:", "sub": [
                    "well-established (\u2265 2 good between-group studies by independent "
                    "researchers, or a large series of single-case studies, showing "
                    "superiority to placebo or another treatment, with manuals and "
                    "clear samples);",
                    "probably efficacious.",
                 ]},
            ]},
            {"type": "subhead", "text": "Hierarchy of evidence"},
            {"type": "bullets", "items": [
                "Systematic reviews / meta-analyses of RCTs; then RCTs; controlled "
                "trials without randomisation; cohort and case-control; single-case "
                "studies; expert opinion.",
                "Efficacy (controlled conditions) vs effectiveness (real-world) vs "
                "clinical utility (feasible, acceptable, cost-effective).",
            ]},
            {"type": "subhead", "text": "Examples of well-supported treatments"},
            {"type": "bullets", "items": [
                "CBT for depression, panic, social anxiety, GAD.",
                "ERP for OCD; prolonged exposure, cognitive processing therapy and EMDR "
                "for PTSD.",
                "Behavioural activation and IPT for depression; DBT for borderline "
                "personality disorder.",
                "Family-based treatment for adolescent anorexia; parent management "
                "training for child conduct problems; CBT for insomnia.",
                "Social skills training, family psychoeducation and cognitive "
                "remediation for schizophrenia.",
            ]},
            {"type": "subhead", "text": "Guidelines and scaling up (WHO / India)"},
            {"type": "bullets", "items": [
                "NICE (UK), APA, and WHO mhGAP recommend specific treatments "
                "(supported).",
                "In India, culturally adapted low-intensity interventions delivered by "
                "lay counsellors (e.g. the Healthy Activity Programme in the PREMIUM "
                "trials in Goa, and the Thinking Healthy Programme) show evidence-based "
                "care can be scaled up (task-shifting).",
            ]},
            {"type": "subhead", "text": "Steps of EBP in the clinic"},
            {"type": "bullets", "items": [
                "Ask a clear clinical question.",
                "Search for the best evidence.",
                "Appraise it.",
                "Integrate it with clinical expertise and the client's values.",
                "Apply; evaluate the outcome.",
            ]},
            {"type": "subhead", "text": "The evidence debate and common factors"},
            {"type": "bullets", "items": [
                "Eysenck (1952) claimed psychotherapy was no more effective than "
                "spontaneous remission.",
                "Smith and Glass (1977) meta-analysed 375 studies and found a mean "
                "effect size of about 0.68 (the average treated person better off than "
                "about 75% of untreated persons).",
                "The \"Dodo bird verdict\" (Rosenzweig; Luborsky): all bona fide "
                "therapies work about equally, pointing to common factors.",
                {"text": "Lambert (1992) estimated contributions (approximate and "
                 "debated, as the backbone frames them):", "sub": [
                    "extratherapeutic / client factors ~40%;",
                    "relationship ~30%;",
                    "expectancy ~15%;",
                    "technique ~15%.",
                 ]},
                "Balance: both specific (model-based) and common (relationship) factors "
                "matter; evidence-based techniques work best within a strong "
                "relationship.",
            ]},
            {"type": "subhead", "text": "The common factors (backbone)"},
            {"type": "bullets", "items": [
                "Therapeutic relationship (alliance, empathy).",
                "Client expectations and hope (placebo).",
                "A healing setting and a credible rationale.",
                "A credible ritual / set of procedures.",
                "New learning experiences; a sense of mastery.",
                "Wampold's contextual model argues these common factors, rather than "
                "specific techniques, explain most of the benefit of therapy.",
            ]},
            {"type": "subhead", "text": "Criticisms and challenges of EBP"},
            {"type": "bullets", "items": [
                "RCT samples often unrepresentative (strict exclusion, no "
                "comorbidity); limited generalisability to diverse and Indian "
                "populations.",
                "Manuals can reduce flexibility; researcher allegiance effects; "
                "publication bias.",
                "Under-emphasis of common factors and the therapist's contribution.",
                "Outcome focus on symptom measures rather than broader well-being or "
                "meaning.",
                "Practice-based evidence and routine outcome monitoring complement "
                "RCTs; cultural adaptation and task shifting are important in India.",
            ]},
            {"type": "flow", "title": "Flow: EBP in the clinic (Ask-Acquire-Appraise-Apply-Assess)", "nodes": [
                "Ask a focused clinical question",
                "Acquire the best available research evidence",
                "Appraise the evidence (hierarchy: meta-analysis > RCT > ...)",
                "Integrate with clinical expertise + client values/culture",
                "Apply the intervention",
                "Assess the outcome (monitor, revise)",
            ]},
            {"type": "mnemonic", "acronym": "5 A's",
             "expansion": "the steps of evidence-based practice",
             "lines": [
                 ["A", "Ask", "ask a clear clinical question"],
                 ["A", "Acquire", "search for the best evidence"],
                 ["A", "Appraise", "critically appraise the evidence"],
                 ["A", "Apply", "integrate with expertise + client values, then apply"],
                 ["A", "Assess", "evaluate and revise the outcome"],
             ]},
            {"type": "flow", "title": "Flow: what accounts for therapy's benefit "
                                      "(Lambert's estimates)", "nodes": [
                "Client / extratherapeutic factors (~40%)",
                "The therapeutic relationship (~30%)",
                "Expectancy / hope (~15%)",
                "Technique (~15%)",
                "\u2192 approximate and debated, as the backbone frames them",
            ]},
            {"type": "mnemonic", "acronym": "RICH-M",
             "expansion": "the common factors across therapies (backbone)",
             "lines": [
                 ["R", "Relationship", "alliance and empathy"],
                 ["I", "Investment / hope", "client expectations and hope (placebo)"],
                 ["C", "Credible rationale", "a healing setting and rationale"],
                 ["H", "Healing ritual", "a credible ritual / set of procedures"],
                 ["M", "Mastery & learning", "new learning experiences and mastery"],
             ]},
        ],
    },
]


# ===========================================================================
# HTML RENDERER
# ===========================================================================

def _esc(text: str) -> str:
    return html.escape(str(text), quote=True)


def _render_bullets(items: list) -> str:
    out = ["<ul>"]
    for item in items:
        if isinstance(item, dict):
            out.append("<li>" + _esc(item["text"]))
            sub = item.get("sub")
            if sub:
                out.append("<ul>")
                for s in sub:
                    out.append("<li>" + _esc(s) + "</li>")
                out.append("</ul>")
            out.append("</li>")
        else:
            out.append("<li>" + _esc(item) + "</li>")
    out.append("</ul>")
    return "\n".join(out)


def _render_flow(block: Block) -> str:
    out = ['<div class="flow avoid-break">']
    out.append('<div class="flow-title">Flow chart: ' + _esc(block["title"]) + "</div>")
    nodes = block.get("nodes", [])
    for i, node in enumerate(nodes):
        out.append('<div class="flow-node">' + _esc(node) + "</div>")
        if i < len(nodes) - 1:
            out.append('<div class="flow-arrow">&#8595;</div>')
    branches = block.get("branches")
    if branches:
        out.append('<div class="flow-arrow">&#8595;</div>')
        out.append('<div class="flow-branches">')
        for br in branches:
            out.append('<div class="flow-branch">')
            out.append('<div class="flow-branch-label">' + _esc(br["label"]) + "</div>")
            out.append('<div class="flow-arrow small">&#8595;</div>')
            out.append('<div class="flow-node branch-node">' + _esc(br["node"]) + "</div>")
            out.append("</div>")
        out.append("</div>")
    out.append("</div>")
    return "\n".join(out)


def _render_mnemonic(block: Block) -> str:
    out = ['<div class="mnemonic avoid-break">']
    out.append('<div class="mnemonic-head">Mnemonic: <span class="acronym">'
               + _esc(block["acronym"]) + "</span> &mdash; " + _esc(block["expansion"])
               + "</div>")
    out.append('<table class="mnemonic-table">')
    for letter, word, meaning in block["lines"]:
        out.append("<tr>"
                   + '<td class="m-letter">' + _esc(letter) + "</td>"
                   + '<td class="m-word">' + _esc(word) + "</td>"
                   + '<td class="m-meaning">' + _esc(meaning) + "</td>"
                   + "</tr>")
    out.append("</table>")
    out.append("</div>")
    return "\n".join(out)


def _render_note(block: Block) -> str:
    out = ['<div class="note avoid-break">']
    out.append('<div class="note-title">' + _esc(block["title"]) + "</div>")
    out.append(_render_bullets(block["items"]))
    out.append("</div>")
    return "\n".join(out)


def _render_block(block: Block) -> str:
    btype = block["type"]
    if btype == "bullets":
        return _render_bullets(block["items"])
    if btype == "subhead":
        return '<h3 class="subhead">' + _esc(block["text"]) + "</h3>"
    if btype == "para":
        return "<p>" + _esc(block["text"]) + "</p>"
    if btype == "note":
        return _render_note(block)
    if btype == "flow":
        return _render_flow(block)
    if btype == "mnemonic":
        return _render_mnemonic(block)
    raise ValueError("Unknown block type: %r" % btype)


def _render_section(section: Section) -> str:
    out = ['<section class="topic">']
    out.append('<h2 class="section-title"><span class="section-num">'
               + str(section["num"]) + ".</span> " + _esc(section["title"]) + "</h2>")
    for block in section["blocks"]:
        out.append(_render_block(block))
    out.append("</section>")
    return "\n".join(out)


CSS = """
:root {
  --ink: #1b2330;
  --muted: #4a5568;
  --accent: #2b4c7e;
  --accent-soft: #e8eef7;
  --line: #b9c4d4;
  --flow-bg: #eef3fb;
  --flow-border: #2b4c7e;
  --mnem-bg: #fff6e6;
  --mnem-border: #c98a1b;
  --note-bg: #fdecef;
  --note-border: #b23a48;
}
* { box-sizing: border-box; }
html, body {
  margin: 0; padding: 0;
  color: var(--ink);
  font-family: Georgia, "Times New Roman", serif;
  font-size: 10.5pt; line-height: 1.42;
}
.page {
  width: 210mm; min-height: 297mm;
  margin: 0 auto; padding: 18mm 16mm 20mm 16mm;
  background: #fff;
}
.sheet { background: #d7dce3; padding: 8mm 0; }
h1, h2, h3 { font-family: "Segoe UI", Arial, Helvetica, sans-serif; color: var(--accent); }
h2.section-title {
  font-size: 15pt; margin: 18pt 0 6pt; padding-bottom: 3pt;
  border-bottom: 2px solid var(--accent);
}
.section-num { color: var(--muted); }
h3.subhead {
  font-size: 11.5pt; margin: 11pt 0 3pt; color: var(--ink);
  border-left: 4px solid var(--accent); padding-left: 7px;
}
p { margin: 4pt 0; }
ul { margin: 3pt 0 6pt 0; padding-left: 20px; }
ul ul { margin: 2pt 0 2pt 0; }
li { margin: 2pt 0; }

/* Cover */
.cover { text-align: center; display: flex; flex-direction: column;
  justify-content: center; min-height: 261mm; }
.cover .eyebrow { letter-spacing: 3px; text-transform: uppercase;
  color: var(--muted); font-family: "Segoe UI", Arial, sans-serif; font-size: 11pt; }
.cover h1 { font-size: 27pt; line-height: 1.25; margin: 14pt 24pt; color: var(--accent); }
.cover .subtitle { font-size: 15pt; color: var(--ink); margin-top: 4pt; }
.cover .rule { width: 60mm; height: 3px; background: var(--accent); margin: 20pt auto; }
.cover .credit { font-family: "Segoe UI", Arial, sans-serif; font-size: 12.5pt;
  margin-top: 8pt; color: var(--ink); font-weight: bold; }
.cover .scope { margin: 18pt auto 0; max-width: 150mm; color: var(--muted);
  font-size: 10pt; text-align: left; }

/* Footer credit (repeats via running element in print; shown on every page div for screen) */
.page-footer { margin-top: 10mm; padding-top: 4pt; border-top: 1px solid var(--line);
  font-family: "Segoe UI", Arial, sans-serif; font-size: 8.5pt; color: var(--muted);
  display: flex; justify-content: space-between; }

/* Note / source box */
.note { background: var(--note-bg); border: 1px solid var(--note-border);
  border-left: 5px solid var(--note-border); border-radius: 4px;
  padding: 6pt 10pt; margin: 8pt 0; }
.note-title { font-family: "Segoe UI", Arial, sans-serif; font-weight: bold;
  color: var(--note-border); margin-bottom: 2pt; }

/* Flow chart */
.flow { background: var(--flow-bg); border: 1px solid var(--line);
  border-radius: 5px; padding: 8pt 10pt 10pt; margin: 10pt 0;
  text-align: center; }
.flow-title { font-family: "Segoe UI", Arial, sans-serif; font-weight: bold;
  color: var(--accent); margin-bottom: 6pt; font-size: 10.5pt; }
.flow-node { display: inline-block; background: #fff; border: 1.5px solid var(--flow-border);
  border-radius: 5px; padding: 5pt 10pt; margin: 0 auto; max-width: 150mm;
  font-family: "Segoe UI", Arial, sans-serif; font-size: 9.5pt; }
.branch-node { background: var(--accent-soft); }
.flow-arrow { color: var(--flow-border); font-size: 14pt; line-height: 1; margin: 2pt 0; }
.flow-arrow.small { font-size: 11pt; }
.flow-branches { display: flex; flex-wrap: wrap; justify-content: center; gap: 6pt; }
.flow-branch { border: 1px dashed var(--line); border-radius: 5px; padding: 5pt;
  flex: 1 1 44%; min-width: 60mm; background: rgba(255,255,255,0.5); }
.flow-branch-label { font-family: "Segoe UI", Arial, sans-serif; font-size: 8.5pt;
  color: var(--muted); font-style: italic; }

/* Mnemonic */
.mnemonic { background: var(--mnem-bg); border: 1px solid var(--mnem-border);
  border-left: 5px solid var(--mnem-border); border-radius: 4px;
  padding: 6pt 10pt; margin: 10pt 0; }
.mnemonic-head { font-family: "Segoe UI", Arial, sans-serif; font-weight: bold;
  color: #8a5a00; margin-bottom: 4pt; }
.mnemonic-head .acronym { font-size: 13pt; letter-spacing: 1px; color: var(--mnem-border); }
.mnemonic-table { border-collapse: collapse; width: 100%; }
.mnemonic-table td { padding: 2pt 6pt; vertical-align: top;
  font-family: "Segoe UI", Arial, sans-serif; font-size: 9.5pt; }
.m-letter { font-weight: bold; color: var(--mnem-border); width: 28pt;
  font-size: 12pt; text-align: center; }
.m-word { font-weight: bold; width: 32%; }
.m-meaning { color: var(--muted); }

.toc { font-family: "Segoe UI", Arial, sans-serif; }
.toc h2 { border: none; }
.toc ol { font-size: 11pt; line-height: 1.8; }

.avoid-break { page-break-inside: avoid; break-inside: avoid; }
.topic { page-break-inside: auto; }
section.topic { margin-bottom: 4pt; }

@page {
  size: A4;
  margin: 16mm 14mm 16mm 14mm;
  @bottom-center {
    content: "Made by Tejas Dangodra | M.Phil. Clinical Psychology";
    font-family: Arial, sans-serif; font-size: 8pt; color: #4a5568;
  }
}
@media print {
  .sheet { background: #fff; padding: 0; }
  .page { width: auto; min-height: auto; margin: 0; padding: 0; box-shadow: none; }
  .page-break { page-break-before: always; }
  .cover { min-height: 250mm; }
  h2.section-title { page-break-after: avoid; }
  h3.subhead { page-break-after: avoid; }
}
@media screen {
  .page { box-shadow: 0 1px 6px rgba(0,0,0,0.25); margin-bottom: 8mm; }
}
"""


def render_html() -> str:
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="en">')
    parts.append("<head>")
    parts.append('<meta charset="utf-8">')
    parts.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    parts.append("<title>" + _esc(TITLE) + "</title>")
    parts.append("<style>" + CSS + "</style>")
    parts.append("</head>")
    parts.append("<body>")
    parts.append('<div class="sheet">')

    # ---- Cover page ----
    parts.append('<div class="page cover-page">')
    parts.append('<div class="cover">')
    parts.append('<div class="eyebrow">M.Phil. Clinical Psychology &bull; Study Notes</div>')
    parts.append("<h1>" + _esc(TITLE) + "</h1>")
    parts.append('<div class="subtitle">' + _esc(SUBTITLE) + "</div>")
    parts.append('<div class="rule"></div>')
    parts.append('<div class="credit">' + _esc(CREDIT) + "</div>")
    parts.append('<div class="scope">')
    parts.append("<strong>Scope &mdash; Unit I: Introduction to Psychotherapy.</strong> "
                 "Definitions; objectives; training of professional therapists; ethical "
                 "and legal issues; rights and responsibilities; consent and assent "
                 "(minors); planning and recording of sessions; structuring and goal "
                 "setting; pre- and post-assessment; and the practice of evidence-based "
                 "therapies. Bullet-point study format, with a flow chart and a "
                 "mnemonic for each of the ten topics.")
    parts.append("</div>")
    parts.append("</div>")
    parts.append('<div class="page-footer"><span>' + _esc(CREDIT)
                 + "</span><span>Unit I &bull; Cover</span></div>")
    parts.append("</div>")

    # ---- Contents + source note ----
    parts.append('<div class="page page-break toc">')
    parts.append("<h2>Contents</h2>")
    parts.append("<ol>")
    for section in SECTIONS:
        parts.append("<li>" + _esc(section["title"]) + "</li>")
    parts.append("</ol>")
    parts.append(_render_note(SOURCE_NOTE))
    parts.append('<div class="page-footer"><span>' + _esc(CREDIT)
                 + "</span><span>Unit I &bull; Contents</span></div>")
    parts.append("</div>")

    # ---- Topic pages (two topics per page container; each topic may paginate) ----
    for idx, section in enumerate(SECTIONS):
        parts.append('<div class="page page-break">')
        parts.append(_render_section(section))
        parts.append('<div class="page-footer"><span>' + _esc(CREDIT)
                     + "</span><span>Unit I &bull; Topic " + str(section["num"])
                     + "</span></div>")
        parts.append("</div>")

    parts.append("</div>")  # .sheet
    parts.append("</body>")
    parts.append("</html>")
    return "\n".join(parts)


# ===========================================================================
# DOCX RENDERER  (reuses the SAME content model as the HTML renderer)
# ===========================================================================
# Drawn natively (tables, shading, borders, down-arrow glyphs) so everything
# survives in Word with NO external images. The .docx and .html are generated
# from the one SECTIONS model, so they stay content-identical.

from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.shared import Mm, Pt, RGBColor  # noqa: E402

# Colour palette mirroring the HTML CSS variables.
_ACCENT = RGBColor(0x2B, 0x4C, 0x7E)
_INK = RGBColor(0x1B, 0x23, 0x30)
_MUTED = RGBColor(0x4A, 0x55, 0x68)
_MNEM_BORDER = RGBColor(0xC9, 0x8A, 0x1B)
_MNEM_HEAD = RGBColor(0x8A, 0x5A, 0x00)
_NOTE_BORDER = RGBColor(0xB2, 0x3A, 0x48)

# Shading fill colours (hex strings, no leading '#').
_FLOW_NODE_FILL = "FFFFFF"
_FLOW_BRANCH_FILL = "E8EEF7"
_MNEM_FILL = "FFF6E6"
_NOTE_FILL = "FDECEF"

_ARROW = "\u2193"  # down arrow


def _set_cell_border(cell, color="2B4C7E", sz="8"):
    """Give a table cell a solid single border on all four sides."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def _shade_cell(cell, fill):
    """Apply a solid background fill to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _no_space(paragraph, before=0, after=0):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)


def _add_run(paragraph, text, *, bold=False, italic=False, size=None,
             color=None, font="Segoe UI"):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if font is not None:
        run.font.name = font
    return run


def _docx_bullets(doc, items, level=0):
    for item in items:
        if isinstance(item, dict):
            p = doc.add_paragraph(style="List Bullet")
            if level:
                p.paragraph_format.left_indent = Mm(10 + 6 * level)
            _no_space(p, 0, 2)
            _add_run(p, item["text"], size=10.5, color=_INK, font="Calibri")
            sub = item.get("sub")
            if sub:
                for s in sub:
                    sp = doc.add_paragraph(style="List Bullet 2")
                    _no_space(sp, 0, 1)
                    _add_run(sp, s, size=10, color=_INK, font="Calibri")
        else:
            p = doc.add_paragraph(style="List Bullet")
            if level:
                p.paragraph_format.left_indent = Mm(10 + 6 * level)
            _no_space(p, 0, 2)
            _add_run(p, item, size=10.5, color=_INK, font="Calibri")


def _docx_flow(doc, block):
    """Render a flow chart as a single-column bordered table with down-arrows."""
    nodes = list(block.get("nodes", []))
    branches = block.get("branches")

    title_p = doc.add_paragraph()
    _no_space(title_p, 8, 2)
    _add_run(title_p, "Flow chart: " + str(block["title"]),
             bold=True, size=10.5, color=_ACCENT)

    # One table, one column; each node is a bordered+shaded cell, with an
    # arrow-only row between nodes.
    rows = []
    for i, node in enumerate(nodes):
        rows.append(("node", node))
        if i < len(nodes) - 1 or branches:
            rows.append(("arrow", _ARROW))

    table = doc.add_table(rows=len(rows), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for (kind, text), row in zip(rows, table.rows):
        cell = row.cells[0]
        cell.width = Mm(150)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _no_space(p, 1, 1)
        if kind == "node":
            _set_cell_border(cell, color="2B4C7E", sz="10")
            _shade_cell(cell, _FLOW_NODE_FILL)
            _add_run(p, text, size=9.5, color=_INK)
        else:
            _add_run(p, text, bold=True, size=13, color=_ACCENT)

    if branches:
        for br in branches:
            lp = doc.add_paragraph()
            lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _no_space(lp, 1, 0)
            _add_run(lp, br["label"], italic=True, size=8.5, color=_MUTED)
            ap = doc.add_paragraph()
            ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _no_space(ap, 0, 0)
            _add_run(ap, _ARROW, bold=True, size=11, color=_ACCENT)
            btable = doc.add_table(rows=1, cols=1)
            btable.alignment = WD_TABLE_ALIGNMENT.CENTER
            bcell = btable.rows[0].cells[0]
            bcell.width = Mm(150)
            _set_cell_border(bcell, color="2B4C7E", sz="8")
            _shade_cell(bcell, _FLOW_BRANCH_FILL)
            bp = bcell.paragraphs[0]
            bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _no_space(bp, 1, 1)
            _add_run(bp, br["node"], size=9.5, color=_INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _docx_mnemonic(doc, block):
    """Render a mnemonic as a distinct shaded single-cell table (callout box)."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.width = Mm(166)
    _set_cell_border(cell, color="C98A1B", sz="8")
    _shade_cell(cell, _MNEM_FILL)

    head = cell.paragraphs[0]
    _no_space(head, 2, 3)
    _add_run(head, "Mnemonic: ", bold=True, size=10.5, color=_MNEM_HEAD)
    _add_run(head, str(block["acronym"]), bold=True, size=13, color=_MNEM_BORDER)
    _add_run(head, "  \u2014  " + str(block["expansion"]),
             bold=True, size=10.5, color=_MNEM_HEAD)

    for letter, word, meaning in block["lines"]:
        lp = cell.add_paragraph()
        _no_space(lp, 0, 1)
        _add_run(lp, str(letter) + "  ", bold=True, size=11, color=_MNEM_BORDER)
        _add_run(lp, str(word) + " \u2014 ", bold=True, size=9.5, color=_INK)
        _add_run(lp, str(meaning), size=9.5, color=_MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _docx_note(doc, block):
    """Render a source / caution box as a shaded single-cell table."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.width = Mm(166)
    _set_cell_border(cell, color="B23A48", sz="8")
    _shade_cell(cell, _NOTE_FILL)

    head = cell.paragraphs[0]
    _no_space(head, 2, 3)
    _add_run(head, str(block["title"]), bold=True, size=10.5, color=_NOTE_BORDER)

    for item in block["items"]:
        ip = cell.add_paragraph(style="List Bullet")
        _no_space(ip, 0, 2)
        _add_run(ip, str(item), size=9.5, color=_INK, font="Calibri")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _docx_subhead(doc, text):
    p = doc.add_paragraph()
    _no_space(p, 9, 2)
    _add_run(p, text, bold=True, size=11.5, color=_INK)


def _docx_para(doc, text):
    p = doc.add_paragraph()
    _no_space(p, 3, 3)
    _add_run(p, text, size=10.5, color=_INK, font="Calibri")


def _docx_block(doc, block):
    btype = block["type"]
    if btype == "bullets":
        _docx_bullets(doc, block["items"])
    elif btype == "subhead":
        _docx_subhead(doc, block["text"])
    elif btype == "para":
        _docx_para(doc, block["text"])
    elif btype == "note":
        _docx_note(doc, block)
    elif btype == "flow":
        _docx_flow(doc, block)
    elif btype == "mnemonic":
        _docx_mnemonic(doc, block)
    else:
        raise ValueError("Unknown block type: %r" % btype)


def _docx_section(doc, section, page_break_before=True):
    # Each topic starts on a fresh page (mirrors the one-topic-per-page layout
    # used in the HTML deliverable) so the two stay structurally consistent.
    h = doc.add_heading(level=1)
    if page_break_before:
        h.paragraph_format.page_break_before = True
    _add_run(h, str(section["num"]) + ". " + str(section["title"]),
             bold=True, size=15, color=_ACCENT)
    for block in section["blocks"]:
        _docx_block(doc, block)


def _set_footer_credit(doc):
    """Put the exact credit string in the section footer (repeats on every page)."""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.text = ""
        _add_run(p, CREDIT, size=8.5, color=_MUTED)


def render_docx(path: str = DOCX_OUT) -> "Document":
    doc = Document()

    # A4 page setup with ~20mm margins on the (single) section.
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(20)
    sec.bottom_margin = Mm(20)
    sec.left_margin = Mm(20)
    sec.right_margin = Mm(20)

    # ---- Cover page ----
    eyebrow = doc.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(eyebrow, 90, 6)
    _add_run(eyebrow, "M.PHIL. CLINICAL PSYCHOLOGY  \u2022  STUDY NOTES",
             size=11, color=_MUTED)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(title_p, 6, 6)
    _add_run(title_p, TITLE, bold=True, size=26, color=_ACCENT)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(sub_p, 2, 10)
    _add_run(sub_p, SUBTITLE, size=15, color=_INK)

    credit_p = doc.add_paragraph()
    credit_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(credit_p, 20, 6)
    _add_run(credit_p, CREDIT, bold=True, size=13, color=_INK)

    scope_p = doc.add_paragraph()
    scope_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(scope_p, 24, 0)
    _add_run(scope_p,
             "Scope \u2014 Unit I: Introduction to Psychotherapy. Definitions; "
             "objectives; training of professional therapists; ethical and legal "
             "issues; rights and responsibilities; consent and assent (minors); "
             "planning and recording of sessions; structuring and goal setting; "
             "pre- and post-assessment; and the practice of evidence-based "
             "therapies. Bullet-point study format, with a flow chart and a "
             "mnemonic for each of the ten topics.",
             size=10, color=_MUTED, font="Calibri")

    # Page break after the cover.
    pb = doc.add_paragraph()
    pb.add_run().add_break(WD_BREAK.PAGE)

    # ---- Contents + source note ----
    toc_h = doc.add_heading(level=1)
    _add_run(toc_h, "Contents", bold=True, size=15, color=_ACCENT)
    for section in SECTIONS:
        tp = doc.add_paragraph(style="List Number")
        _no_space(tp, 0, 1)
        _add_run(tp, str(section["title"]), size=11, color=_INK, font="Calibri")
    _docx_note(doc, SOURCE_NOTE)

    # ---- Topic sections (each starts on a new page) ----
    for section in SECTIONS:
        _docx_section(doc, section, page_break_before=True)

    # Footer credit on every page.
    _set_footer_credit(doc)

    doc.save(path)
    return doc


def main() -> None:
    html_text = render_html()
    with open(HTML_OUT, "w", encoding="utf-8") as fh:
        fh.write(html_text)
    print("Wrote HTML:", HTML_OUT)
    print("  bytes:", len(html_text.encode("utf-8")))

    render_docx(DOCX_OUT)
    print("Wrote DOCX:", DOCX_OUT)
    print("  bytes:", os.path.getsize(DOCX_OUT))

    print("  sections:", len(SECTIONS))
    flows = sum(1 for s in SECTIONS for b in s["blocks"] if b["type"] == "flow")
    mnems = sum(1 for s in SECTIONS for b in s["blocks"] if b["type"] == "mnemonic")
    print("  flow charts:", flows, " mnemonics:", mnems)


if __name__ == "__main__":
    main()
