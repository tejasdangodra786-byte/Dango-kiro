# -*- coding: utf-8 -*-
"""Single source of truth for Psychotherapy Unit II notes.

Both the HTML build (build_unit2.py -> parts) and the DOCX build
(build_docx_unit2.py) import SECTIONS from this module so the two deliverables
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

TITLE = "Psychotherapy \u2014 Unit II: Therapeutic Relationship"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

# Cover / title-page sub-heading used by both deliverables.
SUBTITLE = "M.Phil. Clinical Psychology \u2014 short-form study notes (bullet style)"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens both deliverables.
INTRO = [
    {"t": "lead", "text": "These are concise, bullet-style revision notes for Unit II of the Psychotherapy paper (M.Phil. Clinical Psychology). They cover the therapeutic relationship in full, from the working alliance and its models to the client, therapist, illness, technique and other factors that shape it, and fold in the relevant professional bodies (RCI, IACP, WHO and APA) where each is examinable."},
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
        "A citation in square brackets names the full book and its author, for example {{c:Gabbard}} or {{c:Corey}}, so you can trace any claim back to the work it comes from.",
        "Where a point rests on a professional body rather than a book, the citation names the document or body and its year, for example {{c:APA 2017}}, keeping any section number.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail.",
        "Hard clinical terms are kept, but each one is explained right after it in a coloured bracket, like this: [[alliance|the trusting bond and the agreement on the goals and tasks of therapy]]. Just skim the coloured bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no laws, citations, clause numbers or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "{{SEC}} IACP (Indian Association of Clinical Psychologists): its members' code of conduct is represented only from secondary summaries, so it is described in general terms as broadly aligning with APA and RCI duties, and no IACP clause numbers are quoted.",
        "Primary texts relied on directly: the APA Ethics Code 2017 and the Rehabilitation Council of India (RCI) Act 1992.",
        "Statements about professional bodies are for study; in practice read the current text and seek advice, as codes change.",
    ]},
]

# Short reading list keyed to the topics.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the books used and the primary ethical and online sources consulted."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Glen O. Gabbard \u2014 Long-Term Psychodynamic Psychotherapy: A Basic Text (therapeutic alliance, transference and countertransference).",
        "Gerald Corey \u2014 Theory and Practice of Counseling and Psychotherapy (therapist characteristics and the relationship).",
        "Raymond J. Corsini & Danny Wedding \u2014 Current Psychotherapies (working alliance and common factors).",
        "Richard S. Sharf \u2014 Theories of Psychotherapy and Counseling (Rogers' core conditions).",
        "Christine Lister-Ford \u2014 A Short Introduction to Psychotherapy (the relationship in process).",
    ]},
    {"t": "h3", "text": "Primary ethical and online sources"},
    {"t": "ul", "items": [
        "APA Ethics Code (2017); APA Dictionary of Psychology.",
        "World Health Organization guidance on psychological interventions.",
        "Rehabilitation Council of India (RCI) Act 1992.",
        "{{SEC}} IACP code of conduct \u2014 secondary summaries only; primary text not accessed, so no clause numbers are quoted.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. ALLIANCE & MODELS ===========================
_s("t1", "1", "The therapeutic relationship and working alliance: definition and models", [
    {"t": "lead", "text": "The therapeutic relationship, or alliance, is the collaborative emotional bond and working partnership between client and therapist; it is both the vehicle through which every technique is delivered and, for many authors, a curative factor in its own right."},
    {"t": "fig", "key": "alliance_model", "caption": "How Bordin's bond, goals and tasks and Rogers' core conditions feed into one working alliance."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises the [[therapeutic alliance|the cooperative working relationship between client and therapist]] as the **cooperative working relationship**, seen as **essential** for progress, that combines the client's affective **bond** with the therapist and their agreement on the **goals and tasks** of treatment. {{c:APA Dictionary of Psychology}}",
        "{{GUIDE}} The **World Health Organization**, in its mhGAP and psychological-interventions guidance, characterises effective care as being delivered within a **respectful, trusting and collaborative** provider-client relationship, which it treats as a **foundation** for engagement and adherence rather than an optional extra. {{c:WHO mhGAP}}",
        "{{THEORY}} **Gabbard** describes the therapeutic alliance as a **fundamental concept** that applies across all forms of psychotherapy, denoting the **collaborative bond** that lets client and therapist work together toward shared aims. {{c:Gabbard}}",
        "{{THEORY}} **Greenson (1967)**, reported in Corsini and Wedding, framed the **working alliance** as the **rational, non-neurotic** partnership in which the client aligns with the therapist's method to pursue the work of therapy. {{c:Corsini}}",
    ]},
    {"t": "h3", "text": "Why the relationship matters (exam form)"},
    {"t": "ul", "items": [
        "The alliance is often called the [[vehicle|the carrier through which specific methods take effect]] of therapy, because **techniques work through the relationship**, so even a brilliant method falls flat when the bond with the client is weak. {{c:Gabbard}}",
        "{{THEORY}} **Rogers (1957)** argued that the relationship conditions of empathy, warmth and genuineness are **necessary and sufficient** for change, so that for him the relationship itself is the active ingredient rather than a mere background. {{c:Sharf}}",
        "{{THEORY}} Many later clinicians, especially **behaviour therapists**, call those same conditions **necessary but not sufficient**: a good bond is needed, but it has to be paired with specific techniques before real change follows. {{c:Corey}}",
        "The alliance is a **pantheoretical common factor**, which means it cuts across psychodynamic, cognitive-behavioural, humanistic and systemic schools instead of belonging to any single model. {{c:Corsini}}",
        "A strong early bond helps an anxious or reluctant client feel **safe enough to open up**, so trust comes first and only then can assessment, formulation and treatment really get going. {{c:Lister-Ford}}",
    ]},
    {"t": "h3", "text": "Models and components of the relationship"},
    {"t": "ul", "items": [
        "{{THEORY}} **Bordin (1979)** proposed the **working alliance** as a pantheoretical model with three interdependent parts, namely agreement on **goals**, agreement on **tasks**, and the emotional **bond** of trust and liking between the pair. {{c:Corsini}}",
        "In Bordin's model **goals** are the agreed outcomes the client wants, **tasks** are the in-session activities believed to reach them, and the **bond** is the affective tie that makes the shared work feel safe and worthwhile. {{c:Corsini}}",
        "{{THEORY}} **Rogers (1957)** named three **core conditions**, which are accurate [[empathy|grasping the client's feelings as if they were one's own]], [[unconditional positive regard|non-judgemental warmth and acceptance of the person]] and [[congruence|being genuine, with outward manner matching inner experience]]. {{c:Sharf}}",
        "Rogers stressed that the core conditions must be **perceived by the client**, because warmth and understanding only help when the client actually experiences them, not merely when the therapist intends them. {{c:Corey}}",
        "{{THEORY}} **Greenson (1967) and Gelso** split the relationship into three strands, namely the **working alliance**, the **transference-countertransference** configuration, and the **real relationship** of genuine, realistic perceptions of each other. {{c:Gabbard}}",
        "{{THEORY}} **Luborsky** described a **helping alliance** with two types, the first being the client's experience of the therapist as **supportive**, and the second a sense of **shared work** toward common goals, while **Zetzel** used the alliance as the basis for analytic work. {{c:Gabbard}}",
        "{{THEORY}} From an **attachment** perspective after **Bowlby**, the therapist functions as a [[secure base|a dependable figure from whom the client can safely explore distress]], so the client's own attachment style strongly shapes how the bond forms. {{c:Gabbard}}",
        "[[Transference|the client's displacement onto the therapist of feelings from past relationships]] can be positive or negative and, when understood, becomes rich material rather than an obstacle to the work. {{c:Gabbard}}",
        "[[Countertransference|the therapist's own emotional reactions to the client]] is managed through **self-awareness and supervision** and can be used as **information** about what the client evokes in others. {{c:Gabbard}}",
        "[[Resistance|the client's conscious or unconscious opposition to the therapy process]] shows as silence, lateness, avoidance or intellectualisation, and is explored rather than fought, because it signals where the work is threatening. {{c:Gabbard}}",
        "The relationship moves through **stages**, namely engagement and formation, a deepening **working** phase, and **termination**, where ending is itself therapeutic because it can revive earlier experiences of loss and separation. {{c:Lister-Ford}}",
    ]},
    {"t": "box", "kind": "key", "title": "Bordin vs Rogers at a glance", "items": [
        "Bordin (working alliance): bond + agreement on goals + agreement on tasks; pantheoretical, suits every model.",
        "Rogers (core conditions): empathy + unconditional positive regard + congruence; must be perceived by the client.",
        "Greenson/Gelso: working alliance + transference-countertransference + real relationship.",
        "Shared message: a collaborative, trusting bond is the common ground beneath all schools.",
    ]},
    {"t": "mn", "word": "BeGoT", "text": "Bond, Goals, Tasks \u2014 Bordin's three parts of the working alliance; pair it with Rogers' ECG (Empathy, Congruence, unconditional positive reGard) to recall both leading models."},
])

# ========================= 2. CLIENT CHARACTERISTICS ======================
_s("t2", "2", "Client characteristics that shape the relationship", [
    {"t": "lead", "text": "What the client brings, including their readiness to change, expectations, attachment style, coping style, insight and culture, strongly influences how quickly and how firmly a working alliance can form."},
    {"t": "fig", "key": "client_factors", "caption": "The main client characteristics that feed into, or strain, the alliance."},
    {"t": "ul", "items": [
        "{{THEORY}} **Prochaska and DiClemente's** stages of change run through [[precontemplation|not yet seeing a problem]], contemplation, preparation, action, maintenance and relapse, and interventions work best when **matched to the client's stage**. {{c:Corsini}}",
        "Pushing a **precontemplative** client who does not yet want change usually just **provokes resistance**, so a gentler, exploratory stance such as motivational interviewing builds the bond far better early on. {{c:Corsini}}",
        "A client's **expectations and hope** are a major common factor, because someone who believes therapy can help engages more fully, while **role expectations** shape what they think the work will involve. {{c:Corsini}}",
        "**Attachment style** predicts how the bond forms, since **secure** clients build an alliance quickly whereas insecure styles make trust slower and more effortful to establish. {{c:Gabbard}}",
        "**Avoidant** clients tend to keep **emotional distance**, **anxious-ambivalent** clients may become **over-dependent**, and **disorganised** clients often show frequent ruptures that the therapist must patiently repair. {{c:Gabbard}}",
        "{{THEORY}} **Beutler** distinguishes coping styles, so that **externalising** clients who act out suit symptom-focused methods while **internalising** clients who turn inward suit insight-oriented work. {{c:Corsini}}",
        "{{THEORY}} **Reactance**, meaning a client's [[sensitivity to being controlled|a tendency to resist perceived pressure on their freedom]], matters because **high-reactance** clients do better with **non-directive** approaches that leave them in charge. {{c:Corsini}}",
        "**Psychological mindedness**, insight, verbal ability and intelligence ease [[insight-oriented therapy|therapy that works by understanding hidden patterns and meanings]], because such work depends on reflecting on one's own mind. {{c:Sharf}}",
        "The old **YAVIS** stereotype (young, attractive, verbal, intelligent, successful) versus the **HOUND** stereotype (homely, old, unintelligent, non-verbal, disadvantaged) exposes a long-standing **bias** toward favouring easier, more articulate clients. {{c:Sharf}}",
        "**Culture and demographics** such as age, gender, language, religion and beliefs about illness shape the bond, and a **culturally mismatched** therapy that ignores these often ends in early dropout. {{c:Corey}}",
        "Clients who explain distress through **spiritual or somatic** frames, or who face heavy **stigma** and strong family expectations, need the therapist to work respectfully **within** that worldview rather than against it. {{c:Corey}}",
        "**Severity, distress and functioning** matter too, since severe impairment, marked interpersonal difficulty and **chronic** problems all make a stable alliance slower and harder to build. {{c:Gabbard}}",
        "A client's **previous therapy experience** colours the new relationship, so that past help that failed lowers hope while a past good experience primes trust in the current therapist. {{c:Lister-Ford}}",
        "The **referral source** matters, because a **voluntary** client who seeks help engages more readily than a **mandated** or family-pushed client who may feel coerced into attending. {{c:Corey}}",
        "A client's **social support and life circumstances**, including housing, finances and relationships, shape the energy and stability they can bring to the therapeutic work. {{c:Lister-Ford}}",
    ]},
    {"t": "mn", "word": "CARE-PS", "text": "Change-readiness, Attachment style, Reactance and coping, Expectations and hope, Psychological mindedness, Social and cultural context \u2014 the client characteristics that shape the alliance."},
])

# ========================= 3. THERAPIST CHARACTERISTICS ===================
_s("t3", "3", "Therapist characteristics that shape the relationship", [
    {"t": "lead", "text": "Therapists vary in how well they form alliances; their personal qualities, professional competence, stance and self-awareness matter more than demographic matching, and some therapists are reliably more effective than others."},
    {"t": "fig", "key": "therapist_factors", "caption": "Personal qualities, professional competence and stance combining into therapist effectiveness."},
    {"t": "ul", "items": [
        "The core **personal qualities** are **empathy, warmth and genuineness**, together with respect, trustworthiness, patience and a hope-inspiring attitude that together let the client feel understood and safe. {{c:Corey}}",
        "**Emotional stability, flexibility, humour, self-awareness** and tolerance of ambiguity help the therapist stay steady and responsive when the client is distressed, hostile or ambivalent. {{c:Corey}}",
        "**Professional competence** grows from training, experience and supervision, plus the ability to **formulate** a case and adapt the treatment, so skill and character really work hand in hand. {{c:Gabbard}}",
        "[[Cultural competence|the knowledge, awareness and skills to work respectfully across cultures]] lets the therapist respond to the client's values and context, which strongly supports engagement in diverse settings. {{c:Corey}}",
        "{{THEORY}} **Therapist effects** are real, meaning some therapists consistently achieve **better outcomes** than others, and these differences between therapists are often **larger** than the differences between treatment methods. {{c:Corsini}}",
        "**Therapeutic stance and style**, such as how **directive** or active the therapist is, must be tuned to the client, because a style that suits one client can alienate another. {{c:Sharf}}",
        "[[Self-disclosure|the therapist sharing something personal]] can strengthen the bond when it is **brief and serves the client**, but harms the work when it shifts the focus onto the therapist's own needs. {{c:Gabbard}}",
        "**Countertransference** from unresolved conflicts, a need to **rescue**, over-identification, anger, boredom or attraction can distort the relationship unless the therapist notices and works on it. {{c:Gabbard}}",
        "**Supervision and personal work** are the main safeguards, letting the therapist process their reactions so that personal blind spots do not silently damage the alliance. {{c:Gabbard}}",
        "A therapist's **expectations and belief** in the treatment, together with genuine **positive regard** for the client, raise the client's own hope and strengthen engagement. {{c:Corsini}}",
        "**Demographic matching** of gender, ethnicity or language can comfort some clients, but it matters **less** than the therapist's warmth, respect and cultural responsiveness. {{c:Corey}}",
        "{{CODE}} **Competence and ethical conduct** are professional obligations, and in India the **RCI Act 1992** sets standards of conduct and etiquette for registered rehabilitation professionals including clinical psychologists. {{c:RCI Act 1992}}",
        "**Burnout** erodes empathy and patience, so a depleted or overloaded therapist forms weaker alliances, which makes self-care an ethical as well as a personal concern. {{c:Corey}}",
        "**Harmful therapist qualities** include coldness, rigidity, authoritarianism, criticism, over-involvement and disrespect for boundaries, all of which damage trust and can make the client worse. {{c:Corey}}",
        "The single most decisive factor is the **therapist's attitude** toward the client, because a respectful, responsive and self-aware therapist builds strong alliances across a wide range of clients and problems. {{c:Corsini}}",
    ]},
    {"t": "mn", "word": "WE-CARES", "text": "Warmth and empathy, Experience and competence, Cultural responsiveness, Attitude and positive regard, Reflective self-awareness, Ethical conduct, Stance tuned to the client \u2014 the therapist qualities that build the alliance."},
])

# ========================= 4. ILLNESS FACTORS =============================
_s("t4", "4", "How the nature and severity of illness influence the relationship", [
    {"t": "lead", "text": "The type, severity and chronicity of the disorder, along with insight, risk, trauma history and cognitive state, change how quickly trust can form and how active the therapist must be to hold the relationship."},
    {"t": "fig", "key": "illness_gradient", "caption": "A gradient from easier-to-engage to harder-to-engage presentations, with the demands each places on the bond."},
    {"t": "ul", "items": [
        "Alliances form most **easily** with mild to moderate **anxiety** and **adjustment** problems, because such clients usually retain insight, motivation and the capacity to collaborate. {{c:Gabbard}}",
        "**Personality disorders**, especially **borderline, paranoid and antisocial** presentations, strain the bond through mistrust, testing and rapid shifts, so the alliance needs careful, consistent tending. {{c:Gabbard}}",
        "**Psychosis** makes engagement hard because **suspicion and poor insight** can lead the client to doubt the therapist's intentions, so trust is built slowly and concretely. {{c:Gabbard}}",
        "**Severe depression** brings **withdrawal, hopelessness and low energy**, which slow responsiveness and mean the therapist must carry more of the early activity and warmth. {{c:Gabbard}}",
        "**Substance use** often involves **denial and ambivalence**, while **eating disorders** carry ambivalence about recovery, so both call for a non-confrontational, motivation-building stance early on. {{c:Corsini}}",
        "**Severity and chronicity** stretch out the time it takes to build a bond, and repeated **failed treatments** drain hope, so the therapist has to actively rebuild the client's expectation of improvement. {{c:Gabbard}}",
        "Poor **insight and motivation**, as in acute psychosis or mania, reduce collaboration, so the work may begin with engagement and psychoeducation rather than deeper therapeutic tasks. {{c:Gabbard}}",
        "**Comorbidity**, such as depression combined with anxiety or substance use, complicates the picture and can divide the client's attention and the alliance across competing problems. {{c:Gabbard}}",
        "**Suicidality and risk** demand a more **active, structured** stance and may strain boundaries, because safety planning and monitoring temporarily reshape the usual therapeutic frame. {{c:Corey}}",
        "{{CODE}} Managing risk also engages legal and ethical duties, since the **APA Ethics Code 2017** permits disclosure without consent to protect the client or others from serious harm. {{c:APA 2017, Std 4.05}}",
        "A **trauma and abuse** history makes trust difficult, so the therapist prioritises **safety and gradual pacing**, letting the client set the speed rather than pressing for early disclosure. {{c:Gabbard}}",
        "**Cognitive impairment** from organic disorders or intellectual disability calls for **simplified communication** and often family involvement, which changes the shape of the relationship. {{c:Gabbard}}",
        "**Medication and physical illness**, through sedation, side effects, pain or fatigue, affect engagement and alertness, so the therapist adapts session length and expectations accordingly. {{c:Gabbard}}",
        "In **bipolar disorder**, the raised energy and grandiosity of [[mania|an elevated, overactive mood state with reduced insight]] can shrink **collaboration**, so the therapist waits for stability before attempting deeper work. {{c:Gabbard}}",
        "In **OCD and anxiety**, repeated [[reassurance-seeking|asking the therapist again and again to confirm nothing is wrong]] can pull the therapist into an **unhelpful role** that quietly feeds the very symptom. {{c:Corsini}}",
        "Across every presentation one principle holds: the **sicker or more mistrustful** the client, the more the therapist must **adjust pace, structure and activity** to protect a fragile bond. {{c:Gabbard}}",
    ]},
    {"t": "mn", "word": "SICK-TM", "text": "Severity and chronicity, Insight and motivation, Comorbidity, Kind of disorder, Trauma history, risk and Medication/physical illness \u2014 the illness factors that shape the alliance."},
])

# ========================= 5. TECHNIQUE FACTORS ===========================
_s("t5", "5", "How technique and therapeutic approach influence the relationship", [
    {"t": "lead", "text": "The techniques chosen, how well they fit the client, how clearly they are explained, and how flexibly they are applied all feed back into the alliance, so method and relationship are tightly linked rather than separate."},
    {"t": "fig", "key": "technique_match", "caption": "Matching technique to the client, with timing and flexibility, strengthening the alliance."},
    {"t": "ul", "items": [
        "{{THEORY}} **Beutler's prescriptive matching** tailors the method to the client, so **directive** techniques suit **low-reactance** clients while **non-directive** work suits **high-reactance** clients. {{c:Corsini}}",
        "Matching also follows coping style, with **symptom-focused** techniques suiting **externalisers** and **insight-focused** work suiting **internalisers**, so a good fit itself supports the bond. {{c:Corsini}}",
        "A **clear, credible rationale** for the treatment raises the client's **expectation** and strengthens the alliance, because people engage more when they understand why a task should help. {{c:Corsini}}",
        "**Timing and accuracy** of interventions matter, so a well-timed, accurate [[interpretation|a statement linking the client's pattern to its hidden meaning]] deepens the bond. {{c:Gabbard}}",
        "A **premature or inaccurate** interpretation usually breeds **defensiveness**, because the client feels misunderstood or exposed before enough trust has had time to build. {{c:Gabbard}}",
        "An overly **directive** style can trigger **reactance** in clients who resist being controlled, while an overly **passive** style can leave clients feeling confused, unheld or frustrated. {{c:Sharf}}",
        "**Confrontation, exposure or challenging** techniques used too early can strain the alliance, so the **bond is built first** and demanding work is introduced once trust can bear it. {{c:Corey}}",
        "[[Responsive flexibility|adjusting the plan to the client in the moment]] improves the alliance, whereas **rigid manual adherence** in the face of a rupture tends to worsen it. {{c:Corsini}}",
        "**Homework and between-session tasks** shape the **task agreement**, so negotiating tasks the client sees as relevant and doable keeps the collaborative part of the bond strong. {{c:Beck}}",
        "The **format** of therapy, whether individual, group or family, changes the relationship, since group and family work add multiple bonds and shared cohesion to manage. {{c:Corey}}",
        "**Teletherapy** by video can produce an alliance **comparable** to in-person work, but it is vulnerable to **technical problems** and the loss of some non-verbal cues. {{c:Corey}}",
        "{{GUIDE}} Because flexibility helps, the therapist treats the manual as a **guide not a cage**, adapting technique to the client's reactions while keeping the treatment coherent and evidence-based. {{c:Corsini}}",
        "In CBT, [[collaborative empiricism|client and therapist checking beliefs against the evidence as a team]] is itself a **relationship stance**, because investigating the problem together treats the client as a **partner** rather than a passive recipient. {{c:Beck}}",
        "Using [[Socratic questioning|gently guiding the client with open questions until they reach their own insight]] rather than **lecturing** protects the bond, because conclusions the client reaches themselves feel **their own** rather than forced on them. {{c:Beck}}",
        "The overarching rule is simple: **technique and relationship interact**, so the best outcomes come from applying an effective method **responsively** inside a strong, well-tended alliance. {{c:Corsini}}",
    ]},
    {"t": "mn", "word": "MR-FIT", "text": "Matching (Beutler), Rationale credibility, Flexibility vs rigid adherence, Interpretation timing and accuracy, Tasks, format and teletherapy \u2014 the technique factors that shape the alliance."},
])

# ========================= 6. OTHER FACTORS ===============================
_s("t6", "6", "Other factors: setting, culture, ruptures and repair, transference, boundaries, teletherapy", [
    {"t": "lead", "text": "Beyond the people and the method, the frame and setting, culture and power distance, ruptures and their repair, transference, professional boundaries and technology all shape the relationship, and ethics bodies govern the boundary issues directly."},
    {"t": "fig", "key": "rupture_repair", "caption": "The rupture-repair cycle: a withdrawal or confrontation rupture, met with meta-communication and repair, can strengthen the alliance."},
    {"t": "h3", "text": "Setting, culture and context"},
    {"t": "ul", "items": [
        "The **therapy frame** of privacy, comfort, location, fees, session length, frequency and cancellation policy sets the stable conditions within which trust can grow. {{c:Lister-Ford}}",
        "**Duration** affects depth, because a longer treatment allows a deeper bond while **abrupt endings** strain the relationship and can reactivate earlier experiences of loss. {{c:Gabbard}}",
        "**Culture and power distance** shape the bond, and in many **Indian** settings clients may defer to the therapist as a **doctor figure** and avoid openly disagreeing, which the therapist must actively invite. {{c:Corey}}",
        "**Stigma, language barriers, gender norms and religion** all influence engagement, so the therapist works respectfully within the client's social and cultural world rather than against it. {{c:Corey}}",
        "**Life events and external stress** during therapy, along with **third parties** such as family, insurers or courts in mandated work, pull on the relationship from outside the room. {{c:Corey}}",
        "**Therapist-client fit** or chemistry is hard to measure yet genuinely matters, so inviting **early feedback** and being ready to **refer elsewhere** when the fit is poor is itself good practice. {{c:Corsini}}",
    ]},
    {"t": "h3", "text": "Ruptures and repair"},
    {"t": "ul", "items": [
        "{{THEORY}} **Safran and Muran** describe [[ruptures|tensions or breakdowns in the alliance]] in two forms, namely **withdrawal** ruptures where the client goes silent, compliant or changes the topic, and **confrontation** ruptures where the client criticises, shows anger or makes demands. {{c:Corsini}}",
        "**Repair** uses [[meta-communication|openly talking about what is happening between client and therapist]], empathic validation and flexibility with tasks or goals, and is associated with **better outcome**. {{c:Corsini}}",
        "A **repaired rupture can leave the alliance stronger** than before, so the rupture-repair cycle is treated as valuable therapeutic material rather than as a failure on the therapist's part. {{c:Gabbard}}",
        "The therapist **takes responsibility** for their own contribution to a rupture, because owning the strain models honesty and makes it safe for the client to raise future tensions. {{c:Gabbard}}",
    ]},
    {"t": "box", "kind": "key", "title": "Withdrawal vs confrontation ruptures (Safran & Muran)", "items": [
        "Withdrawal rupture: client turns away - silence, compliance, changing the topic, vague agreement.",
        "Confrontation rupture: client turns against - criticism, anger, complaints, demands on the therapist.",
        "Repair for both: notice it, name it gently (meta-communication), validate, and adjust tasks or goals.",
        "Outcome: a well-repaired rupture often leaves the alliance stronger than before.",
    ]},
    {"t": "h3", "text": "Boundaries, dual relationships and the bodies"},
    {"t": "ul", "items": [
        "[[Transference and countertransference|the client's displaced feelings and the therapist's reactions to them]] are managed openly, with supervision, so that strong feelings in the room inform rather than derail the work. {{c:Gabbard}}",
        "{{CODE}} **Professional boundaries** protect the client, and the **APA Ethics Code 2017, Standard 3.05 Multiple Relationships** warns against second relationships that could impair objectivity or risk exploitation or harm. {{c:APA 2017, Std 3.05}}",
        "{{CODE}} The **APA Ethics Code 2017, Standards 10.05 to 10.08** prohibit **sexual intimacies** with current therapy clients and tightly restrict them with former clients, making this one of the clearest boundary rules in the code. {{c:APA 2017, Std 10.05}}",
        "{{CODE}} In India the **RCI Act 1992** empowers the Rehabilitation Council to prescribe **standards of conduct, etiquette and ethics** for registered professionals, so boundary and competence duties carry statutory backing. {{c:RCI Act 1992}}",
        "{{GUIDE}} The **World Health Organization** likewise frames the provider-client relationship as resting on **trust, respect and confidentiality**, which boundary-keeping is designed to protect. {{c:WHO mhGAP}}",
        "{{SEC}} The **IACP** (Indian Association of Clinical Psychologists) is understood from secondary summaries to broadly **align with APA and RCI** duties on competence and boundaries, and no IACP clause numbers are quoted here because the primary code was not accessed. {{c:IACP}}",
        "**Teletherapy** adds boundary and privacy demands, since the therapist must secure the setting, confirm the client's location for safety, and manage confidentiality across the technology used. {{c:Corey}}",
    ]},
    {"t": "mn", "word": "SCRuB-T", "text": "Setting and frame, Culture and power distance, Ruptures and repair, Boundaries and dual relationships, Transference and teletherapy \u2014 the other factors that shape the alliance."},
])

# ========================= 7. MEASURING & OUTCOME =========================
_s("t7", "7", "Measuring the alliance and its link to outcome", [
    {"t": "lead", "text": "The alliance can be measured with validated scales, and meta-analyses show a modest but remarkably consistent link between a strong alliance, especially early in therapy, and better treatment outcome across diagnoses and formats."},
    {"t": "fig", "key": "alliance_outcome", "caption": "Measure the alliance early, feed it back, and strengthen a bond that predicts outcome at about r = .28."},
    {"t": "ul", "items": [
        "The alliance is measured with validated self-report and observer scales, chiefly the **Working Alliance Inventory (WAI)** of Horvath and Greenberg, the **California Psychotherapy Alliance Scales (CALPAS)**, the **Helping Alliance Questionnaire (HAq)** and the brief **Session Rating Scale**. {{c:Corsini}}",
        "The **WAI** follows Bordin directly, scoring the three components of **bond, goals and tasks**, which makes the model measurable and lets therapist and client track the partnership over time. {{c:Corsini}}",
        "{{THEORY}} **Horvath and Symonds (1991)** and later **Horvath et al. (2011)** found a **moderate and consistent** association between alliance and outcome of about **r = .28** across many studies. {{c:APA Div 29 Norcross}}",
        "{{THEORY}} **Fluckiger et al. (2018)** confirmed this alliance-outcome link in a large updated meta-analysis, showing it holds across **therapies, diagnoses and formats** rather than being tied to one school. {{c:APA Div 29 Norcross}}",
        "For many disorders the relationship is held to explain **more outcome variance** than the specific technique, which is exactly why common-factors researchers place the alliance at the heart of effective therapy. {{c:Corsini}}",
        "The **early alliance**, often measured by sessions three to five, **predicts** later outcome, so monitoring the bond early gives a practical, actionable warning sign well before termination. {{c:APA Div 29 Norcross}}",
        "Causation runs **both ways**, because a good alliance partly **causes** improvement and is partly an **effect** of early symptom relief, so the two reinforce each other over a course of therapy. {{c:Corsini}}",
        "{{GUIDE}} The **Norcross and Lambert** APA task force classified the **alliance, empathy, goal consensus, group cohesion** and **collecting client feedback** as **demonstrably effective** relationship elements. {{c:APA Div 29 Norcross}}",
        "**Routine feedback**, using brief alliance and outcome measures each session, lets the therapist catch a weakening bond and repair it, which is itself linked to better results. {{c:Corsini}}",
        "The clinical message is that measuring the alliance is **not just academic**, because a low early score is a clear cue to slow down, invite feedback and repair the relationship before pressing on with technique. {{c:Corsini}}",
        "Because the correlation is **modest in size but highly consistent**, the alliance is best read as a **necessary foundation** that works together with effective technique rather than a guarantee of success on its own. {{c:Corsini}}",
        "{{THEORY}} The convergence of **Bordin's model, the WAI and the meta-analytic evidence** gives the alliance a rare status in psychotherapy, as a concept that is at once theoretically clear, measurable and empirically supported. {{c:APA Div 29 Norcross}}",
        "A practical strength of the brief measures is their **low burden**, since scales such as the **Session Rating Scale** take under a minute, which makes session-by-session [[alliance monitoring|routinely tracking the bond across sessions]] feasible in busy clinical settings. {{c:Corsini}}",
        "The evidence also warns against **complacency**, because an alliance that looks strong can still mask an unspoken [[rupture|a strain or breakdown in the bond]], so therapists are urged to invite honest feedback rather than assume a high score settles the matter. {{c:Corsini}}",
        "To make this routine, [[measurement-based care|checking the scores at each session and letting them guide treatment]] turns alliance feedback into **regular clinical data**, so a **weakening bond** is caught and acted on early instead of only when therapy has already broken down. {{c:Corsini}}",
    ]},
    {"t": "mn", "word": "WAI-FOR", "text": "WAI and other measures, Alliance-outcome r about .28, Illness-wide consistency, Feedback routinely collected, Outcome predicted by the early alliance, Repair when it dips \u2014 how the alliance is measured and why it matters."},
    {"t": "note", "text": "{{THEORY}} Meta-analytic alliance-outcome findings are credited to the named researchers and the APA task force work (Norcross, 'Psychotherapy Relationships That Work'); no page numbers are attached to these research findings. {{c:APA Div 29 Norcross}}"},
])
