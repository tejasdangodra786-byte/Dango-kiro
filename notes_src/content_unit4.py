# -*- coding: utf-8 -*-
"""Single source of truth for Psychotherapy Unit IV notes (Affective Psychotherapies).

Both the HTML build (build_unit4.py -> parts) and the DOCX build
(build_docx_unit4.py) import SECTIONS from this module so the two deliverables
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

TITLE = "Psychotherapy \u2014 Unit IV: Affective Psychotherapies"
CREDIT = "Dango_Psych by Tejas Dangodra | M.Phil. Clinical Psychology"

# Cover / title-page sub-heading used by both deliverables.
SUBTITLE = "M.Phil. Clinical Psychology study notes (bullet style) covering the affective psychotherapies"

# ---------------------------------------------------------------------------
# Front-matter: a short "how to read" intro that opens both deliverables.
INTRO = [
    {"t": "lead", "text": "These are concise, exam-ready revision notes for Unit IV of the Psychotherapy paper (M.Phil. Clinical Psychology). They cover the affective psychotherapies one by one, and for each therapy they weave in its origin, basis, formulation, goals, procedures and techniques, stages and process, outcome, indications and the research and current status, so you can revise a whole 5-mark answer from a single topic."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one therapy from the syllabus, written as a 5-mark exam answer in full sentences rather than terse fragments, so you can revise straight from the point and reproduce it under exam conditions.",
        "Every topic opens with a one-line summary, carries a flowchart figure that pictures the core idea, gives a short definition box drawn from recognised bodies and classic authors, and closes with a mnemonic that packs the key points into a single memorable word.",
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
        "A citation in square brackets names the full book and its author, for example {{c:Corey}} or {{c:Yalom}}, so you can trace any claim back to the work it comes from.",
        "Foundational ideas are cited to the author and classic work that introduced them, for example {{c:Freud}} or {{c:Rogers}}, and where a point rests on a professional body the citation names that body, for example {{c:APA Dictionary of Psychology}}.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks: skim the bold words first to recall the gist of a sentence, then read the full sentence to fill in the detail.",
        "Hard clinical terms are kept, but each one is explained right after it in a coloured bracket, like this: [[transference|the client's redirection of feelings from past relationships onto the therapist]]. Just skim the coloured bracket to get the plain meaning, so the difficult word never leaves you stuck.",
        "Every topic carries one flowchart figure that pictures the idea and one mnemonic that packs its key points into a single memorable word.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no authors, citations, study results, effect sizes or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "Professional-body characterisations of the therapies (for example the APA Dictionary of Psychology entries for psychotherapy, psychodynamic psychotherapy, person-centered therapy and gestalt therapy) are given as characterisations drawn from standard reference wording, not as verbatim quotations, and no page numbers are attached.",
        "Foundational claims are attributed to the author and classic work that introduced them (Freud, Rogers, Perls, May, Yalom, Frankl, Adler); where research is summarised, only well-known findings and named authors are given, and the strength of evidence is stated honestly rather than dressed up with invented numbers.",
        "{{SEC}} Where an Indian professional body such as the IACP is mentioned, it is represented only in general terms from secondary summaries, and no clause numbers are quoted.",
    ]},
]

# Short reading list keyed to the topics.
REFERENCES = [
    {"t": "lead", "text": "Key sources: the books used and the primary and online sources consulted for the affective-psychotherapies content."},
    {"t": "h3", "text": "Books"},
    {"t": "ul", "items": [
        "Gerald Corey \u2014 Theory and Practice of Counseling and Psychotherapy (psychodynamic, humanistic, existential, Gestalt, person-centred and Adlerian therapies).",
        "Richard S. Sharf \u2014 Theories of Psychotherapy and Counseling (theory, techniques and process of each approach).",
        "Raymond J. Corsini & Danny Wedding \u2014 Current Psychotherapies (comparative overview of the major systems).",
        "Glen O. Gabbard \u2014 Long-Term Psychodynamic Psychotherapy: A Basic Text (psychodynamic technique, transference and the therapeutic frame).",
        "Nancy McWilliams \u2014 Psychoanalytic Diagnosis and Psychoanalytic Case Formulation (psychodynamic formulation and defences).",
        "Sigmund Freud \u2014 The Interpretation of Dreams and Introductory Lectures on Psycho-Analysis (the foundations of psychoanalysis).",
        "Carl R. Rogers \u2014 Client-Centered Therapy and On Becoming a Person (the person-centred approach and the core conditions).",
        "Irvin D. Yalom \u2014 Existential Psychotherapy (death, freedom, isolation and meaninglessness).",
        "Viktor E. Frankl \u2014 Man's Search for Meaning (logotherapy, paradoxical intention and dereflection).",
        "Frederick S. Perls, Ralph Hefferline & Paul Goodman \u2014 Gestalt Therapy (awareness, contact and the here-and-now).",
        "Alfred Adler \u2014 The Science of Living and Understanding Human Nature (Individual Psychology, inferiority and social interest).",
        "Eric Berne \u2014 Transactional Analysis in Psychotherapy and Games People Play (ego states, transactions, games and scripts).",
        "William Glasser \u2014 Reality Therapy: A New Approach to Psychiatry and Choice Theory (choice theory, basic needs and the WDEP system).",
        "Milton H. Erickson & Ernest L. Rossi \u2014 Hypnotherapy: An Exploratory Casebook (permissive, indirect and utilisation hypnotherapy).",
        "Virginia M. Axline \u2014 Play Therapy and Dibs: In Search of Self (non-directive, child-centred play therapy and its eight principles).",
        "Jacob L. Moreno \u2014 Psychodrama (spontaneity, role theory and the five instruments of action-based group therapy).",
    ]},
    {"t": "h3", "text": "Primary and online sources"},
    {"t": "ul", "items": [
        "Affective Therapies (Unit IV compiled reader) \u2014 the primary compiled reader for this unit, drawn on for each therapy's origin, technique and current status.",
        "American Psychological Association \u2014 APA Dictionary of Psychology (characterisations of 'psychotherapy', 'psychodynamic psychotherapy', 'person-centered therapy', 'gestalt therapy' and related entries).",
        "World Health Organization \u2014 mhGAP Intervention Guide (psychological interventions delivered through respectful, collaborative care).",
        "{{SEC}} Indian Association of Clinical Psychologists \u2014 secondary summaries only; primary text not accessed, so no clause numbers are quoted.",
    ]},
]

# ---------------------------------------------------------------------------
SECTIONS = []


def _s(sid, num, title, blocks):
    SECTIONS.append({"id": sid, "num": num, "title": title, "blocks": blocks})


# ========================= 1. PSYCHODYNAMIC PSYCHOTHERAPY ===========================
_s("t1", "1", "Psychodynamic psychotherapy", [
    {"t": "lead", "text": "Psychodynamic psychotherapy is the modern, briefer descendant of Freudian psychoanalysis that treats symptoms as meaningful compromises between unconscious wishes, reality and conscience, and uses the relationship, free association and interpretation to make the unconscious conscious and bring lasting change."},
    {"t": "fig", "key": "psychodynamic_model", "caption": "The Freudian model: unconscious conflict between id, ego and superego surfaces as a symptom, which therapy traces back to insight."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises [[psychodynamic psychotherapy|therapy that works with unconscious conflict and the therapy relationship]] as a form of therapy derived from psychoanalysis that focuses on **unconscious processes** as they show in present behaviour, aiming at self-awareness and insight into the influence of the past on the present. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} **Freud** founded psychoanalysis with Breuer's 'talking cure', moving from hypnosis and catharsis to **free association and interpretation**, and set its aim in the formula that where id was, there **ego shall be**. {{c:Freud}}",
        "{{THEORY}} Later schools extended the approach into **ego psychology, object relations, self psychology and attachment**, so that modern psychodynamic therapy is face-to-face and briefer than classical analysis while keeping its core ideas. {{c:Gabbard}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin** the approach begins with **Freud and Breuer** and the case of Anna O., the so-called talking cure, out of which Freud built psychoanalysis from hypnosis and catharsis toward free association and dream interpretation. {{c:Freud}}",
        "Its later **origin** runs through **Jung, Adler, Anna Freud, Klein, Winnicott, Kohut and Bowlby**, so that object relations, self psychology and attachment all grew from the one psychoanalytic root into today's briefer forms. {{c:Gabbard}}",
        "The first basis is [[psychic determinism|the view that nothing mental is accidental and every symptom has a cause]], the idea that behaviour and symptoms **have meaning and causes** rather than arising by chance, so that even slips of the tongue are read as significant. {{c:Freud}}",
        "A central basis is the **unconscious**, described through the topographic model of conscious, preconscious and unconscious and the structural model in which the **id, ego and superego** pursue pleasure, reality and conscience. {{c:Freud}}",
        "The developmental basis is the **psychosexual stages**, oral, anal, phallic, latency and genital, where excessive frustration or indulgence can cause [[fixation|getting stuck at an early developmental stage]] that later shapes personality and symptoms. {{c:McWilliams}}",
        "Anxiety and the **defence mechanisms** form another basis, so that **repression, denial, projection, displacement and sublimation** keep unacceptable wishes out of awareness while quietly steering conduct. {{c:McWilliams}}",
        "In **formulation**, a symptom is read as a [[compromise formation|a symptom that partly satisfies a wish and partly defends against it]] between a repressed wish, the demands of reality and the superego, so it symbolises an underlying conflict and yields secondary gain. {{c:McWilliams}}",
        "The guiding **goal** is to make the **unconscious conscious**, giving the client insight into hidden conflicts and patterns, strengthening the ego, maturing the defences and so producing lasting change rather than mere symptom removal. {{c:Freud}}",
        "The therapist's stance is a key **procedure**, combining [[neutrality|not taking sides among the client's wishes, defences and conscience]], abstinence and relative anonymity so the therapist acts as a **blank screen** onto which the client's patterns can be projected. {{c:Gabbard}}",
        "The fundamental **technique** is [[free association|saying whatever comes to mind without censoring it]], in which the client is asked to voice every passing thought, because the gaps and detours reveal where the real conflict lies. {{c:Freud}}",
        "**Dream analysis** is a classic technique, treating dreams as the royal road to the unconscious and distinguishing the remembered **manifest content** from the hidden **latent content** reached by exploring the dream's distortions. {{c:Freud}}",
        "**Interpretation** is the therapist's main tool, tactfully and at the right moment explaining the unconscious meaning of associations, dreams and symptoms, so that what was warded off can be put into words and understood. {{c:Gabbard}}",
        "Two further techniques are the **analysis of resistance**, noting lateness, silence or intellectualising, and the [[analysis of transference|exploring the client's transfer of old relationship patterns onto the therapist]], which lets old patterns be relived and understood in the room. {{c:Gabbard}}",
        "In **stages and process** the work moves from an **opening phase** of assessment, contract and alliance, through the emergence and **working through** of transference and resistance, to a planned **termination** that reviews and consolidates the gains. {{c:Gabbard}}",
        "A distinctive process idea is [[working through|repeatedly revisiting an insight until it produces real change]], the patient and slow repetition of insights across many sessions, since a single interpretation rarely shifts a deeply rooted pattern on its own. {{c:Gabbard}}",
        "The **outcome** sought is reduced symptoms, greater insight, more flexible and mature defences and healthier relationships, and gains often continue to grow after therapy ends in what is called a **sleeper effect**. {{c:Corey}}",
        "**Indications** favour clients with good ego strength, psychological-mindedness and motivation facing neurotic, personality or interpersonal problems, while acute psychosis, high suicidality and crisis states are **contraindications** better met by supportive work. {{c:Gabbard}}",
        "In **research and current status**, well-known reviews report that psychodynamic therapy produces **substantial and lasting benefit**, with short-term psychodynamic therapy supported for depression and anxiety and specialised forms helping borderline personality disorder, while classical long analysis remains hard to study. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Psychodynamic therapy at a glance", "items": [
        "Origin: Freud and Breuer, extended by ego psychology, object relations, self psychology and attachment.",
        "Basis: psychic determinism, the unconscious, id-ego-superego, psychosexual stages, defences.",
        "Formulation: the symptom is a compromise formation that symbolises unconscious conflict.",
        "Technique: free association, dream analysis, interpretation, analysis of resistance and transference.",
        "Process: opening, transference, working through, termination; indicated for ego-strong neurotic clients.",
    ]},
    {"t": "mn", "word": "FIDO-WT", "text": "Free association, Interpretation, Dream analysis, (analysis of) resistance and transferO, Working through and Termination \u2014 the techniques and stages of psychodynamic therapy; the aim throughout is to make the unconscious conscious."},
])

# ========================= 2. BRIEF PSYCHOTHERAPY ======================
_s("t2", "2", "Brief psychotherapy", [
    {"t": "lead", "text": "Brief psychotherapy is a planned, time-limited treatment that works on one agreed focus with an active, directive therapist, using the deadline itself to mobilise motivation, so that symptoms ease and functioning is restored without attempting a full personality overhaul."},
    {"t": "fig", "key": "brief_focus", "caption": "Brief therapy keeps a single focus across a short, time-limited course with a fixed end date and a mid-point review."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors and models", "items": [
        "{{THEORY}} Brief therapies share a **planned time limit, a single focus and an active therapist**, and they rest on the finding that most improvement happens early, with roughly half of clients improved within about eight sessions. {{c:Affective Therapies Reader}}",
        "{{THEORY}} Named models include **Malan's focal therapy, Sifneos' short-term anxiety-provoking therapy, Mann's time-limited therapy and Davanloo's intensive short-term dynamic therapy**, alongside interpersonal therapy and solution-focused brief therapy. {{c:Sharf}}",
        "{{GUIDE}} The **World Health Organization** promotes brief, structured psychological interventions in routine care, framing them as effective help delivered through **respectful, collaborative** conversation within a few sessions. {{c:WHO mhGAP}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, brief therapy grows from **Ferenczi and Rank's** active technique and set time limits in the 1920s and from **Alexander and French's** idea of a corrective emotional experience in 1946. {{c:Affective Therapies Reader}}",
        "Its **origin** then branches into many models, including **Malan, Sifneos, Mann and Davanloo** in the psychodynamic stream and **de Shazer and Berg's** solution-focused therapy in the systemic stream. {{c:Sharf}}",
        "A practical **origin** driver was service pressure, since cost, demand and the evidence that **most gains come early** pushed clinicians toward shorter, planned treatments. {{c:Affective Therapies Reader}}",
        "The first basis is that the **time limit is therapeutic**, because a known end date mobilises motivation, concentrates the work and guards against the regression and dependency that an open-ended course can breed. {{c:Affective Therapies Reader}}",
        "A second basis is a tight [[focus|a single agreed problem or theme that organises the whole treatment]], so the therapist is **active, directive and selective**, steering the sessions back to the one theme rather than following every new concern. {{c:Sharf}}",
        "A third basis is optimism about change, emphasising the **here-and-now, the client's strengths** and the assumption that small changes can set off larger ones, especially in the solution-focused models. {{c:Sharf}}",
        "In **formulation** the therapist builds a **focal formulation**, naming a recurring interpersonal pattern, a current loss or an unresolved conflict tied to the present problem, which keeps the brief work on target. {{c:Affective Therapies Reader}}",
        "A well-known formulation grid is **interpersonal therapy's four problem areas**, namely grief, role disputes, role transitions and interpersonal deficits, one of which is chosen as the focus. {{c:Sharf}}",
        "The **goals** are modest and clear, to **resolve the focal problem**, relieve symptoms and restore previous functioning, plus teaching the client skills to use later, rather than reconstructing the whole personality. {{c:Affective Therapies Reader}}",
        "A core **procedure** is careful **patient selection**, and Sifneos' criteria look for a circumscribed complaint, at least one meaningful past relationship, psychological-mindedness, motivation and the capacity to tolerate anxiety. {{c:Affective Therapies Reader}}",
        "The **contract** is itself a technique, typically **6 to 20 sessions**, often around twelve to sixteen, with a fixed end date agreed at the outset so that both parties work against the clock. {{c:Sharf}}",
        "Dynamic **techniques** include focused interpretation, selective attention and inattention, gentle challenge of defences and the deliberate use of **positive transference** to keep the client engaged with the focus. {{c:Sharf}}",
        "Solution-focused **techniques** are distinctive, among them the [[miracle question|asking how life would look if the problem vanished overnight]], scaling questions, hunting for exceptions and complimenting the client's existing resources. {{c:Sharf}}",
        "In **stages and process** the work runs from **assessment and selection** in the first few sessions, through an active **middle phase** on the focus with a mid-point review, to an explicit **termination** that consolidates gains. {{c:Affective Therapies Reader}}",
        "A defining process feature is that **termination is discussed from the start**, so the ending is planned, the gains are reviewed and a booster or follow-up session may be arranged rather than the therapy simply trailing off. {{c:Affective Therapies Reader}}",
        "The **outcome** is symptom relief and improved functioning, and many comparisons find brief therapy **as effective as longer therapy** for common disorders such as depression, anxiety and adjustment problems. {{c:Corey}}",
        "**Indications** are acute and focal problems, adjustment disorders, mild to moderate depression and anxiety, grief and crises in motivated clients, while chronic multi-problem conditions and severe high-risk states are **contraindications**. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is strong for **interpersonal therapy and brief cognitive-behavioural therapy** and supportive for short-term psychodynamic and solution-focused work, and brief models are now standard in primary care, student counselling and telehealth. {{c:Sharf}}",
    ]},
    {"t": "box", "kind": "key", "title": "Brief psychotherapy at a glance", "items": [
        "Origin: Ferenczi and Rank, Alexander and French; models by Malan, Sifneos, Mann, Davanloo, de Shazer.",
        "Basis: the time limit is therapeutic; a single focus; an active, directive therapist.",
        "Formulation: a focal theme; interpersonal therapy's grief, disputes, transitions and deficits.",
        "Technique: selection criteria, a fixed contract, focused interpretation, solution-focused questions.",
        "Process: assessment, focused middle phase, planned termination; strong evidence for IPT and brief CBT.",
    ]},
    {"t": "mn", "word": "FAST", "text": "Focus (one theme), Active and directive therapist, Selection (clear criteria) and Time limit (fixed end date) \u2014 the four features that define brief psychotherapy; the deadline itself is used to drive the work."},
])

# ========================= 3. HUMANISTIC PSYCHOTHERAPY ===================
_s("t3", "3", "Humanistic psychotherapy", [
    {"t": "lead", "text": "Humanistic psychotherapy is the third force in psychology, a growth-oriented movement that trusts the person's own actualising tendency, centres subjective experience, free will and the whole person, and makes the genuine therapeutic relationship, rather than technique, the main agent of change."},
    {"t": "fig", "key": "humanistic_hierarchy", "caption": "Maslow's hierarchy of needs: distress follows when blocked lower needs stop a person from reaching self-actualisation."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises [[humanistic therapy|an approach stressing growth, choice and the whole person]] as an approach arising from humanistic psychology that emphasises the person's **capacity for growth, self-direction and self-actualisation** and the quality of the therapeutic relationship. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} The movement began as the **third force** in the 1950s and 1960s, reacting against the determinism of psychoanalysis and the mechanism of behaviourism, and drew figures such as **Maslow, Rogers, May, Frankl and Perls**. {{c:Corey}}",
        "{{THEORY}} **Maslow's** hierarchy of needs, from physiological and safety needs up to [[self-actualisation|becoming the fullest version of oneself]], frames distress as what happens when growth needs are blocked. {{c:Corey}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, humanistic therapy emerged in the **1950s and 1960s** as the self-styled third force, a deliberate reaction against both psychoanalytic determinism and behaviourist mechanism, with the Association for Humanistic Psychology founded in 1961. {{c:Corey}}",
        "Its **origin** gathers several figures under one banner, since **Maslow, Rogers, May, Frankl and Perls** all stressed growth and choice, and its roots lie in phenomenology and existential philosophy. {{c:Corey}}",
        "The first basis is that people are **inherently good and growth-oriented**, carrying an [[actualising tendency|an innate drive toward growth and fulfilment]] that moves them toward fulfilment when conditions allow. {{c:Corey}}",
        "A second basis is the emphasis on **subjective experience, free will, choice and responsibility**, treating the person's own point of view and the here-and-now as the proper focus rather than drives or conditioning. {{c:Sharf}}",
        "A third basis is holism, the view of the **whole person** in context, so the approach resists reducing a human being to symptoms, reflexes or diagnostic labels. {{c:Corsini}}",
        "A fourth basis, from **Maslow**, is the hierarchy of needs, in which physiological, safety, belonging and esteem needs build toward self-actualisation, and blocked growth needs produce distress. {{c:Corey}}",
        "In **formulation**, distress is understood as **natural growth being blocked** by a hostile environment, by [[conditions of worth|approval given only when a person meets others' standards]], or by losing touch with one's own experience. {{c:Corey}}",
        "A key formulation idea is that symptoms reflect **disconnection from experience** rather than disease, so the aim is to restore awareness and self-direction rather than to correct a defect. {{c:Sharf}}",
        "The **goals** are **self-awareness, self-acceptance, authenticity** and personal growth, helping the person take responsibility and move toward self-actualisation in their own way. {{c:Corey}}",
        "A defining **procedure** is that the **relationship itself is the main agent of change**, so a genuine, accepting and empathic presence matters more than any specific technique. {{c:Corey}}",
        "Core **techniques** are **empathic listening, reflection and unconditional acceptance**, offered in a non-directive, client-led manner that follows the person rather than steering them. {{c:Sharf}}",
        "Experiential **techniques** are also used, including **focusing on bodily felt sense and chair work**, which help the person contact and process emotion in the present moment. {{c:Corey}}",
        "In **stages and process** the work is **client-led and unfolding** rather than staged by the therapist, moving as the person feels safe enough to explore, deepen awareness and reclaim ownership of their experience. {{c:Sharf}}",
        "A process hallmark is **genuine presence and congruence** from the therapist, which lowers the client's defences and frees the actualising tendency to resume its natural course. {{c:Corey}}",
        "The **outcome** sought is increased **self-esteem, congruence and emotional awareness**, together with greater relationship satisfaction and a more internal sense of direction. {{c:Corey}}",
        "**Indications** cover mild to moderate distress, identity and relationship problems, depression and anxiety and clients seeking growth, while severe psychosis or a need for structured skills training are better served **as a contraindication** by more directive work. {{c:Sharf}}",
        "In **research and current status**, well-known meta-analyses by **Elliott and colleagues** report that humanistic-experiential therapies produce large, stable gains and are broadly equivalent to cognitive-behavioural therapy in many comparisons. {{c:Corey}}",
        "The approach remains alive in modern forms such as **emotion-focused therapy and motivational interviewing**, though critics note that its concepts are vague, hard to test and perhaps over-optimistic about human nature. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Humanistic therapy at a glance", "items": [
        "Origin: the 1950s to 1960s third force; Maslow, Rogers, May, Frankl, Perls; phenomenological roots.",
        "Basis: innate goodness and an actualising tendency; subjective experience, free will, the whole person.",
        "Formulation: distress is blocked growth, conditions of worth and disconnection from experience.",
        "Technique: empathic listening, reflection, unconditional acceptance; relationship over technique.",
        "Outcome and research: growth and congruence; Elliott et al. show large gains, equivalent to CBT.",
    ]},
    {"t": "mn", "word": "GROWTH", "text": "Growth-oriented view, Relationship as the agent of change, Own actualising tendency, Whole person, Third force and Here-and-now \u2014 the humanistic essentials; it trusts the person to move toward self-actualisation."},
])

# ========================= 4. EXISTENTIAL PSYCHOTHERAPY =============================
_s("t4", "4", "Existential psychotherapy", [
    {"t": "lead", "text": "Existential psychotherapy helps people face the unavoidable givens of existence, death, freedom, isolation and meaninglessness, so that the anxiety these provoke is met honestly and turned into authentic, responsible and meaningful living rather than avoided through symptoms."},
    {"t": "fig", "key": "existential_concerns", "caption": "Yalom's four ultimate concerns: facing death, freedom, isolation and meaninglessness honestly turns existential anxiety into authentic living."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Yalom** frames existential therapy around four [[ultimate concerns|the unavoidable givens of existence that provoke anxiety]], namely **death, freedom, isolation and meaninglessness**, and holds that confronting them produces anxiety that avoidance turns into symptoms. {{c:Yalom}}",
        "{{THEORY}} **Frankl's** logotherapy centres the **will to meaning**, and his Man's Search for Meaning, formed from his concentration-camp experience, argues that meaning can be found even in unavoidable suffering. {{c:Frankl}}",
        "{{THEORY}} The approach draws on **Kierkegaard, Nietzsche, Heidegger and Sartre** and on clinicians such as **Binswanger, Boss and May**, giving it a strongly philosophical character. {{c:Corey}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, existential therapy has philosophical roots in **Kierkegaard, Nietzsche, Heidegger and Sartre** and clinical roots in Binswanger and Boss's Daseinsanalysis, which applied phenomenology to the suffering person. {{c:Corey}}",
        "Its clinical **origin** runs through **Frankl's** logotherapy from the 1930s, **May's** Existence in 1958 and **Yalom's** Existential Psychotherapy in 1980, which organised the field for English-speaking clinicians. {{c:Yalom}}",
        "The first basis is that existence involves unavoidable **ultimate concerns**, so that death, freedom, isolation and meaninglessness confront everyone and generate an [[existential anxiety|the dread that arises from facing the basic conditions of existence]] that cannot simply be cured away. {{c:Yalom}}",
        "A second basis is that **existence precedes essence**, meaning we create ourselves through our choices, so that [[authenticity|living in accordance with one's own values and choices]] stands against living in bad faith. {{c:Corey}}",
        "A third basis, from **Heidegger**, is being-in-the-world across the physical, social and personal worlds, together with Frankl's conviction that a **will to meaning** is the deepest human motive. {{c:Frankl}}",
        "In **formulation**, psychopathology is read as **inauthentic living**, the avoidance of freedom and responsibility, loss of meaning or failure to cope with death and isolation. {{c:Corey}}",
        "Frankl's distinctive formulation is the [[existential vacuum|a sense of emptiness and boredom from loss of meaning]], an emptiness and boredom arising from lost meaning that he called noogenic neurosis. {{c:Frankl}}",
        "The **goals** are to help clients **face life's givens honestly** and to increase awareness, responsibility, authenticity and meaning, committing to action even in the presence of anxiety. {{c:Yalom}}",
        "A core **procedure** is **phenomenological exploration** of the client's own worldview, with the therapist bracketing assumptions and using Socratic dialogue to examine values, choices and avoidances. {{c:Corey}}",
        "The here-and-now **relationship** is used as a laboratory, and encouraging **'I' statements** that own experience helps the client take responsibility rather than treating problems as things that merely happen to them. {{c:Yalom}}",
        "A signature logotherapy **technique** is [[paradoxical intention|deliberately trying to do the very thing one fears]], in which the client intends the feared symptom, which often loosens the anticipatory anxiety behind phobias and insomnia. {{c:Frankl}}",
        "A second logotherapy **technique** is [[dereflection|shifting attention away from the self toward meaningful activity]], turning the client's focus from anxious self-monitoring onto meaningful activity and other people. {{c:Frankl}}",
        "A third technique is **attitudinal modification**, helping the client find meaning in unavoidable suffering, which is central when the situation itself cannot be changed. {{c:Frankl}}",
        "In **stages and process** the work establishes a **genuine relationship** and explores the worldview, then identifies patterns of avoidance and inauthenticity, deepens awareness of freedom and responsibility, and ends in committed action. {{c:Yalom}}",
        "The process is **collaborative and exploratory** rather than technique-driven, so the therapist accompanies the client in facing hard truths rather than fixing a defect. {{c:Corey}}",
        "The **outcome** sought is greater **meaning, purpose and authenticity** and an increased ability to tolerate anxiety, so the person lives more deliberately in the face of life's givens. {{c:Yalom}}",
        "**Indications** include existential crises such as life transitions, midlife and bereavement, terminal or chronic illness, and depression marked by emptiness, while acute psychosis and a wish for quick symptom control are relative **contraindications**. {{c:Corey}}",
        "In **research and current status**, the evidence is **moderate**, with meaning-centred interventions showing benefit for meaning, mood and quality of life, especially in people with cancer, while critics note the concepts are abstract and hard to operationalise. {{c:Yalom}}",
    ]},
    {"t": "box", "kind": "key", "title": "Existential therapy at a glance", "items": [
        "Origin: Kierkegaard to Sartre; Binswanger, Boss, Frankl, May and Yalom.",
        "Basis: four ultimate concerns (death, freedom, isolation, meaninglessness); existence precedes essence.",
        "Formulation: inauthentic living, avoided responsibility and the existential vacuum.",
        "Technique: phenomenological and Socratic dialogue; paradoxical intention, dereflection, attitude change.",
        "Outcome and research: meaning and authenticity; moderate evidence, strong in meaning-centred care.",
    ]},
    {"t": "mn", "word": "DFIM", "text": "Death, Freedom, Isolation and Meaninglessness \u2014 Yalom's four ultimate concerns at the heart of existential therapy; facing them honestly is what turns dread into authentic living."},
])

# ========================= 5. GESTALT THERAPY =============================
_s("t5", "5", "Gestalt therapy", [
    {"t": "lead", "text": "Gestalt therapy is an experiential, here-and-now approach developed by Fritz and Laura Perls in which awareness itself is curative, the therapist sets up experiments rather than giving interpretations, and the client completes unfinished business and reowns split-off parts to become more whole."},
    {"t": "fig", "key": "gestalt_cycle", "caption": "The Gestalt contact cycle: a need becomes figure, mobilises energy, reaches contact and withdraws; disturbances interrupt the cycle."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises [[gestalt therapy|an experiential therapy centred on present awareness and the whole person]] as an experiential approach that emphasises **present experience and awareness**, integrating feelings, perceptions and actions into the whole person. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} It was developed by **Fritz Perls, Laura Perls and Paul Goodman** from psychoanalysis, Gestalt psychology and field theory, and was made popular in the 1960s at the Esalen Institute. {{c:Corey}}",
        "{{THEORY}} Its guiding idea is the [[paradoxical theory of change|change happens when a person accepts what they are, not when they force what they are not]], which Beisser stated as change coming when a person fully becomes what they are. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, Gestalt therapy was developed by **Fritz and Laura Perls with Paul Goodman** in the 1940s and 1950s, drawing on psychoanalysis, Gestalt psychology, Reich's body work and Lewin's field theory. {{c:Corey}}",
        "Its **origin** reached a wide public in the **1960s at Esalen**, where Perls' charismatic demonstrations made the empty-chair style famous, for better and worse. {{c:Affective Therapies Reader}}",
        "The first basis is **holism and field theory**, the view that person and environment form a single field, so behaviour is understood in context rather than in isolation. {{c:Corey}}",
        "A central basis is the **here-and-now**, focusing on the **what and how** of present experience rather than the why, because awareness in the present moment is held to be curative. {{c:Corey}}",
        "A further basis is **figure and ground**, in which the dominant need becomes the [[figure|the need or interest that stands out from the background at a given moment]] against the background, guided by organismic self-regulation. {{c:Sharf}}",
        "Another basis is [[unfinished business|unexpressed feelings from the past that intrude on the present]], the unexpressed feelings that press for attention and interfere with living fully in the present. {{c:Corey}}",
        "In **formulation**, disturbance arises from **interruptions to contact** at the boundary, named as introjection, projection, [[retroflection|turning back onto oneself what one would like to do to or get from others]], deflection and confluence. {{c:Corey}}",
        "A further formulation idea is **Perls' layers of neurosis**, the phony, phobic, impasse, implosive and explosive layers through which a client passes toward authentic contact. {{c:Affective Therapies Reader}}",
        "A vivid formulation is the **top-dog versus underdog** split, the inner conflict between the demanding shoulds and the resistant wants that keeps a person stuck. {{c:Corey}}",
        "The **goals** are to **increase awareness**, take responsibility, integrate split-off parts, complete unfinished business and shift from relying on the environment toward self-support. {{c:Corey}}",
        "The best-known **technique** is the [[empty-chair technique|speaking to an imagined person or part of oneself placed in an empty chair]], a dialogue with an absent person or a part of the self that brings a conflict into the room. {{c:Corey}}",
        "Further **techniques** include **exaggerating a gesture, reversal and staying with the feeling**, each designed as an experiment that heightens awareness rather than as an interpretation. {{c:Sharf}}",
        "Language **techniques** matter too, replacing 'it' and 'you' with **'I' statements** and 'I can't' with 'I won't', so the client owns experience and choice in the present tense. {{c:Corey}}",
        "**Dream work** is distinctive, treating every part of a dream as a projected piece of the dreamer, who then plays each part to reclaim it. {{c:Corey}}",
        "In **stages and process** the work moves from **contact and the I-Thou relationship**, through awareness work and experiments that heighten contact, to integration and closure of what was unfinished. {{c:Corey}}",
        "The therapist's **process** stance is **authentic, present and non-interpretative**, offering experiments and honest feedback rather than analysing from a distance. {{c:Corey}}",
        "The **outcome** sought is greater **awareness, responsibility and emotional expression**, fuller integration and more vital interpersonal functioning. {{c:Corey}}",
        "**Indications** include anxiety, depression, interpersonal and psychosomatic problems and trauma-related unfinished business, while very confrontational experiments are a **contraindication** for fragile or psychotic clients who may be overwhelmed. {{c:Sharf}}",
        "In **research and current status**, the empirical base is **limited but growing**, and the strongest support comes through emotion-focused therapy, where Greenberg and colleagues show that empty-chair work for unfinished business is effective. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Gestalt therapy at a glance", "items": [
        "Origin: Fritz and Laura Perls with Goodman; Gestalt psychology and field theory; Esalen in the 1960s.",
        "Basis: holism, the here-and-now, figure and ground, unfinished business, awareness is curative.",
        "Formulation: contact disturbances, layers of neurosis, top-dog versus underdog.",
        "Technique: the empty chair, exaggeration, 'I' language, dream work; experiments not interpretations.",
        "Outcome and research: awareness and integration; limited base, supported via emotion-focused therapy.",
    ]},
    {"t": "mn", "word": "AWARE", "text": "Awareness is curative, Who-and-how not why, Appropriate experiments (empty chair), Responsibility and 'I' language, and Ending unfinished business \u2014 the Gestalt essentials in the here-and-now."},
])

# ========================= 6. PERSON-CENTRED THERAPY ===========================
_s("t6", "6", "Person-centred (client-centred) therapy", [
    {"t": "lead", "text": "Person-centred therapy is Carl Rogers' non-directive approach in which the therapist's congruence, unconditional positive regard and empathy are held to be the necessary and sufficient conditions that free the client's actualising tendency, so the client, as the expert on themselves, directs their own change."},
    {"t": "fig", "key": "rogers_conditions", "caption": "Rogers' core conditions: congruence, unconditional positive regard and empathy free the client's actualising tendency."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** characterises [[person-centered therapy|Rogers' non-directive therapy built on empathy, congruence and acceptance]] as the non-directive form of therapy developed by **Carl Rogers**, in which the therapist's genuineness, unconditional positive regard and empathic understanding foster the client's growth. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} **Rogers** introduced non-directive counselling in 1942 and set out his **six necessary and sufficient conditions** of therapeutic change in 1957, grounding the approach in the actualising tendency. {{c:Rogers}}",
        "{{THEORY}} The approach is firmly **humanistic yet distinct**, since it is Rogers' specific model built on the core conditions, whereas humanistic therapy names the broader third-force movement. {{c:Corey}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, the approach is the work of **Carl Rogers**, whose Counseling and Psychotherapy of 1942 introduced non-directive counselling and whose Client-Centered Therapy of 1951 named the model. {{c:Rogers}}",
        "Rogers later broadened it into the **person-centred approach** applied to education, groups and conflict resolution, and he pioneered recording and studying sessions, which began an empirical tradition in therapy research. {{c:Corey}}",
        "The foundational basis is the [[actualising tendency|the single inbuilt drive of the organism toward growth and fulfilment]], the single motivating force that moves every organism toward growth and fulfilment when conditions allow. {{c:Rogers}}",
        "A second basis is the **self-concept and the ideal self**, together with the organismic valuing process by which a healthy person trusts their own experiencing as a guide. {{c:Rogers}}",
        "A third basis is [[conditions of worth|the sense that one is valued only when meeting others' standards]], formed when love and approval are conditional, which lead the person to distort or deny experiences that conflict with them. {{c:Corey}}",
        "A fourth basis is **incongruence**, the gap between self-concept and experience, or between real and ideal self, which produces anxiety and defensiveness, so that restoring congruence is the mark of health. {{c:Rogers}}",
        "In **formulation**, problems are understood as **incongruence caused by conditions of worth**, so therapy is not a technical repair but the provision of conditions in which the actualising tendency can resume. {{c:Corey}}",
        "The decisive formulation claim is **Rogers' six conditions**, of which the three therapist-offered core conditions are congruence, unconditional positive regard and empathy. {{c:Rogers}}",
        "The first core condition is [[congruence|the therapist being genuine and integrated, not playing a role]], the therapist being real and integrated in the relationship rather than hiding behind a professional facade. {{c:Rogers}}",
        "The second core condition is [[unconditional positive regard|accepting and prizing the client without conditions]], a warm, non-judgemental acceptance of the client as a person whatever they bring. {{c:Rogers}}",
        "The third core condition is **empathy**, an accurate understanding of the client's internal frame of reference that is communicated back so the client feels deeply understood. {{c:Corey}}",
        "The guiding **goals** are greater **congruence and self-acceptance**, an internal locus of evaluation and movement toward becoming a fully functioning person who is open to experience. {{c:Rogers}}",
        "The defining **procedure** is being **non-directive**, so the therapist does not interpret, diagnose, advise or reassure, but trusts the client as the expert on themselves. {{c:Corey}}",
        "The core **techniques**, insofar as the model allows them, are **active and reflective listening**, reflection of feeling, paraphrase, clarification and the acceptance of silences, offered within a genuine relationship. {{c:Sharf}}",
        "In **stages and process**, Rogers described a **seven-stage continuum** along which the client moves from rigid remoteness from feelings toward fluid, immediate and trusted experiencing. {{c:Corey}}",
        "In practice the **process** is that the client talks while the therapist reflects, and as the client feels understood and accepted, defences lower, **self-exploration deepens** and self-directed change emerges. {{c:Corey}}",
        "The **outcome** sought is increased self-acceptance and congruence, an internal locus of evaluation, reduced anxiety and defensiveness and improved, self-directed relationships. {{c:Rogers}}",
        "**Indications** are wide, covering mild to moderate emotional, identity and relationship problems, grief, adolescents and student counselling, while clients wanting direct advice, severe disorders or acute crises are relative **contraindications**. {{c:Sharf}}",
        "In **research and current status**, Rogers was a **pioneer of therapy research**, and meta-analyses by Elliott and colleagues find that empathy, positive regard and congruence correlate modestly but reliably with outcome, though critics challenge the claim that the conditions are sufficient on their own. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Person-centred therapy at a glance", "items": [
        "Origin: Carl Rogers (1942 non-directive counselling, 1951 client-centred therapy, 1957 six conditions).",
        "Basis: the actualising tendency, self-concept, conditions of worth and incongruence.",
        "Formulation: problems are incongruence from conditions of worth; provide the core conditions.",
        "Technique: non-directive; congruence, unconditional positive regard and empathy; reflective listening.",
        "Outcome and research: congruence and self-direction; Rogers pioneered research; conditions aid outcome.",
    ]},
    {"t": "mn", "word": "CUE", "text": "Congruence, Unconditional positive regard and Empathy \u2014 Rogers' three therapist-offered core conditions in person-centred therapy; given these, the client's own actualising tendency does the healing."},
])

# ========================= 7. ADLERIAN THERAPY =========================
_s("t7", "7", "Adlerian therapy (Individual Psychology)", [
    {"t": "lead", "text": "Adlerian therapy, or Individual Psychology, is Alfred Adler's approach that views the person as a whole, goal-directed and socially embedded, treats problems as discouragement and mistaken lifestyle beliefs rather than illness, and re-educates the client toward courage, social interest and useful contribution."},
    {"t": "fig", "key": "adlerian_phases", "caption": "The four Adlerian phases: relationship, lifestyle assessment, insight and reorientation toward social interest."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Adler** founded **Individual Psychology** after breaking from Freud's circle in 1911, stressing that the person is an indivisible whole whose behaviour is **goal-directed** and shaped by social context. {{c:Adler}}",
        "{{THEORY}} Its hallmark concept is [[social interest|an innate capacity and willingness to cooperate and contribute to others]], which Adler treated as the very measure of mental health. {{c:Adler}}",
        "{{THEORY}} The approach was carried forward by **Dreikurs, Ansbacher, Mosak and Dinkmeyer** into child guidance, family work and parent education. {{c:Corey}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, Adlerian therapy was founded by **Alfred Adler**, a Viennese physician and early member of Freud's circle who broke away in 1911 to establish Individual Psychology. {{c:Adler}}",
        "Its **origin** also includes **Rudolf Dreikurs**, who developed Adler's ideas for children, families and the democratic classroom, and Adler's founding of the first child-guidance clinics in 1920s Vienna. {{c:Corey}}",
        "The first basis is **holism and teleology**, the view that the person is indivisible and that behaviour is pulled toward a future goal, a [[fictional finalism|a guiding self-created goal that directs a person's striving]] that organises striving. {{c:Corey}}",
        "A central basis is **inferiority and compensation**, so that natural feelings of inferiority spur a striving for superiority or significance, which can become either healthy ambition or a self-defeating complex. {{c:Adler}}",
        "A defining basis is **social interest**, the innate capacity to cooperate and contribute, which Adler made the yardstick of mental health, so that well-being grows with connection to others. {{c:Adler}}",
        "A further basis is the [[lifestyle|the unique pattern of beliefs and goals formed early in childhood]], the unique pattern of beliefs and goals formed by about age five from the family constellation and birth order. {{c:Corey}}",
        "Another basis is the contrast between **private logic and common sense**, where idiosyncratic mistaken beliefs, or basic mistakes, steer a person away from cooperative living. {{c:Sharf}}",
        "In **formulation**, problems arise from **discouragement, mistaken beliefs and low social interest**, so difficulties are seen as problems of living rather than illness, as a way of life that is not working. {{c:Corey}}",
        "A practical formulation tool is **Dreikurs' four mistaken goals** of children's misbehaviour, namely attention, power, revenge and a display of inadequacy. {{c:Corey}}",
        "The **goals** are to **increase social interest and belonging**, correct mistaken lifestyle beliefs, build courage and self-worth and replace inferiority striving with useful contribution. {{c:Adler}}",
        "A core **procedure** is the **four-phase structure** named by Dreikurs, beginning with establishing an equal, respectful and collaborative relationship. {{c:Corey}}",
        "The second phase is **lifestyle assessment**, analysing the family constellation, birth order and especially [[early recollections|the earliest memories, read as clues to current beliefs]], which are read as projective clues to present beliefs. {{c:Corey}}",
        "The third phase is **insight**, in which the therapist discloses and interprets the client's mistaken beliefs, private logic and hidden goals in a warm, tentative way. {{c:Corey}}",
        "The fourth phase is **reorientation**, turning insight into action through new choices, supported by **encouragement**, the central Adlerian technique. {{c:Corey}}",
        "Distinctive **techniques** include [[acting as if|behaving as though one already had the desired quality]], **spitting in the client's soup** to make a payoff unattractive, catching oneself, the push-button technique and setting tasks and homework. {{c:Sharf}}",
        "In **stages and process**, the therapist works as an **encouraging teacher or coach**, treating therapy as an educational and re-educational process rather than a cure of disease. {{c:Corey}}",
        "The **outcome** sought is reduced inferiority and discouragement, increased courage, greater **social interest** and a more adaptive lifestyle and relationships. {{c:Adler}}",
        "**Indications** are broad, covering children and adolescents, parent education and family therapy, school and career counselling and group work, including collectivist cultures whose values fit social interest, while severe psychopathology needs **adaptation**. {{c:Corey}}",
        "In **research and current status**, the evidence is **limited but positive**, with support for Adlerian parent education such as STEP, Adlerian play therapy and counselling, and Adler's ideas are recognised as anticipating cognitive, humanistic and family-systems therapies. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Adlerian therapy at a glance", "items": [
        "Origin: Alfred Adler (Individual Psychology, 1911); Dreikurs, Ansbacher, Mosak, Dinkmeyer.",
        "Basis: holism and teleology, inferiority and compensation, social interest, lifestyle and birth order.",
        "Formulation: discouragement and mistaken beliefs; problems of living, not illness.",
        "Technique: four phases (relationship, assessment, insight, reorientation); encouragement, acting as if.",
        "Outcome and research: courage and social interest; limited but positive evidence; anticipated CBT.",
    ]},
    {"t": "mn", "word": "RAIR", "text": "Relationship, Assessment (lifestyle), Insight and Reorientation \u2014 Dreikurs' four Adlerian phases; the engine throughout is encouragement, aiming the client toward social interest."},
])

# ========================= 8. TRANSACTIONAL ANALYSIS ======================
_s("t8", "8", "Transactional analysis (TA)", [
    {"t": "lead", "text": "Transactional analysis is Eric Berne's social-psychiatry approach that reads personality as three ego states, Parent, Adult and Child, and treats problems as contaminated ego states, repetitive games and a limiting life script, aiming by a clear contract to restore autonomy and game-free intimacy."},
    {"t": "fig", "key": "ta_states", "caption": "The three ego states and the four types of transaction, with the Karpman drama triangle that games play out."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Eric Berne** founded transactional analysis in the late 1950s, setting out its theory in Transactional Analysis in Psychotherapy and popularising it in Games People Play, and he defined the aim of therapy as **autonomy** through awareness, spontaneity and intimacy. {{c:Berne}}",
        "{{THEORY}} A [[transaction|a single exchange of a stimulus and a response between two people]] is the unit of social interaction, and Berne read every exchange as issuing from one of three **ego states**, the Parent, the Adult and the Child. {{c:Berne}}",
        "{{THEORY}} The approach was carried forward by **Harris, Steiner, Karpman and the Gouldings**, who added the OK life positions, scripts, the drama triangle and redecision therapy. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, transactional analysis was created by **Eric Berne**, a Canadian-born psychiatrist who broke from psychoanalysis to build a plain-language social psychiatry, publishing Transactional Analysis in Psychotherapy in 1961 and Games People Play in 1964. {{c:Berne}}",
        "Its **origin** then widened through **Thomas Harris**, whose I'm OK, You're OK spread the life positions, **Claude Steiner** on scripts, **Stephen Karpman** on the drama triangle and the **Gouldings** on redecision therapy. {{c:Affective Therapies Reader}}",
        "The core basis is the three [[ego states|consistent patterns of feeling and behaviour, called Parent, Adult and Child]], where the **Parent** is taught and judging, the **Adult** is thinking and objective and the **Child** is feeling and spontaneous. {{c:Berne}}",
        "A second basis is the analysis of **transactions**, which may be **complementary** and so keep talk flowing, **crossed** and so break it off, or [[ulterior|carrying a hidden psychological message under the social one]] and so drive games. {{c:Sharf}}",
        "A third basis is **strokes**, the units of recognition people hunger for, so that positive and negative strokes, conditional and unconditional, shape how a person structures time from withdrawal through rituals and pastimes to games and intimacy. {{c:Berne}}",
        "A fourth basis is **games and rackets**, which are [[games|repetitive ulterior transactions that end in a predictable bad feeling]] such as 'Why don't you, Yes but', paying off in a familiar racket feeling that props up the script. {{c:Berne}}",
        "A fifth basis is the [[life script|an unconscious life plan decided in childhood from parental messages]], built from injunctions such as 'don't be close', counter-injunction drivers such as 'be perfect' and the child's own early decisions. {{c:Affective Therapies Reader}}",
        "A sixth basis is the four **life positions** of the OK Corral, running from the healthy 'I'm OK, you're OK' to the blaming, the depressive and the hopeless positions that colour a person's whole outlook. {{c:Affective Therapies Reader}}",
        "In **formulation**, problems are traced to **contaminated or excluded ego states**, destructive games, a negative script and an unhealthy life position, so the therapist maps which ego state and game is driving the complaint. {{c:Sharf}}",
        "A vivid formulation tool is **Karpman's drama triangle**, in which clients switch between the roles of Persecutor, Victim and Rescuer, and the healthier exit is assertiveness, vulnerability and genuine caring. {{c:Affective Therapies Reader}}",
        "The **goals** are **autonomy and game-free intimacy**, a free and uncontaminated Adult, a revised script sometimes called a script cure, better social control and relief of the presenting symptom. {{c:Berne}}",
        "A central **procedure** is the **contract**, a clear and mutual agreement on goals and responsibilities that gives TA its collaborative, adult-to-adult, 'OK, OK' stance from the very first session. {{c:Sharf}}",
        "The main **technique** is the **four kinds of analysis**, namely structural analysis of ego states, transactional analysis of exchanges, game analysis of hidden payoffs and script analysis of the life plan. {{c:Berne}}",
        "Berne set out eight **therapeutic operations**, from interrogation, specification and confrontation through explanation, illustration, confirmation and interpretation to [[crystallisation|the moment the client can hold a new, freeing position for themselves]]. {{c:Berne}}",
        "Further **techniques** include the **egogram**, permissions and protection, **redecision work** that re-experiences and changes early decisions through chair work, stroke work, role play and homework, usually in a group format. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work runs from **contracting and assessment**, through strengthening and decontaminating the Adult and the deconfusion and redecision of the Child, to relearning new behaviours and finally autonomy and termination. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is greater autonomy, clearer communication, an **end to games**, a revised script and improved self-esteem, with the simple language letting clients keep using the ideas after therapy. {{c:Sharf}}",
        "**Indications** cover neurotic, relationship and communication problems, personality issues, couples and families, groups, education and organisations, while acute psychosis is a **contraindication** served only by adapted forms. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is **modest**, with some reviews reporting positive outcomes but few rigorous trials, yet TA stays widely taught and accredited by the ITAA and EATA despite criticism of its jargon and limited validation. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Transactional analysis at a glance", "items": [
        "Origin: Eric Berne (1961 to 1964); extended by Harris, Steiner, Karpman and the Gouldings.",
        "Basis: Parent-Adult-Child ego states, transactions, strokes, games, script, OK life positions.",
        "Formulation: contaminated ego states, games, negative script; Karpman's drama triangle.",
        "Technique: the contract; structural, transactional, game and script analysis; redecision chair work.",
        "Process: contract, decontaminate Adult, deconfuse Child, relearn, autonomy; modest evidence base.",
    ]},
    {"t": "mn", "word": "PAC", "text": "Parent, Adult and Child \u2014 the three ego states at the heart of transactional analysis; the contract and the four analyses (structural, transactional, game and script) work to free the Adult and reach autonomy."},
])

# ========================= 9. REALITY THERAPY ======================
_s("t9", "9", "Reality therapy (and choice theory)", [
    {"t": "lead", "text": "Reality therapy is William Glasser's present-focused approach built on choice theory, which holds that all behaviour is chosen to meet five basic needs, and uses the WDEP system within a warm, non-blaming relationship to help clients evaluate and plan more effective, responsible choices."},
    {"t": "fig", "key": "reality_wdep", "caption": "The WDEP cycle inside a trusting environment: Wants, Doing, self-Evaluation and Planning, driving more effective choices."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **William Glasser** developed reality therapy in the 1960s and later grounded it in [[choice theory|the view that we choose almost all of our behaviour to meet basic needs]], the idea that we can directly control only our own acting and thinking. {{c:Glasser}}",
        "{{THEORY}} Glasser held that **most psychological problems are relationship problems**, rooted in unsatisfying connections, and so he rejected diagnostic labels in favour of helping clients make more effective choices. {{c:Glasser}}",
        "{{THEORY}} **Robert Wubbolding** systematised the practice as the **WDEP system** set inside a cycle of counselling that first builds a trusting, friendly environment. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, reality therapy was founded by **William Glasser**, an American psychiatrist who, dissatisfied with Freudian therapy, developed his ideas working at the Ventura School for Girls and a veterans' hospital, publishing Reality Therapy in 1965. {{c:Glasser}}",
        "Its **origin** continued as Glasser's control theory of the 1980s was renamed **choice theory** in 1998, and **Robert Wubbolding** built the WDEP system and the cycle of counselling that organise practice today. {{c:Affective Therapies Reader}}",
        "The first basis is that **all behaviour is chosen**, since we can directly control only our acting and thinking, which in turn steer our feeling and physiology, a picture Glasser likened to the four wheels of a car. {{c:Glasser}}",
        "A second basis is the five genetic [[basic needs|survival, love and belonging, power, freedom and fun]], which drive all behaviour, with love and belonging treated as the most important for mental health. {{c:Glasser}}",
        "A third basis is the [[quality world|an inner picture album of the people and things a person most wants]], against which we compare the perceived world, so that a gap between the two produces frustration and behaviour. {{c:Affective Therapies Reader}}",
        "A fourth basis is that **relationships are central**, so the seven deadly habits of criticising, blaming and complaining damage connection while the seven caring habits of supporting, listening and respecting rebuild it. {{c:Affective Therapies Reader}}",
        "A fifth basis is **responsibility and a present focus**, meeting one's needs without stopping others from meeting theirs, with no excuses, no blame and no punishment, and almost no attention paid to the past or to symptoms. {{c:Glasser}}",
        "In **formulation**, the client is seen as having an **unmet need**, usually love and belonging, and choosing ineffective behaviours such as withdrawing, 'depressing' or blaming to try to satisfy it, forming a failure identity. {{c:Affective Therapies Reader}}",
        "The **goals** are to help clients **take responsibility and make effective choices**, improve their relationships and meet needs in ways that are responsible, realistic and right, Glasser's three Rs. {{c:Glasser}}",
        "The core **procedure** is **Wubbolding's cycle of counselling**, which first builds the environment, a trusting and friendly bond with humour and good listening while avoiding arguing and criticising. {{c:Affective Therapies Reader}}",
        "The main **technique** is the **WDEP system**, where W explores wants and needs and D clarifies current total behaviour and direction. {{c:Affective Therapies Reader}}",
        "The WDEP **technique** continues with E, the client's own **self-evaluation** of whether their behaviour is helping and is realistic, and P, a plan that is committed to and workable. {{c:Affective Therapies Reader}}",
        "A memorable planning **technique** is the [[SAMI2C3 plan|a plan that is simple, attainable, measurable, immediate, involving, controlled, committed and consistent]], which makes the client's commitment concrete and checkable. {{c:Affective Therapies Reader}}",
        "Further **techniques** include skilful questioning, metaphor and humour, paradox, role play, contracts and self-disclosure, with the gentle confrontation 'Is what you are doing helping?' and a firm refusal to accept excuses. {{c:Affective Therapies Reader}}",
        "In **stages and process** the therapist builds the environment, establishes what the client wants, explores current behaviour, guides self-evaluation, helps form and commit to a plan, and then **follows up**, revising the plan if it fails. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is greater responsibility, a stronger **internal locus of control**, better relationships, higher self-esteem and more satisfying behaviour, achieved in a short and practical course of work. {{c:Affective Therapies Reader}}",
        "**Indications** are broad, covering schools, adolescents and offenders, corrections, addictions, marital and group counselling and the workplace, while severe mental illness or trauma needing exploration is a **caution**. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is **moderate but often quasi-experimental**, with studies in schools, correctional and cross-cultural settings showing gains in self-esteem, locus of control and relationships, though critics attack its rejection of diagnosis and simple view of choice. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Reality therapy at a glance", "items": [
        "Origin: William Glasser (Reality Therapy, 1965; Choice Theory, 1998); Wubbolding's WDEP and cycle.",
        "Basis: all behaviour is chosen; five basic needs; the quality world; relationships are central.",
        "Formulation: an unmet need and ineffective chosen behaviours; a failure identity.",
        "Technique: build the environment, then Wants, Doing, Evaluation and Planning (a SAMI2C3 plan).",
        "Process: present-focused, no excuses, no punishment; moderate, mostly quasi-experimental evidence.",
    ]},
    {"t": "mn", "word": "WDEP", "text": "Wants, Doing (and direction), Evaluation and Planning \u2014 Wubbolding's four procedures of reality therapy, applied inside a warm, non-blaming environment to help clients choose more effective, responsible behaviour."},
])

# ========================= 10. SUPPORTIVE PSYCHOTHERAPY ======================
_s("t10", "10", "Supportive psychotherapy", [
    {"t": "lead", "text": "Supportive psychotherapy is the ego-strengthening end of the supportive-expressive continuum, which aims not to uncover the unconscious but to maintain and restore ego functions, self-esteem and coping through a warm, active relationship, using reassurance, advice, praise and problem solving."},
    {"t": "fig", "key": "supportive_continuum", "caption": "The supportive to expressive continuum: supportive work bolsters existing defences; expressive work uncovers the unconscious."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} Supportive psychotherapy sits at the **supportive end of Luborsky's supportive-expressive continuum**, aiming to maintain or improve ego functions and self-esteem rather than to uncover unconscious conflict. {{c:Affective Therapies Reader}}",
        "{{THEORY}} It treats the **therapeutic relationship as the main agent of change**, with the therapist acting as a trusted, caring support or [[auxiliary ego|a dependable figure who lends ego strength the client temporarily lacks]]. {{c:Affective Therapies Reader}}",
        "{{THEORY}} It was described by **Wallerstein, Rockland, and Winston, Pinsker and McCullough** for patients with limited ego strength who need strengthening rather than insight. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, supportive psychotherapy grows from the earliest practice of psychiatry, its **reassurance, advice and moral treatment**, and from psychodynamic therapy at the supportive end of Luborsky's supportive-expressive continuum. {{c:Affective Therapies Reader}}",
        "Its **origin** was given shape by **Wallerstein's Menninger project, Rockland, and Winston, Pinsker and McCullough**, who formalised it for patients whose fragile personality cannot tolerate uncovering work. {{c:Affective Therapies Reader}}",
        "The first basis is a shift of aim, since the goal is to **maintain, restore or improve ego functions** and self-esteem and to shore up existing defences and coping rather than to make the unconscious conscious. {{c:Affective Therapies Reader}}",
        "A second basis is that the **relationship itself heals**, so a dependable, warm therapist serves as an auxiliary ego, lending the strength and stability the client temporarily lacks. {{c:Affective Therapies Reader}}",
        "A third basis is a **here-and-now focus**, keeping the work on present stresses, coping and daily functioning rather than on childhood reconstruction or deep transference. {{c:Affective Therapies Reader}}",
        "A fourth basis is the handling of transference, since **positive transference is used but not interpreted**, and deep interpretation is deliberately avoided so the client is not destabilised. {{c:Affective Therapies Reader}}",
        "In **formulation**, symptoms are read as arising when **stress overwhelms a person's coping resources**, so that an overwhelmed ego or a fragile personality needs its functioning and adaptation strengthened. {{c:Affective Therapies Reader}}",
        "The **goals** are **symptom relief and stabilisation**, improved adaptation and coping, support through crisis or illness, prevention of relapse and rehospitalisation, and enhanced self-esteem. {{c:Affective Therapies Reader}}",
        "A defining **procedure** is the therapist's stance, which is **active, warm, flexible and often directive**, meeting the client where they are rather than waiting in silence for associations. {{c:Affective Therapies Reader}}",
        "Core **techniques** include **reassurance, encouragement and praise**, empathic listening and [[ventilation|letting the client talk out and release pent-up feelings]], and advice and guidance where the client cannot yet decide alone. {{c:Affective Therapies Reader}}",
        "Further **techniques** are **psychoeducation, explanation and clarification**, anticipatory guidance that prepares the client for stressful events, and practical problem solving and reality testing. {{c:Affective Therapies Reader}}",
        "Additional **techniques** include **reframing**, environmental modification such as involving family or arranging services, modelling, limit-setting, the strengthening of adaptive defences and timely crisis intervention. {{c:Affective Therapies Reader}}",
        "The **frequency and duration** are flexible, ranging from weekly to monthly sessions and from brief and time-limited work to long-term maintenance for chronic conditions. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work moves from assessing **strengths and ego capacities**, through building the alliance and defining problems and goals, to selecting and applying supportive techniques. {{c:Affective Therapies Reader}}",
        "The later **process** is monitoring progress with a gradual reduction or ongoing maintenance and an open-ended, flexible ending rather than a sharply scheduled termination. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **better coping and daily functioning**, reduced distress and fewer relapses, with stabilisation valued above the deeper personality change pursued by expressive therapy. {{c:Affective Therapies Reader}}",
        "**Indications** are wide, covering schizophrenia and other severe mental illness, bipolar and personality disorders with low ego strength, crises and bereavement, chronic or terminal medical illness, the elderly and those with low insight. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Wallerstein's Menninger study** found supportive therapy produced more lasting change than expected, and later work shows it comparable to other therapies, yet it is under-researched and often used only as a control condition, and critics warn it may foster dependency. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Supportive psychotherapy at a glance", "items": [
        "Origin: psychiatry's reassurance and advice; the supportive end of Luborsky's continuum; Wallerstein.",
        "Basis: strengthen ego functions and self-esteem; the relationship heals; here-and-now focus.",
        "Formulation: stress overwhelms coping; an overwhelmed ego or fragile personality.",
        "Technique: reassurance, praise, ventilation, advice, psychoeducation, problem solving, reframing.",
        "Indications: severe mental illness, crises, medical illness; comparable but under-researched.",
    ]},
    {"t": "mn", "word": "SUPPORT", "text": "Strengthen the ego, Use the relationship, Praise and reassure, Problem-solve, Offer advice and psychoeducation, Reality-test and Reframe, Tackle the crisis \u2014 the aims and techniques of supportive psychotherapy, which bolsters coping rather than uncovering the unconscious."},
])

# ========================= 11. CLINICAL HYPNOTHERAPY ======================
_s("t11", "11", "Clinical hypnotherapy", [
    {"t": "lead", "text": "Clinical hypnotherapy uses hypnosis, a state of focused attention and heightened responsiveness to suggestion, to help clients modify perception, emotion and behaviour, moving through pre-induction, induction, deepening, a therapeutic phase and re-alerting to treat pain, anxiety and habit problems, usually as an adjunct."},
    {"t": "fig", "key": "hypnotherapy_stages", "caption": "The five stages of a hypnotherapy session: pre-induction, induction, deepening, therapeutic phase and termination."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **American Psychological Association (Division 30)** characterises [[hypnosis|a state of focused attention with reduced peripheral awareness and heightened response to suggestion]] as a state of consciousness involving focused attention and reduced peripheral awareness, marked by an enhanced capacity for response to suggestion. {{c:APA Division 30}}",
        "{{THEORY}} **Milton H. Erickson** shaped modern practice with a **permissive, indirect and utilisation** approach, using the client's own imagery and language rather than authoritarian commands. {{c:Erickson}}",
        "{{THEORY}} Theories divide into **state views**, such as Hilgard's neodissociation, and **non-state or socio-cognitive views** that stress role-taking, expectancy and imagination. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, hypnotherapy descends from **Mesmer's animal magnetism** in the 1770s, through Puysegur's artificial somnambulism and James Braid, who in 1843 coined the word hypnosis from the Greek for sleep. {{c:Affective Therapies Reader}}",
        "Its **origin** runs on through the **Nancy school of Liebeault and Bernheim**, who saw hypnosis as suggestion, Charcot and Janet on hysteria and dissociation, and Breuer and Freud's early catharsis under hypnosis, which Freud later abandoned. {{c:Affective Therapies Reader}}",
        "The modern **origin** is dominated by **Milton Erickson's** permissive and indirect method and by Hilgard's neodissociation theory with the Stanford scales, alongside socio-cognitive theorists such as Barber, Sarbin, Spanos and Kirsch. {{c:Erickson}}",
        "The first basis is the nature of the hypnotic state, marked by **absorption, focused attention and suggestibility** with reduced critical analysis, and it is important that hypnosis is **not sleep** and not a loss of control. {{c:APA Division 30}}",
        "A second basis is the debate between **state and non-state theories**, where one camp sees an altered state of dissociation and the other sees role-taking, expectancy and imagination, with neuroimaging showing changes in attention and default-mode networks. {{c:Affective Therapies Reader}}",
        "A third basis is [[hypnotisability|a stable trait describing how readily a person responds to hypnotic suggestion]], a fairly stable trait that varies across people and is measured by scales such as the Stanford and Harvard group scales. {{c:Affective Therapies Reader}}",
        "In **formulation**, symptoms, habits, pain and anxiety are seen as maintained by **automatic learned responses** and by negative beliefs and imagery, which focused attention and suggestion can be used to modify. {{c:Affective Therapies Reader}}",
        "The **goals** are **symptom control**, especially of pain and anxiety, habit change, relaxation, increased self-efficacy and ego-strength, and the teaching of self-hypnosis so the client can manage symptoms alone. {{c:Affective Therapies Reader}}",
        "A core **procedure** is the **pre-induction phase**, which secures consent, takes a history, explains hypnosis and dispels the myths of lost control or revealed secrets, and builds expectancy and rapport. {{c:Affective Therapies Reader}}",
        "The **induction technique** then narrows attention through eye fixation, progressive relaxation, breathing or imagery, with Ericksonian inductions using **confusion, storytelling and pacing and leading** rather than direct command. {{c:Erickson}}",
        "The **deepening technique** follows, using counting down, staircase imagery or [[fractionation|repeatedly lifting and re-inducing trance to deepen it]], which intensify the hypnotic experience before therapeutic work begins. {{c:Affective Therapies Reader}}",
        "In the **therapeutic phase** the clinician uses **direct and indirect suggestions, metaphors and ego-strengthening**, together with imagery and rehearsal and, where indicated, hypnotic analgesia such as glove anaesthesia. {{c:Affective Therapies Reader}}",
        "Deeper **techniques** in that phase include age regression and [[hypnoanalysis|uncovering and working through buried material under hypnosis]], post-hypnotic suggestions and anchoring, desensitisation and parts therapy, all used with caution. {{c:Affective Therapies Reader}}",
        "The final **technique** is termination, a careful **re-alerting or dehypnosis** by counting up and reorienting the client, followed by discussion and the teaching of self-hypnosis for ongoing practice. {{c:Affective Therapies Reader}}",
        "In **stages and process** the session therefore moves through the fixed sequence of **pre-induction, induction, deepening, therapeutic phase and termination**, each step prepared before the next is attempted. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is reduced pain and anxiety, changed habits and greater self-efficacy, and gains are strengthened when the client practises **self-hypnosis** between sessions. {{c:Affective Therapies Reader}}",
        "**Indications** include chronic and procedural pain, labour and dental pain, anxiety and phobias, irritable bowel syndrome, insomnia, habit disorders and psychosomatic problems, while **contraindications and cautions** cover psychosis, severe personality disorder and the risk of false memory in forensic use. {{c:Affective Therapies Reader}}",
        "In **research and current status**, meta-analyses find that **adding hypnosis to cognitive-behavioural therapy improves outcomes**, with strong evidence for pain, irritable bowel syndrome and procedural distress, moderate evidence for anxiety and weaker or inconclusive evidence for smoking and weight loss. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Clinical hypnotherapy at a glance", "items": [
        "Origin: Mesmer and Braid; the Nancy school; Erickson, Hilgard and the socio-cognitive theorists.",
        "Basis: focused attention and suggestibility (APA Division 30); state vs non-state; hypnotisability.",
        "Formulation: symptoms kept up by automatic responses and imagery, modified by suggestion.",
        "Technique: pre-induction, induction, deepening, therapeutic suggestions, re-alerting; self-hypnosis.",
        "Indications: pain, IBS, anxiety, habits; strong for pain and IBS, weak for smoking and weight.",
    ]},
    {"t": "mn", "word": "PIDTT", "text": "Pre-induction, Induction, Deepening, Therapeutic phase and Termination \u2014 the five stages of a hypnotherapy session; throughout, hypnosis is focused attention with a heightened response to suggestion, used mostly as an adjunct."},
])

# ========================= 12. PLAY THERAPY ======================
_s("t12", "12", "Play therapy", [
    {"t": "lead", "text": "Play therapy treats play as the child's natural language, so that toys become words and play becomes speech, and it uses a safe playroom and a trained therapist, in directive or non-directive form, to help children express and master emotions, trauma and relationship difficulties."},
    {"t": "fig", "key": "play_classification", "caption": "A classification of play therapy from non-directive child-centred work to directive and parent-mediated forms."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} Play therapy rests on the idea that **play is the child's natural language**, so that, in Landreth's phrase, toys are the child's words and play is the child's speech. {{c:Axline}}",
        "{{THEORY}} **Virginia Axline** set out the **non-directive, child-centred** form in 1947 and its eight guiding principles, applying Rogerian conditions to children. {{c:Axline}}",
        "{{THEORY}} Its roots reach back to **Anna Freud and Melanie Klein**, who first used play analytically, and forward to Landreth's child-centred play therapy and Knell's cognitive-behavioural play therapy. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, play therapy begins with **Freud's case of Little Hans** in 1909 and Hermine Hug-Hellmuth, the first to use play with children, before **Anna Freud and Melanie Klein** made play central to child analysis. {{c:Affective Therapies Reader}}",
        "Its **origin** then runs through **Virginia Axline's** non-directive play therapy in 1947 and her book Dibs: In Search of Self, and on to the Guerneys' filial therapy, Landreth's child-centred play therapy and Schaefer's founding of the Association for Play Therapy. {{c:Axline}}",
        "The first basis is that **play is the child's natural language**, since children cannot easily put abstract thoughts and feelings into words and so express them through symbolic play with toys. {{c:Axline}}",
        "A second basis is the [[therapeutic powers of play|the specific ways play heals, such as communication, catharsis and relationship building]], named by Schaefer as communication, emotional wellness, stronger relationships and increased personal strengths. {{c:Affective Therapies Reader}}",
        "A third basis is that play therapy draws on **several theories**, the psychoanalytic, the humanistic and child-centred, the attachment, the cognitive-behavioural and the developmental, so it can be directive or non-directive. {{c:Affective Therapies Reader}}",
        "A fourth basis is **Axline's eight principles**, which include building a warm relationship, accepting the child exactly as they are, reflecting feelings and respecting the child's own capacity to solve problems. {{c:Axline}}",
        "In **formulation**, children's problems are read as **blocked emotions, trauma, insecure attachment or unmet needs**, which play lets them express and master safely at a symbolic distance from the real event. {{c:Affective Therapies Reader}}",
        "The **goals** are **emotional expression and regulation**, a better self-concept and self-esteem, social skills and coping, the resolution of trauma or conflict and a strengthened parent-child attachment. {{c:Affective Therapies Reader}}",
        "A core **procedure** is the **playroom and its toys**, sorted into real-life toys such as a doll house, acting-out and aggressive toys such as puppets and a bop bag, and creative media such as sand, clay and paint. {{c:Affective Therapies Reader}}",
        "Non-directive **techniques** include **tracking** the child's play aloud, reflecting feelings, returning responsibility with phrases such as 'you can decide' and encouragement that follows rather than steers the child. {{c:Affective Therapies Reader}}",
        "Limit-setting is a key **technique**, often taught as [[ACT|acknowledge the feeling, communicate the limit, target an alternative]], which keeps the playroom safe without shaming the child. {{c:Affective Therapies Reader}}",
        "Structured **techniques** include **puppets, sandplay, drawing and storytelling**, Gardner's mutual storytelling, bibliotherapy, role play, therapeutic games and cognitive-behavioural play that uses modelling and exposure. {{c:Affective Therapies Reader}}",
        "Parent-mediated **techniques** are distinctive, since **filial therapy and child-parent relationship therapy** train parents to hold play sessions at home, and Theraplay builds structure, engagement, nurture and challenge. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work moves from assessment, parent consultation and consent, through **relationship building** as the child explores the room and tests the therapist, into the appearance of themes in the play. {{c:Affective Therapies Reader}}",
        "The later **process** is a **working-through phase** of aggression, regression, limit-testing and dramatic play toward mastery, and then a termination for which the child is prepared early, usually across about twelve to twenty weekly sessions. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **better emotional expression, behaviour and self-esteem**, improved social skills and a stronger parent-child relationship, especially when parents are actively involved. {{c:Affective Therapies Reader}}",
        "**Indications** are children of about three to twelve years with anxiety, fears, trauma and abuse, bereavement, parental divorce, adjustment and behaviour problems and selective mutism, while **adolescents** may need talk, art or drama instead. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Bratton and colleagues' 2005 meta-analysis of ninety-three studies** reported a large mean effect, larger when parents took part, and later child-centred play therapy reviews show moderate to large effects, though critics note methodological weaknesses. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Play therapy at a glance", "items": [
        "Origin: Freud's Little Hans; Anna Freud and Klein; Axline (non-directive, 1947); Landreth's CCPT.",
        "Basis: play is the child's language; the therapeutic powers of play; Axline's eight principles.",
        "Formulation: blocked emotions, trauma, insecure attachment expressed through symbolic play.",
        "Technique: playroom and toys, tracking, reflecting, ACT limit-setting, sandplay, filial therapy.",
        "Research: Bratton et al. (2005) large effect; moderate to large for CCPT; children about 3 to 12.",
    ]},
    {"t": "mn", "word": "TRACK", "text": "Toys are words, Relationship first, Accept the child, Children-centred (non-directive) and Keep safe limits (ACT) \u2014 the heart of play therapy, in which play is the child's natural language for expressing and mastering experience."},
])

# ========================= 13. PSYCHODRAMA ======================
_s("t13", "13", "Psychodrama", [
    {"t": "lead", "text": "Psychodrama is J. L. Moreno's action-based group therapy in which a protagonist enacts rather than merely describes their conflicts, using a director, auxiliary egos, an audience and a stage, so that spontaneity and new roles are restored through warm-up, action and sharing."},
    {"t": "fig", "key": "psychodrama_phases", "caption": "The five instruments and three phases of psychodrama: warm-up, action and sharing on the stage."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Jacob L. Moreno** created psychodrama from his Theatre of Spontaneity, coined the term **group psychotherapy** and developed sociometry, the measurement of relationships within a group. {{c:Moreno}}",
        "{{THEORY}} Its foundation is that [[spontaneity|a fresh, adequate response to a new or old situation]] and creativity are the basis of mental health, so that pathology is a loss of spontaneity and rigid roles. {{c:Moreno}}",
        "{{THEORY}} Moreno described **five instruments**, the protagonist, the director, the auxiliary egos, the audience and the stage, that together make enactment possible. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, psychodrama was created by **Jacob L. Moreno** in Vienna, whose Theatre of Spontaneity in 1921 led, after he moved to the United States, to the Beacon sanatorium and the systematic method of action-based group therapy. {{c:Moreno}}",
        "Its **origin** also credits **Zerka Moreno** as a key developer, and the method went on to influence Gestalt therapy, drama therapy, role play and the behavioural rehearsal used in cognitive-behavioural work. {{c:Affective Therapies Reader}}",
        "The first basis is that **spontaneity and creativity are the foundations of mental health**, so that distress is understood as a loss of spontaneity, rigid role patterns and unfinished action. {{c:Moreno}}",
        "A second basis is **role theory**, covering role-taking, role-playing and role-creating, so that role conflict and role rigidity are seen as central to psychological problems. {{c:Affective Therapies Reader}}",
        "A third basis is [[tele|a two-way flow of genuine feeling and understanding between people]], Moreno's term for the mutual empathy that binds a group, distinct from one-way transference. {{c:Affective Therapies Reader}}",
        "A fourth basis is **catharsis, encounter and surplus reality**, the last letting a client enact what never happened or what is feared or wished, all worked in the here-and-now moment. {{c:Affective Therapies Reader}}",
        "In **formulation**, distress is traced to a **limited or conflicting role repertoire**, blocked spontaneity and unresolved relationships and events, which enactment brings into the open to be experienced and changed. {{c:Affective Therapies Reader}}",
        "The **goals** are **catharsis and release, insight and a wider role repertoire**, renewed spontaneity, better interpersonal skills, the resolution of unfinished business and group cohesion and support. {{c:Affective Therapies Reader}}",
        "A defining **procedure** is the use of the **five instruments**, the protagonist whose issue is explored, the director who guides, the auxiliary egos who play significant others, the audience who witness and the stage. {{c:Affective Therapies Reader}}",
        "The most powerful **technique** is [[role reversal|the protagonist plays the other person to see through their eyes]], which builds empathy and insight by letting the protagonist experience the other's position. {{c:Affective Therapies Reader}}",
        "A second key **technique** is **doubling**, in which an auxiliary voices the protagonist's unspoken feelings, and **mirroring**, in which the protagonist watches an auxiliary portray them from the outside. {{c:Affective Therapies Reader}}",
        "Further **techniques** include **soliloquy, the empty chair and future projection**, together with concretisation, sculpting, the social atom and the magic shop, which make abstract feelings concrete. {{c:Affective Therapies Reader}}",
        "A related **technique** is **sociodrama**, which turns the same action methods onto shared group or social issues rather than one person's private conflict. {{c:Affective Therapies Reader}}",
        "In **stages and process**, the work opens with a **warm-up** in which the director builds trust and a protagonist emerges, so the group is ready for action. {{c:Affective Therapies Reader}}",
        "The **action phase** then sets and enacts the scene, often moving from the present into the past, using role reversal, doubling and mirroring to reach a peak of emotion and new solutions before the scene is closed. {{c:Affective Therapies Reader}}",
        "The closing **process** is **sharing and integration**, in which the audience offer their own similar experiences rather than analysis or advice, reducing the protagonist's isolation and supporting re-entry. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **emotional release and insight**, improved relationships and role flexibility and the support of the group, achieved through doing rather than only talking. {{c:Affective Therapies Reader}}",
        "**Indications** are group settings for trauma, grief, addictions, family and relationship issues, adolescents, social-skills training and psychiatric inpatients, while acute psychosis or a very fragile client without containment is a **caution** needing a skilled director. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Kipper and Ritchie's 2003 meta-analysis** reported a large effect with role reversal and doubling the most effective techniques, and later reviews show moderate positive evidence, though rigorous randomised trials remain few. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Psychodrama at a glance", "items": [
        "Origin: J. L. Moreno (Theatre of Spontaneity, 1921); coined group psychotherapy and sociometry.",
        "Basis: spontaneity and creativity; role theory; tele; catharsis, encounter and surplus reality.",
        "Formulation: a limited role repertoire and blocked spontaneity; unresolved relationships.",
        "Instruments and technique: protagonist, director, auxiliary egos, audience, stage; role reversal.",
        "Process: warm-up, action, sharing; Kipper and Ritchie (2003) large effect; few rigorous RCTs.",
    ]},
    {"t": "mn", "word": "PDAAS", "text": "Protagonist, Director, Auxiliary egos, Audience and Stage \u2014 Moreno's five instruments of psychodrama; the work flows through warm-up, action and sharing, with role reversal the most powerful technique."},
])

# ========================= 14. ORIENTAL APPROACHES ======================
_s("t14", "14", "Oriental approaches", [
    {"t": "lead", "text": "Oriental approaches are Asian mind-body practices used for mental health, among them yoga, meditation, shavasana, pranic healing, reiki and tai chi, which share a focus on attention, breath and relaxation, but whose evidence ranges from genuine support for mindfulness, yoga and tai chi down to weak and inconclusive findings for energy healing."},
    {"t": "fig", "key": "oriental_map", "caption": "A map of the named oriental approaches grouped by evidence, from mind-body practices with support to energy healing with weak evidence."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and traditions", "items": [
        "{{THEORY}} **Patanjali's Yoga Sutras** characterise [[yoga|the union of body, mind and spirit through an eight-limbed discipline]] as the cessation of the fluctuations of the mind, achieved through the eight limbs from ethical restraints to absorption. {{c:Affective Therapies Reader}}",
        "{{GUIDE}} Clinical **mindfulness meditation** is characterised as non-judgemental, present-moment awareness, adapted by Kabat-Zinn as Mindfulness-Based Stress Reduction and recommended within the World Health Organization's wider promotion of self-care for mental health. {{c:WHO mhGAP}}",
        "{{THEORY}} **Energy-healing** practices such as pranic healing and reiki are characterised by their own traditions as the clearing or channelling of a life energy, prana or ki, a claim for which there is **no credible scientific evidence**. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, techniques, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, these approaches come from **Asian philosophical and healing traditions**, chiefly the Indian, Buddhist, Taoist and Japanese, and many have since been adapted as structured, sometimes evidence-based, interventions. {{c:Affective Therapies Reader}}",
        "**Yoga** originates in ancient India and was systematised by **Patanjali** in the Yoga Sutras as Ashtanga or eight-limbed yoga, with later Hatha yoga and types such as Raja, Karma, Bhakti and Jnana yoga. {{c:Affective Therapies Reader}}",
        "The basis of **yoga** is [[union|the joining of body, mind and spirit, the literal meaning of the word yoga]], reached through the eight limbs of ethical restraints, observances, postures, breath control, sense withdrawal, concentration, meditation and absorption. {{c:Affective Therapies Reader}}",
        "In therapy, **yoga techniques** are gentle **asanas, pranayama breathing and relaxation**, and well-known studies at NIMHANS and meta-analyses report real benefit as an adjunct for mild to moderate depression and anxiety. {{c:Affective Therapies Reader}}",
        "**Meditation** originates in Hindu, Buddhist and Taoist contemplative traditions, and was adapted clinically as **Mindfulness-Based Stress Reduction** by Jon Kabat-Zinn in 1979 and Mindfulness-Based Cognitive Therapy by Segal, Williams and Teasdale. {{c:Affective Therapies Reader}}",
        "The basis of **meditation** is **training attention and decentring**, learning to observe thoughts as passing mental events rather than facts, which lowers rumination and increases acceptance and emotion regulation. {{c:Affective Therapies Reader}}",
        "The evidence for **meditation** is genuine, since mindfulness-based cognitive therapy is **recommended for preventing relapse** in recurrent depression and a major review by Goyal and colleagues in 2014 found moderate evidence for anxiety, depression and pain. {{c:Affective Therapies Reader}}",
        "**Shavasana**, the corpse pose, is the final relaxation posture of a yoga session, and the related [[Yoga Nidra|a guided deep relaxation also called yogic sleep]] was popularised by Swami Satyananda Saraswati. {{c:Affective Therapies Reader}}",
        "The basis of **shavasana** is a **relaxation response** like Benson's, which lowers sympathetic arousal, heart rate, blood pressure and muscle tension, resembling Jacobson's progressive muscle relaxation, with small trials showing reduced anxiety and better sleep. {{c:Affective Therapies Reader}}",
        "**Pranic healing** was developed by **Master Choa Kok Sui** in 1987 and holds that disease results from depleted or congested prana in an energy body, treated by scanning, cleansing and energising the field, usually without touch. {{c:Affective Therapies Reader}}",
        "The evidence for **pranic healing** is **very limited and poorly controlled**, with no credible evidence for a detectable energy field, so any benefit is probably relaxation, attention and expectancy, and it must not replace evidence-based treatment. {{c:Affective Therapies Reader}}",
        "**Reiki** originates with **Mikao Usui** in Japan in 1922 and has the practitioner channel a universal life energy through the hands to the client, who lies clothed while the practitioner uses set hand positions. {{c:Affective Therapies Reader}}",
        "The evidence for **reiki** is **weak and at high risk of bias**, with reviews such as the Cochrane review for anxiety and depression finding insufficient evidence, so it is best regarded as a complementary comfort measure that should not delay treatment. {{c:Affective Therapies Reader}}",
        "**Tai chi** originates in seventeenth-century China from Taoist philosophy and martial arts and is often called **meditation in motion**, using slow, flowing, coordinated movements with diaphragmatic breathing and focused attention. {{c:Affective Therapies Reader}}",
        "The evidence for **tai chi** is **reasonably good**, with meta-analyses reporting reduced stress, anxiety and depression, benefit as an adjunct in late-life depression and value for fall prevention, chronic pain and cardiac rehabilitation. {{c:Affective Therapies Reader}}",
        "Other related approaches, offered under 'etc.', include **Morita and Naikan therapy from Japan and qigong**, where Morita teaches accepting one's feelings and doing what needs to be done and Naikan uses structured introspection to build gratitude. {{c:Affective Therapies Reader}}",
        "The common **basis and mechanism** across these practices is improved **self-regulation**, raised parasympathetic and vagal tone, lowered cortisol and arousal and better attention and emotion regulation, which explains much of their benefit. {{c:Affective Therapies Reader}}",
        "The honest **outcome and indications** are that mindfulness, yoga and tai chi have a **real evidence base** as adjuncts for stress, anxiety, depression and chronic pain, while **reiki and pranic healing have weak or inconclusive evidence** and carry the risk of replacing proper care. {{c:Affective Therapies Reader}}",
        "In **research and current status**, these approaches fit Indian and Asian clients' cultural beliefs and are low in cost and side effects, but their evidence is of **variable quality**, strongest for mindfulness-based therapies and weakest for energy-healing claims that remain untestable. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "key", "title": "Oriental approaches at a glance", "items": [
        "Named practices: yoga, meditation, shavasana (and Yoga Nidra), pranic healing, reiki, tai chi (etc.).",
        "Basis: attention, breath and relaxation; self-regulation, raised vagal tone and lowered arousal.",
        "Good evidence: mindfulness (MBCT for relapse), yoga and tai chi as adjuncts for anxiety and depression.",
        "Weak or inconclusive: reiki and pranic healing energy healing; benefit likely relaxation and expectancy.",
        "Cautions: must not replace medical or psychological care; watch for unqualified practitioners.",
    ]},
    {"t": "mn", "word": "YMSPRT", "text": "Yoga, Meditation, Shavasana, Pranic healing, Reiki and Tai chi \u2014 the named oriental approaches; mindfulness, yoga and tai chi have a real evidence base, while reiki and pranic healing have weak or inconclusive evidence."},
])
