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
    {"t": "lead", "text": "These are easy, exam-ready revision notes for Unit IV of the Psychotherapy paper (M.Phil. Clinical Psychology). They cover the affective psychotherapies one by one. For each therapy you get its origin, basis, formulation, goals, techniques, process, outcome, indications and research in short, plain sentences. So you can revise a whole 5-mark answer from one topic."},
    {"t": "h3", "text": "What the sections are"},
    {"t": "ul", "items": [
        "Each numbered section is one therapy from the syllabus. It is written as a 5-mark exam answer in full but simple sentences, so you can revise straight from the point and write it back in the exam.",
        "Every topic opens with a one-line summary. It then has a flowchart that pictures the core idea, a short definition box from well-known bodies and classic authors, the answer points, an 'at a glance' box, and a mnemonic that packs the key points into one memorable word.",
        "A short source note and a references list sit at the end. The references name only the books (full title and author) and the online or primary sources that were actually used.",
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
        "A citation in square brackets names the full book and its author, for example {{c:Corey}} or {{c:Yalom}}. So you can trace any claim back to the work it comes from.",
        "Foundational ideas are cited to the author and classic work that introduced them, for example {{c:Freud}} or {{c:Rogers}}. Where a point rests on a professional body, the citation names that body, for example {{c:APA Dictionary of Psychology}}.",
    ]},
    {"t": "h3", "text": "How to read the formatting"},
    {"t": "ul", "items": [
        "Words in **bold** are the memory hooks. Skim the bold words first to recall the gist, then read the full sentence for the detail.",
        "Hard clinical terms are kept, but each one is explained right after it in a coloured bracket, like this: [[transference|reacting to the therapist as if they were someone from your past, like a parent]]. Just read the coloured bracket to get the plain meaning, so the hard word never leaves you stuck.",
        "Tricky ideas also get a short plain sentence that unpacks them, often with a tiny example, right next to the technical wording. Every topic also carries one flowchart and one mnemonic.",
    ]},
]

# ---------------------------------------------------------------------------
# A proper, simple introduction to the affective psychotherapies. Rendered as a
# dedicated "plain" section right after "How to read these notes" and before
# Topic 1, in both deliverables. Kept genuinely introductory and plain.
AFFECTIVE_INTRO = [
    {"t": "lead", "text": "Before the fourteen therapies, here is a simple map of the whole area. Read this once and the rest of the unit will make much more sense, because you will see what all these therapies are trying to do and how they differ."},
    {"t": "h3", "text": "What 'affective psychotherapies' means here"},
    {"t": "ul", "items": [
        "The word [[affective|to do with feelings and emotions]] simply points to the heart of these therapies. They work mainly through **emotion, insight, the relationship and lived experience**, by talking and relating, not by giving medicines and not mainly by drilling new habits. In plain words, they help you by changing how you feel, how you understand yourself, and how you relate to others.",
        "So an **affective (talk-and-relationship) therapy** is a planned, caring conversation. A trained therapist and a client meet regularly, and the healing comes from understanding feelings, from facing them, and from a trusting bond, rather than from a pill or a reward chart.",
        "These therapies sit apart from two neighbours. They differ from the **purely behavioural** approach, which changes behaviour through conditioning and rewards, and from the **medical** approach, which treats the brain with drugs. In practice the lines blur, and many therapists mix ideas, but the affective group leads with feeling, meaning and relationship.",
    ]},
    {"t": "h3", "text": "What the fourteen therapies share"},
    {"t": "ul", "items": [
        "They all rely on the [[therapeutic relationship|the trusting, caring bond between client and therapist, which is itself part of the cure]]. Decades of research show this bond, called the working alliance, is one of the strongest predictors of whether therapy helps, whatever the method. Think of it as the soil: the technique is the seed, but without good soil little grows.",
        "They all treat **emotion as information, not noise**. A feeling is read as a message about an unmet need, a hidden conflict or a blocked goal, so the work is to feel it, understand it and let it change, instead of just switching it off.",
        "They all aim at **lasting inner change**, not only quick symptom relief. The hope is more self-understanding, more freedom to choose, and healthier ways of relating, so the gains keep working after therapy ends.",
        "They all give the client an active part. The client is a partner who explores, chooses and practises, not a passive patient who is simply 'treated'.",
    ]},
    {"t": "h3", "text": "How the fourteen broadly group"},
    {"t": "ul", "items": [
        ("The fourteen fall into a few natural families. Knowing the family tells you, at a glance, what a therapy will emphasise:", [
            "**Insight / depth** therapies look inward and backward to uncover hidden conflict: **psychodynamic** therapy (and its briefer offshoots).",
            "**Humanistic-experiential** therapies trust the person's drive to grow and work in the here-and-now: **humanistic, person-centred, existential and gestalt** therapy.",
            "The **Adlerian** approach stands a little apart, mixing early insight with practical encouragement and a strong social focus.",
            "**Brief and integrative** therapies are short, focused and practical: **brief psychotherapy**, which borrows from many schools around one agreed problem.",
            "**Supportive** therapy strengthens coping and steadies the person, rather than digging deep: **supportive psychotherapy**.",
            "**Action and expressive** therapies heal through doing, playing and enacting: **play therapy** (for children) and **psychodrama** (acting out conflicts in a group).",
            "**Altered-state and mind-body** approaches work through focused attention, relaxation and the body: **clinical hypnotherapy** and the **oriental approaches** such as yoga, meditation and tai chi.",
        ]),
    ]},
    {"t": "box", "kind": "info", "title": "The one idea to carry through the whole unit", "items": [
        "Across all fourteen, two engines do the real work: a **genuine, trusting relationship** and some form of **emotional change** (feeling, understanding or re-living an emotion so it loosens its grip). Every technique in this unit is, in the end, a particular way of switching on one of these two engines.",
        "So as you read each therapy, keep asking three questions: **what is its main focus** (its central idea), **what are its techniques** (what the therapist actually does in the room), and **how does it work** (why that changes the client). Those three threads run through every topic below.",
    ]},
    {"t": "h3", "text": "How to use this unit"},
    {"t": "ul", "items": [
        "Each of the fourteen numbered topics is a self-contained study note on one therapy. You can revise any one on its own, because each gives the main focus, the basis, how a case is understood, the goals, the techniques in detail, the process or stages, the outcome, who it suits, and its research standing.",
        "For every therapy, spend most of your effort on three things: its **main focus**, its **techniques** (what is done and why), and **how it works**. History is included but kept short, because it carries the fewest marks and matters least for understanding the therapy.",
        "Use the **flowchart** to picture the core idea and the **mnemonic** to lock the key points into memory. Read the coloured brackets for the plain meaning of every hard term, so no piece of jargon ever leaves you stuck.",
    ]},
]

# Honest limits and source conventions; rendered as a dedicated section in
# both deliverables so the disclosure travels with the content.
SOURCE_NOTE = [
    {"t": "lead", "text": "Honesty note: no authors, citations, study results, effect sizes or page numbers in these notes are invented. The following limits apply and are disclosed so you can verify independently."},
    {"t": "ul", "items": [
        "Professional-body characterisations of the therapies (for example the APA Dictionary of Psychology entries for psychotherapy, psychodynamic psychotherapy, person-centered therapy and gestalt therapy) are given as characterisations drawn from standard reference wording, not as word-for-word quotations, and no page numbers are attached.",
        "Foundational claims are attributed to the author and classic work that introduced them (Freud, Rogers, Perls, May, Yalom, Frankl, Adler). Where research is summarised, only well-known findings and named authors are given, and the strength of evidence is stated honestly rather than dressed up with invented numbers.",
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
    {"t": "lead", "text": "Psychodynamic psychotherapy is the modern, shorter form of Freud's psychoanalysis. It sees a symptom as a clue to a hidden mental struggle, not a random fault. The therapist uses the relationship, free talking and gentle explaining to bring that hidden struggle into the open, so lasting change can happen."},
    {"t": "fig", "key": "psychodynamic_model", "caption": "The Freudian model: unconscious conflict between id, ego and superego surfaces as a symptom, which therapy traces back to insight."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** describes [[psychodynamic psychotherapy|therapy that works with hidden mental conflicts and with the bond between client and therapist]] as a therapy that grew out of psychoanalysis. It works with [[unconscious processes|mental activity going on outside your awareness, like a motive you don't notice]] that still shape how you act today. The aim is self-awareness and seeing how the past shapes the present. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} **Freud** started psychoanalysis with Breuer's 'talking cure'. He moved on from hypnosis and [[catharsis|releasing bottled-up emotion so it stops pressing on you]] to **free association and interpretation**. He summed up its aim in one line: where id was, there **ego shall be** (the sensible self should take charge). {{c:Freud}}",
        "{{THEORY}} Later schools widened the approach into ego psychology, object relations, **self psychology** and attachment. So modern psychodynamic therapy is face-to-face and briefer than classical analysis, while keeping the same core ideas. {{c:Gabbard}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is the hidden, unconscious part of the mind. Psychodynamic therapy holds that much of what drives us sits below awareness, like the bulk of an iceberg under the water. Old conflicts and feelings we pushed away still pull the strings today and show up as symptoms.",
        "So the whole therapy aims at one thing: bring the hidden into the open. When a buried conflict is understood and felt in words, it loses its grip, and the person is freed to choose instead of being driven. Freud summed it up as **'where id was, there ego shall be'** \u2014 the sensible, choosing self takes charge of the blind, wanting part.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is [[insight|the 'aha' moment of seeing a hidden pattern or feeling in yourself clearly for the first time]] plus a safe relationship. When a client finally sees and feels what was buried \u2014 say, that their constant people-pleasing hides a fear of being abandoned \u2014 the old pattern can relax.",
        "Crucially, insight is not just understanding in the head. The client must **re-live the old feeling in the room**, usually toward the therapist, and have it meet a new, accepting response. That gap between the dreaded reaction and the real, kind one is what loosens the knot.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, the approach begins with **Freud and Breuer** and the case of Anna O., the so-called talking cure. From it Freud built psychoanalysis, moving from hypnosis and catharsis toward free association and dream interpretation. {{c:Freud}}",
        "Its later **origin** runs through **Jung, Adler and Klein**, plus Anna Freud, Winnicott, Kohut and Bowlby. Object relations, self psychology and attachment all grew from this one root into today's briefer, face-to-face forms. {{c:Gabbard}}",
        "The first basis is [[psychic determinism|the idea that nothing in the mind is a pure accident, so every symptom has a cause]]. Behaviour and symptoms **have meaning**, so even a slip of the tongue is read as a clue, not chance. {{c:Freud}}",
        "A central basis is the **unconscious**. The mind is split into conscious, preconscious and unconscious, and into the **id, ego and superego**. The id is the blind 'I want it now' part; the superego is the strict inner parent that judges; the ego is the referee that keeps peace between the two and the real world. {{c:Freud}}",
        "The developmental basis is the **psychosexual stages**: oral, anal, phallic, latency and genital. Too much frustration or too much indulgence can cause [[fixation|getting stuck at an early stage, like a part of you never fully growing up]], which later shapes personality and symptoms. {{c:McWilliams}}",
        "Anxiety and the **defence mechanisms** are another basis. A defence is an automatic mental trick that keeps a painful feeling out of awareness. **Repression** buries it, **denial** refuses it, **projection** sees your own anger in someone else, and **sublimation** channels it into something useful like sport. {{c:McWilliams}}",
        "In **formulation**, a symptom is read as a [[compromise formation|a symptom that half-satisfies a hidden wish and half-hides it, like a nervous habit that both shows and masks a worry]]. Picture a tug-of-war between a forbidden wish, the conscience that forbids it, and reality; the symptom is the knot where the rope stops moving. {{c:McWilliams}}",
        "The guiding **goal** is to make the **unconscious conscious**. The client gains insight into hidden conflicts and patterns, the ego grows stronger, and rigid defences mature into flexible ones. This brings lasting change, not just the quick removal of one symptom. {{c:Freud}}",
        "The therapist's stance is a key **procedure**. It combines [[neutrality|not taking sides between the client's wishes, defences and conscience]], restraint and staying fairly private. The therapist acts partly as a **blank screen**, so the client's old relationship patterns show up clearly in the room. {{c:Gabbard}}",
        "The foundational **technique** is [[free association|saying whatever comes to mind without filtering it, even if it feels silly, rude or random]]. How it is done: the client lies back or sits and simply reports each passing thought. Why it helps: the mind cannot wander freely for long before it bumps into the very conflict it is avoiding, and the pauses, jokes and sudden topic-changes mark exactly where it hides. {{c:Freud}}",
        "**Dream analysis** is a classic technique, called the royal road to the unconscious. How it is done: the client tells the dream, then free-associates to each part of it. The therapist separates the remembered **manifest content** (the surface story) from the hidden **latent content** (the real wish behind it). Example: a dream of missing a train may, on association, lead to a fear of missing out on life. {{c:Freud}}",
        "**Interpretation** is the therapist's main tool. How it is done: at the right moment, gently and tentatively, the therapist puts into words a hidden meaning the client is near to seeing \u2014 never a clever lecture, always a well-timed nudge. Why it helps: naming the buried feeling lets the thinking ego take hold of it instead of being secretly ruled by it. {{c:Gabbard}}",
        "**Analysis of resistance** is a technique in its own right. [[Resistance|anything the client does, often without meaning to, that blocks the therapy \u2014 arriving late, going blank, changing the subject]] is not a nuisance but a signpost. How it is done: the therapist points out the block gently and wonders aloud what feeling it is guarding, because the strongest resistance usually sits right on top of the most important material. {{c:Gabbard}}",
        "**Working with transference** is the deepest technique. [[Transference|reacting to the therapist as if they were an important figure from your past, such as a parent]] turns the session into a live replay of the client's old relationships. How it is done: the therapist notices the client treating them as, say, a critical father, names it kindly, and links it to the past; the client then gets a different, accepting response in place of the dreaded one. Why it helps: the old pattern is corrected in real feeling, not just in talk. {{c:Gabbard}}",
        "**Working-through** is the technique that makes change stick. How it is done: the same insight is returned to again and again, in many situations, over many sessions. Why it is needed: a single 'aha' rarely shifts a lifelong habit, just as understanding why you bite your nails does not stop you the first time \u2014 only patient repetition rewrites the pattern. {{c:Gabbard}}",
        "In **stages and process** the work moves through three phases. It starts with an **opening phase** of assessment, agreeing the frame (times, fee, rules) and building trust. Then transference and resistance appear and are **worked through** in a long middle phase. Finally a planned **termination** reviews and locks in the gains and handles the feelings about ending. {{c:Gabbard}}",
        "The **outcome** sought is fewer symptoms, more insight, more flexible defences and healthier relationships. A striking finding is the **sleeper effect**: gains often keep growing even after therapy ends, because the client carries on self-analysing. {{c:Corey}}",
        "**Indications** favour clients with good ego strength, curiosity about their own mind and real motivation, facing neurotic, personality or relationship problems. Acute psychosis, high suicide risk and raw crisis states are **contraindications** better met first by supportive work. {{c:Gabbard}}",
        "In **research and current status**, well-known reviews (for example Shedler's) find psychodynamic therapy gives **substantial, lasting benefit**. Short-term forms help depression and anxiety, and specialised forms such as transference-focused and mentalisation-based therapy help borderline personality disorder, while classical long analysis stays hard to study with trials. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "key", "title": "Psychodynamic therapy at a glance", "items": [
        "Origin: Freud and Breuer, extended by ego psychology, object relations, self psychology and attachment.",
        "Basis: psychic determinism, the unconscious, id-ego-superego, psychosexual stages, defences.",
        "Formulation: the symptom is a compromise formation that stands for a hidden conflict.",
        "Technique: free association, dream analysis, interpretation, analysis of resistance and transference.",
        "Process: opening, transference, working through, termination; best for ego-strong neurotic clients.",
    ]},
    {"t": "mn", "word": "FIDO-WT", "text": "Free association, Interpretation, Dream analysis, (analysis of) resistance and transferO, Working through and Termination \u2014 the techniques and stages of psychodynamic therapy; the aim throughout is to make the unconscious conscious."},
])


# ========================= 2. BRIEF PSYCHOTHERAPY ======================
_s("t2", "2", "Brief psychotherapy", [
    {"t": "lead", "text": "Brief psychotherapy is planned, time-limited treatment. Therapist and client agree on one problem to work on, and the therapist stays active and directive. The fixed deadline is used on purpose to fire up motivation. The aim is to ease symptoms and restore functioning, not to remake the whole personality."},
    {"t": "fig", "key": "brief_focus", "caption": "Brief therapy keeps a single focus across a short, time-limited course with a fixed end date and a mid-point review."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors and models", "items": [
        "{{THEORY}} Brief therapies share a **planned time limit**, a single focus and an active therapist. They rest on a well-known finding: most improvement happens early, with about half of clients better within roughly eight sessions. {{c:Affective Therapies Reader}}",
        "{{THEORY}} Named models include **Malan, Sifneos, Mann and Davanloo**, each with their own short, focused dynamic therapy. Alongside them sit interpersonal therapy and solution-focused brief therapy. {{c:Sharf}}",
        "{{GUIDE}} The **World Health Organization** promotes brief, structured psychological help in routine care. It frames this as real help given through **respectful, collaborative** talk within a few sessions. {{c:WHO mhGAP}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is one problem, worked on fast, against a clock. Brief therapy does not try to remake the whole personality. It picks a single agreed theme \u2014 a loss, a stuck role, a current conflict \u2014 and keeps every session pointed at it.",
        "Its boldest idea is that the **deadline itself heals**. A fixed end date is not a limitation forced by money or waiting lists; it is used on purpose, like a deadline that makes you finally finish an essay. It concentrates the mind, raises motivation and stops therapy drifting.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The engine is **focus plus pressure plus hope**. By refusing to wander, the therapist lets the client make real progress on one thing; the looming end date keeps urgency high; and in the solution-focused versions, the belief that **small changes snowball** gives the client confidence to start.",
        "There is also a practical truth behind it, called the **dose-response curve**: most of the improvement people ever get from therapy happens in the first handful of sessions. Brief therapy simply builds its whole design around that fact instead of fighting it.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, brief therapy grows from **Ferenczi and Rank's** active, time-limited technique in the 1920s. It also grows from **Alexander and French's** idea of a corrective emotional experience in 1946. {{c:Affective Therapies Reader}}",
        "Its **origin** then branches into many models. The psychodynamic stream has **Malan, Sifneos, Mann and Davanloo**; the systemic stream has **de Shazer and Berg's** solution-focused therapy. {{c:Sharf}}",
        "A practical **origin** driver was service pressure. Cost, long waiting lists and the proof that **most gains come early** all pushed clinicians toward shorter, planned treatment. {{c:Affective Therapies Reader}}",
        "The first basis is that the **time limit is therapeutic**. A known end date lights a fire under the work. It concentrates effort and guards against the drift and over-dependence that an open-ended course can breed. {{c:Affective Therapies Reader}}",
        "A second basis is a tight [[focus|one agreed problem or theme that the whole therapy is built around]]. The therapist is **active and directive**, steering sessions back to the one theme instead of chasing every new worry the client raises. {{c:Sharf}}",
        "A third basis is hope about change. The work stresses the **here-and-now and the client's strengths**. It assumes a small change can set off bigger ones \u2014 like one tidy shelf prompting you to tidy a whole room \u2014 especially in the solution-focused models. {{c:Sharf}}",
        "In **formulation** the therapist builds a **focal formulation**. This names one recurring relationship pattern, a current loss or an unresolved conflict tied to the present problem, and keeps the work on target. {{c:Affective Therapies Reader}}",
        "A well-known formulation grid is **interpersonal therapy's four problem areas**: grief, role disputes (an ongoing clash in a key relationship), role transitions (a hard life change) and interpersonal deficits (few or poor relationships). One of these is picked as the single focus. {{c:Sharf}}",
        "The **goals** are modest and clear: **resolve the focal problem**, relieve symptoms and restore how the person was functioning before. The therapist also teaches skills for later, rather than rebuilding the personality. {{c:Affective Therapies Reader}}",
        "A core **procedure** is careful **patient selection**, because brief work does not suit everyone. Sifneos' criteria look for one clear complaint, at least one meaningful past relationship, curiosity about the mind, motivation and the ability to bear some anxiety. {{c:Affective Therapies Reader}}",
        "The **contract** is itself a technique. How it is done: at the very first meeting, therapist and client fix the number of sessions (usually **6 to 20**, often twelve to sixteen) and the exact end date. Why it helps: naming the finish line turns time into a shared tool both people work against. {{c:Sharf}}",
        "**Maintaining the focus** is a constant technique. How it is done: whenever the client drifts to a new worry, the therapist gently and repeatedly brings the talk back to the agreed theme. Why it helps: depth on one problem in a short time beats shallow coverage of many. {{c:Sharf}}",
        "**Focused interpretation and selective attention** are the dynamic techniques. How it is done: the therapist comments only on material that touches the focus, lets the rest pass, and gently challenges the client's defences around that one theme while using their warm feelings toward the therapist to keep them engaged. {{c:Sharf}}",
        "The **miracle question** is the signature solution-focused technique. How it is done: the therapist asks, [[miracle question|'Suppose you woke tomorrow and the problem had vanished overnight \u2014 what would be the first small sign that told you?']]. Why it helps: it quietly moves the client from staring at the problem to picturing the solution in concrete, doable steps. {{c:Sharf}}",
        "**Scaling and exception questions** complete the solution-focused kit. Scaling asks 'on a scale of 0 to 10, where are you today, and what would 1 point up look like?', making progress measurable. Exception-hunting asks 'when was the problem a little less \u2014 what was different then?', so the client's own past successes become the blueprint. {{c:Sharf}}",
        "In **stages and process** the work runs in three parts. First is **assessment and selection** in the early sessions. Then an active **middle phase** on the focus, with a mid-point review. Finally an explicit **termination** that locks in the gains. {{c:Affective Therapies Reader}}",
        "A defining process feature is that **termination is discussed from the start**. The ending is planned, the gains are reviewed, and a booster or follow-up session may be arranged, so therapy does not just trail off. {{c:Affective Therapies Reader}}",
        "The **outcome** is symptom relief and better functioning. Many comparisons find brief therapy **as effective as longer therapy** for common problems like depression, anxiety and adjustment difficulties. {{c:Corey}}",
        "**Indications** are sharp, focal problems: adjustment disorders, mild to moderate depression and anxiety, grief and crises, in motivated clients. Chronic, many-sided conditions and severe high-risk states are **contraindications**. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is strong for **interpersonal therapy and brief CBT**, and supportive for short-term psychodynamic and solution-focused work. Brief models are now standard in primary care, student counselling and telehealth. {{c:Sharf}}",
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
    {"t": "lead", "text": "Humanistic psychotherapy is the 'third force' in psychology. It is a growth-focused approach that trusts each person's natural drive to grow. It centres on how things feel from the inside, on free choice, and on the whole person. The genuine relationship, not clever technique, is what mostly heals."},
    {"t": "fig", "key": "humanistic_hierarchy", "caption": "Maslow's hierarchy of needs: distress follows when blocked lower needs stop a person from reaching self-actualisation."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** describes [[humanistic therapy|an approach that stresses growth, free choice and treating the whole person, not just symptoms]] as a therapy from humanistic psychology. It emphasises the person's **capacity for growth and self-direction** and the quality of the relationship. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} The movement began as the **third force** in the 1950s and 1960s. It pushed back against the determinism of psychoanalysis and the machine-like view of behaviourism, and drew figures like **Maslow, Rogers, May, Frankl and Perls**. {{c:Corey}}",
        "{{THEORY}} **Maslow's** hierarchy of needs runs from food and safety up to [[self-actualisation|becoming the fullest, best version of yourself]]. It frames distress as what happens when growth needs get blocked. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is the person's own drive to grow. Humanistic therapy trusts that, given the right conditions, every person naturally moves toward becoming their fullest, healthiest self \u2014 much as a plant grows toward the light when it has soil, water and sun.",
        "So the therapist does not diagnose and fix a broken machine. The focus is on the **whole person** \u2014 their inner point of view, their freedom and the here-and-now \u2014 and on providing a warm, honest relationship in which the stalled growth can start up again.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change works through a **safe, genuine relationship** that releases the blocked growth drive. When a person feels truly accepted and understood, their defences drop, they dare to look at feelings they had hidden, and the natural push toward growth resumes on its own.",
        "There is no clever technique doing the work; the relationship **is** the technique. It is like thawing \u2014 warmth does not force the ice to melt, it simply removes what was keeping it frozen, and the natural process takes over.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, humanistic therapy emerged in the **1950s and 1960s** as the self-styled third force. It was a deliberate reaction against both psychoanalytic determinism and behaviourist mechanism. The Association for Humanistic Psychology was founded in 1961. {{c:Corey}}",
        "Its **origin** gathers several figures under one banner. **Maslow, Rogers, May, Frankl and Perls** all stressed growth and choice, and its roots lie in phenomenology and existential philosophy. {{c:Corey}}",
        "The first basis is that people are **inherently good and growth-oriented**. Each person carries an [[actualising tendency|an inborn drive to grow and become your best self, like a seed that naturally pushes toward the light]] that moves them forward when conditions allow. {{c:Corey}}",
        "A second basis is the stress on **subjective experience, free will and responsibility**. The person's own point of view and the here-and-now are the proper focus, rather than drives or conditioning. Put simply, it matters more how life feels from the inside to this person than how it looks on a test sheet. {{c:Sharf}}",
        "A third basis is holism, the view of the **whole person** in context. The approach refuses to shrink a human being down to symptoms, reflexes or a diagnostic label. {{c:Corsini}}",
        "A fourth basis, from **Maslow**, is the hierarchy of needs. Physiological, safety, belonging and esteem needs build up toward self-actualisation. A blocked lower need stalls the climb \u2014 it is hard to chase your dreams when you feel unsafe or unloved \u2014 and that stalling is what shows up as distress. {{c:Corey}}",
        "In **formulation**, distress is seen as **natural growth being blocked**. It is blocked by a harsh environment, by [[conditions of worth|the feeling that you are only loved or valued when you meet other people's standards]], or by losing touch with your own feelings. {{c:Corey}}",
        "A key formulation idea is that symptoms show **disconnection from experience**, not disease. So the aim is to restore awareness and self-direction, not to fix a defect. In plain terms, the person has drifted away from what they truly feel, and the symptom is the alarm bell. {{c:Sharf}}",
        "The **goals** are **self-awareness, self-acceptance and authenticity**, plus personal growth. Therapy helps the person take responsibility and move toward self-actualisation in their own way. {{c:Corey}}",
        "A defining **procedure** is that the **relationship itself** is the main healer. A genuine, accepting, empathic presence matters more than any particular technique. {{c:Corey}}",
        "**Empathic listening** is the core technique. How it is done: the therapist listens for the feeling under the words and reflects it back accurately, so the client hears their own inner world said aloud. Why it helps: feeling deeply understood is, by itself, freeing \u2014 it tells the person they are not alone and need not hide. {{c:Sharf}}",
        "**Reflection and acceptance** are the next techniques. How it is done: the therapist mirrors feelings without judging, advising or steering, following the client's lead rather than leading. Why it helps: unconditional acceptance dissolves the conditions of worth that first blocked the person's growth. {{c:Sharf}}",
        "**Focusing on the felt sense** is an experiential technique (Gendlin). How it is done: the therapist invites the client to notice a vague bodily feeling \u2014 a tightness, a heaviness \u2014 and stay with it until words arrive. Example: attending to a knot in the stomach may slowly reveal an unnamed grief. Why it helps: emotion lives in the body, so the body is a doorway to feelings the mind has hidden. {{c:Corey}}",
        "**Chair and experiential work** round out the techniques. How it is done: the client speaks to an imagined part of themselves or another person, bringing a felt conflict into the room to be experienced, not just discussed. {{c:Corey}}",
        "In **stages and process** the work is **client-led and unfolding**, not staged by the therapist. It moves as the person feels safe enough to explore, deepen awareness and take back ownership of their experience. {{c:Sharf}}",
        "A process hallmark is **genuine presence** and congruence from the therapist. This lowers the client's defences and frees the actualising tendency to resume its natural course. {{c:Corey}}",
        "The **outcome** sought is more **self-esteem and emotional awareness**, together with better relationships and a stronger inner sense of direction. {{c:Corey}}",
        "**Indications** cover mild to moderate distress, identity and relationship problems, depression and anxiety, and clients seeking growth. Severe psychosis or a need for structured skills training is better served, **as a contraindication**, by more directive work. {{c:Sharf}}",
        "In **research and current status**, well-known meta-analyses by **Elliott and colleagues** find humanistic-experiential therapies give large, stable gains. They are broadly equal to CBT in many comparisons. {{c:Corey}}",
        "The approach stays alive in modern forms like **emotion-focused therapy and motivational interviewing**. Critics still say its concepts are vague, hard to test and perhaps too rosy about human nature. {{c:Corey}}",
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
    {"t": "lead", "text": "Existential psychotherapy helps people face the hard facts of being alive: death, freedom, aloneness and the search for meaning. These facts stir up anxiety. Instead of dodging that anxiety through symptoms, the client meets it honestly and turns it into a more genuine, responsible and meaningful life."},
    {"t": "fig", "key": "existential_concerns", "caption": "Yalom's four ultimate concerns: facing death, freedom, isolation and meaninglessness honestly turns existential anxiety into authentic living."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Yalom** builds existential therapy around four [[ultimate concerns|the unavoidable facts of life that stir up anxiety in everyone]]: **death, freedom, isolation and meaninglessness**. Facing them causes anxiety, and dodging that anxiety is what turns into symptoms. {{c:Yalom}}",
        "{{THEORY}} **Frankl's** logotherapy centres the **will to meaning**. His book Man's Search for Meaning, born from his time in a concentration camp, argues that meaning can be found even in suffering we cannot escape. {{c:Frankl}}",
        "{{THEORY}} The approach draws on philosophers **Kierkegaard, Nietzsche, Heidegger and Sartre**, and on clinicians like **Binswanger, Boss and May**. This gives it a strongly philosophical flavour. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is facing the hard, unavoidable facts of being alive \u2014 that we will die, that we are free and responsible, that we are ultimately alone, and that life has no ready-made meaning. These facts stir up anxiety in everyone.",
        "The central claim is that **symptoms often come from dodging this anxiety** rather than from a disease. So the therapy does not try to remove the anxiety; it helps the person stop running from it and instead live more honestly, more responsibly and with more meaning in its presence.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change comes from **honest confrontation, not avoidance**. When the client stops numbing or distracting themselves and squarely faces, say, the fact of death or their own freedom, the dodged anxiety stops leaking out as symptoms and becomes a spur to live more fully.",
        "A simple example: a man who buries himself in overwork to avoid the thought of mortality may, once he faces it in therapy, choose to spend his evenings with his children instead. Facing the fact changed how he used his limited time. The awareness itself is what frees the choice.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, existential therapy has philosophical roots in **Kierkegaard, Nietzsche, Heidegger and Sartre**. Its clinical roots are in Binswanger and Boss's Daseinsanalysis, which applied phenomenology to the suffering person. {{c:Corey}}",
        "Its clinical **origin** runs through **Frankl's** logotherapy from the 1930s, **May's** book Existence in 1958, and **Yalom's** Existential Psychotherapy in 1980, which organised the field for English-speaking clinicians. {{c:Yalom}}",
        "The first basis is that life brings unavoidable **ultimate concerns**. Death, freedom, isolation and meaninglessness face everyone. They create an [[existential anxiety|the deep unease that comes from facing the basic facts of being alive]] that cannot simply be cured away. {{c:Yalom}}",
        "A second basis is that **existence precedes essence**. In plain words, you are not born with a fixed self; you create yourself through your choices. So [[authenticity|living by your own real values and choices rather than just going along with others]] stands against living a borrowed life. {{c:Corey}}",
        "A third basis, from **Heidegger**, is being-in-the-world across the physical, social and personal worlds. Frankl adds that a **will to meaning** \u2014 the hunger for life to mean something \u2014 is the deepest human drive, deeper than pleasure or power. {{c:Frankl}}",
        "In **formulation**, mental trouble is read as **inauthentic living**. It is the avoidance of freedom and responsibility, the loss of meaning, or the failure to cope with death and aloneness. {{c:Corey}}",
        "Frankl's distinctive formulation is the [[existential vacuum|a hollow feeling of emptiness and boredom that comes when life seems to have no meaning]], which he called noogenic neurosis. {{c:Frankl}}",
        "The **goals** are to help clients **face life's givens honestly**. Therapy builds awareness, responsibility, authenticity and meaning, and encourages the person to act even while anxiety is present. {{c:Yalom}}",
        "**Phenomenological exploration** is the core procedure. How it is done: the therapist deliberately sets aside their own assumptions and asks open, curious questions to enter the client's own world exactly as the client sees it. Why it helps: only by seeing the world through the client's eyes can the real avoidances be found. {{c:Corey}}",
        "**Using the here-and-now relationship** is a technique in itself. How it is done: the therapist treats the bond in the room as a live sample of how the client meets others, and encourages **'I' statements** ('I chose', 'I feel') in place of 'it happened to me'. Why it helps: owning your words is a first, concrete step to owning your life. {{c:Yalom}}",
        "**Paradoxical intention** is a signature logotherapy technique. How it is done: the client is asked to try hard to bring on the very thing they dread \u2014 someone who fears blushing tries to blush as red as possible. Why it works: you cannot force an anxious symptom on command, so the effort breaks the vicious circle of 'fear of the fear'. {{c:Frankl}}",
        "**Dereflection** is the second logotherapy technique. How it is done: the therapist turns the client's attention away from anxiously watching their own symptom and onto meaningful people and activities instead. Example: a man obsessed with his sleeplessness is guided to focus on his work and family, and the sleep often returns on its own. {{c:Frankl}}",
        "**Attitude change** (finding meaning in unavoidable suffering) is the third technique. How it is done: when a situation genuinely cannot be changed \u2014 a terminal illness, a bereavement \u2014 the therapist helps the client choose the stance they take toward it, which is the one freedom no one can remove. {{c:Frankl}}",
        "In **stages and process**, the work first builds a **genuine relationship** and explores the client's worldview. It then spots patterns of avoidance, deepens awareness of freedom and responsibility, and ends in committed action. {{c:Yalom}}",
        "The process is **collaborative and exploratory**, not driven by technique. The therapist walks alongside the client in facing hard truths, rather than repairing a defect. {{c:Corey}}",
        "The **outcome** sought is more **meaning and authenticity** and a greater ability to tolerate anxiety. The person then lives more deliberately in the face of life's givens. {{c:Yalom}}",
        "**Indications** include existential crises such as life transitions, midlife and bereavement, terminal or chronic illness, and depression marked by emptiness. Acute psychosis and a wish for quick symptom control are relative **contraindications**. {{c:Corey}}",
        "In **research and current status**, the evidence is **moderate**. Meaning-centred interventions help meaning, mood and quality of life, especially in people with cancer. Critics note the concepts are abstract and hard to pin down for testing. {{c:Yalom}}",
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
    {"t": "lead", "text": "Gestalt therapy is a hands-on, here-and-now approach from Fritz and Laura Perls. Its key idea is that awareness itself heals. The therapist sets up live experiments rather than giving interpretations. The client finishes 'unfinished business' and takes back disowned parts of themselves, so as to become more whole."},
    {"t": "fig", "key": "gestalt_cycle", "caption": "The Gestalt contact cycle: a need becomes figure, mobilises energy, reaches contact and withdraws; disturbances interrupt the cycle."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** describes [[gestalt therapy|a hands-on therapy focused on present-moment awareness and the whole person]] as an experiential approach. It stresses **present experience and awareness** and brings feelings, perceptions and actions together into the whole person. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} It was developed by **Fritz Perls, Laura Perls and Paul Goodman** from psychoanalysis, Gestalt psychology and field theory. It became popular in the 1960s at the Esalen Institute. {{c:Corey}}",
        "{{THEORY}} Its guiding idea is the [[paradoxical theory of change|the odd truth that you change by fully accepting who you are now, not by forcing yourself to be someone else]]. Beisser put it as change coming when a person fully becomes what they are. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is present-moment **awareness**. Gestalt therapy holds that simply becoming fully aware of what you are feeling, doing and avoiding right now is itself what heals \u2014 you do not need the therapist to explain your past to you.",
        "Instead of talking **about** a problem, the therapist sets up live **experiments** so the client experiences it directly in the room. The aim is to finish old 'unfinished business' and take back the parts of yourself you had disowned, so you feel whole again.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change works through **heightened awareness in the here-and-now**. By exaggerating a gesture, staying with a feeling or speaking to an empty chair, the client suddenly notices what they have been doing out of habit \u2014 and that fresh awareness, by itself, opens up a choice that was invisible before.",
        "This rests on the [[paradoxical theory of change|the odd truth that you change by fully accepting and becoming who you are now, not by forcing yourself to be someone else]]. Fighting a feeling keeps you stuck; fully owning it is what lets it move, much as you can only leave a room once you admit you are in it.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, Gestalt therapy was developed by **Fritz and Laura Perls with Paul Goodman** in the 1940s and 1950s. They drew on psychoanalysis, Gestalt psychology, Reich's body work and Lewin's field theory. {{c:Corey}}",
        "Its **origin** reached a wide public in the **1960s at Esalen**. There Perls' bold live demonstrations made the empty-chair style famous, for better and worse. {{c:Affective Therapies Reader}}",
        "The first basis is **holism and field theory**: person and environment form one single field. So behaviour is understood in its context, not in isolation. {{c:Corey}}",
        "A central basis is the **here-and-now**. The focus is on the **what and how** of present experience, not the why. This is because awareness in the present moment, not a theory about the past, is what heals. {{c:Corey}}",
        "A further basis is **figure and ground**. The strongest need of the moment rises up to become the [[figure|the one need or interest that stands out from everything else right now, like a craving when you are hungry]] against the background. Healthy living is a smooth flow of needs rising, being met and fading back. {{c:Sharf}}",
        "Another basis is [[unfinished business|old, unexpressed feelings from the past that keep nagging and spilling into the present]]. Like an argument you never got to finish, these press for attention and stop a person living fully now. {{c:Corey}}",
        "In **formulation**, trouble comes from **interruptions to contact** at the boundary with others. These are named introjection (swallowing others' rules whole), projection (seeing your own feeling in others), [[retroflection|turning back on yourself what you really want to do to someone else, like getting angry at yourself instead of them]], deflection (dodging real contact) and confluence (blurring into others). {{c:Corey}}",
        "A further formulation idea is **Perls' layers of neurosis**: the phony, phobic, impasse, implosive and explosive layers. A client passes through them, like peeling an onion, on the way to genuine contact. {{c:Affective Therapies Reader}}",
        "A vivid formulation is the **top-dog versus underdog** split. This is the inner tug-of-war between the bossy 'you should' (top-dog) and the sly, stubborn 'I won't' (underdog), which keeps a person stuck. {{c:Corey}}",
        "The **goals** are to **increase awareness**, take responsibility, join up split-off parts, finish unfinished business and shift from leaning on others toward self-support. {{c:Corey}}",
        "The **empty-chair technique** is the best-known Gestalt method. How it is done: the client imagines a person, or a part of themselves, sitting in an [[empty chair|an actual empty chair the client talks to, picturing someone absent or an inner part of themselves in it]], speaks to it, then swaps seats and replies as that other side. Why it helps: a bottled-up conflict becomes a live dialogue, so feelings that were stuck get expressed and the two sides can meet. {{c:Corey}}",
        "The **two-chair dialogue** is the related technique for inner splits. How it is done: the top-dog sits in one chair and the underdog in the other, and the client voices both in turn. Why it helps: the warring parts stop fighting in the dark and start hearing each other, which is how they integrate. {{c:Corey}}",
        "**Exaggeration and staying with the feeling** are core experiments. How they are done: the therapist asks the client to make a small gesture bigger (keep tapping that foot, harder) or to not escape a rising feeling but stay with it a little longer. Why they help: both turn up the volume on something half-noticed until it reaches full awareness. {{c:Sharf}}",
        "**'I' language** is a simple but powerful technique. How it is done: the client swaps 'it' and 'you' for **'I'**, and 'I can't' for 'I won't'. Example: 'my hand is shaking' becomes 'I am shaking', which quietly restores ownership of the experience and the choice. {{c:Corey}}",
        "**Dream work** is distinctive in Gestalt. How it is done: rather than interpret a dream, the therapist treats every part of it \u2014 even a wall or a dog \u2014 as a cast-off piece of the dreamer, who speaks as each part in turn. Why it helps: the client reclaims the disowned feelings the dream represents. {{c:Corey}}",
        "In **stages and process** the work moves from **contact and the I-Thou relationship**, through awareness work and experiments that heighten contact, to integration and the closing of what was unfinished. {{c:Corey}}",
        "The therapist's **process** stance is **authentic and present**, not interpretative. They offer experiments and honest, in-the-moment feedback rather than analysing from a distance. {{c:Corey}}",
        "The **outcome** sought is more **awareness and responsibility**, freer emotional expression, fuller integration and more lively relationships. {{c:Corey}}",
        "**Indications** include anxiety, depression, relationship and psychosomatic problems, and trauma-related unfinished business. Very confrontational experiments are a **contraindication** for fragile or psychotic clients, who may be overwhelmed. {{c:Sharf}}",
        "In **research and current status**, the evidence base is **limited but growing**. The strongest support comes through emotion-focused therapy, where Greenberg and colleagues show empty-chair work for unfinished business is effective. {{c:Corey}}",
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
    {"t": "lead", "text": "Person-centred therapy is Carl Rogers' non-directive approach. The therapist offers three things: being real, accepting the client fully, and truly understanding them. Rogers held that these are enough on their own to free the client's natural drive to grow. The client is the expert on themselves and leads their own change."},
    {"t": "fig", "key": "rogers_conditions", "caption": "Rogers' core conditions: congruence, unconditional positive regard and empathy free the client's actualising tendency."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **APA Dictionary of Psychology** describes [[person-centered therapy|Rogers' non-directive therapy built on being genuine, accepting and empathic]] as the non-directive form of therapy developed by **Carl Rogers**. The therapist's genuineness, unconditional positive regard and empathy help the client grow. {{c:APA Dictionary of Psychology}}",
        "{{THEORY}} **Rogers** introduced non-directive counselling in 1942. In 1957 he set out his **six necessary and sufficient conditions** for change, grounding the approach in the actualising tendency. {{c:Rogers}}",
        "{{THEORY}} The approach is firmly **humanistic yet distinct**. It is Rogers' specific model built on the core conditions, whereas humanistic therapy names the broader third-force movement. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is the quality of the relationship, not any technique. Rogers claimed that if the therapist offers just three attitudes \u2014 being genuine, fully accepting the client, and truly understanding them \u2014 that is **enough on its own** to free the client's natural drive to grow.",
        "The client, not the therapist, is treated as the **expert on their own life**. So the therapist never diagnoses, advises or steers. They follow the client, trusting that the person already carries within them the direction they need.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is simple: when a person is **accepted exactly as they are**, they no longer need to defend or hide, so they can look honestly at feelings they once denied. As the hidden feelings come into the open and are accepted, the gap between their real self and their self-image closes, and health returns.",
        "A small example: a man who was taught that anger is 'bad' has buried his, and feels tense and unreal. With a therapist who accepts even his anger without flinching, he slowly dares to feel it, name it, and own it \u2014 and the tension eases because he is no longer at war with part of himself.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, the approach is the work of **Carl Rogers**. His book Counseling and Psychotherapy in 1942 introduced non-directive counselling, and his Client-Centered Therapy in 1951 named the model. {{c:Rogers}}",
        "Rogers later broadened it into the **person-centred approach** for education, groups and conflict resolution. He also pioneered recording and studying real sessions, which started an empirical tradition in therapy research. {{c:Corey}}",
        "The foundational basis is the [[actualising tendency|the one inbuilt drive in every person to grow and fulfil their potential, like a plant growing toward the sun]]. It moves the person toward growth whenever conditions allow. {{c:Rogers}}",
        "A second basis is the **self-concept and the ideal self** (how you see yourself versus who you wish you were). A healthy person also trusts their own gut sense of experience as a guide, rather than only others' opinions. {{c:Rogers}}",
        "A third basis is [[conditions of worth|the belief that you only deserve love and approval when you meet other people's standards]]. They form when love was given with strings attached, and they push the person to deny feelings that clash with them. {{c:Corey}}",
        "A fourth basis is **incongruence**: the gap between how you see yourself and what you actually feel, or between your real and ideal self. This gap produces anxiety and defensiveness; closing it is the mark of health. {{c:Rogers}}",
        "In **formulation**, problems are seen as **incongruence caused by conditions of worth**. So therapy is not a technical repair; it simply provides the conditions in which the growth drive can start up again. {{c:Corey}}",
        "The decisive formulation claim is **Rogers' six conditions** for change, of which the three the therapist offers \u2014 the 'core conditions' \u2014 are congruence, unconditional positive regard and empathy. {{c:Rogers}}",
        "The first core condition is **congruence** (genuineness). How it shows up: the therapist is [[congruent|real and honest in the room, not hiding behind a professional mask or a fixed smile]], matching their outward manner to their true inner feeling. Why it helps: realness invites realness, so the client dares to be genuine too. {{c:Rogers}}",
        "The second core condition is **unconditional positive regard**. How it shows up: the therapist [[prizes the client|warmly accepts and values the client no matter what they bring, with no strings attached]] even when disapproving of a behaviour. Why it helps: being accepted whatever you reveal dissolves the conditions of worth that caused the trouble. {{c:Rogers}}",
        "The third core condition is **empathy**. How it shows up: the therapist accurately grasps the client's inner world and reflects it back, checking they have it right. Why it helps: feeling deeply and accurately understood lets the client trust and explore their own experience. {{c:Corey}}",
        "The guiding **goals** are more **congruence and self-acceptance**, judging yourself by your own standards rather than others', and becoming a fully functioning person who is open to experience. {{c:Rogers}}",
        "The defining **procedure** is being **non-directive**. How it is done: the therapist deliberately does not interpret, diagnose, advise or reassure. Why it matters: every piece of steering would quietly re-impose a new 'condition of worth', so holding back is itself the active ingredient. {{c:Corey}}",
        "**Reflective listening** is the central technique. How it is done: the therapist mirrors back the feeling and meaning underneath the client's words ('it sounds as if you felt abandoned'), paraphrases, clarifies and allows comfortable silences. Why it helps: the client hears their own inner world said back clearly, which deepens their self-understanding. {{c:Sharf}}",
        "In **stages and process**, Rogers described a **seven-stage continuum**. The client moves from being rigid and cut off from feelings toward feelings that flow freely, arrive in the moment, and can be trusted. {{c:Corey}}",
        "In practice the **process** is simple: the client talks and the therapist reflects. As the client feels understood and accepted, defences drop, **self-exploration deepens**, and self-directed change appears. {{c:Corey}}",
        "The **outcome** sought is more self-acceptance and congruence, an inner source of judgement, less anxiety and defensiveness, and better, self-directed relationships. {{c:Rogers}}",
        "**Indications** are wide: mild to moderate emotional, identity and relationship problems, grief, adolescents and student counselling. Clients wanting direct advice, or with severe disorders or acute crises, are relative **contraindications**. {{c:Sharf}}",
        "In **research and current status**, Rogers was a **pioneer of therapy research**. Meta-analyses by Elliott and colleagues find empathy, positive regard and congruence link modestly but reliably with good outcome, though critics doubt they are enough on their own. {{c:Corey}}",
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
    {"t": "lead", "text": "Adlerian therapy, or Individual Psychology, is Alfred Adler's approach. It sees the person as a whole, as aiming at goals, and as shaped by their social world. Problems are treated as discouragement and mistaken beliefs about life, not illness. The therapy re-educates the client toward courage, cooperation and useful contribution."},
    {"t": "fig", "key": "adlerian_phases", "caption": "The four Adlerian phases: relationship, lifestyle assessment, insight and reorientation toward social interest."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Adler** founded **Individual Psychology** after leaving Freud's circle in 1911. He stressed that the person is one indivisible whole, that behaviour is **goal-directed**, and that it is shaped by social context. {{c:Adler}}",
        "{{THEORY}} Its hallmark concept is [[social interest|a built-in capacity and willingness to cooperate with and care about other people]]. Adler treated it as the very measure of mental health. {{c:Adler}}",
        "{{THEORY}} The approach was carried forward by **Dreikurs, Ansbacher, Mosak and Dinkmeyer** into child guidance, family work and parent education. {{c:Corey}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is the person as a whole, aiming at goals, and set in their social world. Adler saw us not as pushed from behind by drives but as **pulled forward by goals** \u2014 a private picture of how to feel capable and belong. Problems are treated as **discouragement and mistaken beliefs about life**, not as illness.",
        "The yardstick of health is [[social interest|a built-in capacity and willingness to cooperate with and care about other people]]. So the whole therapy re-educates the client toward courage, cooperation and useful contribution, rather than curing a disease.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change works through **encouragement plus insight into one's lifestyle**. The therapist first helps the client see the mistaken, self-defeating beliefs they formed as a child, then rebuilds their courage so they dare to act differently and connect with others.",
        "The key idea is that **a discouraged person behaves badly**; a child who believes 'I only matter if I am the centre of attention' will misbehave to get it. Replace the belief and restore the courage, and the behaviour changes on its own \u2014 much as a plant straightens once you stop blocking its light.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, Adlerian therapy was founded by **Alfred Adler**, a Viennese physician and early member of Freud's circle. He broke away in 1911 to set up Individual Psychology. {{c:Adler}}",
        "Its **origin** also includes **Rudolf Dreikurs**, who adapted Adler's ideas for children, families and the democratic classroom. Adler founded the first child-guidance clinics in 1920s Vienna. {{c:Corey}}",
        "The first basis is **holism and teleology** (being pulled by future goals). The person is indivisible, and behaviour aims at a self-made goal, a [[fictional finalism|an imagined future goal you invent for yourself that quietly steers all your striving]]. {{c:Corey}}",
        "A central basis is **inferiority and compensation**. Natural feelings of being 'not enough' spur a striving to feel capable and significant. This can grow into healthy ambition or, if it overwhelms the person, into a self-defeating inferiority complex. {{c:Adler}}",
        "A defining basis is **social interest**: the inborn capacity to cooperate and contribute. Adler made it the yardstick of mental health, so well-being grows as you feel more connected to and useful to others. {{c:Adler}}",
        "A further basis is the [[lifestyle|the unique set of beliefs and goals about life that you form very early in childhood and then live by]]. It is largely shaped by about age five from the family set-up and birth order, and then acts like a pair of tinted glasses through which you read everything. {{c:Corey}}",
        "Another basis is the contrast between **private logic and common sense**. Private logic is a person's own quirky, mistaken beliefs \u2014 their 'basic mistakes' \u2014 which steer them away from cooperative living. {{c:Sharf}}",
        "In **formulation**, problems come from **discouragement, mistaken beliefs and low social interest**. Difficulties are seen as problems of living, not illness; in other words, a way of life that is not working. {{c:Corey}}",
        "A practical formulation tool is **Dreikurs' four mistaken goals** of a child's misbehaviour: getting attention, seeking power, taking revenge, and showing inadequacy to be left alone. Naming the goal tells you what the child is really after. {{c:Corey}}",
        "The **goals** are to **boost social interest and belonging**, correct mistaken lifestyle beliefs, build courage and self-worth, and swap the striving driven by inferiority for useful contribution. {{c:Adler}}",
        "**Phase one, the relationship**, is the first procedure. How it is done: the therapist builds an equal, respectful, collaborative partnership \u2014 two people side by side, never expert-over-patient. Why it matters: a discouraged person can only risk change inside a warm, non-judging bond. {{c:Corey}}",
        "**Phase two, lifestyle assessment**, is a signature technique. How it is done: the therapist explores the family set-up, birth order and especially [[early recollections|your very earliest specific memories, read as clues to the beliefs you live by now]], treating these first memories like a projective snapshot of the client's core beliefs. {{c:Corey}}",
        "**Phase three, insight**, follows. How it is done: the therapist gently and tentatively (never bluntly) points out the client's mistaken beliefs, private logic and hidden goals, offering them as a guess to try on. {{c:Corey}}",
        "**Phase four, reorientation**, turns insight into action, and its engine is **encouragement**, the central Adlerian technique. How it is done: the therapist steadily highlights strengths and effort, not just results, so the client builds the courage to choose new, cooperative behaviour. {{c:Corey}}",
        "**Acting 'as if'** is a distinctive action technique. How it is done: the client behaves for a week as though they already had the quality they lack \u2014 [[acting as if|behaving as though you already had the confidence or skill you want, to try it on for size, like rehearsing a role]] \u2014 which breaks the 'I could never' belief by direct experience. {{c:Sharf}}",
        "**Spitting in the client's soup** and **catching oneself** are further techniques. Spitting in the soup means gently naming the hidden payoff of a behaviour so it loses its appeal (once you see you sulk to get sympathy, sulking feels less sweet). Catching oneself trains the client to spot their old pattern starting and stop it, aided by the push-button technique and homework. {{c:Sharf}}",
        "In **stages and process**, the therapist works like an **encouraging teacher or coach**. Therapy is treated as education and re-education, not the cure of a disease. {{c:Corey}}",
        "The **outcome** sought is less inferiority and discouragement, more courage, greater **social interest**, and a more workable lifestyle and relationships. {{c:Adler}}",
        "**Indications** are broad: children and adolescents, parent education, family therapy, school and career counselling and group work. It fits collectivist cultures whose values match social interest, while severe psychopathology needs **adaptation**. {{c:Corey}}",
        "In **research and current status**, the evidence is **limited but positive**. There is support for Adlerian parent education such as STEP, for Adlerian play therapy and for counselling. Adler's ideas are recognised as foreshadowing cognitive, humanistic and family-systems therapies. {{c:Corey}}",
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
    {"t": "lead", "text": "Transactional analysis is Eric Berne's down-to-earth approach. It reads personality as three inner 'ego states': Parent, Adult and Child. It sees problems as mixed-up ego states, repeated 'games', and a limiting life story. Using a clear, agreed contract, it works to restore autonomy and honest, game-free closeness."},
    {"t": "fig", "key": "ta_states", "caption": "The three ego states and the four types of transaction, with the Karpman drama triangle that games play out."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Eric Berne** founded transactional analysis in the late 1950s. He set out the theory in Transactional Analysis in Psychotherapy and made it popular in Games People Play. He defined the goal as **autonomy** through awareness, spontaneity and intimacy. {{c:Berne}}",
        "{{THEORY}} A [[transaction|one back-and-forth exchange between two people: you say something, I respond]] is the basic unit of social life. Berne saw every exchange as coming from one of three **ego states**: the Parent, the Adult or the Child. {{c:Berne}}",
        "{{THEORY}} The approach was carried forward by **Harris, Steiner, Karpman and the Gouldings**, who added the OK life positions, scripts, the drama triangle and redecision therapy. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is how we think, feel and relate, pictured through three inner 'selves' called [[ego states|three consistent ways of thinking, feeling and acting that we all switch between, called Parent, Adult and Child]]. At any moment you are speaking from one of them: the **Parent** (copied from your caregivers), the **Adult** (calm, here-and-now, factual) or the **Child** (your feelings and playfulness from long ago).",
        "Problems are read as mixed-up ego states, repeated hidden-motive 'games', and a limiting life-story written in childhood. The whole therapy uses a clear, agreed contract to restore **autonomy** and honest, game-free closeness, in plain everyday language anyone can use.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change works by **strengthening the Adult and rewriting the script**. When the client can see which ego state they are in and which game they keep playing, they can choose to answer from the calm, reality-based Adult instead of being yanked around by an old Parent rule or a hurt Child feeling.",
        "A clear example: a woman keeps starting the game 'Why don't you \u2014 Yes but' (she asks for advice, then rejects every suggestion) and ends up feeling misunderstood again, which confirms her childhood script 'nobody can help me'. Once she spots the game, she can stop playing it and ask for help straight, Adult-to-Adult.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, transactional analysis was created by **Eric Berne**, a Canadian-born psychiatrist. He broke from psychoanalysis to build a plain-language social psychiatry, publishing Transactional Analysis in Psychotherapy in 1961 and Games People Play in 1964. {{c:Berne}}",
        "Its **origin** then widened through **Thomas Harris** (whose I'm OK, You're OK spread the life positions), **Claude Steiner** on scripts, **Stephen Karpman** on the drama triangle, and the **Gouldings** on redecision therapy. {{c:Affective Therapies Reader}}",
        "The core basis is the three **ego states**. The **Parent** is the taught, judging, caring voice copied from caregivers; the **Adult** is the calm, factual problem-solver dealing with the here-and-now; the **Child** holds the feelings, wants and playfulness carried from early life. {{c:Berne}}",
        "A second basis is the analysis of [[transaction|one back-and-forth exchange between two people: you say something, I respond]]s. They can be **complementary** (replies come from the expected state, so talk flows), **crossed** (an unexpected state replies, so talk breaks off), or [[ulterior|carrying a hidden message underneath the polite surface one, like saying 'nice shirt' to mean the opposite]], which drives games. {{c:Sharf}}",
        "A third basis is **strokes**, the units of recognition we all hunger for. A stroke is any sign of being noticed; positive or negative, conditional or not. Berne's point is that we need them so badly that we will provoke negative ones rather than be ignored, which shapes how we fill our time. {{c:Berne}}",
        "A fourth basis is **games and rackets**. Games are [[games|repeated hidden-motive exchanges that always end in the same bad feeling]], like 'Why don't you, Yes but'. They pay off in a familiar bad 'racket' feeling (such as righteous anger) that quietly props up the life script. {{c:Berne}}",
        "A fifth basis is the [[life script|an unconscious life story you decided on as a child, based on messages from your parents, and then act out]]. It is built from injunctions like 'don't be close', driver messages like 'be perfect', and the child's own early decisions about how life will go. {{c:Affective Therapies Reader}}",
        "A sixth basis is the four **life positions** of the OK Corral. They run from the healthy 'I'm OK, you're OK' down to 'I'm OK, you're not' (blaming), 'I'm not OK, you're OK' (self-doubting) and 'I'm not OK, you're not OK' (hopeless), each colouring a person's whole outlook. {{c:Affective Therapies Reader}}",
        "In **formulation**, problems are traced to **contaminated or excluded ego states** (for example Parent prejudices leaking into the Adult), destructive games, a negative script and an unhealthy life position. The therapist maps which ego state and which game is driving the complaint. {{c:Sharf}}",
        "A vivid formulation tool is **Karpman's drama triangle**. Clients switch between Persecutor, Victim and Rescuer roles. The healthier exit from each is assertiveness, honest vulnerability and genuine caring. {{c:Affective Therapies Reader}}",
        "The **goals** are **autonomy and game-free intimacy**: a free, uncontaminated Adult, a revised script (a 'script cure'), better self-control and relief of the presenting symptom. {{c:Berne}}",
        "The **contract** is a defining technique, not just a formality. How it is done: at the start, client and therapist agree in plain words on a clear, observable goal and their shared roles. Why it helps: it sets the equal, adult-to-adult, 'you're OK, I'm OK' tone and keeps both accountable. {{c:Sharf}}",
        "The **four analyses** are the main technique set. **Structural analysis** maps the client's three ego states; **transactional analysis** studies their exchanges with others; **game analysis** uncovers the hidden payoff of repeated bad exchanges; and **script analysis** traces the childhood life-story driving it all. Together they move the problem from vague to visible. {{c:Berne}}",
        "**Berne's eight therapeutic operations** are the therapist's fine tools, running from interrogation, specification and confrontation, through explanation, illustration, confirmation and interpretation, up to [[crystallisation|the moment the client can finally hold a new, freeing view of themselves on their own]]. Each is a graded step from gathering facts to the client's own insight. {{c:Berne}}",
        "**Redecision work** is the powerful change technique (the Gouldings). How it is done: using chair work, the client goes back to the early scene where they made a limiting decision ('I must never show need') and, from the Adult, makes a new decision now. Supported by the **egogram**, permissions, stroke work and homework, usually in a group. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work runs from **contracting and assessment**, through strengthening and 'decontaminating' the Adult and freeing the Child, to relearning new behaviours and finally autonomy and termination. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is more autonomy, clearer communication, an **end to games**, a revised script and better self-esteem. The simple language lets clients keep using the ideas after therapy ends. {{c:Sharf}}",
        "**Indications** cover neurotic, relationship and communication problems, personality issues, couples and families, groups, education and organisations. Acute psychosis is a **contraindication** served only by adapted forms. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is **modest**. Some reviews report good outcomes, but few rigorous trials exist. TA stays widely taught and accredited by the ITAA and EATA, despite criticism of its jargon and limited testing. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Reality therapy is William Glasser's present-focused approach, built on choice theory. It holds that we choose almost all of our behaviour to meet five basic needs. Within a warm, non-blaming relationship, it uses the WDEP system to help clients judge their own choices and plan better, more responsible ones."},
    {"t": "fig", "key": "reality_wdep", "caption": "The WDEP cycle inside a trusting environment: Wants, Doing, self-Evaluation and Planning, driving more effective choices."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **William Glasser** developed reality therapy in the 1960s. He later grounded it in [[choice theory|the idea that we choose nearly all of our behaviour to try to meet our basic needs]], and that we can directly control only our own acting and thinking. {{c:Glasser}}",
        "{{THEORY}} Glasser held that **most psychological problems are relationship problems**, rooted in unsatisfying connections. So he rejected diagnostic labels in favour of helping clients make better choices. {{c:Glasser}}",
        "{{THEORY}} **Robert Wubbolding** organised the practice into the **WDEP system**, set inside a cycle of counselling that first builds a trusting, friendly bond. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is present choices and responsibility. Reality therapy holds that we **choose almost all of our behaviour** to try to meet five basic needs, and that we can directly control only what we do and think \u2014 not, directly, how we feel. So the work stays firmly on the present and on better choices, not on the past or on blame.",
        "Glasser pictured it as a car with **four wheels**: thinking and acting are the front wheels we steer with, while feelings and body are the back wheels that follow. Change where you drive (what you do and think) and the back wheels \u2014 your feelings \u2014 come along behind.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "Change works through **self-evaluation inside a caring relationship**. The therapist does not scold or diagnose; within a warm bond they keep asking, kindly, 'Is what you are doing actually getting you what you want?'. When the client honestly answers 'no', they become willing to plan something better.",
        "The crucial move is that the **client judges their own behaviour**, not the therapist. A teenager who skips school to feel powerful is led to see for himself that it is costing him the future he wants; owning that judgement is what powers a realistic new plan he will actually keep.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, reality therapy was founded by **William Glasser**, an American psychiatrist. Unhappy with Freudian therapy, he developed his ideas at the Ventura School for Girls and a veterans' hospital, publishing Reality Therapy in 1965. {{c:Glasser}}",
        "Its **origin** continued as Glasser's 1980s control theory was renamed **choice theory** in 1998. **Robert Wubbolding** then built the WDEP system and the cycle of counselling that organise practice today. {{c:Affective Therapies Reader}}",
        "The first basis is that **all behaviour is chosen**, which Glasser grounded in [[choice theory|the idea that we choose nearly all of our behaviour to try to meet our basic needs]]. We can directly control only our acting and thinking, which then steer our feelings and body \u2014 the four-wheel car picture. {{c:Glasser}}",
        "A second basis is the five genetic [[basic needs|the five needs that drive all of us: survival, love and belonging, power, freedom and fun]]. Love and belonging is treated as the most important for mental health, which is why Glasser said most problems are relationship problems. {{c:Glasser}}",
        "A third basis is the [[quality world|a private mental picture album of the specific people and things you most want in your life]]. We constantly compare the real world to this album, and a gap between the two is what drives our behaviour to close it. {{c:Affective Therapies Reader}}",
        "A fourth basis is that **relationships are central**. The seven deadly habits (criticising, blaming, complaining, nagging, threatening, punishing, bribing) damage connection; the seven caring habits (supporting, encouraging, listening, accepting, trusting, respecting, negotiating differences) rebuild it. {{c:Affective Therapies Reader}}",
        "A fifth basis is **responsibility and a present focus**. You meet your needs without stopping others meeting theirs, with no excuses, no blame and no punishment, and almost no attention to the past or to symptoms. {{c:Glasser}}",
        "In **formulation**, the client is seen as having an **unmet need**, usually love and belonging. They then choose ineffective behaviours \u2014 withdrawing, 'depressing' (Glasser's verb, to stress it is chosen) or blaming \u2014 to try to meet it, which builds a failure identity. {{c:Affective Therapies Reader}}",
        "The **goals** are to help clients **take responsibility and make effective choices**, improve their relationships, and meet needs in ways that are responsible, realistic and right \u2014 Glasser's three Rs. {{c:Glasser}}",
        "**Building the environment** is the first procedure, from Wubbolding's cycle. How it is done: the therapist forms a trusting, friendly bond using warmth, humour and genuine listening, while firmly avoiding arguing, criticising and accepting excuses. Why it matters: a person will only evaluate themselves honestly inside a relationship that feels safe. {{c:Affective Therapies Reader}}",
        "**W \u2014 exploring Wants** is the first WDEP step. How it is done: the therapist asks what the client truly wants \u2014 from life, from others, from therapy \u2014 drawing out the pictures in their quality world. Why it helps: you cannot steer toward a goal you have never named. {{c:Affective Therapies Reader}}",
        "**D \u2014 Doing and direction** is the second step. How it is done: the therapist asks in concrete detail what the client is actually doing right now and where that road is taking them. Why it helps: it makes the current behaviour visible so it can be weighed. {{c:Affective Therapies Reader}}",
        "**E \u2014 self-Evaluation** is the pivot of the whole method. How it is done: the therapist asks the client to judge, for themselves, whether their current behaviour is getting them what they want and is realistic. Why it is central: change only begins when the person \u2014 not the therapist \u2014 decides the old way is not working. {{c:Affective Therapies Reader}}",
        "**P \u2014 Planning** is the final step, made memorable as the [[SAMI2C3 plan|a plan that is Simple, Attainable, Measurable, Immediate, Involving, Controlled, Committed and Consistent]]. How it is done: client and therapist craft one small, concrete, do-it-now step the client commits to, so success is easy to see and build on. {{c:Affective Therapies Reader}}",
        "Further **techniques** include skilful questioning, metaphor and humour, paradox, role play, written contracts and self-disclosure. The gentle, repeated challenge \u2014 'Is what you are doing helping?' \u2014 runs through all of them, always with a firm refusal to accept excuses. {{c:Affective Therapies Reader}}",
        "In **stages and process** the therapist builds the environment, finds out what the client wants, explores current behaviour, guides self-evaluation, helps form and commit to a plan, and then **follows up**, revising the plan if it fails. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is more responsibility, a stronger **internal locus of control** (feeling in charge of your own life), better relationships, higher self-esteem and more satisfying behaviour, in a short and practical course. {{c:Affective Therapies Reader}}",
        "**Indications** are broad: schools, adolescents and offenders, corrections, addictions, marital and group counselling and the workplace. Severe mental illness or trauma needing exploration is a **caution**. {{c:Affective Therapies Reader}}",
        "In **research and current status** the evidence is **moderate but often quasi-experimental**. Studies in schools, correctional and cross-cultural settings show gains in self-esteem, locus of control and relationships. Critics attack its rejection of diagnosis and its simple view of choice. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Supportive psychotherapy sits at the ego-strengthening end of therapy. It does not dig into the unconscious. Instead it props up and restores the client's coping, self-esteem and everyday functioning. It does this through a warm, active relationship, using reassurance, advice, praise and practical problem solving."},
    {"t": "fig", "key": "supportive_continuum", "caption": "The supportive to expressive continuum: supportive work bolsters existing defences; expressive work uncovers the unconscious."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} Supportive psychotherapy sits at the **supportive end of Luborsky's supportive-expressive continuum**. Think of a scale: one end uncovers hidden conflict, the other end strengthens coping. Supportive work aims to keep up or improve ego functions and self-esteem, not to uncover conflict. {{c:Affective Therapies Reader}}",
        "{{THEORY}} It treats the **therapeutic relationship** as the main healer. The therapist is a trusted, caring support, or an [[auxiliary ego|a dependable stand-in who lends the client the mental strength they temporarily lack, like a crutch for a sprained ankle]]. {{c:Affective Therapies Reader}}",
        "{{THEORY}} It was described by **Wallerstein, Rockland, and Winston, Pinsker and McCullough** for patients with limited ego strength who need shoring up rather than insight. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is to **strengthen the person**, not to dig into them. Supportive therapy does the opposite of uncovering work: instead of opening up hidden conflicts, it reinforces the client's existing strengths, defences and coping so they can get through the day.",
        "Picture someone whose house is in a storm. Deep therapy would rebuild the foundations; supportive therapy props up the walls so the house stands through the storm. The aim is to **keep the person steady and functioning**, raise their self-esteem, and carry them through a crisis or a long illness, using the warm relationship itself as the main support.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is a **borrowed strength plus relief of pressure**. A dependable, warm therapist acts as an [[auxiliary ego|a stand-in source of steadiness and judgement that the client leans on until their own recovers]], so the client borrows the calm and good sense they cannot yet manage alone \u2014 much like leaning on a crutch until a sprained ankle heals.",
        "At the same time, letting feelings out, getting clear reassurance and advice, and solving one practical problem at a time all lower the pressure on an overwhelmed mind. As the stress drops and small successes build, the client's own coping and confidence slowly come back, and the borrowed support is needed less.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, supportive psychotherapy grows from the earliest practice of psychiatry, its **reassurance, advice and moral treatment**. It also grows from psychodynamic therapy at the supportive end of Luborsky's continuum. {{c:Affective Therapies Reader}}",
        "Its **origin** was shaped by **Wallerstein's Menninger project**, by Rockland, and by Winston, Pinsker and McCullough. They formalised it for patients whose fragile personality cannot handle uncovering work. {{c:Affective Therapies Reader}}",
        "The first basis is a change of aim. The goal is to **maintain, restore or improve ego functions** and self-esteem, and to shore up existing defences and coping, not to make the unconscious conscious. {{c:Affective Therapies Reader}}",
        "A second basis is that the **relationship itself heals**. A dependable, warm therapist acts as an auxiliary ego, lending the strength and steadiness the client temporarily lacks. {{c:Affective Therapies Reader}}",
        "A third basis is a **here-and-now focus**. The work stays on present stresses, coping and daily functioning, not on childhood reconstruction or deep transference. {{c:Affective Therapies Reader}}",
        "A fourth basis is how transference is handled. **Positive transference is used but not interpreted**, and deep interpretation is deliberately avoided, so the client is not destabilised. In short, the therapist uses the client's trust but does not analyse it. {{c:Affective Therapies Reader}}",
        "In **formulation**, symptoms are read as arising when **stress overwhelms a person's coping**. An overwhelmed ego, or a fragile personality, needs its functioning and adaptation strengthened. {{c:Affective Therapies Reader}}",
        "The **goals** are **symptom relief and stabilisation**, better adaptation and coping, support through a crisis or illness, prevention of relapse and rehospitalisation, and higher self-esteem. {{c:Affective Therapies Reader}}",
        "A defining **procedure** is the therapist's stance, which is **active, warm and flexible**, and often directive. The therapist meets the client where they are, rather than waiting in silence for associations. {{c:Affective Therapies Reader}}",
        "Core **techniques** include **reassurance, encouragement and praise**, empathic listening, and [[ventilation|letting the client talk out and release pent-up feelings, like letting steam out of a pot]]. Advice and guidance are given where the client cannot yet decide alone. {{c:Affective Therapies Reader}}",
        "Further **techniques** are **psychoeducation, explanation and clarification**, anticipatory guidance that readies the client for a stressful event, and practical problem solving and reality testing. {{c:Affective Therapies Reader}}",
        "More **techniques** include **reframing** (helping the client see a situation in a more helpful light), changing the environment (such as involving family or arranging services), modelling, limit-setting, strengthening good defences and timely crisis intervention. {{c:Affective Therapies Reader}}",
        "The **frequency and duration** are flexible. Sessions may be weekly or monthly, and the work can be brief and time-limited or long-term maintenance for chronic conditions. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work moves from assessing **strengths and ego capacities**, through building the alliance and defining problems and goals, to choosing and applying supportive techniques. {{c:Affective Therapies Reader}}",
        "The later **process** is monitoring progress, with a gradual reduction or ongoing maintenance, and an open-ended, flexible ending rather than a sharply scheduled termination. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **better coping and daily functioning**, less distress and fewer relapses. Stabilisation is valued above the deeper personality change that expressive therapy pursues. {{c:Affective Therapies Reader}}",
        "**Indications** are wide: schizophrenia and other severe mental illness, bipolar and personality disorders with low ego strength, crises and bereavement, chronic or terminal medical illness, the elderly and those with low insight. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Wallerstein's Menninger study** found supportive therapy gave more lasting change than expected. Later work shows it comparable to other therapies, yet it is under-researched and often used only as a control condition, and critics warn it may foster dependency. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Clinical hypnotherapy uses hypnosis, a state of narrow, focused attention in which a person responds more readily to suggestion. In that state the therapist helps change how the client perceives, feels and behaves. A session moves through five steps and is mostly used as an add-on to treat pain, anxiety and habits."},
    {"t": "fig", "key": "hypnotherapy_stages", "caption": "The five stages of a hypnotherapy session: pre-induction, induction, deepening, therapeutic phase and termination."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and classic authors", "items": [
        "{{GUIDE}} The **American Psychological Association (Division 30)** describes [[hypnosis|a state of narrowly focused attention, less awareness of surroundings, and a greater readiness to respond to suggestion]]. It is a state of consciousness, not sleep. {{c:APA Division 30}}",
        "{{THEORY}} **Milton H. Erickson** shaped modern practice with a **permissive, indirect** style. He used the client's own imagery and words rather than giving bossy commands. {{c:Erickson}}",
        "{{THEORY}} Theories split into **state views**, like Hilgard's neodissociation (hypnosis is a real altered state), and **non-state or socio-cognitive views** that stress role-taking, expectancy and imagination. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is **focused attention put to use**. In hypnosis the mind narrows to one thing and the usual critical, doubting part steps back, so helpful suggestions can take hold far more easily than in ordinary waking talk.",
        "Think of how a child, lost in a story, really feels afraid or excited \u2014 attention is so absorbed that the imagined becomes almost real. Hypnotherapy deliberately creates that absorbed state and then feeds in useful ideas (calm instead of panic, numbness instead of pain). It is **not sleep and not loss of control**; the client stays aware and could stop any time. It is mostly used as an **add-on** to other treatment, not a therapy on its own.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is **heightened suggestibility inside a narrowed focus**. When attention is absorbed and the critical, judging mind relaxes, a suggestion meets less resistance and is accepted more like a direct instruction \u2014 so 'your hand feels numb' or 'the worry drifts away' can actually reshape how the body and feelings respond.",
        "The gains are then made to last by **post-hypnotic suggestion and self-hypnosis**: the helpful response is linked to a cue the client can use later, and they are taught to re-enter the calm state on their own. In short, hypnosis opens a window; the suggestions and home practice are what carry the change out of the room and into daily life.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, hypnotherapy descends from **Mesmer's animal magnetism** in the 1770s, through Puysegur's artificial somnambulism, and James Braid, who in 1843 coined 'hypnosis' from the Greek for sleep. {{c:Affective Therapies Reader}}",
        "Its **origin** runs on through the **Nancy school of Liebeault and Bernheim** (who saw hypnosis as suggestion), Charcot and Janet on hysteria and dissociation, and Breuer and Freud's early catharsis under hypnosis, which Freud later gave up. {{c:Affective Therapies Reader}}",
        "The modern **origin** is dominated by **Milton Erickson's** permissive, indirect method, by Hilgard's neodissociation theory with the Stanford scales, and by socio-cognitive theorists like Barber, Sarbin, Spanos and Kirsch. {{c:Erickson}}",
        "The first basis is the nature of the hypnotic state: **absorption, focused attention and suggestibility** with reduced critical judgement. Importantly, hypnosis is **not sleep** and not a loss of control. {{c:APA Division 30}}",
        "A second basis is the debate between **state and non-state theories**. One camp sees a real altered state of dissociation; the other sees role-taking, expectancy and imagination. Brain imaging shows changes in attention and default-mode networks. {{c:Affective Therapies Reader}}",
        "A third basis is [[hypnotisability|how easily a given person responds to hypnotic suggestion, which stays fairly fixed, like a stable trait]]. It varies across people and is measured by scales like the Stanford and Harvard group scales. {{c:Affective Therapies Reader}}",
        "In **formulation**, symptoms, habits, pain and anxiety are seen as kept going by **automatic learned responses** and by negative beliefs and imagery. Focused attention and suggestion can be used to change them. {{c:Affective Therapies Reader}}",
        "The **goals** are **symptom control**, especially of pain and anxiety, habit change, relaxation, more self-efficacy and ego-strength, and teaching self-hypnosis so the client can manage symptoms alone. {{c:Affective Therapies Reader}}",
        "A core **procedure** is the **pre-induction phase**. It secures consent, takes a history, explains hypnosis, and clears up the myths of lost control or revealed secrets, while building expectancy and rapport. {{c:Affective Therapies Reader}}",
        "The **induction technique** then narrows attention through eye fixation, progressive relaxation, breathing or imagery. Ericksonian inductions use **confusion, storytelling and pacing and leading** rather than direct command. {{c:Erickson}}",
        "The **deepening technique** follows, using counting down, staircase imagery or [[fractionation|lifting the person out of trance and taking them back in again, repeatedly, so each dip goes deeper]], to intensify the experience before therapeutic work. {{c:Affective Therapies Reader}}",
        "In the **therapeutic phase** the clinician uses **direct and indirect suggestions, metaphors and ego-strengthening**, plus imagery and rehearsal. Where helpful, hypnotic analgesia such as glove anaesthesia is used for pain. {{c:Affective Therapies Reader}}",
        "Deeper **techniques** in that phase include age regression and [[hypnoanalysis|uncovering and working through buried memories or feelings while the person is under hypnosis]], post-hypnotic suggestions and anchoring, desensitisation and parts therapy, all used with care. {{c:Affective Therapies Reader}}",
        "The final **technique** is termination: a careful **re-alerting or dehypnosis** by counting up and reorienting the client, followed by discussion and teaching self-hypnosis for ongoing practice. {{c:Affective Therapies Reader}}",
        "In **stages and process** the session therefore follows the fixed order of **pre-induction, induction, deepening, therapeutic phase and termination**, each step set up before the next is tried. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is less pain and anxiety, changed habits and greater self-efficacy. Gains are stronger when the client practises **self-hypnosis** between sessions. {{c:Affective Therapies Reader}}",
        "**Indications** include chronic and procedural pain, labour and dental pain, anxiety and phobias, irritable bowel syndrome, insomnia, habit disorders and psychosomatic problems. **Contraindications and cautions** cover psychosis, severe personality disorder and the risk of false memory in forensic use. {{c:Affective Therapies Reader}}",
        "In **research and current status**, meta-analyses find that **adding hypnosis to CBT improves outcomes**. Evidence is strong for pain, irritable bowel syndrome and procedural distress, moderate for anxiety, and weaker or unclear for smoking and weight loss. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Play therapy treats play as the child's natural language. In it, toys become the child's words and play becomes the child's speech. Using a safe playroom and a trained therapist, in either directive or non-directive form, it helps children express and work through emotions, trauma and relationship difficulties."},
    {"t": "fig", "key": "play_classification", "caption": "A classification of play therapy from non-directive child-centred work to directive and parent-mediated forms."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} Play therapy rests on the idea that **play is the child's natural language**. As Landreth put it, toys are the child's words and play is the child's speech. Children show in play what they cannot yet say in words. {{c:Axline}}",
        "{{THEORY}} **Virginia Axline** set out the **non-directive, child-centred** form in 1947, with its eight guiding principles. She applied Rogers' conditions to children. {{c:Axline}}",
        "{{THEORY}} Its roots reach back to **Anna Freud and Melanie Klein**, who first used play in analysis, and forward to Landreth's child-centred play therapy and Knell's cognitive-behavioural play therapy. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is to meet children **in their own language, which is play**. A young child cannot sit and explain, in words, that they feel scared, angry or unsafe \u2014 but they can show it with toys. So play therapy treats **toys as words and play as speech**.",
        "A frightened child may make two dolls fight, or bury a toy in the sand and dig it up again, working through a real worry at a safe, pretend distance. The therapist gives a warm, accepting playroom and follows the child's lead, so the child can **express, understand and master** feelings and experiences they could never put into a sentence.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is **safe expression plus mastery through repetition**. By acting a frightening or confusing experience out in symbolic play, the child lets the bottled-up feeling out at a safe distance (it is only the dolls, not really them) and, by replaying it with a better ending, moves from feeling helpless to feeling in charge of the story.",
        "The accepting relationship does the rest: when the therapist reflects the child's feelings and accepts them without blame, the child learns the feelings are allowed and manageable. Over many sessions this builds self-esteem, emotional control and trust \u2014 and when parents are trained to play this way too, as in filial therapy, the healing continues at home.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, play therapy begins with **Freud's case of Little Hans** in 1909 and Hermine Hug-Hellmuth, the first to use play with children. Then **Anna Freud and Melanie Klein** made play central to child analysis. {{c:Affective Therapies Reader}}",
        "Its **origin** then runs through **Virginia Axline's** non-directive play therapy in 1947 and her book Dibs: In Search of Self. It moves on to the Guerneys' filial therapy, Landreth's child-centred play therapy and Schaefer's founding of the Association for Play Therapy. {{c:Axline}}",
        "The first basis is that **play is the child's natural language**. Children cannot easily put abstract thoughts and feelings into words, so they express them through symbolic play with toys. {{c:Axline}}",
        "A second basis is the [[therapeutic powers of play|the specific ways play actually heals, such as helping a child communicate, release feelings and build relationships]], which Schaefer named as communication, emotional wellness, stronger relationships and increased personal strengths. {{c:Affective Therapies Reader}}",
        "A third basis is that play therapy draws on **several theories**: psychoanalytic, humanistic and child-centred, attachment, cognitive-behavioural and developmental. So it can be directive or non-directive. {{c:Affective Therapies Reader}}",
        "A fourth basis is **Axline's eight principles**. These include building a warm relationship, accepting the child exactly as they are, reflecting feelings, and respecting the child's own ability to solve problems. {{c:Axline}}",
        "In **formulation**, children's problems are read as **blocked emotions**, trauma, insecure attachment or unmet needs. Play lets the child express and master these safely, at a symbolic distance from the real event. For example, a scared child may make dolls act out a frightening scene. {{c:Affective Therapies Reader}}",
        "The **goals** are **emotional expression and regulation**, a better self-concept and self-esteem, social skills and coping, the resolution of trauma or conflict, and a stronger parent-child attachment. {{c:Affective Therapies Reader}}",
        "A core **procedure** is the **playroom and its toys**, sorted into three kinds: real-life toys like a doll house, acting-out and aggressive toys like puppets and a bop bag, and creative media like sand, clay and paint. {{c:Affective Therapies Reader}}",
        "Non-directive **techniques** include **tracking** the child's play aloud, reflecting feelings, returning responsibility with phrases like 'you can decide', and encouragement that follows rather than steers the child. {{c:Affective Therapies Reader}}",
        "Limit-setting is a key **technique**, often taught as [[ACT|a three-step way to set a limit: Acknowledge the feeling, Communicate the limit, Target an acceptable alternative]], which keeps the playroom safe without shaming the child. {{c:Affective Therapies Reader}}",
        "Structured **techniques** include **puppets, sandplay, drawing and storytelling**, Gardner's mutual storytelling, bibliotherapy, role play, therapeutic games, and cognitive-behavioural play that uses modelling and exposure. {{c:Affective Therapies Reader}}",
        "Parent-mediated **techniques** are distinctive. In **filial therapy and child-parent relationship therapy** the therapist trains parents to run play sessions at home. Theraplay builds structure, engagement, nurture and challenge. {{c:Affective Therapies Reader}}",
        "In **stages and process** the work moves from assessment, parent consultation and consent, through **relationship building** as the child explores the room and tests the therapist, into the appearance of themes in the play. {{c:Affective Therapies Reader}}",
        "The later **process** is a **working-through phase** of aggression, regression, limit-testing and dramatic play toward mastery. Then comes a termination that the child is prepared for early, usually across about twelve to twenty weekly sessions. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **better emotional expression, behaviour and self-esteem**, improved social skills and a stronger parent-child relationship, especially when parents take an active part. {{c:Affective Therapies Reader}}",
        "**Indications** are children of about three to twelve years with anxiety, fears, trauma and abuse, bereavement, parental divorce, adjustment and behaviour problems, and selective mutism. **Adolescents** may need talk, art or drama instead. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Bratton and colleagues'** 2005 meta-analysis of ninety-three studies reported a large mean effect, larger when parents took part. Later child-centred play therapy reviews show moderate to large effects, though critics note methodological weaknesses. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Psychodrama is J. L. Moreno's action-based group therapy. One group member, the protagonist, acts out their conflict instead of just talking about it. A director, helpers playing other people, an audience and a stage make this possible. The work restores spontaneity and new roles through three phases: warm-up, action and sharing."},
    {"t": "fig", "key": "psychodrama_phases", "caption": "The five instruments and three phases of psychodrama: warm-up, action and sharing on the stage."},
    {"t": "box", "kind": "def", "title": "Characterisations by classic authors", "items": [
        "{{THEORY}} **Jacob L. Moreno** created psychodrama out of his Theatre of Spontaneity. He coined the term **group psychotherapy** and developed sociometry, a way of measuring the relationships within a group. {{c:Moreno}}",
        "{{THEORY}} Its foundation is that [[spontaneity|the ability to respond freshly and well to a situation, old or new, instead of on autopilot]] and creativity are the basis of mental health. So trouble is a loss of spontaneity and getting stuck in rigid roles. {{c:Moreno}}",
        "{{THEORY}} Moreno described **five instruments**: the protagonist, the director, the auxiliary egos, the audience and the stage. Together they make the enactment possible. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is to **act out the problem**, rather than just talk about it. Rather than describe a painful argument with a parent, the client (the protagonist) stages it on the spot, with group members stepping in to play the other people, so the conflict is lived again in the room.",
        "The guiding belief is that health is **spontaneity** \u2014 being able to respond freshly and flexibly \u2014 and that distress comes from getting stuck in rigid, narrow roles. By replaying scenes and trying new responses safely, the person recovers spontaneity and widens their range of roles. Doing it, not only saying it, is what makes the difference.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The change engine is **enactment plus a fresh viewpoint**. Acting a scene out brings the real emotion up and releases it (catharsis), far more fully than describing it would. Then techniques like **role reversal** \u2014 where the protagonist steps into the other person's shoes and plays them \u2014 let the client literally feel the situation from the other side, which produces sudden empathy and insight.",
        "Because the stage is 'surplus reality', the client can also replay what never happened \u2014 say what was never said, meet someone who has died, rehearse a feared future \u2014 and so finish unfinished business and practise new roles safely. The closing **sharing**, where others tell their own similar stories, dissolves the shame of being alone with the problem and settles the person back into real life.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, goals, technique, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, psychodrama was created by **Jacob L. Moreno** in Vienna. His Theatre of Spontaneity in 1921 led, after he moved to the United States, to the Beacon sanatorium and a systematic method of action-based group therapy. {{c:Moreno}}",
        "Its **origin** also credits **Zerka Moreno** as a key developer. The method went on to influence Gestalt therapy, drama therapy, role play and the behavioural rehearsal used in CBT. {{c:Affective Therapies Reader}}",
        "The first basis is that **spontaneity and creativity** are the foundations of mental health. So distress is understood as lost spontaneity, rigid role patterns and unfinished action. {{c:Moreno}}",
        "A second basis is **role theory**, covering role-taking, role-playing and role-creating. Role conflict and role rigidity are seen as central to psychological problems. {{c:Affective Therapies Reader}}",
        "A third basis is [[tele|a genuine two-way flow of feeling and understanding between people, unlike one-sided transference]]. Moreno used the word for the real mutual empathy that binds a group. {{c:Affective Therapies Reader}}",
        "A fourth basis is **catharsis, encounter and surplus reality**. Surplus reality lets a client act out what never happened, or what is feared or wished, all worked in the here-and-now. {{c:Affective Therapies Reader}}",
        "In **formulation**, distress is traced to a **limited or conflicting set of roles**, blocked spontaneity and unresolved relationships and events. Enactment brings these into the open to be experienced and changed. {{c:Affective Therapies Reader}}",
        "The **goals** are **catharsis and release**, insight and a wider range of roles, renewed spontaneity, better interpersonal skills, the resolution of unfinished business, and group cohesion and support. {{c:Affective Therapies Reader}}",
        "A defining **procedure** is the use of the **five instruments**: the protagonist whose issue is explored, the director who guides, the auxiliary egos who play significant others, the audience who witness, and the stage. {{c:Affective Therapies Reader}}",
        "The most powerful **technique** is [[role reversal|the protagonist swaps and plays the other person, so as to see the situation through their eyes]]. It builds empathy and insight by letting the protagonist feel the other's position. {{c:Affective Therapies Reader}}",
        "A second key **technique** is **doubling**, where a helper voices the protagonist's unspoken feelings. There is also **mirroring**, where the protagonist watches a helper play them from the outside, like seeing themselves on film. {{c:Affective Therapies Reader}}",
        "Further **techniques** include **soliloquy, the empty chair and future projection**, plus concretisation, sculpting, the social atom and the magic shop, which make abstract feelings concrete and visible. {{c:Affective Therapies Reader}}",
        "A related **technique** is **sociodrama**, which turns the same action methods onto shared group or social issues, rather than one person's private conflict. {{c:Affective Therapies Reader}}",
        "In **stages and process**, the work opens with a **warm-up**. The director builds trust and a protagonist emerges, so the group is ready for action. {{c:Affective Therapies Reader}}",
        "The **action phase** then sets and acts out the scene, often moving from the present into the past. It uses role reversal, doubling and mirroring to reach a peak of emotion and new solutions before the scene is closed. {{c:Affective Therapies Reader}}",
        "The closing **process** is **sharing and integration**. The audience offer their own similar experiences, rather than analysis or advice. This reduces the protagonist's sense of isolation and helps them re-enter daily life. {{c:Affective Therapies Reader}}",
        "The **outcome** sought is **emotional release and insight**, better relationships and role flexibility, and the support of the group, all reached through doing rather than only talking. {{c:Affective Therapies Reader}}",
        "**Indications** are group settings for trauma, grief, addictions, family and relationship issues, adolescents, social-skills training and psychiatric inpatients. Acute psychosis, or a very fragile client without containment, is a **caution** needing a skilled director. {{c:Affective Therapies Reader}}",
        "In **research and current status**, **Kipper and Ritchie's 2003 meta-analysis** reported a large effect, with role reversal and doubling the most effective techniques. Later reviews show moderate positive evidence, though rigorous randomised trials remain few. {{c:Affective Therapies Reader}}",
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
    {"t": "lead", "text": "Oriental approaches are Asian mind-body practices used for mental health. They include yoga, meditation, shavasana, pranic healing, reiki and tai chi. They all work with attention, breath and relaxation. But their evidence varies a lot: genuine support for mindfulness, yoga and tai chi, down to weak and unclear findings for energy healing."},
    {"t": "fig", "key": "oriental_map", "caption": "A map of the named oriental approaches grouped by evidence, from mind-body practices with support to energy healing with weak evidence."},
    {"t": "box", "kind": "def", "title": "Characterisations by professional bodies and traditions", "items": [
        "{{THEORY}} **Patanjali's Yoga Sutras** describe [[yoga|the joining of body, mind and spirit through an eight-part discipline]] as calming the restless chatter of the mind. This is reached through the eight limbs, from ethical rules to deep absorption. {{c:Affective Therapies Reader}}",
        "{{GUIDE}} Clinical **mindfulness meditation** is described as non-judgemental, present-moment awareness. Kabat-Zinn adapted it as Mindfulness-Based Stress Reduction. It sits within the World Health Organization's wider promotion of self-care for mental health. {{c:WHO mhGAP}}",
        "{{THEORY}} **Energy-healing** practices like pranic healing and reiki claim to clear or channel a life energy, prana or ki. Their own traditions describe it this way, but there is **no credible scientific evidence** that such an energy exists. {{c:Affective Therapies Reader}}",
    ]},
    {"t": "box", "kind": "info", "title": "Main focus (the central idea)", "items": [
        "The **main focus** is to settle the mind through the **body, breath and attention**. Instead of talking a problem out, these practices quiet the whole nervous system \u2014 slow breathing, gentle movement, still postures and focused awareness \u2014 so the body shifts from 'alarm' mode into 'rest' mode, and the restless chatter of the mind calms down.",
        "They are **mind-body** methods: change the body (breath, posture, movement) and attention, and the mind follows. The group is honest about range and evidence, though. Mindfulness, yoga and tai chi have a **real evidence base**; the energy-healing claims of reiki and pranic healing have **weak or inconclusive evidence**, and must never replace proper medical or psychological care.",
    ]},
    {"t": "h3", "text": "How it actually works (the mechanism of change)"},
    {"t": "ul", "items": [
        "The main change engine is **better self-regulation of the nervous system**. Slow breathing, relaxation and steady movement switch on the body's calming (parasympathetic) side \u2014 heart rate, blood pressure, muscle tension and the stress hormone cortisol all drop \u2014 which is the body's natural opposite of the fight-or-flight stress response.",
        "Meditation adds a mental skill called [[decentring|learning to watch a thought as just a passing mental event, not a fact or a command]]: by watching worries drift past like clouds instead of being swept along by them, the person ruminates less and accepts more. For the energy-healing practices, any benefit is best explained by this same relaxation, focused attention and positive expectancy \u2014 not by a special energy, for which there is no credible proof.",
    ]},
    {"t": "h3", "text": "Exam points (origin, basis, formulation, techniques, process, outcome, indications, research)"},
    {"t": "ul", "items": [
        "In **origin**, these approaches come from **Asian philosophical and healing traditions**, chiefly Indian, Buddhist, Taoist and Japanese. Many have since been adapted into structured, sometimes evidence-based, interventions. {{c:Affective Therapies Reader}}",
        "**Yoga** comes from ancient India and was systematised by **Patanjali** in the Yoga Sutras as Ashtanga, or eight-limbed yoga. Later came Hatha yoga and types such as Raja, Karma, Bhakti and Jnana yoga. {{c:Affective Therapies Reader}}",
        "The basis of **yoga** is [[union|the joining of body, mind and spirit, which is the literal meaning of the word yoga]]. It is reached through the eight limbs: ethical restraints, observances, postures, breath control, sense withdrawal, concentration, meditation and absorption. {{c:Affective Therapies Reader}}",
        "In therapy, **yoga techniques** are gentle **asanas, pranayama breathing and relaxation**. Well-known studies at NIMHANS and meta-analyses report real benefit as an add-on for mild to moderate depression and anxiety. {{c:Affective Therapies Reader}}",
        "**Meditation** comes from Hindu, Buddhist and Taoist contemplative traditions. It was adapted clinically as **Mindfulness-Based Stress Reduction** by Jon Kabat-Zinn in 1979, and as Mindfulness-Based Cognitive Therapy by Segal, Williams and Teasdale. {{c:Affective Therapies Reader}}",
        "The basis of **meditation** is **training attention and decentring**. Decentring means learning to watch your thoughts as passing mental events, not facts, like watching clouds drift by. This lowers rumination and raises acceptance and emotion control. {{c:Affective Therapies Reader}}",
        "The evidence for **meditation** is genuine. Mindfulness-based cognitive therapy is **recommended for preventing relapse** in recurrent depression. A major review by Goyal and colleagues in 2014 found moderate evidence for anxiety, depression and pain. {{c:Affective Therapies Reader}}",
        "**Shavasana**, the corpse pose, is the final relaxation posture of a yoga session. The related [[Yoga Nidra|a guided deep relaxation, sometimes called 'yogic sleep', where you stay aware while the body fully rests]] was popularised by Swami Satyananda Saraswati. {{c:Affective Therapies Reader}}",
        "The basis of **shavasana** is a **relaxation response** like Benson's. It lowers sympathetic arousal, heart rate, blood pressure and muscle tension, much like Jacobson's progressive muscle relaxation. Small trials show less anxiety and better sleep. {{c:Affective Therapies Reader}}",
        "**Pranic healing** was developed by **Master Choa Kok Sui** in 1987. It holds that disease comes from depleted or congested prana in an energy body, treated by scanning, cleansing and energising the field, usually without touch. {{c:Affective Therapies Reader}}",
        "The evidence for **pranic healing** is **very limited and poorly controlled**. There is no credible evidence for a detectable energy field. Any benefit is probably relaxation, attention and expectancy, and it must not replace evidence-based treatment. {{c:Affective Therapies Reader}}",
        "**Reiki** comes from **Mikao Usui** in Japan in 1922. The practitioner channels a universal life energy through the hands to the client, who lies clothed while set hand positions are used. {{c:Affective Therapies Reader}}",
        "The evidence for **reiki** is **weak and at high risk of bias**. Reviews such as the Cochrane review for anxiety and depression find insufficient evidence. So it is best seen as a complementary comfort measure that should not delay treatment. {{c:Affective Therapies Reader}}",
        "**Tai chi** comes from seventeenth-century China, from Taoist philosophy and martial arts. It is often called **meditation in motion**, using slow, flowing, coordinated movements with deep breathing and focused attention. {{c:Affective Therapies Reader}}",
        "The evidence for **tai chi** is **reasonably good**. Meta-analyses report less stress, anxiety and depression, benefit as an add-on in late-life depression, and value for fall prevention, chronic pain and cardiac rehabilitation. {{c:Affective Therapies Reader}}",
        "Other related approaches, offered under 'etc.', include **Morita and Naikan therapy** from Japan and qigong. Morita teaches accepting your feelings and doing what needs to be done; Naikan uses structured introspection to build gratitude. {{c:Affective Therapies Reader}}",
        "The common **basis and mechanism** across these practices is better **self-regulation**: raised parasympathetic and vagal tone, lowered cortisol and arousal, and better attention and emotion control. This explains much of their benefit. {{c:Affective Therapies Reader}}",
        "The honest **outcome and indications** are clear. Mindfulness, yoga and tai chi have a **real evidence base** as add-ons for stress, anxiety, depression and chronic pain. Reiki and pranic healing have **weak or inconclusive** evidence and carry the risk of replacing proper care. {{c:Affective Therapies Reader}}",
        "In **research and current status**, these approaches fit Indian and Asian clients' cultural beliefs and are low in cost and side effects. But their evidence is of **variable quality**: strongest for mindfulness-based therapies and weakest for energy-healing claims that remain untestable. {{c:Affective Therapies Reader}}",
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
