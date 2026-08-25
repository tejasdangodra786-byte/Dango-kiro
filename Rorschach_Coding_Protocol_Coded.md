# RORSCHACH INKBLOT TEST — RESPONSE CODING (SEQUENCE OF SCORES)
### Exner Comprehensive System (CS)

> **Coding reference:** Coding rules applied strictly per the repository's `Rorschach_Coding_Complete_Guide_Exner.html` (Exner CS). Form Quality (FQ) and Z-score values are assigned using the standard published Comprehensive System norms (Exner *Table A* Form Quality tables and the card-by-card Z-value table).

---

## ⚠️ IMPORTANT SOURCE NOTE (please read before relying on FQ)

Form Quality in the Comprehensive System is **not** determined by "looking at the blot" — it is a **table lookup** against Exner's normative **Table A** (which lists, for every location on every card, whether a given content is `o`, `u`, or `–`). The same rule applies to exact **Z-score values** (the ZW / ZAdj / ZDist / ZSpace table).

I checked the two candidate reference files in this repository:

| File | What it actually contains | Has FQ Table A / Z table? |
|---|---|---|
| `Rorschach_Manual_Exner_OCR_Complete.txt` | Only **Chapter 13+ (Interpretation)** of Exner Vol. 1 — noisy OCR of interpretive prose | ❌ No |
| `A-Rorschach-Workbook-for-the-Comprehensive-System.pdf` | **100% scanned images** (1495 image objects, 0 embedded fonts, 21 MB) — no extractable text | ❌ Not machine-readable offline |

The sandbox has **no internet access** (INTEGRATIONS_ONLY) and no offline OCR engine, so I could not extract Table A from the scanned workbook. **Therefore the FQ codes below are assigned from the standard, widely-published CS Table A values, not from a lookup verified against your specific workbook.**

➡️ **Action for you:** the FQ column is marked with a confidence flag. Any row flagged **`[verify]`** sits close to an `o/u` or `u/–` boundary and should be confirmed against the workbook's Table A page for that exact location before the Structural Summary is finalized. If you can give me OCR/photos of the relevant Table A pages, I will lock every FQ down with certainty.

---

## Sequence of Scores

Format: **Card · Resp# · Loc + DQ · Determinant(s) + FQ · (2) · Content(s) · P · Z · Special Scores**

| # | Card | Loc & DQ | Determinant(s) & FQ | (2) | Content | P | Z | Special Scores |
|---|------|----------|---------------------|-----|---------|---|---|----------------|
| 1 | I | D4 o | Fo | | (H), Cg | | | |
| 2 | I | Dd99 + | FMp– | | A | | 4.0 | INC1 |
| 3 | I | Wo | FMp o | | A | | 1.0 | |
| 4 | II | Dd99 o | Fo | (2) | Ad | | | |
| 5 | II | DS5 o | Fo | | Art (Hh) | | 4.5 | |
| 6 | II | D2 o | FCo | (2) | A | | | |
| 7 | III | D9 o | Mp– | (2) | (H) | | | MOR, PHR |
| 8 | III | D3 o | CFu | | Bl, Sc | | | MOR |
| 9 | III | D2 v | C | (2) | Bl | | | MOR |
| 10 | IV | Wo | Mp o | | (H), Cg | P | 2.0 | GHR |
| 11 | V | Wo | FMa.Fo | | A, Bt | P | 1.0 | INC1 |
| 12 | VI | Wv | Y– | | Id | | | |
| 13 | VI | Wo | Mp o | | (H) | | 2.5 | PHR |
| 14 | VII | W+ | FMp o | (2) | A, Hd | P | 2.5 | |
| 15 | VIII | W+ | FMa.CFo | | A, Bt | | 4.5 | INC1 |
| 16 | VIII | Dd33 o | Fo | (2) | A | | | |
| 17 | VIII | D7 o | CFo | | Bt | | | |
| 18 | IX | D6 o | mp– | | An, Sx | | | MOR |
| 19 | IX | Wv | C | | Art | | | |
| 20 | X | D1 o | FMa o | (2) | A | P | | |
| 21 | X | D o | Fo | | A | | | PSV |
| 22 | X | DdS22 o | F– | | Cg, Sx | | | |
| 23 | X | D11 o | Fo | | Ad | | | |

---

## Response-by-Response Coding Rationale

Each response is coded across the seven CS categories in sequence: **Location → DQ → Determinant(s) → FQ → Pair → Content → Popular → Special Scores.**

---

### CARD I

**Response 1 — "A doll; I can see its clothes, the frock."**
- **Location:** D4 (the central figure). *A single central figure, not the whole blot → D, not W.*
- **DQ:** `o` — one object with a definite form demand (a doll has specific shape); no meaningful integration of separate areas.
- **Determinant:** `F` — only shape is invoked ("clothes… frock… that is why it looks like a doll"). No movement, colour, or shading articulated.
- **FQ:** `o` — a human-form figure in the central D of Card I is a conventional, well-fitting percept.
- **Pair:** none (one doll).
- **Content:** `(H), Cg` — a doll is a fictional/representational human figure `(H)`; the frock/clothes = `Cg`.
- **P:** no.
- **Z:** none (single D, no organization).
- **Special:** none.
- **CODE:** `Do Fo (H),Cg`

> *Note:* A "doll" is coded `(H)` (a human representation that is not a real person). Since human content is present, a GHR/PHR determination applies — but only responses with **M** or an **FQ that triggers PHR** typically carry the tag in the sequence; a pure-F `(H)` with FQo and no special scores resolves to **GHR** by the algorithm. Tag added in the summary count.

**Response 2 — "Three things joined together making one animal; I can see hands and legs."**
- **Location:** Dd99 — an idiosyncratic combination of areas ("three things attached") that is not a listed D → unusual detail. **[verify location on the location sheet]**
- **DQ:** `+` — separate areas are actively integrated into one related percept ("three things joined… making one animal").
- **Determinant:** `FMp` — an animal to which movement/posture is attributed; the fusion of parts implies a static animal → passive animal movement. *(If, on the sheet, no motion was truly conveyed, step down to `F`.)* **[verify]**
- **FQ:** `–` — a single animal built from three unrelated joined areas does not fit the contours; likely minus. **[verify against Table A for the exact Dd used]**
- **Content:** `A`.
- **Z:** `4.0` — Card I ZDist/adjacent-integration value for combining separate areas into one W-like organization. **[verify Z type: ZAdj vs ZDist]**
- **Special:** `INC1` — implausible single object ("three things" fused into one animal with hands *and* legs) = incongruous combination, mild.
- **CODE:** `Dd99+ FMp– A 4.0 INC1`

**Response 3 — "A bear; it looks like it is eating something. Hands, legs, nose; only one of it."**
- **Location:** W ("only one of it" = the whole figure, not a symmetrical pair).
- **DQ:** `o` — a whole animal with named parts but no integration of *separate* objects into a scene; the eating is attributed to the single figure.
- **Determinant:** `FMp` — an animal in a species-typical activity (eating). Coded passive (eating/consuming, low exertion). *(Active if the sheet conveys vigorous action.)*
- **FQ:** `o` — a bear using the whole of Card I is a reasonable, seeable fit. **[verify — bear as W on I can fall o/u depending on articulation]**
- **Content:** `A`.
- **P:** no (Card I Popular is bat/butterfly, not bear).
- **Z:** `1.0` — ZW for a whole response on Card I.
- **CODE:** `Wo FMp o A 1.0`

---

### CARD II

**Response 1 — "Dogs; two faces looking opposite ways; two legs each; no tails."**
- **Location:** Dd99 (both sides) — the described dog "faces" are not the standard D animal; treated as an unusual detail on each side. **[verify]**
- **DQ:** `o` — animals with named parts, but no synthesis into an interactive scene (they simply face away).
- **Determinant:** `F` — only shape/contour is used (faces, legs described by form; no motion, colour, or shading given for the dogs).
- **FQ:** `o` — dogs/animal heads on the black D of Card II are commonly seeable. **[verify for the exact Dd]**
- **Pair:** `(2)` — two identical dogs on the bilateral symmetry.
- **Content:** `Ad` — "two faces… two legs" describes animal *parts*, not whole animals → animal detail. *(Code `A` instead if whole dogs were clearly seen.)* **[verify]**
- **Z:** none if a single symmetrical detail; assign only if separate areas were organized.
- **CODE:** `Dd99o F o (2) Ad`

**Response 2 — "A temple bell; attached from above, with the part you ring it from."**
- **Location:** DS5 — the central **white space** (the light central column) used as the bell body → S incorporated with a D. **[verify D/Dd number]**
- **DQ:** `o` — one object (bell) with articulated related parts (top attachment + clapper) → arguably `+`; coded `o` conservatively unless the parts are clearly separate integrated areas. **[verify + vs o]**
- **Determinant:** `F` — shape only.
- **FQ:** `o` — a bell in the central space of Card II is a recognized space percept. **[verify]**
- **Content:** `Art` (a temple bell as an ornamental/ceremonial object) or `Hh`; primary `Art`.
- **Z:** `4.5` — Card II **ZSpace** value (integrating white space with the figure).
- **Special:** none.
- **CODE:** `DS5o Fo Art 4.5`

**Response 3 — "Two parrots; legs, face, tail, eyes; red; parrots are like this."**
- **Location:** D2 (the upper/lateral red areas), both sides.
- **DQ:** `o` — whole animals, named parts, no scene integration.
- **Determinant:** `FC` — form is primary (parrot shape, named parts) with chromatic colour used appropriately ("red… parrots are like this") → colour secondary to form.
- **FQ:** `o` — the upper red D of Card II as a bird is a conventional, seeable fit. **[verify — often u]**
- **Pair:** `(2)`.
- **Content:** `A`.
- **P:** no.
- **Z:** none (single symmetrical D).
- **Special:** none. *(No ALOG: "parrots are like this" is a general description, not strained positional/color logic used as sole proof.)*
- **CODE:** `Do FCo (2) A`

---

### CARD III

**Response 1 — "A man cut from the middle; one hand, one leg; looks like a ghost."**
- **Location:** D9 (both sides) — the lateral human figures.
- **DQ:** `o` — a human figure with parts, no integrated interaction between two figures.
- **Determinant:** `Mp` — a human figure with attributed posture/state ("cut… looks like a ghost"), passive.
- **FQ:** `–` — "a man cut from the middle" with only one hand/one leg imposes a damaged, ill-fitting form on the figure → minus. **[verify against Table A]**
- **Pair:** `(2)` — described bilaterally (both sides).
- **Content:** `(H)` — "ghost" = fictional human.
- **Z:** none (single D per side).
- **Special:** `MOR` — "cut from the middle," damaged/dismembered = morbid. **PHR** — human-content response with FQ– → poor human representation.
- **CODE:** `Do Mp– (2) (H) MOR, PHR`

**Response 2 — "A blade; blood; the hand is cut by the blade, so there's blood."**
- **Location:** D3 (the central red detail).
- **DQ:** `o` for the blade (specific object). *(The blood component is formless — see the reasoning; the dominant object here is the blade.)*
- **Determinant:** `CF` — colour drives the blood ("blood… red"), with some object form (blade) present; colour-dominant overall → CF. *(The blade alone would be F; the blood is colour-based; the combined percept is colour-dominant → CF.)* **[verify FC vs CF]**
- **FQ:** `u` — a blade in the central red D is uncommon but seeable.
- **Content:** `Bl, Sc` — blood; a blade = a man-made implement (`Sc`/`Id`). Primary `Bl`.
- **Z:** none.
- **Special:** `MOR` — cutting/blood = damage. *(Consider `ALOG` only if the color-as-proof reasoning was offered spontaneously and rigidly: "it must be blood because it is red." As worded ("blood is red in colour") it is descriptive, so ALOG is **not** coded — flag if the verbatim shows rigid causal logic.)* **[verify ALOG]**
- **CODE:** `Do CFu Bl,Sc MOR`

**Response 3 — "Stains of blood; someone splattered the blood."**
- **Location:** D2 (the outer red areas), both sides.
- **DQ:** `v` — "stains… splattered" has **no form demand** (amorphous) → vague.
- **Determinant:** `C` — pure colour; the blood is defined solely by redness with no shape → Pure C, no FQ. *(Because "splattered" attributes movement, a case exists for `mp` as a blend; if the spreading motion was articulated, code `C.mp`. As primarily a colour-based formless percept, Pure C is the anchor.)* **[verify blend with m]**
- **FQ:** none (Pure C carries no form quality).
- **Pair:** `(2)` (bilateral red).
- **Content:** `Bl`.
- **Z:** none.
- **Special:** `MOR` — blood/splatter = damage/dysphoria.
- **CODE:** `Dv C (2) Bl MOR`

---

### CARD IV

**Response 1 — "The whole thing is a witch; hands, legs, face; black spots in between."**
- **Location:** W.
- **DQ:** `o` — a single whole figure with named parts; not an integration of separate objects into a relationship.
- **Determinant:** `Mp` — a human(-like) figure with implied posture/stance (standing figure), passive. *(If no posture/tension was conveyed, step down to `F`.)* **[verify M vs F]** The "black spots" alone would suggest achromatic colour/shading, but they are mentioned as incidental markings, not as a colour/shading determinant driving the percept.
- **FQ:** `o` — a large human-like figure using the whole of Card IV is the conventional percept.
- **Content:** `(H), Cg` — witch = fictional human; if clothing implied. Primary `(H)`.
- **P:** **P** — Card IV Popular is a human/human-like figure (giant, monster, person) on W/D7.
- **Z:** `2.0` — ZW for Card IV.
- **Special:** **GHR** — human-content, FQo, no cognitive special scores, no AG/MOR → good human representation.
- **CODE:** `Wo Mp o (H),Cg P 2.0 GHR`

---

### CARD V

**Response 1 (+ Inquiry) — "A butterfly / human-butterfly on a tree… complete butterfly, antennae, face, flying, legs."**
- **Location:** W (the inquiry clarifies the whole blot as one complete butterfly).
- **DQ:** `o` — a single whole object (butterfly) with named parts; the "tree/sitting" idea drops out at inquiry in favor of one complete flying butterfly, so no two-object integration survives → `o`. **[verify + vs o if the tree was retained as a separate integrated area]**
- **Determinant:** `FMa.F` → simplify to **`FMa`** — an animal (butterfly) in species-natural motion (flying), active. Form is subsumed within FM. *(The naming of antennae/face/legs is form articulation, not a separate F determinant.)*
- **FQ:** `o` — butterfly on the whole of Card V is a textbook ordinary fit.
- **Content:** `A` (butterfly). *(The dropped "tree" would add `Bt`; retained here only if the sheet keeps it — otherwise omit.)*
- **P:** **P** — Card V Popular is bat/butterfly on W.
- **Z:** `1.0` — ZW for Card V.
- **Special:** `INC1` — the initial "human butterfly" fuses human + butterfly attributes onto one object = incongruous combination, mild. *(If the final inquiry percept is a clean "complete butterfly" with the human idea fully retracted, INC1 may be dropped — code the response as finally clarified.)* **[verify — depends on whether "human butterfly" was retained]**
- **CODE:** `Wo FMa o A P 1.0 (INC1 — see note)`

---

### CARD VI

**Response 1 — "Black spots all over the whole picture."**
- **Location:** W.
- **DQ:** `v` — "black spots all over" has no form demand (amorphous) → vague.
- **Determinant:** `Y` or `C'` — "black spots" references darkness. Per the CS step-down rule, when the intent between achromatic colour and diffuse shading is unclear, **code `Y`** (diffuse shading), no form. **[verify C' vs Y from verbatim: "black" as colour → C'F/C'; "shadowy/dark patches" → Y]**
- **FQ:** none (formless, `Y`).
- **Content:** `Id` — undifferentiated "spots"/ink markings have no standard content category → idiographic. *(If read as darkness/nature, `Na` could apply.)*
- **Z:** none (no organization of a vague whole beyond the blot itself; `Wv` with no form does not earn ZW).
- **Special:** none.
- **CODE:** `Wv Y Id`

**Response 2 — "A person standing straight; like a witch; face, hands, legs."**
- **Location:** W.
- **DQ:** `o` — a single whole figure with named parts, standing.
- **Determinant:** `Mp` — human(-like) figure with attributed posture ("standing straight"), passive.
- **FQ:** `o` — an upright human-like figure using the whole/central axis of Card VI is a seeable fit (the top "totem/figure" percept). **[verify o vs u]**
- **Content:** `(H)` — "witch" = fictional human. *(If read as a real "person," code `H`.)* **[verify H vs (H)]**
- **Z:** `2.5` — ZW for Card VI.
- **Special:** **PHR/GHR** — depends on final FQ and content: with FQo and no special scores → **GHR**; the "witch" fictional label alone does not force PHR. Coded **GHR** here. **[verify]**
- **CODE:** `Wo Mp o (H) 2.5 GHR`

---

### CARD VII

**Response 1 — "Monkeys on both sides, joined at the legs/tail; they have braids in their hair."**
- **Location:** W (both lateral figures + the connecting lower area joined together).
- **DQ:** `+` — two figures actively integrated (joined, related, with the connecting area) → synthesized.
- **Determinant:** `FMp` — animals in posture/attitude (attached, sitting), passive. *(Active if the monkeys were conveyed as moving/reaching.)* **[verify a vs p]**
- **FQ:** `o` — the Card VII figures as human/animal profiles are the classic percept; monkeys with the described parts are seeable. **[verify o vs u]**
- **Pair:** `(2)`.
- **Content:** `A, Hd` — monkeys = animal; "braids in the hair" introduces a human-detail element `Hd`. Primary `A`. *(If braids read as the animals' own fur styling, `Hd` may be dropped.)* **[verify]**
- **P:** **P** — Card VII Popular is human head/face (D1/D9); the upper-third figures here overlap the Popular area. Coded P as the figures include that region. **[verify — P requires the human-head content specifically; if the percept is purely "monkeys," P may not apply]**
- **Z:** `2.5` — Card VII ZW / adjacent integration for joining the halves. **[verify Z type]**
- **Special:** none. *(Consider `COP` only if a clearly positive interaction was described; "attached" is not cooperative action → no COP.)*
- **CODE:** `W+ FMp o (2) A,Hd 2.5`

---

### CARD VIII

**Response 1 — "Insects and a butterfly climbing a tree, sitting on a green leaf; butterfly has wings, tail, green antennae; the leaf is green."**
- **Location:** W (the animals + central/upper areas + the green lower area as leaf, integrated).
- **DQ:** `+` — multiple objects in a meaningful relationship (creatures climbing/sitting on a leaf on a tree) → synthesized.
- **Determinant:** `FMa.CF` — animals in natural motion (climbing), active `FMa`; plus chromatic colour used for the green leaf where colour is prominent and form is secondary for the leaf ("leaf is green… so it looks like a leaf") → `CF`. Blend ordered movement-then-colour.
- **FQ:** `o` — the lateral animals of Card VIII are the Popular animal figures and fit well; the integrated leaf/tree is seeable. **[verify overall + response FQ]**
- **Content:** `A, Bt` — insects/butterfly = animal; leaf/tree = botany.
- **P:** The Card VIII Popular is the **D1 lateral animal**. This response is coded **W+** integrating more than D1; the Popular applies specifically to the D1 animal seen as such. Because the animals here are folded into a larger W scene rather than reported as the standalone D1 four-legged animal, **P is not coded** for the W percept. **[verify — if the D1 animal was clearly identified, add P]**
- **Z:** `4.5` — Card VIII ZW/adjacent-integration value for the synthesized whole. **[verify Z type/value]**
- **Special:** `INC1` — "green antennae/green moustache" and the composite creature carry a mild incongruity (a butterfly with a "moustache"); mild incongruous combination. **[verify — may be dropped if "antennae" is a plain correction]**
- **CODE:** `W+ FMa.CFo A,Bt 4.5 INC1`

**Response 2 — "Spiders below on both sides."**
- **Location:** Dd33 (the lower lateral projections). **[verify Dd number for the exact area]**
- **DQ:** `o` — a discrete creature with form demand.
- **Determinant:** `F` — shape only (no motion/colour/shading articulated).
- **FQ:** `o` — small spider/insect in the lower detail of Card VIII is seeable. **[verify o vs u]**
- **Pair:** `(2)`.
- **Content:** `A`.
- **Z:** none.
- **Special:** none.
- **CODE:** `Dd33o F o (2) A`

**Response 3 — "A green flower."**
- **Location:** D7 (the upper-central grey/green area used as the flower). **[verify D number — the green is lower; if the green D was used, adjust]**
- **DQ:** `o` — a flower has a form demand. *(If described only as "green" with no shape → `v`.)* **[verify o vs v]**
- **Determinant:** `CF` — colour-dominant ("green flower"; the greenness drives it) with some flower form → CF.
- **FQ:** `o` — a flower/botany percept on the green D of Card VIII is seeable. **[verify]**
- **Content:** `Bt`.
- **Z:** none.
- **Special:** none.
- **CODE:** `Do CFo Bt`

---

### CARD IX

**Response 1 — "Eggs of children; like children being formed inside the stomach; four of them."**
- **Location:** D6 (the central/lower area) or the pink areas read as forming figures. **[verify location]**
- **DQ:** `o` — described objects (eggs/forming figures) with some form; not an integrated multi-object scene → `o`. **[verify o vs v — "eggs" can be vague]**
- **Determinant:** `mp` — a developmental/forming process ("being formed inside") attributes a passive, non-volitional inanimate-type change. *(If purely static shapes, step down to `F`.)* **[verify m vs F]**
- **FQ:** `–` — "children being formed inside the stomach / eggs of children" imposes a percept the contours do not support → minus. **[verify against Table A]**
- **Content:** `An, Sx` — internal formation/womb imagery = anatomy `An`; reproductive/gestational content = sex `Sx`. Primary `An`.
- **Z:** none (single area).
- **Special:** `MOR` — gestation-in-the-stomach damage/anomaly imagery carries a morbid, distorted-body quality. **[verify — MOR requires damage/decay; if purely developmental with no damage, MOR may be dropped]**
- **CODE:** `Do mp– An,Sx MOR`

**Response 2 — "Just stains; three kinds — green, red, and turmeric (yellow)."**
- **Location:** W (all three colour areas referenced together).
- **DQ:** `v` — "stains" has no form demand → vague.
- **Determinant:** `C` — pure chromatic colour; the percept is defined solely by the colours with no form → Pure C, no FQ.
- **FQ:** none (Pure C).
- **Content:** `Art` — coloured patches described as kinds of stains/paint → treated as `Art` (design/paint). *(Alternatively `Id`.)* **[verify Art vs Id]**
- **Z:** none.
- **Special:** none. *(No MOR — "stains" here are colour patches, not blood/damage.)*
- **CODE:** `Wv C Art`

---

### CARD X

**Response 1 — "Two spiders; their legs are in front."**
- **Location:** D1 (the lateral blue areas), both sides.
- **DQ:** `o` — a whole creature with named parts.
- **Determinant:** `FMa` — a spider in species-natural posture/motion (legs positioned/reaching), active. *(Passive if only static posture.)* **[verify a vs p]**
- **FQ:** `o` — spider/multi-legged creature on the blue D1 of Card X is the Popular, well-fitting percept.
- **Pair:** `(2)`.
- **Content:** `A`.
- **P:** **P** — Card X Popular is spider/crab on D1.
- **Z:** none (single symmetrical D).
- **Special:** none.
- **CODE:** `Do FMa o (2) A P`

**Response 2 — "An insect; it has a mouth and legs." (multiple D areas, each a separate insect)**
- **Location:** D (multiple separate D areas, each read as a separate insect). Code the location of the specific area; the repetition across areas is what matters. **[verify which D areas]**
- **DQ:** `o` — a creature with named parts.
- **Determinant:** `F` — shape only (mouth, legs by form). *(FM if motion was attributed.)* **[verify]**
- **FQ:** `o` — insects on the various D details of Card X are commonly seeable. **[verify per area]**
- **Content:** `A`.
- **Z:** none.
- **Special:** `PSV` — **within-card perseveration**: the same content ("insect") assigned to multiple different areas of the same card with the same determinant/DQ. **[verify PSV type — within-card]**
- **CODE:** `Do F o A PSV`

**Response 3 — "Clothes, and a chest — like a female's chest."**
- **Location:** DdS22 — uses the central white space (the light central area) plus adjacent detail. **[verify Dd/D + S number]**
- **DQ:** `o` — objects with some form (clothing/chest). *(Consider `+` if clothes and chest were integrated as one dressed figure.)* **[verify]**
- **Determinant:** `F` — shape only (no colour/shading/movement articulated).
- **FQ:** `–` — a female chest/clothing on the central space of Card X is a low-fit, idiosyncratic percept → likely minus. **[verify against Table A]**
- **Content:** `Cg, Sx` — clothes = clothing; chest/breast = sex. Primary `Cg`.
- **Z:** assign a **ZSpace** value if the white space was integrated with adjacent detail. **[verify Z]**
- **Special:** none.
- **CODE:** `DdSo F– Cg,Sx` *(add ZSpace if space integrated)*

**Response 4 — "A mouth; I can see teeth."**
- **Location:** D11 (a central detail read as a mouth). **[verify D number]**
- **DQ:** `o` — a mouth with teeth has a form demand.
- **Determinant:** `F` — shape only.
- **FQ:** `o` — a mouth/teeth on the relevant central D of Card X is seeable. **[verify o vs u]**
- **Content:** `Ad` or `Hd` — a mouth/teeth as an anatomical detail; if human → `Hd`, if animal/unspecified → `Ad`. Coded `Ad` pending clarification. **[verify Hd vs Ad]**
- **Z:** none.
- **Special:** none.
- **CODE:** `Do F o Ad`

---

## Frequency Tallies (from the coding above)

> These counts follow directly from the codes assigned above. **FQ-dependent figures (XA%, WDA%, X+%, X–%, Xu%) and the final Z summary are provisional** until the `[verify]` FQ/Z rows are confirmed against Table A. Everything else (locations, determinants, contents, Populars, special scores) is derived purely from the CS coding rules and is stable.

**R (total responses) = 23** *(this protocol has 23 codable responses across the 10 cards)*

### Location (W : D : Dd)
| Loc | Count | Responses |
|---|---|---|
| W | 8 | #3, #10, #11, #12, #13, #14, #15, #19 |
| D | 10 | #1, #6, #7, #8, #9, #17, #18, #20, #21, #23 |
| Dd | 5 | #2, #4, #5, #16, #22 |
| **S** (modifier only) | 2 | #5 (DS5), #22 (DdS22) |

**W : D : Dd = 8 : 10 : 5** *(S is a modifier on #5 and #22, not a separate location total)*

### Developmental Quality (DQ)
| DQ | Count | Responses |
|---|---|---|
| + | 3 | #2, #14, #15 |
| o | 17 | #1, #3, #4, #5, #6, #7, #8, #10, #11, #13, #16, #17, #18, #20, #21, #22, #23 |
| v/+ | 0 | — |
| v | 3 | #9, #12, #19 |

*(Total 3 + 17 + 3 = 23 ✓. Recount if any `[verify]` +/o/v flag changes.)*

### Determinants (blends counted by each component)
| Determinant | Count | Responses |
|---|---|---|
| M (human movement) | 3 | #7 (Mp), #10 (Mp), #13 (Mp) → **Ma : Mp = 0 : 3** |
| FM (animal movement) | 5 | #2 (FMp), #3 (FMp), #11 (FMa), #14 (FMp), #15 (FMa), #20 (FMa) → **FMa : FMp = 3 : 3** *(6 FM tokens — see note)* |
| m (inanimate movement) | 1 | #18 (mp) |
| Pure F | 8 | #1, #4, #5, #16, #21, #22, #23, + #2 only if stepped down (else 7) |
| FC | 1 | #6 |
| CF | 3 | #8, #17, and #15 (as the colour half of the FMa.CF blend) |
| C (pure) | 2 | #9, #19 |
| Y (diffuse shading) | 1 | #12 |
| C' | 0 | — (see #12 `[verify]` C' vs Y) |

> **FM note:** six FM *tokens* are listed but #2 is coded `FMp–` and could step down to pure `F` on `[verify]`; if it does, FM = 5 and Pure F = 8. As currently coded (FM kept on #2): FM = 6, Pure F = 7. **FMa : FMp = 3 : 3** either way for the a/p ratio.

**Active : Passive (a : p):** active = #11, #15, #20 (FMa) → 3; passive = #2, #3, #7, #10, #13, #14, #18 → 7. **a : p = 3 : 7** *(recheck the a/p `[verify]` flags on #7, #14, #20).*

### Form Quality (PROVISIONAL — pending Table A)
| FQ | Count | Responses |
|---|---|---|
| + | 0 | — |
| o | 15 | #1, #3, #4, #5, #6, #10, #11, #13, #14, #15, #16, #17, #20, #21, #23 |
| u | 1 | #8 |
| – | 4 | #2, #7, #18, #22 |
| none (formless — no FQ) | 3 | #9 (C), #12 (Y), #19 (C) |

*(Totals: o 15 + u 1 + – 4 + none 3 = **23 ✓**. #9 is Pure C → formless → carries no FQ, so it is NOT in the minus row.)*

> **Provisional FQ ratios** (form-bearing R = 20; adequate = 15 o + 1 u = 16): **X–% ≈ 4 / 20 = 0.20**, **XA% ≈ 16 / 20 = 0.80**, **X+% ≈ 15 / 20 = 0.75**, **Xu% ≈ 1 / 20 = 0.05**. *These MUST be recomputed after Table A verification — several `o` calls could shift to `u` or `–`.*

### Contents (each token counted)
| Content | Count | Responses |
|---|---|---|
| (H) | 4 | #1, #7, #10, #13 |
| H | 0 | — → **Pure H = 0** |
| Hd | 1 | #14 (braids) |
| A | 8 | #2, #3, #6, #11, #15, #16, #20, #21 |
| Ad | 2 | #4, #23 |
| An | 1 | #18 |
| Art | 3 | #5, #19, (#8 secondary if read as implement/design) |
| Bl | 2 | #8, #9 |
| Bt | 2 | #15, #17 |
| Cg | 3 | #1, #10, #22 |
| Sc | 1 | #8 |
| Sx | 2 | #18, #22 |
| Id | 1 | #12 |

**Human content total = 5** ( (H) = 4, Hd = 1 ); **Pure H = 0**; **H : (H)+Hd+(Hd) = 0 : 5**

### Populars (P)
**P = 4** — #10 (Card IV, human-like figure), #11 (Card V, butterfly), #14 (Card VII, head — *verify content covers the D1/D9 head area*), #20 (Card X, spider on D1). *Card VIII Popular is **not** counted for #15 because the lateral animal was folded into a larger W scene rather than reported as the standalone D1 animal.* **[verify #14, and #15 if the D1 animal was named separately]**

### Special Scores
| Score | Count | Responses | WSum6 contribution |
|---|---|---|---|
| INC1 | 3 | #2, #11, #15 | 3 × 2 = **6** |
| MOR | 4 | #7, #8, #9, #18 | — |
| PSV | 1 | #21 | — |
| GHR | 2 | #10, #13 | — |
| PHR | 1 | #7 | — |
| DV / DR / FAB / ALOG / CONTAM | 0 | — | 0 |

**Sum6 = 3** (all INC1) · **WSum6 = 6** · **Level-2 special scores = 0** · **MOR = 4** · **GHR : PHR = 2 : 1**

---

## What is solid vs. what needs your workbook

**Solid (rule-based, no table needed):**
- All **Location** codes' W/D/Dd/S *category* logic, **DQ** logic, **Determinant** identification and active/passive, **Content** codes, **Pair (2)**, **Special Scores** (INC1, MOR, PSV, GHR/PHR), the **blend structures**, and the **Popular** logic.

**Needs Table A / Z-table confirmation (flagged `[verify]`):**
1. Every **FQ** that sits near an `o/u` or `u/–` boundary (esp. #2, #5, #6, #7, #8, #9, #11-note, #13, #18, #22).
2. Exact **D vs Dd numbers** and the **S** location numbers (need the location sheet the client marked).
3. Exact **Z values and Z type** (ZW / ZAdj / ZDist / ZSpace) for #2, #5, #7, #14, #15, #22, and the final **Zf / Zsum / Zest / Zd**.
4. A few **content / a-p / M-vs-F / C'-vs-Y** calls tied to the exact verbatim + inquiry (#2, #3-blade, #6, #9, #12, #14, #20).

Once you provide either (a) clear photos/OCR of the workbook's **Table A** pages and the **Z-score table**, or (b) the marked **location sheet**, I will lock down every FQ and Z value and then compute the full, defensible **Structural Summary** (EB, EA, es, D, Adj D, Lambda, XA%/WDA%/X–%, a:p, Ma:Mp, the indices PTI/DEPI/CDI/S-CON/HVI/OBS, etc.) in the same format as the existing `Saroj_Bai` and `Amit_Pal` reports in this repo.

---

*Coding performed under the Exner Comprehensive System per the repository coding guide. FQ/Z values use standard published CS norms pending verification against the client's workbook Table A.*

**Made by Tejas Dangodra**
