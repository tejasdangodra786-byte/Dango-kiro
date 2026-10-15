# MCMI-III Complete Scoring System

## Overview
A **clinically accurate, exam-level** MCMI-III (Millon Clinical Multiaxial Inventory - Third Edition) scoring system based strictly on:
- Official MCMI-III Manual (Millon, Davis, Grossman, Millon)
- MCMI-III Hand-Scoring User's Guide
- Standard BR (Base Rate) scoring methodology

## Quick Start

```bash
# Run OPD (Outpatient) demo
python3 mcmi_main.py --demo

# Run IPD (Inpatient) demo  
python3 mcmi_main.py --ipd-demo

# Interactive mode
python3 mcmi_main.py

# Score from CSV file
python3 mcmi_main.py --file responses.csv
```

## System Architecture

```
MCMI_III_Scoring_System/
|-- mcmi_main.py              # Main application (CLI + demos)
|-- mcmi_scoring_engine.py    # Core scoring algorithm
|-- mcmi_item_keys.py         # All scale item assignments  
|-- mcmi_br_tables.py         # BR conversion + adjustment tables
|-- mcmi_excel_generator.py   # Excel workbook generator
|-- README.md                 # This file (full documentation)
```

## File Structure (Mirrors Excel 5-Sheet Design)

| Module | Corresponds To | Function |
|--------|---------------|----------|
| `mcmi_main.py` | Sheet 1: Data Entry | Input/output, CLI |
| `mcmi_scoring_engine.py` | Sheet 2: Raw Scores + Sheet 3: Adjustments | All calculations |
| `mcmi_br_tables.py` | Sheet 4: BR Conversion | Lookup tables |
| `mcmi_excel_generator.py` | Sheet 5: Profile Output | Excel generation |

---

## Complete Scoring Algorithm

### Step 1: Age Check
- Patient must be >= 18 years old
- If age < 18 or not indicated: **INVALID**

### Step 2: Omits/Double Marks
- Count unanswered items
- If omitted items > 11: **INVALID**
- If <= 11: Cross out double marks, continue

### Step 3: Scale V (Invalidity)
- Items: **65, 110, 157**
- Score = Count of True responses
- V = 0: Valid
- V = 1: Questionable validity (flag, continue)
- V >= 2: **INVALID** (random responding)

### Step 4: Scale W (Inconsistency)
- 44 pairs of concordant items
- Score = number of pairs answered in opposite directions
- W < 8: Valid
- W = 8-9: Questionable
- W >= 10: **INVALID**

### Step 5: Raw Score Calculation
All 26 scales scored by summing weighted item endorsements:
- Weight-1 items: Add 1 if True
- Weight-2 (prototypal) items: Add 2 if True

### Step 6: Scale X (Disclosure) Raw Score
```
X = Sum of raw scores for: 1, 2A, 2B, 3, 4, (5 x 2/3), 6A, 6B, 7, 8A, 8B
```
- Scale 5 is multiplied by 2/3 before adding
- Result rounded to nearest whole number

### Step 7: Scale X Validity Check
- Valid range: **34 - 178**
- < 34: Extreme defensiveness/low disclosure - INVALID
- > 178: Extreme over-reporting - INVALID

### Step 8: Initial BR Score Conversion
- Raw scores converted to BR using Appendix C tables
- Linear interpolation between anchor points
- BR scores range: 0-115

### Step 9: Disclosure Adjustment (Table 1)
Based on raw X score:
- **1-8B Adjustment**: Applied to Clinical Personality scales (1-8B)
- **S-PP Adjustment**: Applied to Severe Personality + all Syndrome scales

| X Raw Score | 1-8B Adj | S-PP Adj |
|------------|----------|----------|
| 34-37 | +20 | +10 |
| 40-41 | +17 | +9 |
| 50 | +9 | +5 |
| 61-123 | 0 | 0 |
| 130 | -3 | -2 |
| 150 | -11 | -8 |
| 178 | -20 | -14 |

### Step 10: Anxiety/Depression (A/D) Adjustment
**Affects: Scales 2A, 2B, 8B, S, C**

1. Compute A/D value:
   - If A < 75 AND D < 75: No adjustment needed
   - If A >= 75 only: A/D value = A - 75
   - If D >= 75 only: A/D value = D - 75
   - If both >= 75: A/D value = (A-75) + (D-75)

2. Select table based on patient setting:
   - **Table 2**: OPD (Non-Inpatient) or IPD with duration > 4 weeks
   - **Table 3**: IPD with Axis I duration < 1 week
   - **Table 4**: IPD with Axis I duration 1-4 weeks

3. Apply:
   - 2B, 8B, C adjustment factor
   - 2A, S adjustment factor

### Step 11: Inpatient Adjustment (Table 5)
**Affects: Scales SS, CC, PP only**

Only applies to IPD patients with Axis I duration <= 4 weeks:

| Duration | SS | CC | PP |
|----------|-----|-----|-----|
| < 1 week | +6 | +10 | +4 |
| 1-4 weeks | +4 | +8 | +2 |
| > 4 weeks | 0 | 0 | 0 |

### Step 12: Denial/Complaint Adjustment
**Affects: Highest of Scales 4, 5, 7 only**

1. Find highest Clinical Personality scale (1-8B) by current BR
2. Tie-breaking order: 2B > 6A > 8B > 6B > 2A > 8A > 7 > 4 > 5 > 3 > 1
3. If highest is 4, 5, or 7: Add **+8** to that scale only
4. If highest is anything else: No adjustment

### Step 13: Final BR Scores
- Transfer all adjusted scores as final BR scores
- Clamp all values: 0 <= BR <= 115
- Check for **FLAT PROFILE**: All personality BR < 60 = Uninterpretable

---

## BR Score Interpretation

| BR Score | Interpretation |
|----------|---------------|
| < 60 | Not significant |
| 60-74 | Trait/Tendency present |
| 75-84 | Clinically Significant |
| >= 85 | **PROMINENT - Disorder Present** |

---

## All 29 Scales

### Modifying Indices
| Scale | Name | Purpose |
|-------|------|---------|
| X | Disclosure | Response openness level |
| Y | Desirability | Social desirability bias |
| Z | Debasement | Over-reporting/cry for help |

### Validity Indices
| Scale | Name | Purpose |
|-------|------|---------|
| V | Invalidity | Random responding detection |
| W | Inconsistency | Response inconsistency |

### Clinical Personality Patterns (Axis II)
| Scale | Name |
|-------|------|
| 1 | Schizoid |
| 2A | Avoidant |
| 2B | Depressive |
| 3 | Dependent |
| 4 | Histrionic |
| 5 | Narcissistic |
| 6A | Antisocial |
| 6B | Aggressive (Sadistic) |
| 7 | Compulsive |
| 8A | Negativistic (Passive-Aggressive) |
| 8B | Masochistic (Self-Defeating) |

### Severe Personality Pathology
| Scale | Name |
|-------|------|
| S | Schizotypal |
| C | Borderline |
| P | Paranoid |

### Clinical Syndromes (Axis I)
| Scale | Name |
|-------|------|
| A | Anxiety |
| H | Somatoform |
| N | Bipolar: Manic |
| D | Dysthymia |
| B | Alcohol Dependence |
| T | Drug Dependence |
| R | PTSD |

### Severe Clinical Syndromes
| Scale | Name |
|-------|------|
| SS | Thought Disorder |
| CC | Major Depression |
| PP | Delusional Disorder |

---

## IPD vs OPD Patient Differences

| Feature | OPD (Outpatient) | IPD (Inpatient) |
|---------|------------------|-----------------|
| A/D Table | Table 2 | Table 3 or 4 (by duration) |
| Inpatient Adj | Not applied | Applied to SS, CC, PP |
| Duration needed | No | Yes (< 1wk, 1-4wk, > 4wk) |
| Time criteria | N/A | Axis I disorder duration |

---

## Input Formats

### CSV Format
```csv
Item,Response
1,T
2,F
3,T
...
175,F
```

### JSON Format
```json
{"1": true, "2": false, "3": true, ... "175": false}
```

### Or as a list:
```json
[true, false, true, ..., false]
```

---

## Important Notes

1. **BR Conversion Tables**: The conversion anchors use linear interpolation between published reference points. For maximum accuracy, the full Appendix C tables from the manual should be entered.

2. **Item Keys**: Based on the official MCMI-III scoring templates. Each scale has specific items with specific weights (1 or 2).

3. **All adjustments are CUMULATIVE** - always use the latest adjusted BR scores for subsequent steps.

4. **After any adjustment**: Clamp BR values to 0-115 range.

5. **Gender**: MCMI-III uses the same BR conversion tables for males and females.

---

## Clinical Use Disclaimer

This tool is designed for **educational and clinical training purposes**. 
Professional use requires:
- Proper MCMI-III licensure from Pearson/NCS
- Clinical training in psychometric assessment
- Adherence to ethical guidelines for test administration
