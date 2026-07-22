#!/usr/bin/env python3
"""
MCMI-III COMPLETE SCORING TOOL - FINAL VERSION
Exact data from Appendix B (pp.181-183), C.1 (pp.186-187), C.2 (p.188)
All in ONE sheet. Weighted scoring: Proto=2, Nonproto=1, False=1
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "MCMI-III Scoring"

# Styles
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
SEC_FONT = Font(name="Calibri", size=11, bold=True, color="1F4E79")
HDR_FONT = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="FFFFCC")
RESULT_FILL = PatternFill("solid", fgColor="C6EFCE")
GREY_FILL = PatternFill("solid", fgColor="E8E8E8")
SEC_FILL = PatternFill("solid", fgColor="D6E4F0")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
YEL_FILL = PatternFill("solid", fgColor="FFEB9C")
GRN_FILL = PatternFill("solid", fgColor="C6EFCE")
THIN = Border(left=Side('thin'), right=Side('thin'),
              top=Side('thin'), bottom=Side('thin'))


# ============================================================
# EXACT SCORING KEYS FROM APPENDIX B (pages 181-183)
# ============================================================
KEYS = {
    # Page 181
    '1': {  # Schizoid (16 items)
        'proto': [10, 27, 46, 92, 105, 148, 165],
        'nonproto': [4, 38, 48, 101, 142, 156, 167],
        'false': [32, 57]
    },
    '2A': {  # Avoidant (16 items)
        'proto': [18, 40, 69, 84, 99, 127, 141, 174],
        'nonproto': [47, 48, 146, 148, 151, 158],
        'false': [57, 80]
    },
    '2B': {  # Depressive (15 items)
        'proto': [20, 25, 47, 112, 123, 133, 145, 151],
        'nonproto': [24, 43, 83, 86, 142, 148, 154],
        'false': []
    },
    '3': {  # Dependent (16 items)
        'proto': [16, 35, 45, 73, 94, 108, 135, 169],
        'nonproto': [47, 56, 84, 120, 133, 141, 151],
        'false': [82]
    },
    '4': {  # Histrionic (17 items)
        'proto': [12, 21, 32, 51, 57, 80, 88],
        'nonproto': [10, 24, 27, 48, 69, 92, 99, 123, 127, 174],
        'false': []
    },
    # Page 182
    '5': {  # Narcissistic (24 items)
        'proto': [5, 26, 31, 67, 85, 93, 144, 159],
        'nonproto': [21, 38, 57, 80, 88, 116],
        'false': [35, 40, 47, 69, 84, 86, 94, 99, 141, 169]
    },
    '6A': {  # Antisocial (17 items)
        'proto': [17, 38, 53, 101, 113, 139, 166],
        'nonproto': [7, 13, 14, 21, 41, 52, 93, 122, 136],
        'false': [172]
    },
    '6B': {  # Sadistic/Aggressive (20 items)
        'proto': [9, 14, 28, 64, 87, 95, 116],
        'nonproto': [7, 13, 17, 33, 36, 39, 41, 49, 53, 79, 93, 96, 166],
        'false': []
    },
    '7': {  # Compulsive (17 items)
        'proto': [2, 29, 59, 82, 97, 114, 137, 172],
        'nonproto': [],
        'false': [7, 14, 22, 41, 53, 72, 101, 139, 166]
    },
    '8A': {  # Negativistic/Passive-Aggressive (16 items)
        'proto': [7, 15, 22, 36, 50, 60, 79, 115, 126],
        'nonproto': [6, 42, 83, 98, 122, 133, 166],
        'false': []
    },
    '8B': {  # Masochistic/Self-Defeating (15 items)
        'proto': [19, 43, 70, 90, 104, 122, 161],
        'nonproto': [18, 24, 25, 35, 40, 98, 148, 169],
        'false': []
    },
    'S': {  # Schizotypal (16 items)
        'proto': [8, 48, 71, 76, 117, 138, 156, 158, 162],
        'nonproto': [69, 99, 102, 134, 141, 148, 151],
        'false': []
    },
    'C': {  # Borderline (16 items)
        'proto': [30, 41, 72, 83, 98, 120, 134, 142, 154],
        'nonproto': [7, 22, 122, 135, 161, 166, 171],
        'false': []
    },
    'P': {  # Paranoid (17 items)
        'proto': [6, 33, 42, 49, 89, 103, 146, 167, 175],
        'nonproto': [8, 48, 60, 63, 115, 138, 158, 159],
        'false': []
    },
    'A': {  # Anxiety (14 items)
        'proto': [58, 75, 124, 147, 164, 170],
        'nonproto': [40, 61, 76, 108, 109, 135, 145, 149],
        'false': []
    },
    'H': {  # Somatoform (12 items)
        'proto': [4, 11, 37, 55, 74],
        'nonproto': [1, 75, 107, 111, 130, 145, 148],
        'false': []
    },
    # Page 183
    'N': {  # Bipolar: Manic (13 items)
        'proto': [3, 54, 96, 106, 125],
        'nonproto': [22, 41, 51, 83, 117, 134, 166, 170],
        'false': []
    },
    'D': {  # Dysthymia (14 items)
        'proto': [24, 56, 62, 86, 111, 130],
        'nonproto': [15, 25, 55, 83, 104, 141, 142, 148],
        'false': []
    },
    'B': {  # Alcohol Dependence (15 items)
        'proto': [52, 77, 100, 131, 152],
        'nonproto': [14, 41, 64, 93, 101, 113, 122, 139, 166],
        'false_proto': [23]  # "False prototypal item (weight = 2)"
    },
    'T': {  # Drug Dependence (14 items)
        'proto': [13, 39, 66, 91, 118, 136],
        'nonproto': [7, 21, 38, 41, 53, 101, 113, 139],
        'false': []
    },
    'R': {  # PTSD (16 items)
        'proto': [109, 129, 149, 160, 173],
        'nonproto': [62, 76, 83, 123, 133, 142, 147, 148, 151, 154, 164],
        'false': []
    },
    'SS': {  # Thought Disorder (17 items)
        'proto': [34, 61, 68, 78, 102, 168],
        'nonproto': [22, 56, 72, 76, 83, 117, 134, 142, 148, 151, 162],
        'false': []
    },
    'CC': {  # Major Depression (17 items)
        'proto': [1, 44, 107, 128, 150, 171],
        'nonproto': [4, 34, 55, 74, 104, 111, 130, 142, 148, 149, 154],
        'false': []
    },
    'PP': {  # Delusional Disorder (13 items)
        'proto': [63, 119, 140, 153],
        'nonproto': [5, 38, 49, 67, 89, 103, 138, 159, 175],
        'false': []
    },
    # Scale Y - Desirability (21 items) - all weight=1
    'Y': {
        'true': [32, 51, 57, 59, 80, 82, 88, 97, 137, 172],
        'false': [20, 35, 40, 69, 104, 112, 123, 141, 142, 148, 151]
    },
    # Scale Z - Debasement (33 items) - all weight=1, TRUE only
    'Z': {
        'true': [1, 4, 8, 15, 22, 24, 30, 34, 36, 37, 44, 55, 56, 58, 62, 63,
                 70, 74, 75, 76, 83, 84, 86, 99, 111, 123, 128, 133, 134, 142,
                 145, 150, 171],
        'false': []
    },
    # Scale V - Invalidity (3 items) - weight=1
    'V': {
        'true': [65, 110, 157],
        'false': []
    },
}


# ============================================================
# BR TABLES FROM APPENDIX C.1 (pages 186-187) - EXACT from photos
# Index = raw score, value = BR score
# ============================================================
# Page 186: Scales 1, 2A, 2B, 3, 4, 5, 6A, 6B, 7, 8A, 8B, S, C, P
BR = {
    '1':  [0,12,24,36,48,60,62,64,66,68,70,72,75,78,81,85,89,93,97,101,105,109,112,115,115,115,115],
    '2A': [0,12,24,36,48,60,63,67,71,75,77,79,81,83,85,88,91,94,97,100,103,106,109,112,115],
    '2B': [0,10,20,30,40,50,60,65,70,75,77,79,81,83,85,89,92,96,99,103,106,109,112,115],
    '3':  [0,10,20,30,40,50,60,65,70,75,78,81,83,85,87,89,91,94,97,100,103,106,109,112,115],
    '4':  [0,4,8,12,16,20,24,28,32,36,40,44,48,51,54,57,60,63,66,69,72,75,79,83,88],
    '5':  [0,5,10,15,20,25,30,35,40,44,48,52,56,60,63,67,71,75,85,89,93,97,101,105,110,115],
    '6A': [0,12,24,36,48,60,62,64,66,69,71,73,75,75,79,82,85,89,92,96,99,103,106,109,112,115],
    '6B': [0,12,24,36,48,60,62,64,66,68,69,70,71,72,73,74,75,78,80,83,85,90,95,100,105,110,115],
    '7':  [0,4,8,12,16,20,24,28,32,36,39,42,45,48,51,54,57,60,63,66,69,72,75,79,83,87],
    '8A': [0,10,20,30,40,50,60,62,64,66,68,70,72,72,75,75,77,77,79,81,83,85,89,93,97,101,105,110,115],
    '8B': [0,20,40,60,63,66,69,75,78,80,82,85,88,91,94,97,100,103,106,109,112,115],
    'S':  [0,20,40,60,62,64,64,67,68,69,70,71,72,73,74,75,78,81,85,90,95,99,103,107,111,115],
    'C':  [0,12,24,36,48,60,63,66,69,72,75,77,77,79,81,83,85,88,91,94,97,100,103,106,109,112,115],
    'P':  [0,15,30,45,60,61,63,64,66,67,69,70,72,73,75,77,79,81,83,85,90,95,100,105,110,115],
}

# Page 187: Scales A, H, N, D, B, T, R, SS, CC, PP, Y, Z
BR2 = {
    'A':  [0,20,40,60,75,77,79,81,83,85,87,89,91,94,97,100,103,106,109,112,115],
    'H':  [0,15,30,45,60,62,64,66,68,70,72,73,74,75,80,85,100,115],
    'N':  [0,12,24,36,48,60,63,66,69,72,75,80,85,90,95,100,105,110,115],
    'D':  [0,12,24,36,48,60,62,64,66,69,72,75,78,80,82,85,91,97,103,109,115],
    'B':  [0,20,40,60,63,67,71,75,77,80,83,85,88,91,95,99,103,107,111,115],
    'T':  [0,20,40,60,63,67,71,75,76,77,78,79,81,83,85,90,95,100,105,110,115],
    'R':  [0,15,30,45,60,62,63,65,66,68,69,71,73,75,77,79,81,83,85,95,105,115],
    'SS': [0,15,30,45,60,62,64,66,67,68,69,70,71,72,73,74,75,79,82,85,93,100,108,115],
    'CC': [0,15,30,45,60,65,70,75,78,81,87,89,91,93,95,97,99,101,103,106,109,112,115],
    'PP': [0,30,60,62,65,68,70,72,75,80,85,90,95,100,105,110,115],
    'Y':  [0,5,10,15,20,25,30,35,39,43,47,51,55,59,63,67,71,75,80,85,93,100],
    'Z':  [0,18,35,38,40,42,45,47,49,52,54,56,59,61,63,66,68,70,73,75,76,78,79,81,82,84,85,86,88,90,92,94,96,98,100],
}


# ============================================================
# APPENDIX C.2 - Scale X BR Transformations (page 188)
# ============================================================
BR_X = {}
for r in range(34, 39): BR_X[r] = 0
BR_X[39] = 2; BR_X[40] = 3; BR_X[41] = 5; BR_X[42] = 6
BR_X[43] = 8; BR_X[44] = 9; BR_X[45] = 11; BR_X[46] = 12
BR_X[47] = 14; BR_X[48] = 15; BR_X[49] = 17; BR_X[50] = 18
BR_X[51] = 20; BR_X[52] = 21; BR_X[53] = 23; BR_X[54] = 24
BR_X[55] = 26; BR_X[56] = 27; BR_X[57] = 29; BR_X[58] = 30
BR_X[59] = 32; BR_X[60] = 33; BR_X[61] = 35; BR_X[62] = 36
for r in [63,64]: BR_X[r] = 37
BR_X[65] = 38
for r in [66,67]: BR_X[r] = 39
BR_X[68] = 40
for r in [69,70]: BR_X[r] = 41
BR_X[71] = 42
for r in [72,73]: BR_X[r] = 43
BR_X[74] = 44
for r in [75,76]: BR_X[r] = 45
BR_X[77] = 46
for r in [78,79]: BR_X[r] = 47
BR_X[80] = 48
for r in [81,82]: BR_X[r] = 49
BR_X[83] = 50
for r in [84,85]: BR_X[r] = 51
BR_X[86] = 52
for r in [87,88]: BR_X[r] = 53
BR_X[89] = 54
for r in [90,91]: BR_X[r] = 55
BR_X[92] = 56
for r in [93,94]: BR_X[r] = 57
BR_X[95] = 58
for r in [96,97]: BR_X[r] = 59
BR_X[98] = 60
for r in [99,100]: BR_X[r] = 61
BR_X[101] = 62
for r in [102,103]: BR_X[r] = 63
BR_X[104] = 64
for r in [105,106]: BR_X[r] = 65
BR_X[107] = 66
for r in [108,109]: BR_X[r] = 67
BR_X[110] = 68
for r in [111,112]: BR_X[r] = 69
BR_X[113] = 70
for r in [114,115]: BR_X[r] = 71
BR_X[116] = 72
for r in [117,118]: BR_X[r] = 73
for r in [119,120]: BR_X[r] = 74
for r in [121,122,123]: BR_X[r] = 75
for r in [124,125,126]: BR_X[r] = 76
for r in [127,128,129]: BR_X[r] = 77
for r in [130,131,132]: BR_X[r] = 78
for r in [133,134,135]: BR_X[r] = 79
for r in [136,137,138]: BR_X[r] = 80
for r in [139,140,141]: BR_X[r] = 81
for r in [142,143,144]: BR_X[r] = 82
for r in [145,146]: BR_X[r] = 83
for r in [147,148]: BR_X[r] = 84
for r in [149,150]: BR_X[r] = 85
for r in [151,152]: BR_X[r] = 86
BR_X[153] = 87
for r in [154,155]: BR_X[r] = 88
for r in [156,157]: BR_X[r] = 89
for r in [158,159]: BR_X[r] = 90
BR_X[160] = 91
for r in [161,162]: BR_X[r] = 92
for r in [163,164]: BR_X[r] = 93
for r in [165,166]: BR_X[r] = 94
for r in [167,168]: BR_X[r] = 95
for r in [169,170]: BR_X[r] = 96
for r in [171,172]: BR_X[r] = 97
BR_X[173] = 98; BR_X[174] = 99
for r in range(175, 179): BR_X[r] = 100

# ============================================================
# DISCLOSURE ADJUSTMENT TABLE (from existing verified data)
# ============================================================
DISC = {}
for i in range(0, 37): DISC[i] = (20, 10)
DISC[37]=(19,10); DISC[38]=(18,10); DISC[39]=(17,9); DISC[40]=(17,9)
DISC[41]=(16,9); DISC[42]=(15,8); DISC[43]=(14,8); DISC[44]=(13,7)
DISC[45]=(13,7); DISC[46]=(12,7); DISC[47]=(11,6); DISC[48]=(10,6)
DISC[49]=(9,5); DISC[50]=(9,5); DISC[51]=(8,5); DISC[52]=(7,4)
DISC[53]=(6,4); DISC[54]=(5,3); DISC[55]=(5,3); DISC[56]=(4,3)
DISC[57]=(3,2); DISC[58]=(2,2); DISC[59]=(1,1); DISC[60]=(1,1)
for i in range(61,124): DISC[i] = (0,0)
DISC[124]=(-1,-1); DISC[125]=(-1,-1); DISC[126]=(-1,-1)
DISC[127]=(-2,-2); DISC[128]=(-2,-2)
DISC[129]=(-3,-2); DISC[130]=(-3,-2); DISC[131]=(-3,-2)
DISC[132]=(-4,-3); DISC[133]=(-4,-3)
DISC[134]=(-5,-3); DISC[135]=(-5,-3); DISC[136]=(-5,-3)
DISC[137]=(-6,-4); DISC[138]=(-6,-4)
DISC[139]=(-7,-4); DISC[140]=(-7,-4); DISC[141]=(-7,-4)
DISC[142]=(-8,-5); DISC[143]=(-8,-5)
DISC[144]=(-9,-5); DISC[145]=(-9,-5); DISC[146]=(-9,-5)
DISC[147]=(-10,-6); DISC[148]=(-10,-6)
DISC[149]=(-11,-6); DISC[150]=(-11,-6); DISC[151]=(-11,-6)
DISC[152]=(-12,-7); DISC[153]=(-12,-7)
DISC[154]=(-13,-7); DISC[155]=(-13,-7); DISC[156]=(-13,-7)
DISC[157]=(-14,-8); DISC[158]=(-14,-8)
DISC[159]=(-15,-8); DISC[160]=(-15,-8); DISC[161]=(-15,-8)
DISC[162]=(-16,-9); DISC[163]=(-16,-9)
DISC[164]=(-17,-9); DISC[165]=(-17,-9); DISC[166]=(-17,-9)
DISC[167]=(-18,-10); DISC[168]=(-18,-10)
DISC[169]=(-19,-10); DISC[170]=(-19,-10); DISC[171]=(-19,-10)
for i in range(172, 200): DISC[i] = (-20,-11)

# A/D Tables
AD_2AS = {}
for i in range(0,8): AD_2AS[i]=-1
for i in range(8,16): AD_2AS[i]=-2
for i in range(16,24): AD_2AS[i]=-3
for i in range(24,32): AD_2AS[i]=-4
for i in range(32,40): AD_2AS[i]=-5
for i in range(40,48): AD_2AS[i]=-6
for i in range(48,56): AD_2AS[i]=-7
for i in range(56,64): AD_2AS[i]=-8
for i in range(64,72): AD_2AS[i]=-9
for i in range(72,81): AD_2AS[i]=-10

AD_2B8BC = {}
for i in range(0,10): AD_2B8BC[i]=-1
for i in range(10,15): AD_2B8BC[i]=-2
for i in range(15,20): AD_2B8BC[i]=-3
for i in range(20,25): AD_2B8BC[i]=-4
for i in range(25,30): AD_2B8BC[i]=-5
for i in range(30,35): AD_2B8BC[i]=-6
for i in range(35,40): AD_2B8BC[i]=-7
for i in range(40,45): AD_2B8BC[i]=-8
for i in range(45,50): AD_2B8BC[i]=-9
for i in range(50,55): AD_2B8BC[i]=-10
for i in range(55,60): AD_2B8BC[i]=-11
for i in range(60,65): AD_2B8BC[i]=-12
for i in range(65,70): AD_2B8BC[i]=-13
for i in range(70,75): AD_2B8BC[i]=-14
for i in range(75,81): AD_2B8BC[i]=-15


# ============================================================
# SCALE ORDER
# ============================================================
SCALES = [
    ('X', 'Disclosure', 'Modifying'),
    ('Y', 'Desirability', 'Modifying'),
    ('Z', 'Debasement', 'Modifying'),
    ('1', 'Schizoid', 'Personality'),
    ('2A', 'Avoidant', 'Personality'),
    ('2B', 'Depressive', 'Personality'),
    ('3', 'Dependent', 'Personality'),
    ('4', 'Histrionic', 'Personality'),
    ('5', 'Narcissistic', 'Personality'),
    ('6A', 'Antisocial', 'Personality'),
    ('6B', 'Sadistic', 'Personality'),
    ('7', 'Compulsive', 'Personality'),
    ('8A', 'Negativistic', 'Personality'),
    ('8B', 'Masochistic', 'Personality'),
    ('S', 'Schizotypal', 'Severe Pers'),
    ('C', 'Borderline', 'Severe Pers'),
    ('P', 'Paranoid', 'Severe Pers'),
    ('A', 'Anxiety', 'Clinical'),
    ('H', 'Somatoform', 'Clinical'),
    ('N', 'Bipolar Manic', 'Clinical'),
    ('D', 'Dysthymia', 'Clinical'),
    ('B', 'Alcohol Dep', 'Clinical'),
    ('T', 'Drug Dep', 'Clinical'),
    ('R', 'PTSD', 'Clinical'),
    ('SS', 'Thought Disorder', 'Severe Clin'),
    ('CC', 'Major Depression', 'Severe Clin'),
    ('PP', 'Delusional Dis', 'Severe Clin'),
]

# ============================================================
# LAYOUT: SECTION 1 - Header + Patient (rows 1-5)
# ============================================================
ws['A1'] = "MCMI-III COMPLETE SCORING TOOL"
ws['A1'].font = TITLE_FONT
ws['A2'] = "Enter 1=TRUE or 0=FALSE for each item. All scoring is automatic."
ws['A2'].font = Font(italic=True, size=9)
ws['A4'] = "Patient:"; ws['B4'].fill = INPUT_FILL
ws['D4'] = "Date:"; ws['E4'].fill = INPUT_FILL
ws['G4'] = "Age:"; ws['H4'].fill = INPUT_FILL
ws['A5'] = "Inpatient(Y/N):"; ws['B5'].fill = INPUT_FILL
ws['D5'] = "Axis I wks:"; ws['E5'].fill = INPUT_FILL
ws['G5'] = "Gender:"; ws['H5'].fill = INPUT_FILL

# ============================================================
# SECTION 2: ITEM RESPONSES (rows 7-43)
# 175 items in 5 groups of 35: A/B, D/E, G/H, J/K, M/N
# ============================================================
ws['A7'] = "ITEM RESPONSES (Enter 1=True, 0=False)"
ws['A7'].font = SEC_FONT; ws['A7'].fill = SEC_FILL

IR = 8  # item header row
item_cols = [('A','B'),('D','E'),('G','H'),('J','K'),('M','N')]
for nc, rc in item_cols:
    ws[f'{nc}{IR}'] = "#"; ws[f'{nc}{IR}'].font = HDR_FONT
    ws[f'{nc}{IR}'].fill = HDR_FILL; ws[f'{nc}{IR}'].border = THIN
    ws[f'{rc}{IR}'] = "T/F"; ws[f'{rc}{IR}'].font = HDR_FONT
    ws[f'{rc}{IR}'].fill = HDR_FILL; ws[f'{rc}{IR}'].border = THIN

for ci, (nc, rc) in enumerate(item_cols):
    start = ci * 35 + 1
    end = min(start + 34, 175)
    for item in range(start, end + 1):
        r = IR + 1 + (item - start)
        ws[f'{nc}{r}'] = item
        ws[f'{nc}{r}'].alignment = Alignment(horizontal='center')
        ws[f'{nc}{r}'].border = THIN
        ws[f'{rc}{r}'].fill = INPUT_FILL
        ws[f'{rc}{r}'].border = THIN
        ws[f'{rc}{r}'].alignment = Alignment(horizontal='center')

def cell_ref(item_num):
    """Absolute cell reference for an item response."""
    ci = (item_num - 1) // 35
    offset = (item_num - 1) % 35
    rc = item_cols[ci][1]
    row = IR + 1 + offset
    return f"${rc}${row}"


# ============================================================
# SECTION 3: SCORING TABLE (row 45 onwards)
# ============================================================
SC_ROW = IR + 1 + 35 + 1  # = 45
ws[f'A{SC_ROW}'] = "SCORING TABLE"
ws[f'A{SC_ROW}'].font = SEC_FONT; ws[f'A{SC_ROW}'].fill = SEC_FILL

HDR_ROW = SC_ROW + 1  # = 46
cols = ['Scale','Name','Cat','Raw','BR','Disc','A/D','Inp','Den','FINAL BR','Signif']
for ci, h in enumerate(cols):
    c = ws.cell(row=HDR_ROW, column=ci+1)
    c.value = h; c.font = HDR_FONT; c.fill = HDR_FILL
    c.border = THIN; c.alignment = Alignment(horizontal='center', wrap_text=True)

DATA_ROW = HDR_ROW + 1  # = 47
scale_rows = {}

for idx, (code, name, cat) in enumerate(SCALES):
    row = DATA_ROW + idx
    scale_rows[code] = row
    
    ws.cell(row=row, column=1, value=code).border = THIN
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=2, value=name).border = THIN
    ws.cell(row=row, column=3, value=cat).border = THIN
    
    # COLUMN D: RAW SCORE FORMULA
    if code == 'X':
        # Scale X = count of ALL items answered TRUE
        parts = [cell_ref(i) for i in range(1, 176)]
        raw_f = "=" + "+".join(parts)
    elif code in ('Y', 'Z', 'V'):
        key = KEYS[code]
        parts = []
        for it in key.get('true', []):
            parts.append(cell_ref(it))
        for it in key.get('false', []):
            parts.append(f"(1-{cell_ref(it)})")
        raw_f = "=" + "+".join(parts) if parts else "=0"
    elif code == 'B':
        # Alcohol Dep: has false_proto (weight=2)
        key = KEYS[code]
        parts = []
        for it in key['proto']:
            parts.append(f"2*{cell_ref(it)}")
        for it in key['nonproto']:
            parts.append(cell_ref(it))
        for it in key.get('false', []):
            parts.append(f"(1-{cell_ref(it)})")
        for it in key.get('false_proto', []):
            parts.append(f"2*(1-{cell_ref(it)})")
        raw_f = "=" + "+".join(parts)
    else:
        key = KEYS[code]
        parts = []
        for it in key['proto']:
            parts.append(f"2*{cell_ref(it)}")
        for it in key['nonproto']:
            parts.append(cell_ref(it))
        for it in key.get('false', []):
            parts.append(f"(1-{cell_ref(it)})")
        raw_f = "=" + "+".join(parts) if parts else "=0"
    
    ws.cell(row=row, column=4, value=raw_f)
    ws.cell(row=row, column=4).fill = RESULT_FILL
    ws.cell(row=row, column=4).border = THIN
    ws.cell(row=row, column=4).font = Font(bold=True)
    ws.cell(row=row, column=4).alignment = Alignment(horizontal='center')


# ============================================================
# EMBED LOOKUP TABLES IN HIDDEN COLUMNS (starting col 16 = P)
# ============================================================
LC = 16  # Lookup start column (P)

# BR lookup for personality/clinical scales
# Col P = raw score index, Q onwards = BR for each scale
br_scales = ['1','2A','2B','3','4','5','6A','6B','7','8A','8B',
             'S','C','P','A','H','N','D','B','T','R','SS','CC','PP','Y','Z']

ws.cell(row=1, column=LC, value="Raw")
for si, sc in enumerate(br_scales):
    ws.cell(row=1, column=LC+1+si, value=sc)

max_raw_len = max(len(BR.get(sc, BR2.get(sc, []))) for sc in br_scales)
for raw_idx in range(max_raw_len):
    ws.cell(row=2+raw_idx, column=LC, value=raw_idx)
    for si, sc in enumerate(br_scales):
        br_list = BR.get(sc, BR2.get(sc, []))
        if raw_idx < len(br_list):
            ws.cell(row=2+raw_idx, column=LC+1+si, value=br_list[raw_idx])
        else:
            ws.cell(row=2+raw_idx, column=LC+1+si, value=br_list[-1] if br_list else 0)

# Scale X BR lookup
XC = LC + len(br_scales) + 2
ws.cell(row=1, column=XC, value="X_Raw")
ws.cell(row=1, column=XC+1, value="X_BR")
for raw in range(34, 179):
    r = 2 + (raw - 34)
    ws.cell(row=r, column=XC, value=raw)
    ws.cell(row=r, column=XC+1, value=BR_X.get(raw, 100))

# Disclosure lookup
DC = XC + 3
ws.cell(row=1, column=DC, value="D_Raw")
ws.cell(row=1, column=DC+1, value="D_18B")
ws.cell(row=1, column=DC+2, value="D_SPP")
for raw in range(200):
    a18, asp = DISC.get(raw, (0,0))
    ws.cell(row=2+raw, column=DC, value=raw)
    ws.cell(row=2+raw, column=DC+1, value=a18)
    ws.cell(row=2+raw, column=DC+2, value=asp)

# AD_2AS lookup
AC = DC + 4
ws.cell(row=1, column=AC, value="AD_v")
ws.cell(row=1, column=AC+1, value="AD_2AS")
for v in range(81):
    ws.cell(row=2+v, column=AC, value=v)
    ws.cell(row=2+v, column=AC+1, value=AD_2AS[v])

# AD_2B8BC lookup
BC = AC + 3
ws.cell(row=1, column=BC, value="AD_v2")
ws.cell(row=1, column=BC+1, value="AD_2B8B")
for v in range(81):
    ws.cell(row=2+v, column=BC, value=v)
    ws.cell(row=2+v, column=BC+1, value=AD_2B8BC[v])

# Get column letters
RAW_LTR = get_column_letter(LC)
X_RAW_LTR = get_column_letter(XC)
X_BR_LTR = get_column_letter(XC+1)
D_RAW_LTR = get_column_letter(DC)
D_18B_LTR = get_column_letter(DC+1)
D_SPP_LTR = get_column_letter(DC+2)
AC_LTR = get_column_letter(AC)
AC_ADJ_LTR = get_column_letter(AC+1)
BC_LTR = get_column_letter(BC)
BC_ADJ_LTR = get_column_letter(BC+1)

# Map each scale to its BR column letter
scale_br_ltr = {}
for si, sc in enumerate(br_scales):
    scale_br_ltr[sc] = get_column_letter(LC+1+si)


# ============================================================
# COLUMNS E-K: BR, Adjustments, Final BR, Significance
# ============================================================
PERS_18B = ['1','2A','2B','3','4','5','6A','6B','7','8A','8B']
SPP = ['S','C','P','A','H','N','D','B','T','R','SS','CC','PP']
AD_2AS_SC = ['2A', 'S']
AD_2B_SC = ['2B', '8B', 'C']

x_raw = f"D{scale_rows['X']}"
a_row = scale_rows['A']
d_row = scale_rows['D']

for idx, (code, name, cat) in enumerate(SCALES):
    row = DATA_ROW + idx
    
    # COL E: BR SCORE (auto-lookup)
    if code == 'X':
        br_f = (f'=IF(D{row}<34,0,IF(D{row}>178,100,'
                f'VLOOKUP(D{row},${X_RAW_LTR}:${X_BR_LTR},2,TRUE)))')
    elif code in scale_br_ltr:
        br_col = scale_br_ltr[code]
        col_offset = ord(br_col[0]) - ord(RAW_LTR[0]) + 1 if len(br_col)==1 and len(RAW_LTR)==1 else None
        # Use INDEX/MATCH for reliability
        br_f = (f'=VLOOKUP(MIN(D{row},{max_raw_len-1}),'
                f'${RAW_LTR}:${br_col},'
                f'COLUMN(${br_col}$1)-COLUMN(${RAW_LTR}$1)+1,TRUE)')
    else:
        br_f = "=0"
    
    ws.cell(row=row, column=5, value=br_f)
    ws.cell(row=row, column=5).fill = RESULT_FILL
    ws.cell(row=row, column=5).border = THIN
    ws.cell(row=row, column=5).alignment = Alignment(horizontal='center')
    
    # COL F: DISCLOSURE ADJUSTMENT
    if code in ['X','Y','Z']:
        ws.cell(row=row, column=6, value=0)
        ws.cell(row=row, column=6).fill = GREY_FILL
    elif code in PERS_18B:
        ws.cell(row=row, column=6, value=(
            f'=VLOOKUP(MIN(MAX({x_raw},0),199),${D_RAW_LTR}:${D_18B_LTR},2,TRUE)'))
    else:
        ws.cell(row=row, column=6, value=(
            f'=VLOOKUP(MIN(MAX({x_raw},0),199),${D_RAW_LTR}:${D_SPP_LTR},3,TRUE)'))
    ws.cell(row=row, column=6).border = THIN
    ws.cell(row=row, column=6).alignment = Alignment(horizontal='center')
    
    # COL G: A/D ADJUSTMENT
    a_adj = f"(E{a_row}+F{a_row})"
    d_adj = f"(E{d_row}+F{d_row})"
    if code in AD_2AS_SC:
        ws.cell(row=row, column=7, value=(
            f'=IF(AND({a_adj}<75,{d_adj}<75),0,'
            f'VLOOKUP(MIN(MAX(IF({a_adj}>=75,{a_adj}-75,0)+IF({d_adj}>=75,{d_adj}-75,0),0),80),'
            f'${AC_LTR}:${AC_ADJ_LTR},2,TRUE))'))
    elif code in AD_2B_SC:
        ws.cell(row=row, column=7, value=(
            f'=IF(AND({a_adj}<75,{d_adj}<75),0,'
            f'VLOOKUP(MIN(MAX(IF({a_adj}>=75,{a_adj}-75,0)+IF({d_adj}>=75,{d_adj}-75,0),0),80),'
            f'${BC_LTR}:${BC_ADJ_LTR},2,TRUE))'))
    else:
        ws.cell(row=row, column=7, value=0)
        ws.cell(row=row, column=7).fill = GREY_FILL
    ws.cell(row=row, column=7).border = THIN
    ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')
    
    # COL H: INPATIENT ADJUSTMENT
    if code == 'SS':
        ws.cell(row=row, column=8, value='=IF($B$5="Y",IF($E$5<1,6,IF($E$5<=4,4,0)),0)')
    elif code == 'CC':
        ws.cell(row=row, column=8, value='=IF($B$5="Y",IF($E$5<1,10,IF($E$5<=4,8,0)),0)')
    elif code == 'PP':
        ws.cell(row=row, column=8, value='=IF($B$5="Y",IF($E$5<1,4,IF($E$5<=4,2,0)),0)')
    else:
        ws.cell(row=row, column=8, value=0)
        ws.cell(row=row, column=8).fill = GREY_FILL
    ws.cell(row=row, column=8).border = THIN
    ws.cell(row=row, column=8).alignment = Alignment(horizontal='center')
    
    # COL I: DENIAL/COMPLAINT (manual)
    ws.cell(row=row, column=9, value=0)
    ws.cell(row=row, column=9).fill = INPUT_FILL if code in PERS_18B else GREY_FILL
    ws.cell(row=row, column=9).border = THIN
    ws.cell(row=row, column=9).alignment = Alignment(horizontal='center')
    
    # COL J: FINAL BR
    ws.cell(row=row, column=10, value=f'=MIN(MAX(E{row}+F{row}+G{row}+H{row}+I{row},0),115)')
    ws.cell(row=row, column=10).border = THIN
    ws.cell(row=row, column=10).font = Font(bold=True, size=11)
    ws.cell(row=row, column=10).alignment = Alignment(horizontal='center')
    
    # COL K: SIGNIFICANCE
    ws.cell(row=row, column=11, value=(
        f'=IF(J{row}>=85,"PATHOLOGICAL",IF(J{row}>=75,"PRESENT",'
        f'IF(J{row}>=60,"Suggestive","Normal")))'))
    ws.cell(row=row, column=11).border = THIN
    ws.cell(row=row, column=11).alignment = Alignment(horizontal='center')


# ============================================================
# CONDITIONAL FORMATTING + VALIDITY + INSTRUCTIONS + SAVE
# ============================================================
last_row = DATA_ROW + len(SCALES) - 1
rng = f"J{DATA_ROW}:J{last_row}"
ws.conditional_formatting.add(rng, CellIsRule(
    operator='greaterThanOrEqual', formula=['85'],
    fill=RED_FILL, font=Font(bold=True, color="9C0006")))
ws.conditional_formatting.add(rng, CellIsRule(
    operator='between', formula=['75','84'],
    fill=YEL_FILL, font=Font(bold=True, color="9C6500")))
ws.conditional_formatting.add(rng, CellIsRule(
    operator='lessThan', formula=['60'],
    fill=GRN_FILL, font=Font(color="006100")))

# Validity
vr = last_row + 2
ws.cell(row=vr, column=1, value="VALIDITY").font = SEC_FONT
ws.cell(row=vr, column=1).fill = SEC_FILL
ws.cell(row=vr+1, column=1, value="V Score:")
ws.cell(row=vr+1, column=2, value=f"={cell_ref(65)}+{cell_ref(110)}+{cell_ref(157)}")
ws.cell(row=vr+1, column=2).fill = RESULT_FILL
ws.cell(row=vr+1, column=3, value=f'=IF(B{vr+1}>1,"INVALID","OK")')
ws.cell(row=vr+2, column=1, value="X Raw:")
ws.cell(row=vr+2, column=2, value=f"=D{scale_rows['X']}")
ws.cell(row=vr+2, column=2).fill = RESULT_FILL
ws.cell(row=vr+2, column=3, value=f'=IF(OR(B{vr+2}<34,B{vr+2}>178),"INVALID","OK")')
ws.cell(row=vr+3, column=1, value="STATUS:")
ws.cell(row=vr+3, column=1).font = Font(bold=True, size=11)
ws.cell(row=vr+3, column=2, value=(
    f'=IF(OR(B{vr+1}>1,B{vr+2}<34,B{vr+2}>178),"INVALID PROTOCOL","VALID")'))
ws.cell(row=vr+3, column=2).font = Font(bold=True, size=12)

# Interpretation + Instructions
ir2 = vr + 5
ws.cell(row=ir2, column=1, value="BR 85-115 = PATHOLOGICAL | 75-84 = PRESENT | 60-74 = Suggestive | 0-59 = Normal")
ws.cell(row=ir2+1, column=1, value="")
ws.cell(row=ir2+2, column=1, value="HOW TO USE:")
ws.cell(row=ir2+2, column=1).font = SEC_FONT
ws.cell(row=ir2+3, column=1, value="1. Enter 1 (TRUE) or 0 (FALSE) for all 175 items in yellow cells")
ws.cell(row=ir2+4, column=1, value="2. Raw scores auto-calculate (proto*2 + nonproto*1 + false*1)")
ws.cell(row=ir2+5, column=1, value="3. BR scores auto-calculate from Appendix C tables")
ws.cell(row=ir2+6, column=1, value="4. Adjustments (Disc, A/D, Inpatient) are automatic")
ws.cell(row=ir2+7, column=1, value="5. For Denial/Complaint: enter 8 in col I if highest 1-8B is scale 4, 5, or 7")
ws.cell(row=ir2+8, column=1, value="6. Final BR and significance auto-calculate")

# Column widths
for col, w in {'A':6,'B':15,'C':11,'D':7,'E':6,'F':6,'G':6,'H':5,'I':5,'J':8,'K':14}.items():
    ws.column_dimensions[col].width = w

# Hide lookup columns
for col_num in range(LC, BC+2):
    ws.column_dimensions[get_column_letter(col_num)].hidden = True

# ============================================================
# SAVE
# ============================================================
out = "/projects/sandbox/Dango-kiro/MCMI-III_Scoring_Tool_V2.xlsx"
wb.save(out)
import os
print(f"SUCCESS! {out}")
print(f"Size: {os.path.getsize(out)} bytes")
print(f"Scales: {len(SCALES)}, Data rows: {DATA_ROW}-{last_row}")
print(f"Lookup cols: {get_column_letter(LC)}-{get_column_letter(BC+1)} (hidden)")
print(f"\nScale rows: {scale_rows}")
