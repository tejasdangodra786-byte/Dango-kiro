#!/usr/bin/env python3
"""
MCMI-III COMPLETE SCORING TOOL - FINAL
Exact item keys provided by licensed test owner.
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
# EXACT SCORING KEYS FROM USER (Appendix B pp.181-183)
# ============================================================
KEYS = {
    '1': {'proto': [10,27,46,92,105,148,165],
          'nonproto': [4,38,48,101,142,156,167],
          'false': [32,57]},
    '2A': {'proto': [18,40,69,84,99,127,141,174],
           'nonproto': [47,48,146,148,151,158],
           'false': [57,80]},
    '2B': {'proto': [20,25,47,112,123,133,145,151],
           'nonproto': [24,43,83,86,142,148,154],
           'false': []},
    '3': {'proto': [16,35,45,73,94,108,135,169],
          'nonproto': [47,56,84,120,133,141,151],
          'false': [82]},
    '4': {'proto': [12,21,32,51,57,80,88],
          'nonproto': [],
          'false': [10,24,27,48,69,92,99,123,127,174]},
    '5': {'proto': [5,26,31,67,85,93,144,159],
          'nonproto': [21,38,57,80,88,116],
          'false': [35,40,47,69,84,86,94,99,141,169]},
    '6A': {'proto': [17,38,53,101,113,139,166],
           'nonproto': [7,13,14,21,41,52,93,122,136],
           'false': [172]},
    '6B': {'proto': [9,14,28,64,87,95,116],
           'nonproto': [7,13,17,33,36,39,41,49,53,79,93,96,166],
           'false': []},
    '7': {'proto': [2,29,59,82,97,114,137,172],
          'nonproto': [],
          'false': [7,14,22,41,53,72,101,139,166]},
    '8A': {'proto': [7,15,22,36,50,60,79,115,126],
           'nonproto': [6,42,83,98,122,133,166],
           'false': []},
    '8B': {'proto': [19,43,70,90,104,122,161],
           'nonproto': [18,24,25,35,40,98,148,169],
           'false': []},
    'S': {'proto': [8,48,71,76,117,138,156,158,162],
          'nonproto': [69,99,102,134,141,148,151],
          'false': []},
    'C': {'proto': [30,41,72,83,98,120,134,142,154],
          'nonproto': [7,22,122,135,161,166,171],
          'false': []},
    'P': {'proto': [6,33,42,49,89,103,146,167,175],
          'nonproto': [8,48,60,63,115,138,158,159],
          'false': []},
    'A': {'proto': [58,75,124,147,164,170],
          'nonproto': [40,61,76,108,109,135,145,149],
          'false': []},
    'H': {'proto': [4,11,37,55,74],
          'nonproto': [1,75,107,111,130,145,148],
          'false': []},
    'N': {'proto': [3,54,96,106,125],
          'nonproto': [22,41,51,83,117,134,166,170],
          'false': []},
    'D': {'proto': [24,56,62,86,111,130],
          'nonproto': [15,25,55,83,104,141,142,148],
          'false': []},
    'B': {'proto': [52,77,100,131,152],
          'nonproto': [14,41,64,93,101,113,122,139,166],
          'false': [],
          'false_proto': [23]},
    'T': {'proto': [13,39,66,91,118,136],
          'nonproto': [7,21,38,41,53,101,113,139],
          'false': []},
    'R': {'proto': [109,129,149,160,173],
          'nonproto': [62,76,83,123,133,142,147,148,151,154,164],
          'false': []},
    'SS': {'proto': [34,61,68,78,102,168],
           'nonproto': [22,56,72,76,83,117,134,142,148,151,162],
           'false': []},
    'CC': {'proto': [1,44,107,128,150,171],
           'nonproto': [4,34,55,74,104,111,130,142,148,149,154],
           'false': []},
    'PP': {'proto': [63,119,140,153],
           'nonproto': [5,38,49,67,89,103,138,159,175],
           'false': []},
    'Y': {'true': [32,51,57,59,80,82,88,97,137,172],
           'false': [20,35,40,69,104,112,123,141,142,148,151]},
    'Z': {'true': [1,4,8,15,22,24,30,34,36,37,44,55,56,58,62,63,
                   70,74,75,76,83,84,86,99,111,123,128,133,134,142,
                   145,150,171],
           'false': []},
    'V': {'true': [65,110,157], 'false': []},
}

# Scale W Inconsistency pairs (from user)
W_PAIRS = [
    (1,'T',4,'F'), (1,'F',4,'T'), (8,'F',141,'T'), (13,'T',66,'F'),
    (13,'F',66,'T'), (15,'T',133,'F'), (20,'T',112,'F'), (20,'F',112,'T'),
    (22,'T',83,'F'), (24,'F',151,'T'), (25,'T',56,'F'), (25,'F',56,'T'),
    (27,'F',92,'T'), (32,'T',80,'F'), (35,'T',84,'F'), (35,'F',84,'T'),
    (39,'T',118,'F'), (39,'F',118,'T'), (41,'F',166,'T'), (44,'T',86,'F'),
    (44,'F',150,'T'), (48,'T',92,'F'), (49,'T',146,'F'), (52,'F',152,'T'),
    (55,'F',130,'T'), (57,'F',80,'T'), (61,'F',76,'T'), (62,'F',86,'T'),
    (68,'F',162,'T'), (69,'T',99,'F'), (70,'F',104,'T'), (72,'F',142,'T'),
    (74,'F',107,'T'), (77,'T',131,'F'), (91,'T',136,'F'), (91,'F',136,'T'),
    (108,'T',135,'F'), (109,'F',164,'T'), (123,'F',128,'T'), (129,'T',173,'F'),
    (133,'T',145,'F'), (147,'F',149,'T'), (160,'T',164,'F'), (160,'F',173,'T'),
]


# ============================================================
# BR TABLES (from Appendix C.1 pp.186-187, C.2 p.188)
# Index = raw score, value = BR score
# ============================================================
BR = {
    '1':  [0,12,24,36,48,60,62,64,66,68,70,72,75,78,81,85,89,93,97,101,105,109,112,115],
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

# Scale X BR (Appendix C.2 p.188)
BR_X = {}
for r in range(34,39): BR_X[r]=0
BR_X[39]=2;BR_X[40]=3;BR_X[41]=5;BR_X[42]=6;BR_X[43]=8;BR_X[44]=9
BR_X[45]=11;BR_X[46]=12;BR_X[47]=14;BR_X[48]=15;BR_X[49]=17;BR_X[50]=18
BR_X[51]=20;BR_X[52]=21;BR_X[53]=23;BR_X[54]=24;BR_X[55]=26;BR_X[56]=27
BR_X[57]=29;BR_X[58]=30;BR_X[59]=32;BR_X[60]=33;BR_X[61]=35;BR_X[62]=36
for r in [63,64]: BR_X[r]=37
BR_X[65]=38
for r in [66,67]: BR_X[r]=39
BR_X[68]=40
for r in [69,70]: BR_X[r]=41
BR_X[71]=42
for r in [72,73]: BR_X[r]=43
BR_X[74]=44
for r in [75,76]: BR_X[r]=45
BR_X[77]=46
for r in [78,79]: BR_X[r]=47
BR_X[80]=48
for r in [81,82]: BR_X[r]=49
BR_X[83]=50
for r in [84,85]: BR_X[r]=51
BR_X[86]=52
for r in [87,88]: BR_X[r]=53
BR_X[89]=54
for r in [90,91]: BR_X[r]=55
BR_X[92]=56
for r in [93,94]: BR_X[r]=57
BR_X[95]=58
for r in [96,97]: BR_X[r]=59
BR_X[98]=60
for r in [99,100]: BR_X[r]=61
BR_X[101]=62
for r in [102,103]: BR_X[r]=63
BR_X[104]=64
for r in [105,106]: BR_X[r]=65
BR_X[107]=66
for r in [108,109]: BR_X[r]=67
BR_X[110]=68
for r in [111,112]: BR_X[r]=69
BR_X[113]=70
for r in [114,115]: BR_X[r]=71
BR_X[116]=72
for r in [117,118]: BR_X[r]=73
for r in [119,120]: BR_X[r]=74
for r in [121,122,123]: BR_X[r]=75
for r in [124,125,126]: BR_X[r]=76
for r in [127,128,129]: BR_X[r]=77
for r in [130,131,132]: BR_X[r]=78
for r in [133,134,135]: BR_X[r]=79
for r in [136,137,138]: BR_X[r]=80
for r in [139,140,141]: BR_X[r]=81
for r in [142,143,144]: BR_X[r]=82
for r in [145,146]: BR_X[r]=83
for r in [147,148]: BR_X[r]=84
for r in [149,150]: BR_X[r]=85
for r in [151,152]: BR_X[r]=86
BR_X[153]=87
for r in [154,155]: BR_X[r]=88
for r in [156,157]: BR_X[r]=89
for r in [158,159]: BR_X[r]=90
BR_X[160]=91
for r in [161,162]: BR_X[r]=92
for r in [163,164]: BR_X[r]=93
for r in [165,166]: BR_X[r]=94
for r in [167,168]: BR_X[r]=95
for r in [169,170]: BR_X[r]=96
for r in [171,172]: BR_X[r]=97
BR_X[173]=98;BR_X[174]=99
for r in range(175,179): BR_X[r]=100


# Disclosure Adjustment Table
DISC = {}
for i in range(0,37): DISC[i]=(20,10)
DISC[37]=(19,10);DISC[38]=(18,10);DISC[39]=(17,9);DISC[40]=(17,9)
DISC[41]=(16,9);DISC[42]=(15,8);DISC[43]=(14,8);DISC[44]=(13,7)
DISC[45]=(13,7);DISC[46]=(12,7);DISC[47]=(11,6);DISC[48]=(10,6)
DISC[49]=(9,5);DISC[50]=(9,5);DISC[51]=(8,5);DISC[52]=(7,4)
DISC[53]=(6,4);DISC[54]=(5,3);DISC[55]=(5,3);DISC[56]=(4,3)
DISC[57]=(3,2);DISC[58]=(2,2);DISC[59]=(1,1);DISC[60]=(1,1)
for i in range(61,124): DISC[i]=(0,0)
DISC[124]=(-1,-1);DISC[125]=(-1,-1);DISC[126]=(-1,-1)
DISC[127]=(-2,-2);DISC[128]=(-2,-2)
DISC[129]=(-3,-2);DISC[130]=(-3,-2);DISC[131]=(-3,-2)
DISC[132]=(-4,-3);DISC[133]=(-4,-3)
DISC[134]=(-5,-3);DISC[135]=(-5,-3);DISC[136]=(-5,-3)
DISC[137]=(-6,-4);DISC[138]=(-6,-4)
DISC[139]=(-7,-4);DISC[140]=(-7,-4);DISC[141]=(-7,-4)
DISC[142]=(-8,-5);DISC[143]=(-8,-5)
DISC[144]=(-9,-5);DISC[145]=(-9,-5);DISC[146]=(-9,-5)
DISC[147]=(-10,-6);DISC[148]=(-10,-6)
DISC[149]=(-11,-6);DISC[150]=(-11,-6);DISC[151]=(-11,-6)
DISC[152]=(-12,-7);DISC[153]=(-12,-7)
DISC[154]=(-13,-7);DISC[155]=(-13,-7);DISC[156]=(-13,-7)
DISC[157]=(-14,-8);DISC[158]=(-14,-8)
DISC[159]=(-15,-8);DISC[160]=(-15,-8);DISC[161]=(-15,-8)
DISC[162]=(-16,-9);DISC[163]=(-16,-9)
DISC[164]=(-17,-9);DISC[165]=(-17,-9);DISC[166]=(-17,-9)
DISC[167]=(-18,-10);DISC[168]=(-18,-10)
DISC[169]=(-19,-10);DISC[170]=(-19,-10);DISC[171]=(-19,-10)
for i in range(172,200): DISC[i]=(-20,-11)

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

SCALES = [
    ('X','Disclosure','Modifying'),('Y','Desirability','Modifying'),
    ('Z','Debasement','Modifying'),('1','Schizoid','Personality'),
    ('2A','Avoidant','Personality'),('2B','Depressive','Personality'),
    ('3','Dependent','Personality'),('4','Histrionic','Personality'),
    ('5','Narcissistic','Personality'),('6A','Antisocial','Personality'),
    ('6B','Sadistic','Personality'),('7','Compulsive','Personality'),
    ('8A','Negativistic','Personality'),('8B','Masochistic','Personality'),
    ('S','Schizotypal','Severe Pers'),('C','Borderline','Severe Pers'),
    ('P','Paranoid','Severe Pers'),('A','Anxiety','Clinical'),
    ('H','Somatoform','Clinical'),('N','Bipolar Manic','Clinical'),
    ('D','Dysthymia','Clinical'),('B','Alcohol Dep','Clinical'),
    ('T','Drug Dep','Clinical'),('R','PTSD','Clinical'),
    ('SS','Thought Dis','Severe Clin'),('CC','Major Dep','Severe Clin'),
    ('PP','Delusional','Severe Clin'),
]


# ============================================================
# BUILD SHEET
# ============================================================
# Section 1: Header (rows 1-5)
ws['A1'] = "MCMI-III COMPLETE SCORING TOOL"
ws['A1'].font = TITLE_FONT
ws['A2'] = "Enter 1=TRUE or 0=FALSE. All scoring automatic."
ws['A2'].font = Font(italic=True, size=9)
ws['A4'] = "Patient:"; ws['B4'].fill = INPUT_FILL
ws['D4'] = "Date:"; ws['E4'].fill = INPUT_FILL
ws['G4'] = "Age:"; ws['H4'].fill = INPUT_FILL
ws['A5'] = "Inpatient(Y/N):"; ws['B5'].fill = INPUT_FILL
ws['D5'] = "Axis I wks:"; ws['E5'].fill = INPUT_FILL
ws['G5'] = "Gender:"; ws['H5'].fill = INPUT_FILL

# Section 2: Item responses (rows 7-43)
ws['A7'] = "ITEM RESPONSES (1=True, 0=False)"
ws['A7'].font = SEC_FONT; ws['A7'].fill = SEC_FILL
IR = 8
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
    ci = (item_num - 1) // 35
    offset = (item_num - 1) % 35
    rc = item_cols[ci][1]
    row = IR + 1 + offset
    return f"${rc}${row}"


# Section 3: Scoring table (row 45+)
SC_ROW = 45
ws[f'A{SC_ROW}'] = "SCORING TABLE"
ws[f'A{SC_ROW}'].font = SEC_FONT; ws[f'A{SC_ROW}'].fill = SEC_FILL
HDR_ROW = SC_ROW + 1
cols = ['Scale','Name','Cat','Raw','BR','Disc','A/D','Inp','Den','FINAL BR','Signif']
for ci, h in enumerate(cols):
    c = ws.cell(row=HDR_ROW, column=ci+1)
    c.value = h; c.font = HDR_FONT; c.fill = HDR_FILL
    c.border = THIN; c.alignment = Alignment(horizontal='center', wrap_text=True)

DATA_ROW = HDR_ROW + 1  # row 47
scale_rows = {}

for idx, (code, name, cat) in enumerate(SCALES):
    row = DATA_ROW + idx
    scale_rows[code] = row
    ws.cell(row=row, column=1, value=code).border = THIN
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=2, value=name).border = THIN
    ws.cell(row=row, column=3, value=cat).border = THIN

    # RAW SCORE (Col D)
    if code == 'X':
        # Scale X (Disclosure) = Sum of raw scores from scales 1 through 8B
        # Scale 5 raw is multiplied by 2/3 before adding
        # We'll reference D{row} for each of scales 1,2A,2B,3,4,6A,6B,7,8A,8B
        # and for Scale 5 we use ROUND(D{scale5_row}*2/3,0)
        # NOTE: scale_rows won't be fully populated yet, so we calculate row numbers
        # Scales order: X=47, Y=48, Z=49, 1=50, 2A=51, 2B=52, 3=53, 4=54, 5=55,
        # 6A=56, 6B=57, 7=58, 8A=59, 8B=60
        x_parts = []
        for sc in ['1','2A','2B','3','4','6A','6B','7','8A','8B']:
            sc_row = DATA_ROW + [s[0] for s in SCALES].index(sc)
            x_parts.append(f"D{sc_row}")
        # Scale 5 with 2/3 multiplier
        sc5_row = DATA_ROW + [s[0] for s in SCALES].index('5')
        x_parts.append(f"ROUND(D{sc5_row}*2/3,0)")
        raw_f = "=" + "+".join(x_parts)
    elif code == '5':
        # Scale 5 raw score is calculated normally from items
        # (the 2/3 multiplication only applies when contributing to Scale X)
        key = KEYS[code]
        parts = []
        for it in key['proto']:
            parts.append(f"2*{cell_ref(it)}")
        for it in key['nonproto']:
            parts.append(cell_ref(it))
        for it in key.get('false', []):
            parts.append(f"(1-{cell_ref(it)})")
        raw_f = "=" + "+".join(parts) if parts else "=0"
    elif code in ('Y', 'Z', 'V'):
        key = KEYS[code]
        parts = []
        for it in key.get('true', []):
            parts.append(cell_ref(it))
        for it in key.get('false', []):
            parts.append(f"(1-{cell_ref(it)})")
        raw_f = "=" + "+".join(parts) if parts else "=0"
    else:
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
        raw_f = "=" + "+".join(parts) if parts else "=0"

    ws.cell(row=row, column=4, value=raw_f)
    ws.cell(row=row, column=4).fill = RESULT_FILL
    ws.cell(row=row, column=4).border = THIN
    ws.cell(row=row, column=4).font = Font(bold=True)
    ws.cell(row=row, column=4).alignment = Alignment(horizontal='center')

    # BR (Col E) - placeholder, filled after lookup tables
    ws.cell(row=row, column=5).border = THIN
    ws.cell(row=row, column=5).fill = RESULT_FILL
    ws.cell(row=row, column=5).alignment = Alignment(horizontal='center')


# ============================================================
# EMBED LOOKUP TABLES IN HIDDEN COLUMNS
# ============================================================
LC = 16  # Col P
br_scales = ['1','2A','2B','3','4','5','6A','6B','7','8A','8B',
             'S','C','P','A','H','N','D','B','T','R','SS','CC','PP','Y','Z']
ws.cell(row=1, column=LC, value="Raw")
for si, sc in enumerate(br_scales):
    ws.cell(row=1, column=LC+1+si, value=sc)
max_br_len = max(len(BR[sc]) for sc in br_scales)
for raw_idx in range(max_br_len):
    ws.cell(row=2+raw_idx, column=LC, value=raw_idx)
    for si, sc in enumerate(br_scales):
        br_list = BR[sc]
        val = br_list[raw_idx] if raw_idx < len(br_list) else br_list[-1]
        ws.cell(row=2+raw_idx, column=LC+1+si, value=val)

XC = LC + len(br_scales) + 2
ws.cell(row=1, column=XC, value="X_Raw")
ws.cell(row=1, column=XC+1, value="X_BR")
for raw in range(34, 179):
    ws.cell(row=2+(raw-34), column=XC, value=raw)
    ws.cell(row=2+(raw-34), column=XC+1, value=BR_X.get(raw, 100))

DC = XC + 3
ws.cell(row=1, column=DC, value="D_Raw")
ws.cell(row=1, column=DC+1, value="D_18B")
ws.cell(row=1, column=DC+2, value="D_SPP")
for raw in range(200):
    a18, asp = DISC.get(raw, (0,0))
    ws.cell(row=2+raw, column=DC, value=raw)
    ws.cell(row=2+raw, column=DC+1, value=a18)
    ws.cell(row=2+raw, column=DC+2, value=asp)

AC = DC + 4
ws.cell(row=1, column=AC, value="AD_v")
ws.cell(row=1, column=AC+1, value="AD_2AS")
for v in range(81):
    ws.cell(row=2+v, column=AC, value=v)
    ws.cell(row=2+v, column=AC+1, value=AD_2AS[v])

BC_COL = AC + 3
ws.cell(row=1, column=BC_COL, value="AD_v2")
ws.cell(row=1, column=BC_COL+1, value="AD_2B8B")
for v in range(81):
    ws.cell(row=2+v, column=BC_COL, value=v)
    ws.cell(row=2+v, column=BC_COL+1, value=AD_2B8BC[v])

# Column letters for formulas
RAW_L = get_column_letter(LC)
X_RAW_L = get_column_letter(XC)
X_BR_L = get_column_letter(XC+1)
D_RAW_L = get_column_letter(DC)
D_18B_L = get_column_letter(DC+1)
D_SPP_L = get_column_letter(DC+2)
AC_L = get_column_letter(AC)
AC_ADJ_L = get_column_letter(AC+1)
BC_L = get_column_letter(BC_COL)
BC_ADJ_L = get_column_letter(BC_COL+1)

scale_br_l = {}
for si, sc in enumerate(br_scales):
    scale_br_l[sc] = get_column_letter(LC+1+si)


# ============================================================
# NOW FILL IN BR + ADJUSTMENT FORMULAS (Cols E-K)
# ============================================================
PERS_18B = ['1','2A','2B','3','4','5','6A','6B','7','8A','8B']
SPP = ['S','C','P','A','H','N','D','B','T','R','SS','CC','PP']
AD_2AS_SC = ['2A','S']
AD_2B_SC = ['2B','8B','C']

x_raw = f"D{scale_rows['X']}"
a_row = scale_rows['A']
d_row = scale_rows['D']

for idx, (code, name, cat) in enumerate(SCALES):
    row = DATA_ROW + idx

    # Col E: BR Score
    if code == 'X':
        br_f = (f'=IF(D{row}<34,0,IF(D{row}>178,100,'
                f'VLOOKUP(D{row},${X_RAW_L}:${X_BR_L},2,TRUE)))')
    elif code in scale_br_l:
        br_col_l = scale_br_l[code]
        br_f = (f'=VLOOKUP(MIN(D{row},{max_br_len-1}),'
                f'${RAW_L}:${br_col_l},'
                f'COLUMN(${br_col_l}$1)-COLUMN(${RAW_L}$1)+1,TRUE)')
    else:
        br_f = "=0"
    ws.cell(row=row, column=5, value=br_f)

    # Col F: Disclosure Adjustment
    if code in ['X','Y','Z']:
        ws.cell(row=row, column=6, value=0)
        ws.cell(row=row, column=6).fill = GREY_FILL
    elif code in PERS_18B:
        ws.cell(row=row, column=6, value=(
            f'=VLOOKUP(MIN(MAX({x_raw},0),199),${D_RAW_L}:${D_18B_L},2,TRUE)'))
    else:
        ws.cell(row=row, column=6, value=(
            f'=VLOOKUP(MIN(MAX({x_raw},0),199),${D_RAW_L}:${D_SPP_L},3,TRUE)'))
    ws.cell(row=row, column=6).border = THIN
    ws.cell(row=row, column=6).alignment = Alignment(horizontal='center')

    # Col G: A/D Adjustment
    a_adj = f"(E{a_row}+F{a_row})"
    d_adj = f"(E{d_row}+F{d_row})"
    if code in AD_2AS_SC:
        ws.cell(row=row, column=7, value=(
            f'=IF(AND({a_adj}<75,{d_adj}<75),0,'
            f'VLOOKUP(MIN(MAX(IF({a_adj}>=75,{a_adj}-75,0)'
            f'+IF({d_adj}>=75,{d_adj}-75,0),0),80),'
            f'${AC_L}:${AC_ADJ_L},2,TRUE))'))
    elif code in AD_2B_SC:
        ws.cell(row=row, column=7, value=(
            f'=IF(AND({a_adj}<75,{d_adj}<75),0,'
            f'VLOOKUP(MIN(MAX(IF({a_adj}>=75,{a_adj}-75,0)'
            f'+IF({d_adj}>=75,{d_adj}-75,0),0),80),'
            f'${BC_L}:${BC_ADJ_L},2,TRUE))'))
    else:
        ws.cell(row=row, column=7, value=0)
        ws.cell(row=row, column=7).fill = GREY_FILL
    ws.cell(row=row, column=7).border = THIN
    ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')

    # Col H: Inpatient
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

    # Col I: Denial/Complaint (manual)
    ws.cell(row=row, column=9, value=0)
    ws.cell(row=row, column=9).fill = INPUT_FILL if code in PERS_18B else GREY_FILL
    ws.cell(row=row, column=9).border = THIN
    ws.cell(row=row, column=9).alignment = Alignment(horizontal='center')

    # Col J: Final BR
    ws.cell(row=row, column=10, value=f'=MIN(MAX(E{row}+F{row}+G{row}+H{row}+I{row},0),115)')
    ws.cell(row=row, column=10).border = THIN
    ws.cell(row=row, column=10).font = Font(bold=True, size=11)
    ws.cell(row=row, column=10).alignment = Alignment(horizontal='center')

    # Col K: Significance
    ws.cell(row=row, column=11, value=(
        f'=IF(J{row}>=85,"PATHOLOGICAL",IF(J{row}>=75,"PRESENT",'
        f'IF(J{row}>=60,"Suggestive","Normal")))'))
    ws.cell(row=row, column=11).border = THIN
    ws.cell(row=row, column=11).alignment = Alignment(horizontal='center')


# ============================================================
# FINAL: Conditional formatting, validity, hide cols, save
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
ws.cell(row=vr+3, column=1, value="STATUS:").font = Font(bold=True)
ws.cell(row=vr+3, column=2, value=(
    f'=IF(OR(B{vr+1}>1,B{vr+2}<34,B{vr+2}>178),"INVALID","VALID")'))
ws.cell(row=vr+3, column=2).font = Font(bold=True, size=12)

# Scale W Inconsistency
wr = vr + 5
ws.cell(row=wr, column=1, value="W INCONSISTENCY").font = SEC_FONT
ws.cell(row=wr, column=1).fill = SEC_FILL
ws.cell(row=wr+1, column=1, value="W Score:")
# Build W formula: each pair scores 1 if the condition is met
# Format: (item_a, dir_a, item_b, dir_b)
# Score 1 point if: item_a answered dir_a AND item_b answered dir_b
# e.g. (1,'T',4,'F') = item1=1 AND item4=0
w_parts = []
for item_a, dir_a, item_b, dir_b in W_PAIRS:
    cond_a = f"{cell_ref(item_a)}=1" if dir_a == 'T' else f"{cell_ref(item_a)}=0"
    cond_b = f"{cell_ref(item_b)}=1" if dir_b == 'T' else f"{cell_ref(item_b)}=0"
    w_parts.append(f"IF(AND({cond_a},{cond_b}),1,0)")

# Split into chunks to avoid formula length limit
chunk_size = 22
w_formula_parts = []
for i in range(0, len(w_parts), chunk_size):
    chunk = w_parts[i:i+chunk_size]
    w_formula_parts.append("+".join(chunk))
w_formula = "=" + "+".join(w_formula_parts)

ws.cell(row=wr+1, column=2, value=w_formula)
ws.cell(row=wr+1, column=2).fill = RESULT_FILL
ws.cell(row=wr+1, column=3, value=f'=IF(B{wr+1}>12,"INVALID (random?)",IF(B{wr+1}>8,"CAUTION","OK"))')

# Instructions
ir2 = wr + 3
ws.cell(row=ir2, column=1, value="BR: 85-115=PATHOLOGICAL | 75-84=PRESENT | 60-74=Suggestive | 0-59=Normal")
ws.cell(row=ir2+2, column=1, value="HOW TO USE:").font = SEC_FONT
ws.cell(row=ir2+3, column=1, value="1. Enter 1(TRUE) or 0(FALSE) for all 175 items")
ws.cell(row=ir2+4, column=1, value="2. Raw + BR scores calculate automatically")
ws.cell(row=ir2+5, column=1, value="3. Adjustments auto-calculate")
ws.cell(row=ir2+6, column=1, value="4. Only manual step: Col I Denial/Complaint (enter 8 if highest 1-8B is 4,5,7)")

# Column widths
for col, w in {'A':6,'B':15,'C':11,'D':7,'E':6,'F':6,'G':6,'H':5,'I':5,'J':8,'K':14}.items():
    ws.column_dimensions[col].width = w

# Hide lookup columns
for col_num in range(LC, BC_COL + 2):
    ws.column_dimensions[get_column_letter(col_num)].hidden = True

# SAVE
out = "/projects/sandbox/Dango-kiro/MCMI-III_Scoring_Tool_V2.xlsx"
wb.save(out)
import os
print(f"SUCCESS! {out} ({os.path.getsize(out)} bytes)")
print(f"Scales: {len(SCALES)}, Rows: {DATA_ROW}-{last_row}")
print(f"Lookup: {get_column_letter(LC)}-{get_column_letter(BC_COL+1)} (hidden)")
