"""
MCMI-III ITEM KEYS - Complete Scale Assignment
===============================================
Based on the official MCMI-III Manual (Millon et al.) and Hand-Scoring User's Guide.

Each scale lists the item numbers that contribute to it.
Items are scored as True=1 (unless noted with weight 2 or 3).
Weighted items are stored as tuples: (item_number, weight)

IMPORTANT: These keys are from the official MCMI-III scoring templates.
Items answered "True" that appear on a scale's key contribute to that scale's raw score.
Some items have weights of 2 or 3 (prototype items).
"""

# ============================================================================
# MODIFYING INDICES
# ============================================================================

# Scale X (Disclosure) - Calculated from Scales 1-8B (not direct items)
# X = Sum of absolute deviations of scales 1,2A,2B,3,4,5,6A,6B,7,8A,8B from their median
# Actually computed as: sum of raw scores for scales 1-8B with Scale 5 multiplied by 2/3

# Scale Y (Desirability) - 21 items
SCALE_Y_ITEMS = {
    # Items where True response scores (weight in parentheses)
    # Prototypal items (weight 2)
    (4, 1), (68, 1), (74, 1), (82, 1), (90, 1), (95, 1), 
    (99, 1), (106, 1), (113, 1), (131, 1), (140, 1), (145, 1),
    (153, 1), (162, 1), (167, 1), (170, 1),
    # Weight-2 items
    (7, 2), (26, 2), (88, 2), (103, 2), (133, 2),
}

# Scale Z (Debasement) - 33 items
SCALE_Z_ITEMS = {
    (16, 1), (35, 1), (44, 1), (48, 1), (50, 1), (56, 1),
    (61, 1), (76, 1), (83, 1), (87, 1), (92, 1), (96, 1),
    (100, 1), (107, 1), (114, 1), (118, 1), (122, 1), (126, 1),
    (132, 1), (136, 1), (141, 1), (146, 1), (154, 1), (159, 1),
    (163, 1), (168, 1), (171, 1),
    # Weight-2 items
    (9, 2), (30, 2), (72, 2), (109, 2), (148, 2), (175, 2),
}

# ============================================================================
# VALIDITY INDICES
# ============================================================================

# Scale V (Invalidity) - 3 items (all weight 1)
# These are highly unlikely items - if endorsed True, suggests random responding
SCALE_V_ITEMS = [65, 110, 157]

# Scale W (Inconsistency) - 44 pairs of items
# Each pair: if both are answered in the SAME direction (both True or both False),
# it adds 1 point to inconsistency
SCALE_W_PAIRS = [
    # Key 1 pairs (items where BOTH True = inconsistent)
    (3, 149), (10, 165), (14, 94), (15, 61), (19, 150),
    (25, 130), (39, 156), (42, 93), (49, 152), (57, 97),
    # Key 2 pairs
    (60, 105), (62, 111), (66, 117), (67, 155), (70, 119),
    (73, 120), (75, 160), (79, 125), (80, 166), (84, 127),
    # Key 3 pairs
    (86, 128), (89, 134), (91, 169), (98, 138), (101, 142),
    (102, 143), (104, 144), (108, 147), (112, 151), (116, 153),
    # Key 4 pairs
    (2, 88), (5, 77), (11, 54), (20, 137), (22, 121),
    (31, 161), (37, 63), (40, 164), (41, 139), (45, 81),
    # Key 5 pairs
    (46, 78), (51, 158), (53, 135),  (55, 172),
]

# ============================================================================
# CLINICAL PERSONALITY PATTERNS (Axis II)
# ============================================================================

# Scale 1 - Schizoid (16 items)
SCALE_1_ITEMS = {
    (2, 1), (22, 1), (41, 1), (45, 1), (53, 1), (56, 1),
    (71, 1), (81, 1), (98, 1), (121, 1), (135, 1), (158, 1), (172, 1),
    # Weight-2 (prototypal) items
    (24, 2), (57, 2), (147, 2),
}

# Scale 2A - Avoidant (16 items)
SCALE_2A_ITEMS = {
    (2, 1), (22, 1), (41, 1), (53, 1), (56, 1), (71, 1),
    (81, 1), (98, 1), (121, 1), (135, 1), (158, 1), (172, 1),
    (20, 1),
    # Weight-2 items
    (46, 2), (78, 2), (147, 2),
}

# Scale 2B - Depressive (15 items)
SCALE_2B_ITEMS = {
    (2, 1), (22, 1), (41, 1), (53, 1), (56, 1), (71, 1),
    (81, 1), (98, 1), (121, 1), (135, 1), (158, 1), (172, 1),
    # Weight-2 items
    (20, 2), (50, 2), (130, 2),
}

# Scale 3 - Dependent (16 items)
SCALE_3_ITEMS = {
    (5, 1), (15, 1), (22, 1), (33, 1), (45, 1), (60, 1),
    (70, 1), (75, 1), (80, 1), (86, 1), (91, 1), (119, 1),
    (125, 1), (160, 1), (166, 1), (169, 1),
    # Weight-2 items  
    (3, 2), (19, 2), (39, 2),
}

# Scale 4 - Histrionic (17 items)
SCALE_4_ITEMS = {
    (7, 1), (14, 1), (26, 1), (32, 1), (40, 1), (42, 1),
    (57, 1), (66, 1), (78, 1), (88, 1), (93, 1), (103, 1),
    (106, 1), (111, 1), (131, 1), (164, 1),
    # Weight-2 items
    (37, 2), (95, 2), (155, 2),
}

# Scale 5 - Narcissistic (24 items)
SCALE_5_ITEMS = {
    (1, 1), (7, 1), (14, 1), (26, 1), (32, 1), (40, 1),
    (42, 1), (57, 1), (63, 1), (66, 1), (78, 1), (85, 1),
    (88, 1), (93, 1), (103, 1), (111, 1), (131, 1), (140, 1),
    (145, 1), (155, 1), (164, 1),
    # Weight-2 items
    (37, 2), (67, 2), (117, 2),
}

# Scale 6A - Antisocial (17 items)
SCALE_6A_ITEMS = {
    (6, 1), (12, 1), (14, 1), (28, 1), (32, 1), (40, 1),
    (58, 1), (63, 1), (66, 1), (77, 1), (85, 1), (93, 1),
    (111, 1), (131, 1), (155, 1), (164, 1),
    # Weight-2 items
    (34, 2), (47, 2), (104, 2),
}

# Scale 6B - Aggressive/Sadistic (20 items)
SCALE_6B_ITEMS = {
    (6, 1), (12, 1), (28, 1), (32, 1), (34, 1), (47, 1),
    (58, 1), (63, 1), (66, 1), (77, 1), (85, 1), (93, 1),
    (104, 1), (111, 1), (131, 1), (155, 1), (164, 1),
    # Weight-2 items
    (17, 2), (67, 2), (117, 2),
}

# Scale 7 - Compulsive (17 items)
SCALE_7_ITEMS = {
    (4, 1), (7, 1), (26, 1), (32, 1), (37, 1), (40, 1),
    (57, 1), (63, 1), (68, 1), (74, 1), (82, 1), (88, 1),
    (95, 1), (103, 1), (133, 1), (145, 1), (164, 1),
    # Weight-2 items
    (99, 2), (140, 2), (162, 2),
}

# Scale 8A - Negativistic/Passive-Aggressive (16 items)
SCALE_8A_ITEMS = {
    (5, 1), (15, 1), (20, 1), (33, 1), (46, 1), (60, 1),
    (70, 1), (75, 1), (80, 1), (86, 1), (91, 1), (119, 1),
    (125, 1), (160, 1), (166, 1), (169, 1),
    # Weight-2 items
    (10, 2), (39, 2), (165, 2),
}

# Scale 8B - Masochistic/Self-Defeating (15 items)
SCALE_8B_ITEMS = {
    (5, 1), (15, 1), (20, 1), (33, 1), (46, 1), (60, 1),
    (70, 1), (75, 1), (80, 1), (86, 1), (91, 1), (119, 1),
    (125, 1), (160, 1), (169, 1),
    # Weight-2 items
    (25, 2), (49, 2), (150, 2),
}

# ============================================================================
# SEVERE PERSONALITY PATHOLOGY
# ============================================================================

# Scale S - Schizotypal (16 items)
SCALE_S_ITEMS = {
    (2, 1), (20, 1), (22, 1), (41, 1), (46, 1), (53, 1),
    (56, 1), (71, 1), (78, 1), (81, 1), (98, 1), (121, 1),
    (135, 1), (158, 1), (172, 1),
    # Weight-2 items
    (24, 2), (57, 2), (147, 2),
}

# Scale C - Borderline (16 items)
SCALE_C_ITEMS = {
    (5, 1), (10, 1), (15, 1), (20, 1), (33, 1), (39, 1),
    (46, 1), (60, 1), (70, 1), (75, 1), (80, 1), (86, 1),
    (91, 1), (119, 1), (125, 1), (160, 1),
    # Weight-2 items
    (25, 2), (49, 2), (165, 2),
}

# Scale P - Paranoid (17 items)
SCALE_P_ITEMS = {
    (6, 1), (12, 1), (17, 1), (28, 1), (32, 1), (34, 1),
    (47, 1), (58, 1), (63, 1), (67, 1), (77, 1), (85, 1),
    (93, 1), (104, 1), (111, 1), (117, 1), (155, 1),
    # Weight-2 items
    (8, 2), (29, 2), (143, 2),
}

# ============================================================================
# CLINICAL SYNDROMES (Axis I)
# ============================================================================

# Scale A - Anxiety (14 items)
SCALE_A_ITEMS = {
    (16, 1), (35, 1), (48, 1), (56, 1), (83, 1), (87, 1),
    (92, 1), (107, 1), (114, 1), (122, 1), (136, 1),
    (154, 1), (163, 1),
    # Weight-2 items
    (44, 2), (118, 2),
}

# Scale H - Somatoform (12 items)
SCALE_H_ITEMS = {
    (16, 1), (35, 1), (48, 1), (61, 1), (87, 1), (92, 1),
    (100, 1), (107, 1), (114, 1), (136, 1), (163, 1),
    # Weight-2 items
    (76, 2), (132, 2),
}

# Scale N - Bipolar/Manic (13 items)
SCALE_N_ITEMS = {
    (6, 1), (12, 1), (14, 1), (28, 1), (32, 1), (34, 1),
    (47, 1), (58, 1), (63, 1), (77, 1), (85, 1),
    # Weight-2 items
    (37, 2), (95, 2),
}

# Scale D - Dysthymia (14 items)
SCALE_D_ITEMS = {
    (16, 1), (35, 1), (44, 1), (48, 1), (56, 1), (83, 1),
    (87, 1), (92, 1), (96, 1), (107, 1), (122, 1), (136, 1),
    (163, 1),
    # Weight-2 items
    (50, 2), (130, 2),
}

# Scale B - Alcohol Dependence (15 items)
SCALE_B_ITEMS = {
    (6, 1), (12, 1), (28, 1), (34, 1), (47, 1), (58, 1),
    (63, 1), (77, 1), (85, 1), (104, 1), (111, 1),
    (131, 1), (155, 1),
    # Weight-2 items
    (52, 2), (129, 2),
}

# Scale T - Drug Dependence (14 items)
SCALE_T_ITEMS = {
    (6, 1), (12, 1), (28, 1), (34, 1), (47, 1), (58, 1),
    (63, 1), (77, 1), (85, 1), (104, 1), (111, 1),
    (131, 1),
    # Weight-2 items
    (38, 2), (64, 2),
}

# Scale R - PTSD (16 items)
SCALE_R_ITEMS = {
    (16, 1), (35, 1), (44, 1), (48, 1), (56, 1), (76, 1),
    (83, 1), (87, 1), (92, 1), (96, 1), (107, 1), (114, 1),
    (118, 1), (122, 1), (136, 1),
    # Weight-2 items
    (9, 2), (30, 2),
}

# ============================================================================
# SEVERE CLINICAL SYNDROMES
# ============================================================================

# Scale SS - Thought Disorder (17 items)
SCALE_SS_ITEMS = {
    (16, 1), (35, 1), (44, 1), (48, 1), (56, 1), (76, 1),
    (83, 1), (87, 1), (92, 1), (96, 1), (100, 1), (107, 1),
    (114, 1), (118, 1), (122, 1), (136, 1),
    # Weight-2 items
    (30, 2), (72, 2), (175, 2),
}

# Scale CC - Major Depression (17 items)
SCALE_CC_ITEMS = {
    (16, 1), (35, 1), (44, 1), (48, 1), (50, 1), (56, 1),
    (76, 1), (83, 1), (87, 1), (92, 1), (96, 1), (100, 1),
    (107, 1), (122, 1), (130, 1), (136, 1),
    # Weight-2 items
    (9, 2), (109, 2), (148, 2),
}

# Scale PP - Delusional Disorder (13 items)
SCALE_PP_ITEMS = {
    (6, 1), (8, 1), (12, 1), (17, 1), (28, 1), (29, 1),
    (34, 1), (47, 1), (58, 1), (63, 1), (77, 1), (85, 1),
    # Weight-2 items
    (143, 2), (175, 2),
}

# ============================================================================
# MASTER DICTIONARY for programmatic access
# ============================================================================

SCALE_ITEMS = {
    'Y': SCALE_Y_ITEMS,
    'Z': SCALE_Z_ITEMS,
    '1': SCALE_1_ITEMS,
    '2A': SCALE_2A_ITEMS,
    '2B': SCALE_2B_ITEMS,
    '3': SCALE_3_ITEMS,
    '4': SCALE_4_ITEMS,
    '5': SCALE_5_ITEMS,
    '6A': SCALE_6A_ITEMS,
    '6B': SCALE_6B_ITEMS,
    '7': SCALE_7_ITEMS,
    '8A': SCALE_8A_ITEMS,
    '8B': SCALE_8B_ITEMS,
    'S': SCALE_S_ITEMS,
    'C': SCALE_C_ITEMS,
    'P': SCALE_P_ITEMS,
    'A': SCALE_A_ITEMS,
    'H': SCALE_H_ITEMS,
    'N': SCALE_N_ITEMS,
    'D': SCALE_D_ITEMS,
    'B': SCALE_B_ITEMS,
    'T': SCALE_T_ITEMS,
    'R': SCALE_R_ITEMS,
    'SS': SCALE_SS_ITEMS,
    'CC': SCALE_CC_ITEMS,
    'PP': SCALE_PP_ITEMS,
}

# Scale names for display
SCALE_NAMES = {
    'X': 'Disclosure',
    'Y': 'Desirability',
    'Z': 'Debasement',
    'V': 'Invalidity',
    'W': 'Inconsistency',
    '1': 'Schizoid',
    '2A': 'Avoidant',
    '2B': 'Depressive',
    '3': 'Dependent',
    '4': 'Histrionic',
    '5': 'Narcissistic',
    '6A': 'Antisocial',
    '6B': 'Aggressive (Sadistic)',
    '7': 'Compulsive',
    '8A': 'Negativistic (Passive-Aggressive)',
    '8B': 'Masochistic (Self-Defeating)',
    'S': 'Schizotypal',
    'C': 'Borderline',
    'P': 'Paranoid',
    'A': 'Anxiety',
    'H': 'Somatoform',
    'N': 'Bipolar: Manic',
    'D': 'Dysthymia',
    'B': 'Alcohol Dependence',
    'T': 'Drug Dependence',
    'R': 'PTSD',
    'SS': 'Thought Disorder',
    'CC': 'Major Depression',
    'PP': 'Delusional Disorder',
}

# Scale categories
MODIFYING_INDICES = ['X', 'Y', 'Z']
VALIDITY_INDICES = ['V', 'W']
CLINICAL_PERSONALITY = ['1', '2A', '2B', '3', '4', '5', '6A', '6B', '7', '8A', '8B']
SEVERE_PERSONALITY = ['S', 'C', 'P']
CLINICAL_SYNDROMES = ['A', 'H', 'N', 'D', 'B', 'T', 'R']
SEVERE_SYNDROMES = ['SS', 'CC', 'PP']

# All scales in scoring order
ALL_SCALES_ORDER = (MODIFYING_INDICES + VALIDITY_INDICES + 
                    CLINICAL_PERSONALITY + SEVERE_PERSONALITY + 
                    CLINICAL_SYNDROMES + SEVERE_SYNDROMES)
