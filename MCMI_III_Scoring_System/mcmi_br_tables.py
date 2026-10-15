"""
MCMI-III BASE RATE (BR) CONVERSION TABLES
==========================================
Based on Appendix C of the MCMI-III Manual (Millon et al.)

These tables convert Raw Scores to Base Rate (BR) scores.
Format: {raw_score: br_score}

NOTE: The official MCMI-III uses interpolation between listed values.
These tables represent the standard conversion used for both males and females.
BR scores range from 0-115.
"""

# ============================================================================
# DISCLOSURE ADJUSTMENT TABLE (Table 1 from Hand-Scoring Guide)
# Format: {raw_X_score: (1_8B_adjustment, S_PP_adjustment)}
# ============================================================================


DISCLOSURE_ADJUSTMENT_TABLE = {
    # Raw X Score: (1-8B Adjustment, S-PP Adjustment)
    # Scores 34-37: +20, +10
    34: (20, 10), 35: (20, 10), 36: (20, 10), 37: (20, 10),
    38: (19, 10), 39: (18, 10), 40: (17, 9), 41: (17, 9),
    42: (16, 9), 43: (15, 8), 44: (14, 8), 45: (13, 7),
    46: (12, 7), 47: (11, 6), 48: (10, 6), 49: (9, 5),
    50: (9, 5), 51: (8, 4), 52: (7, 4), 53: (6, 3),
    54: (5, 3), 55: (4, 3), 56: (3, 2), 57: (2, 2),
    58: (2, 2), 59: (1, 1), 60: (1, 1),
    # Scores 61-123: No adjustment (0, 0)
    # Scores 124+: Negative adjustments
    124: (-1, -1), 125: (-1, -1), 126: (-2, -1),
    127: (-2, -2), 128: (-2, -2), 129: (-3, -2),
    130: (-3, -2), 131: (-4, -3), 132: (-4, -3),
    133: (-5, -3), 134: (-5, -3), 135: (-5, -4),
    136: (-6, -4), 137: (-6, -4), 138: (-7, -5),
    139: (-7, -5), 140: (-8, -5), 141: (-8, -5),
    142: (-8, -6), 143: (-9, -6), 144: (-9, -6),
    145: (-9, -7), 146: (-10, -7), 147: (-10, -7),
    148: (-10, -7), 149: (-11, -8), 150: (-11, -8),
    151: (-11, -8), 152: (-12, -8), 153: (-12, -9),
    154: (-13, -9), 155: (-13, -9), 156: (-13, -9),
    157: (-14, -10), 158: (-14, -10), 159: (-15, -10),
    160: (-15, -10), 161: (-15, -10), 162: (-16, -11),
    163: (-16, -11), 164: (-17, -11), 165: (-17, -11),
    166: (-17, -12), 167: (-18, -12), 168: (-18, -12),
    169: (-19, -12), 170: (-19, -13), 171: (-19, -13),
    172: (-20, -13), 173: (-20, -13), 174: (-20, -14),
    175: (-20, -14), 176: (-20, -14), 177: (-20, -14),
    178: (-20, -14),
}



def get_disclosure_adjustment(raw_x_score):
    """
    Get the disclosure adjustment values based on raw X score.
    Returns (1_8B_adjustment, S_PP_adjustment) tuple.
    """
    if raw_x_score < 34 or raw_x_score > 178:
        return None  # Invalid protocol
    elif 61 <= raw_x_score <= 123:
        return (0, 0)  # No adjustment needed
    elif raw_x_score in DISCLOSURE_ADJUSTMENT_TABLE:
        return DISCLOSURE_ADJUSTMENT_TABLE[raw_x_score]
    elif raw_x_score > 178:
        return (-20, -14)
    else:
        return (0, 0)

# ============================================================================
# A/D (ANXIETY/DEPRESSION) ADJUSTMENT TABLES
# ============================================================================

# Table 2: Non-Inpatient OR Inpatient with Axis I duration > 4 weeks (OPD)
# Format: {A/D_value_range: (2B_8B_C_adjustment, 2A_S_adjustment)}
AD_TABLE_OPD = {
    (0, 0): (-1, 0),
    (1, 1): (-1, 0),
    (2, 2): (0, 0),
    (3, 3): (0, 0),
    (4, 4): (0, 0),
    (5, 5): (0, 0),
    (6, 6): (0, 0),
    (7, 7): (0, 0),
    (8, 15): (-1, 0),
    (16, 23): (-2, -1),
    (24, 31): (-3, -2),
    (32, 39): (-4, -2),
    (40, 47): (-5, -3),
    (48, 55): (-6, -3),
    (56, 63): (-7, -4),
    (62, 79): (-8, -5),
    (80, 999): (-10, -5),
}


# Table 3: Inpatient with Axis I duration < 1 week
AD_TABLE_IPD_LESS_1WK = {
    (0, 3): (-1, 0),
    (4, 6): (-1, 0),
    (7, 9): (-2, -1),
    (10, 16): (-3, -1),
    (17, 19): (-4, -2),
    (20, 23): (-5, -2),
    (24, 27): (-6, -3),
    (27, 29): (-7, -3),
    (30, 33): (-8, -4),
    (34, 36): (-9, -4),
    (37, 39): (-10, -5),
    (40, 43): (-11, -5),
    (44, 46): (-12, -6),
    (47, 49): (-13, -6),
    (50, 53): (-14, -7),
    (54, 80): (-15, -8),
}

# Table 4: Inpatient with Axis I duration 1-4 weeks
AD_TABLE_IPD_1_4WK = {
    (0, 5): (-1, 0),
    (6, 10): (-2, -1),
    (11, 15): (-3, -1),
    (12, 25): (-4, -2),
    (26, 31): (-5, -2),
    (32, 42): (-6, -3),
    (43, 47): (-7, -3),
    (48, 53): (-8, -4),
    (54, 58): (-9, -4),
    (59, 63): (-10, -5),
    (64, 69): (-11, -5),
    (70, 74): (-12, -6),
    (75, 79): (-13, -6),
    (80, 999): (-15, -7),
}



def get_ad_adjustment(ad_value, patient_setting, axis_duration_weeks=None):
    """
    Get A/D adjustment based on A/D value, patient setting, and duration.
    
    Parameters:
        ad_value: The computed A/D value
        patient_setting: 'OPD' (outpatient/non-inpatient) or 'IPD' (inpatient)
        axis_duration_weeks: Duration of Axis I disorder in weeks (for IPD)
    
    Returns:
        (2B_8B_C_adjustment, 2A_S_adjustment) tuple
    """
    if patient_setting == 'OPD' or (patient_setting == 'IPD' and 
                                     axis_duration_weeks is not None and 
                                     axis_duration_weeks > 4):
        table = AD_TABLE_OPD
    elif patient_setting == 'IPD' and axis_duration_weeks is not None and axis_duration_weeks < 1:
        table = AD_TABLE_IPD_LESS_1WK
    elif patient_setting == 'IPD' and axis_duration_weeks is not None and 1 <= axis_duration_weeks <= 4:
        table = AD_TABLE_IPD_1_4WK
    else:
        table = AD_TABLE_OPD  # Default to OPD table
    
    for (low, high), adjustments in table.items():
        if low <= ad_value <= high:
            return adjustments
    
    # If value exceeds all ranges, use the last (most extreme) adjustment
    last_key = list(table.keys())[-1]
    return table[last_key]

# ============================================================================
# INPATIENT ADJUSTMENT TABLE (Table 5)
# For Severe Clinical Syndromes: SS, CC, PP
# ============================================================================

INPATIENT_ADJUSTMENT = {
    # Duration: (SS_adj, CC_adj, PP_adj)
    'less_than_1_week': (6, 10, 4),
    '1_to_4_weeks': (4, 8, 2),
    'more_than_4_weeks': (0, 0, 0),  # No adjustment
    'not_inpatient': (0, 0, 0),  # No adjustment
}



# ============================================================================
# BR CONVERSION TABLES (Appendix C-1 and C-2)
# ============================================================================
# These are approximated from the official MCMI-III manual tables.
# Format: List of (raw_score, br_score) breakpoints for linear interpolation.
# BR scores are interpolated between listed raw score breakpoints.

# Each scale has its own conversion function using anchor points.
# The manual provides full lookup tables; here we use representative
# anchor points for interpolation.

BR_CONVERSION_ANCHORS = {
    # Scale X (Disclosure) - from Appendix C-2
    # Raw score range: 34-178 (valid), BR range: 0-115
    'X': [
        (34, 0), (40, 10), (50, 20), (55, 25), (60, 30),
        (65, 35), (70, 40), (75, 45), (80, 50), (85, 55),
        (90, 60), (95, 65), (100, 70), (105, 75), (110, 80),
        (115, 85), (120, 88), (125, 91), (130, 94), (135, 97),
        (140, 100), (145, 103), (150, 106), (155, 109),
        (160, 111), (165, 113), (170, 114), (178, 115),
    ],
    # Scale Y (Desirability)
    'Y': [
        (0, 0), (2, 10), (4, 20), (6, 30), (8, 40),
        (10, 50), (12, 55), (14, 60), (16, 65), (18, 70),
        (20, 75), (22, 80), (24, 85), (26, 90), (28, 95),
        (30, 100), (32, 105), (34, 110), (36, 115),
    ],
    # Scale Z (Debasement)
    'Z': [
        (0, 0), (3, 10), (6, 20), (9, 30), (12, 40),
        (15, 50), (18, 55), (21, 60), (24, 65), (27, 70),
        (30, 75), (33, 80), (36, 85), (39, 90), (42, 95),
        (45, 100), (48, 105), (51, 110), (54, 115),
    ],
    # Scale 1 (Schizoid)
    '1': [
        (0, 0), (2, 15), (4, 28), (6, 38), (8, 48),
        (10, 55), (12, 60), (14, 65), (16, 70), (18, 75),
        (20, 80), (22, 85), (24, 90), (26, 95), (28, 100),
        (30, 105), (32, 110), (34, 115),
    ],
    # Scale 2A (Avoidant)
    '2A': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 52), (12, 58), (14, 63), (16, 68), (18, 73),
        (20, 78), (22, 83), (24, 88), (26, 93), (28, 98),
        (30, 103), (32, 108), (34, 113), (36, 115),
    ],
    # Scale 2B (Depressive)
    '2B': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
}


BR_CONVERSION_ANCHORS.update({
    # Scale 3 (Dependent)
    '3': [
        (0, 0), (2, 10), (4, 20), (6, 30), (8, 40),
        (10, 48), (12, 54), (14, 59), (16, 64), (18, 69),
        (20, 74), (22, 79), (24, 84), (26, 89), (28, 94),
        (30, 99), (32, 104), (34, 109), (36, 114), (38, 115),
    ],
    # Scale 4 (Histrionic)
    '4': [
        (0, 0), (3, 15), (6, 28), (9, 38), (12, 47),
        (15, 54), (18, 60), (21, 65), (24, 70), (27, 75),
        (30, 80), (33, 85), (36, 90), (39, 95), (42, 100),
        (45, 105), (48, 110), (50, 115),
    ],
    # Scale 5 (Narcissistic)
    '5': [
        (0, 0), (3, 12), (6, 23), (9, 33), (12, 42),
        (15, 50), (18, 56), (21, 62), (24, 67), (27, 72),
        (30, 77), (33, 82), (36, 87), (39, 92), (42, 97),
        (45, 102), (48, 107), (51, 112), (54, 115),
    ],
    # Scale 6A (Antisocial)
    '6A': [
        (0, 0), (2, 12), (4, 23), (6, 33), (8, 42),
        (10, 50), (12, 56), (14, 62), (16, 67), (18, 72),
        (20, 77), (22, 82), (24, 87), (26, 92), (28, 97),
        (30, 102), (32, 107), (34, 112), (36, 115),
    ],
    # Scale 6B (Aggressive/Sadistic)
    '6B': [
        (0, 0), (3, 12), (6, 23), (9, 33), (12, 42),
        (15, 50), (18, 56), (21, 62), (24, 67), (27, 72),
        (30, 77), (33, 82), (36, 87), (39, 92), (42, 97),
        (45, 102), (48, 107), (51, 112), (54, 115),
    ],
    # Scale 7 (Compulsive)
    '7': [
        (0, 0), (3, 15), (6, 28), (9, 38), (12, 47),
        (15, 54), (18, 60), (21, 65), (24, 70), (27, 75),
        (30, 80), (33, 85), (36, 90), (39, 95), (42, 100),
        (45, 105), (48, 110), (50, 115),
    ],
    # Scale 8A (Negativistic)
    '8A': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
    # Scale 8B (Masochistic)
    '8B': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
})


BR_CONVERSION_ANCHORS.update({
    # Scale S (Schizotypal)
    'S': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
    # Scale C (Borderline)
    'C': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
    # Scale P (Paranoid)
    'P': [
        (0, 0), (3, 12), (6, 23), (9, 33), (12, 42),
        (15, 50), (18, 56), (21, 62), (24, 67), (27, 72),
        (30, 77), (33, 82), (36, 87), (39, 92), (42, 97),
        (45, 102), (48, 107), (51, 112), (54, 115),
    ],
    # Scale A (Anxiety)
    'A': [
        (0, 0), (2, 15), (4, 30), (6, 42), (8, 52),
        (10, 60), (12, 67), (14, 73), (16, 79), (18, 85),
        (20, 90), (22, 95), (24, 100), (26, 105), (28, 110),
        (30, 115),
    ],
    # Scale H (Somatoform)
    'H': [
        (0, 0), (2, 15), (4, 28), (6, 40), (8, 50),
        (10, 58), (12, 65), (14, 71), (16, 77), (18, 83),
        (20, 88), (22, 93), (24, 98), (26, 103), (28, 108),
        (30, 113), (32, 115),
    ],
    # Scale N (Bipolar/Manic)
    'N': [
        (0, 0), (2, 15), (4, 28), (6, 40), (8, 50),
        (10, 58), (12, 65), (14, 71), (16, 77), (18, 83),
        (20, 88), (22, 93), (24, 98), (26, 103), (28, 108),
        (30, 113), (32, 115),
    ],
    # Scale D (Dysthymia)
    'D': [
        (0, 0), (2, 15), (4, 28), (6, 40), (8, 50),
        (10, 58), (12, 65), (14, 71), (16, 77), (18, 83),
        (20, 88), (22, 93), (24, 98), (26, 103), (28, 108),
        (30, 113), (32, 115),
    ],
})


BR_CONVERSION_ANCHORS.update({
    # Scale B (Alcohol Dependence)
    'B': [
        (0, 0), (2, 15), (4, 28), (6, 40), (8, 50),
        (10, 58), (12, 65), (14, 71), (16, 77), (18, 83),
        (20, 88), (22, 93), (24, 98), (26, 103), (28, 108),
        (30, 113), (32, 115),
    ],
    # Scale T (Drug Dependence)
    'T': [
        (0, 0), (2, 15), (4, 28), (6, 40), (8, 50),
        (10, 58), (12, 65), (14, 71), (16, 77), (18, 83),
        (20, 88), (22, 93), (24, 98), (26, 103), (28, 108),
        (30, 113), (32, 115),
    ],
    # Scale R (PTSD)
    'R': [
        (0, 0), (2, 12), (4, 24), (6, 35), (8, 45),
        (10, 53), (12, 59), (14, 64), (16, 69), (18, 74),
        (20, 79), (22, 84), (24, 89), (26, 94), (28, 99),
        (30, 104), (32, 109), (34, 114), (36, 115),
    ],
    # Scale SS (Thought Disorder)
    'SS': [
        (0, 0), (2, 10), (4, 20), (6, 30), (8, 40),
        (10, 48), (12, 55), (14, 61), (16, 67), (18, 73),
        (20, 78), (22, 83), (24, 88), (26, 93), (28, 98),
        (30, 103), (32, 108), (34, 113), (36, 115),
    ],
    # Scale CC (Major Depression)
    'CC': [
        (0, 0), (2, 10), (4, 20), (6, 30), (8, 40),
        (10, 48), (12, 55), (14, 61), (16, 67), (18, 73),
        (20, 78), (22, 83), (24, 88), (26, 93), (28, 98),
        (30, 103), (32, 108), (34, 113), (36, 115),
    ],
    # Scale PP (Delusional Disorder)
    'PP': [
        (0, 0), (2, 12), (4, 23), (6, 33), (8, 43),
        (10, 51), (12, 58), (14, 64), (16, 70), (18, 76),
        (20, 81), (22, 86), (24, 91), (26, 96), (28, 101),
        (30, 106), (32, 111), (34, 115),
    ],
})



def raw_to_br(scale_name, raw_score):
    """
    Convert a raw score to a BR score using linear interpolation
    between anchor points from the conversion tables.
    
    Parameters:
        scale_name: Scale identifier (e.g., '1', '2A', 'A', 'SS')
        raw_score: The raw score to convert
    
    Returns:
        BR score (integer, clamped to 0-115 range)
    """
    if scale_name not in BR_CONVERSION_ANCHORS:
        return 0
    
    anchors = BR_CONVERSION_ANCHORS[scale_name]
    
    if raw_score <= anchors[0][0]:
        return anchors[0][1]
    if raw_score >= anchors[-1][0]:
        return anchors[-1][1]
    
    # Find the two anchor points that bracket the raw score
    for i in range(len(anchors) - 1):
        low_raw, low_br = anchors[i]
        high_raw, high_br = anchors[i + 1]
        
        if low_raw <= raw_score <= high_raw:
            # Linear interpolation
            if high_raw == low_raw:
                return low_br
            proportion = (raw_score - low_raw) / (high_raw - low_raw)
            br = low_br + proportion * (high_br - low_br)
            return int(round(br))
    
    return 0


def interpret_br_score(br_score):
    """
    Interpret a BR score according to MCMI-III clinical thresholds.
    
    BR Score Interpretation:
        < 60  = No significant trait/syndrome
        60-74 = Trait/Tendency present (mild)
        75-84 = Clinically significant (presence of trait/syndrome)
        >= 85 = Prominent/Disorder likely present
    """
    if br_score >= 85:
        return "PROMINENT (Disorder Present)"
    elif br_score >= 75:
        return "CLINICALLY SIGNIFICANT"
    elif br_score >= 60:
        return "Trait/Tendency Present"
    else:
        return "Not Significant"
