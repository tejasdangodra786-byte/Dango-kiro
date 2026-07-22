"""
MCMI-III SCORING ENGINE
========================
Complete scoring implementation following the official MCMI-III
Hand-Scoring User's Guide (Millon et al.)

This engine performs:
1. Validity checks (V, W scales)
2. Raw score calculation for all 26+ scales
3. Scale X (Disclosure) calculation
4. Initial BR score conversion
5. Disclosure adjustment (1-8B and S-PP)
6. Anxiety/Depression (A/D) adjustment
7. Inpatient adjustment (SS, CC, PP)
8. Denial/Complaint adjustment (Scales 4, 5, 7)
9. Final BR score determination
"""

import math
from mcmi_item_keys import (
    SCALE_V_ITEMS, SCALE_W_PAIRS, SCALE_ITEMS, SCALE_NAMES,
    CLINICAL_PERSONALITY, SEVERE_PERSONALITY,
    CLINICAL_SYNDROMES, SEVERE_SYNDROMES,
    MODIFYING_INDICES, ALL_SCALES_ORDER
)
from mcmi_br_tables import (
    get_disclosure_adjustment, get_ad_adjustment,
    INPATIENT_ADJUSTMENT, raw_to_br, interpret_br_score
)



class MCMIIIScorer:
    """
    Complete MCMI-III Scoring System.
    
    Usage:
        scorer = MCMIIIScorer()
        scorer.set_responses(responses_dict)  # {1: True/False, 2: True/False, ...}
        scorer.set_patient_info(age=35, setting='OPD', axis_duration_weeks=None)
        results = scorer.score()
    """
    
    def __init__(self):
        self.responses = {}  # Item responses: {item_num: True/False}
        self.age = None
        self.patient_setting = 'OPD'  # 'OPD' or 'IPD'
        self.axis_duration_weeks = None  # For inpatient: weeks of Axis I disorder
        self.gender = None  # 'M' or 'F' (both use same tables in MCMI-III)
        
        # Results storage
        self.raw_scores = {}
        self.initial_br_scores = {}
        self.adjusted_br_scores = {}
        self.final_br_scores = {}
        self.validity_status = 'NOT SCORED'
        self.validity_notes = []
        self.adjustment_log = []
    
    def set_responses(self, responses):
        """
        Set item responses.
        
        Parameters:
            responses: dict {item_number: True/False} for items 1-175
                       OR list of 175 True/False values (index 0 = item 1)
        """
        if isinstance(responses, list):
            self.responses = {i+1: bool(v) for i, v in enumerate(responses)}
        elif isinstance(responses, dict):
            self.responses = {int(k): bool(v) for k, v in responses.items()}
        else:
            raise ValueError("Responses must be dict or list")
    
    def set_patient_info(self, age, setting='OPD', axis_duration_weeks=None, gender=None):
        """
        Set patient demographic and clinical information.
        
        Parameters:
            age: Patient age (must be >= 18)
            setting: 'OPD' (outpatient/non-inpatient) or 'IPD' (inpatient)
            axis_duration_weeks: For IPD patients, duration of Axis I disorder
            gender: 'M' or 'F' (MCMI-III uses same tables for both)
        """
        self.age = age
        self.patient_setting = setting.upper()
        self.axis_duration_weeks = axis_duration_weeks
        self.gender = gender


    
    def score(self):
        """
        Execute the complete MCMI-III scoring algorithm.
        Returns a comprehensive results dictionary.
        """
        self.validity_notes = []
        self.adjustment_log = []
        
        # Step 1: Age check
        if not self._check_age():
            return self._build_results()
        
        # Step 2: Omits/Double marks check
        if not self._check_omits():
            return self._build_results()
        
        # Step 3: Scale V (Invalidity) check
        if not self._check_invalidity():
            return self._build_results()
        
        # Step 4: Scale W (Inconsistency) check
        if not self._check_inconsistency():
            return self._build_results()
        
        # Step 5: Calculate raw scores for all scales (Y, Z, 1-PP)
        self._calculate_raw_scores()
        
        # Step 6: Calculate Scale X raw score
        self._calculate_scale_x()
        
        # Step 7: Scale X raw score validity check
        if not self._check_scale_x_validity():
            return self._build_results()
        
        # Step 8: Initial BR score conversion
        self._calculate_initial_br_scores()
        
        # Step 9: Disclosure adjustment
        self._apply_disclosure_adjustment()
        
        # Step 10: Anxiety/Depression adjustment
        self._apply_ad_adjustment()
        
        # Step 11: Inpatient adjustment
        self._apply_inpatient_adjustment()
        
        # Step 12: Denial/Complaint adjustment
        self._apply_denial_complaint_adjustment()
        
        # Step 13: Finalize BR scores
        self._finalize_scores()
        
        return self._build_results()


    
    # =========================================================================
    # STEP 1: AGE CHECK
    # =========================================================================
    def _check_age(self):
        """Step 1-1: Check if age is indicated and >= 18."""
        if self.age is None:
            self.validity_status = 'INVALID'
            self.validity_notes.append("Age not indicated. Test is INVALID.")
            return False
        if self.age < 18:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Age ({self.age}) is less than 18. Test is INVALID."
            )
            return False
        return True
    
    # =========================================================================
    # STEP 2: OMITS/DOUBLE MARKS CHECK
    # =========================================================================
    def _check_omits(self):
        """Step 2-1: Count omitted items. If > 11, invalid."""
        total_items = 175
        answered = sum(1 for i in range(1, total_items + 1) if i in self.responses)
        omitted = total_items - answered
        
        self.omitted_count = omitted
        
        if omitted > 11:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Too many omitted items ({omitted}). Maximum allowed: 11. "
                "Test is INVALID."
            )
            return False
        elif omitted > 0:
            self.validity_notes.append(
                f"Note: {omitted} item(s) omitted (within acceptable range)."
            )
        return True
    
    # =========================================================================
    # STEP 3: SCALE V (INVALIDITY) CHECK
    # =========================================================================
    def _check_invalidity(self):
        """
        Step 3-1 to 3-3: Calculate Scale V score.
        V items: 65, 110, 157
        Score = count of True responses to these items.
        If V >= 2: INVALID
        If V == 1: Questionable validity
        If V == 0: Valid
        """
        v_score = sum(
            1 for item in SCALE_V_ITEMS 
            if self.responses.get(item, False)
        )
        self.raw_scores['V'] = v_score
        
        if v_score >= 2:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Scale V (Invalidity) = {v_score}. Score >= 2 indicates "
                "random/invalid responding. Test is INVALID."
            )
            return False
        elif v_score == 1:
            self.validity_status = 'QUESTIONABLE'
            self.validity_notes.append(
                f"Scale V (Invalidity) = {v_score}. Questionable validity "
                "based on Scale V score. Interpret with caution."
            )
        else:
            self.validity_notes.append(f"Scale V (Invalidity) = {v_score}. VALID.")
        
        return True


    
    # =========================================================================
    # STEP 4: SCALE W (INCONSISTENCY) CHECK
    # =========================================================================
    def _check_inconsistency(self):
        """
        Step 4-1 to 4-3: Calculate Scale W (Inconsistency) score.
        
        Per manual: "The Inconsistency scale consists of 44 pairs of
        item responses. Each blackened pair of responses adds 1 to 
        the scale score."
        
        This means: For each pair (item_a, item_b), if BOTH items
        are answered TRUE (both blackened/endorsed), that pair scores 1.
        The score is the count of pairs where both items are TRUE.
        
        Uses 5 plastic overlay keys. Each key shows pairs of response
        circles connected by lines. If both circles in a pair are 
        filled in (TRUE), score 1 point for that pair.
        
        If W >= 10: INVALID (further scoring not recommended)
        If W == 8 or 9: Questionable
        If W < 8: Valid
        
        Returns: True if scoring should continue, False if invalid.
        """
        w_score = 0
        
        for item_a, item_b in SCALE_W_PAIRS:
            resp_a = self.responses.get(item_a, False)
            resp_b = self.responses.get(item_b, False)
            
            # Score 1 if BOTH items in the pair are TRUE (both blackened)
            if resp_a and resp_b:
                w_score += 1
        
        self.raw_scores['W'] = w_score
        
        if w_score >= 10:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Scale W (Inconsistency) = {w_score}. Score >= 10 indicates "
                "inconsistent responding. Test is INVALID. "
                "Further scoring/interpretation NOT recommended."
            )
            return False
        elif w_score >= 8:
            if self.validity_status != 'INVALID':
                self.validity_status = 'QUESTIONABLE'
            self.validity_notes.append(
                f"Scale W (Inconsistency) = {w_score}. Questionable validity "
                "based on Scale W score."
            )
        else:
            self.validity_notes.append(
                f"Scale W (Inconsistency) = {w_score}. Within acceptable range."
            )
        return True
    
    # =========================================================================
    # STEP 5: RAW SCORE CALCULATION
    # =========================================================================
    def _calculate_raw_scores(self):
        """
        Step 5-1 to 5-4: Calculate raw scores for all scales
        (except V, W, and X which are handled separately).
        
        For each scale, sum the weights of items answered True.
        Weight-1 items: add 1 if True
        Weight-2 items: add 2 if True
        Weight-3 items: add 3 if True
        """
        for scale_name, items in SCALE_ITEMS.items():
            raw_score = 0
            for item_info in items:
                if isinstance(item_info, tuple):
                    item_num, weight = item_info
                else:
                    item_num = item_info
                    weight = 1
                
                if self.responses.get(item_num, False):
                    raw_score += weight
            
            self.raw_scores[scale_name] = raw_score


    
    # =========================================================================
    # STEP 6: SCALE X (DISCLOSURE) RAW SCORE
    # =========================================================================
    def _calculate_scale_x(self):
        """
        Step 6-1 to 6-6: Calculate Scale X raw score.
        
        Scale X = Sum of raw scores for Scales 1, 2A, 2B, 3, 4, 5(*2/3),
                  6A, 6B, 7, 8A, 8B
        
        Per manual: X is based on the raw scores for Scales 1 through 8B.
        Scale 5 raw score is multiplied by 2/3 before adding to sum.
        """
        personality_scales = ['1', '2A', '2B', '3', '4', '5', '6A', '6B', '7', '8A', '8B']
        
        x_score = 0.0
        x_components = {}
        
        for scale in personality_scales:
            raw = self.raw_scores.get(scale, 0)
            if scale == '5':
                # Scale 5 is multiplied by 2/3
                contribution = raw * (2.0 / 3.0)
                x_components[scale] = f"{raw} x 2/3 = {contribution:.2f}"
            else:
                contribution = raw
                x_components[scale] = str(raw)
            x_score += contribution
        
        # Step 6-5: Round to whole number
        # Round down if decimal < 0.5, up if >= 0.5
        x_score_rounded = round(x_score)
        
        self.raw_scores['X'] = x_score_rounded
        self.x_score_detail = {
            'components': x_components,
            'sum_decimal': x_score,
            'final_rounded': x_score_rounded
        }
        
        self.adjustment_log.append(
            f"Scale X Calculation: Sum = {x_score:.2f}, "
            f"Rounded = {x_score_rounded}"
        )
    
    # =========================================================================
    # STEP 7: SCALE X VALIDITY CHECK
    # =========================================================================
    def _check_scale_x_validity(self):
        """
        Step 7-1: Check if Scale X raw score is within valid range.
        Valid range: 34-178
        If < 34 or > 178: INVALID
        """
        x_raw = self.raw_scores.get('X', 0)
        
        if x_raw < 34:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Scale X (Disclosure) raw score = {x_raw}. "
                "Score < 34 indicates extreme defensiveness/low disclosure. "
                "Test is INVALID."
            )
            return False
        elif x_raw > 178:
            self.validity_status = 'INVALID'
            self.validity_notes.append(
                f"Scale X (Disclosure) raw score = {x_raw}. "
                "Score > 178 indicates extreme over-reporting. "
                "Test is INVALID."
            )
            return False
        else:
            self.validity_notes.append(
                f"Scale X (Disclosure) raw score = {x_raw}. "
                "Within valid range (34-178)."
            )
            if self.validity_status not in ['INVALID', 'QUESTIONABLE']:
                self.validity_status = 'VALID'
            return True


    
    # =========================================================================
    # STEP 8: INITIAL BR SCORE CONVERSION
    # =========================================================================
    def _calculate_initial_br_scores(self):
        """
        Step 8-1 to 8-3: Convert raw scores to initial BR scores
        using the BR conversion tables (Appendix C).
        """
        # Convert all scales from raw to BR
        scales_to_convert = (['X', 'Y', 'Z'] + CLINICAL_PERSONALITY + 
                            SEVERE_PERSONALITY + CLINICAL_SYNDROMES + 
                            SEVERE_SYNDROMES)
        
        for scale in scales_to_convert:
            raw = self.raw_scores.get(scale, 0)
            br = raw_to_br(scale, raw)
            self.initial_br_scores[scale] = br
            self.adjusted_br_scores[scale] = br  # Start with initial values
        
        self.adjustment_log.append("Initial BR scores calculated from conversion tables.")
    
    # =========================================================================
    # STEP 9: DISCLOSURE ADJUSTMENT
    # =========================================================================
    def _apply_disclosure_adjustment(self):
        """
        Step 9-1 to 9-4: Apply Disclosure adjustment.
        
        Based on raw X score, apply:
        - 1-8B adjustment to Clinical Personality scales (1-8B)
        - S-PP adjustment to Severe Personality + Clinical/Severe Syndromes
        """
        x_raw = self.raw_scores.get('X', 0)
        adjustment = get_disclosure_adjustment(x_raw)
        
        if adjustment is None:
            self.adjustment_log.append(
                "Disclosure adjustment: SKIPPED (invalid X score)"
            )
            return
        
        adj_1_8b, adj_s_pp = adjustment
        
        self.adjustment_log.append(
            f"Disclosure Adjustment (X raw={x_raw}): "
            f"1-8B = {adj_1_8b:+d}, S-PP = {adj_s_pp:+d}"
        )
        
        # Apply 1-8B adjustment to Clinical Personality scales
        for scale in CLINICAL_PERSONALITY:
            old_br = self.adjusted_br_scores[scale]
            new_br = max(0, min(115, old_br + adj_1_8b))
            self.adjusted_br_scores[scale] = new_br
        
        # Apply S-PP adjustment to Severe Personality + all Syndrome scales
        affected_scales = SEVERE_PERSONALITY + CLINICAL_SYNDROMES + SEVERE_SYNDROMES
        for scale in affected_scales:
            old_br = self.adjusted_br_scores[scale]
            new_br = max(0, min(115, old_br + adj_s_pp))
            self.adjusted_br_scores[scale] = new_br


    
    # =========================================================================
    # STEP 10: ANXIETY/DEPRESSION (A/D) ADJUSTMENT
    # =========================================================================
    def _apply_ad_adjustment(self):
        """
        Step 10-1 to 10-3: Apply A/D adjustment.
        
        Affects: Scales 2A, 2B, 8B, S, C
        
        Algorithm:
        1. Compute A/D value based on current BR scores for A and D
        2. Look up adjustment in appropriate table (OPD/IPD)
        3. Apply 2B,8B,C adjustment and 2A,S adjustment
        """
        br_a = self.adjusted_br_scores.get('A', 0)
        br_d = self.adjusted_br_scores.get('D', 0)
        
        # Determine A/D value
        if br_a < 75 and br_d < 75:
            # No A/D adjustment necessary
            self.adjustment_log.append(
                f"A/D Adjustment: NOT NEEDED (A BR={br_a}, D BR={br_d}, "
                "both < 75)"
            )
            return
        elif br_a >= 75 and br_d < 75:
            ad_value = br_a - 75
        elif br_d >= 75 and br_a < 75:
            ad_value = br_d - 75
        else:  # Both >= 75
            ad_value = (br_a - 75) + (br_d - 75)
        
        # Get adjustment factors
        adj_2b_8b_c, adj_2a_s = get_ad_adjustment(
            ad_value, self.patient_setting, self.axis_duration_weeks
        )
        
        self.adjustment_log.append(
            f"A/D Adjustment (A BR={br_a}, D BR={br_d}, A/D value={ad_value}, "
            f"Setting={self.patient_setting}): "
            f"2B,8B,C adj = {adj_2b_8b_c:+d}, 2A,S adj = {adj_2a_s:+d}"
        )
        
        # Apply 2B, 8B, C adjustment
        for scale in ['2B', '8B', 'C']:
            old_br = self.adjusted_br_scores[scale]
            new_br = max(0, min(115, old_br + adj_2b_8b_c))
            self.adjusted_br_scores[scale] = new_br
        
        # Apply 2A, S adjustment
        for scale in ['2A', 'S']:
            old_br = self.adjusted_br_scores[scale]
            new_br = max(0, min(115, old_br + adj_2a_s))
            self.adjusted_br_scores[scale] = new_br
    
    # =========================================================================
    # STEP 11: INPATIENT ADJUSTMENT
    # =========================================================================
    def _apply_inpatient_adjustment(self):
        """
        Step 11-1 to 11-4: Apply Inpatient adjustment.
        
        Affects: Scales SS, CC, PP (Severe Clinical Syndromes)
        
        Only applies if patient is IPD (inpatient) with Axis I
        duration <= 4 weeks.
        """
        if self.patient_setting != 'IPD':
            self.adjustment_log.append(
                "Inpatient Adjustment: NOT APPLICABLE (patient is not inpatient)"
            )
            return
        
        if self.axis_duration_weeks is None or self.axis_duration_weeks > 4:
            self.adjustment_log.append(
                "Inpatient Adjustment: NOT APPLICABLE "
                "(Axis I duration > 4 weeks or not specified)"
            )
            return
        
        # Determine duration category
        if self.axis_duration_weeks < 1:
            duration_key = 'less_than_1_week'
            duration_label = "< 1 week"
        elif 1 <= self.axis_duration_weeks <= 4:
            duration_key = '1_to_4_weeks'
            duration_label = "1-4 weeks"
        else:
            duration_key = 'more_than_4_weeks'
            duration_label = "> 4 weeks"
        
        adj_ss, adj_cc, adj_pp = INPATIENT_ADJUSTMENT[duration_key]
        
        self.adjustment_log.append(
            f"Inpatient Adjustment (Duration: {duration_label}): "
            f"SS = {adj_ss:+d}, CC = {adj_cc:+d}, PP = {adj_pp:+d}"
        )
        
        # Apply adjustments
        adjustments = {'SS': adj_ss, 'CC': adj_cc, 'PP': adj_pp}
        for scale, adj in adjustments.items():
            old_br = self.adjusted_br_scores[scale]
            new_br = max(0, min(115, old_br + adj))
            self.adjusted_br_scores[scale] = new_br


    
    # =========================================================================
    # STEP 12: DENIAL/COMPLAINT ADJUSTMENT
    # =========================================================================
    def _apply_denial_complaint_adjustment(self):
        """
        Step 12-1 to 12-2: Apply Denial/Complaint adjustment.
        
        Algorithm:
        1. Find the highest-ranked Clinical Personality Pattern scale (1-8B)
           based on current adjusted BR scores
        2. If highest is Scale 4, 5, or 7: add 8 points to that scale
        3. If highest is anything else: no adjustment
        
        Tie-breaking order (highest priority first):
        2B, 6A, 8B, 6B, 2A, 8A, 7, 4, 5, 3, 1
        """
        # Tie-breaking priority (higher index = higher priority for tie-breaking)
        tie_break_order = ['1', '3', '5', '4', '7', '8A', '2A', '6B', '8B', '6A', '2B']
        
        # Find highest Clinical Personality Pattern scale
        highest_scale = None
        highest_br = -1
        
        for scale in CLINICAL_PERSONALITY:
            br = self.adjusted_br_scores.get(scale, 0)
            if br > highest_br:
                highest_br = br
                highest_scale = scale
            elif br == highest_br:
                # Tie-breaking: scale appearing later in tie_break_order wins
                if (scale in tie_break_order and highest_scale in tie_break_order):
                    if tie_break_order.index(scale) > tie_break_order.index(highest_scale):
                        highest_scale = scale
        
        # Check if highest is 4, 5, or 7
        if highest_scale in ['4', '5', '7']:
            old_br = self.adjusted_br_scores[highest_scale]
            new_br = max(0, min(115, old_br + 8))
            self.adjusted_br_scores[highest_scale] = new_br
            
            self.adjustment_log.append(
                f"Denial/Complaint Adjustment: Highest personality scale = "
                f"{highest_scale} ({SCALE_NAMES[highest_scale]}) with BR={old_br}. "
                f"Added +8 -> {new_br}"
            )
        else:
            self.adjustment_log.append(
                f"Denial/Complaint Adjustment: NO adjustment needed. "
                f"Highest scale = {highest_scale} "
                f"({SCALE_NAMES.get(highest_scale, 'N/A')}) "
                f"with BR={highest_br}. "
                f"(Adjustment only applies if highest is 4, 5, or 7)"
            )
    
    # =========================================================================
    # STEP 13: FINALIZE SCORES
    # =========================================================================
    def _finalize_scores(self):
        """
        Step 13-1 to 13-2: Transfer final BR scores.
        Also checks for FLAT PROFILE (all Clinical Personality BR < 60).
        """
        # Copy adjusted BR scores to final
        all_scales = (['X', 'Y', 'Z'] + CLINICAL_PERSONALITY + 
                     SEVERE_PERSONALITY + CLINICAL_SYNDROMES + SEVERE_SYNDROMES)
        
        for scale in all_scales:
            self.final_br_scores[scale] = self.adjusted_br_scores.get(scale, 0)
        
        # Check for flat profile (Step 14-1)
        personality_brs = [
            self.final_br_scores.get(s, 0) for s in CLINICAL_PERSONALITY
        ]
        
        if all(br < 60 for br in personality_brs):
            self.validity_notes.append(
                "WARNING: FLAT PROFILE detected. All Clinical Personality "
                "Pattern BR scores < 60. Profile is considered UNINTERPRETABLE."
            )
            self.flat_profile = True
        else:
            self.flat_profile = False


    
    # =========================================================================
    # BUILD RESULTS
    # =========================================================================
    def _build_results(self):
        """Build comprehensive results dictionary."""
        results = {
            'validity': {
                'status': self.validity_status,
                'notes': self.validity_notes,
                'V_score': self.raw_scores.get('V', None),
                'W_score': self.raw_scores.get('W', None),
                'X_raw': self.raw_scores.get('X', None),
                'omitted_items': getattr(self, 'omitted_count', None),
            },
            'patient_info': {
                'age': self.age,
                'setting': self.patient_setting,
                'axis_duration_weeks': self.axis_duration_weeks,
                'gender': self.gender,
            },
            'raw_scores': dict(self.raw_scores),
            'initial_br_scores': dict(self.initial_br_scores),
            'adjusted_br_scores': dict(self.adjusted_br_scores),
            'final_br_scores': dict(self.final_br_scores),
            'interpretations': {},
            'adjustment_log': self.adjustment_log,
            'flat_profile': getattr(self, 'flat_profile', None),
            'elevated_scales': {
                'prominent_85+': [],
                'significant_75_84': [],
                'trait_60_74': [],
            }
        }
        
        # Generate interpretations for each scale
        for scale in self.final_br_scores:
            br = self.final_br_scores[scale]
            results['interpretations'][scale] = {
                'scale_name': SCALE_NAMES.get(scale, scale),
                'raw_score': self.raw_scores.get(scale, 'N/A'),
                'initial_br': self.initial_br_scores.get(scale, 'N/A'),
                'final_br': br,
                'interpretation': interpret_br_score(br),
            }
            
            # Categorize elevated scales
            if br >= 85:
                results['elevated_scales']['prominent_85+'].append(
                    f"{scale} ({SCALE_NAMES.get(scale, '')}): BR={br}"
                )
            elif br >= 75:
                results['elevated_scales']['significant_75_84'].append(
                    f"{scale} ({SCALE_NAMES.get(scale, '')}): BR={br}"
                )
            elif br >= 60:
                results['elevated_scales']['trait_60_74'].append(
                    f"{scale} ({SCALE_NAMES.get(scale, '')}): BR={br}"
                )
        
        return results


    
    # =========================================================================
    # UTILITY METHODS
    # =========================================================================
    def get_scale_detail(self, scale_name):
        """Get detailed information about a specific scale's scoring."""
        return {
            'scale': scale_name,
            'full_name': SCALE_NAMES.get(scale_name, 'Unknown'),
            'raw_score': self.raw_scores.get(scale_name, 'Not calculated'),
            'initial_br': self.initial_br_scores.get(scale_name, 'Not calculated'),
            'final_br': self.final_br_scores.get(scale_name, 'Not calculated'),
            'interpretation': interpret_br_score(
                self.final_br_scores.get(scale_name, 0)
            ),
        }
    
    def print_summary(self):
        """Print a formatted summary of results."""
        results = self._build_results()
        
        print("=" * 70)
        print("          MCMI-III SCORING RESULTS SUMMARY")
        print("=" * 70)
        
        # Patient Info
        print(f"\n{'PATIENT INFORMATION':=^70}")
        print(f"  Age: {self.age}")
        print(f"  Setting: {self.patient_setting}")
        if self.patient_setting == 'IPD':
            print(f"  Axis I Duration: {self.axis_duration_weeks} weeks")
        print(f"  Gender: {self.gender or 'Not specified'}")
        
        # Validity
        print(f"\n{'VALIDITY STATUS':=^70}")
        print(f"  Overall: {self.validity_status}")
        for note in self.validity_notes:
            print(f"  - {note}")
        
        if self.validity_status == 'INVALID':
            print("\n  *** SCORING STOPPED: Protocol is INVALID ***")
            return
        
        # Raw Scores
        print(f"\n{'RAW SCORES':=^70}")
        print(f"  {'Scale':<8} {'Name':<35} {'Raw Score':<10}")
        print(f"  {'-'*8} {'-'*35} {'-'*10}")
        
        all_scales = (['V', 'W', 'X', 'Y', 'Z'] + CLINICAL_PERSONALITY + 
                     SEVERE_PERSONALITY + CLINICAL_SYNDROMES + SEVERE_SYNDROMES)
        for scale in all_scales:
            if scale in self.raw_scores:
                name = SCALE_NAMES.get(scale, '')
                print(f"  {scale:<8} {name:<35} {self.raw_scores[scale]:<10}")
        
        # BR Scores
        print(f"\n{'BR SCORES (Initial -> Final)':=^70}")
        print(f"  {'Scale':<6} {'Name':<30} {'Raw':<5} {'Init BR':<8} "
              f"{'Final BR':<9} {'Interpretation':<20}")
        print(f"  {'-'*6} {'-'*30} {'-'*5} {'-'*8} {'-'*9} {'-'*20}")
        
        scoring_scales = (['X', 'Y', 'Z'] + CLINICAL_PERSONALITY + 
                         SEVERE_PERSONALITY + CLINICAL_SYNDROMES + SEVERE_SYNDROMES)
        for scale in scoring_scales:
            if scale in self.final_br_scores:
                name = SCALE_NAMES.get(scale, '')[:30]
                raw = self.raw_scores.get(scale, '-')
                init_br = self.initial_br_scores.get(scale, '-')
                final_br = self.final_br_scores[scale]
                interp = interpret_br_score(final_br)
                
                # Highlight elevated scales
                marker = ""
                if final_br >= 85:
                    marker = " ***"
                elif final_br >= 75:
                    marker = " **"
                
                print(f"  {scale:<6} {name:<30} {raw:<5} {init_br:<8} "
                      f"{final_br:<9} {interp}{marker}")
        
        # Adjustment Log
        print(f"\n{'ADJUSTMENT LOG':=^70}")
        for entry in self.adjustment_log:
            print(f"  {entry}")
        
        # Elevated Scales Summary
        print(f"\n{'ELEVATED SCALES SUMMARY':=^70}")
        elevated = results['elevated_scales']
        
        if elevated['prominent_85+']:
            print("\n  PROMINENT (BR >= 85) - Disorder Likely Present:")
            for s in elevated['prominent_85+']:
                print(f"    *** {s}")
        
        if elevated['significant_75_84']:
            print("\n  CLINICALLY SIGNIFICANT (BR 75-84):")
            for s in elevated['significant_75_84']:
                print(f"    ** {s}")
        
        if elevated['trait_60_74']:
            print("\n  Trait/Tendency Present (BR 60-74):")
            for s in elevated['trait_60_74']:
                print(f"    * {s}")
        
        if not any(elevated.values()):
            print("  No scales elevated above BR 60.")
        
        # Flat profile warning
        if getattr(self, 'flat_profile', False):
            print(f"\n{'!!! WARNING: FLAT PROFILE !!!':^70}")
            print("  All Clinical Personality BR scores < 60.")
            print("  Profile is considered UNINTERPRETABLE.")
        
        print("\n" + "=" * 70)
