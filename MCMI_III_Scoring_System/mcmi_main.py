#!/usr/bin/env python3
"""
MCMI-III COMPLETE SCORING APPLICATION
=======================================
A clinically accurate, exam-level MCMI-III scoring system.

Based on:
- Official MCMI-III Manual (Millon, Davis, Grossman, Millon)
- MCMI-III Hand-Scoring User's Guide
- Standard BR (Base Rate) scoring methodology

Usage:
    python mcmi_main.py                    # Interactive mode
    python mcmi_main.py --file input.csv   # Score from file
    python mcmi_main.py --demo             # Run demo scoring

Author: Clinical Psychometric Scoring System
"""

import sys
import os
import json
import csv
from datetime import datetime

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcmi_scoring_engine import MCMIIIScorer
from mcmi_item_keys import SCALE_NAMES, ALL_SCALES_ORDER



def interactive_data_entry():
    """
    Interactive mode for entering 175 MCMI-III item responses.
    Returns dict of {item_number: True/False}.
    """
    print("\n" + "=" * 60)
    print("    MCMI-III DATA ENTRY")
    print("    Enter responses for 175 items")
    print("    Format: T (True) or F (False)")
    print("    Enter 'q' to quit, 'b' to go back")
    print("=" * 60)
    
    responses = {}
    i = 1
    
    while i <= 175:
        try:
            resp = input(f"  Item {i:3d}: ").strip().upper()
            
            if resp == 'Q':
                print("  Exiting data entry...")
                return None
            elif resp == 'B' and i > 1:
                i -= 1
                continue
            elif resp in ['T', 'TRUE', '1', 'Y', 'YES']:
                responses[i] = True
                i += 1
            elif resp in ['F', 'FALSE', '0', 'N', 'NO']:
                responses[i] = False
                i += 1
            else:
                print("  Invalid input. Enter T (True) or F (False).")
        except (EOFError, KeyboardInterrupt):
            print("\n  Data entry interrupted.")
            return None
    
    return responses


def get_patient_info():
    """Get patient demographic and clinical information."""
    print("\n" + "=" * 60)
    print("    PATIENT INFORMATION")
    print("=" * 60)
    
    # Age
    while True:
        try:
            age = int(input("  Patient Age: "))
            if age < 0 or age > 120:
                print("  Please enter a valid age.")
                continue
            break
        except ValueError:
            print("  Please enter a number.")
    
    # Gender
    while True:
        gender = input("  Gender (M/F): ").strip().upper()
        if gender in ['M', 'F']:
            break
        print("  Please enter M or F.")
    
    # Patient Setting
    while True:
        setting = input("  Setting - OPD (Outpatient) or IPD (Inpatient): ").strip().upper()
        if setting in ['OPD', 'IPD']:
            break
        print("  Please enter OPD or IPD.")
    
    # Axis I duration (for inpatients)
    axis_duration = None
    if setting == 'IPD':
        print("\n  Axis I Disorder Duration Options:")
        print("    1 = Less than 1 week")
        print("    2 = 1 to 4 weeks")
        print("    3 = More than 4 weeks")
        
        while True:
            try:
                choice = int(input("  Duration (1/2/3): "))
                if choice == 1:
                    axis_duration = 0.5  # Less than 1 week
                elif choice == 2:
                    axis_duration = 2  # 1-4 weeks
                elif choice == 3:
                    axis_duration = 5  # More than 4 weeks
                else:
                    print("  Enter 1, 2, or 3.")
                    continue
                break
            except ValueError:
                print("  Please enter a number.")
    
    return age, gender, setting, axis_duration



def load_responses_from_csv(filepath):
    """
    Load responses from a CSV file.
    Expected format: One column with 175 rows of T/F or 1/0 values.
    OR: Two columns (item_number, response).
    """
    responses = {}
    
    with open(filepath, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # Skip header if present
    start_idx = 0
    if rows and rows[0][0].lower() in ['item', 'item_number', 'question', '#']:
        start_idx = 1
    
    for idx, row in enumerate(rows[start_idx:], 1):
        if len(row) >= 2:
            # Two-column format: item_number, response
            try:
                item_num = int(row[0])
                resp = row[1].strip().upper()
            except ValueError:
                continue
        elif len(row) == 1:
            # Single column: just responses in order
            item_num = idx
            resp = row[0].strip().upper()
        else:
            continue
        
        if resp in ['T', 'TRUE', '1', 'Y', 'YES']:
            responses[item_num] = True
        elif resp in ['F', 'FALSE', '0', 'N', 'NO']:
            responses[item_num] = False
    
    return responses


def load_responses_from_json(filepath):
    """Load responses from a JSON file."""
    with open(filepath, 'r') as f:
        data = json.load(f)
    
    if isinstance(data, list):
        return {i+1: bool(v) for i, v in enumerate(data)}
    elif isinstance(data, dict):
        return {int(k): bool(v) for k, v in data.items()}
    else:
        raise ValueError("JSON must be a list or dict")


def generate_report(results, output_path=None):
    """Generate a detailed clinical report."""
    report = []
    report.append("=" * 70)
    report.append("       MCMI-III COMPREHENSIVE SCORING REPORT")
    report.append(f"       Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("=" * 70)
    
    # Patient Information
    pi = results['patient_info']
    report.append(f"\n{'PATIENT INFORMATION':=^70}")
    report.append(f"  Age: {pi['age']}")
    report.append(f"  Gender: {pi['gender'] or 'Not specified'}")
    report.append(f"  Setting: {pi['setting']} "
                  f"({'Outpatient/Non-Inpatient' if pi['setting']=='OPD' else 'Inpatient'})")
    if pi['setting'] == 'IPD' and pi['axis_duration_weeks'] is not None:
        report.append(f"  Axis I Duration: {pi['axis_duration_weeks']} weeks")
    
    # Validity Section
    report.append(f"\n{'VALIDITY ASSESSMENT':=^70}")
    v = results['validity']
    report.append(f"  Overall Validity: {v['status']}")
    report.append(f"  Scale V (Invalidity): {v['V_score']}")
    report.append(f"  Scale W (Inconsistency): {v['W_score']}")
    report.append(f"  Scale X (Disclosure) Raw: {v['X_raw']}")
    report.append(f"  Omitted Items: {v['omitted_items']}")
    report.append("")
    for note in v['notes']:
        report.append(f"  {note}")
    
    if v['status'] == 'INVALID':
        report.append("\n  *** PROTOCOL IS INVALID - CANNOT BE INTERPRETED ***")
        report.append("=" * 70)
        full_report = "\n".join(report)
        if output_path:
            with open(output_path, 'w') as f:
                f.write(full_report)
        return full_report
    
    # Modifying Indices
    report.append(f"\n{'MODIFYING INDICES':=^70}")
    report.append(f"  {'Scale':<6} {'Name':<20} {'Raw':<6} {'BR':<6} "
                  f"{'Interpretation':<25}")
    report.append(f"  {'-'*6} {'-'*20} {'-'*6} {'-'*6} {'-'*25}")
    for scale in ['X', 'Y', 'Z']:
        if scale in results['final_br_scores']:
            name = SCALE_NAMES[scale]
            raw = results['raw_scores'].get(scale, '-')
            br = results['final_br_scores'][scale]
            interp = ""
            if scale == 'X':
                if br < 35:
                    interp = "Low Disclosure (Defensive)"
                elif br > 75:
                    interp = "High Disclosure (Over-reporting)"
                else:
                    interp = "Normal Disclosure"
            elif scale == 'Y':
                if br >= 75:
                    interp = "High Desirability"
                else:
                    interp = "Normal"
            elif scale == 'Z':
                if br >= 75:
                    interp = "High Debasement"
                else:
                    interp = "Normal"
            report.append(f"  {scale:<6} {name:<20} {raw:<6} {br:<6} {interp}")
    
    # Clinical Personality Patterns
    report.append(f"\n{'CLINICAL PERSONALITY PATTERNS (Axis II)':=^70}")
    report.append(f"  {'Scale':<6} {'Name':<35} {'Raw':<5} {'BR':<5} "
                  f"{'Level':<25}")
    report.append(f"  {'-'*6} {'-'*35} {'-'*5} {'-'*5} {'-'*25}")
    
    from mcmi_item_keys import CLINICAL_PERSONALITY
    for scale in CLINICAL_PERSONALITY:
        if scale in results['final_br_scores']:
            name = SCALE_NAMES[scale]
            raw = results['raw_scores'].get(scale, '-')
            br = results['final_br_scores'][scale]
            from mcmi_br_tables import interpret_br_score as interp_fn
            interp = interp_fn(br)
            marker = "***" if br >= 85 else "**" if br >= 75 else ""
            report.append(f"  {scale:<6} {name:<35} {raw:<5} {br:<5} "
                         f"{interp} {marker}")
    
    # Severe Personality Pathology
    report.append(f"\n{'SEVERE PERSONALITY PATHOLOGY':=^70}")
    report.append(f"  {'Scale':<6} {'Name':<35} {'Raw':<5} {'BR':<5} "
                  f"{'Level':<25}")
    report.append(f"  {'-'*6} {'-'*35} {'-'*5} {'-'*5} {'-'*25}")
    
    from mcmi_item_keys import SEVERE_PERSONALITY
    for scale in SEVERE_PERSONALITY:
        if scale in results['final_br_scores']:
            name = SCALE_NAMES[scale]
            raw = results['raw_scores'].get(scale, '-')
            br = results['final_br_scores'][scale]
            interp = interp_fn(br)
            marker = "***" if br >= 85 else "**" if br >= 75 else ""
            report.append(f"  {scale:<6} {name:<35} {raw:<5} {br:<5} "
                         f"{interp} {marker}")
    
    # Clinical Syndromes
    report.append(f"\n{'CLINICAL SYNDROMES (Axis I)':=^70}")
    report.append(f"  {'Scale':<6} {'Name':<35} {'Raw':<5} {'BR':<5} "
                  f"{'Level':<25}")
    report.append(f"  {'-'*6} {'-'*35} {'-'*5} {'-'*5} {'-'*25}")
    
    from mcmi_item_keys import CLINICAL_SYNDROMES
    for scale in CLINICAL_SYNDROMES:
        if scale in results['final_br_scores']:
            name = SCALE_NAMES[scale]
            raw = results['raw_scores'].get(scale, '-')
            br = results['final_br_scores'][scale]
            interp = interp_fn(br)
            marker = "***" if br >= 85 else "**" if br >= 75 else ""
            report.append(f"  {scale:<6} {name:<35} {raw:<5} {br:<5} "
                         f"{interp} {marker}")
    
    # Severe Clinical Syndromes
    report.append(f"\n{'SEVERE CLINICAL SYNDROMES':=^70}")
    report.append(f"  {'Scale':<6} {'Name':<35} {'Raw':<5} {'BR':<5} "
                  f"{'Level':<25}")
    report.append(f"  {'-'*6} {'-'*35} {'-'*5} {'-'*5} {'-'*25}")
    
    from mcmi_item_keys import SEVERE_SYNDROMES
    for scale in SEVERE_SYNDROMES:
        if scale in results['final_br_scores']:
            name = SCALE_NAMES[scale]
            raw = results['raw_scores'].get(scale, '-')
            br = results['final_br_scores'][scale]
            interp = interp_fn(br)
            marker = "***" if br >= 85 else "**" if br >= 75 else ""
            report.append(f"  {scale:<6} {name:<35} {raw:<5} {br:<5} "
                         f"{interp} {marker}")
    
    # Adjustment Log
    report.append(f"\n{'ADJUSTMENT LOG':=^70}")
    for entry in results['adjustment_log']:
        report.append(f"  {entry}")
    
    # Elevated Scales Summary
    report.append(f"\n{'PROFILE SUMMARY':=^70}")
    elevated = results['elevated_scales']
    
    if elevated['prominent_85+']:
        report.append("\n  PROMINENT SCALES (BR >= 85) - Disorder Present:")
        for s in elevated['prominent_85+']:
            report.append(f"    *** {s}")
    
    if elevated['significant_75_84']:
        report.append("\n  CLINICALLY SIGNIFICANT (BR 75-84):")
        for s in elevated['significant_75_84']:
            report.append(f"    ** {s}")
    
    if elevated['trait_60_74']:
        report.append("\n  Trait Present (BR 60-74):")
        for s in elevated['trait_60_74']:
            report.append(f"    * {s}")
    
    if results.get('flat_profile'):
        report.append("\n  !!! FLAT PROFILE: All personality BR < 60 !!!")
        report.append("  Profile is considered UNINTERPRETABLE.")
    
    # Legend
    report.append(f"\n{'LEGEND':=^70}")
    report.append("  BR < 60  : Not significant")
    report.append("  BR 60-74 : Trait/Tendency present")
    report.append("  BR 75-84 : Clinically significant (**)")
    report.append("  BR >= 85 : Prominent - Disorder present (***)")
    
    report.append("\n" + "=" * 70)
    report.append("  END OF REPORT")
    report.append("=" * 70)
    
    full_report = "\n".join(report)
    
    if output_path:
        with open(output_path, 'w') as f:
            f.write(full_report)
        print(f"\n  Report saved to: {output_path}")
    
    return full_report



def run_demo():
    """
    Run a demonstration scoring with sample data.
    This demonstrates an OPD patient with moderate pathology.
    """
    print("\n" + "=" * 60)
    print("    MCMI-III SCORING SYSTEM - DEMONSTRATION")
    print("=" * 60)
    print("\n  Running demo with sample patient data...")
    print("  Setting: OPD (Outpatient)")
    print("  Age: 35, Gender: Female")
    
    # Generate sample responses (simulating a patient with
    # depressive/avoidant features)
    import random
    random.seed(42)
    
    # Start with all False
    responses = {i: False for i in range(1, 176)}
    
    # Endorse items that would elevate certain scales
    # Depressive items (Scale 2B related)
    depressive_items = [2, 20, 22, 41, 50, 53, 56, 71, 81, 98, 
                        121, 130, 135, 158, 172]
    for item in depressive_items:
        responses[item] = True
    
    # Anxiety items (Scale A related)
    anxiety_items = [16, 35, 44, 48, 83, 87, 92, 107, 114, 118, 
                     122, 136, 154, 163]
    for item in anxiety_items:
        responses[item] = True
    
    # Avoidant items (Scale 2A related)
    avoidant_items = [46, 78, 147]
    for item in avoidant_items:
        responses[item] = True
    
    # Some dependent items
    dependent_items = [5, 15, 33, 60, 70, 75, 80, 86, 91]
    for item in dependent_items:
        responses[item] = True
    
    # Some debasement items
    debasement_items = [9, 30, 72, 109, 148, 175, 76, 96, 100]
    for item in debasement_items:
        responses[item] = True
    
    # Some normal/desirability items for balance
    desirability_items = [4, 68, 74, 82, 95, 99]
    for item in desirability_items:
        responses[item] = True
    
    # A few more scattered True responses
    other_items = [3, 10, 19, 25, 39, 45, 49, 57, 119, 125, 
                   150, 160, 165, 166, 169]
    for item in other_items:
        responses[item] = True
    
    # Ensure W (Inconsistency) pairs are concordant to keep W score low
    # W pairs are items that should be answered in the same direction
    # Make paired items match each other
    from mcmi_item_keys import SCALE_W_PAIRS
    for item_a, item_b in SCALE_W_PAIRS:
        # Make both items agree (both True or both False)
        if item_a in responses and responses[item_a]:
            responses[item_b] = True
        elif item_b in responses and responses[item_b]:
            responses[item_a] = True
        # If neither is True, both stay False (already concordant)
    
    # Make sure V items are False (valid protocol)
    responses[65] = False
    responses[110] = False
    responses[157] = False
    
    # Score
    scorer = MCMIIIScorer()
    scorer.set_responses(responses)
    scorer.set_patient_info(age=35, setting='OPD', gender='F')
    results = scorer.score()
    
    # Print summary
    scorer.print_summary()
    
    # Generate report
    report = generate_report(results)
    print(report)
    
    # Save demo report
    output_dir = os.path.dirname(os.path.abspath(__file__))
    report_path = os.path.join(output_dir, 'demo_report.txt')
    with open(report_path, 'w') as f:
        f.write(report)
    print(f"\n  Demo report saved to: {report_path}")
    
    return results


def run_ipd_demo():
    """
    Run a demonstration with an IPD (Inpatient) patient.
    Demonstrates the inpatient adjustment and different A/D table.
    """
    print("\n" + "=" * 60)
    print("    MCMI-III SCORING - IPD (INPATIENT) DEMONSTRATION")
    print("=" * 60)
    print("\n  Running demo with inpatient data...")
    print("  Setting: IPD (Inpatient)")
    print("  Axis I Duration: < 1 week")
    print("  Age: 28, Gender: Male")
    
    import random
    random.seed(123)
    
    responses = {i: False for i in range(1, 176)}
    
    # More severe pathology for inpatient
    # Borderline features
    borderline_items = [5, 10, 15, 20, 25, 33, 39, 46, 49, 60, 
                        70, 75, 80, 86, 91, 119, 125, 160, 165]
    for item in borderline_items:
        responses[item] = True
    
    # Severe depression
    depression_items = [9, 16, 35, 44, 48, 50, 56, 76, 83, 87, 
                        92, 96, 100, 107, 109, 122, 130, 136, 148]
    for item in depression_items:
        responses[item] = True
    
    # Thought disorder items
    thought_items = [30, 72, 114, 118, 163, 175]
    for item in thought_items:
        responses[item] = True
    
    # Antisocial/aggressive features
    antisocial_items = [6, 12, 17, 28, 34, 47, 58, 63, 77, 85, 
                        104, 111, 117]
    for item in antisocial_items:
        responses[item] = True
    
    # Additional items
    other = [8, 29, 32, 40, 42, 57, 66, 67, 93, 131, 143, 
             150, 155, 164, 166, 169, 172]
    for item in other:
        responses[item] = True
    
    # Keep V items False
    responses[65] = False
    responses[110] = False
    responses[157] = False
    
    # Ensure W (Inconsistency) pairs are concordant for valid protocol
    from mcmi_item_keys import SCALE_W_PAIRS
    for item_a, item_b in SCALE_W_PAIRS:
        if item_a in responses and responses[item_a]:
            responses[item_b] = True
        elif item_b in responses and responses[item_b]:
            responses[item_a] = True
    
    # Score with IPD setting
    scorer = MCMIIIScorer()
    scorer.set_responses(responses)
    scorer.set_patient_info(
        age=28, 
        setting='IPD', 
        axis_duration_weeks=0.5,  # Less than 1 week
        gender='M'
    )
    results = scorer.score()
    
    # Print summary
    scorer.print_summary()
    
    # Generate report
    report = generate_report(results)
    print(report)
    
    return results



def main():
    """Main entry point for the MCMI-III scoring application."""
    print("\n" + "=" * 60)
    print("    MCMI-III COMPREHENSIVE SCORING SYSTEM")
    print("    Based on Official MCMI-III Manual (Millon et al.)")
    print("    Hand-Scoring User's Guide Protocol")
    print("=" * 60)
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--demo':
            run_demo()
            return
        elif sys.argv[1] == '--ipd-demo':
            run_ipd_demo()
            return
        elif sys.argv[1] == '--file' and len(sys.argv) > 2:
            filepath = sys.argv[2]
            if filepath.endswith('.json'):
                responses = load_responses_from_json(filepath)
            else:
                responses = load_responses_from_csv(filepath)
            
            if not responses:
                print("  Error: Could not load responses from file.")
                return
            
            age, gender, setting, axis_duration = get_patient_info()
            
            scorer = MCMIIIScorer()
            scorer.set_responses(responses)
            scorer.set_patient_info(
                age=age, setting=setting,
                axis_duration_weeks=axis_duration, gender=gender
            )
            results = scorer.score()
            scorer.print_summary()
            
            # Generate report
            report = generate_report(results, 'mcmi_report.txt')
            print(report)
            return
        elif sys.argv[1] == '--help':
            print("\n  Usage:")
            print("    python mcmi_main.py              # Interactive mode")
            print("    python mcmi_main.py --demo       # OPD demo")
            print("    python mcmi_main.py --ipd-demo   # IPD demo")
            print("    python mcmi_main.py --file <path> # Score from file")
            print("    python mcmi_main.py --help       # Show this help")
            return
    
    # Interactive mode
    print("\n  Select mode:")
    print("    1. Enter responses manually (175 items)")
    print("    2. Load from CSV file")
    print("    3. Load from JSON file")
    print("    4. Run OPD (Outpatient) demo")
    print("    5. Run IPD (Inpatient) demo")
    print("    6. Exit")
    
    while True:
        try:
            choice = input("\n  Choice (1-6): ").strip()
            
            if choice == '1':
                responses = interactive_data_entry()
                if responses is None:
                    continue
            elif choice == '2':
                filepath = input("  Enter CSV file path: ").strip()
                responses = load_responses_from_csv(filepath)
            elif choice == '3':
                filepath = input("  Enter JSON file path: ").strip()
                responses = load_responses_from_json(filepath)
            elif choice == '4':
                run_demo()
                continue
            elif choice == '5':
                run_ipd_demo()
                continue
            elif choice == '6':
                print("  Goodbye!")
                return
            else:
                print("  Invalid choice.")
                continue
            
            if not responses:
                print("  Error: No responses loaded.")
                continue
            
            # Get patient info
            age, gender, setting, axis_duration = get_patient_info()
            
            # Score
            scorer = MCMIIIScorer()
            scorer.set_responses(responses)
            scorer.set_patient_info(
                age=age, setting=setting,
                axis_duration_weeks=axis_duration, gender=gender
            )
            results = scorer.score()
            
            # Display
            scorer.print_summary()
            
            # Save report?
            save = input("\n  Save report to file? (Y/N): ").strip().upper()
            if save == 'Y':
                filename = input("  Filename (default: mcmi_report.txt): ").strip()
                if not filename:
                    filename = 'mcmi_report.txt'
                report = generate_report(results, filename)
                print(report)
            else:
                report = generate_report(results)
                print(report)
            
        except (EOFError, KeyboardInterrupt):
            print("\n  Exiting...")
            return


if __name__ == '__main__':
    main()
