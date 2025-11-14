#!/usr/bin/env python3
"""
Quick verification that T0006 issues are fixed in new mappings.json
"""

import json
from pathlib import Path

def main():
    base_dir = Path(__file__).parent.parent

    old_json = base_dir / 'data' / 'mapping' / 'xbrl_mappings_complete.json'
    new_json = base_dir / 'data' / 'mapping' / 'mappings.json'

    print("=" * 80)
    print("T0006 FIX VERIFICATION")
    print("=" * 80)
    print()

    # Load files
    with open(old_json, 'r', encoding='utf-8') as f:
        old_data = json.load(f)
    with open(new_json, 'r', encoding='utf-8') as f:
        new_data = json.load(f)

    # Get T0006 data
    old_t0006 = old_data.get('mappature', {}).get('T0006', [])
    new_t0006 = new_data.get('mappature', {}).get('T0006', [])

    print(f"Old JSON: {len(old_t0006)} entries")
    print(f"New JSON: {len(new_t0006)} entries")
    print(f"Difference: +{len(new_t0006) - len(old_t0006)} entries")
    print()

    # Check the 4 critical codes that were wrong
    critical_codes = [
        'T0006.D01.1.001.002.003.001.000',  # Variazioni rimanenze
        'T0006.D01.1.001.004.000.000.000',  # Differenza A-B
        'T0006.D01.1.001.007.000.000.000',  # Risultato prima imposte
        'T0006.D01.1.001.009.000.000.000',  # Utile/perdita esercizio
    ]

    print("CRITICAL FIXES:")
    print("-" * 80)

    for code in critical_codes:
        # Find in old
        old_entry = next((e for e in old_t0006 if e.get('codice_excel') == code), None)
        # Find in new
        new_entry = next((e for e in new_t0006 if e.get('code') == code), None)

        if old_entry and new_entry:
            old_abstract = old_entry.get('ui', {}).get('is_abstract')
            new_abstract = new_entry.get('ui', {}).get('is_abstract')

            status = "✓ FIXED" if old_abstract != new_abstract else "  (no change)"

            print(f"\n{code}")
            print(f"  Label: {new_entry['ui']['label'][:60]}")
            print(f"  Old is_abstract: {old_abstract}")
            print(f"  New is_abstract: {new_abstract}  {status}")
            print(f"  XBRL type: {new_entry.get('xbrl', {}).get('type', 'N/A')}")

    print()
    print("-" * 80)

    # Count abstract entries
    old_abstract_count = sum(1 for e in old_t0006 if e.get('ui', {}).get('is_abstract'))
    new_abstract_count = sum(1 for e in new_t0006 if e.get('ui', {}).get('is_abstract'))

    print()
    print("ABSTRACT ENTRIES (headers):")
    print(f"  Old: {old_abstract_count}")
    print(f"  New: {new_abstract_count}  (+{new_abstract_count - old_abstract_count})")
    print()

    print("=" * 80)
    print("✓ Verification complete!")
    print("=" * 80)

if __name__ == '__main__':
    main()
