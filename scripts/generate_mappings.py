#!/usr/bin/env python3
"""
Generate optimized mappings.json from XBRL taxonomy XML files.

This script reads:
- mapping.xml: XBRL metadata (name, type, prefix, period_type)
- dimension.xml: UI metadata (labels, hierarchy levels)

And generates a clean, optimized JSON structure.
"""

import xml.etree.ElementTree as ET
import json
from pathlib import Path
from datetime import datetime
from collections import defaultdict


def parse_mapping_xml(xml_path):
    """
    Parse mapping.xml to extract XBRL metadata for each code.

    Returns:
        dict: {code: {name, type, prefix, period_type, def_code}}
    """
    print(f"Parsing {xml_path}...")
    tree = ET.parse(xml_path)
    root = tree.getroot()

    mappings = {}

    # Find all <report> elements
    for report in root.findall('.//report'):
        report_code = report.get('code')
        if not report_code:
            continue

        # Find all <cell> elements within this report
        for cell in report.findall('.//cell'):
            code = cell.get('code')
            if not code:
                continue

            # Extract XBRL attributes
            xbrl_name = cell.get('{http://www.xbrl.org}name')
            xbrl_type = cell.get('{http://www.xbrl.org}type')
            xbrl_prefix = cell.get('{http://www.xbrl.org}prefix')
            xbrl_period_type = cell.get('{http://www.xbrl.org}periodType')
            def_code = cell.get('def_code')

            mappings[code] = {
                'name': xbrl_name,
                'type': xbrl_type,
                'prefix': xbrl_prefix,
                'period_type': xbrl_period_type,
                'def_code': def_code,
                'report': report_code
            }

    print(f"  Found {len(mappings)} XBRL mappings")
    return mappings


def parse_dimension_xml(xml_path):
    """
    Parse dimension.xml to extract UI metadata (labels, indent levels).

    Returns:
        dict: {code: {label, indent_level, type, order}}
    """
    print(f"Parsing {xml_path}...")
    tree = ET.parse(xml_path)
    root = tree.getroot()

    dimensions = {}

    # Recursive function to traverse the hierarchy
    def traverse_children(element, parent_report_code, depth=0):
        """Traverse child elements and extract metadata."""
        for child in element.findall('child'):
            code = child.get('code')
            name = child.get('name')
            child_type = child.get('type')  # 'abstract', 'item', 'group'
            level = int(child.get('level', 0))
            order = int(child.get('order', 0))
            fullname = child.get('fullname', name)

            if code:
                # Calculate indent level based on hierarchy level
                # Level typically starts at 2 for root items
                indent_level = max(0, level - 2)

                dimensions[code] = {
                    'label': name,
                    'fullname': fullname,
                    'indent_level': indent_level,
                    'dim_type': child_type,  # from dimension.xml
                    'order': order,
                    'report': parent_report_code
                }

            # Recurse into nested children
            traverse_children(child, parent_report_code, depth + 1)

    # Find all <report> elements
    for report in root.findall('.//report'):
        report_code = report.get('code')
        if not report_code:
            continue

        # Find all <dimension> elements
        for dimension in report.findall('.//dimension'):
            traverse_children(dimension, report_code)

    print(f"  Found {len(dimensions)} dimension entries")
    return dimensions


def merge_mappings(xbrl_mappings, ui_dimensions):
    """
    Merge XBRL and UI metadata into final structure.

    Args:
        xbrl_mappings: dict from parse_mapping_xml
        ui_dimensions: dict from parse_dimension_xml

    Returns:
        dict: {report_code: [mapping_entries]}
    """
    print("Merging mappings...")

    merged = defaultdict(list)
    all_codes = set(xbrl_mappings.keys()) | set(ui_dimensions.keys())

    for code in sorted(all_codes):
        xbrl = xbrl_mappings.get(code)
        ui = ui_dimensions.get(code)

        # Determine report code
        report_code = None
        if xbrl and xbrl.get('report'):
            report_code = xbrl['report']
        elif ui and ui.get('report'):
            report_code = ui['report']
        else:
            # Try to extract from code (e.g., T0006.D01.1.001...)
            parts = code.split('.')
            if parts:
                report_code = parts[0]

        if not report_code:
            continue

        # Build entry
        entry = {
            'code': code
        }

        # UI section
        ui_data = {}
        if ui:
            # Use 'name' (short label) for UI, not 'fullname' (full path)
            ui_data['label'] = ui.get('label', '')
            ui_data['indent_level'] = ui.get('indent_level', 0)
        else:
            # Fallback for codes without dimension data
            ui_data['label'] = ''
            ui_data['indent_level'] = 0

        # Determine is_abstract from XBRL type or dimension type
        # Both 'abstract' and 'group' types should be rendered as headers (no input fields)
        if xbrl and xbrl.get('type'):
            ui_data['is_abstract'] = (xbrl['type'] == 'abstract')
        elif ui and ui.get('dim_type') in ('abstract', 'group'):
            ui_data['is_abstract'] = True
        else:
            ui_data['is_abstract'] = False

        entry['ui'] = ui_data

        # XBRL section (only if exists)
        if xbrl and xbrl.get('name'):
            xbrl_data = {
                'name': xbrl['name']
            }
            # Only add non-null fields to reduce size
            if xbrl.get('prefix'):
                xbrl_data['prefix'] = xbrl['prefix']
            if xbrl.get('type'):
                xbrl_data['type'] = xbrl['type']
            if xbrl.get('period_type'):
                xbrl_data['period_type'] = xbrl['period_type']
            if xbrl.get('def_code'):
                xbrl_data['def_code'] = xbrl['def_code']
            entry['xbrl'] = xbrl_data
        else:
            # Don't include xbrl key if no data
            pass

        merged[report_code].append(entry)

    # Convert defaultdict to regular dict and sort entries by code
    result = {}
    for report_code, entries in sorted(merged.items()):
        result[report_code] = sorted(entries, key=lambda x: x['code'])

    total_entries = sum(len(entries) for entries in result.values())
    print(f"  Merged {total_entries} entries across {len(result)} reports")

    return result


def generate_json(mappings, output_path):
    """Generate final JSON file with metadata."""
    print(f"Generating {output_path}...")

    output = {
        'metadata': {
            'generated': datetime.now().isoformat(),
            'version': '2.0',
            'source_files': [
                'data/taxonomy/mapping.xml',
                'data/taxonomy/dimension.xml'
            ],
            'description': 'XBRL mappings for Italian GAAP financial statements (Principi Contabili Italiani)'
        },
        'mappature': mappings
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    # Calculate size
    size_kb = Path(output_path).stat().st_size / 1024
    print(f"  ✓ Generated {output_path} ({size_kb:.1f} KB)")


def validate_output(mappings, sample_table='T0006'):
    """Quick validation of output."""
    print(f"\nValidating output (sample: {sample_table})...")

    if sample_table not in mappings:
        print(f"  ⚠ Warning: {sample_table} not found in mappings")
        return

    entries = mappings[sample_table]
    print(f"  {sample_table}: {len(entries)} entries")

    # Count abstract vs data entries
    abstract_count = sum(1 for e in entries if e['ui']['is_abstract'])
    data_count = len(entries) - abstract_count
    xbrl_count = sum(1 for e in entries if e.get('xbrl'))

    print(f"    - Abstract entries: {abstract_count}")
    print(f"    - Data entries: {data_count}")
    print(f"    - With XBRL metadata: {xbrl_count}")

    # Show a few sample entries
    print(f"\n  Sample entries:")
    for entry in entries[:3]:
        abstract_str = " (abstract)" if entry['ui']['is_abstract'] else ""
        xbrl_str = f" → {entry.get('xbrl', {}).get('name', 'no XBRL')}" if entry.get('xbrl') else " (no XBRL)"
        print(f"    {entry['code']}: {entry['ui']['label'][:50]}{abstract_str}{xbrl_str}")


def main():
    """Main execution."""
    print("=" * 80)
    print("XBRL Mappings Generator v2.0")
    print("=" * 80)
    print()

    # Paths
    base_dir = Path(__file__).parent.parent
    mapping_xml = base_dir / 'data' / 'taxonomy' / 'mapping.xml'
    dimension_xml = base_dir / 'data' / 'taxonomy' / 'dimension.xml'
    output_json = base_dir / 'data' / 'mapping' / 'mappings.json'

    # Verify input files exist
    if not mapping_xml.exists():
        print(f"✗ Error: {mapping_xml} not found")
        return 1
    if not dimension_xml.exists():
        print(f"✗ Error: {dimension_xml} not found")
        return 1

    # Parse XML files
    xbrl_mappings = parse_mapping_xml(mapping_xml)
    ui_dimensions = parse_dimension_xml(dimension_xml)

    # Merge data
    merged_mappings = merge_mappings(xbrl_mappings, ui_dimensions)

    # Generate JSON
    generate_json(merged_mappings, output_json)

    # Validate
    validate_output(merged_mappings)

    print()
    print("=" * 80)
    print("✓ Generation complete!")
    print("=" * 80)

    return 0


if __name__ == '__main__':
    exit(main())
