#!/usr/bin/env python3
"""
Generate sparse format workbookabb.json from dense format.

This script converts the template file from dense 2D arrays (99% null cells)
to sparse format (only non-null cells stored), reducing file size by ~95%.

Input:  data/template/workbookabb.json (44 MB dense)
Output: data/template/workbookabb-sparse.json (~0.9 MB sparse)
"""

import json
import sys
from pathlib import Path
from datetime import datetime


def dense_to_sparse(dense_array):
    """
    Convert dense 2D array to sparse format.

    Args:
        dense_array: 2D array with mostly null cells

    Returns:
        dict: {
            'meta': {'rows': int, 'cols': int},
            'data': {'row_idx': {'col_idx': value, ...}, ...}
        }
    """
    if not dense_array:
        return {'meta': {'rows': 0, 'cols': 0}, 'data': {}}

    rows = len(dense_array)
    cols = len(dense_array[0]) if dense_array else 0

    sparse_data = {}
    cell_count = 0

    for r, row in enumerate(dense_array):
        row_data = {}
        for c, cell in enumerate(row):
            if cell is not None:
                row_data[str(c)] = cell
                cell_count += 1

        if row_data:  # Only store rows with data
            sparse_data[str(r)] = row_data

    return {
        'meta': {'rows': rows, 'cols': cols},
        'data': sparse_data,
        '_cells': cell_count  # For statistics
    }


def convert_sheets_to_sparse(sheets_dict):
    """Convert all sheets to sparse format."""
    print("Converting sheets to sparse format...")

    sparse_sheets = {}
    total_dense_size = 0
    total_sparse_size = 0
    total_cells = 0
    total_non_null = 0

    sheet_count = len(sheets_dict)

    for idx, (sheet_name, sheet_data) in enumerate(sheets_dict.items(), 1):
        if idx % 20 == 0 or idx == sheet_count:
            print(f"  Processing {idx}/{sheet_count} sheets...", end='\r')

        # Convert to sparse
        sparse = dense_to_sparse(sheet_data)
        sparse_sheets[sheet_name] = {
            'meta': sparse['meta'],
            'data': sparse['data']
        }

        # Statistics
        rows = sparse['meta']['rows']
        cols = sparse['meta']['cols']
        cells = rows * cols
        non_null = sparse['_cells']

        total_cells += cells
        total_non_null += non_null

        dense_size = len(json.dumps(sheet_data))
        sparse_size = len(json.dumps(sparse_sheets[sheet_name]))

        total_dense_size += dense_size
        total_sparse_size += sparse_size

    print()  # New line after progress
    print(f"✓ Converted {sheet_count} sheets")
    print(f"  Total cells: {total_cells:,}")
    print(f"  Non-null cells: {total_non_null:,} ({100*total_non_null/total_cells:.1f}%)")
    print(f"  Null cells: {total_cells - total_non_null:,} ({100*(total_cells-total_non_null)/total_cells:.1f}%)")
    print()
    print(f"  Dense size: {total_dense_size:,} bytes ({total_dense_size/1024/1024:.2f} MB)")
    print(f"  Sparse size: {total_sparse_size:,} bytes ({total_sparse_size/1024/1024:.2f} MB)")
    print(f"  Savings: {total_dense_size - total_sparse_size:,} bytes ({100*(total_dense_size-total_sparse_size)/total_dense_size:.1f}%)")

    return sparse_sheets


def main():
    """Main execution."""
    print("=" * 80)
    print("SPARSE TEMPLATE GENERATOR")
    print("=" * 80)
    print()

    base_dir = Path(__file__).parent.parent
    input_path = base_dir / 'data' / 'template' / 'workbookabb.json'
    output_path = base_dir / 'data' / 'template' / 'workbookabb-sparse.json'

    # Verify input exists
    if not input_path.exists():
        print(f"✗ Error: {input_path} not found")
        return 1

    # Load dense JSON
    print(f"Loading: {input_path}")
    input_size = input_path.stat().st_size
    print(f"  Size: {input_size:,} bytes ({input_size/1024/1024:.2f} MB)")
    print()

    with open(input_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f"Loaded JSON structure:")
    print(f"  Keys: {list(data.keys())}")
    print(f"  Sheets: {len(data.get('sheets', {}))}")
    print()

    # Convert sheets to sparse
    sparse_sheets = convert_sheets_to_sparse(data.get('sheets', {}))

    # Build output structure
    output_data = {
        'metadata': {
            'format': 'sparse',
            'version': '2.0',
            'generated': datetime.now().isoformat(),
            'description': 'Sparse format template - only non-null cells stored',
            'original_file': 'workbookabb.json (dense format)',
            'conversion_script': 'scripts/generate_sparse_template.py'
        },
        'config': data.get('config'),
        'index': data.get('index'),
        'sheets': sparse_sheets
    }

    # Write output
    print()
    print(f"Writing: {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=None, separators=(',', ':'))

    output_size = output_path.stat().st_size
    print(f"  Size: {output_size:,} bytes ({output_size/1024/1024:.2f} MB)")
    print()

    # Final summary
    total_savings = input_size - output_size
    savings_pct = 100 * total_savings / input_size

    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Input:   {input_size:>15,} bytes ({input_size/1024/1024:>7.2f} MB)")
    print(f"Output:  {output_size:>15,} bytes ({output_size/1024/1024:>7.2f} MB)")
    print(f"Savings: {total_savings:>15,} bytes ({total_savings/1024/1024:>7.2f} MB)")
    print(f"Reduction: {savings_pct:>13.1f}%")
    print()
    print("✓ Sparse template generated successfully!")
    print()
    print("Next steps:")
    print("  1. Update app.js to use sparse-loader.js")
    print("  2. Test with sample tables")
    print("  3. Replace workbookabb.json with workbookabb-sparse.json")

    return 0


if __name__ == '__main__':
    sys.exit(main())
