#!/usr/bin/env python3
"""
Test sparse conversion to ensure no data loss.

This script:
1. Loads selected tables from workbookabb.json
2. Converts them to sparse format
3. Converts back to dense format
4. Verifies they are identical to the original
"""

import json
import sys
from pathlib import Path


def dense_to_sparse(dense_array):
    """Convert dense 2D array to sparse format."""
    if not dense_array:
        return {'meta': {'rows': 0, 'cols': 0}, 'data': {}}

    rows = len(dense_array)
    cols = len(dense_array[0]) if dense_array else 0

    sparse_data = {}
    for r, row in enumerate(dense_array):
        for c, cell in enumerate(row):
            if cell is not None:
                if str(r) not in sparse_data:
                    sparse_data[str(r)] = {}
                sparse_data[str(r)][str(c)] = cell

    return {
        'meta': {'rows': rows, 'cols': cols},
        'data': sparse_data
    }


def sparse_to_dense(sparse_obj):
    """Convert sparse format back to dense 2D array."""
    meta = sparse_obj.get('meta', {})
    sparse_data = sparse_obj.get('data', {})

    rows = meta.get('rows', 0)
    cols = meta.get('cols', 0)

    # Create dense array filled with None
    dense = [[None for _ in range(cols)] for _ in range(rows)]

    # Fill in non-null values
    for r_str, row_data in sparse_data.items():
        r = int(r_str)
        for c_str, value in row_data.items():
            c = int(c_str)
            dense[r][c] = value

    return dense


def arrays_equal(arr1, arr2):
    """Deep equality check for 2D arrays."""
    if len(arr1) != len(arr2):
        return False

    for r in range(len(arr1)):
        if len(arr1[r]) != len(arr2[r]):
            return False
        for c in range(len(arr1[r])):
            if arr1[r][c] != arr2[r][c]:
                return False

    return True


def test_table(table_name, original_data):
    """Test sparse conversion for a single table."""
    print(f"\n{'='*60}")
    print(f"Testing: {table_name}")
    print('='*60)

    # Original stats
    rows = len(original_data)
    cols = len(original_data[0]) if original_data else 0
    non_null = sum(1 for row in original_data for cell in row if cell is not None)
    total = rows * cols

    print(f"Original: {rows} rows × {cols} cols")
    print(f"Non-null cells: {non_null}/{total} ({100*non_null/total if total else 0:.1f}%)")

    # Convert to sparse
    sparse = dense_to_sparse(original_data)

    # Size comparison
    original_json = json.dumps(original_data)
    sparse_json = json.dumps(sparse)

    original_size = len(original_json)
    sparse_size = len(sparse_json)
    savings = 100 * (original_size - sparse_size) / original_size if original_size else 0

    print(f"\nSize comparison:")
    print(f"  Dense:  {original_size:>10,} bytes ({original_size/1024:>7.1f} KB)")
    print(f"  Sparse: {sparse_size:>10,} bytes ({sparse_size/1024:>7.1f} KB)")
    print(f"  Savings: {original_size - sparse_size:>9,} bytes ({savings:>5.1f}%)")

    # Convert back to dense
    reconstructed = sparse_to_dense(sparse)

    # Verify equality
    if arrays_equal(original_data, reconstructed):
        print(f"\n✓ PASS: Reconstructed data matches original exactly")
        return True
    else:
        print(f"\n✗ FAIL: Reconstructed data differs from original!")

        # Find first difference
        for r in range(min(len(original_data), len(reconstructed))):
            for c in range(min(len(original_data[r]), len(reconstructed[r]))):
                if original_data[r][c] != reconstructed[r][c]:
                    print(f"  First diff at [{r},{c}]:")
                    print(f"    Original: {original_data[r][c]}")
                    print(f"    Reconstructed: {reconstructed[r][c]}")
                    return False

        return False


def main():
    """Main test execution."""
    print("="*60)
    print("SPARSE CONVERSION TEST")
    print("="*60)
    print()

    # Load workbookabb.json
    base_dir = Path(__file__).parent.parent
    json_path = base_dir / 'data' / 'template' / 'workbookabb.json'

    if not json_path.exists():
        print(f"✗ Error: {json_path} not found")
        return 1

    print(f"Loading: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    sheets = data.get('sheets', {})
    print(f"Loaded {len(sheets)} sheets")

    # Test tables of different types
    test_cases = [
        ('T0000', 'Informazioni generali - tipo 1'),
        ('T0002', 'Stato patrimoniale abbreviato - tipo 1'),
        ('T0006', 'Conto economico abbreviato - tipo 1'),
        ('T0009', 'Rendiconto finanziario - tipo 2'),
        ('T0154', 'Nota integrativa - tipo 2'),
        ('T0151', 'Nota integrativa - tipo 1'),
    ]

    results = []

    for table_code, description in test_cases:
        if table_code not in sheets:
            print(f"\n⚠ Warning: {table_code} not found in sheets")
            continue

        table_data = sheets[table_code]
        result = test_table(f"{table_code} ({description})", table_data)
        results.append((table_code, result))

    # Summary
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)

    passed = sum(1 for _, r in results if r)
    total = len(results)

    for table_code, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"  {status}  {table_code}")

    print()
    print(f"Total: {passed}/{total} tests passed")

    if passed == total:
        print("\n✓ All tests passed! Sparse conversion is safe.")
        return 0
    else:
        print(f"\n✗ {total - passed} test(s) failed!")
        return 1


if __name__ == '__main__':
    sys.exit(main())
