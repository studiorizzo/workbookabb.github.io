# XBRL Mapping Scripts

This directory contains Python scripts for generating and validating XBRL mappings.

## Scripts

### `generate_mappings.py`

Generates the optimized `mappings.json` file from the XBRL taxonomy XML files.

**Input:**
- `data/taxonomy/mapping.xml` - XBRL metadata (names, types, prefixes)
- `data/taxonomy/dimension.xml` - UI metadata (labels, hierarchy, indent levels)

**Output:**
- `data/mapping/mappings.json` - Optimized JSON mapping file

**Usage:**
```bash
python3 scripts/generate_mappings.py
```

**What it does:**
1. Parses mapping.xml to extract XBRL element metadata
2. Parses dimension.xml to extract UI labels and hierarchy
3. Merges the data, determining `is_abstract` from XBRL types
4. Generates clean JSON structure without redundant fields
5. Validates output for sample tables

**Structure improvements over old format:**
- ✅ Correct `is_abstract` determination from XBRL type
- ✅ Proper UI labels from dimension hierarchy
- ✅ All abstract (header) entries included
- ✅ Cleaner structure (removed redundant coordinate data)
- ✅ Complete coverage (465 tables vs 5 in old version)

### `verify_t0006_fix.py`

Verifies that the critical T0006 issues have been fixed in the new mappings.json.

**Usage:**
```bash
python3 scripts/verify_t0006_fix.py
```

**What it checks:**
- The 4 critical monetary fields that were incorrectly marked as abstract
- Abstract entry count (headers)
- Total entry count

## Requirements

- Python 3.6+
- No external dependencies (uses stdlib only)

## Regenerating Mappings

If the taxonomy XML files are updated, regenerate the mappings:

```bash
python3 scripts/generate_mappings.py
```

Then test the app to ensure everything works correctly.
