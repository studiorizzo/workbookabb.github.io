// workbookabb - Sparse Array Loader Module
//
// Provides transparent conversion from sparse format to dense arrays.
// Lazy conversion: only converts sheets/rows when accessed.

/**
 * Convert sparse sheet format to dense 2D array.
 *
 * @param {Object} sparseSheet - Sparse format: {meta: {rows, cols}, data: {row: {col: val}}}
 * @returns {Array} Dense 2D array
 */
function sparseToDense(sparseSheet) {
    if (!sparseSheet || !sparseSheet.meta) {
        console.warn('Invalid sparse sheet format');
        return [];
    }

    const { rows, cols } = sparseSheet.meta;
    const sparseData = sparseSheet.data || {};

    // Create dense array filled with null
    const dense = new Array(rows);
    for (let r = 0; r < rows; r++) {
        dense[r] = new Array(cols).fill(null);
    }

    // Fill in non-null values from sparse data
    for (const [rowStr, rowData] of Object.entries(sparseData)) {
        const r = parseInt(rowStr);
        if (r >= 0 && r < rows) {
            for (const [colStr, value] of Object.entries(rowData)) {
                const c = parseInt(colStr);
                if (c >= 0 && c < cols) {
                    dense[r][c] = value;
                }
            }
        }
    }

    return dense;
}

/**
 * Create a lazy-loading wrapper for sparse sheets.
 * Converts to dense format only when the sheet is first accessed.
 *
 * @param {Object} sparseSheets - Object with sheet codes as keys
 * @returns {Object} Wrapper object with lazy conversion
 */
function createLazySparseLoader(sparseSheets) {
    const cache = new Map();
    let stats = {
        totalSheets: Object.keys(sparseSheets).length,
        loadedSheets: 0,
        cacheHits: 0
    };

    return new Proxy(sparseSheets, {
        get(target, sheetCode) {
            // Return stats if requested
            if (sheetCode === '_stats') {
                return stats;
            }

            // Check if sheet exists
            if (!(sheetCode in target)) {
                return undefined;
            }

            // Check cache first
            if (cache.has(sheetCode)) {
                stats.cacheHits++;
                return cache.get(sheetCode);
            }

            // Convert sparse to dense on first access
            const sparseSheet = target[sheetCode];
            const denseArray = sparseToDense(sparseSheet);

            // Cache the result
            cache.set(sheetCode, denseArray);
            stats.loadedSheets++;

            console.log(`✓ Loaded sheet ${sheetCode} (${stats.loadedSheets}/${stats.totalSheets})`);

            return denseArray;
        },

        // Support Object.keys(), for...in, etc.
        ownKeys(target) {
            return Reflect.ownKeys(target);
        },

        has(target, key) {
            return key in target;
        },

        getOwnPropertyDescriptor(target, key) {
            return Reflect.getOwnPropertyDescriptor(target, key);
        }
    });
}

/**
 * Check if workbookData uses sparse format.
 *
 * @param {Object} workbookData - The loaded workbook data
 * @returns {boolean} True if sparse format
 */
function isSparseFormat(workbookData) {
    if (!workbookData) return false;

    // Check for metadata indicating sparse format
    if (workbookData.metadata && workbookData.metadata.format === 'sparse') {
        return true;
    }

    // Check if first sheet has sparse structure
    const sheets = workbookData.sheets;
    if (sheets) {
        const firstSheetKey = Object.keys(sheets)[0];
        if (firstSheetKey) {
            const firstSheet = sheets[firstSheetKey];
            // Sparse format has meta and data properties
            if (firstSheet && typeof firstSheet === 'object' &&
                'meta' in firstSheet && 'data' in firstSheet) {
                return true;
            }
        }
    }

    return false;
}

/**
 * Load and prepare workbook data (handles both dense and sparse formats).
 *
 * @param {Object} rawData - Raw JSON data from fetch
 * @returns {Object} Processed workbook data with sheets ready for use
 */
function prepareWorkbookData(rawData) {
    if (!rawData) {
        throw new Error('No workbook data provided');
    }

    // Check format
    const isSparse = isSparseFormat(rawData);

    if (isSparse) {
        console.log('✓ Sparse format detected, using lazy loader');

        // Wrap sheets with lazy loader
        const lazySheets = createLazySparseLoader(rawData.sheets);

        return {
            config: rawData.config,
            index: rawData.index,
            sheets: lazySheets,
            metadata: rawData.metadata,
            _format: 'sparse'
        };
    } else {
        console.log('✓ Dense format detected, using as-is');

        return {
            config: rawData.config,
            index: rawData.index,
            sheets: rawData.sheets,
            _format: 'dense'
        };
    }
}

/**
 * Get loader statistics (only for sparse format).
 *
 * @param {Object} workbookData - Prepared workbook data
 * @returns {Object|null} Statistics or null if not sparse
 */
function getSparseStats(workbookData) {
    if (workbookData && workbookData._format === 'sparse' && workbookData.sheets) {
        return workbookData.sheets._stats;
    }
    return null;
}

// Export functions for use in app.js
// (In a module system, you'd use export { ... })
if (typeof window !== 'undefined') {
    window.SparseLoader = {
        sparseToDense,
        createLazySparseLoader,
        isSparseFormat,
        prepareWorkbookData,
        getSparseStats
    };
}
