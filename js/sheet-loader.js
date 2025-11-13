// workbookabb - Sheet Loader Module (SheetJS) - VERSIONE MIGLIORATA

// Parse file XLS/XLSX
async function parseXLS(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {
                    type: 'array',
                    cellFormula: false,
                    cellStyles: false
                });
                
                resolve(workbook);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Errore lettura file'));
        reader.readAsArrayBuffer(file);
    });
}

// Importa da XLS a struttura bilancio
async function importFromXLS(workbook) {
    const bilancio = {
        metadata: {
            versione: '1.0',
            data_creazione: new Date().toISOString(),
            data_modifica: new Date().toISOString(),
            ragione_sociale: null,
            anno_esercizio: new Date().getFullYear()
        },
        fogli: {}
    };
    
    // IMPORTA CONFIGURAZIONE (date, CF, valuta)
    importConfigurazione(workbook, bilancio);
    
    let celleImportate = 0;
    let fogliImportati = 0;
    
    // Per ogni foglio nel workbook
    for (const sheetName of workbook.SheetNames) {
        try {
            // Carica template per questo foglio
            const template = getTemplate(sheetName);
            if (!template || !template.loaded) {
                console.warn(`Template ${sheetName} non disponibile, skip`);
                continue;
            }
            
            const sheet = workbook.Sheets[sheetName];
            const templateData = template.data;

            // Parse configurazione (righe 0-1) in oggetto
            const configMap = parseSheetConfig(templateData);
            if (!configMap || !configMap.tipo_tab) {
                console.warn(`Template ${sheetName} configurazione invalida`);
                continue;
            }

            const tipoTab = parseInt(configMap.tipo_tab);

            bilancio.fogli[sheetName] = {};
            let celleSheet = 0;

            if (tipoTab === 1) {
                // TIPO 1: Foglio semplice
                celleSheet = importTipo1(sheet, templateData, configMap, bilancio.fogli[sheetName], sheetName);
            } else if (tipoTab === 2) {
                // TIPO 2: Tabella 2D
                celleSheet = importTipo2(sheet, templateData, configMap, bilancio.fogli[sheetName]);
            } else if (tipoTab === 3) {
                // TIPO 3: Tuple
                celleSheet = importTipo3(sheet, templateData, configMap, bilancio.fogli[sheetName]);
            } else {
                console.warn(`Tipo ${tipoTab} non supportato per ${sheetName}`);
                continue;
            }
            
            if (celleSheet > 0) {
                celleImportate += celleSheet;
                fogliImportati++;
                console.log(`✓ Imported ${sheetName}: ${celleSheet} celle (tipo ${tipoTab})`);
            }
            
        } catch (error) {
            console.error(`Error importing ${sheetName}:`, error);
        }
    }
    
    console.log(`✓ Import completato: ${fogliImportati} fogli, ${celleImportate} celle totali`);
    
    return bilancio;
}

// Import TIPO 1: Foglio semplice
function importTipo1(sheet, templateData, configMap, datiSheet, sheetName) {
    // Usa configMap invece di indici hardcoded
    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;

    let count = 0;

    // Se 1 riga × 1 colonna = textBlock
    if (numRows === 1 && numCols === 1) {
        // Il codice è in riga colCodeRow, colonna firstCol
        const codiceCell = templateData[colCodeRow]?.[firstCol];
        if (codiceCell) {
            const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c: firstCol });
            const cell = sheet[cellAddress];
            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                datiSheet[codiceCell] = cell.v;
                count++;
            }
        }
    } else {
        // Determina se è T0000 (importa solo prima colonna)
        const isT0000 = sheetName === 'T0000';
        const colsToImport = isT0000 ? 1 : numCols;

        // Tabella semplice con codici riga
        for (let r = 0; r < numRows; r++) {
            const rowData = templateData[firstRow + r];
            if (!rowData) continue;

            // Leggi row_code dalla colonna specificata
            const codiceRiga = rowData[rowCodeCol];
            if (!codiceRiga) continue;

            // Per T0000: importa solo la prima colonna (corrente)
            // Per altri fogli tipo 1: importa tutte le colonne (sovrascrivendo)
            for (let c = 0; c < colsToImport; c++) {
                const xlsRow = firstRow + r;
                const xlsCol = firstCol + c;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });

                const cell = sheet[cellAddress];
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    // Usa codice_riga direttamente (no suffisso colonna)
                    datiSheet[codiceRiga] = cell.v;
                    count++;
                }
            }
        }
    }

    return count;
}

// Import TIPO 2: Tabella 2D
function importTipo2(sheet, templateData, configMap, datiSheet) {
    // Usa configMap invece di indici hardcoded (come Java getSheetMap)
    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;

    let count = 0;

    for (let r = 0; r < numRows; r++) {
        const rowData = templateData[firstRow + r];
        if (!rowData) continue;

        // Leggi row_code dalla colonna specificata (non hardcoded a 0!)
        const codiceRiga = rowData[rowCodeCol];
        if (!codiceRiga) continue;

        for (let c = 0; c < numCols; c++) {
            // Leggi col_code dalla riga specificata
            const colCodeData = templateData[colCodeRow];
            if (!colCodeData) continue;

            const codiceColonna = colCodeData[firstCol + c];
            if (!codiceColonna) continue;

            const codiceCella = `${codiceRiga}_${codiceColonna}`;

            const xlsRow = firstRow + r;
            const xlsCol = firstCol + c;
            const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });

            const cell = sheet[cellAddress];
            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                datiSheet[codiceCella] = cell.v;
                count++;
            }
        }
    }

    return count;
}

// Import TIPO 3: Tuple
function importTipo3(sheet, templateData, configMap, datiSheet) {
    // Usa configMap invece di indici hardcoded
    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;

    let count = 0;

    for (let r = 0; r < numRows; r++) {
        const rowData = templateData[firstRow + r];
        if (!rowData) continue;

        // Leggi row_code dalla colonna specificata
        const codiceRiga = rowData[rowCodeCol];
        if (!codiceRiga) continue;

        // Se una sola colonna, usa codiceRiga diretto
        if (numCols === 1) {
            const xlsRow = firstRow + r;
            const xlsCol = firstCol;
            const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
            const cell = sheet[cellAddress];

            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                datiSheet[codiceRiga] = cell.v;
                count++;
            }
        } else {
            // Più colonne: leggi col_code dalla riga specificata
            for (let c = 0; c < numCols; c++) {
                const colCodeData = templateData[colCodeRow];
                if (!colCodeData) continue;

                const codiceColonna = colCodeData[firstCol + c];
                if (!codiceColonna) continue;

                const codiceCella = `${codiceRiga}_${codiceColonna}`;
                const xlsRow = firstRow + r;
                const xlsCol = firstCol + c;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });

                const cell = sheet[cellAddress];
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    datiSheet[codiceCella] = cell.v;
                    count++;
                }
            }
        }
    }

    return count;
}

// Estrai dati foglio XLS (helper)
function getSheetData(sheet) {
    if (!sheet || !sheet['!ref']) return [];
    
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const data = [];
    
    for (let R = range.s.r; R <= range.e.r; R++) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = sheet[cellAddress];
            row.push(cell ? (cell.v || null) : null);
        }
        data.push(row);
    }
    
    return data;
}

// Get cell value (helper)
function getCellValue(sheet, row, col) {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = sheet[cellAddress];
    return cell ? (cell.v || null) : null;
}

// Parse sheet configuration (righe 0-1) in oggetto HashMap
// Come getSheetMap() nell'app Java
function parseSheetConfig(templateData) {
    if (!templateData || templateData.length < 2) {
        return null;
    }

    const keys = templateData[0];   // Riga 0: nomi parametri
    const values = templateData[1]; // Riga 1: valori

    if (!Array.isArray(keys) || !Array.isArray(values)) {
        return null;
    }

    const config = {};
    for (let i = 0; i < keys.length && i < values.length; i++) {
        if (keys[i]) {
            config[keys[i]] = values[i];
        }
    }

    return config;
}

// Mappa Named Ranges → coordinate celle
// Basata sull'analisi dell'Excel originale e del codice Java
function getNamedRangeLocations() {
    return {
        // Date INPUT utente (da foglio index)
        'c_this_end_input': { sheet: 'index', row: 2, col: 2 },     // C3
        'c_this_start_import': { sheet: 'index', row: 2, col: 6 },  // G3
        'c_prev_end_import': { sheet: 'index', row: 3, col: 8 },    // I4
        'c_prev_start_import': { sheet: 'index', row: 3, col: 6 },  // G4

        // Date calcolate (da foglio config)
        'c_this_end': { sheet: 'config', row: 8, col: 2 },   // C9
        'c_this_start': { sheet: 'config', row: 8, col: 3 }, // D9
        'c_this': { sheet: 'config', row: 8, col: 4 },       // E9
        'c_prev_end': { sheet: 'config', row: 9, col: 2 },   // C10
        'c_prev_start': { sheet: 'config', row: 9, col: 3 }, // D10
        'c_prev': { sheet: 'config', row: 9, col: 4 },       // E10

        // Altri
        'cf': { sheet: 'config', row: 13, col: 2 },          // C14 - Codice Fiscale
        'unit': { sheet: 'config', row: 11, col: 2 }         // C12 - Valuta
    };
}

// Risolvi Named Range: legge valore da Excel workbook
function resolveNamedRange(workbook, rangeName) {
    const locations = getNamedRangeLocations();
    const location = locations[rangeName];

    if (!location) {
        console.warn(`Named range "${rangeName}" non trovato`);
        return null;
    }

    // Trova il foglio
    const sheet = workbook.Sheets[location.sheet];
    if (!sheet) {
        console.warn(`Foglio "${location.sheet}" non trovato per named range "${rangeName}"`);
        return null;
    }

    // Leggi valore dalla cella
    return getCellValue(sheet, location.row, location.col);
}

// Importa configurazione usando Named Ranges (come app Java originale)
function importConfigurazione(workbook, bilancio) {
    console.log('📋 Import configurazione usando Named Ranges...');

    try {
        // STRATEGIA: Leggi date dai Named Ranges come fa l'app Java
        // Priorità 1: Foglio Indice (INPUT utente)
        // Priorità 2: Foglio Configurazione (valori calcolati) - fallback

        // Leggi date corrente
        let fineCorrenteSerial = resolveNamedRange(workbook, 'c_this_end_input');
        let inizioCorrenteSerial = resolveNamedRange(workbook, 'c_this_start_import');
        let annoCorrenteStr = resolveNamedRange(workbook, 'c_this');

        // Fallback a Configurazione se Indice non disponibile
        if (!fineCorrenteSerial) {
            fineCorrenteSerial = resolveNamedRange(workbook, 'c_this_end');
        }
        if (!inizioCorrenteSerial) {
            inizioCorrenteSerial = resolveNamedRange(workbook, 'c_this_start');
        }

        // Leggi date precedente
        let finePrecedenteSerial = resolveNamedRange(workbook, 'c_prev_end_import');
        let inizioPrecedenteSerial = resolveNamedRange(workbook, 'c_prev_start_import');
        let annoPrecedenteStr = resolveNamedRange(workbook, 'c_prev');

        // Fallback a Configurazione
        if (!finePrecedenteSerial) {
            finePrecedenteSerial = resolveNamedRange(workbook, 'c_prev_end');
        }
        if (!inizioPrecedenteSerial) {
            inizioPrecedenteSerial = resolveNamedRange(workbook, 'c_prev_start');
        }

        // Leggi altri metadati
        const codiceFiscale = resolveNamedRange(workbook, 'cf');
        const valuta = resolveNamedRange(workbook, 'unit');

        console.log('📥 Valori letti da Named Ranges:', {
            'c_this_end': fineCorrenteSerial,
            'c_this_start': inizioCorrenteSerial,
            'c_this': annoCorrenteStr,
            'c_prev_end': finePrecedenteSerial,
            'c_prev_start': inizioPrecedenteSerial,
            'c_prev': annoPrecedenteStr,
            'cf': codiceFiscale,
            'unit': valuta
        });

        // Converti date seriali Excel in ISO
        if (fineCorrenteSerial && typeof fineCorrenteSerial === 'number') {
            bilancio.metadata.fine_corrente = excelSerialToISO(fineCorrenteSerial);
        }
        if (inizioCorrenteSerial && typeof inizioCorrenteSerial === 'number') {
            bilancio.metadata.inizio_corrente = excelSerialToISO(inizioCorrenteSerial);
        }
        if (finePrecedenteSerial && typeof finePrecedenteSerial === 'number') {
            bilancio.metadata.fine_precedente = excelSerialToISO(finePrecedenteSerial);
        }
        if (inizioPrecedenteSerial && typeof inizioPrecedenteSerial === 'number') {
            bilancio.metadata.inizio_precedente = excelSerialToISO(inizioPrecedenteSerial);
        }

        // Estrai anni da stringhe "c2022" → 2022
        let annoCorrente = null;
        let annoPrecedente = null;

        if (annoCorrenteStr && typeof annoCorrenteStr === 'string') {
            const match = annoCorrenteStr.match(/(\d{4})/);
            if (match) {
                annoCorrente = parseInt(match[1]);
                console.log('📅 Anno corrente estratto da "' + annoCorrenteStr + '":', annoCorrente);
            }
        }

        if (annoPrecedenteStr && typeof annoPrecedenteStr === 'string') {
            const match = annoPrecedenteStr.match(/(\d{4})/);
            if (match) {
                annoPrecedente = parseInt(match[1]);
                console.log('📅 Anno precedente estratto da "' + annoPrecedenteStr + '":', annoPrecedente);
            }
        }

        // Fallback: calcola anni dalle date se non specificati
        if (!annoCorrente && bilancio.metadata.fine_corrente) {
            annoCorrente = new Date(bilancio.metadata.fine_corrente).getFullYear();
            console.log('📅 Anno corrente calcolato da data fine:', annoCorrente);
        }
        if (!annoPrecedente && bilancio.metadata.fine_precedente) {
            annoPrecedente = new Date(bilancio.metadata.fine_precedente).getFullYear();
            console.log('📅 Anno precedente calcolato da data fine:', annoPrecedente);
        }

        // Imposta metadati
        if (annoCorrente) {
            bilancio.metadata.anno_esercizio = Math.round(annoCorrente);
        }
        if (annoPrecedente) {
            bilancio.metadata.anno_precedente = Math.round(annoPrecedente);
        }
        if (valuta) {
            bilancio.metadata.valuta = valuta;
        }
        if (codiceFiscale) {
            bilancio.metadata.codice_fiscale = codiceFiscale;
        }

        console.log('✓ Configurazione importata:', {
            anno_corrente: bilancio.metadata.anno_esercizio,
            anno_precedente: bilancio.metadata.anno_precedente,
            fine_corrente: bilancio.metadata.fine_corrente,
            fine_precedente: bilancio.metadata.fine_precedente,
            cf: bilancio.metadata.codice_fiscale,
            valuta: bilancio.metadata.valuta
        });

    } catch (error) {
        console.error('❌ Errore import configurazione:', error);
    }
}

// Converti seriale Excel in data ISO (YYYY-MM-DD)
function excelSerialToISO(serial) {
    if (!serial || isNaN(serial)) return null;
    
    // Excel usa 1900-01-01 come giorno 1 (con bug leap year 1900)
    const excelEpoch = new Date(1899, 11, 30); // 30 dic 1899
    const date = new Date(excelEpoch.getTime() + serial * 86400000);
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
}
