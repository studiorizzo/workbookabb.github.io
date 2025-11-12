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
            
            // Estrai configurazione
            if (!templateData || templateData.length < 2) {
                console.warn(`Template ${sheetName} formato invalido`);
                continue;
            }
            
            const config = templateData[1];
            
            // Verifica che config sia un array valido
            if (!Array.isArray(config) || config.length === 0) {
                console.warn(`Config ${sheetName} non valida, skip`);
                continue;
            }
            
            const tipoTab = config[0];
            
            bilancio.fogli[sheetName] = {};
            let celleSheet = 0;
            
            if (tipoTab === 1) {
                // TIPO 1: Foglio semplice
                celleSheet = importTipo1(sheet, templateData, config, bilancio.fogli[sheetName]);
            } else if (tipoTab === 2) {
                // TIPO 2: Tabella 2D
                celleSheet = importTipo2(sheet, templateData, config, bilancio.fogli[sheetName]);
            } else if (tipoTab === 3) {
                // TIPO 3: Tuple
                celleSheet = importTipo3(sheet, templateData, config, bilancio.fogli[sheetName]);
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
function importTipo1(sheet, templateData, config, datiSheet) {
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const numCols = config[2];
    
    let count = 0;
    
    // Se 1 riga × 1 colonna = textBlock
    if (numRows === 1 && numCols === 1) {
        // Il codice è in riga 2, colonna 3
        const codiceCell = templateData[2]?.[3];
        if (codiceCell) {
            const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c: firstCol });
            const cell = sheet[cellAddress];
            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                datiSheet[codiceCell] = cell.v;
                count++;
            }
        }
    } else {
        // Altrimenti tabella semplice con codici riga
        for (let r = 0; r < numRows; r++) {
            const rowData = templateData[firstRow + r];
            if (!rowData) continue;
            
            const codiceRiga = rowData[0];
            if (!codiceRiga) continue;
            
            for (let c = 0; c < numCols; c++) {
                const xlsRow = firstRow + r;
                const xlsCol = firstCol + c;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
                
                const cell = sheet[cellAddress];
                if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                    // Per tipo 1 con più colonne, usa codice_riga direttamente
                    // (se servisse distinguere le colonne, si dovrebbe estendere)
                    datiSheet[codiceRiga] = cell.v;
                    count++;
                }
            }
        }
    }
    
    return count;
}

// Import TIPO 2: Tabella 2D
function importTipo2(sheet, templateData, config, datiSheet) {
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const numCols = config[2];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    let count = 0;
    
    for (let r = 0; r < numRows; r++) {
        const rowData = templateData[firstRow + r];
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        for (let c = 0; c < numCols && c < codiciColonne.length; c++) {
            const codiceColonna = codiciColonne[c];
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
function importTipo3(sheet, templateData, config, datiSheet) {
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const numCols = config[2];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    let count = 0;
    
    for (let r = 0; r < numRows; r++) {
        const rowData = templateData[firstRow + r];
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        // Se una sola colonna o nessun codice colonna, usa codiceRiga diretto
        if (numCols === 1 || codiciColonne.length === 0) {
            const xlsRow = firstRow + r;
            const xlsCol = firstCol;
            const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
            const cell = sheet[cellAddress];
            
            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                datiSheet[codiceRiga] = cell.v;
                count++;
            }
        } else {
            // Più colonne
            for (let c = 0; c < numCols && c < codiciColonne.length; c++) {
                const codiceColonna = codiciColonne[c];
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
