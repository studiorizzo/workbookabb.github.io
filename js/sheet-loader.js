// workbookabb - Sheet Loader Module (SheetJS)

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
    
    // Per ogni foglio nel workbook
    for (const sheetName of workbook.SheetNames) {
        try {
            // Carica template per questo foglio
            const template = await loadTemplate(sheetName);
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
            const tipoTab = config[0];
            
            bilancio.fogli[sheetName] = {};
            
            if (tipoTab === 1) {
                // TextBlock
                const codiceCell = templateData[2]?.[3];
                const firstRow = config[3];
                const firstCol = config[4];
                
                if (codiceCell) {
                    const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c: firstCol });
                    const cell = sheet[cellAddress];
                    bilancio.fogli[sheetName][codiceCell] = cell ? (cell.v || '') : '';
                }
                
            } else if (tipoTab === 2) {
                // Tabella 2D
                const firstRow = config[3];
                const firstCol = config[4];
                const numRows = config[1];
                const numCols = config[2];
                const codiciColonne = templateData[2]?.slice(3) || [];
                
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
                        bilancio.fogli[sheetName][codiceCella] = cell ? (cell.v || null) : null;
                    }
                }
            }
            
            console.log(`Imported ${sheetName}: ${Object.keys(bilancio.fogli[sheetName]).length} celle`);
            
        } catch (error) {
            console.error(`Error importing ${sheetName}:`, error);
        }
    }
    
    return bilancio;
}

// Estrai dati foglio XLS
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

// Get cell value
function getCellValue(sheet, row, col) {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = sheet[cellAddress];
    return cell ? (cell.v || null) : null;
}
