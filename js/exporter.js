// workbookabb - Exporter Module (VERSIONE MIGLIORATA)

// Export XLS
async function exportXLS() {
    const bilancio = getBilancio();
    if (!bilancio) {
        showToast('Nessun bilancio da esportare', 'error');
        return;
    }
    
    try {
        showLoading('Generazione file Excel...');
        
        const workbook = XLSX.utils.book_new();
        let fogliEsportati = 0;
        
        // Per ogni foglio nel bilancio
        for (const codice in bilancio.fogli) {
            const template = getTemplate(codice);
            if (!template || !template.loaded) {
                console.warn(`Template ${codice} non disponibile, skip export`);
                continue;
            }
            
            const dati = bilancio.fogli[codice];
            const templateData = template.data;
            
            if (!templateData || templateData.length < 2) continue;
            
            const config = templateData[1];
            if (!Array.isArray(config) || config.length === 0) continue;
            
            const tipoTab = config[0];
            
            // Crea worksheet vuoto
            const worksheet = {};
            let celleEsportate = 0;
            
            if (tipoTab === 1) {
                celleEsportate = exportTipo1(worksheet, dati, templateData, config);
            } else if (tipoTab === 2) {
                celleEsportate = exportTipo2(worksheet, dati, templateData, config);
            } else if (tipoTab === 3) {
                celleEsportate = exportTipo3(worksheet, dati, templateData, config);
            } else {
                console.warn(`Tipo ${tipoTab} non supportato per export ${codice}`);
                continue;
            }
            
            if (celleEsportate > 0) {
                // Imposta range (espansione automatica)
                worksheet['!ref'] = 'A1:Z100';
                
                // Aggiungi al workbook
                XLSX.utils.book_append_sheet(workbook, worksheet, codice);
                fogliEsportati++;
                
                console.log(`✓ Exported ${codice}: ${celleEsportate} celle (tipo ${tipoTab})`);
            }
        }
        
        if (fogliEsportati === 0) {
            hideLoading();
            showToast('Nessun foglio da esportare', 'warning');
            return;
        }
        
        // Download
        const filename = `bilancio_abbreviato_${new Date().toISOString().split('T')[0]}.xls`;
        XLSX.writeFile(workbook, filename);
        
        hideLoading();
        showToast(`File Excel esportato (${fogliEsportati} fogli)`, 'success');
        
    } catch (error) {
        console.error('Export XLS error:', error);
        hideLoading();
        showToast('Errore export XLS: ' + error.message, 'error');
    }
}

// Export TIPO 1: Foglio semplice
function exportTipo1(worksheet, dati, templateData, config) {
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const numCols = config[2];
    
    let count = 0;
    
    // Se 1×1 = textBlock
    if (numRows === 1 && numCols === 1) {
        // Il codice è in riga 2, colonna 3
        const codiceCell = templateData[2]?.[3];
        if (codiceCell && dati[codiceCell]) {
            const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c: firstCol });
            worksheet[cellAddress] = { 
                t: 's', 
                v: dati[codiceCell] 
            };
            count++;
        }
    } else {
        // Tabella semplice
        for (let r = 0; r < numRows; r++) {
            const rowData = templateData[firstRow + r];
            if (!rowData) continue;
            
            const codiceRiga = rowData[0];
            if (!codiceRiga) continue;
            
            const valore = dati[codiceRiga];
            if (valore !== null && valore !== undefined && valore !== '') {
                const xlsRow = firstRow + r;
                const xlsCol = firstCol;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
                
                worksheet[cellAddress] = {
                    t: typeof valore === 'number' ? 'n' : 's',
                    v: valore
                };
                count++;
            }
        }
    }
    
    return count;
}

// Export TIPO 2: Tabella 2D
function exportTipo2(worksheet, dati, templateData, config) {
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    let count = 0;
    
    for (let r = 0; r < numRows; r++) {
        const rowData = templateData[firstRow + r];
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        for (let c = 0; c < codiciColonne.length; c++) {
            const codiceColonna = codiciColonne[c];
            if (!codiceColonna) continue;
            
            const codiceCella = `${codiceRiga}_${codiceColonna}`;
            const valore = dati[codiceCella];
            
            if (valore !== null && valore !== undefined && valore !== '') {
                const xlsRow = firstRow + r;
                const xlsCol = firstCol + c;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
                
                worksheet[cellAddress] = {
                    t: typeof valore === 'number' ? 'n' : 's',
                    v: valore
                };
                count++;
            }
        }
    }
    
    return count;
}

// Export TIPO 3: Tuple
function exportTipo3(worksheet, dati, templateData, config) {
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
        
        // Se una sola colonna, usa codiceRiga diretto
        if (numCols === 1 || codiciColonne.length === 0) {
            const valore = dati[codiceRiga];
            if (valore !== null && valore !== undefined && valore !== '') {
                const xlsRow = firstRow + r;
                const xlsCol = firstCol;
                const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
                
                worksheet[cellAddress] = {
                    t: typeof valore === 'number' ? 'n' : 's',
                    v: valore
                };
                count++;
            }
        } else {
            // Più colonne
            for (let c = 0; c < numCols && c < codiciColonne.length; c++) {
                const codiceColonna = codiciColonne[c];
                if (!codiceColonna) continue;
                
                const codiceCella = `${codiceRiga}_${codiceColonna}`;
                const valore = dati[codiceCella];
                
                if (valore !== null && valore !== undefined && valore !== '') {
                    const xlsRow = firstRow + r;
                    const xlsCol = firstCol + c;
                    const cellAddress = XLSX.utils.encode_cell({ r: xlsRow, c: xlsCol });
                    
                    worksheet[cellAddress] = {
                        t: typeof valore === 'number' ? 'n' : 's',
                        v: valore
                    };
                    count++;
                }
            }
        }
    }
    
    return count;
}

// Export XBRL (fase 2 - placeholder)
function exportXBRL() {
    showToast('Export XBRL non ancora implementato (Fase 2)', 'info');
    
    // TODO: Implementare in fase 2
    // 1. Per ogni cella valorizzata
    // 2. Cerca mapping in mappings.json
    // 3. Genera fact XBRL con elemento, contesto, unità
    // 4. Valida con XML tassonomia
    // 5. Download istanza XBRL
}

// Genera preview XBRL (debug)
function generateXBRLPreview() {
    const bilancio = getBilancio();
    if (!bilancio) return '';
    
    const xbrlMappings = getXBRLMappings();
    if (!xbrlMappings) return '';
    
    let xbrl = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xbrl += '<xbrl xmlns="http://www.xbrl.org/2003/instance">\n\n';
    
    // Contesti base
    xbrl += '  <context id="c2024">\n';
    xbrl += '    <entity><identifier scheme="...">...</identifier></entity>\n';
    xbrl += '    <period><instant>2024-12-31</instant></period>\n';
    xbrl += '  </context>\n\n';
    
    // Facts (primi 10 per preview)
    let count = 0;
    for (const foglio in bilancio.fogli) {
        if (count >= 10) break;
        
        const dati = bilancio.fogli[foglio];
        const foglioMappings = xbrlMappings.mappature?.[foglio];
        
        if (!foglioMappings) continue;
        
        for (const codiceExcel in dati) {
            if (count >= 10) break;
            
            const valore = dati[codiceExcel];
            if (valore === null || valore === undefined || valore === '') continue;
            
            const codiceBase = codiceExcel.split('_')[0];
            const mapping = foglioMappings.find(m => 
                m.codice_excel === codiceExcel || 
                m.codice_excel === codiceBase
            );
            
            if (mapping && mapping.xbrl) {
                const prefix = mapping.xbrl.prefix || 'itcc-ci';
                const name = mapping.xbrl.name;
                const type = mapping.xbrl.type;
                
                if (type === 'nonnum:textBlock') {
                    xbrl += `  <${prefix}:${name} contextRef="c2024">\n`;
                    xbrl += `    <![CDATA[${valore}]]>\n`;
                    xbrl += `  </${prefix}:${name}>\n`;
                } else if (type === 'monetary') {
                    xbrl += `  <${prefix}:${name} contextRef="c2024" unitRef="EUR" decimals="2">${valore}</${prefix}:${name}>\n`;
                } else {
                    xbrl += `  <${prefix}:${name} contextRef="c2024">${valore}</${prefix}:${name}>\n`;
                }
                
                count++;
            }
        }
    }
    
    xbrl += '\n</xbrl>';
    
    return xbrl;
}
