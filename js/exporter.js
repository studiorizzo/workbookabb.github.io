// workbookabb - Exporter Module

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
        
        // Per ogni foglio
        for (const codice in bilancio.fogli) {
            const template = getTemplate(codice);
            if (!template || !template.loaded) {
                console.warn(`Template ${codice} non disponibile, skip`);
                continue;
            }
            
            const dati = bilancio.fogli[codice];
            const templateData = template.data;
            
            if (!templateData || templateData.length < 2) continue;
            
            const config = templateData[1];
            const tipoTab = config[0];
            
            // Crea worksheet vuoto
            const worksheet = {};
            
            if (tipoTab === 1) {
                // TextBlock
                const codiceCell = templateData[2]?.[3];
                const firstRow = config[3];
                const firstCol = config[4];
                
                if (codiceCell && dati[codiceCell]) {
                    const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c: firstCol });
                    worksheet[cellAddress] = { 
                        t: 's', 
                        v: dati[codiceCell] 
                    };
                }
                
            } else if (tipoTab === 2) {
                // Tabella 2D
                const firstRow = config[3];
                const firstCol = config[4];
                const numRows = config[1];
                const codiciColonne = templateData[2]?.slice(3) || [];
                
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
                        }
                    }
                }
            }
            
            // Imposta range
            worksheet['!ref'] = 'A1:Z100'; // Range default
            
            // Aggiungi al workbook
            XLSX.utils.book_append_sheet(workbook, worksheet, codice);
        }
        
        // Download
        const filename = `bilancio_abbreviato_${new Date().toISOString().split('T')[0]}.xls`;
        XLSX.writeFile(workbook, filename);
        
        hideLoading();
        showToast('File Excel esportato', 'success');
        
    } catch (error) {
        console.error('Export XLS error:', error);
        hideLoading();
        showToast('Errore export XLS: ' + error.message, 'error');
    }
}

// Export XBRL (fase 2 - placeholder)
function exportXBRL() {
    showToast('Export XBRL non ancora implementato (Fase 2)', 'info');
    
    // TODO: Implementare in fase 2
    // 1. Per ogni cella valorizzata
    // 2. Cerca mapping in xbrl_mappings_complete.json
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
            
            const mapping = foglioMappings.find(m => 
                m.codice_excel === codiceExcel || 
                m.codice_excel?.startsWith(codiceExcel.split('_')[0])
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
