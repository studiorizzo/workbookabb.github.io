// workbookabb - Form Renderer Module (VERSIONE MIGLIORATA)

// Renderizza Configurazione (date, codice fiscale, valuta)
function renderConfigurazione(content) {
    const bilancio = getBilancio();
    const workbookData = getWorkbookData();
    
    if (!bilancio) {
        content.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <h3>Nessun bilancio caricato</h3>
                <p>Crea un nuovo bilancio o carica un file Excel esistente.</p>
            </div>
        `;
        return;
    }
    
    // Leggi dati da bilancio.metadata o da workbookData.config
    const metadata = bilancio.metadata || {};
    const config = workbookData?.config || [];
    
    // Date esercizio corrente (da config[8])
    const inizioCorrente = metadata.inizio_corrente || (config[8]?.[2] || '2024-01-01');
    const fineCorrente = metadata.fine_corrente || (config[8]?.[1] || '2024-12-31');
    const annoCorrente = metadata.anno_esercizio || (config[8]?.[7] || new Date().getFullYear());
    
    // Date esercizio precedente (da config[9])
    const inizioPrecedente = metadata.inizio_precedente || (config[9]?.[2] || '2023-01-01');
    const finePrecedente = metadata.fine_precedente || (config[9]?.[1] || '2023-12-31');
    const annoPrecedente = metadata.anno_precedente || (config[9]?.[7] || annoCorrente - 1);
    
    const codiceFiscale = metadata.codice_fiscale || (config[13]?.[1] || '');
    const valuta = metadata.valuta || (config[11]?.[1] || 'EUR');
    
    // Converti date ISO in formato input date
    const toDateInput = (isoDate) => {
        if (!isoDate) return '';
        return isoDate.split('T')[0];
    };
    
    const html = `
        <div class="sheet-header">
            <h2>⚙️ Configurazione</h2>
            <p>Imposta i parametri del bilancio</p>
        </div>
        
        <div class="form-container">
            <div style="max-width: 700px; margin: 0 auto;">
                <div style="background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    
                    <h3 style="margin-bottom: 20px; color: #1976d2; border-bottom: 2px solid #1976d2; padding-bottom: 10px;">
                        📅 Esercizio Corrente (${annoCorrente})
                    </h3>
                    
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px;">
                        <div>
                            <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                                Data Inizio
                            </label>
                            <input type="date" 
                                   id="config_inizio_corrente"
                                   value="${toDateInput(inizioCorrente)}"
                                   style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px;" />
                        </div>
                        <div>
                            <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                                Data Fine
                            </label>
                            <input type="date" 
                                   id="config_fine_corrente"
                                   value="${toDateInput(fineCorrente)}"
                                   style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px;" />
                        </div>
                    </div>
                    
                    <h3 style="margin-bottom: 20px; color: #1976d2; border-bottom: 2px solid #1976d2; padding-bottom: 10px;">
                        📅 Esercizio Precedente (${annoPrecedente})
                    </h3>
                    
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px;">
                        <div>
                            <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                                Data Inizio
                            </label>
                            <input type="date" 
                                   id="config_inizio_precedente"
                                   value="${toDateInput(inizioPrecedente)}"
                                   style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px;" />
                        </div>
                        <div>
                            <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                                Data Fine
                            </label>
                            <input type="date" 
                                   id="config_fine_precedente"
                                   value="${toDateInput(finePrecedente)}"
                                   style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px;" />
                        </div>
                    </div>
                    
                    <h3 style="margin-bottom: 20px; color: #1976d2; border-bottom: 2px solid #1976d2; padding-bottom: 10px;">
                        🏢 Dati Azienda
                    </h3>
                    
                    <div style="margin-bottom: 25px;">
                        <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                            Codice Fiscale
                        </label>
                        <input type="text" 
                               id="config_codice_fiscale"
                               value="${escapeHtml(codiceFiscale)}"
                               maxlength="16"
                               style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px; text-transform: uppercase;" />
                        <small style="color: #666; font-size: 13px;">11 o 16 caratteri</small>
                    </div>
                    
                    <div style="margin-bottom: 25px;">
                        <label style="display: block; font-weight: bold; margin-bottom: 8px; color: #333;">
                            Valuta
                        </label>
                        <select id="config_valuta" 
                                style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 16px;">
                            <option value="EUR" ${valuta === 'EUR' ? 'selected' : ''}>EUR - Euro</option>
                            <option value="USD" ${valuta === 'USD' ? 'selected' : ''}>USD - Dollaro USA</option>
                            <option value="GBP" ${valuta === 'GBP' ? 'selected' : ''}>GBP - Sterlina</option>
                        </select>
                    </div>
                    
                    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
                        <button onclick="salvaConfigurazione()" 
                                style="width: 100%; padding: 12px; background: #4CAF50; color: white; border: none; border-radius: 4px; font-size: 16px; font-weight: bold; cursor: pointer;">
                            💾 Salva Configurazione
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    content.innerHTML = html;
}

// Renderizza foglio corrente
function renderFoglio(codice) {
    const content = document.getElementById('content');
    
    // GESTIONE SPECIALE CONFIGURAZIONE
    if (codice === 'Configurazione') {
        renderConfigurazione(content);
        return;
    }
    
    const template = getTemplate(codice);
    const bilancio = getBilancio();
    
    if (!template || !template.loaded) {
        content.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📋</div>
                <h3>Template non disponibile</h3>
                <p>Il template per ${codice} non è stato caricato.</p>
            </div>
        `;
        return;
    }
    
    if (!bilancio) {
        content.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <h3>Nessun bilancio caricato</h3>
                <p>Crea un nuovo bilancio o carica un file Excel esistente.</p>
            </div>
        `;
        return;
    }
    
    const templateData = template.data;
    const dati = bilancio.fogli[codice] || {};
    
    // Estrai configurazione
    if (!templateData || templateData.length < 2) {
        content.innerHTML = `<div class="empty-state"><p>Template formato non valido</p></div>`;
        return;
    }
    
    const config = templateData[1];
    
    // Verifica che config sia un array valido
    if (!Array.isArray(config) || config.length === 0) {
        content.innerHTML = `<div class="empty-state"><p>Configurazione template non valida</p></div>`;
        return;
    }
    
    const tipoTab = config[0];
    
    let html = '';
    
    // Header foglio
    const titolo = getTitoloFoglio(templateData);
    html += `
        <div class="sheet-header">
            <h2>${codice}</h2>
            <p>${titolo}</p>
        </div>
    `;
    
    // Renderizza in base al tipo
    if (tipoTab === 1) {
        html += renderTipo1(templateData, dati, codice);
    } else if (tipoTab === 2) {
        html += renderTipo2(templateData, dati, codice);
    } else if (tipoTab === 3) {
        html += renderTipo3(templateData, dati, codice);
    } else if (tipoTab === 4) {
        html += renderTipo4(templateData, dati, codice);
    } else {
        // Tipo non riconosciuto - mostra informazioni di debug
        html += `
            <div class="info-box">
                <h3>⚠️ Tipo foglio non standard: ${tipoTab}</h3>
                <p>Questo foglio utilizza un formato non ancora implementato.</p>
                <details>
                    <summary>Configurazione foglio (debug)</summary>
                    <pre>${JSON.stringify(config, null, 2)}</pre>
                </details>
            </div>
        `;
    }
    
    content.innerHTML = html;
    
    // Attach event listeners agli input
    attachInputListeners();
}

// Estrai titolo foglio
function getTitoloFoglio(templateData) {
    // Cerca in riga 6 o 4
    if (templateData[6] && templateData[6][1]) {
        return templateData[6][1];
    }
    if (templateData[4] && templateData[4][1]) {
        return templateData[4][1];
    }
    return 'Sezione bilancio';
}

// TIPO 1: Foglio semplice con 1-2 colonne di valori
function renderTipo1(templateData, dati, foglioCode) {
    const config = templateData[1];
    const firstRow = config[3];
    const firstCol = config[4];
    const numRows = config[1];
    const numCols = config[2];
    
    const xbrlMappings = getXBRLMappings();
    
    // Se c'è una sola riga e una sola cella, è un textBlock
    if (numRows === 1 && numCols === 1) {
        // Il codice cella è in riga 2, colonna 3 (non in firstRow)
        const codiceCell = templateData[2]?.[3];
        if (!codiceCell) {
            return '<div class="empty-state"><p>Codice cella non trovato</p></div>';
        }
        
        const valore = dati[codiceCell] || '';
        const titolo = getTitoloFoglio(templateData);
        
        const mapping = findMappingByCode(codiceCell, foglioCode, xbrlMappings);
        const label = mapping?.ui?.label || titolo;
        
        return `
            <div class="textblock">
                <h3>${escapeHtml(label)}</h3>
                <textarea 
                    class="cell-textarea"
                    data-cell="${codiceCell}"
                    data-foglio="${foglioCode}"
                    placeholder="Inserisci il testo...">` + escapeHtml(valore) + `</textarea>
            </div>
        `;
    }
    
    // Altrimenti è una tabella semplice (es. T0000)
    let html = '<div class="form-container"><table class="bilancio-table">';
    
    // Header colonne (cerca nella riga 8 o firstRow-1)
    const headerRow = templateData[Math.max(8, firstRow - 1)] || [];
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    for (let c = 0; c < numCols; c++) {
        const colHeader = headerRow[firstCol + c] || '';
        html += `<th>${escapeHtml(colHeader)}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        // Cerca mapping per label e indentazione
        const mapping = findMappingByCode(codiceRiga, foglioCode, xbrlMappings);
        const labelRiga = mapping?.ui?.label || rowData[2] || '';
        const indentLevel = mapping?.ui?.indent_level || 0;
        const isAbstract = mapping?.ui?.is_abstract || false;
        
        const indentPx = indentLevel * 20;
        const labelClass = isAbstract ? 'label-abstract' : 'label';
        
        html += `<tr>
            <td class="${labelClass}" style="padding-left: ${indentPx}px">
                ${escapeHtml(labelRiga)}
            </td>`;
        
        // Celle input
        if (isAbstract) {
            for (let c = 0; c < numCols; c++) {
                html += '<td class="cell-abstract">—</td>';
            }
        } else {
            for (let c = 0; c < numCols; c++) {
                const valore = dati[codiceRiga] !== null && dati[codiceRiga] !== undefined 
                    ? dati[codiceRiga] 
                    : '';
                
                html += `<td>
                    <input type="text" 
                           class="cell-input"
                           data-cell="${codiceRiga}"
                           data-foglio="${foglioCode}"
                           value="${escapeHtml(valore)}"
                           placeholder="" />
                </td>`;
            }
        }
        
        html += '</tr>';
    }
    
    html += '</tbody></table></div>';
    return html;
}

// TIPO 2: Tabella 2D con codici riga × codici colonna
function renderTipo2(templateData, dati, foglioCode) {
    const config = templateData[1];
    const firstRow = config[3];
    const numRows = config[1];
    const headerRow = templateData[9] || templateData[Math.max(8, firstRow - 1)] || [];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    const xbrlMappings = getXBRLMappings();
    
    // Leggi anni da bilancio per header dinamici
    const bilancio = getBilancio();
    const annoCorrente = bilancio?.metadata?.anno_esercizio || new Date().getFullYear();
    const annoPrecedente = bilancio?.metadata?.anno_precedente || (annoCorrente - 1);
    
    let html = '<div class="form-container"><table class="bilancio-table">';
    
    // Header colonne
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    for (let i = 3; i < headerRow.length && (i-3) < codiciColonne.length; i++) {
        let headerText = escapeHtml(headerRow[i] || '');
        // Sostituisci anni hardcoded con valori reali
        if (i === 3) headerText = annoCorrente;
        if (i === 4) headerText = annoPrecedente;
        html += `<th>${headerText}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        // Cerca mapping per label e indentazione
        const mapping = findMappingByCode(codiceRiga, foglioCode, xbrlMappings);
        const labelRiga = mapping?.ui?.label || rowData[2] || '';
        const indentLevel = mapping?.ui?.indent_level || 0;
        const isAbstract = mapping?.ui?.is_abstract || false;
        
        const indentPx = indentLevel * 20;
        const labelClass = isAbstract ? 'label-abstract' : 'label';
        
        html += `<tr>
            <td class="${labelClass}" style="padding-left: ${indentPx}px">
                ${escapeHtml(labelRiga)}
            </td>`;
        
        // Celle input
        if (isAbstract) {
            // Abstract: nessun input
            for (let c = 0; c < codiciColonne.length; c++) {
                html += '<td class="cell-abstract">—</td>';
            }
        } else {
            // Input normali
            for (let c = 0; c < codiciColonne.length; c++) {
                const codiceColonna = codiciColonne[c];
                if (!codiceColonna) {
                    html += '<td></td>';
                    continue;
                }
                
                const codiceCella = `${codiceRiga}_${codiceColonna}`;
                const valore = dati[codiceCella] !== null && dati[codiceCella] !== undefined 
                    ? dati[codiceCella] 
                    : '';
                
                html += `<td>
                    <input type="number" 
                           class="cell-input"
                           data-cell="${codiceCella}"
                           data-foglio="${foglioCode}"
                           value="${valore}"
                           step="0.01"
                           placeholder="0" />
                </td>`;
            }
        }
        
        html += '</tr>';
    }
    
    html += '</tbody></table></div>';
    return html;
}

// TIPO 3: Tabella con tuple (righe multiple espandibili)
function renderTipo3(templateData, dati, foglioCode) {
    const config = templateData[1];
    const firstRow = config[3];
    const numRows = config[1];
    const numCols = config[2];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    const xbrlMappings = getXBRLMappings();
    
    // Leggi anni da bilancio per header dinamici
    const bilancio = getBilancio();
    const annoCorrente = bilancio?.metadata?.anno_esercizio || new Date().getFullYear();
    const annoPrecedente = bilancio?.metadata?.anno_precedente || (annoCorrente - 1);
    
    let html = '<div class="form-container"><table class="bilancio-table">';
    
    // Header (cerca in riga 9 o firstRow-1)
    const headerRow = templateData[9] || templateData[Math.max(8, firstRow - 1)] || [];
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    for (let i = 3; i < headerRow.length && (i-3) < numCols; i++) {
        let headerText = escapeHtml(headerRow[i] || '');
        // Sostituisci anni hardcoded con valori reali
        if (i === 3) headerText = annoCorrente;
        if (i === 4) headerText = annoPrecedente;
        html += `<th>${headerText}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati - tipo 3 è simile a tipo 1 ma con possibilità di più valori per riga
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[0];
        if (!codiceRiga) continue;
        
        const mapping = findMappingByCode(codiceRiga, foglioCode, xbrlMappings);
        const labelRiga = mapping?.ui?.label || rowData[2] || '';
        const indentLevel = mapping?.ui?.indent_level || 0;
        const isAbstract = mapping?.ui?.is_abstract || false;
        
        const indentPx = indentLevel * 20;
        const labelClass = isAbstract ? 'label-abstract' : 'label';
        
        html += `<tr>
            <td class="${labelClass}" style="padding-left: ${indentPx}px">
                ${escapeHtml(labelRiga)}
            </td>`;
        
        if (isAbstract) {
            for (let c = 0; c < numCols; c++) {
                html += '<td class="cell-abstract">—</td>';
            }
        } else {
            // Per tipo 3: se c'è un solo valore, usa codiceRiga diretto
            // Se ci sono più colonne, usa codiceRiga_codiceColonna
            if (numCols === 1 || codiciColonne.length === 0) {
                const valore = dati[codiceRiga] !== null && dati[codiceRiga] !== undefined 
                    ? dati[codiceRiga] 
                    : '';
                
                html += `<td>
                    <input type="text" 
                           class="cell-input"
                           data-cell="${codiceRiga}"
                           data-foglio="${foglioCode}"
                           value="${escapeHtml(valore)}"
                           placeholder="" />
                </td>`;
            } else {
                // Più colonne
                for (let c = 0; c < numCols && c < codiciColonne.length; c++) {
                    const codiceColonna = codiciColonne[c];
                    const codiceCella = codiceColonna ? `${codiceRiga}_${codiceColonna}` : codiceRiga;
                    const valore = dati[codiceCella] !== null && dati[codiceCella] !== undefined 
                        ? dati[codiceCella] 
                        : '';
                    
                    html += `<td>
                        <input type="text" 
                               class="cell-input"
                               data-cell="${codiceCella}"
                               data-foglio="${foglioCode}"
                               value="${escapeHtml(valore)}"
                               placeholder="" />
                    </td>`;
                }
            }
        }
        
        html += '</tr>';
    }
    
    html += '</tbody></table></div>';
    return html;
}

// TIPO 4: Altri formati speciali
function renderTipo4(templateData, dati, foglioCode) {
    // TODO: Implementare altri formati
    return `
        <div class="info-box">
            <h3>⚠️ Tipo 4: Formato speciale (non ancora implementato)</h3>
            <p>Questo foglio utilizza un formato speciale.</p>
        </div>
    `;
}

// Attach event listeners agli input
function attachInputListeners() {
    // Input numerici e testuali
    document.querySelectorAll('.cell-input').forEach(input => {
        input.addEventListener('change', handleCellChange);
        input.addEventListener('blur', handleCellChange);
    });
    
    // Textarea
    document.querySelectorAll('.cell-textarea').forEach(textarea => {
        textarea.addEventListener('blur', handleCellChange);
    });
}

// Handle cambio valore cella
function handleCellChange(event) {
    const input = event.target;
    const codiceCella = input.dataset.cell;
    const foglioCode = input.dataset.foglio;
    
    let valore;
    if (input.type === 'number') {
        valore = input.value ? parseFloat(input.value) : null;
    } else {
        valore = input.value;
    }
    
    updateCellValue(foglioCode, codiceCella, valore);
}

// Trova mapping per codice
function findMappingByCode(codiceExcel, foglioCode, xbrlMappings) {
    if (!xbrlMappings || !xbrlMappings.mappature) return null;
    
    const foglioMappings = xbrlMappings.mappature[foglioCode];
    if (!foglioMappings) return null;
    
    // Cerca match esatto
    let mapping = foglioMappings.find(m => m.codice_excel === codiceExcel);
    
    // Se non trovato, cerca per prefisso (per celle combinate riga_colonna)
    if (!mapping) {
        const codiceBase = codiceExcel.split('_')[0];
        mapping = foglioMappings.find(m => m.codice_excel === codiceBase);
    }
    
    return mapping;
}

// Escape HTML
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}
