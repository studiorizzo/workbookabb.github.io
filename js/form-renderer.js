// workbookabb - Form Renderer Module (VERSIONE MIGLIORATA)

// Parse sheet config (righe 0-1) in HashMap
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

// Parse formula value - estrae valori da formule RAW (es. "=c2022" → 2022)
function parseFormulaValue(value, bilancio) {
    // Se non è una stringa o non inizia con =, restituisci come è
    if (typeof value !== 'string' || !value.startsWith('=')) {
        return value;
    }

    // Estrai anno da formule come "=c2022", "=c_2022", ecc.
    const yearMatch = value.match(/(\d{4})/);
    if (yearMatch) {
        return parseInt(yearMatch[1]);
    }

    // Se la formula è un Named Range, prova a risolverlo dal bilancio
    const namedRangeMatch = value.match(/^=([a-zA-Z_][a-zA-Z0-9_]*)/);
    if (namedRangeMatch && bilancio) {
        const rangeName = namedRangeMatch[1];

        // Mappa Named Ranges comuni a valori dal bilancio.metadata
        if (rangeName === 'c_this' || rangeName === 'anno_corrente') {
            return bilancio.metadata?.anno_esercizio || new Date().getFullYear();
        }
        if (rangeName === 'c_prev' || rangeName === 'anno_precedente') {
            return bilancio.metadata?.anno_precedente || (new Date().getFullYear() - 1);
        }
        if (rangeName === 'cf') {
            return bilancio.metadata?.codice_fiscale || '';
        }
        if (rangeName === 'unit') {
            return bilancio.metadata?.valuta || 'EUR';
        }
    }

    // Se non riesci a parsare, restituisci la formula originale senza il =
    return value.substring(1);
}

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
                        📅 Esercizio Corrente
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
                        📅 Esercizio Precedente
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
                    
                    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
                        <button onclick="salvaConfigurazione()" 
                                style="width: 100%; padding: 12px; background: #4CAF50; color: white; border: none; border-radius: 4px; font-size: 16px; font-weight: bold; cursor: pointer;">
                            💾 Salva Configurazione
                        </button>
                        <p style="margin-top: 15px; font-size: 13px; color: #666; text-align: center;">
                            <strong>Nota:</strong> Codice Fiscale si inserisce in T0000 (Dati Anagrafici)<br>
                            Valuta: EUR (Euro) - Standard per bilanci italiani
                        </p>
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

    const tipoTab = parseInt(config[0]);
    
    let html = '';

    // Header foglio
    const titolo = getTitoloFoglio(templateData, codice);
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

// Estrai titolo foglio - cerca in più posizioni e usa fallback
function getTitoloFoglio(templateData, foglioCode = null) {
    // Parse config per trovare first_row
    const configMap = parseSheetConfig(templateData);
    const firstRow = configMap ? parseInt(configMap.first_row) : 9;

    // POSIZIONE PRINCIPALE: Prima riga dati, colonna 2 (es. sheet[9][2] per T0006)
    // Il titolo si trova nella prima riga dei dati (first_row), colonna 2
    if (templateData[firstRow] && templateData[firstRow][2]) {
        const value = templateData[firstRow][2];
        if (typeof value === 'string' && value.length > 3 && value.length < 100) {
            if (!value.startsWith('=') && !value.match(/^[A-Z0-9_\.]{10,}$/)) {
                return value;
            }
        }
    }

    // FALLBACK: Cerca titolo in altre posizioni specifiche [riga, colonna]
    const possiblePositions = [
        [6, 1],  // Titolo esteso (es. "Conto economico abbreviato")
        [5, 2],  // Posizione JSON compresso (compatibilità)
        [3, 1],  // Altra posizione
        [firstRow, 1],  // Prima riga dati, colonna 1
        [7, 2]   // Altra posizione
    ];

    for (const [rowIdx, colIdx] of possiblePositions) {
        if (templateData[rowIdx] && templateData[rowIdx][colIdx]) {
            const value = templateData[rowIdx][colIdx];
            if (value && typeof value === 'string' && value.length > 3 && value.length < 100) {
                // Escludi valori che sembrano codici XBRL o formule
                if (!value.startsWith('=') && !value.match(/^[A-Z0-9_\.]{10,}$/)) {
                    return value;
                }
            }
        }
    }

    // Fallback: cerca nel nome del foglio usando i mapping XBRL
    if (foglioCode) {
        const xbrlMappings = getXBRLMappings();
        if (xbrlMappings && xbrlMappings.fogli && xbrlMappings.fogli[foglioCode]) {
            const foglioInfo = xbrlMappings.fogli[foglioCode];
            if (foglioInfo.nome || foglioInfo.label) {
                return foglioInfo.nome || foglioInfo.label;
            }
        }

        // Nomi conosciuti hardcoded come fallback
        const knownNames = {
            'T0000': 'Dati anagrafici',
            'T0002': 'Stato patrimoniale - Attivo',
            'T0004': 'Stato patrimoniale - Passivo',
            'T0006': 'Conto economico',
            'T0008': 'Rendiconto finanziario',
            'T0010': 'Movimentazione patrimonio netto'
        };

        if (knownNames[foglioCode]) {
            return knownNames[foglioCode];
        }
    }

    return 'Sezione bilancio';
}

// TIPO 1: Foglio semplice con 1-2 colonne di valori
function renderTipo1(templateData, dati, foglioCode) {
    const configMap = parseSheetConfig(templateData);
    if (!configMap) {
        return '<div class="empty-state"><p>Configurazione template non valida</p></div>';
    }

    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;

    const xbrlMappings = getXBRLMappings();

    // Se c'è una sola riga e una sola cella, è un textBlock
    if (numRows === 1 && numCols === 1) {
        // Il codice cella è in riga 2, colonna 3 (non in firstRow)
        const codiceCell = templateData[2]?.[3];
        if (!codiceCell) {
            return '<div class="empty-state"><p>Codice cella non trovato</p></div>';
        }

        const valore = dati[codiceCell] || '';
        const titolo = getTitoloFoglio(templateData, foglioCode);

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

    // Leggi anni da bilancio per header dinamici
    const bilancio = getBilancio();
    const annoCorrente = bilancio?.metadata?.anno_esercizio || new Date().getFullYear();
    const annoPrecedente = bilancio?.metadata?.anno_precedente || (annoCorrente - 1);

    // Determina se è T0000 o foglio senza codici colonna (stesso check di Tipo 2)
    const isT0000 = foglioCode === 'T0000';
    const effectiveNumCols = isT0000 ? 1 : numCols;

    // Altrimenti è una tabella semplice (es. T0000)
    let html = '<div class="form-container"><table class="bilancio-table">';

    // Header colonne - usa colCodeRow per gli header o firstRow-1 come fallback
    const headerRowIndex = colCodeRow > 0 ? colCodeRow : Math.max(8, firstRow - 1);
    const headerRow = templateData[headerRowIndex] || [];
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';

    if (isT0000) {
        // T0000: una sola colonna con anno corrente
        html += `<th>${annoCorrente}</th>`;
    } else {
        // Altri fogli tipo 1: usa header standard
        for (let c = 0; c < numCols; c++) {
            let colHeader = headerRow[firstCol + c] || '';

            // Parse formule RAW (es. =c2022 → 2022)
            colHeader = parseFormulaValue(colHeader, bilancio);

            // Sostituisci anni hardcoded con valori dinamici
            if (typeof colHeader === 'number' && colHeader >= 1900 && colHeader <= 2100) {
                if (c === 0) colHeader = annoCorrente;
                else if (c === 1) colHeader = annoPrecedente;
            }

            html += `<th>${escapeHtml(colHeader)}</th>`;
        }
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[rowCodeCol];
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
            for (let c = 0; c < effectiveNumCols; c++) {
                html += '<td class="cell-abstract">—</td>';
            }
        } else {
            // Per T0000: UNA sola cella, per altri: tutte le colonne
            for (let c = 0; c < effectiveNumCols; c++) {
                const valore = dati[codiceRiga] !== null && dati[codiceRiga] !== undefined
                    ? dati[codiceRiga]
                    : '';

                // Determina il tipo di input dalla mappatura
                const mapping = findMappingByCode(codiceRiga, foglioCode, xbrlMappings);
                const xbrlType = mapping?.xbrl?.type || 'string';
                const inputType = (xbrlType === 'monetary' || xbrlType === 'decimal') ? 'number' : 'text';

                html += `<td>
                    <input type="${inputType}"
                           class="cell-input"
                           data-cell="${codiceRiga}"
                           data-foglio="${foglioCode}"
                           value="${escapeHtml(valore)}"
                           ${inputType === 'number' ? 'step="0.01" placeholder="0"' : 'placeholder=""'} />
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
    const configMap = parseSheetConfig(templateData);
    if (!configMap) return '<div class="empty-state"><p>Configurazione non valida</p></div>';
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;
    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;

    // Header row - usa colCodeRow o firstRow-1 come fallback (NO hardcoded [9])
    const headerRowIndex = colCodeRow > 0 ? colCodeRow : Math.max(8, firstRow - 1);
    const headerRow = templateData[headerRowIndex] || [];
    const codiciColonne = templateData[colCodeRow]?.slice(firstCol) || [];
    
    const xbrlMappings = getXBRLMappings();
    
    // Leggi anni da bilancio per header dinamici
    const bilancio = getBilancio();
    const annoCorrente = bilancio?.metadata?.anno_esercizio || new Date().getFullYear();
    const annoPrecedente = bilancio?.metadata?.anno_precedente || (annoCorrente - 1);
    
    // Determina se è T0000 (dati anagrafici) o foglio senza codici colonna
    const isT0000 = foglioCode === 'T0000';
    const hasColumnCodes = codiciColonne.length > 0 && codiciColonne.some(c => c && c !== '-');
    
    // Per T0000 o fogli senza codici: usa una sola colonna (corrente)
    const effectiveNumCols = (isT0000 || !hasColumnCodes) ? 1 : codiciColonne.length;
    
    let html = '<div class="form-container"><table class="bilancio-table">';
    
    // Header colonne
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    
    if (isT0000 || !hasColumnCodes) {
        // T0000 o fogli senza codici colonna: una sola colonna (anno corrente)
        html += `<th>${annoCorrente}</th>`;
    } else {
        // Altri fogli: usa codici colonna
        for (let i = firstCol; i < headerRow.length && (i - firstCol) < codiciColonne.length; i++) {
            let headerValue = headerRow[i];

            // Parse formule RAW (es. =c2022 → 2022)
            headerValue = parseFormulaValue(headerValue, bilancio);

            let headerText = escapeHtml(headerValue || '');

            // Sostituisci anni hardcoded con valori dinamici da metadata
            if (typeof headerValue === 'number' && headerValue >= 1900 && headerValue <= 2100) {
                const colIndex = i - firstCol;
                if (colIndex === 0) {
                    headerText = annoCorrente;
                } else if (colIndex === 1) {
                    headerText = annoPrecedente;
                }
            } else if (typeof headerValue === 'string') {
                const parsed = parseInt(headerValue);
                if (!isNaN(parsed) && parsed >= 1900 && parsed <= 2100) {
                    const colIndex = i - firstCol;
                    if (colIndex === 0) {
                        headerText = annoCorrente;
                    } else if (colIndex === 1) {
                        headerText = annoPrecedente;
                    }
                }
            }

            html += `<th>${headerText}</th>`;
        }
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[rowCodeCol];
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
            for (let c = 0; c < effectiveNumCols; c++) {
                html += '<td class="cell-abstract">—</td>';
            }
        } else if (isT0000 || !hasColumnCodes) {
            // T0000 o foglio senza codici colonna: UNA sola cella con codiceRiga semplice
            const valore = dati[codiceRiga] !== null && dati[codiceRiga] !== undefined 
                ? dati[codiceRiga] 
                : '';
            
            // Determina il tipo di input dalla mappatura
            const mapping = findMappingByCode(codiceRiga, foglioCode, xbrlMappings);
            const xbrlType = mapping?.xbrl?.type || 'string';
            const inputType = (xbrlType === 'monetary' || xbrlType === 'decimal') ? 'number' : 'text';
            
            html += `<td>
                <input type="${inputType}" 
                       class="cell-input"
                       data-cell="${codiceRiga}"
                       data-foglio="${foglioCode}"
                       value="${escapeHtml(valore)}"
                       ${inputType === 'number' ? 'step="0.01" placeholder="0"' : 'placeholder=""'} />
            </td>`;
        } else {
            // Input normali con codici colonna
            for (let c = 0; c < codiciColonne.length; c++) {
                const codiceColonna = codiciColonne[c];
                if (!codiceColonna || codiceColonna === '-') {
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
    const configMap = parseSheetConfig(templateData);
    if (!configMap) return '<div class="empty-state"><p>Configurazione non valida</p></div>';
    const rowCodeCol = parseInt(configMap.row_code_nrcol) || 0;
    const colCodeRow = parseInt(configMap.col_code_nrrow) || 2;
    const firstRow = parseInt(configMap.first_row) || 0;
    const firstCol = parseInt(configMap.first_col) || 0;
    const numRows = parseInt(configMap.nr_row) || 0;
    const numCols = parseInt(configMap.nr_col) || 0;
    const codiciColonne = templateData[colCodeRow]?.slice(firstCol) || [];

    const xbrlMappings = getXBRLMappings();

    // Leggi anni da bilancio per header dinamici
    const bilancio = getBilancio();
    const annoCorrente = bilancio?.metadata?.anno_esercizio || new Date().getFullYear();
    const annoPrecedente = bilancio?.metadata?.anno_precedente || (annoCorrente - 1);

    let html = '<div class="form-container"><table class="bilancio-table">';

    // Header row - usa colCodeRow o firstRow-1 come fallback (NO hardcoded [9])
    const headerRowIndex = colCodeRow > 0 ? colCodeRow : Math.max(8, firstRow - 1);
    const headerRow = templateData[headerRowIndex] || [];
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    for (let i = firstCol; i < headerRow.length && (i - firstCol) < numCols; i++) {
        let headerValue = headerRow[i];

        // Parse formule RAW (es. =c2022 → 2022)
        headerValue = parseFormulaValue(headerValue, bilancio);

        let headerText = escapeHtml(headerValue || '');

        // Sostituisci anni hardcoded con valori dinamici
        const colIndex = i - firstCol;
        if (typeof headerValue === 'number' && headerValue >= 1900 && headerValue <= 2100) {
            if (colIndex === 0) headerText = annoCorrente;
            else if (colIndex === 1) headerText = annoPrecedente;
        }

        html += `<th>${headerText}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Righe dati - tipo 3 è simile a tipo 1 ma con possibilità di più valori per riga
    for (let r = 0; r < numRows; r++) {
        const rowIndex = firstRow + r;
        const rowData = templateData[rowIndex];
        
        if (!rowData) continue;
        
        const codiceRiga = rowData[rowCodeCol];
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
