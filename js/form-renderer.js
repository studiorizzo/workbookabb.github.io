// workbookabb - Form Renderer Module

// Renderizza foglio corrente
function renderFoglio(codice) {
    const content = document.getElementById('content');
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
        html += renderTextBlock(templateData, dati, codice);
    } else if (tipoTab === 2) {
        html += renderTabella2D(templateData, dati, codice);
    } else {
        html += `<div class="empty-state"><p>Tipo tabella non supportato: ${tipoTab}</p></div>`;
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

// Render TextBlock (campo testo libero)
function renderTextBlock(templateData, dati, foglioCode) {
    const codiceCell = templateData[2]?.[3];
    if (!codiceCell) {
        return '<div class="empty-state"><p>Codice cella non trovato</p></div>';
    }
    
    const valore = dati[codiceCell] || '';
    const titolo = getTitoloFoglio(templateData);
    
    // Cerca label e indent da xbrl_mappings
    const xbrlMappings = getXBRLMappings();
    const mapping = findMappingByCode(codiceCell, foglioCode, xbrlMappings);
    const label = mapping?.ui?.label || titolo;
    
    return `
        <div class="textblock">
            <h3>${label}</h3>
            <textarea 
                class="cell-textarea"
                data-cell="${codiceCell}"
                data-foglio="${foglioCode}"
                placeholder="Inserisci il testo...">` + escapeHtml(valore) + `</textarea>
        </div>
    `;
}

// Render Tabella 2D
function renderTabella2D(templateData, dati, foglioCode) {
    const config = templateData[1];
    const firstRow = config[3];
    const numRows = config[1];
    const headerRow = templateData[9] || [];
    const codiciColonne = templateData[2]?.slice(3) || [];
    
    const xbrlMappings = getXBRLMappings();
    
    let html = '<div class="form-container"><table class="bilancio-table">';
    
    // Header colonne
    html += '<thead><tr><th style="min-width: 250px;">Descrizione</th>';
    for (let i = 3; i < headerRow.length && (i-3) < codiciColonne.length; i++) {
        html += `<th>${escapeHtml(headerRow[i] || '')}</th>`;
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

// Attach event listeners agli input
function attachInputListeners() {
    // Input numerici
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
    if (input.classList.contains('cell-input')) {
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
    
    return foglioMappings.find(m => 
        m.codice_excel === codiceExcel ||
        m.codice_excel?.startsWith(codiceExcel)
    );
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
