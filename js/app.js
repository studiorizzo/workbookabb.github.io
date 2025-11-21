// workbookabb - Main Application Module (VERSIONE UNIFICATA)

// Variabili globali
let workbookData = null; // Contiene config, index, sheets
let configurazione = null;
let templates = {}; // Cache templates già processati
let templatesList = []; // Lista codici fogli disponibili
let xbrlMappings = null;
let currentFoglio = null;
let bilancio = null;

// Inizializzazione
document.addEventListener('DOMContentLoaded', async () => {
    console.log('=== Initializing workbookabb ===');
    
    try {
        showLoading('Caricamento risorse...');
        
        // 1. Carica workbookapp.json unificato
        await loadWorkbookData();
        
        // 2. Estrai configurazione
        extractConfigurazione();
        
        // 3. Scopri tutti i fogli disponibili
        discoverSheets();
        
        // 4. Carica mappature XBRL (label, indentazioni)
        await loadXBRLMappings();
        
        // 5. Genera indice sidebar
        renderIndex();
        
        // 6. Setup event listeners
        setupEventListeners();
        
        // 7. Verifica bilancio salvato
        loadSavedBilancio();
        
        hideLoading();
        
        console.log('=== App initialized successfully ===');
    } catch (error) {
        console.error('=== Initialization error ===', error);
        hideLoading();
        showToast('Errore inizializzazione: ' + error.message, 'error');
    }
});

// Carica workbookabb.json unificato (supporta sia dense che sparse format)
async function loadWorkbookData() {
    try {
        // Try sparse format first (smaller, faster)
        let response = await fetch('data/template/workbookabb-sparse.json');
        if (!response.ok) {
            // Fallback to dense format
            console.log('Sparse format not found, loading dense format...');
            response = await fetch('data/template/workbookabb.json');
            if (!response.ok) throw new Error('workbookabb.json non trovato');
        }

        const rawData = await response.json();
        console.log('✓ Template file loaded');

        // Prepare workbook data (handles both sparse and dense formats)
        workbookData = window.SparseLoader.prepareWorkbookData(rawData);
        console.log(`✓ Using ${workbookData._format} format`);
    } catch (error) {
        throw new Error('Impossibile caricare template: ' + error.message);
    }
}

// Estrai configurazione da workbookData
function extractConfigurazione() {
    if (!workbookData || !workbookData.config) {
        console.warn('⚠ Configurazione non disponibile');
        configurazione = null;
        return;
    }
    
    configurazione = workbookData.config;
    console.log('✓ Configurazione estratta');
}

// Scopri tutti i fogli in workbookData.sheets
function discoverSheets() {
    if (!workbookData || !workbookData.sheets) {
        console.warn('⚠ Nessun foglio trovato');
        templatesList = [];
        return;
    }
    
    // Lista ordinata dei fogli
    templatesList = Object.keys(workbookData.sheets).sort();
    
    console.log(`✓ Discovered ${templatesList.length} sheets:`, templatesList.slice(0, 10), '...');
}

// Carica mappature XBRL
async function loadXBRLMappings() {
    try {
        const response = await fetch('data/mapping/mappings.json');
        if (!response.ok) throw new Error('mappings.json non trovato');
        xbrlMappings = await response.json();
        console.log('✓ XBRL mappings loaded:', Object.keys(xbrlMappings.mappature || {}).length, 'fogli');
    } catch (error) {
        console.warn('⚠ mappings.json not available:', error.message);
        xbrlMappings = { mappature: {} };
    }
}

// Carica template foglio specifico (da workbookData)
function loadTemplate(codice) {
    // Se già in cache, ritorna
    if (templates[codice] && templates[codice].loaded) {
        return templates[codice];
    }
    
    // Cerca in workbookData.sheets
    if (!workbookData || !workbookData.sheets || !workbookData.sheets[codice]) {
        console.error(`✗ Template ${codice} non trovato in workbookData`);
        return null;
    }
    
    // Crea oggetto template
    templates[codice] = {
        code: codice,
        data: workbookData.sheets[codice],
        loaded: true
    };
    
    console.log(`✓ Template ${codice} loaded from workbookData`);
    return templates[codice];
}

// Genera indice navigazione da templatesList
function renderIndex() {
    const indexTree = document.getElementById('index-tree');
    
    if (templatesList.length === 0) {
        indexTree.innerHTML = '<div class="empty-state"><p>Nessun template disponibile</p></div>';
        return;
    }
    
    // Raggruppa fogli per categoria (basato su naming pattern)
    const grouped = {
        'Principali': [],
        'Nota Integrativa': [],
        'Altri': []
    };
    
    templatesList.forEach(code => {
        const num = parseInt(code.substring(1));
        if (num <= 13) {
            grouped['Principali'].push(code);
        } else if (num >= 14 && num <= 615) {
            grouped['Nota Integrativa'].push(code);
        } else {
            grouped['Altri'].push(code);
        }
    });
    
    // Genera HTML
    let html = '<div class="tree-submenu">';
    
    // CONFIGURAZIONE in cima
    html += '<div class="tree-section"><strong>Configurazione</strong></div>';
    html += `<div class="tree-item" data-foglio="Configurazione" onclick="navigateToFoglio('Configurazione')">
        ⚙️ Configurazione Bilancio
    </div>`;
    
    // Principali (sempre espansi)
    if (grouped['Principali'].length > 0) {
        html += '<div class="tree-section" style="margin-top: 15px;"><strong>Fogli Principali</strong></div>';
        grouped['Principali'].forEach(code => {
            const label = getFoglioLabel(code);
            html += `<div class="tree-item" data-foglio="${code}" onclick="navigateToFoglio('${code}')">
                ${code}${label ? ' - ' + label : ''}
            </div>`;
        });
    }
    
    // Nota Integrativa (collassabile)
    if (grouped['Nota Integrativa'].length > 0) {
        html += `<div class="tree-section" style="margin-top: 15px;">
            <strong>Nota Integrativa (${grouped['Nota Integrativa'].length} fogli)</strong>
        </div>`;
        html += '<div class="tree-collapsible" style="max-height: 300px; overflow-y: auto;">';
        grouped['Nota Integrativa'].forEach(code => {
            const label = getFoglioLabel(code);
            html += `<div class="tree-item" data-foglio="${code}" onclick="navigateToFoglio('${code}')">
                ${code}${label ? ' - ' + label : ''}
            </div>`;
        });
        html += '</div>';
    }
    
    // Altri
    if (grouped['Altri'].length > 0) {
        html += `<div class="tree-section" style="margin-top: 15px;">
            <strong>Altri (${grouped['Altri'].length})</strong>
        </div>`;
        html += '<div class="tree-collapsible" style="max-height: 200px; overflow-y: auto;">';
        grouped['Altri'].forEach(code => {
            const label = getFoglioLabel(code);
            html += `<div class="tree-item" data-foglio="${code}" onclick="navigateToFoglio('${code}')">
                ${code}${label ? ' - ' + label : ''}
            </div>`;
        });
        html += '</div>';
    }
    
    html += '</div>';
    indexTree.innerHTML = html;
}

// Get label foglio (helper) - usa stessa fonte del breadcrumb
function getFoglioLabel(code) {
    // Leggi dinamicamente dal workbookData
    if (!workbookData || !workbookData.sheets || !workbookData.sheets[code]) {
        return '';
    }

    const sheetData = workbookData.sheets[code];

    // POSIZIONE PRINCIPALE: riga 6, colonna 1 (es. "Conto economico abbreviato")
    if (sheetData[6] && sheetData[6][1]) {
        return sheetData[6][1];
    }

    // FALLBACK: riga 4, colonna 1
    if (sheetData[4] && sheetData[4][1]) {
        return sheetData[4][1];
    }

    // Ultimo fallback: stringa vuota
    return '';
}

// Navigazione a foglio
async function navigateToFoglio(codice) {
    try {
        showLoading(`Caricamento ${codice}...`);
        
        // Gestione speciale per Configurazione
        if (codice === 'Configurazione') {
            currentFoglio = codice;

            // Evidenzia nell'indice
            document.querySelectorAll('.tree-item').forEach(item => {
                item.classList.remove('active');
                if (item.dataset.foglio === codice) {
                    item.classList.add('active');
                }
            });
            
            // Renderizza configurazione
            renderFoglio(codice);
            hideLoading();
            return;
        }
        
        // Carica template se necessario
        const template = loadTemplate(codice);
        if (!template) {
            throw new Error(`Template ${codice} non disponibile`);
        }
        
        // Aggiorna stato UI
        currentFoglio = codice;

        // Evidenzia nell'indice
        document.querySelectorAll('.tree-item').forEach(item => {
            item.classList.remove('active');
            if (item.dataset.foglio === codice) {
                item.classList.add('active');
            }
        });
        
        // Renderizza foglio
        renderFoglio(codice);
        
        hideLoading();
    } catch (error) {
        console.error('Navigation error:', error);
        hideLoading();
        showToast('Errore caricamento foglio: ' + error.message, 'error');
    }
}

// Setup event listeners
function setupEventListeners() {
    // Nuovo Bilancio
    document.getElementById('btn-nuovo').addEventListener('click', nuovoBilancio);
    
    // Carica Excel
    document.getElementById('btn-carica').addEventListener('click', () => {
        document.getElementById('file-input').click();
    });
    
    document.getElementById('file-input').addEventListener('change', handleFileUpload);
    
    // Salva
    document.getElementById('btn-salva').addEventListener('click', salvaBilancio);

    // Export JSON
    document.getElementById('btn-export-json').addEventListener('click', exportJSON);
}

// Nuovo bilancio
function nuovoBilancio() {
    if (bilancio && !confirm('Creare un nuovo bilancio? I dati non salvati andranno persi.')) {
        return;
    }
    
    // Usa date da Configurazione se disponibili
    let annoCorrente = new Date().getFullYear();
    let annoPrecedente = annoCorrente - 1;
    
    if (configurazione && configurazione[8]) {
        const dataStr = configurazione[8][1];
        if (dataStr) {
            const year = new Date(dataStr).getFullYear();
            if (!isNaN(year)) annoCorrente = year;
            annoPrecedente = annoCorrente - 1;
        }
    }
    
    bilancio = {
        metadata: {
            versione: '1.0',
            data_creazione: new Date().toISOString(),
            data_modifica: new Date().toISOString(),
            ragione_sociale: null,
            anno_esercizio: annoCorrente,
            anno_precedente: annoPrecedente
        },
        fogli: {}
    };
    
    // Inizializza struttura vuota per tutti i fogli disponibili
    templatesList.forEach(code => {
        bilancio.fogli[code] = {};
    });
    
    // Abilita pulsanti
    enableButtons();
    
    // Vai al primo foglio
    if (templatesList.length > 0) {
        navigateToFoglio(templatesList[0]);
    }
    
    showToast('Nuovo bilancio creato', 'success');
}

// Upload file XLS
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        showLoading('Caricamento file Excel...');
        
        const xlsData = await parseXLS(file);
        
        // Crea bilancio da XLS
        bilancio = await importFromXLS(xlsData);
        
        enableButtons();
        
        // Vai al primo foglio disponibile nel bilancio
        const firstFoglio = Object.keys(bilancio.fogli).find(f => 
            Object.keys(bilancio.fogli[f]).length > 0
        );
        
        if (firstFoglio) {
            navigateToFoglio(firstFoglio);
        }
        
        hideLoading();
        showToast('File caricato con successo', 'success');
    } catch (error) {
        console.error('File upload error:', error);
        hideLoading();
        showToast('Errore caricamento file: ' + error.message, 'error');
    }
}

// Carica bilancio salvato
function loadSavedBilancio() {
    const saved = localStorage.getItem('workbookabb_bilancio');
    if (saved) {
        try {
            bilancio = JSON.parse(saved);
            enableButtons();
            console.log('✓ Bilancio salvato caricato');
        } catch (error) {
            console.error('Error loading saved bilancio:', error);
        }
    }
}

// Salva bilancio
function salvaBilancio() {
    if (!bilancio) return;

    // Cancella auto-save pendente per evitare doppi messaggi
    if (typeof autoSaveTimeout !== 'undefined' && autoSaveTimeout !== null) {
        clearTimeout(autoSaveTimeout);
        autoSaveTimeout = null;
    }

    bilancio.metadata.data_modifica = new Date().toISOString();
    localStorage.setItem('workbookabb_bilancio', JSON.stringify(bilancio));

    showToast('Bilancio salvato', 'success');
}

// Enable buttons
function enableButtons() {
    document.getElementById('btn-salva').disabled = false;
    document.getElementById('btn-export-json').disabled = false;
}

// Export JSON
function exportJSON() {
    if (!bilancio) return;
    
    const json = JSON.stringify(bilancio, null, 2);
    downloadFile(json, 'bilancio_abbreviato.json', 'application/json');
    
    showToast('JSON esportato', 'success');
}

// Download helper
function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// UI Helpers
function showLoading(text = 'Caricamento...') {
    document.getElementById('loading-text').textContent = text;
    document.getElementById('loading').classList.add('show');
}

function hideLoading() {
    document.getElementById('loading').classList.remove('show');
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

// Getter per altre parti dell'app
function getBilancio() {
    return bilancio;
}

function getCurrentFoglio() {
    return currentFoglio;
}

function getXBRLMappings() {
    return xbrlMappings;
}

function getTemplate(codice) {
    return loadTemplate(codice);
}

function getConfigurazione() {
    return configurazione;
}

function getTemplatesList() {
    return templatesList;
}

function getWorkbookData() {
    return workbookData;
}

// 🔍 DEBUG: Funzione helper per analizzare struttura bilancio
window.debugBilancio = function() {
    const bilancio = getBilancio();
    if (!bilancio) {
        console.log('❌ Nessun bilancio caricato');
        return;
    }

    console.log('=== 📊 STRUTTURA BILANCIO ===\n');
    console.log('Metadata:', bilancio.metadata);

    console.log('\n=== 📋 FOGLI COMPILATI ===');
    for (const [codice, dati] of Object.entries(bilancio.fogli)) {
        const celle = Object.keys(dati).filter(k => dati[k] !== null && dati[k] !== '' && dati[k] !== undefined);
        if (celle.length > 0) {
            console.log(`\n${codice} (${celle.length} celle):`);
            celle.slice(0, 5).forEach(cella => {
                const hasSuffix = cella.includes('_');
                const marker = hasSuffix ? '✓' : '✗';
                console.log(`  ${marker} ${cella} = ${JSON.stringify(dati[cella])}`);
            });
            if (celle.length > 5) {
                console.log(`  ... altre ${celle.length - 5} celle`);
            }
        }
    }

    console.log('\n=== 📈 STATISTICHE ===');
    const totalSheets = Object.keys(bilancio.fogli).length;
    const compiledSheets = Object.entries(bilancio.fogli).filter(([_, dati]) =>
        Object.values(dati).some(v => v !== null && v !== '' && v !== undefined)
    ).length;
    const totalCells = Object.values(bilancio.fogli).reduce((sum, dati) =>
        sum + Object.values(dati).filter(v => v !== null && v !== '' && v !== undefined).length, 0
    );

    console.log(`Fogli totali: ${totalSheets}`);
    console.log(`Fogli compilati: ${compiledSheets}`);
    console.log(`Celle compilate: ${totalCells}`);

    console.log('\n💡 Usa JSON.stringify(getBilancio(), null, 2) per vedere la struttura completa');

    return bilancio;
};

// Salva configurazione bilancio
function salvaConfigurazione() {
    if (!bilancio) return;
    
    const inizioCorrente = document.getElementById('config_inizio_corrente').value;
    const fineCorrente = document.getElementById('config_fine_corrente').value;
    const inizioPrecedente = document.getElementById('config_inizio_precedente').value;
    const finePrecedente = document.getElementById('config_fine_precedente').value;
    
    // Calcola anni da date fine
    const annoCorrente = fineCorrente ? new Date(fineCorrente).getFullYear() : new Date().getFullYear();
    const annoPrecedente = finePrecedente ? new Date(finePrecedente).getFullYear() : annoCorrente - 1;
    
    // Aggiorna metadata
    bilancio.metadata.inizio_corrente = inizioCorrente;
    bilancio.metadata.fine_corrente = fineCorrente;
    bilancio.metadata.inizio_precedente = inizioPrecedente;
    bilancio.metadata.fine_precedente = finePrecedente;
    bilancio.metadata.anno_esercizio = annoCorrente;
    bilancio.metadata.anno_precedente = annoPrecedente;
    bilancio.metadata.valuta = 'EUR'; // Sempre EUR per bilanci italiani
    
    // Salva
    salvaBilancio();
    
    // Ricarica il foglio corrente se siamo in un foglio (per aggiornare gli header anni)
    if (currentFoglio && currentFoglio !== 'Configurazione') {
        renderFoglio(currentFoglio);
    }
    
    showToast('Configurazione salvata', 'success');
}
