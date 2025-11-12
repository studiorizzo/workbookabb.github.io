// workbookabb - Main Application Module (FIXED)

// Variabili globali
let configurazione = null;
let indice = null;
let templates = {};
let xbrlMappings = null;
let currentFoglio = null;
let bilancio = null;

// Inizializzazione
document.addEventListener('DOMContentLoaded', async () => {
    console.log('Initializing workbookabb...');
    
    try {
        showLoading('Caricamento risorse...');
        
        // 1. Carica Configurazione (date, contesti)
        await loadConfigur

azione();
        
        // 2. Carica Indice (struttura sidebar)
        await loadIndice();
        
        // 3. Carica mappature XBRL
        await loadXBRLMappings();
        
        // 4. Genera indice sidebar
        renderIndex();
        
        // 5. Setup event listeners
        setupEventListeners();
        
        // 6. Verifica bilancio salvato
        loadSavedBilancio();
        
        hideLoading();
        
        console.log('App initialized successfully');
    } catch (error) {
        console.error('Initialization error:', error);
        hideLoading();
        showToast('Errore inizializzazione: ' + error.message, 'error');
    }
});

// Carica Configurazione.json
async function loadConfigurazione() {
    try {
        const response = await fetch('data/template/Configurazione.json');
        if (!response.ok) throw new Error('Configurazione.json non trovato');
        configurazione = await response.json();
        console.log('Configurazione loaded:', configurazione);
    } catch (error) {
        console.error('Error loading Configurazione:', error);
        configurazione = null;
    }
}

// Carica Indice.json
async function loadIndice() {
    try {
        const response = await fetch('data/template/Indice.json');
        if (!response.ok) throw new Error('Indice.json non trovato');
        indice = await response.json();
        console.log('Indice loaded:', indice.length, 'righe');
    } catch (error) {
        console.error('Error loading Indice:', error);
        // Fallback a lista hardcoded minima
        indice = null;
    }
}

// Carica mappature XBRL
async function loadXBRLMappings() {
    try {
        const response = await fetch('data/mapping/xbrl_mappings_complete.json');
        if (!response.ok) throw new Error('xbrl_mappings non trovato');
        xbrlMappings = await response.json();
        console.log('XBRL mappings loaded:', Object.keys(xbrlMappings.mappature || {}).length, 'fogli');
    } catch (error) {
        console.warn('xbrl_mappings not available:', error.message);
        xbrlMappings = { mappature: {} };
    }
}

// Carica template foglio specifico
async function loadTemplate(codice) {
    if (templates[codice] && templates[codice].loaded) {
        return templates[codice];
    }
    
    try {
        const response = await fetch(`data/template/${codice}.json`);
        if (!response.ok) throw new Error(`Template ${codice} non trovato`);
        
        const data = await response.json();
        templates[codice] = {
            code: codice,
            data: data,
            loaded: true
        };
        
        console.log(`Template ${codice} loaded:`, data.length, 'righe');
        return templates[codice];
    } catch (error) {
        console.error(`Error loading template ${codice}:`, error);
        return null;
    }
}

// Genera indice navigazione da Indice.json
function renderIndex() {
    const indexTree = document.getElementById('index-tree');
    
    if (!indice || indice.length === 0) {
        // Fallback a struttura minima hardcoded
        indexTree.innerHTML = renderFallbackIndex();
        return;
    }
    
    // Costruisci albero da Indice.json
    let html = '<div class="tree-submenu">';
    
    // Mappa foglio corrente (T0000, T0002, etc.)
    let currentTCode = null;
    
    indice.forEach((riga, idx) => {
        // Skip righe vuote o header
        if (!riga || riga.length < 3) return;
        
        const col1 = riga[1]; // Vuoto o codice foglio
        const col2 = riga[2]; // Label/titolo
        
        // Rileva se è un foglio navigabile (inizia con spazio = indentato)
        const isIndented = col2 && col2.startsWith('  ');
        const label = col2 ? col2.trim() : '';
        
        if (!label || label === '') return;
        
        // Cerca codice T0xxx nelle righe successive
        // (L'indice ha una struttura particolare)
        // Logica semplificata: fogli principali non indentati
        
        if (!isIndented && label.length > 3) {
            // Possibile foglio principale
            // Mappa manualmente i principali (quick fix)
            let tCode = null;
            if (label.includes('Informazioni generali')) tCode = 'T0000';
            else if (label.includes('Stato patrimoniale')) tCode = 'T0002';
            else if (label.includes('Conto economico')) tCode = 'T0006';
            else if (label.includes('Rendiconto finanziario, metodo indiretto')) tCode = 'T0009';
            else if (label.includes('Rendiconto finanziario, metodo diretto')) tCode = 'T0011';
            
            if (tCode) {
                html += `<div class="tree-item" data-foglio="${tCode}" onclick="navigateToFoglio('${tCode}')">
                    ${label}
                </div>`;
                currentTCode = tCode;
            } else {
                // Header di sezione
                html += `<div class="tree-item has-children">${label}</div>`;
            }
        } else if (isIndented) {
            // Foglio indentato (Nota Integrativa)
            // Per ora skip (troppo complesso mapparli tutti)
            // TODO: mappare T0166, T0167, etc.
        }
    });
    
    html += '</div>';
    indexTree.innerHTML = html;
}

// Fallback index se Indice.json non carica
function renderFallbackIndex() {
    const structure = {
        'Informazioni Generali': 'T0000',
        'Stato Patrimoniale': 'T0002',
        'Conto Economico': 'T0006',
        'Rendiconto Finanziario (indiretto)': 'T0009',
        'Rendiconto Finanziario (diretto)': 'T0011'
    };
    
    let html = '<div class="tree-submenu">';
    for (const [label, code] of Object.entries(structure)) {
        html += `<div class="tree-item" data-foglio="${code}" onclick="navigateToFoglio('${code}')">
            ${label}
        </div>`;
    }
    html += '</div>';
    return html;
}

// Navigazione a foglio
async function navigateToFoglio(codice) {
    try {
        showLoading(`Caricamento ${codice}...`);
        
        // Carica template se necessario
        const template = await loadTemplate(codice);
        if (!template) {
            throw new Error(`Template ${codice} non disponibile`);
        }
        
        // Aggiorna stato UI
        currentFoglio = codice;
        updateBreadcrumb(codice);
        
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

// Update breadcrumb
function updateBreadcrumb(codice) {
    const breadcrumb = document.getElementById('breadcrumb');
    const template = templates[codice];
    
    // Estrai titolo dal template
    let titolo = codice;
    if (template && template.data && template.data.length > 6) {
        titolo = template.data[6]?.[1] || template.data[4]?.[1] || codice;
    }
    
    breadcrumb.textContent = `${codice} - ${titolo}`;
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
    
    // Export XLS
    document.getElementById('btn-export-xls').addEventListener('click', exportXLS);
    
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
    let annoPrec = annoCorrente - 1;
    
    if (configurazione && configurazione[8]) {
        annoCorrente = configurazione[8][7] || annoCorrente;
        annoPrec = configurazione[9]?.[7] || annoPrec;
    }
    
    bilancio = {
        metadata: {
            versione: '1.0',
            data_creazione: new Date().toISOString(),
            data_modifica: new Date().toISOString(),
            ragione_sociale: null,
            anno_esercizio: annoCorrente,
            anno_precedente: annoPrec
        },
        fogli: {}
    };
    
    // Abilita pulsanti
    enableButtons();
    
    // Vai al primo foglio
    navigateToFoglio('T0000');
    
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
        navigateToFoglio('T0000');
        
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
            console.log('Bilancio salvato caricato');
        } catch (error) {
            console.error('Error loading saved bilancio:', error);
        }
    }
}

// Salva bilancio
function salvaBilancio() {
    if (!bilancio) return;
    
    bilancio.metadata.data_modifica = new Date().toISOString();
    localStorage.setItem('workbookabb_bilancio', JSON.stringify(bilancio));
    
    showToast('Bilancio salvato', 'success');
}

// Enable buttons
function enableButtons() {
    document.getElementById('btn-salva').disabled = false;
    document.getElementById('btn-export-xls').disabled = false;
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
    return templates[codice];
}

function getConfigurazione() {
    return configurazione;
}
