// workbookabb - Main Application Module

// Variabili globali
let templates = {};
let xbrlMappings = null;
let currentFoglio = null;
let bilancio = null;

// Inizializzazione
document.addEventListener('DOMContentLoaded', async () => {
    console.log('Initializing workbookabb...');
    
    try {
        showLoading('Caricamento risorse...');
        
        // Carica mappature XBRL
        await loadXBRLMappings();
        
        // Carica lista fogli disponibili
        await loadTemplatesList();
        
        // Genera indice
        renderIndex();
        
        // Setup event listeners
        setupEventListeners();
        
        // Verifica se c'è un bilancio salvato
        loadSavedBilancio();
        
        hideLoading();
        
        console.log('App initialized successfully');
    } catch (error) {
        console.error('Initialization error:', error);
        hideLoading();
        showToast('Errore inizializzazione: ' + error.message, 'error');
    }
});

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

// Carica lista template fogli
async function loadTemplatesList() {
    // Lista fogli principali (in produzione caricare da index o file)
    const fogliPrincipali = [
        'T0000', 'T0002', 'T0006', 'T0009', 'T0011',
        'T0166', 'T0167' // Esempi
    ];
    
    templates = {};
    for (const codice of fogliPrincipali) {
        templates[codice] = { code: codice };
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
        
        return templates[codice];
    } catch (error) {
        console.error(`Error loading template ${codice}:`, error);
        return null;
    }
}

// Genera indice navigazione
function renderIndex() {
    const indexTree = document.getElementById('index-tree');
    
    const structure = {
        'Informazioni Generali': 'T0000',
        'Stato Patrimoniale': 'T0002',
        'Conto Economico': 'T0006',
        'Rendiconto Finanziario (indiretto)': 'T0009',
        'Rendiconto Finanziario (diretto)': 'T0011',
        'Nota Integrativa': {
            'Immobilizzazioni Immateriali': {
                'Movimenti': 'T0166',
                'Commento': 'T0167'
            }
        }
    };
    
    indexTree.innerHTML = renderTreeItems(structure);
}

function renderTreeItems(items, level = 0) {
    let html = '<div class="tree-submenu">';
    
    for (const [label, value] of Object.entries(items)) {
        if (typeof value === 'string') {
            // Foglio
            html += `<div class="tree-item" data-foglio="${value}" onclick="navigateToFoglio('${value}')">
                ${label}
            </div>`;
        } else {
            // Sezione con sottomenu
            html += `<div class="tree-item has-children">${label}</div>`;
            html += renderTreeItems(value, level + 1);
        }
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
    breadcrumb.textContent = `📄 ${codice}`;
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
    
    bilancio = {
        metadata: {
            versione: '1.0',
            data_creazione: new Date().toISOString(),
            data_modifica: new Date().toISOString(),
            ragione_sociale: null,
            anno_esercizio: new Date().getFullYear()
        },
        fogli: {}
    };
    
    // Inizializza tutti i fogli vuoti
    for (const codice in templates) {
        bilancio.fogli[codice] = {};
    }
    
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
