// workbookabb - Data Manager Module

// Aggiorna valore cella
function updateCellValue(foglioCode, codiceCella, valore) {
    const bilancio = getBilancio();
    
    if (!bilancio) {
        console.error('Nessun bilancio caricato');
        return;
    }
    
    if (!bilancio.fogli[foglioCode]) {
        bilancio.fogli[foglioCode] = {};
    }
    
    bilancio.fogli[foglioCode][codiceCella] = valore;
    
    // Auto-save in localStorage
    autoSave();
    
    console.log(`Updated ${foglioCode}.${codiceCella} = ${valore}`);
}

// Auto-save
let autoSaveTimeout = null;
function autoSave() {
    clearTimeout(autoSaveTimeout);
    autoSaveTimeout = setTimeout(() => {
        salvaBilancio();
    }, 2000); // Salva dopo 2 secondi di inattività
}

// Valida valore cella
function validateCellValue(codiceCella, valore, foglioCode) {
    // Cerca tipo da xbrl_mappings
    const xbrlMappings = getXBRLMappings();
    if (!xbrlMappings) return { valid: true };
    
    const mapping = findMappingByCode(codiceCella, foglioCode, xbrlMappings);
    if (!mapping || !mapping.xbrl) return { valid: true };
    
    const type = mapping.xbrl.type;
    
    // Validazione base per tipo
    if (type === 'monetary' || type === 'decimal') {
        if (valore !== null && valore !== '' && isNaN(parseFloat(valore))) {
            return {
                valid: false,
                message: 'Valore numerico non valido'
            };
        }
    }
    
    return { valid: true };
}

// Calcola formule automatiche (esempio base)
function recalculateFormulas(foglioCode) {
    // TODO: Implementare logica formule
    // Esempio: Totale Immobilizzazioni = Immateriali + Materiali + Finanziarie
    
    console.log('Formule calcolo non ancora implementate');
}

// Esporta dati foglio come array (per debug)
function exportFoglioData(foglioCode) {
    const bilancio = getBilancio();
    if (!bilancio) return null;
    
    return bilancio.fogli[foglioCode] || {};
}

// Statistiche bilancio
function getBilancioStats() {
    const bilancio = getBilancio();
    if (!bilancio) return null;
    
    let totalCelle = 0;
    let celleCompilate = 0;
    
    for (const foglio in bilancio.fogli) {
        const dati = bilancio.fogli[foglio];
        for (const cella in dati) {
            totalCelle++;
            const valore = dati[cella];
            if (valore !== null && valore !== '' && valore !== undefined) {
                celleCompilate++;
            }
        }
    }
    
    return {
        totalCelle,
        celleCompilate,
        percentualeCompletamento: totalCelle > 0 ? Math.round((celleCompilate / totalCelle) * 100) : 0,
        ultimaModifica: bilancio.metadata?.data_modifica
    };
}

// Pulisci bilancio (reset)
function resetBilancio() {
    if (!confirm('Sei sicuro di voler cancellare tutti i dati? Questa operazione non può essere annullata.')) {
        return false;
    }
    
    localStorage.removeItem('workbookabb_bilancio');
    location.reload();
    return true;
}

// Helper per trovare mapping
function findMappingByCode(codiceExcel, foglioCode, xbrlMappings) {
    if (!xbrlMappings || !xbrlMappings.mappature) return null;

    const foglioMappings = xbrlMappings.mappature[foglioCode];
    if (!foglioMappings) return null;

    // Cerca match esatto (supporta sia 'code' che 'codice_excel' per compatibilità)
    let mapping = foglioMappings.find(m => m.code === codiceExcel || m.codice_excel === codiceExcel);

    // Se non trovato, cerca per prefisso (per celle combinate riga_colonna)
    if (!mapping) {
        const codiceBase = codiceExcel.split('_')[0];
        mapping = foglioMappings.find(m => m.code === codiceBase || m.codice_excel === codiceBase);
    }

    return mapping;
}
