/* =========================================
   1. DATA CONFIGURATION & STORAGE
   ========================================= */

const STORAGE_KEY = 'ratrix_data_v3_clients'; 
const OLD_STORAGE_KEY = 'ratrix_data_v2_profiles'; 

const DEFAULT_LIMITS = [50, 100, 150, 500];

const MODEL_KEYS = [
    'fixed', 'flat', 'minFixed', 'cumulative', 
    'minCumulative', 'minExcess', 'excess'
];

const SERVICE_MODES = [
    "DOOR TO DOOR",
    "PORT TO PORT",
    "DOOR TO PORT",
    "PORT TO DOOR"
];

/* DATA STRUCTURE (V3) */
let GLOBAL_STORE = {
    activeClientId: null,
    clients: {}
};

// --- HELPERS ---

function generateId(prefix = 'id_') {
    return prefix + Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function createEmptyRow(limitCount) {
    return {
        origin: "",
        dest: "",
        rates: new Array(limitCount).fill(null) 
    };
}

function createFreshClientStore() {
    let store = {};
    MODEL_KEYS.forEach(key => {
        const pid = generateId('p_');
        const initialLimits = [...DEFAULT_LIMITS];
        
        store[key] = {
            activeId: pid,
            profiles: {}
        };
        
        store[key].profiles[pid] = {
            name: "Standard Table",
            limits: initialLimits,
            rows: [createEmptyRow(initialLimits.length)]
        };
    });
    return store;
}

// --- PERSISTENCE ---

function saveToLocal() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(GLOBAL_STORE));
}

function loadFromLocal() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
        try {
            GLOBAL_STORE = JSON.parse(raw);
            return true;
        } catch (e) { console.error("Corrupted V3 store", e); }
    }
    
    // --- MIGRATION (Legacy) ---
    const oldRaw = localStorage.getItem(OLD_STORAGE_KEY);
    if (oldRaw) {
        try {
            console.log("Migrating V2 Data to V3 Clients...");
            const oldData = JSON.parse(oldRaw);
            
            let migratedStore = oldData.store || oldData.data_store;
            if(!migratedStore) return false;

            const globalLimits = oldData.limits || DEFAULT_LIMITS;
            MODEL_KEYS.forEach(key => {
                if (migratedStore[key] && migratedStore[key].profiles) {
                    Object.values(migratedStore[key].profiles).forEach(p => {
                        if (!p.limits) p.limits = [...globalLimits];
                        p.rows.forEach(r => delete r.serviceMode); 
                    });
                }
            });

            const clientId = generateId('c_');
            GLOBAL_STORE.clients[clientId] = {
                id: clientId,
                name: "Legacy Data (Migrated)",
                description: "Imported from previous version",
                data_store: migratedStore
            };
            GLOBAL_STORE.activeClientId = clientId;
            saveToLocal();
            return true;
        } catch(e) { console.error("Migration failed", e); }
    }

    return false;
}

// --- INITIALIZATION ---

function initApp() {
    const loaded = loadFromLocal();
    if (!loaded) {
        const cid = generateId('c_');
        GLOBAL_STORE.clients[cid] = {
            id: cid,
            name: "Default Client",
            description: "",
            data_store: createFreshClientStore()
        };
        GLOBAL_STORE.activeClientId = cid;
        saveToLocal();
    }
    renderClientList();
    renderAppForActiveClient();
}

// --- CLIENT MANAGEMENT LOGIC ---

function getActiveClient() {
    if(!GLOBAL_STORE.activeClientId || !GLOBAL_STORE.clients[GLOBAL_STORE.activeClientId]) {
        const keys = Object.keys(GLOBAL_STORE.clients);
        if(keys.length > 0) {
            GLOBAL_STORE.activeClientId = keys[0];
            return GLOBAL_STORE.clients[keys[0]];
        }
        return null;
    }
    return GLOBAL_STORE.clients[GLOBAL_STORE.activeClientId];
}

/* =========================================
   2. UI: CLIENT SIDEBAR
   ========================================= */

const elClientList = document.getElementById('clientList');
const elClientDesc = document.getElementById('clientDescription');
const elHeaderClientName = document.getElementById('headerClientName');
const elViewClientName = document.getElementById('viewClientName');

function renderClientList() {
    elClientList.innerHTML = '';
    const activeId = GLOBAL_STORE.activeClientId;

    Object.values(GLOBAL_STORE.clients).forEach(client => {
        const div = document.createElement('div');
        div.className = `client-item ${client.id === activeId ? 'active' : ''}`;
        
        const nameSpan = document.createElement('span');
        nameSpan.className = 'client-name-area';
        nameSpan.textContent = client.name;
        nameSpan.onclick = () => switchClient(client.id, 'editor');

        const viewBtn = document.createElement('button');
        viewBtn.className = 'btn-sidebar-view';
        viewBtn.textContent = 'View Tables';
        viewBtn.title = "View Read-Only Tables";
        viewBtn.onclick = (e) => {
            e.stopPropagation(); 
            switchClient(client.id, 'viewer');
        };

        div.appendChild(nameSpan);
        div.appendChild(viewBtn);
        elClientList.appendChild(div);
    });

    const current = getActiveClient();
    if (current) {
        elHeaderClientName.textContent = current.name;
        elViewClientName.textContent = current.name; 
        elClientDesc.value = current.description || "";
    }
}

function switchClient(id, viewMode = 'editor') {
    GLOBAL_STORE.activeClientId = id;
    saveToLocal();
    renderClientList();
    renderAppForActiveClient(); 
    switchView(viewMode);
}

document.getElementById('btnNewClient').addEventListener('click', () => {
    const name = prompt("Enter Client Name:");
    if (!name) return;
    
    const cid = generateId('c_');
    GLOBAL_STORE.clients[cid] = {
        id: cid,
        name: name,
        description: "",
        data_store: createFreshClientStore() 
    };
    switchClient(cid, 'editor');
});

elClientDesc.addEventListener('input', (e) => {
    const client = getActiveClient();
    if(client) {
        client.description = e.target.value;
        saveToLocal();
    }
});

document.getElementById('btnDeleteClient').addEventListener('click', () => {
    const client = getActiveClient();
    if(!client) return;
    
    if(Object.keys(GLOBAL_STORE.clients).length <= 1) {
        alert("Cannot delete the last remaining client.");
        return;
    }

    if(confirm(`Delete client "${client.name}" and all their rates? This cannot be undone.`)) {
        delete GLOBAL_STORE.clients[client.id];
        GLOBAL_STORE.activeClientId = Object.keys(GLOBAL_STORE.clients)[0];
        saveToLocal();
        renderClientList();
        renderAppForActiveClient();
    }
});

/* =========================================
   3. VIEW TOGGLING (Editor vs Viewer)
   ========================================= */

function switchView(mode) {
    const editor = document.getElementById('editor-view');
    const viewer = document.getElementById('viewer-view');
    
    if (mode === 'editor') {
        editor.classList.add('active-view');
        viewer.classList.remove('active-view');
    } else {
        saveToLocal();
        renderReadOnlyDashboard(); 
        editor.classList.remove('active-view');
        viewer.classList.add('active-view');
    }
}

window.switchView = switchView; 

document.getElementById('btnSaveAndOpenView').addEventListener('click', () => {
    saveToLocal();
    switchView('viewer');
});


/* =========================================
   4. READ-ONLY DASHBOARD RENDERER
   ========================================= */

function renderReadOnlyDashboard() {
    const container = document.getElementById('viewer-content');
    container.innerHTML = '';
    
    const client = getActiveClient();
    if (!client) return;

    const dataStore = client.data_store;
    let hasData = false;

    MODEL_KEYS.forEach(modelKey => {
        const modelData = dataStore[modelKey];
        if (!modelData || !modelData.profiles) return;

        const profiles = Object.values(modelData.profiles);
        if (profiles.length === 0) return;

        hasData = true;

        // Create Section for Model (Wrapper)
        const section = document.createElement('div');
        section.className = 'model-section';
        
        // --- UPDATED: REMOVED MODEL TITLE ---
        // The title creation code has been deleted here.

        // Render each table in this model
        profiles.forEach(profile => {
            const card = document.createElement('div');
            card.className = 'ro-table-card';
            
            const tableTitle = document.createElement('div');
            tableTitle.className = 'ro-table-title';
            tableTitle.innerHTML = `📄 ${profile.name}`;
            card.appendChild(tableTitle);

            // Build Table HTML
            let thHtml = `<th class="ro-origin-dest">Origin</th><th class="ro-origin-dest">Destination</th>`;
            profile.limits.forEach((lim, i) => {
                const prev = i === 0 ? 1 : profile.limits[i-1] + 1;
                thHtml += `<th>${prev}-${lim}</th>`;
            });

            let rowsHtml = '';
            profile.rows.forEach(row => {
                // Skip empty rows
                if(!row.origin && !row.dest) return;
                
                let tds = `<td class="ro-origin-dest">${row.origin}</td><td class="ro-origin-dest">${row.dest}</td>`;
                row.rates.forEach(r => {
                    tds += `<td>${r === null ? '-' : r}</td>`;
                });
                rowsHtml += `<tr>${tds}</tr>`;
            });

            if(rowsHtml === '') rowsHtml = '<tr><td colspan="100%" style="font-style:italic; color:var(--text-muted);">No routes defined.</td></tr>';

            const tableHtml = `
                <div class="ro-table-wrapper">
                    <table class="ro-table">
                        <thead><tr>${thHtml}</tr></thead>
                        <tbody>${rowsHtml}</tbody>
                    </table>
                </div>
            `;
            
            card.innerHTML += tableHtml;
            section.appendChild(card);
        });

        container.appendChild(section);
    });

    if(!hasData) {
        container.innerHTML = '<div style="text-align:center; color:var(--text-muted);">No tables found for this client.</div>';
    }
}


/* =========================================
   5. APP LOGIC (Editor & Calc)
   ========================================= */

function getActiveClientDataStore() {
    const client = getActiveClient();
    return client ? client.data_store : null;
}

const elTableSelect = document.getElementById('tableProfileSelect');
const elModel = document.getElementById('pricingModel');

function renderAppForActiveClient() {
    renderProfileDropdown();
    renderTableStructure();
    calculate();
}

function getActiveModelData() {
    const store = getActiveClientDataStore();
    const model = elModel.value;
    if (!store[model]) {
        const pid = generateId('p_');
        const initialLimits = [...DEFAULT_LIMITS];
        store[model] = { activeId: pid, profiles: {} };
        store[model].profiles[pid] = { name: "Standard Table", limits: initialLimits, rows: [createEmptyRow(initialLimits.length)] };
    }
    return store[model];
}

function getActiveProfile() {
    const modelData = getActiveModelData();
    const pid = modelData.activeId;
    if (!modelData.profiles[pid]) {
        const firstKey = Object.keys(modelData.profiles)[0];
        if(!firstKey) return null; 
        modelData.activeId = firstKey;
        return modelData.profiles[firstKey];
    }
    return modelData.profiles[pid];
}

function renderProfileDropdown() {
    const modelData = getActiveModelData();
    elTableSelect.innerHTML = '';
    
    Object.keys(modelData.profiles).forEach(pid => {
        const p = modelData.profiles[pid];
        const option = document.createElement('option');
        option.value = pid;
        option.text = p.name;
        if (pid === modelData.activeId) option.selected = true;
        elTableSelect.appendChild(option);
    });
}

// HANDLERS

elTableSelect.addEventListener('change', (e) => {
    const modelData = getActiveModelData();
    modelData.activeId = e.target.value;
    saveToLocal();
    renderTableStructure();
    calculate();
});

document.getElementById('btnNewProfile').addEventListener('click', () => {
    const name = prompt("Enter table name (e.g. 'Promo Rates'):");
    if (!name) return;

    const modelData = getActiveModelData();
    const newId = generateId('p_');
    const initialLimits = [...DEFAULT_LIMITS];

    modelData.profiles[newId] = {
        name: name,
        limits: initialLimits,
        rows: [createEmptyRow(initialLimits.length)]
    };
    modelData.activeId = newId; 
    
    saveToLocal();
    renderProfileDropdown();
    renderTableStructure();
    calculate();
});

document.getElementById('btnRenameProfile').addEventListener('click', () => {
    const profile = getActiveProfile();
    const newName = prompt("Rename table:", profile.name);
    if (newName && newName.trim() !== "") {
        profile.name = newName;
        saveToLocal();
        renderProfileDropdown();
    }
});

document.getElementById('btnDeleteProfile').addEventListener('click', () => {
    const modelData = getActiveModelData();
    const keys = Object.keys(modelData.profiles);
    
    if (keys.length <= 1) {
        alert("You must have at least one table for this model.");
        return;
    }
    
    const profileName = modelData.profiles[modelData.activeId].name;
    if (confirm(`Delete table "${profileName}"?`)) {
        delete modelData.profiles[modelData.activeId];
        modelData.activeId = Object.keys(modelData.profiles)[0];
        
        saveToLocal();
        renderProfileDropdown();
        renderTableStructure();
        calculate();
    }
});


/* =========================================
   6. CALCULATION FORMULAS
   ========================================= */

const strategies = {
    fixed: (w, rates, limits) => {
        const index = getBracketIndex(w, limits);
        if (index === -1 || !isValid(rates[index])) return null; 
        return w * rates[index];
    },
    flat: (w, rates, limits) => {
        const index = getBracketIndex(w, limits);
        if (index === -1 || !isValid(rates[index])) return null;
        return rates[index];
    },
    minFixed: (w, rates, limits) => {
        if (w <= limits[0]) return isValid(rates[0]) ? rates[0] : null;
        const index = getBracketIndex(w, limits);
        if (index === -1 || !isValid(rates[index])) return null;
        return w * rates[index];
    },
    cumulative: (w, rates, limits) => {
        let total = 0;
        let remaining = w;
        let prevMax = 0;
        for (let i = 0; i < limits.length; i++) {
            if (!isValid(rates[i])) return null; 
            const capacity = limits[i] - prevMax;
            const fill = Math.min(remaining, capacity);
            if (fill > 0) {
                total += fill * rates[i];
                remaining -= fill;
            }
            prevMax = limits[i];
            if (remaining <= 0) break;
        }
        return remaining > 0 ? null : total; 
    },
    minCumulative: (w, rates, limits) => {
        let total = 0;
        let remaining = w;
        let prevMax = 0;
        for (let i = 0; i < limits.length; i++) {
            if (!isValid(rates[i])) return null;
            const capacity = limits[i] - prevMax;
            const fill = Math.min(remaining, capacity);
            if (fill > 0) {
                if (i === 0) total += rates[0]; 
                else total += fill * rates[i];
                remaining -= fill;
            }
            prevMax = limits[i];
            if (remaining <= 0) break;
        }
        return remaining > 0 ? null : total;
    },
    minExcess: (w, rates, limits) => {
        if (!isValid(rates[0]) || !isValid(rates[1])) return null;
        const limit = limits[0]; 
        const baseFlat = rates[0];      
        const excessRate = rates[1];     
        if (w <= limit) return baseFlat;
        return baseFlat + ((w - limit) * excessRate);
    },
    excess: (w, rates, limits) => {
        if (!isValid(rates[0]) || !isValid(rates[1])) return null;
        const limit = limits[0]; 
        const baseRate = rates[0];       
        const excessRate = rates[1];     
        if (w <= limit) return w * baseRate;
        return (limit * baseRate) + ((w - limit) * excessRate);
    }
};

function isValid(val) { return val !== null && val !== "" && !isNaN(val); }

function getBracketIndex(w, limits) {
    const effectiveWeight = Math.floor(w);
    for (let i = 0; i < limits.length; i++) {
        if (effectiveWeight <= limits[i]) return i;
    }
    return -1;
}

/* =========================================
   7. UI RENDERING & EDITING (EDITOR)
   ========================================= */

const elTableHeader = document.getElementById('tableHeaderRow');
const elTableBody = document.getElementById('rateTableBody');
const elResult = document.getElementById('resultPrice');
const elDesc = document.getElementById('formulaDesc');
const elWeight = document.getElementById('weightInput');
const elOrigin = document.getElementById('origin');
const elDest = document.getElementById('destination');
const elServiceMode = document.getElementById('serviceModeInput');
const elDimL = document.getElementById('dimL');
const elDimW = document.getElementById('dimW');
const elDimH = document.getElementById('dimH');
const elVolDivisor = document.getElementById('volDivisor');
const elChargeBasis = document.getElementById('chargeBasis');
const elDispActual = document.getElementById('dispActual');
const elDispVol = document.getElementById('dispVol');
const elDispCbm = document.getElementById('dispCbm');

function renderTableStructure() {
    const profile = getActiveProfile();
    if(!profile) return; 

    const activeRows = profile.rows;
    const currentLimits = profile.limits; 

    // Header - NO Service Mode
    let headerHtml = `
        <th style="width: 100px;">Origin</th>
        <th style="width: 100px;">Destination</th>`;

    currentLimits.forEach((limit, index) => {
        const prevLimit = index === 0 ? 0 : currentLimits[index - 1];
        const startVal = prevLimit + 1;
        
        headerHtml += `
        <th>
            <div class="col-header-container">
                <button class="btn-col-delete" onclick="deleteColumn(${index})" title="Delete Column">✕</button>
                <div class="range-wrapper">
                    <span id="start-${index}" class="header-range-text">${startVal}</span> - 
                    <input type="number" class="header-input" 
                           value="${limit}" 
                           onchange="handleEditLimit(${index}, this.value)">
                </div>
            </div>
        </th>`;
    });
    headerHtml += `<th style="width: 60px;">Action</th>`;
    elTableHeader.innerHTML = headerHtml;

    // Body
    let bodyHtml = '';
    activeRows.forEach((row, rowIndex) => {
        let ratesHtml = '';
        while(row.rates.length < currentLimits.length) row.rates.push(null);

        row.rates.forEach((rate, colIndex) => {
            if(colIndex < currentLimits.length) {
                const displayVal = (rate === null) ? "" : rate;
                ratesHtml += `
                <td>
                    <input type="number" class="editable-input" 
                           value="${displayVal}" 
                           placeholder="-"
                           onchange="handleEditRate(${rowIndex}, ${colIndex}, this.value)">
                </td>`;
            }
        });

        bodyHtml += `
        <tr id="row-${rowIndex}">
            <td>
                <input type="text" class="editable-input text-input" 
                       value="${row.origin}" 
                       placeholder="Origin"
                       onchange="handleEditRoute('origin', ${rowIndex}, this.value)">
            </td>
            <td>
                <input type="text" class="editable-input text-input" 
                       value="${row.dest}" 
                       placeholder="Dest"
                       onchange="handleEditRoute('dest', ${rowIndex}, this.value)">
            </td>
            ${ratesHtml}
            <td>
                <button class="btn-delete" onclick="deleteRow(${rowIndex})">✕</button>
            </td>
        </tr>`;
    });

    elTableBody.innerHTML = bodyHtml;
    syncDropdowns(activeRows);
}

function syncDropdowns(rows) {
    const origins = [...new Set(rows.map(item => item.origin).filter(x => x))];
    const dests = [...new Set(rows.map(item => item.dest).filter(x => x))];
    
    const currOrigin = elOrigin.value;
    const currDest = elDest.value;

    elOrigin.innerHTML = origins.length ? origins.map(o => `<option value="${o}">${o}</option>`).join('') : '<option value="">-</option>';
    elDest.innerHTML = dests.length ? dests.map(d => `<option value="${d}">${d}</option>`).join('') : '<option value="">-</option>';

    if (origins.includes(currOrigin)) elOrigin.value = currOrigin;
    if (dests.includes(currDest)) elDest.value = currDest;
}

function highlightRow(index) {
    const rows = document.querySelectorAll('#rateTableBody tr');
    rows.forEach(r => r.classList.remove('active-route-row'));
    if(index === -1) return;

    const activeRow = document.getElementById(`row-${index}`);
    if (activeRow) {
        activeRow.classList.add('active-route-row');
        activeRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

// ---- ACTIONS ----

document.getElementById('addColBtn').addEventListener('click', () => {
    const profile = getActiveProfile();
    const currentLimits = profile.limits;
    const lastLimit = currentLimits[currentLimits.length - 1] || 0;
    
    profile.limits.push(lastLimit + 50);
    profile.rows.forEach(r => r.rates.push(null));
    
    saveToLocal();
    renderTableStructure();
    calculate();
});

document.getElementById('addRowBtn').addEventListener('click', () => {
    const profile = getActiveProfile();
    profile.rows.push(createEmptyRow(profile.limits.length));
    saveToLocal();
    renderTableStructure();
});

// ---- EDIT HANDLERS ----

window.handleEditLimit = function(index, value) {
    const val = parseFloat(value);
    if (isNaN(val) || val <= 0) return;
    
    const profile = getActiveProfile();
    profile.limits[index] = val;
    saveToLocal();
    renderTableStructure(); 
    calculate();
};

window.handleEditRate = function(rowIndex, colIndex, value) {
    const profile = getActiveProfile();
    if (value === "") profile.rows[rowIndex].rates[colIndex] = null;
    else profile.rows[rowIndex].rates[colIndex] = parseFloat(value);
    
    saveToLocal();
    calculate(); 
};

window.handleEditRoute = function(field, rowIndex, value) {
    const profile = getActiveProfile();
    profile.rows[rowIndex][field] = value;
    saveToLocal();
    syncDropdowns(profile.rows); 
    calculate();
};

window.deleteRow = function(index) {
    const profile = getActiveProfile();
    if (profile.rows.length <= 1) {
        alert("Cannot delete the last remaining row.");
        return;
    }
    profile.rows.splice(index, 1);
    saveToLocal();
    renderTableStructure();
    calculate();
};

window.deleteColumn = function(index) {
    const profile = getActiveProfile();

    if (profile.limits.length <= 1) {
        alert("Cannot delete the last remaining weight column.");
        return;
    }

    if (!confirm("Delete this weight column for THIS table?")) {
        return;
    }

    profile.limits.splice(index, 1);
    profile.rows.forEach(row => {
        if (row.rates.length > index) {
            row.rates.splice(index, 1);
        }
    });

    saveToLocal();
    renderTableStructure();
    calculate();
};

/* =========================================
   8. CALCULATION
   ========================================= */

function calculate() {
    const actualWeight = parseFloat(elWeight.value) || 0;
    const origin = elOrigin.value;
    const dest = elDest.value;
    const serviceMode = elServiceMode.value; 
    const model = elModel.value; 

    const L = parseFloat(elDimL.value) || 0;
    const W = parseFloat(elDimW.value) || 0;
    const H = parseFloat(elDimH.value) || 0;
    const divisor = parseFloat(elVolDivisor.value) || 6000;

    let volWeight = (divisor > 0) ? (L * W * H) / divisor : 0;
    volWeight = parseFloat(volWeight.toFixed(2));
    const cbm = (L * W * H) / 1000000;

    elDispActual.textContent = `${actualWeight} kg`;
    elDispVol.textContent = `${volWeight} kg`; 
    elDispCbm.textContent = `${cbm.toLocaleString(undefined, {maximumFractionDigits: 4})}`;

    const chargeBasis = elChargeBasis.value;
    let chargeableWeight = (chargeBasis === 'volumetric') ? volWeight : actualWeight;

    if (chargeableWeight <= 0) {
        elResult.textContent = "Php 0.00";
        elDesc.textContent = "Please enter valid weight/dimensions";
        highlightRow(-1); 
        return;
    }

    const profile = getActiveProfile();
    if(!profile) return;
    const currentLimits = profile.limits;
    
    // UPDATED MATCHING: Removed Service Mode
    const routeIndex = profile.rows.findIndex(r => 
        r.origin === origin && 
        r.dest === dest
    );
    
    if (routeIndex === -1) {
        elResult.textContent = "Route Not Found";
        elDesc.textContent = `No match for ${origin} -> ${dest}`;
        highlightRow(-1);
        return;
    }
    highlightRow(routeIndex);
    const route = profile.rows[routeIndex];

    const strategyFn = strategies[model] || strategies.fixed;
    const total = strategyFn(chargeableWeight, route.rates, currentLimits);

    if (total === null) {
        const maxLimit = currentLimits[currentLimits.length - 1];
        if (chargeableWeight > maxLimit) {
             elResult.textContent = "Over Limit";
             elDesc.textContent = `Weight (${chargeableWeight}kg) exceeds ${maxLimit}kg limit`;
        } else {
             elResult.textContent = "Invalid Rate";
             elDesc.textContent = "Rate is blank/missing for this bracket.";
        }
    } else {
        elResult.textContent = `Php ${total.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
        const bracketIdx = getBracketIndex(chargeableWeight, currentLimits);
        
        const basisLabel = chargeBasis === 'volumetric' ? 'Vol. Wt.' : 'Act. Wt.';
        elDesc.textContent = `${model} | ${basisLabel} | ${serviceMode}`;
    }
}

// EVENTS
elModel.addEventListener('change', () => {
    renderProfileDropdown();
    renderTableStructure();
    calculate();
});

elWeight.addEventListener('input', calculate);
elOrigin.addEventListener('change', calculate);
elDest.addEventListener('change', calculate);
elServiceMode.addEventListener('change', calculate); 

elDimL.addEventListener('input', calculate);
elDimW.addEventListener('input', calculate);
elDimH.addEventListener('input', calculate);
elVolDivisor.addEventListener('input', calculate);
elChargeBasis.addEventListener('change', calculate);

document.getElementById('calculateBtn').addEventListener('click', calculate);

/* =========================================
   9. EXPORT / IMPORT
   ========================================= */

function downloadFile(content, fileName, mimeType) {
    const a = document.createElement("a");
    const file = new Blob([content], { type: mimeType });
    a.href = URL.createObjectURL(file);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

document.getElementById('btnExportJson').addEventListener('click', () => {
    const client = getActiveClient();
    const exportObj = {
        app_version: "ratrix_v3_client_backup",
        client_data: client
    };
    const json = JSON.stringify(exportObj, null, 2);
    const safeName = client.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    downloadFile(json, `backup_${safeName}.json`, "application/json");
});

document.getElementById('jsonUpload').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const imported = JSON.parse(e.target.result);
            if (imported.app_version === "ratrix_v3_client_backup" && imported.client_data) {
                const clientName = imported.client_data.name;
                if(confirm(`Restore data for "${clientName}"? This will create a NEW client entry.`)) {
                    const newId = generateId('c_');
                    imported.client_data.id = newId; 
                    GLOBAL_STORE.clients[newId] = imported.client_data;
                    switchClient(newId);
                    alert("Client restored successfully!");
                }
            } else {
                alert("Invalid or incompatible backup file.");
            }
        } catch (err) {
            console.error(err);
            alert("Error parsing JSON.");
        }
        event.target.value = '';
    };
    reader.readAsText(file);
});

// --- EXCEL LOGIC (Updated Headers with DASH) ---
document.getElementById('btnExportCsv').addEventListener('click', async () => {
    const model = elModel.value;
    const profile = getActiveProfile();
    const currentLimits = profile.limits;
    const client = getActiveClient();
    const serviceMode = elServiceMode.value;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(profile.name || 'Sheet1');

    // --- ENHANCED HEADER SECTION ---
    
    const rowClient = worksheet.addRow([`Client: ${client.name}`]);
    rowClient.font = { bold: true, size: 14 };
    rowClient.height = 20;

    const rowTable = worksheet.addRow([`Table: ${profile.name}`]);
    rowTable.font = { size: 12 };

    const rowService = worksheet.addRow([`Service Mode: ${serviceMode}`]);
    rowService.font = { size: 12 };

    worksheet.addRow([]); 

    // DATA HEADERS - CHANGED to "1-50" style
    const headers = ["Origin", "Destination"];
    currentLimits.forEach((lim, i) => {
        const prev = i === 0 ? 1 : currentLimits[i-1] + 1;
        headers.push(`${prev}-${lim}`); // Dash separator
    });
    
    const headerRow = worksheet.addRow(headers);
    headerRow.height = 20;
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }; 
    
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' } 
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // --- DATA ROWS ---
    profile.rows.forEach(row => {
        const rowData = [row.origin, row.dest];
        row.rates.forEach(r => rowData.push(r === null ? "" : r));
        worksheet.addRow(rowData);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const safeName = profile.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const clientSafe = client.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    downloadFile(blob, `${clientSafe}_${model}_${safeName}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
});

document.getElementById('csvUpload').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    if (file.name.endsWith('.xlsx')) {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const buffer = e.target.result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            processXLSXImport(workbook.getWorksheet(1));
            event.target.value = '';
        };
        reader.readAsArrayBuffer(file);
    }
});

function processXLSXImport(worksheet) {
    if (!worksheet || worksheet.rowCount < 2) { alert("Invalid Excel file."); return; }
    
    let headerRowIdx = -1;
    for(let i=1; i<=10; i++) {
        const row = worksheet.getRow(i);
        const cell1 = row.getCell(1).value ? row.getCell(1).value.toString().trim() : '';
        if(cell1 === 'Origin') {
            headerRowIdx = i;
            break;
        }
    }

    if (headerRowIdx === -1) { alert("Could not find header row (Origin...). Ensure format matches export."); return; }

    const headerRow = worksheet.getRow(headerRowIdx);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => headers[colNumber - 1] = cell.value ? cell.value.toString() : "");

    const hasServiceMode = headers[2] && headers[2].trim().toLowerCase() === 'servicemode';
    const rateStartIndex = hasServiceMode ? 3 : 2;

    const newLimits = [];
    const rateColIndices = [];

    for (let i = rateStartIndex; i < headers.length; i++) {
        const colName = headers[i].trim();
        // Handle "Rate_1_50" or "1_50" or "1-50"
        const normalized = colName.replace(/rate_/i, '').replace(/-/g, '_'); 
        const parts = normalized.split('_');
        
        const limitVal = parseFloat(parts[parts.length - 1]);
        if (!isNaN(limitVal)) {
            newLimits.push(limitVal);
            rateColIndices.push(i + 1); 
        }
    }

    if (newLimits.length === 0) { alert("No rate columns found (Format: X-Y or X_Y)."); return; }

    const profile = getActiveProfile();
    const limitsChanged = JSON.stringify(profile.limits) !== JSON.stringify(newLimits);
    if (limitsChanged) {
        if (!confirm("Weight columns differ. Update THIS table's structure?")) return;
        profile.limits = newLimits;
        profile.rows.forEach(r => {
            while (r.rates.length < newLimits.length) r.rates.push(null);
            if (r.rates.length > newLimits.length) r.rates = r.rates.slice(0, newLimits.length);
        });
    }

    const newRows = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowIdx) return; // Skip headers and title
        const rowOrigin = row.getCell(1).value ? row.getCell(1).value.toString().trim() : "";
        const rowDest = row.getCell(2).value ? row.getCell(2).value.toString().trim() : "";
        if (!rowOrigin && !rowDest) return;

        const rowRates = [];
        rateColIndices.forEach(colIdx => {
            const cellVal = row.getCell(colIdx).value;
            let val = (cellVal && typeof cellVal === 'object' && cellVal.result !== undefined) ? cellVal.result : cellVal;
            val = parseFloat(val);
            rowRates.push(isNaN(val) ? null : val);
        });

        newRows.push({ origin: rowOrigin, dest: rowDest, rates: rowRates });
    });

    profile.rows = newRows;
    saveToLocal();
    renderTableStructure();
    calculate();
    alert("Import Successful!");
}

// Theme
const themeToggleBtn = document.getElementById('theme-toggle');
const htmlElement = document.documentElement;
const currentTheme = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
htmlElement.setAttribute('data-theme', currentTheme);

if(themeToggleBtn) {
    themeToggleBtn.addEventListener('click', () => {
        let theme = htmlElement.getAttribute('data-theme');
        let newTheme = theme === 'light' ? 'dark' : 'light';
        htmlElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
    });
}

// START
initApp();