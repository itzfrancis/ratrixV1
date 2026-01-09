/* =========================================
   1. DATA CONFIGURATION & STORAGE
   ========================================= */

const STORAGE_KEY = 'ratrix_data_v2_profiles';

// EDITABLE BRACKETS: Shared columns (Weight Limits) across the entire app
let BRACKET_LIMITS = [50, 100, 150, 500];

// Available Pricing Models
const MODEL_KEYS = [
    'fixed', 'flat', 'minFixed', 'cumulative', 
    'minCumulative', 'minExcess', 'excess'
];

/* DATA STRUCTURE:
   DATA_STORE = {
       "fixed": {
           activeId: "timestamp_id_1",
           profiles: {
               "timestamp_id_1": { name: "Default Table", rows: [...] },
               "timestamp_id_2": { name: "VIP Rates", rows: [...] }
           }
       },
       "flat": { ... }
   }
*/
let DATA_STORE = {};

// --- HELPERS ---

function generateId() {
    return 'p_' + Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function createEmptyRow() {
    return {
        origin: "",
        dest: "",
        rates: new Array(BRACKET_LIMITS.length).fill(null) 
    };
}

function createDefaultProfile() {
    return {
        name: "Standard Table",
        rows: [createEmptyRow()]
    };
}

// --- PERSISTENCE ---

function saveToLocal() {
    const payload = {
        limits: BRACKET_LIMITS,
        store: DATA_STORE
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function loadFromLocal() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
        try {
            const saved = JSON.parse(raw);
            if (saved.limits && saved.store) {
                BRACKET_LIMITS = saved.limits;
                if (saved.store['fixed'] && saved.store['fixed'].profiles) {
                    DATA_STORE = saved.store;
                    return true;
                }
            }
        } catch (e) {
            console.error("Storage corrupted", e);
        }
    }
    return false;
}

// --- INITIALIZATION ---

function initData() {
    if (!loadFromLocal()) {
        MODEL_KEYS.forEach(key => {
            const pid = generateId();
            DATA_STORE[key] = {
                activeId: pid,
                profiles: {}
            };
            DATA_STORE[key].profiles[pid] = createDefaultProfile();
        });
        saveToLocal();
    }
}

initData();

/* =========================================
   2. PROFILE MANAGEMENT
   ========================================= */

const elTableSelect = document.getElementById('tableProfileSelect');
const elModel = document.getElementById('pricingModel');

function getActiveModelData() {
    const model = elModel.value;
    if (!DATA_STORE[model]) {
        const pid = generateId();
        DATA_STORE[model] = { activeId: pid, profiles: {} };
        DATA_STORE[model].profiles[pid] = createDefaultProfile();
    }
    return DATA_STORE[model];
}

function getActiveProfile() {
    const modelData = getActiveModelData();
    const pid = modelData.activeId;
    if (!modelData.profiles[pid]) {
        const firstKey = Object.keys(modelData.profiles)[0];
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

document.getElementById('tableProfileSelect').addEventListener('change', (e) => {
    const modelData = getActiveModelData();
    modelData.activeId = e.target.value;
    saveToLocal();
    renderTableStructure();
    calculate();
});

document.getElementById('btnNewProfile').addEventListener('click', () => {
    const name = prompt("Enter a name for the new table (e.g., 'Client B Rates'):");
    if (!name) return;

    const modelData = getActiveModelData();
    const newId = generateId();
    
    modelData.profiles[newId] = {
        name: name,
        rows: [createEmptyRow()]
    };
    modelData.activeId = newId; 
    
    saveToLocal();
    renderProfileDropdown();
    renderTableStructure();
    calculate();
});

document.getElementById('btnRenameProfile').addEventListener('click', () => {
    const profile = getActiveProfile();
    const newName = prompt("Rename current table:", profile.name);
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
    if (confirm(`Are you sure you want to delete table "${profileName}"?`)) {
        delete modelData.profiles[modelData.activeId];
        modelData.activeId = Object.keys(modelData.profiles)[0];
        
        saveToLocal();
        renderProfileDropdown();
        renderTableStructure();
        calculate();
    }
});


/* =========================================
   3. CALCULATION FORMULAS
   ========================================= */

const strategies = {
    fixed: (w, rates) => {
        const index = getBracketIndex(w);
        if (index === -1 || !isValid(rates[index])) return null; 
        return w * rates[index];
    },
    flat: (w, rates) => {
        const index = getBracketIndex(w);
        if (index === -1 || !isValid(rates[index])) return null;
        return rates[index];
    },
    minFixed: (w, rates) => {
        // For MinFixed, we check against the raw limit first for the minimum
        if (w <= BRACKET_LIMITS[0]) return isValid(rates[0]) ? rates[0] : null;
        
        const index = getBracketIndex(w);
        if (index === -1 || !isValid(rates[index])) return null;
        return w * rates[index];
    },
    cumulative: (w, rates) => {
        let total = 0;
        let remaining = w;
        let prevMax = 0;
        for (let i = 0; i < BRACKET_LIMITS.length; i++) {
            if (!isValid(rates[i])) return null; 
            const capacity = BRACKET_LIMITS[i] - prevMax;
            const fill = Math.min(remaining, capacity);
            if (fill > 0) {
                total += fill * rates[i];
                remaining -= fill;
            }
            prevMax = BRACKET_LIMITS[i];
            if (remaining <= 0) break;
        }
        return remaining > 0 ? null : total; 
    },
    minCumulative: (w, rates) => {
        let total = 0;
        let remaining = w;
        let prevMax = 0;
        for (let i = 0; i < BRACKET_LIMITS.length; i++) {
            if (!isValid(rates[i])) return null;
            const capacity = BRACKET_LIMITS[i] - prevMax;
            const fill = Math.min(remaining, capacity);
            if (fill > 0) {
                if (i === 0) total += rates[0]; 
                else total += fill * rates[i];
                remaining -= fill;
            }
            prevMax = BRACKET_LIMITS[i];
            if (remaining <= 0) break;
        }
        return remaining > 0 ? null : total;
    },
    minExcess: (w, rates) => {
        if (!isValid(rates[0]) || !isValid(rates[1])) return null;
        const limit = BRACKET_LIMITS[0]; 
        const baseFlat = rates[0];      
        const excessRate = rates[1];     
        if (w <= limit) return baseFlat;
        return baseFlat + ((w - limit) * excessRate);
    },
    excess: (w, rates) => {
        if (!isValid(rates[0]) || !isValid(rates[1])) return null;
        const limit = BRACKET_LIMITS[0]; 
        const baseRate = rates[0];       
        const excessRate = rates[1];     
        if (w <= limit) return w * baseRate;
        return (limit * baseRate) + ((w - limit) * excessRate);
    }
};

function isValid(val) { return val !== null && val !== "" && !isNaN(val); }

// UPDATED FUNCTION: Rounds down weight for lookup
function getBracketIndex(w) {
    const effectiveWeight = Math.floor(w); // 5.1 -> 5, 5.9 -> 5
    for (let i = 0; i < BRACKET_LIMITS.length; i++) {
        if (effectiveWeight <= BRACKET_LIMITS[i]) return i;
    }
    return -1;
}

/* =========================================
   4. UI LOGIC & RENDERERS
   ========================================= */

const elTableHeader = document.getElementById('tableHeaderRow');
const elTableBody = document.getElementById('rateTableBody');
const elResult = document.getElementById('resultPrice');
const elDesc = document.getElementById('formulaDesc');
const elWeight = document.getElementById('weightInput');
const elOrigin = document.getElementById('origin');
const elDest = document.getElementById('destination');

// DIMENSION ELEMENTS
const elDimL = document.getElementById('dimL');
const elDimW = document.getElementById('dimW');
const elDimH = document.getElementById('dimH');
const elVolDivisor = document.getElementById('volDivisor');
const elChargeBasis = document.getElementById('chargeBasis');

// STAT DISPLAY ELEMENTS
const elDispActual = document.getElementById('dispActual');
const elDispVol = document.getElementById('dispVol');
const elDispCbm = document.getElementById('dispCbm');

function renderTableStructure() {
    renderProfileDropdown();
    const profile = getActiveProfile();
    const activeRows = profile.rows;

    let headerHtml = `
        <th style="width: 100px;">Origin</th>
        <th style="width: 100px;">Destination</th>
    `;

    BRACKET_LIMITS.forEach((limit, index) => {
        const prevLimit = index === 0 ? 0 : BRACKET_LIMITS[index - 1];
        const startVal = prevLimit + 1;
        headerHtml += `
        <th>
            <span id="start-${index}" class="header-range-text">${startVal}</span> - 
            <input type="number" class="header-input" 
                   value="${limit}" 
                   onchange="handleEditLimit(${index}, this.value)">
        </th>`;
    });
    headerHtml += `<th style="width: 60px;">Action</th>`;
    elTableHeader.innerHTML = headerHtml;

    let bodyHtml = '';
    activeRows.forEach((row, rowIndex) => {
        let ratesHtml = '';
        while(row.rates.length < BRACKET_LIMITS.length) row.rates.push(null);

        row.rates.forEach((rate, colIndex) => {
            if(colIndex < BRACKET_LIMITS.length) {
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
    const lastLimit = BRACKET_LIMITS[BRACKET_LIMITS.length - 1] || 0;
    BRACKET_LIMITS.push(lastLimit + 50);
    
    MODEL_KEYS.forEach(key => {
        if(DATA_STORE[key] && DATA_STORE[key].profiles) {
            Object.values(DATA_STORE[key].profiles).forEach(prof => {
                prof.rows.forEach(r => r.rates.push(null));
            });
        }
    });
    
    saveToLocal();
    renderTableStructure();
    calculate();
});

document.getElementById('addRowBtn').addEventListener('click', () => {
    const profile = getActiveProfile();
    profile.rows.push(createEmptyRow());
    saveToLocal();
    renderTableStructure();
});

// ---- EDIT HANDLERS ----

window.handleEditLimit = function(index, value) {
    const val = parseFloat(value);
    if (isNaN(val) || val <= 0) return;
    BRACKET_LIMITS[index] = val;
    saveToLocal();
    renderTableStructure(); 
    calculate();
};

window.handleEditRate = function(rowIndex, colIndex, value) {
    const profile = getActiveProfile();
    if (value === "") {
        profile.rows[rowIndex].rates[colIndex] = null;
    } else {
        profile.rows[rowIndex].rates[colIndex] = parseFloat(value);
    }
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

/* =========================================
   5. CALCULATION LOGIC
   ========================================= */

function calculate() {
    const actualWeight = parseFloat(elWeight.value) || 0;
    const origin = elOrigin.value;
    const dest = elDest.value;
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
    const routeIndex = profile.rows.findIndex(r => r.origin === origin && r.dest === dest);
    
    if (routeIndex === -1) {
        elResult.textContent = "Route Not Found";
        elDesc.textContent = `No configured rate for this route in table: ${profile.name}`;
        highlightRow(-1);
        return;
    }
    highlightRow(routeIndex);
    const route = profile.rows[routeIndex];

    const strategyFn = strategies[model] || strategies.fixed;
    const total = strategyFn(chargeableWeight, route.rates);

    if (total === null) {
        const maxLimit = BRACKET_LIMITS[BRACKET_LIMITS.length - 1];
        if (chargeableWeight > maxLimit) {
             elResult.textContent = "Over Limit";
             elDesc.textContent = `Weight (${chargeableWeight}kg) exceeds ${maxLimit}kg limit`;
        } else {
             elResult.textContent = "Invalid Rate";
             elDesc.textContent = "Rate is blank/missing for this bracket.";
        }
    } else {
        elResult.textContent = `Php ${total.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
        
        const bracketIdx = getBracketIndex(chargeableWeight);
        let bracketLabel = "Unknown";
        if (bracketIdx !== -1) {
            const min = bracketIdx === 0 ? 1 : BRACKET_LIMITS[bracketIdx - 1] + 1;
            const max = BRACKET_LIMITS[bracketIdx];
            bracketLabel = `${min}-${max}kg`;
        }
        
        const basisLabel = chargeBasis === 'volumetric' ? 'Vol. Wt.' : 'Act. Wt.';
        elDesc.textContent = `${model} | ${basisLabel} (${chargeableWeight}kg) | [Bracket: ${bracketLabel}]`;
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

elDimL.addEventListener('input', calculate);
elDimW.addEventListener('input', calculate);
elDimH.addEventListener('input', calculate);
elVolDivisor.addEventListener('input', calculate);
elChargeBasis.addEventListener('change', calculate);

document.getElementById('calculateBtn').addEventListener('click', calculate);

/* =========================================
   6. EXPORT / IMPORT
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

document.getElementById('btnExportCsv').addEventListener('click', () => {
    const model = elModel.value;
    const profile = getActiveProfile();
    
    let headers = "Origin,Destination";
    BRACKET_LIMITS.forEach((lim, i) => {
        const prev = i === 0 ? 1 : BRACKET_LIMITS[i-1] + 1;
        headers += `,Rate_${prev}_${lim}`;
    });
    headers += "\n";
    let csv = headers;
    profile.rows.forEach(row => {
        const safeRates = row.rates.map(r => r === null ? "" : r);
        csv += `${row.origin},${row.dest},${safeRates.join(',')}\n`;
    });
    
    const safeName = profile.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    downloadFile(csv, `${model}_${safeName}.csv`, "text/csv");
});

document.getElementById('btnExportJson').addEventListener('click', () => {
    const exportObj = {
        app_version: "ratrix_v2_profiles",
        limits: BRACKET_LIMITS,
        data_store: DATA_STORE
    };
    const json = JSON.stringify(exportObj, null, 2);
    downloadFile(json, "full_matrix_db.json", "application/json");
});

document.getElementById('jsonUpload').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const imported = JSON.parse(e.target.result);
            if (imported.app_version === "ratrix_v2_profiles") {
                BRACKET_LIMITS = imported.limits || BRACKET_LIMITS;
                DATA_STORE = imported.data_store || DATA_STORE;
            } 
            else {
                alert("Older format detected. Please check column limits.");
                if (imported.limits) BRACKET_LIMITS = imported.limits;
            }

            saveToLocal();
            renderProfileDropdown();
            renderTableStructure();
            calculate();
            alert("Matrix data loaded successfully!");
        } catch (err) {
            console.error(err);
            alert("Error parsing JSON.");
        }
        event.target.value = '';
    };
    reader.readAsText(file);
});

// Theme
const themeToggleBtn = document.getElementById('theme-toggle');
const htmlElement = document.documentElement;
const currentTheme = localStorage.getItem('theme') || 
                     (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
htmlElement.setAttribute('data-theme', currentTheme);

if(themeToggleBtn) {
    themeToggleBtn.addEventListener('click', () => {
        let theme = htmlElement.getAttribute('data-theme');
        let newTheme = theme === 'light' ? 'dark' : 'light';
        htmlElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
    });
}

// Init
renderTableStructure();