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

// Service Modes
const SERVICE_MODES = [
    "DOOR TO DOOR",
    "PORT TO PORT",
    "DOOR TO PORT",
    "PORT TO DOOR"
];

/* DATA STRUCTURE:
   DATA_STORE = {
       "fixed": {
           activeId: "timestamp_id_1",
           profiles: {
               "timestamp_id_1": { name: "Default Table", rows: [...] }
           }
       },
       ...
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
        serviceMode: "DOOR TO DOOR", // Default
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

function getBracketIndex(w) {
    const effectiveWeight = Math.floor(w);
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
    renderProfileDropdown();
    const profile = getActiveProfile();
    const activeRows = profile.rows;

    // Build Header
    let headerHtml = `
        <th style="width: 100px;">Origin</th>
        <th style="width: 100px;">Destination</th>
        <th style="width: 140px;">Service Mode</th> `;

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

    // Build Body
    let bodyHtml = '';
    activeRows.forEach((row, rowIndex) => {
        let ratesHtml = '';
        while(row.rates.length < BRACKET_LIMITS.length) row.rates.push(null);

        const currentService = row.serviceMode || SERVICE_MODES[0];

        // Build Service Dropdown for this specific row
        const serviceOptions = SERVICE_MODES.map(mode => 
            `<option value="${mode}" ${mode === currentService ? 'selected' : ''}>${mode}</option>`
        ).join('');

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
            <td>
                <select class="editable-input" style="width: 100%; text-align: left;"
                        onchange="handleEditServiceMode(${rowIndex}, this.value)">
                    ${serviceOptions}
                </select>
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

window.handleEditServiceMode = function(rowIndex, value) {
    const profile = getActiveProfile();
    profile.rows[rowIndex].serviceMode = value;
    saveToLocal();
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
    
    // Find row matching Origin, Dest AND ServiceMode
    const routeIndex = profile.rows.findIndex(r => 
        r.origin === origin && 
        r.dest === dest && 
        (r.serviceMode || "DOOR TO DOOR") === serviceMode
    );
    
    if (routeIndex === -1) {
        elResult.textContent = "Route Not Found";
        elDesc.textContent = `No match for ${origin} -> ${dest} via ${serviceMode}`;
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
   6. EXPORT / IMPORT (XLSX & JSON)
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

// --- EXPORT TO EXCEL WITH DROPDOWN ---
document.getElementById('btnExportCsv').addEventListener('click', async () => {
    const model = elModel.value;
    const profile = getActiveProfile();
    
    // Create new Workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(profile.name || 'Sheet1');

    // Headers
    const headers = ["Origin", "Destination", "ServiceMode"];
    BRACKET_LIMITS.forEach((lim, i) => {
        const prev = i === 0 ? 1 : BRACKET_LIMITS[i-1] + 1;
        headers.push(`Rate_${prev}_${lim}`);
    });

    worksheet.addRow(headers);
    worksheet.getRow(1).font = { bold: true };

    // Rows
    profile.rows.forEach(row => {
        const rowData = [
            row.origin, 
            row.dest, 
            row.serviceMode || "DOOR TO DOOR"
        ];
        row.rates.forEach(r => {
            rowData.push(r === null ? "" : r);
        });
        worksheet.addRow(rowData);
    });

    // ADD DROPDOWN TO SERVICE MODE (Column 3)
    const serviceModeColIndex = 3; 
    const rowCount = worksheet.rowCount;
    // Comma-separated string for list validation
    const dropdownList = `"${SERVICE_MODES.join(',')}"`;

    for (let i = 2; i <= rowCount; i++) {
        const cell = worksheet.getCell(i, serviceModeColIndex);
        cell.dataValidation = {
            type: 'list',
            allowBlank: false,
            formulae: [dropdownList],
            showErrorMessage: true,
            errorTitle: 'Invalid Service Mode',
            error: 'Select a valid Service Mode from the list.'
        };
    }

    // Write & Download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    
    const safeName = profile.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const fileName = `${model}_${safeName}.xlsx`;

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});

// --- IMPORT (CSV OR XLSX) ---
document.getElementById('csvUpload').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    // Check extension
    if (file.name.endsWith('.xlsx')) {
        // Handle Excel Import
        const reader = new FileReader();
        reader.onload = async function(e) {
            const buffer = e.target.result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1); // Read first sheet
            processXLSXImport(worksheet);
            event.target.value = ''; // Reset
        };
        reader.readAsArrayBuffer(file);
    } else {
        // Handle CSV Import (Legacy)
        const reader = new FileReader();
        reader.onload = function(e) {
            const text = e.target.result;
            processCSVImport(text);
            event.target.value = '';
        };
        reader.readAsText(file);
    }
});

function processXLSXImport(worksheet) {
    if (!worksheet || worksheet.rowCount < 2) {
        alert("Invalid Excel file.");
        return;
    }

    // 1. Parse Headers (Row 1)
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        headers[colNumber - 1] = cell.value ? cell.value.toString() : ""; 
    });

    const hasServiceMode = headers[2].trim().toLowerCase() === 'servicemode';
    const rateStartIndex = hasServiceMode ? 3 : 2;

    const newLimits = [];
    const rateColIndices = []; // 1-based index for ExcelJS getCell

    for (let i = rateStartIndex; i < headers.length; i++) {
        const colName = headers[i].trim();
        const parts = colName.split('_');
        const limitVal = parseFloat(parts[parts.length - 1]);
        if (!isNaN(limitVal)) {
            newLimits.push(limitVal);
            rateColIndices.push(i + 1); // ExcelJS uses 1-based columns
        }
    }

    if (newLimits.length === 0) {
        alert("No rate columns found (Format: Rate_X_Y).");
        return;
    }

    // 2. Check Limits
    const limitsChanged = JSON.stringify(BRACKET_LIMITS) !== JSON.stringify(newLimits);
    if (limitsChanged) {
        const confirmMsg = "Warning: The weight columns in the Excel file are different from the current setup.\n\n" +
                           "Importing this will update the weight brackets for ALL tables in the app.\n\n" +
                           "Proceed?";
        if (!confirm(confirmMsg)) return;
        BRACKET_LIMITS = newLimits;
        // Update all profiles to match new width
        MODEL_KEYS.forEach(key => {
            if (DATA_STORE[key] && DATA_STORE[key].profiles) {
                Object.values(DATA_STORE[key].profiles).forEach(p => {
                    p.rows.forEach(r => {
                        while (r.rates.length < newLimits.length) r.rates.push(null);
                        if (r.rates.length > newLimits.length) r.rates = r.rates.slice(0, newLimits.length);
                    });
                });
            }
        });
    }

    // 3. Parse Data Rows
    const profile = getActiveProfile();
    const newRows = [];

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header

        const rowOrigin = row.getCell(1).value ? row.getCell(1).value.toString().trim() : "";
        const rowDest = row.getCell(2).value ? row.getCell(2).value.toString().trim() : "";
        
        if (!rowOrigin && !rowDest) return;

        let rowService = "DOOR TO DOOR";
        if (hasServiceMode) {
            const rawService = row.getCell(3).value ? row.getCell(3).value.toString().trim().toUpperCase() : "";
            if(SERVICE_MODES.includes(rawService)) {
                rowService = rawService;
            }
        }

        const rowRates = [];
        rateColIndices.forEach(colIdx => {
            const cellVal = row.getCell(colIdx).value;
            // Handle Excel richness (if cell is object) or simple value
            let val = (cellVal && typeof cellVal === 'object' && cellVal.result !== undefined) ? cellVal.result : cellVal;
            val = parseFloat(val);
            rowRates.push(isNaN(val) ? null : val);
        });

        newRows.push({
            origin: rowOrigin,
            dest: rowDest,
            serviceMode: rowService,
            rates: rowRates
        });
    });

    profile.rows = newRows;
    saveToLocal();
    renderTableStructure();
    calculate();
    alert("Excel Imported Successfully! Rates and Service Modes updated.");
}

function processCSVImport(csvText) {
    const lines = csvText.trim().split(/\r?\n/);
    if (lines.length < 2) {
        alert("Invalid CSV: Not enough data.");
        return;
    }

    const headers = lines[0].split(',');
    const hasServiceMode = headers[2].trim().toLowerCase() === 'servicemode';
    const rateStartIndex = hasServiceMode ? 3 : 2;

    const newLimits = [];
    const rateColIndices = []; 

    for (let i = rateStartIndex; i < headers.length; i++) {
        const colName = headers[i].trim();
        const parts = colName.split('_');
        const limitVal = parseFloat(parts[parts.length - 1]);

        if (!isNaN(limitVal)) {
            newLimits.push(limitVal);
            rateColIndices.push(i);
        }
    }

    if (newLimits.length === 0) {
        alert("No rate columns found (Format: Rate_X_Y).");
        return;
    }

    const limitsChanged = JSON.stringify(BRACKET_LIMITS) !== JSON.stringify(newLimits);
    if (limitsChanged) {
        const confirmMsg = "Warning: The weight columns in the CSV are different from the current setup.\n\n" +
                           "Importing this will update the weight brackets for ALL tables in the app.\n\n" +
                           "Proceed?";
        if (!confirm(confirmMsg)) return;
        BRACKET_LIMITS = newLimits;
        MODEL_KEYS.forEach(key => {
            if (DATA_STORE[key] && DATA_STORE[key].profiles) {
                Object.values(DATA_STORE[key].profiles).forEach(p => {
                    p.rows.forEach(r => {
                        while (r.rates.length < newLimits.length) r.rates.push(null);
                        if (r.rates.length > newLimits.length) r.rates = r.rates.slice(0, newLimits.length);
                    });
                });
            }
        });
    }

    const profile = getActiveProfile();
    const newRows = [];

    for (let i = 1; i < lines.length; i++) {
        const cells = lines[i].split(',');
        if (cells.length < 2) continue;

        const rowOrigin = cells[0].trim();
        const rowDest = cells[1].trim();
        
        let rowService = "DOOR TO DOOR";
        if (hasServiceMode) {
            const rawService = cells[2].trim().toUpperCase();
            if(SERVICE_MODES.includes(rawService)) {
                rowService = rawService;
            }
        }
        
        if (!rowOrigin && !rowDest) continue;

        const rowRates = [];
        rateColIndices.forEach(colIndex => {
            const valStr = cells[colIndex] ? cells[colIndex].trim() : "";
            const val = parseFloat(valStr);
            rowRates.push(isNaN(val) ? null : val);
        });

        newRows.push({
            origin: rowOrigin,
            dest: rowDest,
            serviceMode: rowService,
            rates: rowRates
        });
    }

    profile.rows = newRows;
    saveToLocal();
    renderTableStructure();
    calculate();
    alert("CSV Imported Successfully!");
}

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