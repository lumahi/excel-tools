window.onload = function() {
    resetApplicationState(); // Start with a clean slate
    loadSettings();          // Then load any saved settings over the top
};

const vlookupBtn = document.getElementById('vlookupBtn');
const sourceFileInput = document.getElementById('fileInputSource');
const lookupFileInput = document.getElementById('fileInputLookup');
const downloadResultsBtn = document.getElementById('downloadResultsBtn');
const resultsColumn = document.getElementById('results-column');
const resetAppBtn = document.getElementById('resetAppBtn');
const saveSettingsBtn = document.getElementById('saveSettingsBtn');
const SETTINGS_KEY = 'vlookupToolSettings';

let enrichedResults = [], cachedSourceData = [], cachedLookupData = [];

function parseSingleColumn(input) {
    if (!input) return NaN;
    const part = input.trim().toUpperCase();
    if (!isNaN(parseInt(part))) {
        const colIndex = parseInt(part, 10) - 1;
        return colIndex >= 0 ? colIndex : NaN;
    }
    if (/^[A-Z]+$/.test(part)) {
        return part.split('').reduce((acc, char) => (acc * 26) + char.charCodeAt(0) - 64, 0) - 1;
    }
    return NaN;
}

function parseColumnRangeInput(input) {
    if (!input) return [];
    const indices = new Set();
    const parts = input.split(',');
    for (const part of parts) {
        const trimmedPart = part.trim();
        if (trimmedPart.includes('-')) {
            const [startStr, endStr] = trimmedPart.split('-');
            const start = parseSingleColumn(startStr);
            const end = parseSingleColumn(endStr);
            if (!isNaN(start) && !isNaN(end) && end >= start) {
                for (let i = start; i <= end; i++) indices.add(i);
            } else return [NaN];
        } else {
            const index = parseSingleColumn(trimmedPart);
            if (!isNaN(index)) indices.add(index);
            else return [NaN];
        }
    }
    return Array.from(indices).sort((a, b) => a - b);
}

async function handleFileUpload(file, cacheTarget, previewContainerId) {
    const container = document.getElementById(previewContainerId);
    // Hide results when a new file is selected to avoid showing stale data
    resultsColumn.style.display = 'none';

    if (!file) {
        if (cacheTarget === 'source') cachedSourceData = [];
        else cachedLookupData = [];
        container.innerHTML = '';
        return;
    }
    try {
        const data = await readFile(file);
        if (cacheTarget === 'source') cachedSourceData = data;
        else cachedLookupData = data;
        renderPreview(data, previewContainerId, { rows: 5, cols: 5, auto: false });
    } catch (err) {
        container.innerHTML = `<div class="alert alert-danger">Error reading file.</div>`;
    }
}

function colIndexToLetter(index) {
    let letter = '';
    while (index >= 0) {
        letter = String.fromCharCode(index % 26 + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}


function renderPreview(data, containerId, options) {
    const { rows, cols, auto } = options;
    const container = document.getElementById(containerId);
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="alert alert-light">File is empty.</div>';
        return;
    }
    
    const dataCache = containerId.startsWith('source') ? cachedSourceData : cachedLookupData;
    let numRows, numCols;

    if (auto) {
        numRows = dataCache.length;
        numCols = dataCache.reduce((max, row) => Math.max(max, (row || []).length), 0) || 1;
    } else {
        numRows = Math.max(1, rows);
        numCols = Math.max(1, cols);
    }
    
    let previewData = data.slice(0, numRows).map(row => {
        const newRow = (row || []).slice(0, numCols);
        while(newRow.length < numCols) newRow.push(undefined);
        return newRow;
    });

    while (previewData.length < numRows) {
        previewData.push(Array(numCols).fill(undefined));
    }

    const controlsHTML = `
        <div class="d-flex align-items-center gap-2 mt-3 preview-controls">
            <label for="${containerId}-rows" class="form-label mb-0 small">Preview:</label>
            <input type="number" class="form-control form-control-sm" id="${containerId}-rows" value="${auto ? numRows : rows}" min="1" ${auto ? 'disabled' : ''}>
            <span class="text-muted">x</span>
            <input type="number" class="form-control form-control-sm" id="${containerId}-cols" value="${auto ? numCols : cols}" min="1" ${auto ? 'disabled' : ''}>
            <div class="form-check form-check-inline ms-auto">
                <input class="form-check-input" type="checkbox" id="${containerId}-autoresize" ${auto ? 'checked' : ''}>
                <label class="form-check-label small" for="${containerId}-autoresize">Auto-size</label>
            </div>
        </div>
        <p class="text-muted small fst-italic mt-1">This is a cosmetic preview and does not affect the final VLOOKUP results.</p>`;
    
    // Column Headers (A, B, C...)
    let colHeaderHTML = `<tr><th class="sticky-top-left"></th>`;
    for (let i = 0; i < numCols; i++) colHeaderHTML += `<th class="sticky-col-header">${colIndexToLetter(i)}</th>`;
    colHeaderHTML += '</tr>';

    // Table Body (Row numbers + data)
    let tableBodyHTML = previewData.map((row, i) => `<tr><th class="sticky-row-header">${i + 1}</th>${(row || []).map(cell => `<td>${cell ?? ''}</td>`).join('')}</tr>`).join('');

    const tableHTML = `
        <div class="preview-container">
            <table class="table table-bordered table-sm">
                <thead class="table-light">${colHeaderHTML}</thead>
                <tbody>${tableBodyHTML}</tbody>
            </table>
        </div>`;

    container.innerHTML = controlsHTML + tableHTML;
    
    const rowsInput = document.getElementById(`${containerId}-rows`);
    const colsInput = document.getElementById(`${containerId}-cols`);
    const autoResizeCheckbox = document.getElementById(`${containerId}-autoresize`);

    rowsInput.addEventListener('change', (e) => renderPreview(dataCache, containerId, { rows: parseInt(e.target.value), cols: parseInt(colsInput.value), auto: false }));
    colsInput.addEventListener('change', (e) => renderPreview(dataCache, containerId, { rows: parseInt(rowsInput.value), cols: parseInt(e.target.value), auto: false }));
    autoResizeCheckbox.addEventListener('change', (e) => {
        if (e.target.checked) {
            renderPreview(dataCache, containerId, { rows: 5, cols: 5, auto: true });
        } else {
            renderPreview(dataCache, containerId, { rows: 5, cols: 5, auto: false });
        }
    });
}


function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                resolve(XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, blankrows: true, defval: null }));
            } catch (err) { reject(err); }
        };
        reader.onerror = err => reject(err);
        reader.readAsArrayBuffer(file);
    });
}

function cleanValue(value, options) {
    if (value === undefined || value === null) return value;
    let strValue = String(value);
    if (options.trim) strValue = strValue.trim();
    if (options.ignoreSpecial) strValue = strValue.replace(/[^a-zA-Z0-9]/g, '');
    if (!options.caseSensitive) strValue = strValue.toLowerCase();
    return strValue;
}

function handleLookup(sourceData, lookupData) {
    const lookupValueCol = parseSingleColumn(document.getElementById('lookupValueColumn').value);
    const lookupCol = parseSingleColumn(document.getElementById('lookupColumn').value);
    const returnCols = parseColumnRangeInput(document.getElementById('returnColumn').value);
    const cleaningOptions = {
        trim: document.getElementById('trimWhitespaceCheckbox').checked,
        ignoreSpecial: document.getElementById('ignoreSpecialCharsCheckbox').checked,
        caseSensitive: document.getElementById('caseSensitiveCheckbox').checked,
    };
    const outputOptions = {
        includeAllLookupCols: document.getElementById('includeAllLookupColsCheckbox').checked,
    };
    const matchFilterMode = document.getElementById('matchFilterMode').value;

    if (isNaN(lookupValueCol) || isNaN(lookupCol) || returnCols.some(isNaN)) {
        displayMessage('Please provide valid column identifiers for all fields. Check for errors in your column ranges.', 'danger');
        return;
    }

    const lookupMap = new Map();
    const sourceHeaders = sourceData[0] || [];
    for (const row of sourceData.slice(1)) {
         if (!row || row.length === 0) continue; 
        const key = cleanValue(row[lookupCol], cleaningOptions);
        if (key !== undefined && key !== null && key !== '') {
            if (!lookupMap.has(key)) lookupMap.set(key, row);
        }
    }
    
    const lookupHeaders = lookupData[0] || [];
    const returnHeaders = returnCols.map(c => sourceHeaders[c] || `Column ${String.fromCharCode(65 + c)}`);
    const finalHeaders = outputOptions.includeAllLookupCols
        ? [...lookupHeaders, ...returnHeaders]
        : [lookupHeaders[lookupValueCol] || `Column ${String.fromCharCode(65 + lookupValueCol)}`, ...returnHeaders];
    
    enrichedResults = [finalHeaders];
    let recordsProcessed = 0;
    for (const row of lookupData.slice(1)) {
        if (!row || row.length === 0) continue;
        recordsProcessed++;
        const originalValue = row[lookupValueCol];
        const cleanedValue = cleanValue(originalValue, cleaningOptions);
        const matchedRow = (cleanedValue !== undefined && cleanedValue !== null) ? lookupMap.get(cleanedValue) : undefined;

        if (matchedRow) {
            if (matchFilterMode === 'all' || matchFilterMode === 'found') {
                const baseRow = outputOptions.includeAllLookupCols ? [...row] : [originalValue];
                const foundValues = returnCols.map(c => matchedRow[c] ?? '');
                enrichedResults.push([...baseRow, ...foundValues]);
            }
        } else { // Not found
            if (matchFilterMode === 'all' || matchFilterMode === 'notfound') {
                const baseRow = outputOptions.includeAllLookupCols ? [...row] : [originalValue];
                enrichedResults.push([...baseRow, ...Array(returnCols.length).fill('Not Found')]);
            }
        }
    }
    
    displayResultsTable(enrichedResults, recordsProcessed);
}

vlookupBtn.addEventListener('click', async () => {
    if (!cachedSourceData.length || !cachedLookupData.length) {
        displayMessage('Please upload both a Source File and a Lookup File.', 'danger');
        return;
    }
    setLoading(true);
    try {
        handleLookup(cachedSourceData, cachedLookupData);
    } catch (err) {
        console.error(err);
        displayMessage('An error occurred while processing the files.', 'danger');
    } finally {
        setLoading(false);
    }
});

function displayResultsTable(data, processedCount) {
    const outputDiv = document.getElementById('output');
    const headers = data[0];
    const body = data.slice(1);
    const matchFilterMode = document.getElementById('matchFilterMode').value;
    let message;
    
    switch(matchFilterMode) {
        case 'found':
            message = `${body.length} of ${processedCount} records from the lookup file resulted in a match.`;
            break;
        case 'notfound':
            message = `${body.length} of ${processedCount} records from the lookup file could not be found in the source file.`;
            break;
        case 'all':
        default:
            message = `${processedCount} records from the lookup file were processed, showing all results.`;
            break;
    }
    
    outputDiv.innerHTML = `<p class="text-muted px-3 pt-3">${message}</p>
        <div class="table-responsive" style="max-height: 400px; overflow-y: auto;">
            <table class="table table-striped table-hover table-sm">
                <thead class="table-dark" style="position: sticky; top: 0;">
                    <tr>${headers.map(h => `<th>${h ?? ''}</th>`).join('')}</tr>
                </thead>
                <tbody>
                    ${body.map(row => `<tr>${(row || []).map(cell => `<td>${cell ?? ''}</td>`).join('')}</tr>`).join('')}
                </tbody>
            </table>
        </div>`;
    
    resultsColumn.style.display = 'block';
    downloadResultsBtn.style.display = data.length > 1 ? 'inline-block' : 'none';
}

function setLoading(isLoading) {
    document.getElementById('btn-text').style.display = isLoading ? 'none' : 'inline-block';
    document.getElementById('btn-spinner').style.display = isLoading ? 'inline-block' : 'none';
    vlookupBtn.disabled = isLoading;
}

function displayMessage(message, type) {
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = `<div class="alert alert-${type} m-3">${message}</div>`;
    resultsColumn.style.display = 'block';
    downloadResultsBtn.style.display = 'none';
}

function downloadCSV() {
    const csvRows = [];
    enrichedResults.forEach(rowArray => {
        let row = rowArray.map(cell => {
            const cellStr = (cell === null || cell === undefined) ? '' : String(cell);
            if (cellStr.search(/("|,|\n)/g) >= 0) {
                return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
        }).join(',');
        csvRows.push(row);
    });
    
    const csvString = csvRows.join('\r\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'enriched_results.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// --- NEW --- Save/Load/Reset Logic ---

function saveSettings() {
    const settings = {
        // We can't save the file object, but we can save the filename as a hint for the user.
        sourceFileName: sourceFileInput.files.length > 0 ? sourceFileInput.files[0].name : null,
        lookupFileName: lookupFileInput.files.length > 0 ? lookupFileInput.files[0].name : null,

        // Parameters
        lookupValueColumn: document.getElementById('lookupValueColumn').value,
        lookupColumn: document.getElementById('lookupColumn').value,
        returnColumn: document.getElementById('returnColumn').value,
        includeAllLookupCols: document.getElementById('includeAllLookupColsCheckbox').checked,
        matchFilterMode: document.getElementById('matchFilterMode').value,

        // Cleaning options
        trimWhitespace: document.getElementById('trimWhitespaceCheckbox').checked,
        ignoreSpecialChars: document.getElementById('ignoreSpecialCharsCheckbox').checked,
        caseSensitive: document.getElementById('caseSensitiveCheckbox').checked
    };

    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));

    // Provide user feedback by temporarily changing the button
    const originalText = saveSettingsBtn.innerHTML;
    saveSettingsBtn.innerHTML = `<i class="bi bi-check-lg"></i> Saved!`;
    saveSettingsBtn.classList.remove('btn-outline-success');
    saveSettingsBtn.classList.add('btn-success');
    saveSettingsBtn.disabled = true;

    setTimeout(() => {
        saveSettingsBtn.innerHTML = originalText;
        saveSettingsBtn.classList.remove('btn-success');
        saveSettingsBtn.classList.add('btn-outline-success');
        saveSettingsBtn.disabled = false;
    }, 2000);
}

function loadSettings() {
    const savedSettings = localStorage.getItem(SETTINGS_KEY);
    if (savedSettings) {
        const settings = JSON.parse(savedSettings);

        // Restore text and select inputs
        document.getElementById('lookupValueColumn').value = settings.lookupValueColumn || '';
        document.getElementById('lookupColumn').value = settings.lookupColumn || '';
        document.getElementById('returnColumn').value = settings.returnColumn || '';
        document.getElementById('matchFilterMode').value = settings.matchFilterMode || 'found';

        // Restore checkboxes (handle undefined for backward compatibility)
        document.getElementById('includeAllLookupColsCheckbox').checked = settings.includeAllLookupCols !== false;
        document.getElementById('trimWhitespaceCheckbox').checked = settings.trimWhitespace || false;
        document.getElementById('ignoreSpecialCharsCheckbox').checked = settings.ignoreSpecialChars || false;
        document.getElementById('caseSensitiveCheckbox').checked = settings.caseSensitive || false;

        // Display saved filenames as a hint
        if (settings.sourceFileName) {
            document.getElementById('savedSourceFilename').innerHTML = `<i class="bi bi-info-circle"></i> Last saved with: <strong>${settings.sourceFileName}</strong>`;
        }
        if (settings.lookupFileName) {
            document.getElementById('savedLookupFilename').innerHTML = `<i class="bi bi-info-circle"></i> Last saved with: <strong>${settings.lookupFileName}</strong>`;
        }
        console.log("Settings loaded from localStorage.");
    }
}


function resetApplicationState() {
    // 1. Clear the in-memory data arrays
    enrichedResults = [];
    cachedSourceData = [];
    cachedLookupData = [];

    // 2. Reset the main form to its default values
    // This clears the current values. A page refresh will reload them from localStorage if they exist.
    document.getElementById('main-form').reset();

    // 3. Clear the file input elements themselves
    sourceFileInput.value = '';
    lookupFileInput.value = '';

    // 4. Clear the preview and result areas from the DOM
    document.getElementById('sourcePreview').innerHTML = '';
    document.getElementById('lookupPreview').innerHTML = '';
    document.getElementById('output').innerHTML = '';
    resultsColumn.style.display = 'none';
    
    // Note: Saved settings in localStorage and their UI hints are intentionally NOT cleared.
    console.log("Current session has been reset. Saved settings are preserved in local storage.");
}


sourceFileInput.addEventListener('change', e => handleFileUpload(e.target.files[0], 'source', 'sourcePreview'));
lookupFileInput.addEventListener('change', e => handleFileUpload(e.target.files[0], 'lookup', 'lookupPreview'));
downloadResultsBtn.addEventListener('click', downloadCSV);
resetAppBtn.addEventListener('click', resetApplicationState);
saveSettingsBtn.addEventListener('click', saveSettings);


// --- SAMPLE FILE GENERATION ---

function generateSampleData(rowCount, isPerfect) {
    const sourceData = [["ID", "Product Name", "Category", "Unit Price", "Stock Level", "Supplier", "Date Added"]];
    const lookupData = [["ProductID", "Quantity Needed", "Urgency"]];
    
    const categories = ["Electronics", "Accessories", "Software", "Peripherals", "Components"];
    const suppliers = ["TechCorp", "Innovate Inc.", "Gadgetron", "Supply Co", "Global Parts"];
    const urgencies = ["High", "Medium", "Low"];
    const p_names = ["Quantum", "Aero", "Hyper", "Stellar", "Nova", "Omega", "Alpha", "Delta", "Orion", "Pulsar"];
    const p_models = ["Pro", "Max", "XT", "Core", "Mini", "Plus", "HD", "4K", "Stealth", "Prime"];

    let sourceIds = [];
    for (let i = 1; i <= rowCount; i++) {
        const id = `PROD-${String(i).padStart(3, '0')}`;
        sourceIds.push(id);

        const name = `${p_names[i % p_names.length]} ${p_models[i % p_models.length]} ${i}`;
        const category = categories[i % categories.length];
        const price = (Math.random() * 1490 + 10).toFixed(2);
        const stock = Math.floor(Math.random() * 500);
        const supplier = suppliers[i % suppliers.length];
        const date = new Date(Date.now() - Math.random() * 31536000000).toISOString().split('T')[0]; // Within last year
        
        sourceData.push([id, name, category, price, stock, supplier, date]);
    }

    if (!isPerfect) {
        sourceData.splice(7, 0, []); // Add a blank row
        sourceData[11][5] = null; // Add a blank cell
    }
    
    let lookupIds = [];
    if (isPerfect) {
        lookupIds = [...sourceIds].sort(() => 0.5 - Math.random()); // Shuffle for realism
    } else {
        lookupIds = [...sourceIds].sort(() => 0.5 - Math.random()).slice(0, rowCount - 4); // Take most
        lookupIds.push("PROD-888", "PROD-999"); // Add some non-existent ones
        lookupIds = lookupIds.sort(() => 0.5 - Math.random()); // Re-shuffle
    }

    // Ensure lookupData also has `rowCount` rows
    for(let i = 0; i < rowCount; i++) {
        const id = lookupIds[i] || `PROD-${String(800+i).padStart(3, '0')}`; // Add more non-existent if needed
        const quantity = Math.floor(Math.random() * 50) + 1;
        const urgency = urgencies[i % urgencies.length];
        lookupData.push([id, quantity, urgency]);
    }
    
    return { sourceData, lookupData };
}

function downloadSample(data, fileName) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SampleData");
    XLSX.writeFile(wb, fileName);
}

document.getElementById('downloadSourceStandardBtn').addEventListener('click', () => {
     const { sourceData } = generateSampleData(20, false);
     downloadSample(sourceData, 'sample_source_standard.xlsx');
});

document.getElementById('downloadLookupStandardBtn').addEventListener('click', () => {
     const { lookupData } = generateSampleData(20, false);
     downloadSample(lookupData, 'sample_lookup_standard.xlsx');
});

document.getElementById('downloadSourcePerfectBtn').addEventListener('click', () => {
     const { sourceData } = generateSampleData(20, true);
     downloadSample(sourceData, 'sample_source_perfect.xlsx');
});

document.getElementById('downloadLookupPerfectBtn').addEventListener('click', () => {
     const { lookupData } = generateSampleData(20, true);
     downloadSample(lookupData, 'sample_lookup_perfect.xlsx');
});