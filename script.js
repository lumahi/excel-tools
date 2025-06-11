window.onload = function() {
    document.getElementById('main-form').reset();
    document.getElementById('fileInputSource').value = '';
    document.getElementById('fileInputLookup').value = '';
    document.getElementById('sourcePreview').innerHTML = '';
    document.getElementById('lookupPreview').innerHTML = '';
};

const vlookupBtn = document.getElementById('vlookupBtn');
const sourceFileInput = document.getElementById('fileInputSource');
const lookupFileInput = document.getElementById('fileInputLookup');
const downloadResultsBtn = document.getElementById('downloadResultsBtn');

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
        renderPreview(data, previewContainerId, 5, 5);
    } catch (err) {
        container.innerHTML = `<div class="alert alert-danger">Error reading file.</div>`;
    }
}

function renderPreview(data, containerId, rows, cols) {
    const container = document.getElementById(containerId);
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="alert alert-light">File is empty.</div>';
        return;
    }
    
    const numRows = Math.max(1, rows);
    const numCols = Math.max(1, cols);
    
    let previewData = data.slice(0, numRows + 1).map(row => {
        const newRow = (row || []).slice(0, numCols);
        while(newRow.length < numCols) {
            newRow.push(undefined);
        }
        return newRow;
    });

    const headers = previewData[0] || [];
    const body = previewData.slice(1);

    const controlsHTML = `
        <div class="d-flex align-items-center gap-2 mt-3 preview-controls">
            <label for="${containerId}-rows" class="form-label mb-0 small">Preview:</label>
            <input type="number" class="form-control form-control-sm" id="${containerId}-rows" value="${rows}" min="1">
            <span class="text-muted">x</span>
            <input type="number" class="form-control form-control-sm" id="${containerId}-cols" value="${cols}" min="1">
        </div>
        <p class="text-muted small fst-italic mt-1">This is a cosmetic preview and does not affect the final VLOOKUP results.</p>`;
    
    const tableHTML = `
        <div class="preview-container">
            <table class="table table-bordered table-sm">
                <thead class="table-light"><tr>${headers.map(h => `<th>${h ?? ''}</th>`).join('')}</tr></thead>
                <tbody>${body.map(row => `<tr>${(row || []).map(cell => `<td>${cell ?? ''}</td>`).join('')}</tr>`).join('')}</tbody>
            </table>
        </div>`;

    container.innerHTML = controlsHTML + tableHTML;
    
    const dataCache = containerId.startsWith('source') ? cachedSourceData : cachedLookupData;
    document.getElementById(`${containerId}-rows`).addEventListener('change', (e) => renderPreview(dataCache, containerId, parseInt(e.target.value), parseInt(document.getElementById(`${containerId}-cols`).value)));
    document.getElementById(`${containerId}-cols`).addEventListener('change', (e) => renderPreview(dataCache, containerId, parseInt(document.getElementById(`${containerId}-rows`).value), parseInt(e.target.value)));
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
    const resultCard = document.getElementById('result-card');
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
    
    outputDiv.innerHTML = `<p class="text-muted">${message}</p>
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
    
    resultCard.style.display = 'block';
    downloadResultsBtn.style.display = data.length > 1 ? 'inline-block' : 'none';
}

function setLoading(isLoading) {
    document.getElementById('btn-text').style.display = isLoading ? 'none' : 'inline-block';
    document.getElementById('btn-spinner').style.display = isLoading ? 'inline-block' : 'none';
    vlookupBtn.disabled = isLoading;
}

function displayMessage(message, type) {
    const resultCard = document.getElementById('result-card');
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
    resultCard.style.display = 'block';
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

sourceFileInput.addEventListener('change', e => handleFileUpload(e.target.files[0], 'source', 'sourcePreview'));
lookupFileInput.addEventListener('change', e => handleFileUpload(e.target.files[0], 'lookup', 'lookupPreview'));
downloadResultsBtn.addEventListener('click', downloadCSV);

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