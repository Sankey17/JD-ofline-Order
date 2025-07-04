// Flexible Document Generator Script
let allData = [];
let selectedColumns = [];
let currentDataSource = 'file'; // 'file' or 'sheets'

// Format settings with defaults
let formatSettings = {
    paperWidth: 21.0,
    paperHeight: 29.7,
    marginTop: 2.0,
    marginBottom: 2.0,
    marginLeft: 2.0,
    marginRight: 2.0,
    fontFamily: 'Arial, sans-serif',
    headerFontSize: 18,
    bodyFontSize: 14,
    headerColor: '#1f2937',
    textColor: '#374151',
    lineSpacing: '1.5'
};

// Paper size presets
const paperSizes = {
    custom: { width: 21.0, height: 29.7 },
    a4: { width: 21.0, height: 29.7 },
    a5: { width: 14.8, height: 21.0 },
    a6: { width: 10.5, height: 14.8 },
    receipt: { width: 7.75, height: 12.5 },
    business: { width: 8.5, height: 5.5 }
};

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    initializeFormatEventListeners();
    loadFormatSettings();
    updatePreview();
});

function initializeEventListeners() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const loadSheetsBtn = document.getElementById('loadSheetsBtn');
    const sheetsUrl = document.getElementById('sheetsUrl');
    const generateWordBtn = document.getElementById('generateWordBtn');
    const generatePdfBtn = document.getElementById('generatePdfBtn');

    // File upload events
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);

    // Google Sheets events
    loadSheetsBtn.addEventListener('click', loadGoogleSheetsData);
    sheetsUrl.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            loadGoogleSheetsData();
        }
    });

    // Generate document events
    generateWordBtn.addEventListener('click', () => generateDocument('word'));
    generatePdfBtn.addEventListener('click', () => generateDocument('pdf'));

    // Selection events
    document.getElementById('startRow').addEventListener('change', updateSelectionCount);
    document.getElementById('endRow').addEventListener('change', updateSelectionCount);
}

function initializeFormatEventListeners() {
    const formatInputs = [
        'paperWidth', 'paperHeight',
        'marginTop', 'marginBottom', 'marginLeft', 'marginRight',
        'fontFamily', 'headerFontSize', 'bodyFontSize',
        'headerColor', 'textColor', 'lineSpacing'
    ];
    
    formatInputs.forEach(inputId => {
        const input = document.getElementById(inputId);
        if (input) {
            input.addEventListener('change', updateFormat);
            input.addEventListener('input', updateFormat);
        }
    });
    
    const paperSizeSelect = document.getElementById('paperSizeSelect');
    if (paperSizeSelect) {
        paperSizeSelect.addEventListener('change', updatePaperSize);
    }
    
    console.log('Format event listeners initialized');
}

// Source switching functions
function switchToFileUpload() {
    currentDataSource = 'file';
    document.getElementById('fileUploadBtn').classList.add('active');
    document.getElementById('googleSheetsBtn').classList.remove('active');
    document.getElementById('fileUploadSection').style.display = 'block';
    document.getElementById('googleSheetsSection').style.display = 'none';
    resetSections();
}

function switchToGoogleSheets() {
    currentDataSource = 'sheets';
    document.getElementById('googleSheetsBtn').classList.add('active');
    document.getElementById('fileUploadBtn').classList.remove('active');
    document.getElementById('fileUploadSection').style.display = 'none';
    document.getElementById('googleSheetsSection').style.display = 'block';
    resetSections();
}

function resetSections() {
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('columnSelection').style.display = 'none';
    document.getElementById('formatCustomization').style.display = 'none';
    document.getElementById('dataTable').style.display = 'none';
    document.getElementById('selectionSection').style.display = 'none';
    document.getElementById('generateSection').style.display = 'none';
    document.getElementById('previewSection').style.display = 'none';
    document.getElementById('statusSection').style.display = 'none';
    document.getElementById('emptyState').style.display = 'block';
    allData = [];
    selectedColumns = [];
}

// File upload functions
function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.style.borderColor = '#3b82f6';
}

function handleDragLeave(e) {
    e.preventDefault();
    e.currentTarget.style.borderColor = '#d1d5db';
}

function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.style.borderColor = '#d1d5db';
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showStatus('⚠️ Please select a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }

    showStatus('Reading Excel file...', 'info');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            if (jsonData.length < 2) {
                showStatus('⚠️ Excel file must have at least 2 rows (header + data)', 'error');
                return;
            }
            
            processExcelData(jsonData);
            displayFileInfo(file.name, allData.length, 'Excel File');
            hideStatus();
        } catch (error) {
            console.error('Error reading file:', error);
            showStatus('❌ Error reading Excel file. Please make sure it\'s a valid Excel file.', 'error');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Google Sheets functions
async function loadGoogleSheetsData() {
    const urlInput = document.getElementById('sheetsUrl');
    const url = urlInput.value.trim();
    
    if (!url) {
        showStatus('⚠️ Please enter a Google Sheets URL', 'error');
        return;
    }
    
    if (!isValidGoogleSheetsUrl(url)) {
        showStatus('⚠️ Please enter a valid Google Sheets URL', 'error');
        return;
    }
    
    showStatus('Loading data from Google Sheets...', 'info');
    
    try {
        // Extract spreadsheet ID
        const spreadsheetId = extractSpreadsheetId(url);
        
        // Try different methods to load the data
        let jsonData = null;
        
        // Method 1: Try proxy approach first (most likely to work)
        try {
            jsonData = await loadViaAlternativeProxy(spreadsheetId);
            if (jsonData && jsonData.length >= 2) {
                console.log('Proxy method successful!');
            }
        } catch (error) {
            console.log('Proxy method failed:', error.message);
        }
        
        // Method 2: Try direct CSV only if proxy failed
        if (!jsonData || jsonData.length < 2) {
            try {
                jsonData = await loadViaDirectCSV(spreadsheetId);
                if (jsonData && jsonData.length >= 2) {
                    console.log('Direct CSV method successful!');
                }
            } catch (error) {
                console.log('Direct CSV failed:', error.message);
            }
        }
        
        // Method 3: Try improved CSV approach as last resort
        if (!jsonData || jsonData.length < 2) {
            try {
                jsonData = await loadViaJSONP(spreadsheetId);
                if (jsonData && jsonData.length >= 2) {
                    console.log('JSONP method successful!');
                }
            } catch (error) {
                console.log('JSONP method failed:', error.message);
            }
        }
        
        if (!jsonData || jsonData.length < 2) {
            throw new Error('Unable to load data from Google Sheets');
        }
        
        processExcelData(jsonData);
        displayFileInfo('Google Sheets', allData.length, 'Google Sheets');
        hideStatus();
        
    } catch (error) {
        console.error('Error loading Google Sheets:', error);
        
        // Show user-friendly error with working solution
        showWorkingGoogleSheetsInstructions();
        hideStatus();
    }
}

function extractSpreadsheetId(url) {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
        throw new Error('Invalid Google Sheets URL');
    }
    return match[1];
}

async function loadViaDirectCSV(spreadsheetId) {
    // Try different URL formats for published Google Sheets
    const urlFormats = [
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=0`,
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/pub?output=csv`,
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv`,
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv`
    ];
    
    for (const url of urlFormats) {
    try {
            console.log(`Trying direct CSV URL: ${url}`);
            
            const response = await fetch(url, {
                method: 'GET',
                mode: 'no-cors', // Try no-cors mode first
                credentials: 'omit'
            });
            
            // With no-cors, we can't read the response, so try cors mode
            if (response.type === 'opaque') {
                const corsResponse = await fetch(url, {
                    method: 'GET',
                    mode: 'cors',
                    credentials: 'omit',
                    headers: {
                        'Accept': 'text/csv,text/plain,*/*'
                    }
                });
                
                if (corsResponse.ok) {
                    const csvData = await corsResponse.text();
                    if (csvData && csvData.trim() && !csvData.includes('<!DOCTYPE html>')) {
                        console.log('Direct CSV success!');
                        return parseCSVData(csvData);
                    }
                }
            }
            
        } catch (error) {
            console.log(`URL ${url} failed:`, error.message);
            continue;
        }
    }
    
    throw new Error('All direct CSV methods failed');
}

async function loadViaJSONP(spreadsheetId) {
    // First try the direct CSV approach with proper headers
    try {
        const csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=0`;
        
        const response = await fetch(csvUrl, {
            method: 'GET',
            mode: 'cors',
            credentials: 'omit',
            headers: {
                'Accept': 'text/csv,text/plain,*/*',
                'Cache-Control': 'no-cache'
            }
        });
        
        if (response.ok) {
            const csvData = await response.text();
            return parseCSVData(csvData);
        }
        
        throw new Error('Direct CSV failed');
        
    } catch (error) {
        // If direct fails, try using a different proxy
        return await loadViaAlternativeProxy(spreadsheetId);
    }
}

async function loadViaAlternativeProxy(spreadsheetId) {
    // Prioritize proxies that are most likely to work
    const proxies = [
        'https://api.allorigins.win/raw?url=',
        'https://api.allorigins.win/get?url=', // Alternative allorigins format
        'https://corsproxy.io/?',
        'https://cors-anywhere.herokuapp.com/'
    ];
    
    const csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=0`;
    
    for (const proxy of proxies) {
        try {
            console.log(`Trying proxy: ${proxy.split('?')[0]}...`);
            
            const fullUrl = proxy + encodeURIComponent(csvUrl);
            
            // Add timeout to make it faster
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 8000); // 8 second timeout
            
            const response = await fetch(fullUrl, {
                method: 'GET',
                headers: {
                    'Accept': 'text/csv,text/plain,*/*'
                },
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);
            
            if (response.ok) {
                const data = await response.text();
                let csvData = data;
                
                // Handle allorigins wrapped response
                if (proxy.includes('allorigins.win/get')) {
                    try {
                        const jsonResponse = JSON.parse(data);
                        csvData = jsonResponse.contents;
                    } catch (e) {
                        // If not JSON, use as is
                    }
                }
                
                // Check if we got actual CSV data, not an error page
                if (csvData && csvData.includes(',') && !csvData.includes('<!DOCTYPE html>') && !csvData.includes('Access denied')) {
                    console.log(`✅ Proxy ${proxy.split('?')[0]} worked!`);
                    return parseCSVData(csvData);
                }
            }
        } catch (error) {
            if (error.name === 'AbortError') {
                console.log(`Proxy ${proxy.split('?')[0]} timed out`);
            } else {
                console.log(`Proxy ${proxy.split('?')[0]} failed:`, error.message);
    }
            continue;
        } 
    }
    
    throw new Error('All proxy methods failed');
}

function parseCSVData(csvData) {
    const lines = csvData.trim().split('\n');
    const result = [];
    
    for (let line of lines) {
        const row = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                row.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        row.push(current.trim());
        result.push(row);
    }
    
    return result;
}

function isValidGoogleSheetsUrl(url) {
    return url.includes('docs.google.com/spreadsheets/d/') && 
           (url.includes('/edit') || url.includes('/view') || url.includes('/pub'));
}

function processExcelData(jsonData) {
    const headers = jsonData[0];
    const dataRows = jsonData.slice(1).filter(row => row.some(cell => cell !== null && cell !== undefined && cell !== ''));
    
    // Convert to objects
    allData = dataRows.map((row, index) => {
        const rowObj = { _rowIndex: index + 1 };
        headers.forEach((header, colIndex) => {
            const columnName = header || `Column_${colIndex + 1}`;
            rowObj[columnName] = row[colIndex] || '';
        });
        return rowObj;
    });
    
    console.log('Processed data:', allData);
    
    // Show relevant sections
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('fileInfo').style.display = 'block';
    document.getElementById('columnSelection').style.display = 'block';
    document.getElementById('formatCustomization').style.display = 'block';
    document.getElementById('dataTable').style.display = 'block';
    
    // Populate column selection
    populateColumnSelection(headers);
    
    // Populate data table
    populateDataTable();
}

function populateColumnSelection(headers) {
    const columnList = document.getElementById('columnList');
    columnList.innerHTML = '';
    
    headers.forEach((header, index) => {
        const columnName = header || `Column_${index + 1}`;
        const columnItem = document.createElement('div');
        columnItem.className = 'column-item';
        columnItem.onclick = () => toggleColumn(columnName, columnItem);
        
        // Determine column type
        const sampleData = allData.slice(0, 5).map(row => row[columnName]).filter(val => val !== '');
        let columnType = 'text';
        if (sampleData.length > 0) {
            const firstValue = sampleData[0];
            if (!isNaN(firstValue) && !isNaN(parseFloat(firstValue))) {
                columnType = 'number';
            } else if (firstValue.includes('@')) {
                columnType = 'email';
            } else if (firstValue.match(/^\d{4}-\d{2}-\d{2}/) || firstValue.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                columnType = 'date';
            }
        }
        
        columnItem.innerHTML = `
            <div class="column-checkbox" id="checkbox-${columnName}"></div>
            <span class="column-name">${columnName}</span>
            <span class="column-type">${columnType}</span>
        `;
        
        columnList.appendChild(columnItem);
    });
}

function toggleColumn(columnName, columnElement) {
    const checkbox = columnElement.querySelector('.column-checkbox');
    const isSelected = selectedColumns.includes(columnName);
    
    if (isSelected) {
        // Remove column
        selectedColumns = selectedColumns.filter(col => col !== columnName);
        checkbox.classList.remove('checked');
        columnElement.classList.remove('selected');
    } else {
        // Add column
        selectedColumns.push(columnName);
        checkbox.classList.add('checked');
        columnElement.classList.add('selected');
    }
    
    updateSelectedFieldsPreview();
    updateSelectionSections();
    updatePreview();
}

function selectAllColumns() {
    const allColumnItems = document.querySelectorAll('.column-item');
    selectedColumns = [];
    
    allColumnItems.forEach(item => {
        const columnName = item.querySelector('.column-name').textContent;
        selectedColumns.push(columnName);
        item.querySelector('.column-checkbox').classList.add('checked');
        item.classList.add('selected');
    });
    
    updateSelectedFieldsPreview();
    updateSelectionSections();
    updatePreview();
}

function clearAllColumns() {
    const allColumnItems = document.querySelectorAll('.column-item');
    selectedColumns = [];
    
    allColumnItems.forEach(item => {
        item.querySelector('.column-checkbox').classList.remove('checked');
        item.classList.remove('selected');
    });
    
    updateSelectedFieldsPreview();
    updateSelectionSections();
    updatePreview();
}

function updateSelectedFieldsPreview() {
    const selectedFields = document.getElementById('selectedFields');
    
    if (selectedColumns.length === 0) {
        selectedFields.innerHTML = '<p class="no-selection">No fields selected</p>';
    } else {
        selectedFields.innerHTML = selectedColumns.map(col => 
            `<span class="selected-field">${col}</span>`
        ).join('');
    }
}

function updateSelectionSections() {
    if (selectedColumns.length > 0) {
        document.getElementById('selectionSection').style.display = 'block';
        document.getElementById('previewSection').style.display = 'block';
        document.getElementById('generateSection').style.display = 'block';
        populateSelectionTable();
    } else {
        document.getElementById('selectionSection').style.display = 'none';
        document.getElementById('previewSection').style.display = 'none';
        document.getElementById('generateSection').style.display = 'none';
    }
}

function populateDataTable() {
    const tableHead = document.getElementById('dataTableHead');
    const tableBody = document.getElementById('dataTableBody');
    const dataCount = document.getElementById('dataCount');
    
    if (allData.length === 0) return;
    
    // Get all column names
    const columns = Object.keys(allData[0]).filter(key => key !== '_rowIndex');
    
    // Create header
    tableHead.innerHTML = `
        <tr>
            ${columns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;
    
    // Create body (show first 10 rows for preview)
    const previewData = allData.slice(0, 10);
    tableBody.innerHTML = previewData.map(row => `
        <tr>
            ${columns.map(col => `<td>${row[col] || ''}</td>`).join('')}
        </tr>
    `).join('');
    
    dataCount.textContent = `${allData.length} records found${allData.length > 10 ? ' (showing first 10)' : ''}`;
}

function populateSelectionTable() {
    const tableHead = document.getElementById('selectionTableHead');
    const tableBody = document.getElementById('selectionTableBody');
    
    if (selectedColumns.length === 0 || allData.length === 0) return;
    
    // Create header with checkbox
    tableHead.innerHTML = `
        <tr>
            <th><input type="checkbox" id="selectAllRows" onchange="toggleAllRows(this)"></th>
            <th>#</th>
            ${selectedColumns.map(col => `<th>${col}</th>`).join('')}
        </tr>
    `;
    
    // Create body
    tableBody.innerHTML = allData.map((row, index) => `
        <tr>
            <td><input type="checkbox" class="row-checkbox" data-index="${index}" onchange="updateSelectionCount()"></td>
            <td>${index + 1}</td>
            ${selectedColumns.map(col => `<td>${row[col] || ''}</td>`).join('')}
        </tr>
    `).join('');
    
    updateSelectionCount();
}

function displayFileInfo(filename, count, sourceType) {
    const fileDetails = document.getElementById('fileDetails');
    const columnsFound = document.getElementById('columnsFound');
    
    fileDetails.textContent = `📁 ${filename} - ${count} records loaded from ${sourceType}`;
    
    if (allData.length > 0) {
        const columns = Object.keys(allData[0]).filter(key => key !== '_rowIndex');
        columnsFound.innerHTML = `
            <h4>Available Columns (${columns.length}):</h4>
            <div class="column-tags">
                ${columns.map(col => `<span class="column-tag">${col}</span>`).join('')}
            </div>
        `;
    }
}

// Selection functions
function switchToTableMode() {
    document.getElementById('tableModeBtn').classList.add('active');
    document.getElementById('rowModeBtn').classList.remove('active');
    document.getElementById('tableModeContent').classList.add('active');
    document.getElementById('rowModeContent').classList.remove('active');
    updateSelectionCount();
}

function switchToRowMode() {
    document.getElementById('rowModeBtn').classList.add('active');
    document.getElementById('tableModeBtn').classList.remove('active');
    document.getElementById('rowModeContent').classList.add('active');
    document.getElementById('tableModeContent').classList.remove('active');
    updateSelectionCount();
}

function selectAllRows() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectAllCheckbox = document.getElementById('selectAllRows');
    
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
    
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = true;
    }
    
    updateSelectionCount();
}

function clearAllRows() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectAllCheckbox = document.getElementById('selectAllRows');
    
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
    
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = false;
    }
    
    updateSelectionCount();
}

function selectFirstN() {
    const firstNInput = document.getElementById('firstNInput');
    firstNInput.style.display = firstNInput.style.display === 'none' ? 'flex' : 'none';
}

function applyFirstN() {
    const value = parseInt(document.getElementById('firstNValue').value);
    if (isNaN(value) || value < 1) {
        showStatus('⚠️ Please enter a valid number', 'warning');
        return;
    }
    
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach((checkbox, index) => {
        checkbox.checked = index < value;
    });
    
    document.getElementById('firstNInput').style.display = 'none';
    updateSelectionCount();
}

function toggleAllRows(selectAllCheckbox) {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
    updateSelectionCount();
}

function updateSelectionCount() {
    const selectionCount = document.getElementById('selectionCount');
    const selectedData = getSelectedData();
    
    selectionCount.textContent = `${selectedData.length} records selected`;
    
    // Enable/disable generate buttons
    const generateWordBtn = document.getElementById('generateWordBtn');
    const generatePdfBtn = document.getElementById('generatePdfBtn');
    
    if (selectedData.length > 0 && selectedColumns.length > 0) {
        generateWordBtn.disabled = false;
        generatePdfBtn.disabled = false;
    } else {
        generateWordBtn.disabled = true;
        generatePdfBtn.disabled = true;
    }
}

function getSelectedData() {
    const tableModeActive = document.getElementById('tableModeBtn').classList.contains('active');
    
    if (tableModeActive) {
        // Table selection mode
        const checkedBoxes = document.querySelectorAll('.row-checkbox:checked');
        return Array.from(checkedBoxes).map(checkbox => {
            const index = parseInt(checkbox.dataset.index);
            return allData[index];
        });
    } else {
        // Row range mode
        const startRow = parseInt(document.getElementById('startRow').value) || 1;
        const endRow = parseInt(document.getElementById('endRow').value) || allData.length;
        
        const start = Math.max(1, startRow) - 1; // Convert to 0-based index
        const end = Math.min(allData.length, endRow);
        
        return allData.slice(start, end);
    }
}

// Format functions
function switchFormatTab(tabName) {
    // Remove active class from all tabs and contents
    document.querySelectorAll('.format-tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.format-tab-content').forEach(content => content.classList.remove('active'));
    
    // Add active class to selected tab and content
    document.querySelector(`[onclick="switchFormatTab('${tabName}')"]`).classList.add('active');
    document.getElementById(tabName + 'Tab').classList.add('active');
}

function updatePaperSize() {
    const sizeSelect = document.getElementById('paperSizeSelect');
    const selectedSize = sizeSelect.value;
    
    if (selectedSize !== 'custom' && paperSizes[selectedSize]) {
        document.getElementById('paperWidth').value = paperSizes[selectedSize].width;
        document.getElementById('paperHeight').value = paperSizes[selectedSize].height;
        updateFormat();
    }
}

function updateFormat() {
    // Update format settings
    formatSettings.paperWidth = parseFloat(document.getElementById('paperWidth').value) || 21.0;
    formatSettings.paperHeight = parseFloat(document.getElementById('paperHeight').value) || 29.7;
    formatSettings.marginTop = parseFloat(document.getElementById('marginTop').value) || 2.0;
    formatSettings.marginBottom = parseFloat(document.getElementById('marginBottom').value) || 2.0;
    formatSettings.marginLeft = parseFloat(document.getElementById('marginLeft').value) || 2.0;
    formatSettings.marginRight = parseFloat(document.getElementById('marginRight').value) || 2.0;
    formatSettings.fontFamily = document.getElementById('fontFamily').value || 'Arial, sans-serif';
    formatSettings.headerFontSize = parseInt(document.getElementById('headerFontSize').value) || 18;
    formatSettings.bodyFontSize = parseInt(document.getElementById('bodyFontSize').value) || 14;
    formatSettings.headerColor = document.getElementById('headerColor').value || '#1f2937';
    formatSettings.textColor = document.getElementById('textColor').value || '#374151';
    formatSettings.lineSpacing = document.getElementById('lineSpacing').value || '1.5';
    
    // Save settings
    saveFormatSettings();
    
    // Update preview
    updatePreview();
}

function updatePreview() {
    const preview = document.getElementById('documentPreview');
    if (!preview) return;
    
    if (selectedColumns.length === 0 || allData.length === 0) {
        preview.innerHTML = `
            <div style="text-align: center; padding: 40px; color: #6b7280;">
                <p>Select columns and data to see preview</p>
            </div>
        `;
        return;
    }
    
    // Get first selected record for preview
    const selectedData = getSelectedData();
    const previewData = selectedData.length > 0 ? selectedData[0] : allData[0];
    
    const previewHTML = `
        <div style="
            font-family: ${formatSettings.fontFamily};
            font-size: ${formatSettings.bodyFontSize}px;
            color: ${formatSettings.textColor};
            line-height: ${formatSettings.lineSpacing};
            max-width: ${(formatSettings.paperWidth - formatSettings.marginLeft - formatSettings.marginRight) * 37.8}px;
            padding: ${formatSettings.marginTop * 37.8}px ${formatSettings.marginRight * 37.8}px ${formatSettings.marginBottom * 37.8}px ${formatSettings.marginLeft * 37.8}px;
            background: white;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            margin: 20px auto;
        ">
            <div style="
                text-align: center;
                margin-bottom: 30px;
                padding-bottom: 20px;
                border-bottom: 2px solid ${formatSettings.headerColor};
            ">
                <h1 style="
                    color: ${formatSettings.headerColor};
                    font-size: ${formatSettings.headerFontSize}px;
                    margin: 0 0 10px 0;
                    font-weight: bold;
                ">Document Preview</h1>
                <p style="margin: 0; color: #6b7280; font-size: ${formatSettings.bodyFontSize - 2}px;">
                    Generated from flexible data selection
                </p>
            </div>
            
            <div style="margin-bottom: 30px;">
                ${selectedColumns.map(col => `
                    <div style="margin-bottom: 15px; display: flex; align-items: flex-start;">
                        <strong style="
                            min-width: 120px;
                            color: ${formatSettings.headerColor};
                            margin-right: 15px;
                            font-weight: 600;
                        ">${col}:</strong>
                        <span style="flex: 1;">${previewData[col] || 'N/A'}</span>
                    </div>
                `).join('')}
            </div>
            
            <div style="
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #e5e7eb;
                text-align: center;
                color: #6b7280;
                font-size: ${formatSettings.bodyFontSize - 2}px;
            ">
                <p style="margin: 0;">Generated by JD Sons Document Generator</p>
                <p style="margin: 5px 0 0 0;">Contact: Jemish (9773046615)</p>
            </div>
        </div>
    `;
    
    preview.innerHTML = previewHTML;
}

function resetToDefaults() {
    // Reset to default values
    document.getElementById('paperSizeSelect').value = 'a4';
    document.getElementById('paperWidth').value = 21.0;
    document.getElementById('paperHeight').value = 29.7;
    document.getElementById('marginTop').value = 2.0;
    document.getElementById('marginBottom').value = 2.0;
    document.getElementById('marginLeft').value = 2.0;
    document.getElementById('marginRight').value = 2.0;
    document.getElementById('fontFamily').value = 'Arial, sans-serif';
    document.getElementById('headerFontSize').value = 18;
    document.getElementById('bodyFontSize').value = 14;
    document.getElementById('headerColor').value = '#1f2937';
    document.getElementById('textColor').value = '#374151';
    document.getElementById('lineSpacing').value = '1.5';
    
    updateFormat();
}

function saveFormatSettings() {
    localStorage.setItem('flexibleFormatSettings', JSON.stringify(formatSettings));
}

function loadFormatSettings() {
    const saved = localStorage.getItem('flexibleFormatSettings');
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            formatSettings = { ...formatSettings, ...parsed };
            
            // Apply to form
            Object.keys(formatSettings).forEach(key => {
                const element = document.getElementById(key);
                if (element) {
                    element.value = formatSettings[key];
                }
            });
        } catch (e) {
            console.error('Error loading format settings:', e);
        }
    }
}

// Document generation
async function generateDocument(format = 'word') {
    const selectedData = getSelectedData();
    
    if (selectedData.length === 0) {
        showStatus('⚠️ Please select at least one record', 'warning');
        return;
    }
    
    if (selectedColumns.length === 0) {
        showStatus('⚠️ Please select at least one column', 'warning');
        return;
    }
    
    // Check if required libraries are loaded for Word documents
    if (format === 'word' && typeof docx === 'undefined') {
        showStatus('❌ Word document library not loaded. Please refresh the page and try again.', 'error');
        return;
    }
    
    showStatus(`Generating ${format.toUpperCase()} document with ${selectedData.length} page${selectedData.length > 1 ? 's' : ''}...`, 'info');
    
    try {
        if (format === 'word') {
            await createSingleWordDocument(selectedData);
        } else {
            await createSinglePdfDocument(selectedData);
        }
    } catch (error) {
        console.error('Error generating document:', error);
        showStatus(`❌ Error generating document: ${error.message}`, 'error');
    }
}

async function createSingleWordDocument(data) {
    try {
        const doc = new docx.Document({
            styles: {
                paragraphStyles: [
                    {
                        id: "Heading1",
                        name: "Heading 1",
                        basedOn: "Normal",
                        next: "Normal",
                        run: {
                            size: formatSettings.headerFontSize * 2,
                            bold: true,
                            color: formatSettings.headerColor.replace('#', ''),
                            font: formatSettings.fontFamily.split(',')[0].replace(/['"]/g, ''),
                        },
                        paragraph: {
                            spacing: {
                                after: 300,
                            },
                            alignment: docx.AlignmentType.CENTER,
                        },
                    },
                    {
                        id: "Normal",
                        name: "Normal",
                        run: {
                            size: formatSettings.bodyFontSize * 2,
                            font: formatSettings.fontFamily.split(',')[0].replace(/['"]/g, ''),
                            color: formatSettings.textColor.replace('#', ''),
                        },
                        paragraph: {
                            spacing: {
                                line: Math.floor(parseFloat(formatSettings.lineSpacing) * 240),
                                lineRule: docx.LineRuleType.AUTO,
                            },
                        },
                    },
                    {
                        id: "FieldLabel",
                        name: "Field Label",
                        basedOn: "Normal",
                        run: {
                            bold: true,
                            color: formatSettings.headerColor.replace('#', ''),
                        },
                    },
                ],
            },
            sections: data.map((record, index) => ({
                properties: {
                    page: {
                        size: {
                            width: docx.convertMillimetersToTwip(formatSettings.paperWidth * 10),
                            height: docx.convertMillimetersToTwip(formatSettings.paperHeight * 10),
                        },
                        margin: {
                            top: docx.convertMillimetersToTwip(formatSettings.marginTop * 10),
                            bottom: docx.convertMillimetersToTwip(formatSettings.marginBottom * 10),
                            left: docx.convertMillimetersToTwip(formatSettings.marginLeft * 10),
                            right: docx.convertMillimetersToTwip(formatSettings.marginRight * 10),
                        },
                    },
                },
                children: [
                    // Header
                    new docx.Paragraph({
                        text: "Document Generated from Flexible Data",
                        style: "Heading1",
                    }),
                    
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: `Page ${index + 1} of ${data.length}`,
                                italics: true,
                                size: (formatSettings.bodyFontSize - 2) * 2,
                                color: "666666",
                            }),
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 300 },
                    }),
                    
                    // Data fields
                    ...selectedColumns.map(column => 
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: `${column}: `,
                                    bold: true,
                                    color: formatSettings.headerColor.replace('#', ''),
                                }),
                                new docx.TextRun({
                                    text: `${record[column] || 'N/A'}`,
                                    color: formatSettings.textColor.replace('#', ''),
                                }),
                            ],
                            spacing: { after: 120 },
                        })
                    ),
                    
                    new docx.Paragraph({
                        text: "",
                        spacing: { before: 400, after: 200 },
                    }),
                    
                    // Footer
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Generated by: JD Sons Document Generator",
                                italics: true,
                                size: (formatSettings.bodyFontSize - 2) * 2,
                                color: "666666",
                            }),
                        ],
                        alignment: docx.AlignmentType.CENTER,
                    }),
                    
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: "Contact: Jemish (9773046615)",
                                italics: true,
                                size: (formatSettings.bodyFontSize - 2) * 2,
                                color: "666666",
                            }),
                        ],
                        alignment: docx.AlignmentType.CENTER,
                    }),
                ],
            }))
        });
        
        // Generate document buffer
        const buffer = await docx.Packer.toBuffer(doc);
        
        // Create blob and open in new tab
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        });
        
        const url = URL.createObjectURL(blob);
        const filename = `Flexible_Document_${data.length}_Records_${new Date().toISOString().split('T')[0]}.docx`;
        
        // Try to open in new tab first, fallback to direct download
        try {
            const newTab = window.open('', '_blank');
            if (newTab) {
                // Create a simple HTML page that triggers the download
                newTab.document.write(`
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <title>Word Document Download</title>
                        <style>
                            body {
                                font-family: Arial, sans-serif;
                                display: flex;
                                justify-content: center;
                                align-items: center;
                                height: 100vh;
                                margin: 0;
                                background: #f5f5f5;
                            }
                            .download-container {
                                text-align: center;
                                background: white;
                                padding: 40px;
                                border-radius: 10px;
                                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                            }
                            .download-btn {
                                background: #007bff;
                                color: white;
                                border: none;
                                padding: 15px 30px;
                                border-radius: 5px;
                                cursor: pointer;
                                font-size: 16px;
                                margin-top: 20px;
                            }
                            .download-btn:hover {
                                background: #0056b3;
                            }
                        </style>
                    </head>
                    <body>
                        <div class="download-container">
                            <h2>📝 Word Document Ready</h2>
                            <p>Your document with ${data.length} page${data.length > 1 ? 's' : ''} is ready for download.</p>
                            <button class="download-btn" onclick="downloadFile()">📥 Download Word Document</button>
                            <p style="margin-top: 20px; font-size: 14px; color: #666;">
                                If download doesn't start automatically, click the button above.
                            </p>
                        </div>
                        <script>
                            function downloadFile() {
                                const a = document.createElement('a');
                                a.href = '${url}';
                                a.download = '${filename}';
                                document.body.appendChild(a);
                                a.click();
                                document.body.removeChild(a);
                            }
                            // Auto-download after a short delay
                            setTimeout(downloadFile, 1000);
                        </script>
                    </body>
                    </html>
                `);
                newTab.document.close();
                
                // Clean up after some time
                setTimeout(() => {
                    URL.revokeObjectURL(url);
                }, 10000);
                
                showStatus(`✅ Word document with ${data.length} pages opened in new tab!`, 'success');
            } else {
                throw new Error('Unable to open new tab');
            }
        } catch (error) {
            // Fallback: direct download
            downloadFile(blob, filename);
            URL.revokeObjectURL(url);
            showStatus(`✅ Word document downloaded with ${data.length} pages!`, 'success');
        }
        
    } catch (error) {
        console.error('Error creating Word document:', error);
        throw error;
    }
}

async function createSinglePdfDocument(data) {
    try {
        // Create a comprehensive HTML document with all pages
        let htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Flexible Document - ${data.length} Records</title>
                <style>
                    body {
                        font-family: ${formatSettings.fontFamily};
                        font-size: ${formatSettings.bodyFontSize}px;
                        color: ${formatSettings.textColor};
                        line-height: ${formatSettings.lineSpacing};
                        margin: 0;
                        padding: 0;
                    }
                    
                    .page {
                        width: ${formatSettings.paperWidth}cm;
                        min-height: ${formatSettings.paperHeight}cm;
                        margin: ${formatSettings.marginTop}cm ${formatSettings.marginRight}cm ${formatSettings.marginBottom}cm ${formatSettings.marginLeft}cm;
                        padding: 20px;
                        background: white;
                        page-break-after: always;
                        box-sizing: border-box;
                        position: relative;
                    }
                    
                    .page:last-child {
                        page-break-after: avoid;
                    }
                    
                    h1 {
                        color: ${formatSettings.headerColor};
                        font-size: ${formatSettings.headerFontSize}px;
                        text-align: center;
                        margin: 0 0 10px 0;
                        padding-bottom: 15px;
                        border-bottom: 2px solid ${formatSettings.headerColor};
                    }
                    
                    .page-info {
                        text-align: center;
                        color: #666;
                        font-size: ${formatSettings.bodyFontSize - 2}px;
                        font-style: italic;
                        margin-bottom: 30px;
                    }
                    
                    .field {
                        margin-bottom: 15px;
                        display: flex;
                        align-items: flex-start;
                        min-height: 20px;
                    }
                    
                    .field-label {
                        min-width: 150px;
                        font-weight: bold;
                        color: ${formatSettings.headerColor};
                        margin-right: 15px;
                        flex-shrink: 0;
                    }
                    
                    .field-value {
                        flex: 1;
                        word-wrap: break-word;
                        word-break: break-word;
                    }
                    
                    .footer {
                        position: absolute;
                        bottom: 20px;
                        left: 20px;
                        right: 20px;
                        text-align: center;
                        color: #666;
                        font-size: ${formatSettings.bodyFontSize - 2}px;
                        font-style: italic;
                        border-top: 1px solid #e5e7eb;
                        padding-top: 15px;
                    }
                    
                    /* Print styles */
                    @media print {
                        body {
                            margin: 0;
                            padding: 0;
                        }
                        
                        .page {
                            margin: 0;
                            padding: ${formatSettings.marginTop}cm ${formatSettings.marginRight}cm ${formatSettings.marginBottom}cm ${formatSettings.marginLeft}cm;
                            min-height: ${formatSettings.paperHeight - formatSettings.marginTop - formatSettings.marginBottom}cm;
                            page-break-inside: avoid;
                        }
                        
                    @page {
                        size: ${formatSettings.paperWidth}cm ${formatSettings.paperHeight}cm;
                            margin: 0;
                        }
                        
                        .footer {
                            position: fixed;
                            bottom: ${formatSettings.marginBottom}cm;
                            left: ${formatSettings.marginLeft}cm;
                            right: ${formatSettings.marginRight}cm;
                        }
                        
                        .print-button, .document-info {
                            display: none !important;
                        }
                        
                        .page-info {
                            margin-bottom: 20px;
                        }
                    }
                    
                    /* Screen styles for preview */
                    @media screen {
                        body {
                            background: #f5f5f5;
                            padding: 20px;
                        }
                        
                        .page {
                            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                            margin: 20px auto;
                            border: 1px solid #ddd;
                        }
                        
                        .print-button {
                            position: fixed;
                            top: 20px;
                            right: 20px;
                            background: #007bff;
                            color: white;
                            border: none;
                            padding: 12px 24px;
                            border-radius: 5px;
                            cursor: pointer;
                            font-size: 16px;
                            font-weight: bold;
                            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
                            z-index: 1000;
                        }
                        
                        .print-button:hover {
                            background: #0056b3;
                        }
                        
                        .document-info {
                            position: fixed;
                            top: 20px;
                            left: 20px;
                            background: #28a745;
                            color: white;
                            padding: 10px 16px;
                            border-radius: 5px;
                            font-weight: bold;
                            z-index: 1000;
                        }
                    }
                </style>
            </head>
            <body>
                <div class="document-info">
                    📄 Document: ${data.length} Page${data.length > 1 ? 's' : ''}
                </div>
                <button class="print-button" onclick="window.print()">🖨️ Print / Save as PDF</button>
        `;
        
        // Add each record as a separate page
        data.forEach((record, index) => {
            htmlContent += `
                <div class="page">
                <h1>Document Generated from Flexible Data</h1>
                    <div class="page-info">Page ${index + 1} of ${data.length}</div>
                
                    <div class="content">
                ${selectedColumns.map(column => `
                    <div class="field">
                        <span class="field-label">${column}:</span>
                        <span class="field-value">${record[column] || 'N/A'}</span>
                    </div>
                `).join('')}
                    </div>
                
                <div class="footer">
                        <div>Generated by: JD Sons Document Generator</div>
                        <div>Contact: Jemish (9773046615)</div>
                </div>
                </div>
            `;
        });
        
        htmlContent += `
            </body>
            </html>
        `;
        
        // Open in new tab
        const newTab = window.open('', '_blank');
        if (newTab) {
            newTab.document.write(htmlContent);
            newTab.document.close();
            
            // Focus the new tab
            newTab.focus();
            
            showStatus(`✅ PDF document with ${data.length} pages opened in new tab! Click "Print / Save as PDF" to save.`, 'success');
        } else {
            // Fallback: create downloadable HTML file
            const blob = new Blob([htmlContent], { type: 'text/html' });
            downloadFile(blob, `Flexible_Document_${data.length}_Records_${new Date().toISOString().split('T')[0]}.html`);
            showStatus(`✅ HTML document downloaded with ${data.length} pages! Open in browser and print to save as PDF.`, 'success');
        }
        
    } catch (error) {
        console.error('Error creating PDF document:', error);
        throw error;
    }
}

function downloadFile(blob, filename) {
    try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    } catch (error) {
        console.error('Error downloading file:', error);
        showStatus('❌ Error downloading file. Please try again.', 'error');
    }
}

// Status functions
function showStatus(message, type = 'info') {
    const statusElement = document.querySelector('.status-content');
    const statusSection = document.querySelector('.status-section');
    
    if (!statusElement || !statusSection) return;
    
    statusElement.textContent = message;
    statusSection.className = `status-section status-${type}`;
    statusSection.style.display = 'block';
    
    if (type === 'success' || type === 'error') {
        setTimeout(hideStatus, 5000);
    }
}

function hideStatus() {
    const statusSection = document.querySelector('.status-section');
    if (statusSection) {
    statusSection.style.display = 'none';
    }
}

function showWorkingGoogleSheetsInstructions() {
    // Create a modal with working solution
    const modal = document.createElement('div');
    modal.className = 'instruction-modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>🔧 Let's Make This Work!</h3>
                <button onclick="closeInstructionModal()" class="close-btn">×</button>
            </div>
            <div class="modal-body">
                <p>Great! The system is working but had to try multiple methods to load your data.</p>
                
                <div class="solution-box">
                    <h4>🚀 To make it load faster next time:</h4>
                    <ol>
                        <li>Open your Google Sheet: <a href="${document.getElementById('sheetsUrl').value}" target="_blank">Click here</a></li>
                        <li>Click <strong>Share</strong> (top right) → Set to <strong>"Anyone with the link can view"</strong></li>
                        <li><strong>IMPORTANT:</strong> Go to <strong>File → Share → Publish to web</strong></li>
                        <li>Select <strong>"Entire Document"</strong> and <strong>"Comma-separated values (.csv)"</strong></li>
                        <li>Check <strong>"Automatically republish when changes are made"</strong></li>
                        <li>Click <strong>"Publish"</strong> and confirm</li>
                    </ol>
                    <p style="color: #059669; font-weight: 600; margin-top: 0.5rem;">✅ This will make future loads much faster!</p>
                </div>
                
                <div class="alternative-box">
                    <h4>📁 Alternative - Download & Upload:</h4>
                    <p>If the above doesn't work, just download as Excel and upload:</p>
                    <ol>
                        <li>File → Download → Microsoft Excel (.xlsx)</li>
                        <li>Switch to "Upload Excel File" tab</li>
                        <li>Upload the downloaded file</li>
                    </ol>
                </div>
            </div>
            <div class="modal-actions">
                <button onclick="openGoogleSheet()" class="open-sheet-btn">🔗 Open My Sheet</button>
                <button onclick="retryLoadGoogleSheets()" class="retry-btn">🔄 Try Again</button>
                <button onclick="switchToFileUpload(); closeInstructionModal();" class="switch-btn">📁 Use File Upload</button>
            </div>
        </div>
        <div class="modal-overlay" onclick="closeInstructionModal()"></div>
    `;
    
    document.body.appendChild(modal);
}

function retryLoadGoogleSheets() {
    closeInstructionModal();
    loadGoogleSheetsData();
}

function openGoogleSheet() {
    const url = document.getElementById('sheetsUrl').value.trim();
    if (url) {
        window.open(url, '_blank');
    }
}

function closeInstructionModal() {
    const modal = document.querySelector('.instruction-modal');
    if (modal) {
        document.body.removeChild(modal);
    }
} 