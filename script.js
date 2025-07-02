let customerData = [];
let currentDataSource = 'file'; // 'file' or 'sheets'

// Format settings with defaults
let formatSettings = {
    paperWidth: 7.75,
    paperHeight: 12.5,
    marginTop: 0.2,
    marginBottom: 0.2,
    marginLeft: 0.2,
    marginRight: 0.2,
    fontFamily: 'Arial, sans-serif',
    headerFontSize: 22,
    bodyFontSize: 20,
    headerColor: '#dc2626',
    textColor: '#1f2937',
    companyColor: '#2563eb',
    lineSpacing: '1.2'
};

// Paper size presets
const paperSizes = {
    custom: { width: 7.75, height: 12.5 },
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
}

function initializeFormatEventListeners() {
    // Add event listeners to all format inputs for real-time updates
    const formatInputs = [
        'paperWidth', 'paperHeight',
        'marginTop', 'marginBottom', 'marginLeft', 'marginRight',
        'fontFamily', 'headerFontSize', 'bodyFontSize',
        'headerColor', 'textColor', 'companyColor', 'lineSpacing'
    ];
    
    formatInputs.forEach(inputId => {
        const input = document.getElementById(inputId);
        if (input) {
            // Add event listeners for immediate updates
            input.addEventListener('change', updateFormat);
            input.addEventListener('input', updateFormat);
            
            // For color inputs, also listen to color picker changes
            if (input.type === 'color') {
                input.addEventListener('input', updateFormat);
            }
        }
    });
    
    // Special handling for paper size select
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
    document.getElementById('formatCustomization').style.display = 'none'; // Hide format customization
    document.getElementById('selectionSection').style.display = 'none';
    document.getElementById('generateSection').style.display = 'none';
    document.getElementById('previewSection').style.display = 'none'; // Hide preview section
    document.getElementById('statusSection').style.display = 'none';
    document.getElementById('customerTable').style.display = 'none'; // Hide customer table
    document.getElementById('emptyState').style.display = 'block'; // Show empty state
    customerData = [];
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
        alert('Please select a valid Excel file (.xlsx or .xls)');
        return;
    }

    showStatus('Reading Excel file...');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            if (jsonData.length < 2) {
                alert('Excel file must have at least 2 rows (header + data)');
                hideStatus();
                return;
            }
            
            processExcelData(jsonData);
            displayFileInfo(file.name, customerData.length, 'Excel File');
            hideStatus();
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading Excel file. Please make sure it\'s a valid Excel file.');
            hideStatus();
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Google Sheets functions
async function loadGoogleSheetsData() {
    const urlInput = document.getElementById('sheetsUrl');
    const url = urlInput.value.trim();
    
    if (!url) {
        alert('Please enter a Google Sheets URL');
        return;
    }
    
    if (!isValidGoogleSheetsUrl(url)) {
        alert('Please enter a valid Google Sheets URL');
        return;
    }
    
    showStatus('Loading data from Google Sheets...');
    
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
        displayFileInfo(extractSheetName(url), customerData.length, 'Google Sheets');
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
    const lines = csvData.split('\n');
    const result = [];
    
    for (let line of lines) {
        if (line.trim()) {
            // Simple CSV parsing - handle quoted values
            const row = [];
            let inQuotes = false;
            let currentField = '';
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                
                if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    row.push(currentField.trim());
                    currentField = '';
                } else {
                    currentField += char;
                }
            }
            
            // Add the last field
            row.push(currentField.trim());
            result.push(row);
        }
    }
    
    return result;
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

function isValidGoogleSheetsUrl(url) {
    return url.includes('docs.google.com/spreadsheets/d/') && 
           (url.includes('/edit') || url.includes('/view') || url.includes('/pub'));
}

function extractSheetName(url) {
    return 'Google Sheets Document';
}

// Common data processing functions
function processExcelData(jsonData) {
    const headers = jsonData[0];
    const rows = jsonData.slice(1);
    
    // Find required column indices
    const customerNameIndex = findColumnIndex(headers, ['customer name', 'name', 'customer']);
    const addressIndex = findColumnIndex(headers, ['address', 'addr']);
    const contactIndex = findColumnIndex(headers, ['contact number', 'contact', 'phone', 'mobile']);
    
    if (customerNameIndex === -1 || addressIndex === -1 || contactIndex === -1) {
        alert('Required columns not found. Please ensure your data has columns for Customer Name, Address, and Contact Number.');
        return;
    }
    
    // Process rows and maintain correct Excel row numbers
    customerData = [];

    console.log('Rows:ss', rows);
    
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const customerName = row[customerNameIndex] || '';
        const address = row[addressIndex] || '';
        const contactNumber = row[contactIndex] || '';
        
        // Only include rows with customer name (but keep track of actual Excel row number)
        if (customerName.trim() !== '') {
            customerData.push({
                excelRowNumber: i, // Actual Excel row number (i + 2 because row 1 is header, and i starts at 0)
                customerName: customerName,
                address: address,
                contactNumber: contactNumber
            });
        }
    }
    
    console.log('Processed customer data:', customerData);
    console.log('Excel row numbers:', customerData.map(c => `Row ${c.excelRowNumber}: ${c.customerName}`));
}

function findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i]?.toString().toLowerCase() || '';
        for (const name of possibleNames) {
            if (header.includes(name)) {
                return i;
            }
        }
    }
    return -1;
}

function displayFileInfo(filename, count, sourceType) {
    const fileInfo = document.getElementById('fileInfo');
    const fileDetails = document.getElementById('fileDetails');
    const columnsFound = document.getElementById('columnsFound');
    
    fileDetails.innerHTML = `
        <strong>📊 ${sourceType}:</strong> ${filename}<br>
        <strong>👥 Total Customers Found:</strong> ${count}
    `;
    
    // Show which columns were detected
    const sampleCustomer = customerData[0];
    if (sampleCustomer) {
        columnsFound.innerHTML = `
            <h4>✅ Detected Columns:</h4>
            <ul>
                <li><strong>Customer Name:</strong> ${sampleCustomer.customerName}</li>
                <li><strong>Address:</strong> ${sampleCustomer.address}</li>
                <li><strong>Contact Number:</strong> ${sampleCustomer.contactNumber}</li>
            </ul>
        `;
    }
    
    // Hide empty state and show content sections
    document.getElementById('emptyState').style.display = 'none';
    fileInfo.style.display = 'block';
    document.getElementById('formatCustomization').style.display = 'block'; // Show format customization
    document.getElementById('selectionSection').style.display = 'block';
    document.getElementById('generateSection').style.display = 'block';
    document.getElementById('previewSection').style.display = 'block'; // Show preview section
    
    // Show customer table
    populateCustomerTable();
    document.getElementById('customerTable').style.display = 'block';
    
    // Update row number limits
    const startRow = document.getElementById('startRow');
    const endRow = document.getElementById('endRow');
    const maxRow = customerData.length + 1;
    
    startRow.max = maxRow;
    endRow.max = maxRow;
    endRow.value = Math.min(26, maxRow);
}

function populateCustomerTable() {
    const tableBody = document.getElementById('customersTableBody');
    tableBody.innerHTML = '';
    
    customerData.forEach((customer, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td style="text-align: center;">
                <input type="checkbox" class="customer-checkbox" 
                       onchange="updateSelection()" 
                       data-customer-index="${index}">
            </td>
            <td style="text-align: center;">
                <span class="row-number">${customer.excelRowNumber}</span>
            </td>
            <td>
                <span class="customer-name">${customer.customerName}</span>
            </td>
            <td>
                <span class="customer-address">${customer.address}</span>
            </td>
            <td>
                <span class="customer-contact">${customer.contactNumber}</span>
            </td>
        `;
        tableBody.appendChild(row);
    });
    
    updateSelectionCount();
}

function selectAllCustomers() {
    const checkboxes = document.querySelectorAll('.customer-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
        checkbox.closest('tr').classList.add('selected');
    });
    updateSelectionCount();
}

function clearAllCustomers() {
    const checkboxes = document.querySelectorAll('.customer-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
        checkbox.closest('tr').classList.remove('selected');
    });
    updateSelectionCount();
}

function selectFirstN() {
    const checkboxes = document.querySelectorAll('.customer-checkbox');
    const maxSelect = Math.min(25, checkboxes.length);
    
    // Clear all first
    clearAllCustomers();
    
    // Select first N
    for (let i = 0; i < maxSelect; i++) {
        checkboxes[i].checked = true;
        checkboxes[i].closest('tr').classList.add('selected');
    }
    updateSelectionCount();
}

function updateSelection() {
    const checkboxes = document.querySelectorAll('.customer-checkbox');
    checkboxes.forEach(checkbox => {
        const row = checkbox.closest('tr');
        if (checkbox.checked) {
            row.classList.add('selected');
        } else {
            row.classList.remove('selected');
        }
    });
    updateSelectionCount();
}

function updateSelectionCount() {
    const selectedCheckboxes = document.querySelectorAll('.customer-checkbox:checked');
    document.getElementById('selectedCount').textContent = selectedCheckboxes.length;
}

function getSelectedCustomers() {
    const selectedCheckboxes = document.querySelectorAll('.customer-checkbox:checked');
    const selectedCustomers = [];
    
    selectedCheckboxes.forEach(checkbox => {
        const customerIndex = parseInt(checkbox.dataset.customerIndex);
        selectedCustomers.push(customerData[customerIndex]);
    });
    
    return selectedCustomers;
}

function switchToTableMode() {
    document.getElementById('tableSelectionMode').style.display = 'block';
    document.getElementById('rowSelectionMode').style.display = 'none';
    
    document.getElementById('tableModeBtn').classList.add('active');
    document.getElementById('rowModeBtn').classList.remove('active');
}

function switchToRowMode() {
    document.getElementById('tableSelectionMode').style.display = 'none';
    document.getElementById('rowSelectionMode').style.display = 'block';
    
    document.getElementById('tableModeBtn').classList.remove('active');
    document.getElementById('rowModeBtn').classList.add('active');
}

// Status functions
function showStatus(message) {
    const statusSection = document.getElementById('statusSection');
    const statusMessage = document.getElementById('statusMessage');
    
    statusMessage.textContent = message;
    statusSection.style.display = 'block';
}

function hideStatus() {
    document.getElementById('statusSection').style.display = 'none';
}

function validateInputs() {
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    if (isNaN(startRow) || isNaN(endRow)) {
        alert('Please enter valid row numbers');
        return false;
    }
    
    if (startRow < 2) {
        alert('Start row must be at least 2 (row 1 is the header)');
        return false;
    }
    
    if (startRow > endRow) {
        alert('Start row cannot be greater than end row');
        return false;
    }
    
    const maxRow = customerData.length + 1;
    if (endRow > maxRow) {
        alert(`End row cannot be greater than ${maxRow} (total data rows + 1)`);
        return false;
    }
    
    if (customerData.length === 0) {
        alert('No customer data available. Please load data first.');
        return false;
    }
    
    return true;
}

async function generateDocument(format = 'word') {
    let selectedCustomers = [];
    
    // Make sure format settings are up-to-date before generating
    updateFormat();
    
    // Check which mode is active
    const tableMode = document.getElementById('tableSelectionMode').style.display !== 'none';
    
    if (tableMode) {
        // Table selection mode
        selectedCustomers = getSelectedCustomers();
        
        if (selectedCustomers.length === 0) {
            alert('Please select at least one customer from the table above.');
            return;
        }
    } else {
        // Row number mode
        if (!validateInputs()) {
            return;
        }
        
        const startRowNumber = parseInt(document.getElementById('startRow').value);
        const endRowNumber = parseInt(document.getElementById('endRow').value);
        
        // Filter customers based on selected row range
        selectedCustomers = customerData.filter(customer => 
            customer.excelRowNumber >= startRowNumber && customer.excelRowNumber <= endRowNumber
        );
        
        if (selectedCustomers.length === 0) {
            alert('No customers found in the selected row range.');
            return;
        }
    }
    
    if (customerData.length === 0) {
        alert('No customer data available. Please load data first.');
        return;
    }
    
    showStatus(`Generating ${format.toUpperCase()} document for ${selectedCustomers.length} customers...`);
    
    try {
        // Get row numbers for filename
        const rowNumbers = selectedCustomers.map(c => c.excelRowNumber).sort((a, b) => a - b);
        const startRow = rowNumbers[0];
        const endRow = rowNumbers[rowNumbers.length - 1];
        
        if (format === 'word') {
            const doc = await createWordDocument(selectedCustomers);
            const blob = await docx.Packer.toBlob(doc);
            downloadDocument(blob, `JD_Sons_Customers_Rows_${startRow}-${endRow}.docx`);
        } else if (format === 'pdf') {
            await generatePdfDocument(selectedCustomers, startRow, endRow);
        }
        
        hideStatus();
        
        console.log(`Document generated successfully - ${format.toUpperCase()}: ${formatSettings.paperWidth}×${formatSettings.paperHeight}cm, Font: ${formatSettings.fontFamily.split(',')[0]} ${formatSettings.bodyFontSize}px, Customers: ${selectedCustomers.length}`);
        
    } catch (error) {
        console.error(`Error generating ${format} document:`, error);
        alert(`Error generating ${format.toUpperCase()} document. Please try again.`);
        hideStatus();
    }
}

async function createWordDocument(customers) {
    const children = [];
    
    // Convert hex colors to docx format (remove #)
    const headerColor = formatSettings.headerColor.replace('#', '');
    const textColor = formatSettings.textColor.replace('#', '');
    const companyColor = formatSettings.companyColor.replace('#', '');
    
    // Convert font sizes to Word format (multiply by 2 for half-points)
    const headerSize = formatSettings.headerFontSize * 2;
    const bodySize = formatSettings.bodyFontSize * 2;
    
    for (let i = 0; i < customers.length; i++) {
        const customer = customers[i];
        
        // Create page content for each customer
        const pageChildren = [
            // Customer details first
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Customer Details:",
                        bold: true,
                        size: headerSize,
                        color: headerColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 600, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            }),
            
            // Customer Name
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Customer Name: ",
                        bold: true,
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    }),
                    new docx.TextRun({
                        text: customer.customerName,
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 500, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            }),
            
            // Address
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Address: ",
                        bold: true,
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    }),
                    new docx.TextRun({
                        text: customer.address,
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 500, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            }),
            
            // Contact Number
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Contact Number: ",
                        bold: true,
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    }),
                    new docx.TextRun({
                        text: customer.contactNumber.toString(),
                        size: bodySize,
                        color: textColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 700, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            }),
            
            // Divider line
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "─".repeat(40),
                        color: "666666",
                        size: Math.round(bodySize * 0.8),
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 600 }
            }),
            
            // From details (Jemish's details)
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "From: Jemish (JD Jewellery)",
                        bold: true,
                        size: bodySize,
                        color: companyColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 400, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Contact Number: 9773046615",
                        bold: true,
                        size: Math.round(bodySize * 0.95),
                        color: companyColor,
                        font: formatSettings.fontFamily.split(',')[0].trim()
                    })
                ],
                spacing: { after: 500, line: Math.round(parseFloat(formatSettings.lineSpacing) * 240) }
            })
        ];
        
        // Add all page elements
        children.push(...pageChildren);
        
        // Add page break (except for the last customer)
        if (i < customers.length - 1) {
            children.push(
                new docx.Paragraph({
                    children: [new docx.PageBreak()]
                })
            );
        }
    }
    
    // Use the exact paper dimensions as specified by the user
    const pageWidth = formatSettings.paperWidth;
    const pageHeight = formatSettings.paperHeight;
    
    // Convert cm to twips (1 cm = 567 twips)
    const widthInTwips = Math.round(pageWidth * 567);
    const heightInTwips = Math.round(pageHeight * 567);
    
    const doc = new docx.Document({
        sections: [
            {
                properties: {
                    page: {
                        size: {
                            width: widthInTwips,
                            height: heightInTwips
                        },
                        margin: {
                            top: Math.round(formatSettings.marginTop * 567),
                            right: Math.round(formatSettings.marginRight * 567),
                            bottom: Math.round(formatSettings.marginBottom * 567),
                            left: Math.round(formatSettings.marginLeft * 567)
                        }
                    }
                },
                children: children
            }
        ]
    });
    
    return doc;
}

async function generatePdfDocument(customers, startRowNumber, endRowNumber) {
    // Use the exact paper dimensions as specified by the user
    const pageWidth = formatSettings.paperWidth;
    const pageHeight = formatSettings.paperHeight;
    
    // Create HTML content for PDF generation with print-specific styles
    let htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>JD Sons Customer Documents</title>
        <style>
            @page {
                size: ${pageWidth}cm ${pageHeight}cm;
                margin: 0;
            }
            @media print {
                @page {
                    size: ${pageWidth}cm ${pageHeight}cm;
                    margin: 0;
                }
                body {
                    -webkit-print-color-adjust: exact;
                    color-adjust: exact;
                    print-color-adjust: exact;
                }
            }
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            html, body {
                width: ${pageWidth}cm;
                height: ${pageHeight}cm;
                font-family: ${formatSettings.fontFamily};
                margin: 0;
                padding: 0;
                font-size: ${formatSettings.bodyFontSize}px;
                line-height: ${formatSettings.lineSpacing};
                color: ${formatSettings.textColor};
                background: white;
                overflow: hidden;
            }
            .page {
                width: ${pageWidth}cm;
                height: ${pageHeight}cm;
                padding: ${formatSettings.marginTop}cm ${formatSettings.marginRight}cm ${formatSettings.marginBottom}cm ${formatSettings.marginLeft}cm;
                page-break-after: always;
                box-sizing: border-box;
                overflow: hidden;
                display: block;
                position: relative;
            }
            .page:last-child {
                page-break-after: avoid;
            }
            .customer-header {
                font-weight: bold;
                color: ${formatSettings.headerColor};
                font-size: ${formatSettings.headerFontSize}px;
                margin-bottom: 0.6cm;
                line-height: ${formatSettings.lineSpacing};
                display: block;
            }
            .field-label {
                font-weight: bold;
                display: inline;
                font-size: ${formatSettings.bodyFontSize}px;
                color: ${formatSettings.textColor};
                line-height: ${formatSettings.lineSpacing};
            }
            .field-value {
                color: ${formatSettings.textColor};
                display: inline;
                font-size: ${formatSettings.bodyFontSize}px;
                line-height: ${formatSettings.lineSpacing};
            }
            .field {
                margin-bottom: 0.5cm;
                line-height: ${formatSettings.lineSpacing};
                display: block;
            }
            .divider {
                border-top: 3px solid #666;
                margin: 0.7cm 0;
                width: 100%;
                display: block;
            }
            .from-details {
                color: ${formatSettings.companyColor};
                font-weight: bold;
                font-size: ${formatSettings.bodyFontSize}px;
                margin-bottom: 0.4cm;
                line-height: ${formatSettings.lineSpacing};
                display: block;
            }
            .from-contact {
                color: ${formatSettings.companyColor};
                font-weight: bold;
                font-size: ${Math.round(formatSettings.bodyFontSize * 0.95)}px;
                line-height: ${formatSettings.lineSpacing};
                display: block;
            }
        </style>
    </head>
    <body>`;

    customers.forEach((customer, index) => {
        htmlContent += `
        <div class="page">
            <div class="customer-header">Customer Details:</div>
            
            <div class="field">
                <span class="field-label">Customer Name: </span>
                <span class="field-value">${customer.customerName}</span>
            </div>
            
            <div class="field">
                <span class="field-label">Address: </span>
                <span class="field-value">${customer.address}</span>
            </div>
            
            <div class="field">
                <span class="field-label">Contact Number: </span>
                <span class="field-value">${customer.contactNumber}</span>
            </div>
            
            <div class="divider"></div>
            
            <div class="from-details">From: Jemish (JD Jewellery)</div>
            <div class="from-contact">Contact Number: 9773046615</div>
        </div>`;
    });

    htmlContent += `
    </body>
    </html>`;

    // Create a temporary window for PDF generation
    const printWindow = window.open('', '_blank');
    printWindow.document.write(htmlContent);
    printWindow.document.close();
    
    // Wait for content to load, then open print dialog
    setTimeout(() => {
        printWindow.focus();
        printWindow.print();
    }, 1000);
}

function downloadDocument(blob, filename) {
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    
    URL.revokeObjectURL(url);
}

// Format customization functions
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
        formatSettings.paperWidth = paperSizes[selectedSize].width;
        formatSettings.paperHeight = paperSizes[selectedSize].height;
        updateFormat();
    }
}

function updateFormat() {
    // Update format settings from form inputs - ensure we get current values
    const paperWidthInput = document.getElementById('paperWidth');
    const paperHeightInput = document.getElementById('paperHeight');
    const marginTopInput = document.getElementById('marginTop');
    const marginBottomInput = document.getElementById('marginBottom');
    const marginLeftInput = document.getElementById('marginLeft');
    const marginRightInput = document.getElementById('marginRight');
    const fontFamilyInput = document.getElementById('fontFamily');
    const headerFontSizeInput = document.getElementById('headerFontSize');
    const bodyFontSizeInput = document.getElementById('bodyFontSize');
    const headerColorInput = document.getElementById('headerColor');
    const textColorInput = document.getElementById('textColor');
    const companyColorInput = document.getElementById('companyColor');
    const lineSpacingInput = document.getElementById('lineSpacing');
    
    // Update all format settings with current form values
    if (paperWidthInput) formatSettings.paperWidth = parseFloat(paperWidthInput.value) || 7.75;
    if (paperHeightInput) formatSettings.paperHeight = parseFloat(paperHeightInput.value) || 12.5;
    if (marginTopInput) formatSettings.marginTop = parseFloat(marginTopInput.value) || 0.2;
    if (marginBottomInput) formatSettings.marginBottom = parseFloat(marginBottomInput.value) || 0.2;
    if (marginLeftInput) formatSettings.marginLeft = parseFloat(marginLeftInput.value) || 0.2;
    if (marginRightInput) formatSettings.marginRight = parseFloat(marginRightInput.value) || 0.2;
    if (fontFamilyInput) formatSettings.fontFamily = fontFamilyInput.value || 'Arial, sans-serif';
    if (headerFontSizeInput) formatSettings.headerFontSize = parseInt(headerFontSizeInput.value) || 22;
    if (bodyFontSizeInput) formatSettings.bodyFontSize = parseInt(bodyFontSizeInput.value) || 20;
    if (headerColorInput) formatSettings.headerColor = headerColorInput.value || '#dc2626';
    if (textColorInput) formatSettings.textColor = textColorInput.value || '#1f2937';
    if (companyColorInput) formatSettings.companyColor = companyColorInput.value || '#2563eb';
    if (lineSpacingInput) formatSettings.lineSpacing = lineSpacingInput.value || '1.2';
    
    // Log current settings for debugging
    console.log('Format settings updated:', formatSettings);
    
    updatePreview();
}

function updatePreview() {
    const previewPage = document.getElementById('previewPage');
    const previewContent = previewPage.querySelector('.preview-content');
    const previewSize = document.getElementById('previewSize');
    
    if (!previewPage || !previewContent || !previewSize) {
        console.log('Preview elements not found');
        return;
    }
    
    // Use the exact paper dimensions as specified by the user - no swapping based on orientation
    const pageWidth = formatSettings.paperWidth;
    const pageHeight = formatSettings.paperHeight;
    
    // Update preview page size to match exactly
    previewPage.style.width = `${pageWidth}cm`;
    previewPage.style.height = `${pageHeight}cm`;
    
    // Update content styling to match exactly what will be generated
    previewContent.style.fontFamily = formatSettings.fontFamily;
    previewContent.style.fontSize = `${formatSettings.bodyFontSize}px`;
    previewContent.style.lineHeight = formatSettings.lineSpacing;
    previewContent.style.color = formatSettings.textColor;
    previewContent.style.padding = `${formatSettings.marginTop}cm ${formatSettings.marginRight}cm ${formatSettings.marginBottom}cm ${formatSettings.marginLeft}cm`;
    
    // Update header styling
    const header = previewContent.querySelector('.preview-header');
    if (header) {
        header.style.fontSize = `${formatSettings.headerFontSize}px`;
        header.style.color = formatSettings.headerColor;
        header.style.lineHeight = formatSettings.lineSpacing;
    }
    
    // Update field labels styling
    const labels = previewContent.querySelectorAll('.preview-label');
    labels.forEach(label => {
        label.style.fontSize = `${formatSettings.bodyFontSize}px`;
        label.style.color = formatSettings.textColor;
        label.style.lineHeight = formatSettings.lineSpacing;
    });
    
    // Update field values styling
    const values = previewContent.querySelectorAll('.preview-value');
    values.forEach(value => {
        value.style.fontSize = `${formatSettings.bodyFontSize}px`;
        value.style.color = formatSettings.textColor;
        value.style.lineHeight = formatSettings.lineSpacing;
    });
    
    // Update company info styling
    const company = previewContent.querySelector('.preview-company');
    const companyContact = previewContent.querySelector('.preview-company-contact');
    if (company) {
        company.style.color = formatSettings.companyColor;
        company.style.fontSize = `${formatSettings.bodyFontSize}px`;
        company.style.lineHeight = formatSettings.lineSpacing;
    }
    if (companyContact) {
        companyContact.style.color = formatSettings.companyColor;
        companyContact.style.fontSize = `${Math.round(formatSettings.bodyFontSize * 0.95)}px`;
        companyContact.style.lineHeight = formatSettings.lineSpacing;
    }
    
    // Update preview fields to match document structure
    const fields = previewContent.querySelectorAll('.preview-field');
    fields.forEach(field => {
        field.style.lineHeight = formatSettings.lineSpacing;
        field.style.marginBottom = '0.5cm';
    });
    
    // Update size display
    previewSize.textContent = `${pageWidth.toFixed(1)}cm × ${pageHeight.toFixed(1)}cm`;
    
    console.log(`Preview updated: ${pageWidth.toFixed(1)}cm × ${pageHeight.toFixed(1)}cm`);
}

function resetToDefaults() {
    // Reset to default values
    formatSettings = {
        paperWidth: 7.75,
        paperHeight: 12.5,
        marginTop: 0.2,
        marginBottom: 0.2,
        marginLeft: 0.2,
        marginRight: 0.2,
        fontFamily: 'Arial, sans-serif',
        headerFontSize: 22,
        bodyFontSize: 20,
        headerColor: '#dc2626',
        textColor: '#1f2937',
        companyColor: '#2563eb',
        lineSpacing: '1.2'
    };
    
    // Update form inputs
    document.getElementById('paperSizeSelect').value = 'receipt';
    document.getElementById('paperWidth').value = formatSettings.paperWidth;
    document.getElementById('paperHeight').value = formatSettings.paperHeight;
    document.getElementById('marginTop').value = formatSettings.marginTop;
    document.getElementById('marginBottom').value = formatSettings.marginBottom;
    document.getElementById('marginLeft').value = formatSettings.marginLeft;
    document.getElementById('marginRight').value = formatSettings.marginRight;
    document.getElementById('fontFamily').value = formatSettings.fontFamily;
    document.getElementById('headerFontSize').value = formatSettings.headerFontSize;
    document.getElementById('bodyFontSize').value = formatSettings.bodyFontSize;
    document.getElementById('headerColor').value = formatSettings.headerColor;
    document.getElementById('textColor').value = formatSettings.textColor;
    document.getElementById('companyColor').value = formatSettings.companyColor;
    document.getElementById('lineSpacing').value = formatSettings.lineSpacing;
    
    updatePreview();
    
    // Show confirmation
    alert('✅ Format settings have been reset to defaults!');
}

function saveFormatSettings() {
    try {
        // Make sure we have the latest format settings before saving
        updateFormat();
        localStorage.setItem('jdSonsFormatSettings', JSON.stringify(formatSettings));
        alert('✅ Format settings have been saved successfully!');
    } catch (error) {
        console.error('Error saving format settings:', error);
        alert('❌ Error saving format settings. Please try again.');
    }
}

function loadFormatSettings() {
    try {
        const saved = localStorage.getItem('jdSonsFormatSettings');
        if (saved) {
            const savedSettings = JSON.parse(saved);
            // Merge saved settings with current formatSettings to ensure all properties exist
            formatSettings = { ...formatSettings, ...savedSettings };
            
            // Update form inputs with loaded settings
            if (document.getElementById('paperWidth')) {
                document.getElementById('paperWidth').value = formatSettings.paperWidth;
            }
            if (document.getElementById('paperHeight')) {
                document.getElementById('paperHeight').value = formatSettings.paperHeight;
            }
            if (document.getElementById('marginTop')) {
                document.getElementById('marginTop').value = formatSettings.marginTop;
            }
            if (document.getElementById('marginBottom')) {
                document.getElementById('marginBottom').value = formatSettings.marginBottom;
            }
            if (document.getElementById('marginLeft')) {
                document.getElementById('marginLeft').value = formatSettings.marginLeft;
            }
            if (document.getElementById('marginRight')) {
                document.getElementById('marginRight').value = formatSettings.marginRight;
            }
            if (document.getElementById('fontFamily')) {
                document.getElementById('fontFamily').value = formatSettings.fontFamily;
            }
            if (document.getElementById('headerFontSize')) {
                document.getElementById('headerFontSize').value = formatSettings.headerFontSize;
            }
            if (document.getElementById('bodyFontSize')) {
                document.getElementById('bodyFontSize').value = formatSettings.bodyFontSize;
            }
            if (document.getElementById('headerColor')) {
                document.getElementById('headerColor').value = formatSettings.headerColor;
            }
            if (document.getElementById('textColor')) {
                document.getElementById('textColor').value = formatSettings.textColor;
            }
            if (document.getElementById('companyColor')) {
                document.getElementById('companyColor').value = formatSettings.companyColor;
            }
            if (document.getElementById('lineSpacing')) {
                document.getElementById('lineSpacing').value = formatSettings.lineSpacing;
            }
            
            // Set paper size selector based on current dimensions
            const paperSizeSelect = document.getElementById('paperSizeSelect');
            if (paperSizeSelect) {
                let matchedSize = 'custom';
                for (const [key, size] of Object.entries(paperSizes)) {
                    if (Math.abs(size.width - formatSettings.paperWidth) < 0.1 && 
                        Math.abs(size.height - formatSettings.paperHeight) < 0.1) {
                        matchedSize = key;
                        break;
                    }
                }
                paperSizeSelect.value = matchedSize;
            }
            
            console.log('Format settings loaded successfully:', formatSettings);
        }
    } catch (error) {
        console.error('Error loading format settings:', error);
    }
    
    // Update preview with loaded or default settings
    updatePreview();
} 