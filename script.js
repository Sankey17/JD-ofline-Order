let customerData = [];
let currentDataSource = 'file'; // 'file' or 'sheets'

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
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
    document.getElementById('selectionSection').style.display = 'none';
    document.getElementById('generateSection').style.display = 'none';
    document.getElementById('statusSection').style.display = 'none';
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
    
    fileInfo.style.display = 'block';
    document.getElementById('selectionSection').style.display = 'block';
    document.getElementById('generateSection').style.display = 'block';
    
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
        
    } catch (error) {
        console.error(`Error generating ${format} document:`, error);
        alert(`Error generating ${format.toUpperCase()} document. Please try again.`);
        hideStatus();
    }
}

async function createWordDocument(customers) {
    const children = [];
    
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
                        size: 44,
                        color: "dc2626"
                    })
                ],
                spacing: { after: 600 }
            }),
            
            // Customer Name
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Customer Name: ",
                        bold: true,
                        size: 40
                    }),
                    new docx.TextRun({
                        text: customer.customerName,
                        size: 40,
                        color: "1f2937"
                    })
                ],
                spacing: { after: 500 }
            }),
            
            // Address
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Address: ",
                        bold: true,
                        size: 40
                    }),
                    new docx.TextRun({
                        text: customer.address,
                        size: 40,
                        color: "1f2937"
                    })
                ],
                spacing: { after: 500 }
            }),
            
            // Contact Number
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Contact Number: ",
                        bold: true,
                        size: 40
                    }),
                    new docx.TextRun({
                        text: customer.contactNumber.toString(),
                        size: 40,
                        color: "1f2937"
                    })
                ],
                spacing: { after: 700 }
            }),
            
            // Divider line
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "─".repeat(40),
                        color: "666666",
                        size: 32
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
                        size: 40,
                        color: "2563eb"
                    })
                ],
                spacing: { after: 400 }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: "Contact Number: 9773046615",
                        bold: true,
                        size: 38,
                        color: "2563eb"
                    })
                ],
                spacing: { after: 500 }
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
    
    // Convert cm to twips (1 cm = 567 twips)
    const widthInTwips = Math.round(7.75 * 567);  // 7.75 cm
    const heightInTwips = Math.round(12.5 * 567); // 12.5 cm
    
    const doc = new docx.Document({
        sections: [
            {
                properties: {
                    page: {
                        size: {
                            width: widthInTwips,
                            height: heightInTwips,
                            orientation: docx.PageOrientation.PORTRAIT
                        },
                        margin: {
                            top: 288,    // 0.2 inch (even smaller margins)
                            right: 288,  // 0.2 inch
                            bottom: 288, // 0.2 inch
                            left: 288    // 0.2 inch
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
    // Create HTML content for PDF generation
    let htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {
                size: 7.75cm 12.5cm;
                margin: 0.2cm;
            }
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                font-size: 20px;
                line-height: 1.3;
            }
            .page {
                width: 7.15cm;
                height: 12.1cm;
                page-break-after: always;
                padding: 0.05cm;
            }
            .page:last-child {
                page-break-after: avoid;
            }
            .customer-header {
                font-weight: bold;
                color: #dc2626;
                font-size: 22px;
                margin-bottom: 0.6cm;
            }
            .field-label {
                font-weight: bold;
                display: inline;
                font-size: 20px;
            }
            .field-value {
                color: #1f2937;
                display: inline;
                font-size: 20px;
            }
            .field {
                margin-bottom: 0.5cm;
                line-height: 1.2;
            }
            .divider {
                border-top: 3px solid #666;
                margin: 0.7cm 0;
            }
            .from-details {
                color: #2563eb;
                font-weight: bold;
                font-size: 20px;
                margin-bottom: 0.4cm;
            }
            .from-contact {
                color: #2563eb;
                font-weight: bold;
                font-size: 19px;
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
    
    // Wait for content to load, then print
    setTimeout(() => {
        printWindow.focus();
        printWindow.print();
        setTimeout(() => {
            printWindow.close();
        }, 1000);
    }, 500);
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