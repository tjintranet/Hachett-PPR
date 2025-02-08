// Global variable to store processed data
let processedData = [];

document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('excelFile');
    fileInput.addEventListener('change', handleFileSelect);
});

// Handle file selection
async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        const data = await readExcelFile(file);
        console.log("Raw Excel data:", data);  // Debug log
        processData(data);
        updatePreviewTable();
        updateButtonStates();
        showStatus('File processed successfully', 'success');
    } catch (error) {
        showStatus(`Error processing file: ${error.message}`, 'danger');
        console.error(error);
    }
}

// Read Excel file
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    cellDates: true,
                    cellNF: true,
                    cellText: false
                });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // Convert to array format with all rows
                const rows = XLSX.utils.sheet_to_json(firstSheet, {
                    header: 1,
                    defval: '',
                    raw: true
                });
                
                console.log("Parsed Excel rows:", rows);  // Debug log
                resolve(rows);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process Excel data
function processData(data) {
    if (!data || data.length < 2) return;

    // Process all rows except header
    processedData = data.slice(1)  // Skip header row
        .filter(row => row.length > 0 && row.some(cell => cell))  // Remove empty rows
        .map((row, index) => {
            console.log("Processing row:", row);  // Debug log
            return {
                reference: row[0]?.toString() || '',
                lineNumber: String(index + 1).padStart(5, '0'),
                isbn: row[2]?.toString() || '',  // ISBN from column C
                status: 'AR',
                result: 'OK'
            };
        });
    
    console.log("Processed data:", processedData);  // Debug log
}

// Update preview table
function updatePreviewTable() {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';
    
    if (!processedData || processedData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center">No data loaded</td></tr>';
        return;
    }
    
    processedData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="checkbox" class="row-checkbox"></td>
            <td>
                <button class="btn btn-sm btn-danger" onclick="deleteRow(this)">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
            <td>${row.reference}</td>
            <td>${row.lineNumber}</td>
            <td>${row.isbn}</td>
            <td>${row.status}</td>
            <td>${row.result}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Download CSV file
function downloadCsv() {
    if (processedData.length === 0) {
        showStatus('No data to download', 'warning');
        return;
    }
    
    const orderNumber = processedData[0].reference;
    const csvRows = processedData.map(row => 
        [
            row.reference,
            row.lineNumber,
            row.isbn,
            row.status,
            row.result
        ].join(',')
    );
    
    const csvContent = csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    
    a.href = url;
    a.download = `T1.M${orderNumber}.PPR`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showStatus('PPR file downloaded successfully', 'success');
}

// Clear all data
function clearAll() {
    processedData = [];
    document.getElementById('excelFile').value = '';
    updatePreviewTable();
    updateButtonStates();
    showStatus('All data cleared', 'info');
}

// Update button states
function updateButtonStates() {
    const hasData = processedData && processedData.length > 0;
    document.getElementById('clearBtn').disabled = !hasData;
    document.getElementById('downloadBtn').disabled = !hasData;
    document.getElementById('deleteSelectedBtn').disabled = !hasData;
}

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = `alert alert-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 3000);
}

// Toggle all checkboxes
function toggleAllCheckboxes() {
    const selectAll = document.getElementById('selectAll');
    document.querySelectorAll('.row-checkbox').forEach(checkbox => {
        checkbox.checked = selectAll.checked;
    });
}

// Delete selected rows
function deleteSelected() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    if (checkboxes.length === 0) {
        showStatus('No rows selected', 'warning');
        return;
    }
    
    const rowsToDelete = Array.from(checkboxes).map(checkbox => 
        checkbox.closest('tr'));
    
    rowsToDelete.forEach(row => {
        const index = Array.from(row.parentNode.children).indexOf(row);
        processedData.splice(index, 1);
        row.remove();
    });
    
    // Update line numbers
    processedData = processedData.map((row, index) => ({
        ...row,
        lineNumber: String(index + 1).padStart(5, '0')
    }));
    
    updateButtonStates();
    showStatus(`Deleted ${checkboxes.length} row(s)`, 'success');
}

// Delete single row
function deleteRow(button) {
    const row = button.closest('tr');
    const rowIndex = Array.from(row.parentNode.children).indexOf(row);
    
    processedData.splice(rowIndex, 1);
    
    // Update line numbers
    processedData = processedData.map((row, index) => ({
        ...row,
        lineNumber: String(index + 1).padStart(5, '0')
    }));
    
    updatePreviewTable();
    updateButtonStates();
    showStatus('Row deleted', 'success');
}