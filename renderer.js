const { ipcRenderer } = require('electron');

// Elements
const selectExcelButton = document.getElementById('select-excel');
const excelFileInput = document.getElementById('excel-file');
const syncExcelToMondayButton = document.getElementById('sync-excel-to-monday');
const syncMondayToExcelButton = document.getElementById('sync-monday-to-excel');
const boardIdInput = document.getElementById('board-id');
const apiKeyInput = document.getElementById('api-key');
const logOutput = document.getElementById('log-output');

// Helper function to log messages
function logMessage(message) {
    const logEntry = document.createElement('div');
    logEntry.textContent = message;
    logOutput.appendChild(logEntry);
    logOutput.scrollTop = logOutput.scrollHeight;
}

// Select Excel file
selectExcelButton.addEventListener('click', async () => {
    const filePath = await ipcRenderer.invoke('select-file', 'open');
    if (filePath) {
        excelFileInput.value = filePath;
        logMessage(`Selected Excel file: ${filePath}`);
    } else {
        logMessage('No file selected.');
    }
});

// Sync Excel to Monday
syncExcelToMondayButton.addEventListener('click', () => {
    const filePath = excelFileInput.value.trim();
    const boardId = boardIdInput.value.trim();
    const apiKey = apiKeyInput.value.trim();

    if (!filePath || !boardId || !apiKey) {
        logMessage('Please provide all required inputs: Excel file, Board ID, and API Key.');
        return;
    }

    logMessage('Starting sync from Excel to Monday...');
    ipcRenderer.send('start-processing', { filePath, boardId, apiKey });
});

// Sync Monday to Excel
syncMondayToExcelButton.addEventListener('click', () => {
    const filePath = excelFileInput.value.trim();
    const boardId = boardIdInput.value.trim();
    const apiKey = apiKeyInput.value.trim();

    if (!filePath || !boardId || !apiKey) {
        logMessage('Please provide all required inputs: Excel file, Board ID, and API Key.');
        return;
    }

    logMessage('Starting sync from Monday to Excel...');
    ipcRenderer.send('sync-monday-to-excel', { filePath, boardId, apiKey });
});

// Listen for results
ipcRenderer.on('processing-result', (event, result) => {
    if (result.success) {
        logMessage('Sync from Excel to Monday completed successfully.');
    } else {
        logMessage(`Error during sync from Excel to Monday: ${result.message}`);
    }
});

ipcRenderer.on('sync-result', (event, result) => {
    if (result.success) {
        logMessage('Sync from Monday to Excel completed successfully.');
    } else {
        logMessage(`Error during sync from Monday to Excel: ${result.message}`);
    }
});

// Listen for log messages from the main process
ipcRenderer.on('log-message', (event, message) => {
    logMessage(message);
});
