// renderer.js

const boardIdInput = document.getElementById('boardId');
const apiKeyInput = document.getElementById('apiKey');
const filePathInput = document.getElementById('filePath');
const selectFileButton = document.getElementById('selectFile');
const startButton = document.getElementById('start');
const logDiv = document.getElementById('log');
const modeSwitch = document.getElementById('modeSwitch');
const modeLabel = document.getElementById('modeLabel');

modeSwitch.addEventListener('change', () => {
    if (modeSwitch.checked) {
        // Switch is ON - Sync Monday to Excel
        modeLabel.textContent = 'Sync Monday to Excel';
        filePathInput.placeholder = 'Select existing Excel file to update';
        selectFileButton.textContent = 'Select Excel File';
    } else {
        // Switch is OFF - Sync Excel to Monday
        modeLabel.textContent = 'Sync Excel to Monday';
        filePathInput.placeholder = 'Excel File Path';
        selectFileButton.textContent = 'Select File';
    }
});

selectFileButton.addEventListener('click', async () => {
    const mode = 'open'; // Always open an existing file
    const filePath = await window.electronAPI.selectFile(mode);
    if (filePath) {
        filePathInput.value = filePath;
    }
});

startButton.addEventListener('click', () => {
    const boardId = boardIdInput.value.trim();
    const apiKey = apiKeyInput.value.trim();
    const filePath = filePathInput.value.trim();

    if (!boardId || !apiKey) {
        alert('Please fill in the Board ID and API Key.');
        return;
    }

    if (!filePath) {
        alert('Please select a file or save location.');
        return;
    }

    logDiv.textContent = 'Processing...\n';

    if (modeSwitch.checked) {
        // Sync Monday to Excel
        window.electronAPI.syncMondayToExcel({ boardId, apiKey, filePath });
    } else {
        // Sync Excel to Monday
        window.electronAPI.startProcessing({ boardId, apiKey, filePath });
    }
});

window.electronAPI.onProcessingResult((event, result) => {
    if (result.success) {
        logDiv.textContent += 'Processing completed successfully.\n';
    } else {
        logDiv.textContent += `Error: ${result.message}\n`;
    }
    logDiv.scrollTop = logDiv.scrollHeight; // Auto-scroll to the bottom
});

window.electronAPI.onSyncResult((event, result) => {
    if (result.success) {
        logDiv.textContent += 'Processing completed successfully.\n';
    } else {
        logDiv.textContent += `Error: ${result.message}\n`;
    }
    logDiv.scrollTop = logDiv.scrollHeight; // Auto-scroll to the bottom
});

window.electronAPI.onLogMessage((event, message) => {
    logDiv.textContent += message + '\n';
    logDiv.scrollTop = logDiv.scrollHeight; // Auto-scroll to the bottom
});
