// renderer.js

const boardIdInput = document.getElementById('boardId');
const apiKeyInput = document.getElementById('apiKey');
const filePathInput = document.getElementById('filePath');
const selectFileButton = document.getElementById('selectFile');
const startButton = document.getElementById('start');
const logDiv = document.getElementById('log');

selectFileButton.addEventListener('click', async () => {
    const filePath = await window.electronAPI.selectFile();
    if (filePath) {
        filePathInput.value = filePath;
    }
});

startButton.addEventListener('click', () => {
    const boardId = boardIdInput.value.trim();
    const apiKey = apiKeyInput.value.trim();
    const filePath = filePathInput.value.trim();

    if (!boardId || !apiKey || !filePath) {
        alert('Please fill in all fields.');
        return;
    }

    logDiv.textContent = 'Processing...\n';
    window.electronAPI.startProcessing({ boardId, apiKey, filePath });
});

window.electronAPI.onProcessingResult((event, result) => {
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
