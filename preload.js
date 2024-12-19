// preload.js

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectFile: (mode) => ipcRenderer.invoke('select-file', mode),
    startProcessing: (filePaths) => ipcRenderer.send('start-processing', filePaths),
    syncMondayToExcel: (args) => ipcRenderer.send('sync-monday-to-excel', args),
    onProcessingResult: (callback) => ipcRenderer.on('processing-result', callback),
    onSyncResult: (callback) => ipcRenderer.on('sync-result', callback),
    onLogMessage: (callback) => ipcRenderer.on('log-message', callback),
    runDocuparse: (pdfPaths) => ipcRenderer.send('run-docuparse', pdfPaths),
});
