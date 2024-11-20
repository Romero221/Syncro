// preload.js

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectFile: (mode) => ipcRenderer.invoke('select-file', 'open'),
    startProcessing: (args) => ipcRenderer.send('start-processing', args),
    syncMondayToExcel: (args) => ipcRenderer.send('sync-monday-to-excel', args),
    onProcessingResult: (callback) => ipcRenderer.on('processing-result', callback),
    onSyncResult: (callback) => ipcRenderer.on('sync-result', callback),
    onLogMessage: (callback) => ipcRenderer.on('log-message', callback),
});
