// preload.js

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectFile: () => ipcRenderer.invoke('select-file'),
    startProcessing: (args) => ipcRenderer.send('start-processing', args),
    onProcessingResult: (callback) => ipcRenderer.on('processing-result', callback),
});
