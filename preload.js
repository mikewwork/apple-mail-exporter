const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getMailAccounts: () => ipcRenderer.invoke('get-mail-accounts'),
  selectOutputFolder: () => ipcRenderer.invoke('select-output-folder'),
  exportEmails: (params) => ipcRenderer.invoke('export-emails', params),
  openFolder: (path) => ipcRenderer.invoke('openFolder', path),
  openPDF: (path) => ipcRenderer.invoke('openPDF', path),
  
  // Progress updates
  onProgress: (callback) => {
    ipcRenderer.on('export-progress', callback);
  },
  
  // Remove listeners
  removeAllListeners: (channel) => {
    ipcRenderer.removeAllListeners(channel);
  }
}); 