const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
  // Get the current AD domain
  getDomain: () => ipcRenderer.invoke('get-domain'),

  // Generate the AD report
  generateReport: () => ipcRenderer.invoke('generate-report'),

  // Open the generated report
  openReport: (reportPath) => ipcRenderer.invoke('open-report', reportPath),

  // Listen for script output
  onScriptOutput: (callback) => {
    ipcRenderer.on('script-output', (event, data) => callback(data));
  }
});
