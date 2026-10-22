const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktop', Object.freeze({
  getSettings: () => ipcRenderer.invoke('desktop:get-settings'),
  saveServerUrl: serverUrl => ipcRenderer.invoke('desktop:save-server', serverUrl),
}));
