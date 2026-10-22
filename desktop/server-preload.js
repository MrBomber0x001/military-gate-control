const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('gateServer', Object.freeze({
  getStatus: () => ipcRenderer.invoke('server:get-status'),
  openAddress: address => ipcRenderer.invoke('server:open-address', address),
  retryPort: port => ipcRenderer.invoke('server:retry-port', port),
  onStatus: callback => {
    const listener = (_event, state) => callback(state);
    ipcRenderer.on('server:status', listener);
    return () => ipcRenderer.removeListener('server:status', listener);
  },
}));
