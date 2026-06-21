const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  importContentPack: () => ipcRenderer.invoke('content-pack:import'),
  closeLibrary: () => ipcRenderer.send('library:close'),
  listLibraryItems: (options = {}) => ipcRenderer.invoke('library:list', options),
  selectLibraryItem: (payload) => ipcRenderer.send('library:select', payload),
  onLibraryScope: (callback) => ipcRenderer.on('library:scope', (_event, scope) => callback(scope))
});
