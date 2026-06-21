const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  hideStage: () => ipcRenderer.send('stage:hide'),
  onStageState: (callback) => ipcRenderer.on('stage:state', (_event, state) => callback(state))
});
