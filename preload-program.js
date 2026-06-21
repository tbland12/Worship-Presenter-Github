const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  hideProgram: () => ipcRenderer.send('program:hide'),
  sendProgramEvent: (payload) => ipcRenderer.send('program:event', payload),
  onProgramState: (callback) => ipcRenderer.on('program:state', (_event, state) => callback(state))
});
