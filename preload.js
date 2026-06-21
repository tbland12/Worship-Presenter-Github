const { contextBridge, ipcRenderer, webUtils } = require('electron');

function mediaUrl(host, value, prefix) {
  const cleaned = String(value || '').replace(prefix, '').replace(/\\/g, '/');
  const encoded = cleaned.split('/').filter(Boolean).map(encodeURIComponent).join('/');
  return encoded ? `worship-media://${host}/${encoded}` : '';
}

contextBridge.exposeInMainWorld('api', {
  resetSession: () => ipcRenderer.invoke('session:reset'),
  openProjectFile: () => ipcRenderer.invoke('project:open-file'),
  openRecentProject: (id) => ipcRenderer.invoke('project:open-recent', id),
  listRecentProjects: () => ipcRenderer.invoke('project:list-recent'),
  saveProjectFile: (project) => ipcRenderer.invoke('project:save-file', project),
  writeRecovery: (project) => ipcRenderer.invoke('project:write-recovery', project),
  clearRecovery: () => ipcRenderer.invoke('project:clear-recovery'),
  checkRecovery: () => ipcRenderer.invoke('project:check-recovery'),
  listSongLibrary: (query = '') => ipcRenderer.invoke('song-library:list', query),
  saveSongToLibrary: (song) => ipcRenderer.invoke('song-library:save', { song }),
  instantiateLibrarySong: (id) => ipcRenderer.invoke('song-library:instantiate', id),
  removeLibrarySong: (id) => ipcRenderer.invoke('song-library:remove', id),
  confirmProjectTransition: () => ipcRenderer.invoke('project:confirm-transition'),
  setProjectDirty: (dirty) => ipcRenderer.send('project:set-dirty', dirty === true),
  completeProjectClose: (saved) => ipcRenderer.send('project:complete-close', saved === true),
  onSaveBeforeClose: (callback) => ipcRenderer.on('project:save-before-close', () => callback()),
  pickLyrics: () => ipcRenderer.invoke('lyrics:pick'),
  readTextFile: (token) => ipcRenderer.invoke('file:read-text', token),
  editCommand: (command) => ipcRenderer.send('edit:command', command),
  readPptxSlides: (token) => ipcRenderer.invoke('pptx:read-slides', token),
  pickMedia: () => ipcRenderer.invoke('media:pick'),
  importLibrary: (payload) => ipcRenderer.invoke('library:import', payload),
  importAnnouncement: (token) => ipcRenderer.invoke('announcement:import', token),
  pickAnnouncement: () => ipcRenderer.invoke('announcement:pick'),
  openLibrary: (options = {}) => ipcRenderer.send('library:open', options),
  onLibrarySelected: (callback) => ipcRenderer.on('library:selected', (_event, payload) => callback(payload)),
  checkAssets: (payload) => ipcRenderer.invoke('assets:check', payload),
  onMenuAction: (callback) => ipcRenderer.on('menu:action', (_event, action) => callback(action)),
  registerDroppedFile: (file) => {
    try {
      const filePath = webUtils.getPathForFile(file);
      return filePath ? ipcRenderer.invoke('file:register-drop', filePath) : Promise.resolve(null);
    } catch (error) {
      return Promise.resolve(null);
    }
  },
  listDisplays: () => ipcRenderer.invoke('display:list'),
  showProgram: (displayId) => ipcRenderer.send('program:show', displayId),
  hideProgram: () => ipcRenderer.send('program:hide'),
  setProgramDisplay: (displayId) => ipcRenderer.send('program:set-display', displayId),
  toggleProgramFullscreen: () => ipcRenderer.send('program:toggle-fullscreen'),
  sendProgramState: (state) => ipcRenderer.send('program:state', state),
  onProgramEvent: (callback) => ipcRenderer.on('program:event', (_event, payload) => callback(payload)),
  onDisplayChanged: (callback) => ipcRenderer.on('display:changed', (_event, payload) => callback(payload)),
  getStageConfig: () => ipcRenderer.invoke('stage:get-config'),
  showStage: (displayId) => ipcRenderer.send('stage:show', displayId),
  hideStage: () => ipcRenderer.send('stage:hide'),
  setStageDisplay: (displayId) => ipcRenderer.send('stage:set-display', displayId),
  sendStageState: (state) => ipcRenderer.send('stage:state', state),
  onStageEvent: (callback) => ipcRenderer.on('stage:event', (_event, payload) => callback(payload)),
  resolveLibraryUrl: (assetPath) => mediaUrl('library', assetPath, /^library[\\/]/),
  resolveSessionUrl: (assetPath) => mediaUrl('asset', assetPath, /^session\//),
  getAppVersion: () => ipcRenderer.invoke('app:get-version')
});
