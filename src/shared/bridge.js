export function ensureApiBridge() {
  if (window.api) {
    return window.api;
  }

  if (!window.require) {
    return null;
  }

  try {
    const { ipcRenderer } = window.require('electron');
    const path = window.require('path');
    const { pathToFileURL } = window.require('url');

    function resolveMediaUrl(projectFolder, relativePath) {
      if (!projectFolder || !relativePath) {
        return '';
      }
      const absolutePath = path.join(projectFolder, relativePath);
      return pathToFileURL(absolutePath).toString();
    }

    const api = {
      newProject: (options = {}) => ipcRenderer.invoke('project:new', options),
      openProject: () => ipcRenderer.invoke('project:open'),
      loadProject: (folderPath) => ipcRenderer.invoke('project:load', folderPath),
      saveProject: (folderPath, project) => ipcRenderer.invoke('project:save', folderPath, project),
      openProjectFile: () => ipcRenderer.invoke('project:open-file'),
      saveProjectFile: (filePath, project, options = {}) => ipcRenderer.invoke('project:save-file', filePath, project, options),
      pickLyrics: () => ipcRenderer.invoke('lyrics:pick'),
      readTextFile: (filePath) => ipcRenderer.invoke('file:read-text', filePath),
      pickMedia: () => ipcRenderer.invoke('media:pick'),
      importMedia: (projectFolder, sourcePath) => ipcRenderer.invoke('media:import', projectFolder, sourcePath),
      importLibrary: (sourcePath) => ipcRenderer.invoke('library:import', sourcePath),
      importAnnouncement: (sourcePath) => ipcRenderer.invoke('announcement:import', sourcePath),
      pickAnnouncement: () => ipcRenderer.invoke('announcement:pick'),
      openLibrary: (options = {}) => ipcRenderer.send('library:open', options),
      closeLibrary: () => ipcRenderer.send('library:close'),
      listLibraryItems: (options = {}) => ipcRenderer.invoke('library:list', options),
      selectLibraryItem: (payload) => ipcRenderer.send('library:select', payload),
      onLibrarySelected: (callback) => ipcRenderer.on('library:selected', (_event, payload) => callback(payload)),
      onLibraryScope: (callback) => ipcRenderer.on('library:scope', (_event, scope) => callback(scope)),
      checkAssets: (payload) => ipcRenderer.invoke('assets:check', payload),
      onMenuAction: (callback) => ipcRenderer.on('menu:action', (_event, action) => callback(action)),
      resolveLibraryUrl: (relativePath) => {
        if (!relativePath) {
          return '';
        }
        const cleaned = relativePath.replace(/^library[\\/]/, '');
        const absolutePath = path.join(__dirname, 'library', cleaned);
        return pathToFileURL(absolutePath).toString();
      },
      listDisplays: () => ipcRenderer.invoke('display:list'),
      showProgram: (displayId) => ipcRenderer.send('program:show', displayId),
      hideProgram: () => ipcRenderer.send('program:hide'),
      setProgramDisplay: (displayId) => ipcRenderer.send('program:set-display', displayId),
      toggleProgramFullscreen: () => ipcRenderer.send('program:toggle-fullscreen'),
      sendProgramState: (state) => ipcRenderer.send('program:state', state),
      sendProgramEvent: (payload) => ipcRenderer.send('program:event', payload),
      onProgramState: (callback) => ipcRenderer.on('program:state', (_event, state) => callback(state)),
      onProgramEvent: (callback) => ipcRenderer.on('program:event', (_event, payload) => callback(payload)),
      resolveMediaUrl
    };

    window.api = api;
    return api;
  } catch (error) {
    console.error('Failed to build renderer bridge', error);
    return null;
  }
}
