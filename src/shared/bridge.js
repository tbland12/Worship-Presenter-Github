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
    const fssync = window.require('fs');
    const { pathToFileURL } = window.require('url');
    function resolveBundledLibraryRoot() {
      const candidates = [];
      if (process.resourcesPath) {
        candidates.push(path.join(process.resourcesPath, 'library'));
      }
      candidates.push(path.join(__dirname, 'library'));
      for (const candidate of candidates) {
        if (candidate && fssync.existsSync(candidate)) {
          return candidate;
        }
      }
      return candidates[0] || '';
    }

    let libraryRoot = '';
    let bundledLibraryRoot = resolveBundledLibraryRoot();
    let libraryRootsPromise = null;

    async function refreshLibraryRoots() {
      if (libraryRootsPromise) {
        return libraryRootsPromise;
      }
      libraryRootsPromise = (async () => {
        try {
          const roots = await ipcRenderer.invoke('library:roots');
          if (roots && typeof roots === 'object') {
            libraryRoot = roots.libraryRoot || libraryRoot;
            bundledLibraryRoot = roots.bundledLibraryRoot || bundledLibraryRoot;
          }
        } catch (error) {
          // Ignore and use fallbacks.
          libraryRootsPromise = null;
        }
        return { libraryRoot, bundledLibraryRoot };
      })();
      return libraryRootsPromise;
    }

    refreshLibraryRoots();

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
      editCommand: (command) => ipcRenderer.send('edit:command', command),
      pickMedia: () => ipcRenderer.invoke('media:pick'),
      importMedia: (projectFolder, sourcePath) => ipcRenderer.invoke('media:import', projectFolder, sourcePath),
      importLibrary: (sourcePath) => ipcRenderer.invoke('library:import', sourcePath),
      importContentPack: () => ipcRenderer.invoke('content-pack:import'),
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
      ensureLibraryRoots: () => refreshLibraryRoots(),
      resolveLibraryUrl: (relativePath) => {
        if (!relativePath) {
          return '';
        }
        const cleaned = relativePath.replace(/^library[\\/]/, '');
        const basePath = libraryRoot || bundledLibraryRoot || path.join(__dirname, 'library');
        const absolutePath = path.join(basePath, cleaned);
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
      onDisplayChanged: (callback) => ipcRenderer.on('display:changed', (_event, payload) => callback(payload)),
      resolveMediaUrl,
      getAppVersion: () => ipcRenderer.invoke('app:get-version')
    };

    window.api = api;
    return api;
  } catch (error) {
    console.error('Failed to build renderer bridge', error);
    return null;
  }
}
