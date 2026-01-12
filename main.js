const { app, BrowserWindow, dialog, ipcMain, screen, Menu, autoUpdater } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const fssync = require('fs');
const { pathToFileURL, fileURLToPath } = require('url');

if (typeof BrowserWindow.prototype.isReadyToShow !== 'function') {
  BrowserWindow.prototype.isReadyToShow = function isReadyToShow() {
    if (this.isDestroyed()) {
      return false;
    }
    if (!this.webContents) {
      return false;
    }
    return !this.webContents.isLoading();
  };
}

let mainWindow;
let programWindow;
let libraryWindow;
let libraryWindowScope = 'background';
let updateCheckInProgress = false;
let updateCheckRequested = false;

const UPDATE_REPOSITORY = 'tbland12/Worship-Presenter-Github';

const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']);
const VIDEO_EXTS = new Set(['.mp4', '.mov', '.mkv', '.avi', '.webm']);
const SESSION_VERSION = 2;

function logError(message, error) {
  console.error(message, error);
}

function buildTimestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

async function createDefaultProjectFolder() {
  const baseDir = path.join(app.getPath('documents'), 'WorshipPresenter Projects');
  await fs.mkdir(baseDir, { recursive: true });
  const folderName = `Untitled-${buildTimestamp()}`;
  const projectDir = path.join(baseDir, folderName);
  await fs.mkdir(projectDir, { recursive: true });
  return projectDir;
}

function getDialogParent() {
  const focused = BrowserWindow.getFocusedWindow();
  if (focused && !focused.isDestroyed()) {
    return focused;
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    return mainWindow;
  }
  return null;
}

async function showOpenDialogSafe(options) {
  const parent = getDialogParent();
  if (parent) {
    try {
      return await dialog.showOpenDialog(parent, options);
    } catch (error) {
      return await dialog.showOpenDialog(options);
    }
  }
  return await dialog.showOpenDialog(options);
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1200,
    minHeight: 720,
    backgroundColor: '#101418',
    webPreferences: {
      autoplayPolicy: 'no-user-gesture-required',
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'operator.html'));
  mainWindow.on('closed', () => {
    if (programWindow && !programWindow.isDestroyed()) {
      programWindow.close();
    }
    if (libraryWindow && !libraryWindow.isDestroyed()) {
      libraryWindow.close();
    }
    mainWindow = null;
  });
}

function sendMenuAction(action) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('menu:action', action);
  }
}

function createAppMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'New Project',
          accelerator: 'Ctrl+N',
          click: () => sendMenuAction('new-project')
        },
        {
          label: 'Open Project',
          accelerator: 'Ctrl+O',
          click: () => sendMenuAction('open-project')
        },
        {
          label: 'Save',
          accelerator: 'Ctrl+S',
          click: () => sendMenuAction('save-project')
        },
        {
          label: 'Check for Updates',
          click: () => checkForUpdates(true)
        },
        { type: 'separator' },
        { role: 'quit' }
      ]
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
    { role: 'help', submenu: [] }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function getUpdateFeedUrl() {
  const platform = process.platform;
  const arch = process.arch;
  return `https://update.electronjs.org/${UPDATE_REPOSITORY}/${platform}-${arch}/${app.getVersion()}`;
}

function canCheckForUpdates() {
  return app.isPackaged && process.platform === 'win32';
}

function checkForUpdates(interactive = false) {
  if (!canCheckForUpdates()) {
    if (interactive) {
      dialog.showMessageBox(getDialogParent() || undefined, {
        type: 'info',
        message: 'Updates are available after the app is packaged.',
        detail: 'Build and install the app to enable update checks.'
      });
    }
    return;
  }
  if (updateCheckInProgress) {
    return;
  }
  updateCheckInProgress = true;
  updateCheckRequested = interactive;
  try {
    autoUpdater.setFeedURL({ url: getUpdateFeedUrl() });
    autoUpdater.checkForUpdates();
  } catch (error) {
    updateCheckInProgress = false;
    updateCheckRequested = false;
    logError('Failed to check for updates', error);
    if (interactive) {
      dialog.showErrorBox('Update Error', String(error?.message || error));
    }
  }
}

function initAutoUpdater() {
  if (!canCheckForUpdates()) {
    return;
  }
  autoUpdater.on('update-available', () => {
    if (updateCheckRequested) {
      dialog.showMessageBox(getDialogParent() || undefined, {
        type: 'info',
        message: 'Update available',
        detail: 'The update is downloading in the background.'
      });
    }
  });
  autoUpdater.on('update-not-available', () => {
    if (updateCheckRequested) {
      dialog.showMessageBox(getDialogParent() || undefined, {
        type: 'info',
        message: 'You are up to date.'
      });
    }
    updateCheckInProgress = false;
    updateCheckRequested = false;
  });
  autoUpdater.on('error', (error) => {
    logError('Auto updater error', error);
    if (updateCheckRequested) {
      dialog.showErrorBox('Update Error', String(error?.message || error));
    }
    updateCheckInProgress = false;
    updateCheckRequested = false;
  });
  autoUpdater.on('update-downloaded', () => {
    updateCheckInProgress = false;
    if (!updateCheckRequested) {
      updateCheckRequested = false;
    }
    const response = dialog.showMessageBoxSync(getDialogParent() || undefined, {
      type: 'question',
      buttons: ['Restart', 'Later'],
      defaultId: 0,
      cancelId: 1,
      message: 'Update ready',
      detail: 'Restart the app to apply the update.'
    });
    if (response === 0) {
      autoUpdater.quitAndInstall();
    }
    updateCheckRequested = false;
  });
}

function getDisplayById(displayId) {
  const displays = screen.getAllDisplays();
  return displays.find((d) => d.id === displayId) || screen.getPrimaryDisplay();
}

function positionProgramWindow(displayId) {
  if (!programWindow) {
    return;
  }
  const display = getDisplayById(displayId);
  programWindow.setBounds(display.bounds);
  programWindow.setFullScreen(true);
  programWindow.setAlwaysOnTop(true, 'screen-saver');
}

function createProgramWindow(displayId) {
  if (programWindow) {
    positionProgramWindow(displayId);
    programWindow.showInactive();
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.focus();
    }
    return;
  }

  programWindow = new BrowserWindow({
    show: false,
    frame: false,
    backgroundColor: '#000000',
    webPreferences: {
      autoplayPolicy: 'no-user-gesture-required',
      backgroundThrottling: false,
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  programWindow.loadFile(path.join(__dirname, 'src', 'program.html'));
  programWindow.once('ready-to-show', () => {
    positionProgramWindow(displayId);
    programWindow.showInactive();
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.focus();
    }
  });
  programWindow.on('closed', () => {
    programWindow = null;
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('program:event', { type: 'program-hidden' });
    }
  });
}

function getLibraryFolder() {
  return path.join(app.getAppPath(), 'library');
}

async function ensureLibraryFolder() {
  const folder = getLibraryFolder();
  await fs.mkdir(folder, { recursive: true });
  return folder;
}

async function walkLibraryFolder(root, dir, items, options = {}) {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (options.excludeDirs && options.excludeDirs.has(entry.name.toLowerCase())) {
        continue;
      }
      await walkLibraryFolder(root, fullPath, items, options);
      continue;
    }
    const ext = path.extname(entry.name).toLowerCase();
    if (!IMAGE_EXTS.has(ext) && !VIDEO_EXTS.has(ext)) {
      continue;
    }
    const type = IMAGE_EXTS.has(ext) ? 'image' : 'video';
    if (options.typeFilter && !options.typeFilter.has(type)) {
      continue;
    }
    let size = null;
    try {
      const stat = await fs.stat(fullPath);
      size = stat.size;
    } catch (error) {
      size = null;
    }
    items.push({
      name: entry.name,
      relativePath: path.relative(root, fullPath),
      absolutePath: fullPath,
      type,
      fileUrl: pathToFileURL(fullPath).toString(),
      size
    });
  }
}

async function listLibraryItems(options = {}) {
  const scope = options.scope || 'background';
  const baseFolder = await ensureLibraryFolder();
  let root = baseFolder;
  const walkOptions = {};
  if (scope === 'announcements') {
    root = path.join(baseFolder, 'Announcements');
    await fs.mkdir(root, { recursive: true });
    walkOptions.typeFilter = new Set(['image']);
  } else if (scope === 'background') {
    walkOptions.excludeDirs = new Set(['announcements']);
  }
  const items = [];
  await walkLibraryFolder(root, root, items, walkOptions);
  items.sort((a, b) => a.name.localeCompare(b.name));
  return { folder: root, items };
}

function createLibraryWindow(scope = 'background') {
  libraryWindowScope = scope || 'background';
  if (libraryWindow && !libraryWindow.isDestroyed()) {
    libraryWindow.webContents.send('library:scope', libraryWindowScope);
    libraryWindow.show();
    libraryWindow.focus();
    return;
  }
  libraryWindow = new BrowserWindow({
    width: 1120,
    height: 700,
    minWidth: 1024,
    minHeight: 480,
    parent: mainWindow || undefined,
    modal: false,
    backgroundColor: '#101418',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });
  libraryWindow.loadFile(path.join(__dirname, 'src', 'library.html'), {
    query: { scope: libraryWindowScope }
  });
  libraryWindow.on('closed', () => {
    libraryWindow = null;
  });
}

async function ensureDir(folderPath) {
  await fs.mkdir(folderPath, { recursive: true });
}

async function saveProject(folderPath, project) {
  await ensureDir(folderPath);
  await ensureDir(path.join(folderPath, 'media'));
  await ensureDir(path.join(folderPath, 'thumbnails'));
  const projectPath = path.join(folderPath, 'project.json');
  await fs.writeFile(projectPath, JSON.stringify(project, null, 2), 'utf8');
}

function uniqueFileName(targetDir, baseName, ext) {
  let candidate = `${baseName}${ext}`;
  let counter = 1;
  while (fssync.existsSync(path.join(targetDir, candidate))) {
    candidate = `${baseName}_${counter}${ext}`;
    counter += 1;
  }
  return candidate;
}

function mediaTypeFromExt(ext) {
  const lower = ext.toLowerCase();
  if (VIDEO_EXTS.has(lower)) {
    return 'video';
  }
  if (IMAGE_EXTS.has(lower)) {
    return 'image';
  }
  return 'image';
}

function libraryFolderForType(type) {
  if (type === 'video') {
    return 'Videos';
  }
  if (type === 'image') {
    return 'Images';
  }
  return 'Images';
}

function buildLibraryRelativePath(baseName, ext, type) {
  const folder = libraryFolderForType(type);
  return `library/${folder}/${baseName}${ext}`;
}

function buildAnnouncementRelativePath(baseName, ext) {
  return `library/Announcements/${baseName}${ext}`;
}

function normalizeLibraryRelativePath(relativePath) {
  if (!relativePath) {
    return '';
  }
  const cleaned = relativePath.replace(/^library[\\/]/, '');
  return `library/${cleaned.replace(/\\/g, '/')}`;
}

function ensureUniqueRelativePath(relativePath, used) {
  if (!relativePath) {
    return '';
  }
  let candidate = relativePath.replace(/\\/g, '/');
  if (!used.has(candidate)) {
    used.add(candidate);
    return candidate;
  }
  const dir = path.posix.dirname(candidate);
  const ext = path.posix.extname(candidate);
  const base = path.posix.basename(candidate, ext);
  let counter = 1;
  while (true) {
    const name = `${base}_${counter}${ext}`;
    const next = dir === '.' ? name : path.posix.join(dir, name);
    if (!used.has(next)) {
      used.add(next);
      return next;
    }
    counter += 1;
  }
}

function resolveAssetPath(assetPath, options = {}) {
  if (!assetPath || typeof assetPath !== 'string') {
    return null;
  }
  if (assetPath.startsWith('library/') || assetPath.startsWith('library\\')) {
    const cleaned = assetPath.replace(/^library[\\/]/, '');
    return path.join(getLibraryFolder(), cleaned);
  }
  if (assetPath.startsWith('media/') || assetPath.startsWith('media\\')) {
    if (!options.projectFolder) {
      return null;
    }
    return path.join(options.projectFolder, assetPath);
  }
  if (assetPath.startsWith('file://')) {
    try {
      return fileURLToPath(assetPath);
    } catch (error) {
      return null;
    }
  }
  if (/^[a-zA-Z]:[\\/]/.test(assetPath)) {
    return assetPath;
  }
  return null;
}

async function statIfExists(filePath) {
  try {
    return await fs.stat(filePath);
  } catch (error) {
    return null;
  }
}

async function collectSessionAssets(project, options = {}) {
  const projectCopy = project ? JSON.parse(JSON.stringify(project)) : {};
  const assets = [];
  const usedRelativePaths = new Set();
  const pathMap = new Map();
  const addAssetFromPath = async (sourcePath, defaultBuilder, typeOverride) => {
    if (!sourcePath) {
      return null;
    }
    const sourceKey = sourcePath.replace(/\\/g, '/');
    if (pathMap.has(sourceKey)) {
      return pathMap.get(sourceKey);
    }
    const resolvedPath = resolveAssetPath(sourcePath, options);
    if (!resolvedPath) {
      return null;
    }
    let data;
    try {
      data = await fs.readFile(resolvedPath);
    } catch (error) {
      return null;
    }
    const ext = path.extname(resolvedPath).toLowerCase();
    const type = typeOverride || mediaTypeFromExt(ext);
    const base = path.basename(resolvedPath, ext) || 'background';
    let relativePath = /^library[\\/]/.test(sourcePath)
      ? normalizeLibraryRelativePath(sourcePath)
      : defaultBuilder(base, ext, type);
    relativePath = ensureUniqueRelativePath(relativePath, usedRelativePaths);
    assets.push({
      relativePath,
      sourcePath,
      type,
      size: data.length,
      data: data.toString('base64')
    });
    pathMap.set(sourceKey, relativePath);
    return relativePath;
  };

  const songs = projectCopy.songs || {};
  for (const song of Object.values(songs)) {
    if (!song || !song.background || !song.background.path) {
      continue;
    }
    const relativePath = await addAssetFromPath(song.background.path, buildLibraryRelativePath);
    if (relativePath) {
      song.background.path = relativePath;
    }
  }

  const announcements = projectCopy.announcements || {};
  if (Array.isArray(announcements.slides)) {
    for (const slide of announcements.slides) {
      if (!slide || !slide.mediaPath) {
        continue;
      }
      const relativePath = await addAssetFromPath(
        slide.mediaPath,
        (base, ext) => buildAnnouncementRelativePath(base, ext),
        'image'
      );
      if (relativePath) {
        slide.mediaPath = relativePath;
      }
    }
  }

  const timer = projectCopy.timer || {};
  if (Array.isArray(timer.slides)) {
    for (const slide of timer.slides) {
      if (!slide || !slide.mediaPath) {
        continue;
      }
      const relativePath = await addAssetFromPath(slide.mediaPath, buildLibraryRelativePath);
      if (relativePath) {
        slide.mediaPath = relativePath;
      }
    }
  }

  return { project: projectCopy, assets };
}

async function restoreSessionAssets(project, assets) {
  if (!Array.isArray(assets) || assets.length === 0) {
    return { project, restored: 0 };
  }
  const libraryDir = await ensureLibraryFolder();
  const libraryIndex = await listLibraryItems({ scope: 'all' });
  const existingByRelative = new Map();
  const existingByNameSize = new Map();
  const existingByNameSizeAnnouncements = new Map();
  (libraryIndex.items || []).forEach((item) => {
    const relative = `library/${item.relativePath.replace(/\\/g, '/')}`;
    existingByRelative.set(relative, item);
    if (item.size != null) {
      const key = `${item.name.toLowerCase()}|${item.size}`;
      const isAnnouncement = /library\/announcements\//i.test(relative);
      if (isAnnouncement) {
        if (!existingByNameSizeAnnouncements.has(key)) {
          existingByNameSizeAnnouncements.set(key, relative);
        }
      } else if (!existingByNameSize.has(key)) {
        existingByNameSize.set(key, relative);
      }
    }
  });
  const pathMap = new Map();
  let restored = 0;

  for (const asset of assets) {
    if (!asset || !asset.relativePath || !asset.data) {
      continue;
    }
    const normalizedRelative = normalizeLibraryRelativePath(asset.relativePath);
    const basename = path.posix.basename(normalizedRelative).toLowerCase();
    const existingEntry = existingByRelative.get(normalizedRelative);
    if (existingEntry && asset.size != null && existingEntry.size === asset.size) {
      pathMap.set(normalizedRelative, normalizedRelative);
      pathMap.set(normalizedRelative.replace(/\//g, '\\'), normalizedRelative);
      if (asset.sourcePath) {
        pathMap.set(asset.sourcePath, normalizedRelative);
        pathMap.set(asset.sourcePath.replace(/\\/g, '/'), normalizedRelative);
      }
      continue;
    }
    if (asset.size != null) {
      const matchKey = `${basename}|${asset.size}`;
      const isAnnouncement = /library\/announcements\//i.test(normalizedRelative);
      const matchingRelative = isAnnouncement
        ? existingByNameSizeAnnouncements.get(matchKey)
        : existingByNameSize.get(matchKey);
      if (matchingRelative) {
        pathMap.set(normalizedRelative, matchingRelative);
        pathMap.set(normalizedRelative.replace(/\//g, '\\'), matchingRelative);
        if (asset.sourcePath) {
          pathMap.set(asset.sourcePath, matchingRelative);
          pathMap.set(asset.sourcePath.replace(/\\/g, '/'), matchingRelative);
        }
        continue;
      }
    }

    const ext = path.extname(normalizedRelative).toLowerCase();
    const base = path.basename(normalizedRelative, ext) || 'background';
    const type = asset.type || mediaTypeFromExt(ext);
    let finalRelative = normalizedRelative;
    if (!/library\/(Images|Videos)\//i.test(finalRelative)) {
      finalRelative = buildLibraryRelativePath(base, ext, type);
    }
    const cleaned = finalRelative.replace(/^library[\\/]/, '');
    const targetDir = path.join(libraryDir, path.dirname(cleaned));
    await fs.mkdir(targetDir, { recursive: true });
    let targetPath = path.join(libraryDir, cleaned);

    const buffer = Buffer.from(asset.data, 'base64');
    const existingStat = await statIfExists(targetPath);
    if (existingStat && asset.size && existingStat.size !== asset.size) {
      const uniqueName = uniqueFileName(path.dirname(targetPath), base, ext);
      targetPath = path.join(path.dirname(targetPath), uniqueName);
      finalRelative = path.posix.join('library', path.relative(libraryDir, targetPath).replace(/\\/g, '/'));
    }
    if (!existingStat || (asset.size && existingStat.size !== asset.size)) {
      await fs.writeFile(targetPath, buffer);
      restored += 1;
    }
    const normalizedFinal = finalRelative.replace(/\\/g, '/');
    if (asset.size != null) {
      const key = `${path.posix.basename(normalizedFinal).toLowerCase()}|${asset.size}`;
      if (/library\/announcements\//i.test(normalizedFinal)) {
        existingByNameSizeAnnouncements.set(key, normalizedFinal);
      } else {
        existingByNameSize.set(key, normalizedFinal);
      }
    }
    existingByRelative.set(normalizedFinal, { relativePath: normalizedFinal, size: asset.size });
    pathMap.set(normalizedRelative, normalizedFinal);
    pathMap.set(normalizedRelative.replace(/\//g, '\\'), normalizedFinal);
    if (asset.sourcePath) {
      pathMap.set(asset.sourcePath, normalizedFinal);
      pathMap.set(asset.sourcePath.replace(/\\/g, '/'), normalizedFinal);
    }
  }

  if (project && project.songs) {
    Object.values(project.songs).forEach((song) => {
      if (song && song.background && song.background.path) {
        const replacement = pathMap.get(song.background.path);
        if (replacement) {
          song.background.path = replacement;
        }
      }
    });
  }

  if (project && project.announcements && Array.isArray(project.announcements.slides)) {
    project.announcements.slides.forEach((slide) => {
      if (slide && slide.mediaPath) {
        const replacement = pathMap.get(slide.mediaPath);
        if (replacement) {
          slide.mediaPath = replacement;
        }
      }
    });
  }

  if (project && project.timer && Array.isArray(project.timer.slides)) {
    project.timer.slides.forEach((slide) => {
      if (slide && slide.mediaPath) {
        const replacement = pathMap.get(slide.mediaPath);
        if (replacement) {
          slide.mediaPath = replacement;
        }
      }
    });
  }

  return { project, restored };
}

ipcMain.handle('project:new', async (_event, options = {}) => {
  try {
    const mode = options.mode || 'prompt';
    if (mode === 'auto') {
      return await createDefaultProjectFolder();
    }

    const result = await showOpenDialogSafe({
      title: options.title || 'Create or Select Project Folder',
      properties: ['openDirectory', 'createDirectory']
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  } catch (error) {
    logError('Failed to resolve project folder', error);
    dialog.showErrorBox('Project Folder Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('project:open', async () => {
  try {
    const result = await showOpenDialogSafe({
      title: 'Open Project Folder',
      properties: ['openDirectory']
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  } catch (error) {
    logError('Failed to open project folder dialog', error);
    dialog.showErrorBox('Project Folder Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('project:load', async (_event, folderPath) => {
  const projectPath = path.join(folderPath, 'project.json');
  const raw = await fs.readFile(projectPath, 'utf8');
  return JSON.parse(raw);
});

ipcMain.handle('project:save', async (_event, folderPath, project) => {
  await saveProject(folderPath, project);
  return true;
});

ipcMain.handle('project:save-file', async (_event, filePath, project, options = {}) => {
  try {
    let targetPath = filePath;
    if (!targetPath) {
      const result = await dialog.showSaveDialog(getDialogParent() || undefined, {
        title: 'Save Session',
        defaultPath: path.join(app.getPath('documents'), 'WorshipSession.wpjson'),
        filters: [{ name: 'Worship Presenter Session', extensions: ['wpjson'] }]
      });
      if (result.canceled || !result.filePath) {
        return null;
      }
      targetPath = result.filePath;
    }
    const { project: projectCopy, assets } = await collectSessionAssets(project, options);
    const payload = {
      sessionVersion: SESSION_VERSION,
      project: projectCopy,
      assets
    };
    await fs.writeFile(targetPath, JSON.stringify(payload, null, 2), 'utf8');
    return targetPath;
  } catch (error) {
    logError('Failed to save session file', error);
    dialog.showErrorBox('Save Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('project:open-file', async () => {
  try {
    const result = await showOpenDialogSafe({
      title: 'Open Session',
      properties: ['openFile'],
      filters: [{ name: 'Worship Presenter Session', extensions: ['wpjson'] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    const filePath = result.filePaths[0];
    const raw = await fs.readFile(filePath, 'utf8');
    const parsed = JSON.parse(raw);
    const project = parsed && parsed.project ? parsed.project : parsed;
    const assets = parsed && parsed.assets ? parsed.assets : [];
    const restored = await restoreSessionAssets(project, assets);
    return { filePath, project: restored.project || project };
  } catch (error) {
    logError('Failed to open session file', error);
    dialog.showErrorBox('Open Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('assets:check', async (_event, payload = {}) => {
  const paths = Array.isArray(payload.paths) ? payload.paths : [];
  const projectFolder = payload.projectFolder || null;
  const missing = [];
  for (const assetPath of paths) {
    const resolved = resolveAssetPath(assetPath, { projectFolder });
    if (!resolved) {
      missing.push(assetPath);
      continue;
    }
    const exists = await statIfExists(resolved);
    if (!exists) {
      missing.push(assetPath);
    }
  }
  return { missing };
});

ipcMain.handle('lyrics:pick', async () => {
  try {
    const result = await showOpenDialogSafe({
      title: 'Import Lyrics',
      properties: ['openFile'],
      filters: [{ name: 'Text Files', extensions: ['txt'] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  } catch (error) {
    logError('Failed to open lyrics picker', error);
    dialog.showErrorBox('Lyrics Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('file:read-text', async (_event, filePath) => {
  try {
    return await fs.readFile(filePath, 'utf8');
  } catch (error) {
    logError('Failed to read text file', error);
    dialog.showErrorBox('Read Error', String(error?.message || error));
    return '';
  }
});

ipcMain.handle('media:pick', async () => {
  try {
    const result = await showOpenDialogSafe({
      title: 'Select Background Media',
      properties: ['openFile'],
      filters: [
        { name: 'Media', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'mp4', 'mov', 'mkv', 'avi', 'webm'] }
      ]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  } catch (error) {
    logError('Failed to open media picker', error);
    dialog.showErrorBox('Media Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('announcement:pick', async () => {
  try {
    const result = await showOpenDialogSafe({
      title: 'Select Announcement Image',
      properties: ['openFile'],
      filters: [
        { name: 'Images', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'] }
      ]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  } catch (error) {
    logError('Failed to open announcement picker', error);
    dialog.showErrorBox('Announcement Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('media:import', async (_event, projectFolder, sourcePath) => {
  const mediaDir = path.join(projectFolder, 'media');
  await ensureDir(mediaDir);
  const ext = path.extname(sourcePath);
  const base = path.basename(sourcePath, ext);
  const fileName = uniqueFileName(mediaDir, base, ext);
  const target = path.join(mediaDir, fileName);
  await fs.copyFile(sourcePath, target);
  return {
    relativePath: path.posix.join('media', fileName),
    absolutePath: target
  };
});

ipcMain.handle('library:import', async (_event, sourcePath) => {
  const libraryDir = await ensureLibraryFolder();
  const ext = path.extname(sourcePath);
  const base = path.basename(sourcePath, ext);
  const type = mediaTypeFromExt(ext);
  const subfolder = libraryFolderForType(type);
  const targetDir = path.join(libraryDir, subfolder);
  await ensureDir(targetDir);
  const fileName = uniqueFileName(targetDir, base, ext);
  const target = path.join(targetDir, fileName);
  await fs.copyFile(sourcePath, target);
  const relative = path.relative(libraryDir, target).replace(/\\/g, '/');
  return {
    relativePath: path.posix.join('library', relative),
    absolutePath: target
  };
});

ipcMain.handle('announcement:import', async (_event, sourcePath) => {
  const libraryDir = await ensureLibraryFolder();
  const ext = path.extname(sourcePath);
  const base = path.basename(sourcePath, ext);
  const targetDir = path.join(libraryDir, 'Announcements');
  await ensureDir(targetDir);
  const fileName = uniqueFileName(targetDir, base, ext);
  const target = path.join(targetDir, fileName);
  await fs.copyFile(sourcePath, target);
  const relative = path.relative(libraryDir, target).replace(/\\/g, '/');
  return {
    relativePath: path.posix.join('library', relative),
    absolutePath: target
  };
});

ipcMain.handle('display:list', async () => {
  return screen.getAllDisplays().map((display) => ({
    id: display.id,
    label: `${display.id} - ${display.bounds.width}x${display.bounds.height}`,
    bounds: display.bounds,
    scaleFactor: display.scaleFactor
  }));
});

ipcMain.on('program:show', (_event, displayId) => {
  createProgramWindow(displayId);
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.focus();
  }
});

ipcMain.on('program:hide', () => {
  if (programWindow) {
    programWindow.hide();
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('program:event', { type: 'program-hidden' });
  }
});

ipcMain.on('program:set-display', (_event, displayId) => {
  if (programWindow) {
    positionProgramWindow(displayId);
  }
});

ipcMain.on('program:toggle-fullscreen', () => {
  if (programWindow) {
    programWindow.setFullScreen(!programWindow.isFullScreen());
  }
});

ipcMain.on('program:state', (_event, state) => {
  if (programWindow && !programWindow.isDestroyed()) {
    programWindow.webContents.send('program:state', state);
  }
});

ipcMain.on('program:event', (_event, payload) => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('program:event', payload);
  }
});

ipcMain.handle('library:list', async (_event, options = {}) => {
  return listLibraryItems(options);
});

ipcMain.on('library:open', (_event, options = {}) => {
  const scope = options.scope || 'background';
  createLibraryWindow(scope);
});

ipcMain.on('library:close', () => {
  if (libraryWindow && !libraryWindow.isDestroyed()) {
    libraryWindow.close();
  }
});

ipcMain.on('library:select', (_event, payload) => {
  const sourcePath = payload && payload.sourcePath ? payload.sourcePath : payload;
  const scope = payload && payload.scope ? payload.scope : libraryWindowScope || 'background';
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('library:selected', { scope, sourcePath });
  }
  if (libraryWindow && !libraryWindow.isDestroyed()) {
    libraryWindow.close();
  }
});

app.whenReady().then(() => {
  createMainWindow();
  createAppMenu();
  initAutoUpdater();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
