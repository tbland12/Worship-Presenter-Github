const { app, BrowserWindow, dialog, ipcMain, screen, Menu, autoUpdater } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const fssync = require('fs');
const crypto = require('crypto');
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
const programWindows = new Map();
let libraryWindow;
let libraryWindowScope = 'background';
let updateCheckInProgress = false;
let updateCheckRequested = false;
let libraryInitPromise = null;
let programVisible = false;
let programDisplayTarget = 'all-external';
let programDisplayPreference = null;
let displayChangeTimer = null;

const UPDATE_REPOSITORY = 'tbland12/Worship-Presenter-Github';

const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']);
const VIDEO_EXTS = new Set(['.mp4', '.mov', '.mkv', '.avi', '.webm']);
const SESSION_VERSION = 2;
const LIBRARY_SEED_FILE = '.seeded';
const DISPLAY_TARGET_ALL_EXTERNAL = 'all-external';
const TIMER_LIBRARY_FOLDER = 'Timers';

programDisplayPreference = { mode: DISPLAY_TARGET_ALL_EXTERNAL };

function normalizeDisplayTarget(target) {
  if (target == null) {
    return DISPLAY_TARGET_ALL_EXTERNAL;
  }
  if (target === DISPLAY_TARGET_ALL_EXTERNAL) {
    return DISPLAY_TARGET_ALL_EXTERNAL;
  }
  return String(target);
}

function nearlyEqual(first, second, epsilon = 0.01) {
  if (first == null || second == null) {
    return false;
  }
  return Math.abs(first - second) <= epsilon;
}

function buildDisplayPreference(display) {
  if (!display) {
    return { mode: DISPLAY_TARGET_ALL_EXTERNAL };
  }
  return {
    mode: 'display',
    id: String(display.id),
    bounds: {
      x: display.bounds.x,
      y: display.bounds.y,
      width: display.bounds.width,
      height: display.bounds.height
    },
    scaleFactor: display.scaleFactor,
    rotation: display.rotation
  };
}

function setProgramDisplayPreference(target, displays = []) {
  const normalizedTarget = normalizeDisplayTarget(target);
  programDisplayTarget = normalizedTarget;
  if (normalizedTarget === DISPLAY_TARGET_ALL_EXTERNAL) {
    programDisplayPreference = { mode: DISPLAY_TARGET_ALL_EXTERNAL };
    return;
  }
  const match = displays.find((display) => String(display.id) === normalizedTarget);
  if (match) {
    programDisplayPreference = buildDisplayPreference(match);
  } else {
    programDisplayPreference = { mode: 'display', id: normalizedTarget };
  }
}

function centerFromBounds(bounds) {
  return {
    x: bounds.x + bounds.width / 2,
    y: bounds.y + bounds.height / 2
  };
}

function distanceBetween(first, second) {
  const dx = first.x - second.x;
  const dy = first.y - second.y;
  return Math.hypot(dx, dy);
}

function matchDisplayPreference(preference, displays) {
  if (!preference || preference.mode !== 'display') {
    return null;
  }
  const byId = displays.find((display) => String(display.id) === preference.id);
  if (byId) {
    return byId;
  }
  if (!preference.bounds) {
    return null;
  }
  const exactBounds = displays.find((display) => {
    if (!boundsEqual(display.bounds, preference.bounds)) {
      return false;
    }
    if (preference.scaleFactor != null && !nearlyEqual(display.scaleFactor, preference.scaleFactor)) {
      return false;
    }
    if (preference.rotation != null && display.rotation !== preference.rotation) {
      return false;
    }
    return true;
  });
  if (exactBounds) {
    return exactBounds;
  }
  let candidates = displays.filter((display) => {
    return display.bounds.width === preference.bounds.width
      && display.bounds.height === preference.bounds.height;
  });
  if (preference.scaleFactor != null) {
    const scaleMatches = candidates.filter((display) => nearlyEqual(display.scaleFactor, preference.scaleFactor));
    if (scaleMatches.length > 0) {
      candidates = scaleMatches;
    }
  }
  if (preference.rotation != null) {
    const rotationMatches = candidates.filter((display) => display.rotation === preference.rotation);
    if (rotationMatches.length > 0) {
      candidates = rotationMatches;
    }
  }
  if (candidates.length === 0) {
    return null;
  }
  const preferredCenter = centerFromBounds(preference.bounds);
  let best = candidates[0];
  let bestDistance = distanceBetween(preferredCenter, centerFromBounds(best.bounds));
  for (let i = 1; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    const distance = distanceBetween(preferredCenter, centerFromBounds(candidate.bounds));
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
}

function resolvePreferredTarget(displays) {
  const normalizedTarget = normalizeDisplayTarget(programDisplayTarget);
  if (!programDisplayPreference || programDisplayPreference.mode !== 'display') {
    return normalizedTarget;
  }
  if (normalizedTarget === DISPLAY_TARGET_ALL_EXTERNAL) {
    return normalizedTarget;
  }
  const match = matchDisplayPreference(programDisplayPreference, displays);
  if (!match) {
    return normalizedTarget;
  }
  const matchedId = String(match.id);
  if (normalizedTarget !== matchedId) {
    programDisplayTarget = matchedId;
  }
  programDisplayPreference = buildDisplayPreference(match);
  return matchedId;
}

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
    programWindows.forEach((window) => {
      if (window && !window.isDestroyed()) {
        window.close();
      }
    });
    if (libraryWindow && !libraryWindow.isDestroyed()) {
      libraryWindow.close();
    }
    mainWindow = null;
  });
  mainWindow.on('move', scheduleDisplayChange);
  mainWindow.on('resize', scheduleDisplayChange);
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
          accelerator: 'CmdOrCtrl+N',
          click: () => sendMenuAction('new-project')
        },
        {
          label: 'Open Project',
          accelerator: 'CmdOrCtrl+O',
          click: () => sendMenuAction('open-project')
        },
        {
          label: 'Save',
          accelerator: 'CmdOrCtrl+S',
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

function getOperatorDisplay() {
  if (mainWindow && !mainWindow.isDestroyed()) {
    const bounds = mainWindow.isMinimized()
      ? mainWindow.getNormalBounds()
      : mainWindow.getBounds();
    return screen.getDisplayMatching(bounds);
  }
  return screen.getPrimaryDisplay();
}

function resolveTargetDisplayIds(target, displays) {
  const displayIds = displays.map((display) => display.id);
  if (displayIds.length === 0) {
    return { target: null, ids: [] };
  }
  const displayIdStrings = displayIds.map((id) => String(id));
  const operatorDisplay = getOperatorDisplay();
  const operatorId = operatorDisplay ? operatorDisplay.id : displayIds[0];
  const operatorIdString = operatorId != null ? String(operatorId) : null;
  const externalIds = operatorId != null
    ? displayIds.filter((id) => id !== operatorId)
    : displayIds;
  let resolvedTarget = target ?? DISPLAY_TARGET_ALL_EXTERNAL;
  const resolvedTargetString = String(resolvedTarget);

  if (resolvedTargetString === DISPLAY_TARGET_ALL_EXTERNAL) {
    if (externalIds.length > 0) {
      return { target: DISPLAY_TARGET_ALL_EXTERNAL, ids: externalIds };
    }
    if (operatorId != null) {
      return { target: DISPLAY_TARGET_ALL_EXTERNAL, ids: [operatorId] };
    }
    return { target: DISPLAY_TARGET_ALL_EXTERNAL, ids: displayIds.slice(0, 1) };
  }

  const targetIndex = displayIdStrings.indexOf(resolvedTargetString);
  if (targetIndex !== -1) {
    return { target: displayIdStrings[targetIndex], ids: [displayIds[targetIndex]] };
  }

  if (externalIds.length > 0) {
    return { target: DISPLAY_TARGET_ALL_EXTERNAL, ids: externalIds };
  }
  if (operatorId != null) {
    return { target: operatorIdString, ids: [operatorId] };
  }
  return { target: displayIdStrings[0] || null, ids: displayIds.slice(0, 1) };
}

function boundsEqual(first, second) {
  return first.x === second.x
    && first.y === second.y
    && first.width === second.width
    && first.height === second.height;
}

function positionProgramWindow(window, displayId) {
  if (!window || window.isDestroyed()) {
    return;
  }
  const display = getDisplayById(displayId);
  const bounds = display.bounds;
  const currentBounds = window.getBounds();
  if (!boundsEqual(currentBounds, bounds)) {
    window.setBounds(bounds);
  }
  if (!window.isFullScreen()) {
    window.setFullScreen(true);
  }
  if (!window.isAlwaysOnTop()) {
    window.setAlwaysOnTop(true, 'screen-saver');
  }
}

function createProgramWindow(displayId) {
  const existing = programWindows.get(displayId);
  if (existing && !existing.isDestroyed()) {
    positionProgramWindow(existing, displayId);
    if (programVisible) {
      existing.showInactive();
    }
    return existing;
  }

  const window = new BrowserWindow({
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

  programWindows.set(displayId, window);
  window.loadFile(path.join(__dirname, 'src', 'program.html'));
  window.once('ready-to-show', () => {
    positionProgramWindow(window, displayId);
    if (programVisible) {
      window.showInactive();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.focus();
      }
    }
  });
  window.on('closed', () => {
    programWindows.delete(displayId);
    if (programVisible && programWindows.size === 0 && mainWindow && !mainWindow.isDestroyed()) {
      programVisible = false;
      mainWindow.webContents.send('program:event', { type: 'program-hidden' });
    }
  });
  return window;
}

function closeMissingProgramWindows(displayIds) {
  const displayIdSet = new Set(displayIds);
  programWindows.forEach((window, displayId) => {
    if (displayIdSet.has(displayId)) {
      return;
    }
    if (window && !window.isDestroyed()) {
      window.close();
    } else {
      programWindows.delete(displayId);
    }
  });
}

function showProgramWindows(target, options = {}) {
  const focusMain = options.focusMain !== false;
  const updatePreference = options.updatePreference !== false;
  const displays = options.displays || screen.getAllDisplays();
  const displayIds = displays.map((display) => display.id);
  programVisible = true;
  closeMissingProgramWindows(displayIds);
  if (updatePreference && target != null) {
    setProgramDisplayPreference(target, displays);
  }
  const preferredTarget = resolvePreferredTarget(displays);
  const effectiveTarget = target ?? preferredTarget;
  const resolved = resolveTargetDisplayIds(effectiveTarget, displays);
  const targetIds = resolved.ids;
  const targetSet = new Set(targetIds);
  targetIds.forEach((displayId) => {
    createProgramWindow(displayId);
  });
  programWindows.forEach((window, displayId) => {
    if (!targetSet.has(displayId)) {
      if (window && !window.isDestroyed()) {
        window.close();
      } else {
        programWindows.delete(displayId);
      }
      return;
    }
    positionProgramWindow(window, displayId);
    if (!window.isVisible() && window.isReadyToShow()) {
      window.showInactive();
    }
  });
  if (focusMain && mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.focus();
  }
}

function hideProgramWindows() {
  programVisible = false;
  programWindows.forEach((window) => {
    if (window && !window.isDestroyed()) {
      window.hide();
    }
  });
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('program:event', { type: 'program-hidden' });
  }
}

function handleDisplayChange() {
  displayChangeTimer = null;
  const displays = screen.getAllDisplays();
  const displayIds = displays.map((display) => display.id);
  const preferredTarget = resolvePreferredTarget(displays);
  const resolved = resolveTargetDisplayIds(preferredTarget, displays);
  if (programVisible) {
    showProgramWindows(preferredTarget, { focusMain: false, displays, updatePreference: false });
  } else {
    closeMissingProgramWindows(displayIds);
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('display:changed', {
      target: programDisplayTarget,
      resolvedTarget: resolved.target,
      resolvedDisplayIds: resolved.ids
    });
  }
}

function scheduleDisplayChange() {
  if (displayChangeTimer) {
    clearTimeout(displayChangeTimer);
  }
  displayChangeTimer = setTimeout(handleDisplayChange, 200);
}

function getBundledLibraryFolder() {
  if (app.isPackaged) {
    const packaged = path.join(process.resourcesPath, 'library');
    if (fssync.existsSync(packaged)) {
      return packaged;
    }
  }
  return path.join(app.getAppPath(), 'library');
}

function getLibraryFolder() {
  return path.join(app.getPath('userData'), 'library');
}

async function copyLibraryDir(sourceDir, targetDir) {
  let entries;
  try {
    entries = await fs.readdir(sourceDir);
  } catch (error) {
    return;
  }
  await fs.mkdir(targetDir, { recursive: true });
  for (const entry of entries) {
    const sourcePath = path.join(sourceDir, entry);
    const targetPath = path.join(targetDir, entry);
    let stat;
    try {
      stat = await fs.stat(sourcePath);
    } catch (error) {
      continue;
    }
    if (stat.isDirectory()) {
      await copyLibraryDir(sourcePath, targetPath);
      continue;
    }
    if (stat.isFile()) {
      let targetStat = null;
      try {
        targetStat = await fs.stat(targetPath);
      } catch (error) {
        targetStat = null;
      }
      if (targetStat && targetStat.isDirectory()) {
        continue;
      }
      try {
        await fs.copyFile(sourcePath, targetPath);
      } catch (error) {
        logError('Failed to copy library asset', error);
      }
    }
  }
}

async function collectLibraryFiles(rootDir, dir, files) {
  let entries;
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch (error) {
    return;
  }
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectLibraryFiles(rootDir, fullPath, files);
      continue;
    }
    if (!entry.isFile()) {
      continue;
    }
    let stat;
    try {
      stat = await fs.stat(fullPath);
    } catch (error) {
      continue;
    }
    files.push({
      relativePath: path.relative(rootDir, fullPath).replace(/\\/g, '/'),
      size: stat.size,
      mtimeMs: stat.mtimeMs
    });
  }
}

async function buildLibrarySignature(rootDir) {
  const files = [];
  await collectLibraryFiles(rootDir, rootDir, files);
  files.sort((first, second) => first.relativePath.localeCompare(second.relativePath));
  const hash = crypto.createHash('sha256');
  files.forEach((file) => {
    hash.update(`${file.relativePath}|${file.size}|${file.mtimeMs}\n`);
  });
  return hash.digest('hex');
}

async function ensureLibraryFolder() {
  if (libraryInitPromise) {
    return libraryInitPromise;
  }
  libraryInitPromise = (async () => {
    const folder = getLibraryFolder();
    await fs.mkdir(folder, { recursive: true });
    try {
      const seedPath = path.join(folder, LIBRARY_SEED_FILE);
      let seedSignature = null;
      let seedVersion = null;
      try {
        const raw = await fs.readFile(seedPath, 'utf8');
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
          seedSignature = parsed.bundleSignature || null;
          seedVersion = parsed.version || null;
        }
      } catch (error) {
        seedSignature = null;
        seedVersion = null;
      }
      const bundled = getBundledLibraryFolder();
      if (!fssync.existsSync(bundled)) {
        logError('Bundled library folder is missing', bundled);
        return folder;
      }
      let signature = null;
      try {
        signature = await buildLibrarySignature(bundled);
      } catch (error) {
        logError('Failed to build bundled library manifest', error);
      }
      const currentVersion = app.getVersion();
      if (!signature) {
        return folder;
      }
      if (seedSignature !== signature) {
        await copyLibraryDir(bundled, folder);
        const payload = {
          bundleSignature: signature,
          version: currentVersion || seedVersion || null,
          updatedAt: new Date().toISOString()
        };
        await fs.writeFile(seedPath, JSON.stringify(payload, null, 2), 'utf8');
      }
    } catch (error) {
      logError('Failed to initialize library folder', error);
    }
    return folder;
  })();
  return libraryInitPromise;
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
  } else if (scope === 'timer') {
    root = path.join(baseFolder, TIMER_LIBRARY_FOLDER);
    await fs.mkdir(root, { recursive: true });
  } else if (scope === 'background') {
    walkOptions.excludeDirs = new Set(['announcements', TIMER_LIBRARY_FOLDER.toLowerCase()]);
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

function libraryFolderForScope(scope, type) {
  if (scope === 'timer') {
    return TIMER_LIBRARY_FOLDER;
  }
  return libraryFolderForType(type);
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

function sanitizeRelativePath(value) {
  if (!value || typeof value !== 'string') {
    return '';
  }
  const cleaned = value.replace(/\\/g, '/').trim().replace(/^\/+/, '');
  if (!cleaned) {
    return '';
  }
  const normalized = path.posix.normalize(cleaned);
  if (!normalized || normalized === '.' || normalized === '..') {
    return '';
  }
  if (normalized.startsWith('..') || path.posix.isAbsolute(normalized) || /^[a-zA-Z]:/.test(normalized)) {
    return '';
  }
  return normalized.replace(/^(\.\/)+/, '');
}

function sanitizePackRoot(root, fallback) {
  const cleanedRoot = sanitizeRelativePath(root || '');
  if (cleanedRoot) {
    return cleanedRoot;
  }
  const cleanedFallback = sanitizeRelativePath(fallback || '');
  if (cleanedFallback) {
    return cleanedFallback;
  }
  return 'Content Pack';
}

function normalizePackManifest(manifest) {
  if (!manifest || typeof manifest !== 'object') {
    return null;
  }
  const name = typeof manifest.name === 'string' ? manifest.name.trim() : '';
  const root = typeof manifest.root === 'string' ? manifest.root.trim() : '';
  const files = Array.isArray(manifest.files)
    ? manifest.files.filter((entry) => typeof entry === 'string' && entry.trim()).map((entry) => entry.trim())
    : [];
  return { name, root, files };
}

async function readPackManifestFromFile(filePath) {
  if (!filePath) {
    return null;
  }
  try {
    const raw = await fs.readFile(filePath, 'utf8');
    return normalizePackManifest(JSON.parse(raw));
  } catch (error) {
    return null;
  }
}

async function collectPackEntriesFromFolder(rootDir) {
  const entries = [];
  const walk = async (dir) => {
    let items;
    try {
      items = await fs.readdir(dir, { withFileTypes: true });
    } catch (error) {
      return;
    }
    for (const item of items) {
      const fullPath = path.join(dir, item.name);
      if (item.isDirectory()) {
        await walk(fullPath);
        continue;
      }
      if (!item.isFile()) {
        continue;
      }
      if (item.name.toLowerCase() === 'manifest.json') {
        continue;
      }
      const ext = path.extname(item.name).toLowerCase();
      if (!IMAGE_EXTS.has(ext) && !VIDEO_EXTS.has(ext)) {
        continue;
      }
      const relativePath = path.relative(rootDir, fullPath).replace(/\\/g, '/');
      const stat = await statIfExists(fullPath);
      entries.push({
        relativePath,
        absolutePath: fullPath,
        size: stat ? stat.size : null
      });
    }
  };
  await walk(rootDir);
  return entries;
}

async function importContentPackFile(libraryDir, baseRoot, entry) {
  const outcome = { imported: 0, skipped: 0, failed: 0, renamed: 0 };
  const entryRelative = sanitizeRelativePath(entry.relativePath);
  if (!entryRelative) {
    outcome.failed += 1;
    return outcome;
  }
  const ext = path.posix.extname(entryRelative).toLowerCase();
  if (!IMAGE_EXTS.has(ext) && !VIDEO_EXTS.has(ext)) {
    outcome.skipped += 1;
    return outcome;
  }
  let targetRelative = baseRoot ? path.posix.join(baseRoot, entryRelative) : entryRelative;
  targetRelative = sanitizeRelativePath(targetRelative);
  if (!targetRelative) {
    outcome.failed += 1;
    return outcome;
  }
  let targetPath = path.join(libraryDir, targetRelative);
  const relativeCheck = path.relative(libraryDir, targetPath);
  if (relativeCheck.startsWith('..') || path.isAbsolute(relativeCheck)) {
    outcome.failed += 1;
    return outcome;
  }
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  let dataSize = entry.size;
  let data = null;
  if (entry.data) {
    data = entry.data;
    if (dataSize == null) {
      dataSize = data.length;
    }
  }
  const existing = await statIfExists(targetPath);
  if (existing && dataSize != null && existing.size === dataSize) {
    outcome.skipped += 1;
    return outcome;
  }
  if (existing && dataSize != null && existing.size !== dataSize) {
    const base = path.basename(targetPath, path.extname(targetPath));
    const fileName = uniqueFileName(path.dirname(targetPath), base, ext);
    targetPath = path.join(path.dirname(targetPath), fileName);
    outcome.renamed += 1;
  }
  try {
    if (data) {
      await fs.writeFile(targetPath, data);
    } else if (entry.absolutePath) {
      await fs.copyFile(entry.absolutePath, targetPath);
    } else {
      outcome.failed += 1;
      return outcome;
    }
  } catch (error) {
    logError('Failed to import content pack file', error);
    outcome.failed += 1;
    return outcome;
  }
  outcome.imported += 1;
  return outcome;
}

async function importContentPackFromFolder(folderPath) {
  const manifestPath = path.join(folderPath, 'manifest.json');
  const manifest = await readPackManifestFromFile(manifestPath);
  const folderName = path.basename(folderPath);
  const manifestRoot = manifest ? sanitizeRelativePath(manifest.root || '') : '';
  const baseRoot = sanitizePackRoot(manifestRoot || (manifest && manifest.name) || folderName, folderName);
  const sourceRoot = manifestRoot ? path.join(folderPath, manifestRoot) : folderPath;
  let entries = [];
  if (manifest && Array.isArray(manifest.files) && manifest.files.length > 0) {
    const sanitizedFiles = manifest.files.map((file) => sanitizeRelativePath(file)).filter(Boolean);
    entries = sanitizedFiles.map((file) => ({
      relativePath: file,
      absolutePath: path.join(sourceRoot, file),
      size: null
    }));
  } else {
    entries = await collectPackEntriesFromFolder(sourceRoot);
  }
  if (entries.length === 0) {
    return { error: 'No supported media files found in the selected folder.' };
  }
  const libraryDir = await ensureLibraryFolder();
  const summary = { imported: 0, skipped: 0, failed: 0, renamed: 0, total: entries.length };
  for (const entry of entries) {
    if (entry.absolutePath) {
      const stat = await statIfExists(entry.absolutePath);
      entry.size = stat ? stat.size : null;
    }
    const outcome = await importContentPackFile(libraryDir, baseRoot, entry);
    summary.imported += outcome.imported;
    summary.skipped += outcome.skipped;
    summary.failed += outcome.failed;
    summary.renamed += outcome.renamed;
  }
  return {
    ...summary,
    packName: manifest && manifest.name ? manifest.name : folderName
  };
}

async function importContentPack(sourcePath) {
  if (!sourcePath) {
    return { error: 'No content pack selected.' };
  }
  const stat = await statIfExists(sourcePath);
  if (!stat) {
    return { error: 'Content pack not found.' };
  }
  if (stat.isDirectory()) {
    return importContentPackFromFolder(sourcePath);
  }
  if (stat.isFile() && path.extname(sourcePath).toLowerCase() === '.zip') {
    return { error: 'Zip content packs are not supported. Please unzip and import the folder.' };
  }
  return { error: 'Select a folder content pack to import.' };
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
    const libraryPath = path.join(getLibraryFolder(), cleaned);
    if (fssync.existsSync(libraryPath)) {
      return libraryPath;
    }
    return path.join(getBundledLibraryFolder(), cleaned);
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
    if (!/library\/(Images|Videos|Timers)\//i.test(finalRelative)) {
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

ipcMain.handle('library:roots', async () => {
  await ensureLibraryFolder();
  return {
    libraryRoot: getLibraryFolder(),
    bundledLibraryRoot: getBundledLibraryFolder()
  };
});
ipcMain.handle('app:get-version', () => app.getVersion());

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

ipcMain.handle('library:import', async (_event, payload) => {
  const sourcePath = typeof payload === 'string' ? payload : payload && payload.sourcePath;
  const scope = payload && typeof payload === 'object' ? payload.scope : null;
  if (!sourcePath) {
    return null;
  }
  const libraryDir = await ensureLibraryFolder();
  const ext = path.extname(sourcePath);
  const base = path.basename(sourcePath, ext);
  const type = mediaTypeFromExt(ext);
  const subfolder = libraryFolderForScope(scope, type);
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

ipcMain.handle('content-pack:import', async (_event, options = {}) => {
  try {
    let sourcePath = options && options.sourcePath ? options.sourcePath : null;
    if (!sourcePath) {
      const result = await showOpenDialogSafe({
        title: 'Import Content Pack (Folder)',
        properties: ['openDirectory']
      });
      if (result.canceled || result.filePaths.length === 0) {
        return { canceled: true };
      }
      sourcePath = result.filePaths[0];
    }
    return await importContentPack(sourcePath);
  } catch (error) {
    logError('Failed to import content pack', error);
    return { error: 'Failed to import content pack.' };
  }
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
  showProgramWindows(displayId);
});

ipcMain.on('program:hide', () => {
  hideProgramWindows();
});

ipcMain.on('program:set-display', (_event, displayId) => {
  const displays = screen.getAllDisplays();
  setProgramDisplayPreference(displayId, displays);
  if (programVisible) {
    showProgramWindows(programDisplayTarget, { displays, updatePreference: false });
  }
});

ipcMain.on('program:toggle-fullscreen', () => {
  programWindows.forEach((window) => {
    if (window && !window.isDestroyed()) {
      window.setFullScreen(!window.isFullScreen());
    }
  });
});

ipcMain.on('program:state', (_event, state) => {
  programWindows.forEach((window) => {
    if (window && !window.isDestroyed()) {
      window.webContents.send('program:state', state);
    }
  });
});

ipcMain.on('program:event', (_event, payload) => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('program:event', payload);
  }
});

ipcMain.on('edit:command', (_event, command) => {
  const allowed = new Set(['undo', 'redo', 'cut', 'copy', 'paste', 'selectAll']);
  if (!allowed.has(command)) {
    return;
  }
  const focused = BrowserWindow.getFocusedWindow();
  if (!focused || focused.isDestroyed()) {
    return;
  }
  const contents = focused.webContents;
  const handler = contents && typeof contents[command] === 'function' ? contents[command] : null;
  if (handler) {
    handler.call(contents);
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
  screen.on('display-added', scheduleDisplayChange);
  screen.on('display-removed', scheduleDisplayChange);
  screen.on('display-metrics-changed', scheduleDisplayChange);

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
