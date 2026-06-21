const { app, BrowserWindow, dialog, ipcMain, screen, Menu, autoUpdater, session, protocol, net } = require('electron');
if (require('electron-squirrel-startup')) {
  app.quit();
}
const path = require('path');
const fs = require('fs/promises');
const fssync = require('fs');
const crypto = require('crypto');
const { pathToFileURL } = require('url');
const JSZip = require('jszip');
const { describeAsset, readSession, writeSession } = require('./src/main/session-v3');
const { createPersistenceStores } = require('./src/main/persistence-stores');
const {
  buildHistoryEntry,
  publicHistoryEntries,
  removeHistoryEntry,
  upsertHistory
} = require('./src/main/project-history');
const { migrateProject } = require('./src/main/project-v2');
const { createSongLibrary, SongLibraryConflictError } = require('./src/main/song-library');

protocol.registerSchemesAsPrivileged([
  { scheme: 'worship-media', privileges: { secure: true, standard: true, stream: true, supportFetchAPI: true } }
]);

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
let stageWindow;
let libraryWindow;
let libraryWindowScope = 'background';
let updateCheckInProgress = false;
let updateCheckRequested = false;
let libraryInitPromise = null;
let programVisible = false;
let stageVisible = false;
let stageDisplayTarget = null;
let latestStageState = null;
let programDisplayTarget = 'all-external';
let programDisplayPreference = null;
let displayChangeTimer = null;
let activeSessionPath = null;
let sessionAssetPaths = new Map();
let projectDirty = false;
let allowMainWindowClose = false;
let closePromptOpen = false;
let sessionWriteQueue = Promise.resolve();
let recoveryWriteQueue = Promise.resolve();
let historyWriteQueue = Promise.resolve();
let songLibraryWriteQueue = Promise.resolve();
let preferenceWriteQueue = Promise.resolve();
let _persistenceStores = null;
let _songLibrary = null;
let recentSessionEntries = [];
const fileTokens = new Map();
const FILE_TOKEN_TTL_MS = 5 * 60 * 1000;

function issueFileToken(event, filePath) {
  const extension = path.extname(filePath).toLowerCase();
  if (!IMPORT_EXTENSIONS.has(extension)) throw new Error('Unsupported file type.');
  const token = crypto.randomUUID();
  fileTokens.set(token, { filePath, senderId: event.sender.id, expiresAt: Date.now() + FILE_TOKEN_TTL_MS });
  return { token, name: path.basename(filePath), extension };
}

function assertOperatorSender(event) {
  if (!mainWindow || mainWindow.isDestroyed() || event.sender.id !== mainWindow.webContents.id) {
    throw new Error('IPC request is not authorized for this window.');
  }
}

function assertLibrarySender(event) {
  if (!libraryWindow || libraryWindow.isDestroyed() || event.sender.id !== libraryWindow.webContents.id) {
    throw new Error('IPC request is not authorized for this window.');
  }
}

function assertProgramSender(event) {
  const authorized = [...programWindows.values()].some((window) => {
    return window && !window.isDestroyed() && event.sender.id === window.webContents.id;
  });
  if (!authorized) throw new Error('IPC request is not authorized for this window.');
}

function assertStageSender(event) {
  if (!stageWindow || stageWindow.isDestroyed() || event.sender.id !== stageWindow.webContents.id) {
    throw new Error('IPC request is not authorized for this window.');
  }
}

function assertOperatorOrProgramSender(event) {
  if (mainWindow && !mainWindow.isDestroyed() && event.sender.id === mainWindow.webContents.id) return;
  assertProgramSender(event);
}

function assertOperatorOrStageSender(event) {
  if (mainWindow && !mainWindow.isDestroyed() && event.sender.id === mainWindow.webContents.id) return;
  assertStageSender(event);
}

function consumeFileToken(event, token, allowedExtensions) {
  const record = fileTokens.get(token);
  fileTokens.delete(token);
  if (!record || record.senderId !== event.sender.id || record.expiresAt < Date.now()) {
    throw new Error('File selection expired or is invalid.');
  }
  const extension = path.extname(record.filePath).toLowerCase();
  if (allowedExtensions && !allowedExtensions.has(extension)) throw new Error('Unsupported file type.');
  return record.filePath;
}

const UPDATE_REPOSITORY = 'tbland12/Worship-Presenter-Github';

const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']);
const VIDEO_EXTS = new Set(['.mp4', '.mov', '.mkv', '.avi', '.webm']);
const IMPORT_EXTENSIONS = new Set([...IMAGE_EXTS, ...VIDEO_EXTS, '.txt', '.pptx']);
const LIBRARY_SEED_FILE = '.seeded';
const DISPLAY_TARGET_ALL_EXTERNAL = 'all-external';
const TIMER_LIBRARY_FOLDER = 'Timers';
const MAX_PPTX_BYTES = 100 * 1024 * 1024;
const MAX_PPTX_SLIDES = 512;
const MAX_PPTX_XML_BYTES = 256 * 1024 * 1024;

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

function hardenWindow(window) {
  if (!window || window.isDestroyed()) {
    return;
  }
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => {
    event.preventDefault();
  });
  window.webContents.on('will-attach-webview', (event) => {
    event.preventDefault();
  });
}

function createMainWindow() {
  allowMainWindowClose = false;
  closePromptOpen = false;
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1200,
    minHeight: 720,
    backgroundColor: '#101418',
    webPreferences: {
      autoplayPolicy: 'no-user-gesture-required',
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  hardenWindow(mainWindow);
  mainWindow.loadFile(path.join(__dirname, 'src', 'operator.html'));
  mainWindow.on('close', async (event) => {
    if (allowMainWindowClose || !projectDirty) return;
    if (closePromptOpen) {
      event.preventDefault();
      return;
    }
    event.preventDefault();
    closePromptOpen = true;
    let awaitingSave = false;
    try {
      const result = await dialog.showMessageBox(mainWindow, {
        type: 'warning',
        title: 'Unsaved Changes',
        message: 'Save changes before closing Worship Presenter?',
        detail: 'Unsaved changes will be lost if you discard them.',
        buttons: ['Save', 'Discard', 'Cancel'],
        defaultId: 0,
        cancelId: 2,
        noLink: true
      });
      if (result.response === 0) {
        awaitingSave = true;
        mainWindow.webContents.send('project:save-before-close');
      } else if (result.response === 1) {
        await clearRecoverySnapshot().catch((error) => logError('Failed to clear discarded recovery snapshot', error));
        allowMainWindowClose = true;
        mainWindow.close();
      }
    } finally {
      if (!awaitingSave) closePromptOpen = false;
    }
  });
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

function recentProjectMenuItems() {
  if (recentSessionEntries.length === 0) {
    return [{ label: 'No Recent Projects', enabled: false }];
  }
  return recentSessionEntries.slice(0, 10).map((entry) => ({
    label: entry.title,
    click: () => sendMenuAction(`open-recent:${entry.id}`)
  }));
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
          label: 'Open Recent',
          submenu: recentProjectMenuItems()
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
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload-program.js')
    }
  });

  programWindows.set(displayId, window);
  hardenWindow(window);
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

function resolveStageDisplayId(target, displays = screen.getAllDisplays()) {
  if (displays.length === 0) return null;
  const requested = target == null ? null : String(target);
  const exact = requested && displays.find((display) => String(display.id) === requested);
  if (exact) return exact.id;
  const operatorDisplay = getOperatorDisplay();
  const external = displays.find((display) => !operatorDisplay || display.id !== operatorDisplay.id);
  return (external || operatorDisplay || displays[0]).id;
}

function positionStageWindow(window, displayId) {
  if (!window || window.isDestroyed() || displayId == null) return;
  const display = getDisplayById(displayId);
  if (!boundsEqual(window.getBounds(), display.bounds)) window.setBounds(display.bounds);
  if (!window.isFullScreen()) window.setFullScreen(true);
  if (!window.isAlwaysOnTop()) window.setAlwaysOnTop(true, 'screen-saver');
}

function createStageWindow(displayId) {
  if (stageWindow && !stageWindow.isDestroyed()) {
    positionStageWindow(stageWindow, displayId);
    return stageWindow;
  }
  stageWindow = new BrowserWindow({
    show: false,
    frame: false,
    backgroundColor: '#080d12',
    webPreferences: {
      backgroundThrottling: false,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload-stage.js')
    }
  });
  hardenWindow(stageWindow);
  stageWindow.loadFile(path.join(__dirname, 'src', 'stage-display.html'));
  stageWindow.once('ready-to-show', () => {
    positionStageWindow(stageWindow, displayId);
    if (latestStageState) stageWindow.webContents.send('stage:state', latestStageState);
    if (stageVisible) stageWindow.showInactive();
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.focus();
  });
  stageWindow.on('closed', () => {
    stageWindow = null;
    if (stageVisible) {
      stageVisible = false;
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('stage:event', { type: 'stage-hidden' });
      }
    }
  });
  return stageWindow;
}

function showStageWindow(target) {
  const displays = screen.getAllDisplays();
  const displayId = resolveStageDisplayId(target ?? stageDisplayTarget, displays);
  if (displayId == null) return;
  stageVisible = true;
  stageDisplayTarget = String(displayId);
  const window = createStageWindow(displayId);
  positionStageWindow(window, displayId);
  if (window.isReadyToShow()) window.showInactive();
  if (latestStageState && !window.webContents.isLoading()) {
    window.webContents.send('stage:state', latestStageState);
  }
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.focus();
}

function hideStageWindow() {
  stageVisible = false;
  if (stageWindow && !stageWindow.isDestroyed()) stageWindow.hide();
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('stage:event', { type: 'stage-hidden' });
  }
}

function sanitizeStageState(value = {}) {
  const text = (input, max = 50000) => String(input || '').slice(0, max);
  return {
    active: value.active === true,
    panic: value.panic === true,
    section: text(value.section, 64),
    itemTitle: text(value.itemTitle, 1000),
    currentLabel: text(value.currentLabel, 1000),
    currentText: text(value.currentText),
    currentNotes: text(value.currentNotes, 20000),
    nextLabel: text(value.nextLabel, 1000),
    nextText: text(value.nextText)
  };
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
  if (stageVisible) {
    const previousStageTarget = stageDisplayTarget;
    const stageDisplayId = resolveStageDisplayId(stageDisplayTarget, displays);
    if (stageDisplayId != null) {
      stageDisplayTarget = String(stageDisplayId);
      positionStageWindow(stageWindow, stageDisplayId);
      if (stageDisplayTarget !== previousStageTarget) {
        persistStageDisplayTarget(stageDisplayTarget).catch((error) => {
          logError('Failed to save fallback stage display target', error);
        });
      }
    }
  }
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('display:changed', {
      target: programDisplayTarget,
      resolvedTarget: resolved.target,
      resolvedDisplayIds: resolved.ids
    });
    mainWindow.webContents.send('stage:event', {
      type: 'stage-display-changed',
      displayTarget: stageDisplayTarget,
      visible: stageVisible
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
    if (!entry.isFile()) continue;
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
      type,
      fileUrl: `worship-media://library/${path.relative(options.mediaRoot || root, fullPath)
        .replace(/\\/g, '/')
        .split('/')
        .map(encodeURIComponent)
        .join('/')}`,
      size
    });
  }
}

async function listLibraryItems(options = {}) {
  const scope = options.scope || 'background';
  const baseFolder = await ensureLibraryFolder();
  let root = baseFolder;
  const walkOptions = {};
  walkOptions.mediaRoot = baseFolder;
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
  return { folder: 'Media Library', items };
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
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload-library.js')
    }
  });
  hardenWindow(libraryWindow);
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
  if (!relativePath) return '';
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

async function resolveContainedFile(root, relativePath) {
  const cleaned = sanitizeRelativePath(relativePath);
  if (!cleaned) return null;
  try {
    const [realRoot, realCandidate] = await Promise.all([
      fs.realpath(root),
      fs.realpath(path.join(root, cleaned))
    ]);
    const relative = path.relative(realRoot, realCandidate);
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) return null;
    const stat = await fs.stat(realCandidate);
    return stat.isFile() ? realCandidate : null;
  } catch (error) {
    return null;
  }
}

function ensureUniqueRelativePath(relativePath, used) {
  if (!relativePath) return '';
  let candidate = relativePath.replace(/\\/g, '/');
  if (!used.has(candidate)) {
    used.add(candidate);
    return candidate;
  }
  const dir = path.posix.dirname(candidate);
  const ext = path.posix.extname(candidate);
  const base = path.posix.basename(candidate, ext);
  let counter = 1;
  while (used.has(candidate)) {
    const name = `${base}_${counter}${ext}`;
    candidate = dir === '.' ? name : path.posix.join(dir, name);
    counter += 1;
  }
  used.add(candidate);
  return candidate;
}

async function resolveAssetPath(assetPath) {
  if (!assetPath || typeof assetPath !== 'string') {
    return null;
  }
  if (assetPath.startsWith('library/') || assetPath.startsWith('library\\')) {
    const cleaned = sanitizeRelativePath(assetPath.replace(/^library[\\/]/, ''));
    if (!cleaned) return null;
    return await resolveContainedFile(getLibraryFolder(), cleaned)
      || resolveContainedFile(getBundledLibraryFolder(), cleaned);
  }
  if (assetPath.startsWith('session/')) {
    return sessionAssetPaths.get(assetPath.slice('session/'.length)) || null;
  }
  return null;
}

async function collectV3Session(project) {
  const projectCopy = project ? JSON.parse(JSON.stringify(project)) : {};
  const assets = [];
  const bySource = new Map();
  const addAsset = async (assetPath, mediaType) => {
    if (!assetPath) return null;
    const resolvedPath = await resolveAssetPath(assetPath);
    if (!resolvedPath) return null;
    const sourceKey = path.resolve(resolvedPath).toLowerCase();
    if (bySource.has(sourceKey)) return bySource.get(sourceKey).id;
    const descriptor = await describeAsset(resolvedPath, mediaType);
    const asset = { ...descriptor, sourcePath: resolvedPath };
    assets.push(asset);
    bySource.set(sourceKey, asset);
    return descriptor.id;
  };
  for (const song of Object.values(projectCopy.songs || {})) {
    if (!song?.background?.path) continue;
    const id = await addAsset(song.background.path, song.background.type);
    if (id) song.background.path = `session/${id}`;
  }
  for (const slide of projectCopy.announcements?.slides || []) {
    if (!slide?.mediaPath) continue;
    const id = await addAsset(slide.mediaPath, 'image');
    if (id) slide.mediaPath = `session/${id}`;
  }
  for (const slide of projectCopy.timer?.slides || []) {
    if (!slide?.mediaPath) continue;
    const id = await addAsset(slide.mediaPath, slide.mediaType);
    if (id) slide.mediaPath = `session/${id}`;
  }
  return { project: projectCopy, assets };
}

async function statIfExists(filePath) {
  try {
    return await fs.stat(filePath);
  } catch (error) {
    return null;
  }
}

async function _collectLegacySessionAssets(project, options = {}) {
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
    const resolvedPath = await resolveAssetPath(sourcePath, options);
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

async function _restoreLegacySessionAssets(project, assets) {
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

function enqueueRecoveryOperation(operation) {
  const result = recoveryWriteQueue.then(operation, operation);
  recoveryWriteQueue = result.then(() => undefined, () => undefined);
  return result;
}

function enqueueHistoryOperation(operation) {
  const result = historyWriteQueue.then(operation, operation);
  historyWriteQueue = result.then(() => undefined, () => undefined);
  return result;
}

function enqueueSongLibraryOperation(operation) {
  const result = songLibraryWriteQueue.then(operation, operation);
  songLibraryWriteQueue = result.then(() => undefined, () => undefined);
  return result;
}

function enqueuePreferenceOperation(operation) {
  const result = preferenceWriteQueue.then(operation, operation);
  preferenceWriteQueue = result.then(() => undefined, () => undefined);
  return result;
}

async function persistStageDisplayTarget(target) {
  if (!_persistenceStores) return;
  await enqueuePreferenceOperation(() => _persistenceStores.preferences.update((preferences) => ({
    ...preferences,
    stageDisplayTarget: target
  })));
}

async function clearRecoverySnapshot() {
  if (!_persistenceStores) return;
  await enqueueRecoveryOperation(() => _persistenceStores.recovery.replace(null));
}

async function refreshRecentSessions({ rebuildMenu = true } = {}) {
  if (!_persistenceStores) return;
  const history = await enqueueHistoryOperation(() => _persistenceStores.history.read());
  recentSessionEntries = history.entries;
  if (rebuildMenu && app.isReady()) createAppMenu();
}

async function recordRecentSession(filePath, project) {
  if (!_persistenceStores) return;
  try {
    const entry = buildHistoryEntry(filePath, project);
    const history = await enqueueHistoryOperation(() => {
      return _persistenceStores.history.update((current) => upsertHistory(current, entry));
    });
    if (stageWindow && !stageWindow.isDestroyed()) {
      stageWindow.close();
    }
    recentSessionEntries = history.entries;
    if (app.isReady()) createAppMenu();
  } catch (error) {
    logError('Failed to update recent projects', error);
  }
}

async function forgetRecentSession(id) {
  if (!_persistenceStores) return;
  try {
    const history = await enqueueHistoryOperation(() => {
      return _persistenceStores.history.update((current) => removeHistoryEntry(current, id));
    });
    recentSessionEntries = history.entries;
    if (app.isReady()) createAppMenu();
  } catch (error) {
    logError('Failed to remove recent project', error);
  }
}

async function readSessionFromPath(filePath) {
  const cacheKey = crypto.createHash('sha256').update(filePath).digest('hex').slice(0, 24);
  const cacheDir = path.join(app.getPath('userData'), 'session-cache', cacheKey);
  const opened = await readSession({ filePath, cacheDir });
  activeSessionPath = filePath;
  sessionAssetPaths = opened.assetPaths;
  projectDirty = false;
  return { filePath, project: opened.project };
}

async function prepareSongForLibrary(song) {
  const prepared = structuredClone(song);
  const backgroundPath = prepared?.background?.path || '';
  if (!backgroundPath) return prepared;
  if (backgroundPath.startsWith('library/') || backgroundPath.startsWith('library\\')) {
    if (!await resolveAssetPath(backgroundPath)) {
      throw new Error('Song background is unavailable and cannot be added to the library.');
    }
    return prepared;
  }
  if (!backgroundPath.startsWith('session/')) {
    throw new Error('Song background must come from the media library or current session.');
  }
  const sourcePath = await resolveAssetPath(backgroundPath);
  if (!sourcePath) throw new Error('Song background is unavailable and cannot be added to the library.');
  const descriptor = await describeAsset(sourcePath, prepared.background.type);
  const folder = prepared.background.type === 'video' ? 'Videos' : 'Images';
  const relativePath = path.posix.join(folder, 'Song Library', `${descriptor.id}${descriptor.extension}`);
  const targetPath = path.join(getLibraryFolder(), ...relativePath.split('/'));
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  if (!await statIfExists(targetPath)) await fs.copyFile(sourcePath, targetPath);
  prepared.background.path = `library/${relativePath}`;
  return prepared;
}

ipcMain.handle('app:get-version', (event) => {
  assertOperatorSender(event);
  return app.getVersion();
});
ipcMain.handle('file:register-drop', async (event, filePath) => {
  assertOperatorSender(event);
  if (!filePath || typeof filePath !== 'string') throw new Error('Dropped file path is invalid.');
  const stat = await fs.stat(filePath);
  if (!stat.isFile()) throw new Error('Dropped item is not a regular file.');
  return issueFileToken(event, filePath);
});

ipcMain.handle('project:save-file', async (event, project) => {
  assertOperatorSender(event);
  const saveOperation = async () => {
    try {
    let targetPath = activeSessionPath;
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
    const prepared = await collectV3Session(project);
    await writeSession({
      targetPath,
      appVersion: app.getVersion(),
      project: prepared.project,
      assets: prepared.assets
    });
    activeSessionPath = targetPath;
    await recordRecentSession(targetPath, prepared.project);
    await clearRecoverySnapshot().catch((error) => logError('Failed to clear recovery snapshot after save', error));
    return targetPath;
    } catch (error) {
      logError('Failed to save session file', error);
      dialog.showErrorBox('Save Error', String(error?.message || error));
      return null;
    }
  };
  const result = sessionWriteQueue.then(saveOperation, saveOperation);
  sessionWriteQueue = result.then(() => undefined, () => undefined);
  return result;
});

ipcMain.handle('project:confirm-transition', async (event) => {
  assertOperatorSender(event);
  if (!projectDirty) return 'discard';
  const result = await dialog.showMessageBox(mainWindow, {
    type: 'warning',
    title: 'Unsaved Changes',
    message: 'Save changes before continuing?',
    detail: 'Unsaved changes will be lost if you discard them.',
    buttons: ['Save', 'Discard', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
    noLink: true
  });
  return ['save', 'discard', 'cancel'][result.response] || 'cancel';
});

ipcMain.on('project:set-dirty', (event, dirty) => {
  assertOperatorSender(event);
  projectDirty = dirty === true;
});

ipcMain.handle('project:write-recovery', async (event, project) => {
  assertOperatorSender(event);
  if (!_persistenceStores) return false;
  const migratedProject = migrateProject(project);
  await enqueueRecoveryOperation(() => _persistenceStores.recovery.replace({
    version: 1,
    updatedAt: new Date().toISOString(),
    sourceSessionPath: activeSessionPath,
    project: migratedProject
  }));
  return true;
});

ipcMain.handle('project:clear-recovery', async (event) => {
  assertOperatorSender(event);
  await clearRecoverySnapshot();
  return true;
});

ipcMain.handle('project:check-recovery', async (event) => {
  assertOperatorSender(event);
  if (!_persistenceStores) return null;
  const recovery = await enqueueRecoveryOperation(() => _persistenceStores.recovery.read());
  if (!recovery) return null;
  const sourceName = recovery.sourceSessionPath ? path.basename(recovery.sourceSessionPath) : 'an unsaved session';
  const result = await dialog.showMessageBox(getDialogParent() || undefined, {
    type: 'question',
    title: 'Recover Unsaved Session',
    message: 'Worship Presenter found unsaved work.',
    detail: `Recover changes from ${sourceName}?`,
    buttons: ['Recover', 'Discard'],
    defaultId: 0,
    cancelId: 0,
    noLink: true
  });
  if (result.response === 1) {
    await clearRecoverySnapshot();
    return null;
  }

  let filePath = null;
  activeSessionPath = null;
  sessionAssetPaths = new Map();
  if (recovery.sourceSessionPath) {
    try {
      const source = await readSessionFromPath(recovery.sourceSessionPath);
      filePath = source.filePath;
    } catch (error) {
      logError('Failed to reopen the recovery source session', error);
      activeSessionPath = null;
      sessionAssetPaths = new Map();
    }
  }
  projectDirty = true;
  return { filePath, project: recovery.project, recoveredAt: recovery.updatedAt };
});

ipcMain.on('project:complete-close', (event, saved) => {
  assertOperatorSender(event);
  closePromptOpen = false;
  if (!saved || !mainWindow || mainWindow.isDestroyed()) return;
  projectDirty = false;
  allowMainWindowClose = true;
  mainWindow.close();
});

ipcMain.handle('project:open-file', async (event) => {
  try {
    assertOperatorSender(event);
    const result = await showOpenDialogSafe({
      title: 'Open Session',
      properties: ['openFile'],
      filters: [{ name: 'Worship Presenter Session', extensions: ['wpjson'] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    const payload = await readSessionFromPath(result.filePaths[0]);
    await recordRecentSession(payload.filePath, payload.project);
    await clearRecoverySnapshot().catch((error) => logError('Failed to clear recovery snapshot after open', error));
    return payload;
  } catch (error) {
    logError('Failed to open session file', error);
    dialog.showErrorBox('Open Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('assets:check', async (_event, payload = {}) => {
  assertOperatorSender(_event);
  const paths = Array.isArray(payload.paths) ? payload.paths : [];
  const missing = [];
  for (const assetPath of paths) {
    const resolved = await resolveAssetPath(assetPath);
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

ipcMain.handle('lyrics:pick', async (event) => {
  try {
    assertOperatorSender(event);
    const result = await showOpenDialogSafe({
      title: 'Import Lyrics',
      properties: ['openFile'],
      filters: [{ name: 'Text Files', extensions: ['txt'] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return issueFileToken(event, result.filePaths[0]);
  } catch (error) {
    logError('Failed to open lyrics picker', error);
    dialog.showErrorBox('Lyrics Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('file:read-text', async (event, token) => {
  try {
    assertOperatorSender(event);
    const filePath = consumeFileToken(event, token, new Set(['.txt']));
    return await fs.readFile(filePath, 'utf8');
  } catch (error) {
    logError('Failed to read text file', error);
    dialog.showErrorBox('Read Error', String(error?.message || error));
    return '';
  }
});

ipcMain.handle('project:list-recent', async (event) => {
  assertOperatorSender(event);
  return publicHistoryEntries({ version: 1, entries: recentSessionEntries });
});

ipcMain.handle('project:open-recent', async (event, id) => {
  try {
    assertOperatorSender(event);
    if (typeof id !== 'string' || !/^[a-f0-9]{32}$/.test(id)) throw new Error('Recent project ID is invalid.');
    const entry = recentSessionEntries.find((item) => item.id === id);
    if (!entry) throw new Error('Recent project is no longer available.');
    const payload = await readSessionFromPath(entry.filePath);
    await recordRecentSession(payload.filePath, payload.project);
    await clearRecoverySnapshot().catch((error) => logError('Failed to clear recovery snapshot after open', error));
    return payload;
  } catch (error) {
    if (typeof id === 'string') await forgetRecentSession(id);
    logError('Failed to open recent session file', error);
    dialog.showErrorBox('Open Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('song-library:list', async (event, query) => {
  assertOperatorSender(event);
  if (!_songLibrary) return [];
  return enqueueSongLibraryOperation(() => _songLibrary.list(query));
});

ipcMain.handle('song-library:save', async (event, payload = {}) => {
  let preparedSong = null;
  try {
    assertOperatorSender(event);
    if (!_songLibrary) throw new Error('Song library is unavailable.');
    preparedSong = await prepareSongForLibrary(payload.song);
    const result = await enqueueSongLibraryOperation(() => {
      return _songLibrary.save(preparedSong, { force: payload.force === true });
    });
    return { ok: true, ...result };
  } catch (error) {
    if (error instanceof SongLibraryConflictError) {
      const result = await dialog.showMessageBox(getDialogParent() || undefined, {
        type: 'warning',
        title: 'Library Song Changed',
        message: 'This song has a newer revision in the library.',
        detail: 'Replace the newer library revision with the version in this project?',
        buttons: ['Replace Library Version', 'Cancel'],
        defaultId: 1,
        cancelId: 1,
        noLink: true
      });
      if (result.response !== 0) {
        return { ok: false, conflict: true, currentRevision: error.currentRevision };
      }
      try {
        const forced = await enqueueSongLibraryOperation(() => _songLibrary.save(preparedSong, { force: true }));
        return { ok: true, ...forced };
      } catch (forceError) {
        logError('Failed to replace song library revision', forceError);
        return { ok: false, error: String(forceError?.message || forceError) };
      }
    }
    logError('Failed to save song to library', error);
    return { ok: false, error: String(error?.message || error) };
  }
});

ipcMain.handle('song-library:instantiate', async (event, id) => {
  assertOperatorSender(event);
  if (!_songLibrary) throw new Error('Song library is unavailable.');
  return enqueueSongLibraryOperation(() => _songLibrary.instantiate(id));
});

ipcMain.handle('song-library:remove', async (event, id) => {
  assertOperatorSender(event);
  if (!_songLibrary) throw new Error('Song library is unavailable.');
  const item = await enqueueSongLibraryOperation(() => _songLibrary.getItem(id));
  const result = await dialog.showMessageBox(getDialogParent() || undefined, {
    type: 'warning',
    title: 'Delete Library Song',
    message: `Delete “${item.title}” from the song library?`,
    detail: 'Songs already added to a project will not be removed.',
    buttons: ['Delete', 'Cancel'],
    defaultId: 1,
    cancelId: 1,
    noLink: true
  });
  if (result.response !== 0) return { deleted: false };
  await enqueueSongLibraryOperation(() => _songLibrary.remove(id));
  return { deleted: true };
});

ipcMain.handle('session:reset', async (event) => {
  assertOperatorSender(event);
  activeSessionPath = null;
  sessionAssetPaths = new Map();
  projectDirty = false;
  await clearRecoverySnapshot().catch((error) => logError('Failed to clear recovery snapshot after reset', error));
  return true;
});

ipcMain.handle('pptx:read-slides', async (event, token) => {
  assertOperatorSender(event);
  if (!token || typeof token !== 'string') {
    return [];
  }
  const filePath = consumeFileToken(event, token, new Set(['.pptx']));
  const stat = await fs.stat(filePath);
  if (!stat.isFile() || stat.size > MAX_PPTX_BYTES) {
    throw new Error('PowerPoint file is too large or is not a regular file.');
  }
  const data = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(data);
  const entries = Object.keys(zip.files).map((entry) => ({
    entry,
    normalized: entry.replace(/\\/g, '/')
  }));
  let slidePaths = entries.filter(({ normalized }) => /ppt\/slides\/slide\d+\.xml$/i.test(normalized));
  if (slidePaths.length === 0) {
    slidePaths = entries.filter(
      ({ normalized }) => /ppt\/slides\/[^/]+\.xml$/i.test(normalized) && !/ppt\/slides\/_rels\//i.test(normalized)
    );
  }
  if (slidePaths.length > MAX_PPTX_SLIDES) {
    throw new Error('PowerPoint file contains too many slides.');
  }
  slidePaths.sort((a, b) => {
    const aMatch = a.normalized.match(/slide(\d+)\.xml$/i);
    const bMatch = b.normalized.match(/slide(\d+)\.xml$/i);
    const aIndex = aMatch ? Number(aMatch[1]) : 0;
    const bIndex = bMatch ? Number(bMatch[1]) : 0;
    return aIndex - bIndex;
  });
  const slides = [];
  let totalXmlBytes = 0;
  for (const slidePath of slidePaths) {
    const slideFile = zip.file(slidePath.entry);
    const declaredSize = Number(slideFile?._data?.uncompressedSize);
    if (Number.isFinite(declaredSize) && declaredSize >= 0) {
      totalXmlBytes += declaredSize;
      if (totalXmlBytes > MAX_PPTX_XML_BYTES) {
        throw new Error('PowerPoint slide content is too large.');
      }
    }
    const xml = await slideFile.async('string');
    if (!Number.isFinite(declaredSize)) {
      totalXmlBytes += Buffer.byteLength(xml, 'utf8');
    }
    if (totalXmlBytes > MAX_PPTX_XML_BYTES) {
      throw new Error('PowerPoint slide content is too large.');
    }
    slides.push(xml);
  }
  return slides;
});

ipcMain.handle('media:pick', async (event) => {
  try {
    assertOperatorSender(event);
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
    return issueFileToken(event, result.filePaths[0]);
  } catch (error) {
    logError('Failed to open media picker', error);
    dialog.showErrorBox('Media Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('announcement:pick', async (event) => {
  try {
    assertOperatorSender(event);
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
    return issueFileToken(event, result.filePaths[0]);
  } catch (error) {
    logError('Failed to open announcement picker', error);
    dialog.showErrorBox('Announcement Picker Error', String(error?.message || error));
    return null;
  }
});

ipcMain.handle('library:import', async (event, payload) => {
  assertOperatorSender(event);
  const token = typeof payload === 'string' ? payload : payload && payload.sourcePath;
  const scope = payload && typeof payload === 'object' ? payload.scope : null;
  if (!token) {
    return null;
  }
  const sourcePath = consumeFileToken(event, token, new Set([...IMAGE_EXTS, ...VIDEO_EXTS]));
  const libraryDir = await ensureLibraryFolder();
  const ext = path.extname(sourcePath);
  const type = mediaTypeFromExt(ext);
  const subfolder = libraryFolderForScope(scope, type);
  const targetDir = path.join(libraryDir, subfolder);
  await ensureDir(targetDir);
  const descriptor = await describeAsset(sourcePath, type);
  const fileName = `${descriptor.id}${descriptor.extension}`;
  const target = path.join(targetDir, fileName);
  if (!await statIfExists(target)) await fs.copyFile(sourcePath, target);
  const relative = path.relative(libraryDir, target).replace(/\\/g, '/');
  return {
    relativePath: path.posix.join('library', relative)
  };
});

ipcMain.handle('content-pack:import', async (event) => {
  try {
    assertLibrarySender(event);
    const result = await showOpenDialogSafe({
      title: 'Import Content Pack (Folder)',
      properties: ['openDirectory']
    });
    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true };
    }
    return await importContentPack(result.filePaths[0]);
  } catch (error) {
    logError('Failed to import content pack', error);
    return { error: 'Failed to import content pack.' };
  }
});

ipcMain.handle('announcement:import', async (event, token) => {
  assertOperatorSender(event);
  const sourcePath = consumeFileToken(event, token, IMAGE_EXTS);
  const libraryDir = await ensureLibraryFolder();
  const targetDir = path.join(libraryDir, 'Announcements');
  await ensureDir(targetDir);
  const descriptor = await describeAsset(sourcePath, 'image');
  const fileName = `${descriptor.id}${descriptor.extension}`;
  const target = path.join(targetDir, fileName);
  if (!await statIfExists(target)) await fs.copyFile(sourcePath, target);
  const relative = path.relative(libraryDir, target).replace(/\\/g, '/');
  return {
    relativePath: path.posix.join('library', relative)
  };
});

ipcMain.handle('display:list', async (event) => {
  assertOperatorSender(event);
  const operatorDisplay = getOperatorDisplay();
  return screen.getAllDisplays().map((display) => ({
    id: display.id,
    label: `${display.id} - ${display.bounds.width}x${display.bounds.height}`,
    bounds: display.bounds,
    scaleFactor: display.scaleFactor,
    isOperator: operatorDisplay?.id === display.id
  }));
});

ipcMain.on('program:show', (event, displayId) => {
  assertOperatorSender(event);
  showProgramWindows(displayId);
});

ipcMain.on('program:hide', (event) => {
  assertOperatorOrProgramSender(event);
  hideProgramWindows();
});

ipcMain.on('program:set-display', (event, displayId) => {
  assertOperatorSender(event);
  const displays = screen.getAllDisplays();
  setProgramDisplayPreference(displayId, displays);
  if (programVisible) {
    showProgramWindows(programDisplayTarget, { displays, updatePreference: false });
  }
});

ipcMain.on('program:toggle-fullscreen', (event) => {
  assertOperatorSender(event);
  programWindows.forEach((window) => {
    if (window && !window.isDestroyed()) {
      window.setFullScreen(!window.isFullScreen());
    }
  });
});

ipcMain.on('program:state', (event, state) => {
  assertOperatorSender(event);
  programWindows.forEach((window) => {
    if (window && !window.isDestroyed()) {
      window.webContents.send('program:state', state);
    }
  });
});

ipcMain.on('program:event', (event, payload) => {
  assertProgramSender(event);
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('program:event', payload);
  }
});

ipcMain.handle('stage:get-config', async (event) => {
  assertOperatorSender(event);
  return { displayTarget: stageDisplayTarget, visible: stageVisible };
});

ipcMain.on('stage:show', (event, displayId) => {
  assertOperatorSender(event);
  const resolved = resolveStageDisplayId(displayId);
  if (resolved == null) return;
  stageDisplayTarget = String(resolved);
  persistStageDisplayTarget(stageDisplayTarget).catch((error) => logError('Failed to save stage display target', error));
  showStageWindow(stageDisplayTarget);
});

ipcMain.on('stage:hide', (event) => {
  assertOperatorOrStageSender(event);
  hideStageWindow();
});

ipcMain.on('stage:set-display', (event, displayId) => {
  assertOperatorSender(event);
  const resolved = resolveStageDisplayId(displayId);
  if (resolved == null) return;
  stageDisplayTarget = String(resolved);
  persistStageDisplayTarget(stageDisplayTarget).catch((error) => logError('Failed to save stage display target', error));
  if (stageVisible) showStageWindow(stageDisplayTarget);
});

ipcMain.on('stage:state', (event, state) => {
  assertOperatorSender(event);
  latestStageState = sanitizeStageState(state);
  if (stageWindow && !stageWindow.isDestroyed() && !stageWindow.webContents.isLoading()) {
    stageWindow.webContents.send('stage:state', latestStageState);
  }
});

ipcMain.on('edit:command', (event, command) => {
  assertOperatorSender(event);
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

ipcMain.handle('library:list', async (event, options = {}) => {
  assertLibrarySender(event);
  return listLibraryItems(options);
});

ipcMain.on('library:open', (event, options = {}) => {
  assertOperatorSender(event);
  const scope = options.scope || 'background';
  createLibraryWindow(scope);
});

ipcMain.on('library:close', (event) => {
  assertLibrarySender(event);
  if (libraryWindow && !libraryWindow.isDestroyed()) {
    libraryWindow.close();
  }
});

ipcMain.on('library:select', (event, payload) => {
  assertLibrarySender(event);
  const sourcePath = payload && payload.sourcePath ? payload.sourcePath : payload;
  const scope = payload && payload.scope ? payload.scope : libraryWindowScope || 'background';
  const relativePath = typeof sourcePath === 'string'
    ? sanitizeRelativePath(sourcePath.replace(/^library[\\/]/, ''))
    : '';
  if (!relativePath || !['background', 'announcements', 'timer'].includes(scope)) return;
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('library:selected', { scope, sourcePath: `library/${relativePath}` });
  }
  if (libraryWindow && !libraryWindow.isDestroyed()) {
    libraryWindow.close();
  }
});

app.whenReady().then(async () => {
  _persistenceStores = createPersistenceStores(app.getPath('userData'));
  try {
    await _persistenceStores.initialize();
    const preferences = await _persistenceStores.preferences.read();
    stageDisplayTarget = preferences.stageDisplayTarget;
    _songLibrary = createSongLibrary(_persistenceStores);
    await refreshRecentSessions({ rebuildMenu: false });
  } catch (error) {
    logError('Failed to initialize local persistence stores', error);
  }
  protocol.handle('worship-media', async (request) => {
    try {
      const url = new URL(request.url);
      const value = decodeURIComponent(url.pathname.replace(/^\//, ''));
      let filePath = null;
      if (url.hostname === 'asset' && /^[a-f0-9]{64}$/.test(value)) {
        filePath = sessionAssetPaths.get(value) || null;
      } else if (url.hostname === 'library') {
        const relativePath = sanitizeRelativePath(value);
        const extension = path.extname(relativePath).toLowerCase();
        if (relativePath && (IMAGE_EXTS.has(extension) || VIDEO_EXTS.has(extension))) {
          filePath = await resolveContainedFile(getLibraryFolder(), relativePath)
            || await resolveContainedFile(getBundledLibraryFolder(), relativePath);
        }
      }
      if (!filePath) return new Response('Not found', { status: 404 });
      return net.fetch(pathToFileURL(filePath).toString());
    } catch (error) {
      return new Response('Invalid media URL', { status: 400 });
    }
  });
  session.defaultSession.setPermissionCheckHandler(() => false);
  session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false));
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
