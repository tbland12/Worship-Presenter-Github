const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const read = (relativePath) => fs.readFileSync(path.join(root, relativePath), 'utf8');

test('renderer code does not use executable HTML interpolation', () => {
  const rendererSources = [
    'src/operator.js',
    'src/library.js',
    'src/program.js',
    'src/stage-display.js',
    'src/shared/stage.js',
    'src/shared/ui.js'
  ].map(read);
  rendererSources.forEach((source) => assert.doesNotMatch(source, /\.innerHTML\s*=/));
});

test('operator and library use the shared accessible UI system', () => {
  const operatorHtml = read('src/operator.html');
  const libraryHtml = read('src/library.html');
  const uiCss = read('src/shared/ui.css');
  const uiJs = read('src/shared/ui.js');
  const rendererJs = `${read('src/operator.js')}\n${read('src/library.js')}`;
  assert.match(operatorHtml, /id="appearance-toggle"/);
  assert.match(operatorHtml, /id="save-status"[^>]*role="status"/);
  assert.match(operatorHtml, /id="toast-region"[^>]*aria-live="polite"/);
  assert.match(libraryHtml, /id="library-search"/);
  assert.match(uiCss, /:root\[data-theme="light"\]/);
  assert.match(uiCss, /:focus-visible/);
  assert.match(uiCss, /prefers-reduced-motion/);
  assert.match(uiJs, /worship-presenter-theme/);
  assert.doesNotMatch(rendererJs, /window\.alert/);
});

test('preview carousel and live slide visibility are wired', () => {
  const operatorHtml = read('src/operator.html');
  const operatorCss = read('src/operator.css');
  const operatorJs = read('src/operator.js');
  assert.match(operatorHtml, /id="preview-card-previous"/);
  assert.match(operatorHtml, /id="preview-card-next"/);
  assert.match(operatorCss, /preview-card-adjacent/);
  assert.match(operatorCss, /mask-image:\s*linear-gradient/);
  assert.match(operatorCss, /preview-wrap::before/);
  assert.match(operatorCss, /data-theme="light"[^}]*presenter-controls button\.primary/);
  assert.match(operatorJs, /DEFAULT_PREVIEW_SCALE = 0\.7/);
  assert.match(operatorJs, /function getAdjacentPreviewSelection/);
  assert.match(operatorJs, /function scrollLiveSlideIntoView\(\)/);
  assert.match(operatorJs, /list\.scrollTo\(\{ top: Math\.max\(0, centeredTop\), behavior: 'smooth' \}\)/);
});

test('all BrowserWindows disable Node and enable isolation and sandboxing', () => {
  const main = read('main.js');
  const browserWindowCount = (main.match(/new BrowserWindow\s*\(/g) || []).length;
  assert.equal(browserWindowCount, 4);
  assert.equal((main.match(/nodeIntegration:\s*false/g) || []).length, browserWindowCount);
  assert.equal((main.match(/contextIsolation:\s*true/g) || []).length, browserWindowCount);
  assert.equal((main.match(/sandbox:\s*true/g) || []).length, browserWindowCount);
  assert.doesNotMatch(main, /nodeIntegration:\s*true/);
});

test('renderer documents enforce a restrictive script policy', () => {
  ['src/operator.html', 'src/program.html', 'src/library.html', 'src/stage-display.html'].forEach((file) => {
    const html = read(file);
    assert.match(html, /Content-Security-Policy/);
    assert.match(html, /script-src 'self'/);
    assert.match(html, /object-src 'none'/);
    assert.match(html, /connect-src 'none'/);
    assert.doesNotMatch(html, /script-src[^;]*'unsafe-inline'/);
  });
});

test('renderer bridge has no require fallback', () => {
  const bridge = read('src/shared/bridge.js');
  assert.doesNotMatch(bridge, /window\.require|require\s*\(/);
});

test('navigation, popup, webview, and permission guards are installed', () => {
  const main = read('main.js');
  assert.match(main, /setWindowOpenHandler/);
  assert.match(main, /will-navigate/);
  assert.match(main, /will-attach-webview/);
  assert.match(main, /setPermissionCheckHandler/);
  assert.match(main, /setPermissionRequestHandler/);
});

test('session IPC does not expose renderer-controlled filesystem paths', () => {
  const main = read('main.js');
  const preload = read('preload.js');
  assert.doesNotMatch(main, /ipcMain\.handle\('project:(?:new|open|load|save)'/);
  assert.doesNotMatch(main, /ipcMain\.handle\('media:import'/);
  assert.doesNotMatch(preload, /getPathForFile\s*:|projectFolder.*project:save-file/);
  assert.match(main, /sanitizeRelativePath\(assetPath\.replace/);
});

test('session saves are serialized and dirty state is protected', () => {
  const main = read('main.js');
  const operator = read('src/operator.js');
  assert.match(main, /sessionWriteQueue\.then\(saveOperation, saveOperation\)/);
  assert.match(main, /project:confirm-transition/);
  assert.match(main, /project:save-before-close/);
  assert.match(operator, /while \(saveRequested\)/);
  assert.match(operator, /state\.autoSaveEnabled = Boolean\(filePath\)/);
  assert.match(operator, /setProjectDirty\(true\)/);
});

test('recovery and recent-project workflows keep filesystem paths in the main process', () => {
  const main = read('main.js');
  const preload = read('preload.js');
  const operator = read('src/operator.js');
  assert.match(main, /sourceSessionPath: activeSessionPath/);
  assert.match(main, /recentSessionEntries\.find\(\(item\) => item\.id === id\)/);
  assert.match(main, /publicHistoryEntries/);
  assert.doesNotMatch(preload, /openRecentProject: \(filePath\)/);
  assert.match(operator, /scheduleRecoverySnapshot\(\)/);
  assert.match(operator, /restoreRecoveryAtStartup\(\)/);
});

test('windows use separate least-privilege preloads', () => {
  const main = read('main.js');
  const operatorPreload = read('preload.js');
  const libraryPreload = read('preload-library.js');
  const programPreload = read('preload-program.js');
  const stagePreload = read('preload-stage.js');
  assert.match(main, /preload-program\.js/);
  assert.match(main, /preload-library\.js/);
  assert.match(main, /preload-stage\.js/);
  assert.doesNotMatch(libraryPreload, /project:save-file|program:show|file:register-drop/);
  assert.doesNotMatch(programPreload, /project:save-file|library:list|file:register-drop/);
  assert.doesNotMatch(operatorPreload, /['"]library:list['"]|content-pack:import|program:state.*callback/);
  assert.doesNotMatch(stagePreload, /project:|library:|program:|file:register-drop/);
});

test('renderer media access uses the validated custom protocol', () => {
  const main = read('main.js');
  const operator = read('src/operator.js');
  const html = ['src/operator.html', 'src/library.html', 'src/program.html'].map(read).join('\n');
  assert.match(main, /url\.hostname === 'library'/);
  assert.match(main, /sanitizeRelativePath\(value\)/);
  assert.doesNotMatch(operator, /file:\/\//);
  assert.doesNotMatch(html, /(?:img-src|media-src)[^;]*file:/);
});

test('song library uses opaque IDs, accessible dialog controls, and safe DOM rendering', () => {
  const main = read('main.js');
  const operator = read('src/operator.js');
  const html = read('src/operator.html');
  const songLibrary = read('src/main/song-library.js');
  assert.match(html, /<dialog[^>]*id="song-library-dialog"[^>]*aria-labelledby=/);
  assert.match(html, /id="song-library-search"[^>]*type="search"/);
  assert.match(html, /id="add-library-song"/);
  assert.match(main, /song-library:instantiate/);
  assert.match(main, /prepareSongForLibrary/);
  assert.match(songLibrary, /SONG_LIBRARY_ID/);
  assert.match(operator, /songLibraryResults\.replaceChildren\(\)/);
  assert.doesNotMatch(operator, /songLibraryResults\.innerHTML/);
});

test('stage display is isolated and receives bounded current, next, and note state', () => {
  const main = read('main.js');
  const operator = read('src/operator.js');
  const html = read('src/operator.html');
  const stageHtml = read('src/stage-display.html');
  assert.match(html, /id="speaker-notes"[^>]*maxlength="20000"/);
  assert.match(html, /id="announcement-speaker-notes"/);
  assert.match(html, /id="timer-speaker-notes"/);
  assert.match(stageHtml, /id="current-text"/);
  assert.match(stageHtml, /id="notes-text"/);
  assert.match(stageHtml, /id="next-text"/);
  assert.match(main, /function sanitizeStageState/);
  assert.match(main, /stageDisplayTarget: target/);
  assert.match(operator, /function getNextLiveSelection/);
  assert.match(operator, /currentNotes: selection\?\.slide\?\.speakerNotes/);
});
