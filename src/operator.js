import {
  createNewProject,
  createSong,
  createSlide,
  createMediaSlide,
  richTextFromPlain,
  plainFromRichText,
  getDefaultTheme
} from './shared/model.js';
import { StageRenderer } from './shared/stage.js';
import { ensureApiBridge } from './shared/bridge.js';

const elements = {
  newProject: document.getElementById('new-project'),
  openProject: document.getElementById('open-project'),
  saveProject: document.getElementById('save-project'),
  fileMenu: document.getElementById('file-menu'),
  assetWarning: document.getElementById('asset-warning'),
  projectPath: document.getElementById('project-path'),
  addAnnouncement: document.getElementById('add-announcement'),
  announcements: document.getElementById('announcements-list'),
  announcementsEmpty: document.getElementById('announcements-empty'),
  announcementsPreview: document.getElementById('announcements-preview-list'),
  announcementsPreviewEmpty: document.getElementById('announcements-preview-empty'),
  announcementsButton: document.getElementById('select-announcements'),
  addTimerSlide: document.getElementById('add-timer-slide'),
  timerList: document.getElementById('timer-list'),
  timerEmpty: document.getElementById('timer-empty'),
  timerPreviewList: document.getElementById('timer-preview-list'),
  timerPreviewEmpty: document.getElementById('timer-preview-empty'),
  timerButton: document.getElementById('select-timer'),
  addSong: document.getElementById('add-song'),
  setlist: document.getElementById('setlist'),
  setlistEmpty: document.getElementById('setlist-empty'),
  previewStage: document.getElementById('preview-stage'),
  previewSplitter: document.getElementById('preview-splitter'),
  previewCard: document.getElementById('preview-card'),
  previewWrap: document.getElementById('preview-wrap'),
  songEditor: document.getElementById('song-editor'),
  announcementEditor: document.getElementById('announcement-editor'),
  timerEditor: document.getElementById('timer-editor'),
  songTitle: document.getElementById('song-title'),
  importLyrics: document.getElementById('import-lyrics'),
  addSlide: document.getElementById('add-slide'),
  slides: document.getElementById('slides'),
  slidesEmpty: document.getElementById('slides-empty'),
  songContext: document.getElementById('song-context'),
  showTitle: document.getElementById('show-title'),
  showLyrics: document.getElementById('show-lyrics'),
  showFooter: document.getElementById('show-footer'),
  slideLabel: document.getElementById('slide-label'),
  titleText: document.getElementById('title-text'),
  lyricsText: document.getElementById('lyrics-text'),
  footerText: document.getElementById('footer-text'),
  footerAuto: document.getElementById('footer-auto'),
  titleRow: document.getElementById('title-row'),
  lyricsRow: document.getElementById('lyrics-row'),
  footerRow: document.getElementById('footer-row'),
  songStatus: document.getElementById('song-status'),
  backgroundInspector: document.getElementById('background-inspector'),
  openLibrary: document.getElementById('open-library'),
  uploadBackground: document.getElementById('upload-background'),
  backgroundPath: document.getElementById('background-path'),
  announcementInspector: document.getElementById('announcement-inspector'),
  openAnnouncementLibrary: document.getElementById('open-announcement-library'),
  uploadAnnouncement: document.getElementById('upload-announcement'),
  announcementPath: document.getElementById('announcement-path'),
  announcementAuto: document.getElementById('announcement-auto'),
  announcementAdvance: document.getElementById('announcement-advance'),
  announcementLoop: document.getElementById('announcement-loop'),
  timerInspector: document.getElementById('timer-inspector'),
  openTimerLibrary: document.getElementById('open-timer-library'),
  uploadTimerMedia: document.getElementById('upload-timer-media'),
  timerMediaPath: document.getElementById('timer-media-path'),
  timerAutoVideo: document.getElementById('timer-auto-video'),
  timerAutoImages: document.getElementById('timer-auto-images'),
  timerAdvance: document.getElementById('timer-advance'),
  timerAdvanceRow: document.getElementById('timer-advance-row'),
  themeTarget: document.getElementById('theme-target'),
  themeFont: document.getElementById('theme-font'),
  themeBase: document.getElementById('theme-base'),
  themeColor: document.getElementById('theme-color'),
  themeStrokeToggle: document.getElementById('theme-stroke-toggle'),
  themeStroke: document.getElementById('theme-stroke'),
  themeStrokeColor: document.getElementById('theme-stroke-color'),
  themeShadowToggle: document.getElementById('theme-shadow-toggle'),
  themeShadowX: document.getElementById('theme-shadow-x'),
  themeShadowY: document.getElementById('theme-shadow-y'),
  themeShadowBlur: document.getElementById('theme-shadow-blur'),
  themeShadowColor: document.getElementById('theme-shadow-color'),
  strokeOptions: document.getElementById('stroke-options'),
  shadowOptions: document.getElementById('shadow-options'),
  themeToggle: document.getElementById('theme-toggle'),
  themeSection: document.getElementById('theme-section'),
  themeGroup: document.getElementById('theme-group'),
  themePosition: document.getElementById('theme-position'),
  themeDim: document.getElementById('theme-dim'),
  ccliToggle: document.getElementById('ccli-toggle'),
  ccliSection: document.getElementById('ccli-section'),
  ccliGroup: document.getElementById('ccli-group'),
  ccliNumber: document.getElementById('ccli-number'),
  ccliAuthors: document.getElementById('ccli-authors'),
  ccliPublisher: document.getElementById('ccli-publisher'),
  ccliCopyright: document.getElementById('ccli-copyright'),
  hideProgram: document.getElementById('hide-program'),
  goLive: document.getElementById('go-live'),
  prevSlide: document.getElementById('prev-slide'),
  nextSlide: document.getElementById('next-slide'),
  panic: document.getElementById('panic'),
  autoGoLive: document.getElementById('auto-go-live'),
  followLive: document.getElementById('follow-live'),
  displaySelect: document.getElementById('display-select'),
  liveStatus: document.getElementById('live-status'),
  slideContext: document.getElementById('slide-context'),
  mediaContext: document.getElementById('media-context'),
  dropOverlay: document.getElementById('drop-overlay'),
  inspectorEmpty: document.getElementById('inspector-empty')
};

const previewRenderer = new StageRenderer(elements.previewStage);

const apiBridge = ensureApiBridge();
if (!apiBridge) {
  window.setTimeout(() => {
    window.alert('Electron bridge is unavailable. Dialogs will not open. Please restart the app.');
  }, 200);
}

const state = {
  project: createNewProject(),
  projectFolder: null,
  projectFile: null,
  autoSaveEnabled: false,
  selectedSection: 'setlist',
  selectedSongId: null,
  selectedSlideIndex: 0,
  selectedAnnouncementIndex: -1,
  selectedTimerIndex: -1,
  preview: { section: 'setlist', songId: null, slideIndex: 0 },
  live: { section: null, songId: null, slideIndex: 0 },
  panic: false,
  autoGoLive: false,
  followLive: false,
  displayId: null,
  themeTarget: 'lyrics',
  runtimeBackgrounds: {},
  libraryTarget: null
};

let saveTimer = null;
let assetCheckTimer = null;
let assetCheckToken = 0;
let dropCounter = 0;
let internalDragActive = false;
let announcementAdvanceTimer = null;
let timerImageAdvanceTimer = null;
const PREVIEW_BASE = { width: 1920, height: 1080 };
let previewObserver = null;
let previewSplitState = null;
const SECTION_HEADER_REGEX = /^(verse|chorus|bridge|pre-chorus|prechorus|intro|outro|tag|refrain|ending|hook)(\s+\d+)?$/i;

function updatePreviewScale() {
  const stage = elements.previewStage;
  if (!stage) {
    return;
  }
  const width = stage.clientWidth;
  const height = stage.clientHeight;
  if (!width || !height) {
    return;
  }
  const scale = Math.min(1, width / PREVIEW_BASE.width, height / PREVIEW_BASE.height);
  previewRenderer.setScale(scale);
}

function getPreviewConstraints() {
  if (!elements.previewWrap || !elements.previewCard) {
    return null;
  }
  const wrapWidth = elements.previewWrap.clientWidth;
  if (!wrapWidth) {
    return null;
  }
  const maxHeight = wrapWidth * (9 / 16);
  const minHeight = Math.min(140, maxHeight);
  return { wrapWidth, maxHeight, minHeight };
}

function setPreviewCardHeight(height) {
  const constraints = getPreviewConstraints();
  if (!constraints) {
    return;
  }
  const clampedHeight = Math.max(constraints.minHeight, Math.min(constraints.maxHeight, height));
  const width = Math.min(constraints.wrapWidth, clampedHeight * (16 / 9));
  const finalHeight = width * (9 / 16);
  elements.previewCard.style.width = `${width}px`;
  elements.previewCard.style.height = `${finalHeight}px`;
  if (previewSplitState) {
    previewSplitState.height = finalHeight;
    previewSplitState.maxHeight = constraints.maxHeight;
    previewSplitState.minHeight = constraints.minHeight;
  }
}

function initPreviewSplitter() {
  if (!elements.previewCard || !elements.previewSplitter) {
    return;
  }
  if (previewSplitState) {
    return;
  }
  const constraints = getPreviewConstraints();
  if (!constraints) {
    return;
  }
  previewSplitState = {
    maxHeight: constraints.maxHeight,
    minHeight: constraints.minHeight,
    height: constraints.maxHeight
  };
  setPreviewCardHeight(previewSplitState.height);

  elements.previewSplitter.addEventListener('mousedown', (event) => {
    event.preventDefault();
    if (!previewSplitState) {
      return;
    }
    const startY = event.clientY;
    const startHeight = elements.previewCard.getBoundingClientRect().height;
    elements.previewSplitter.classList.add('dragging');
    document.body.style.cursor = 'row-resize';

    const handleMove = (moveEvent) => {
      const delta = moveEvent.clientY - startY;
      const next = Math.max(
        previewSplitState.minHeight,
        Math.min(previewSplitState.maxHeight, startHeight + delta)
      );
      setPreviewCardHeight(next);
      updatePreviewScale();
      updatePreview();
    };

    const handleUp = () => {
      elements.previewSplitter.classList.remove('dragging');
      document.body.style.cursor = '';
      window.removeEventListener('mousemove', handleMove);
      window.removeEventListener('mouseup', handleUp);
    };

    window.addEventListener('mousemove', handleMove);
    window.addEventListener('mouseup', handleUp);
  });
}

function collectBackgroundPaths() {
  const paths = new Set();
  Object.values(state.project.songs || {}).forEach((song) => {
    const background = song && song.background ? song.background.path : '';
    if (background) {
      paths.add(background);
    }
  });
  const announcementSlides = (state.project.announcements && state.project.announcements.slides) || [];
  announcementSlides.forEach((slide) => {
    if (slide && slide.mediaPath) {
      paths.add(slide.mediaPath);
    }
  });
  const timerSlides = (state.project.timer && state.project.timer.slides) || [];
  timerSlides.forEach((slide) => {
    if (slide && slide.mediaPath) {
      paths.add(slide.mediaPath);
    }
  });
  return Array.from(paths);
}

async function checkAssets() {
  if (!elements.assetWarning) {
    return;
  }
  if (!window.api || !window.api.checkAssets) {
    elements.assetWarning.hidden = true;
    return;
  }
  const paths = collectBackgroundPaths();
  if (paths.length === 0) {
    elements.assetWarning.hidden = true;
    elements.assetWarning.removeAttribute('title');
    return;
  }
  const requestId = ++assetCheckToken;
  let result = null;
  try {
    result = await window.api.checkAssets({ paths, projectFolder: state.projectFolder });
  } catch (error) {
    console.error('Failed to check media assets', error);
  }
  if (requestId !== assetCheckToken) {
    return;
  }
  const missing = (result && result.missing) || [];
  if (!missing.length) {
    elements.assetWarning.hidden = true;
    elements.assetWarning.removeAttribute('title');
    return;
  }
  elements.assetWarning.hidden = false;
  elements.assetWarning.textContent = `Missing media (${missing.length})`;
  elements.assetWarning.title = missing.join('\n');
}

function scheduleAssetCheck() {
  if (!window.api || !window.api.checkAssets) {
    return;
  }
  if (assetCheckTimer) {
    window.clearTimeout(assetCheckTimer);
  }
  assetCheckTimer = window.setTimeout(() => {
    checkAssets();
    assetCheckTimer = null;
  }, 200);
}

function buildDefaultTextStyle(theme, key) {
  const basePx = theme.baseFontPx || 70;
  let defaultFontPx = basePx;
  if (key === 'title') {
    defaultFontPx = 70;
  } else if (key === 'footer') {
    defaultFontPx = 25;
  }

  const shadow = theme.shadow || { dx: 2, dy: 2, blur: 6, color: '#000000' };
  return {
    fontFamily: theme.fontFamily || 'Segoe UI',
    fontPx: defaultFontPx,
    color: theme.textColor || '#FFFFFF',
    strokeWidthPx: theme.strokeWidthPx ?? 1,
    strokeColor: theme.strokeColor || '#000000',
    shadow: { ...shadow }
  };
}

function resolveTextStyle(theme, key) {
  const base = buildDefaultTextStyle(theme, key);
  const style = (theme.textStyles && theme.textStyles[key]) || {};
  const shadow = style.shadow || {};
  return {
    fontFamily: style.fontFamily || base.fontFamily,
    fontPx: style.fontPx || base.fontPx,
    color: style.color || base.color,
    strokeWidthPx: style.strokeWidthPx ?? base.strokeWidthPx,
    strokeColor: style.strokeColor || base.strokeColor,
    shadow: {
      dx: shadow.dx ?? base.shadow.dx,
      dy: shadow.dy ?? base.shadow.dy,
      blur: shadow.blur ?? base.shadow.blur,
      color: shadow.color || base.shadow.color
    }
  };
}

function normalizeProject(project) {
  const base = createNewProject();
  const normalized = project || base;

  normalized.settings = { ...base.settings, ...(normalized.settings || {}) };
  normalized.announcements = { ...base.announcements, ...(normalized.announcements || {}) };
  normalized.timer = { ...base.timer, ...(normalized.timer || {}) };
  normalized.setlist = Array.isArray(normalized.setlist) ? normalized.setlist : [];
  normalized.songs = normalized.songs || {};

  const normalizeMediaSlide = (slide, index) => {
    const mediaPath = slide.mediaPath || slide.path || (slide.background && slide.background.path) || '';
    const mediaType = slide.mediaType || slide.type || (slide.background && slide.background.type) || 'image';
    return {
      id: slide.id || `media-${crypto.randomUUID()}`,
      label: slide.label || `Slide ${index + 1}`,
      mediaPath,
      mediaType: mediaType === 'video' ? 'video' : 'image'
    };
  };

  normalized.announcements.slides = Array.isArray(normalized.announcements.slides)
    ? normalized.announcements.slides.map(normalizeMediaSlide)
    : [];
  normalized.timer.slides = Array.isArray(normalized.timer.slides)
    ? normalized.timer.slides.map(normalizeMediaSlide)
    : [];
  delete normalized.announcements.enabled;
  delete normalized.timer.enabled;

  Object.entries(normalized.songs).forEach(([songId, song]) => {
    const defaultTheme = getDefaultTheme();
    song.id = song.id || songId;
    song.title = song.title || 'Untitled Song';
    song.ccli = {
      songNumber: '',
      authors: [],
      publisher: '',
      copyright: '',
      ...(song.ccli || {})
    };
    song.background = {
      type: 'image',
      path: '',
      ...(song.background || {})
    };
    const mergedTheme = {
      ...defaultTheme,
      ...(song.theme || {}),
      shadow: {
        ...defaultTheme.shadow,
        ...((song.theme || {}).shadow || {})
      }
    };
    mergedTheme.textStyles = {
      title: resolveTextStyle(mergedTheme, 'title'),
      lyrics: resolveTextStyle(mergedTheme, 'lyrics'),
      footer: resolveTextStyle(mergedTheme, 'footer')
    };
    song.theme = mergedTheme;
    song.slides = Array.isArray(song.slides) ? song.slides : [];
    song.slides = song.slides.map((slide, index) => {
      const titleText = slide.titleText || richTextFromPlain('');
      const lyricsText = slide.lyricsText || richTextFromPlain('');
      const footerText = slide.footerText || richTextFromPlain('');
      return {
        id: slide.id || `slide-${crypto.randomUUID()}`,
        label: slide.label || `Slide ${index + 1}`,
        template: slide.template || 'TitleLyricsFooter',
        showTitle: slide.showTitle ?? Boolean(plainFromRichText(titleText)),
        showLyrics: slide.showLyrics ?? Boolean(plainFromRichText(lyricsText)),
        showFooter: slide.showFooter ?? Boolean(plainFromRichText(footerText)),
        footerAutoCcli: slide.footerAutoCcli === true,
        titleText,
        lyricsText,
        footerText
      };
    });
  });

  return normalized;
}

function setProject(project, filePath = null, folderPath = null) {
  const normalized = normalizeProject(project);
  state.project = normalized;
  state.projectFolder = folderPath;
  state.projectFile = filePath;
  state.autoSaveEnabled = false;
  state.runtimeBackgrounds = {};
  elements.projectPath.textContent = filePath || folderPath || 'Unsaved session';
  state.selectedAnnouncementIndex = -1;
  state.selectedTimerIndex = -1;
  state.selectedSection = 'setlist';
  state.selectedSongId = null;
  state.selectedSlideIndex = 0;
  state.preview = { section: 'setlist', songId: null, slideIndex: 0 };

  state.live = { section: null, songId: null, slideIndex: 0 };

  refreshAll();
}

function getSelectedSong() {
  return state.selectedSongId ? state.project.songs[state.selectedSongId] : null;
}

function getSelectedSlide() {
  if (state.selectedSection !== 'setlist') {
    return null;
  }
  const song = getSelectedSong();
  if (!song) {
    return null;
  }
  return song.slides[state.selectedSlideIndex] || null;
}

function getSelectedAnnouncementSlide() {
  const slides = (state.project.announcements && state.project.announcements.slides) || [];
  return slides[state.selectedAnnouncementIndex] || null;
}

function getSelectedTimerSlide() {
  const slides = (state.project.timer && state.project.timer.slides) || [];
  return slides[state.selectedTimerIndex] || null;
}

function selectSection(section) {
  if (section === 'announcements') {
    state.selectedSection = 'announcements';
    const slides = state.project.announcements.slides || [];
    if (state.selectedAnnouncementIndex < 0 || state.selectedAnnouncementIndex >= slides.length) {
      state.selectedAnnouncementIndex = -1;
    }
    state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
    return;
  }
  if (section === 'timer') {
    state.selectedSection = 'timer';
    const slides = state.project.timer.slides || [];
    if (state.selectedTimerIndex < 0 || state.selectedTimerIndex >= slides.length) {
      state.selectedTimerIndex = -1;
    }
    state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
    return;
  }
  state.selectedSection = 'setlist';
  if (!state.selectedSongId || !state.project.songs[state.selectedSongId]) {
    state.selectedSongId = state.project.setlist[0] || null;
    state.selectedSlideIndex = 0;
  }
  if (state.selectedSongId) {
    const song = state.project.songs[state.selectedSongId];
    const maxIndex = song && song.slides ? song.slides.length - 1 : 0;
    state.selectedSlideIndex = Math.min(state.selectedSlideIndex, Math.max(maxIndex, 0));
    state.preview = { section: 'setlist', songId: state.selectedSongId, slideIndex: state.selectedSlideIndex };
  } else {
    state.preview = { section: 'setlist', songId: null, slideIndex: 0 };
  }
}


function getPreviewSelection() {
  const section = state.preview.section || 'setlist';
  if (section === 'announcements') {
    const slides = (state.project.announcements && state.project.announcements.slides) || [];
    const slide = slides[state.preview.slideIndex];
    if (!slide) {
      return null;
    }
    return { section, song: null, slide, slideIndex: state.preview.slideIndex };
  }
  if (section === 'timer') {
    const slides = (state.project.timer && state.project.timer.slides) || [];
    const slide = slides[state.preview.slideIndex];
    if (!slide) {
      return null;
    }
    return { section, song: null, slide, slideIndex: state.preview.slideIndex };
  }
  if (!state.preview.songId) {
    return null;
  }
  const song = state.project.songs[state.preview.songId];
  if (!song) {
    return null;
  }
  const slide = song.slides[state.preview.slideIndex];
  if (!slide) {
    return null;
  }
  return { section: 'setlist', song, slide, slideIndex: state.preview.slideIndex };
}

function getLiveSelection() {
  const section = state.live.section;
  if (section === 'announcements') {
    const slides = (state.project.announcements && state.project.announcements.slides) || [];
    const slide = slides[state.live.slideIndex];
    if (!slide) {
      return null;
    }
    return { section, song: null, slide, slideIndex: state.live.slideIndex };
  }
  if (section === 'timer') {
    const slides = (state.project.timer && state.project.timer.slides) || [];
    const slide = slides[state.live.slideIndex];
    if (!slide) {
      return null;
    }
    return { section, song: null, slide, slideIndex: state.live.slideIndex };
  }
  if (!state.live.songId) {
    return null;
  }
  const song = state.project.songs[state.live.songId];
  if (!song) {
    return null;
  }
  const slide = song.slides[state.live.slideIndex];
  if (!slide) {
    return null;
  }
  return { section: 'setlist', song, slide, slideIndex: state.live.slideIndex };
}

function buildTextFromSlide(slide) {
  return {
    title: plainFromRichText(slide.titleText),
    lyrics: plainFromRichText(slide.lyricsText),
    footer: plainFromRichText(slide.footerText),
    showTitle: slide.showTitle,
    showLyrics: slide.showLyrics,
    showFooter: slide.showFooter
  };
}

function buildCcliFooter(song) {
  if (!song || !song.ccli) {
    return '';
  }
  const parts = [];
  if (song.ccli.songNumber) {
    parts.push(`CCLI #${song.ccli.songNumber}`);
  }
  if (song.ccli.authors && song.ccli.authors.length > 0) {
    parts.push(song.ccli.authors.join(', '));
  }
  if (song.ccli.publisher) {
    parts.push(song.ccli.publisher);
  }
  if (song.ccli.copyright) {
    parts.push(song.ccli.copyright);
  }
  return parts.join(' - ');
}

function buildTextFromSlideWithSong(song, slide) {
  const footerText = slide.footerAutoCcli ? buildCcliFooter(song) : plainFromRichText(slide.footerText);
  return {
    title: plainFromRichText(slide.titleText),
    lyrics: plainFromRichText(slide.lyricsText),
    footer: footerText,
    showTitle: slide.showTitle,
    showLyrics: slide.showLyrics,
    showFooter: slide.showFooter
  };
}

function buildBackground(song) {
  if (!song || !song.background || !song.background.path) {
    return null;
  }
  const rawPath = song.background.path;
  if (rawPath.startsWith('file://')) {
    return { type: song.background.type, path: rawPath };
  }
  if (rawPath.startsWith('library/') || rawPath.startsWith('library\\')) {
    if (window.api && window.api.resolveLibraryUrl) {
      return { type: song.background.type, path: window.api.resolveLibraryUrl(rawPath) };
    }
  }
  if (/^[a-zA-Z]:[\\/]/.test(rawPath)) {
    const normalized = rawPath.replace(/\\/g, '/');
    return { type: song.background.type, path: `file:///${encodeURI(normalized)}` };
  }
  const runtimePath = state.runtimeBackgrounds[song.id];
  if (runtimePath) {
    return { type: song.background.type, path: runtimePath };
  }
  if (!window.api || !window.api.resolveMediaUrl) {
    return null;
  }
  if (!state.projectFolder) {
    return null;
  }
  return {
    type: song.background.type,
    path: window.api.resolveMediaUrl(state.projectFolder, song.background.path)
  };
}

function resolveMediaPath(rawPath) {
  if (!rawPath) {
    return '';
  }
  if (rawPath.startsWith('file://')) {
    return rawPath;
  }
  if (rawPath.startsWith('library/') || rawPath.startsWith('library\\')) {
    if (window.api && window.api.resolveLibraryUrl) {
      return window.api.resolveLibraryUrl(rawPath);
    }
  }
  if (/^[a-zA-Z]:[\\/]/.test(rawPath)) {
    const normalized = rawPath.replace(/\\/g, '/');
    return `file:///${encodeURI(normalized)}`;
  }
  if (window.api && window.api.resolveMediaUrl && state.projectFolder) {
    return window.api.resolveMediaUrl(state.projectFolder, rawPath);
  }
  return '';
}

function buildRenderState(selection, panic, options = {}) {
  if (!selection) {
    return null;
  }
  const section = selection.section || 'setlist';
  if (section !== 'setlist') {
    const slide = selection.slide;
    const background = slide && slide.mediaPath
      ? {
        type: slide.mediaType || 'image',
        path: resolveMediaPath(slide.mediaPath),
        loop: section === 'timer' && slide.mediaType === 'video' ? false : true
      }
      : null;
    return {
      section,
      slideKey: `${section}:${slide ? slide.id : 'empty'}`,
      background,
      backgroundKey:
        slide && slide.mediaType === 'video'
          ? `${section}:${slide.id}:${slide.mediaPath || ''}`
          : undefined,
      theme: getDefaultTheme(),
      text: {
        title: '',
        lyrics: '',
        footer: '',
        showTitle: false,
        showLyrics: false,
        showFooter: false
      },
      panic: Boolean(panic),
      settings: state.project.settings,
      textImmediate: Boolean(options.textImmediate)
    };
  }
  return {
    section: 'setlist',
    slideKey: `${selection.song.id}:${selection.slide.id}`,
    background: buildBackground(selection.song),
    theme: selection.song.theme || getDefaultTheme(),
    text: buildTextFromSlideWithSong(selection.song, selection.slide),
    panic: Boolean(panic),
    settings: state.project.settings,
    textImmediate: Boolean(options.textImmediate)
  };
}

function refreshAll() {
  updateEditorVisibility();
  renderLineupButtons();
  renderAnnouncements();
  renderTimerSlides();
  renderSetlist();
  renderSongDetails();
  renderSlides();
  renderSlideEditor();
  renderInspector();
  initPreviewSplitter();
  updatePreviewScale();
  updatePreview();
  updateLiveStatus();
  scheduleAssetCheck();
}

function renderLineupButtons() {
  if (elements.announcementsButton) {
    elements.announcementsButton.classList.toggle('active', state.selectedSection === 'announcements');
  }
  if (elements.timerButton) {
    elements.timerButton.classList.toggle('active', state.selectedSection === 'timer');
  }
}

function updateEditorVisibility() {
  const section = state.selectedSection || 'setlist';
  const hasSong = Boolean(getSelectedSong());
  if (elements.songEditor) {
    elements.songEditor.hidden = section !== 'setlist' || !hasSong;
  }
  if (elements.announcementEditor) {
    elements.announcementEditor.hidden = section !== 'announcements';
  }
  if (elements.timerEditor) {
    elements.timerEditor.hidden = section !== 'timer';
  }
}

function renderAnnouncements() {
  const slides = (state.project.announcements && state.project.announcements.slides) || [];
  const targets = [
    { list: elements.announcements, empty: elements.announcementsEmpty, showEmpty: true },
    { list: elements.announcementsPreview, empty: elements.announcementsPreviewEmpty, showEmpty: false }
  ];

  targets.forEach(({ list, empty, showEmpty }) => {
    if (!list) {
      return;
    }
    list.innerHTML = '';
    if (empty) {
      empty.style.display = slides.length === 0 && showEmpty ? 'block' : 'none';
    }

    slides.forEach((slide, index) => {
      const item = document.createElement('li');
      item.className = 'list-item draggable';
      item.draggable = true;
      item.dataset.index = String(index);
      if (state.selectedSection === 'announcements' && state.selectedAnnouncementIndex === index) {
        item.classList.add('active');
      }

      const thumb = document.createElement('div');
      thumb.className = 'media-thumb';
      const mediaUrl = resolveMediaPath(slide.mediaPath);
      if (mediaUrl) {
        const img = document.createElement('img');
        img.src = mediaUrl;
        img.alt = slide.label || `Announcement ${index + 1}`;
        thumb.appendChild(img);
      } else {
        thumb.innerHTML = '<span class="media-label">No image</span>';
      }

      item.appendChild(thumb);

      item.addEventListener('click', () => selectAnnouncementSlide(index));
      item.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        event.stopPropagation();
        selectAnnouncementSlide(index);
        showMediaContextMenu(event.clientX, event.clientY, index, 'announcements');
      });
      item.addEventListener('dragstart', (event) => {
        internalDragActive = true;
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(index));
        event.dataTransfer.setData('application/x-wp-internal', 'announcements');
        item.classList.add('dragging');
        hideMediaContextMenu();
      });
      item.addEventListener('dragend', () => {
        internalDragActive = false;
        item.classList.remove('dragging');
        clearAnnouncementDragIndicators();
      });
      item.addEventListener('dragover', (event) => {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
        const rect = item.getBoundingClientRect();
        const position = event.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
        if (item.dataset.dropPosition !== position) {
          clearAnnouncementDragIndicators();
          item.dataset.dropPosition = position;
          item.classList.toggle('drag-over-before', position === 'before');
          item.classList.toggle('drag-over-after', position === 'after');
        }
      });
      item.addEventListener('dragleave', (event) => {
        if (!item.contains(event.relatedTarget)) {
          item.classList.remove('drag-over-before', 'drag-over-after');
          delete item.dataset.dropPosition;
        }
      });
      item.addEventListener('drop', (event) => {
        event.preventDefault();
        internalDragActive = false;
        const fromIndex = Number(event.dataTransfer.getData('text/plain'));
        const toIndex = Number(item.dataset.index);
        const position = item.dataset.dropPosition || 'before';
        if (Number.isNaN(fromIndex) || Number.isNaN(toIndex)) {
          return;
        }
        reorderAnnouncement(fromIndex, toIndex, position);
        clearAnnouncementDragIndicators();
      });

      list.appendChild(item);
    });
  });
}

function renderTimerSlides() {
  const slides = (state.project.timer && state.project.timer.slides) || [];
  const targets = [
    { list: elements.timerList, empty: elements.timerEmpty, showEmpty: true },
    { list: elements.timerPreviewList, empty: elements.timerPreviewEmpty, showEmpty: false }
  ];

  targets.forEach(({ list, empty, showEmpty }) => {
    if (!list) {
      return;
    }
    list.innerHTML = '';
    if (empty) {
      empty.style.display = slides.length === 0 && showEmpty ? 'block' : 'none';
    }

    slides.forEach((slide, index) => {
      const item = document.createElement('li');
      item.className = 'list-item draggable';
      item.draggable = true;
      item.dataset.index = String(index);
      if (state.selectedSection === 'timer' && state.selectedTimerIndex === index) {
        item.classList.add('active');
      }

      const thumb = document.createElement('div');
      thumb.className = 'media-thumb';
      const mediaUrl = resolveMediaPath(slide.mediaPath);
      if (mediaUrl) {
        if (slide.mediaType === 'video') {
          const video = document.createElement('video');
          video.src = mediaUrl;
          video.muted = true;
          video.playsInline = true;
          video.preload = 'metadata';
          thumb.appendChild(video);
        } else {
          const img = document.createElement('img');
          img.src = mediaUrl;
          img.alt = slide.label || `Timer ${index + 1}`;
          thumb.appendChild(img);
        }
      } else {
        thumb.innerHTML = '<span class="media-label">No media</span>';
      }

      item.appendChild(thumb);

      item.addEventListener('click', () => selectTimerSlide(index));
      item.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        event.stopPropagation();
        selectTimerSlide(index);
        showMediaContextMenu(event.clientX, event.clientY, index, 'timer');
      });
      item.addEventListener('dragstart', (event) => {
        internalDragActive = true;
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(index));
        event.dataTransfer.setData('application/x-wp-internal', 'timer');
        item.classList.add('dragging');
        hideMediaContextMenu();
      });
      item.addEventListener('dragend', () => {
        internalDragActive = false;
        item.classList.remove('dragging');
        clearTimerDragIndicators();
      });
      item.addEventListener('dragover', (event) => {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
        const rect = item.getBoundingClientRect();
        const position = event.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
        if (item.dataset.dropPosition !== position) {
          clearTimerDragIndicators();
          item.dataset.dropPosition = position;
          item.classList.toggle('drag-over-before', position === 'before');
          item.classList.toggle('drag-over-after', position === 'after');
        }
      });
      item.addEventListener('dragleave', (event) => {
        if (!item.contains(event.relatedTarget)) {
          item.classList.remove('drag-over-before', 'drag-over-after');
          delete item.dataset.dropPosition;
        }
      });
      item.addEventListener('drop', (event) => {
        event.preventDefault();
        internalDragActive = false;
        const fromIndex = Number(event.dataTransfer.getData('text/plain'));
        const toIndex = Number(item.dataset.index);
        const position = item.dataset.dropPosition || 'before';
        if (Number.isNaN(fromIndex) || Number.isNaN(toIndex)) {
          return;
        }
        reorderTimer(fromIndex, toIndex, position);
        clearTimerDragIndicators();
      });

      list.appendChild(item);
    });
  });
}

function renderSetlist() {
  elements.setlist.innerHTML = '';
  const songs = state.project.setlist;
  elements.setlistEmpty.style.display = songs.length === 0 ? 'block' : 'none';

  songs.forEach((songId, index) => {
    const song = state.project.songs[songId];
    if (!song) {
      return;
    }
    const item = document.createElement('li');
    item.className = 'list-item draggable';
    item.draggable = true;
    item.dataset.index = String(index);
    if (state.selectedSection === 'setlist' && songId === state.selectedSongId) {
      item.classList.add('active');
    }
    item.innerHTML = `
      <strong>${song.title || 'Untitled Song'}</strong>
      <div class="meta">
        <span>${song.slides.length} slides</span>
      </div>
    `;

    item.addEventListener('click', () => selectSong(songId));
    item.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectSong(songId);
      showSongContextMenu(event.clientX, event.clientY, index);
    });
    item.addEventListener('dragstart', (event) => {
      internalDragActive = true;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', String(index));
      event.dataTransfer.setData('application/x-wp-internal', 'setlist');
      item.classList.add('dragging');
      hideSongContextMenu();
    });
    item.addEventListener('dragend', () => {
      internalDragActive = false;
      item.classList.remove('dragging');
      clearSongDragIndicators();
    });
    item.addEventListener('dragover', (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      const rect = item.getBoundingClientRect();
      const position = event.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
      if (item.dataset.dropPosition !== position) {
        clearSongDragIndicators();
        item.dataset.dropPosition = position;
        item.classList.toggle('drag-over-before', position === 'before');
        item.classList.toggle('drag-over-after', position === 'after');
      }
    });
    item.addEventListener('dragleave', (event) => {
      if (!item.contains(event.relatedTarget)) {
        item.classList.remove('drag-over-before', 'drag-over-after');
        delete item.dataset.dropPosition;
      }
    });
    item.addEventListener('drop', (event) => {
      event.preventDefault();
      internalDragActive = false;
      const fromIndex = Number(event.dataTransfer.getData('text/plain'));
      const toIndex = Number(item.dataset.index);
      const position = item.dataset.dropPosition || 'before';
      if (Number.isNaN(fromIndex) || Number.isNaN(toIndex)) {
        return;
      }
      reorderSong(fromIndex, toIndex, position);
      clearSongDragIndicators();
    });

    elements.setlist.appendChild(item);
  });
}

function renderSongDetails() {
  const song = state.selectedSection === 'setlist' ? getSelectedSong() : null;
  elements.songTitle.value = song ? song.title : '';
  elements.songTitle.disabled = !song;
}

function renderSlides() {
  elements.slides.innerHTML = '';
  if (state.selectedSection !== 'setlist') {
    elements.slidesEmpty.style.display = 'block';
    return;
  }
  const song = getSelectedSong();
  if (!song) {
    elements.slidesEmpty.style.display = 'block';
    return;
  }
  elements.slidesEmpty.style.display = song.slides.length === 0 ? 'block' : 'none';

  song.slides.forEach((slide, index) => {
    const item = document.createElement('li');
    item.className = 'list-item draggable';
    item.draggable = true;
    item.dataset.index = String(index);
    if (index === state.selectedSlideIndex) {
      item.classList.add('active');
    }

    const rawText = plainFromRichText(slide.lyricsText) || plainFromRichText(slide.titleText) || 'Slide';
    const previewText = rawText.replace(/\s+/g, ' ').trim();
    const displayText = previewText.length > 30 ? `${previewText.slice(0, 30)}...` : previewText;
    const label = slide.label || `Slide ${index + 1}`;

    item.innerHTML = `
      <strong>${label}</strong>
      <div class="meta">
        <span>${displayText}</span>
      </div>
    `;

    item.addEventListener('click', () => selectSlide(index));
    item.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectSlide(index);
      showSlideContextMenu(event.clientX, event.clientY, index);
    });
    item.addEventListener('dragstart', (event) => {
      internalDragActive = true;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', String(index));
      event.dataTransfer.setData('application/x-wp-internal', 'slides');
      item.classList.add('dragging');
      hideSlideContextMenu();
    });
    item.addEventListener('dragend', () => {
      internalDragActive = false;
      item.classList.remove('dragging');
      clearSlideDragIndicators();
    });
    item.addEventListener('dragover', (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      const rect = item.getBoundingClientRect();
      const position = event.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
      if (item.dataset.dropPosition !== position) {
        clearSlideDragIndicators();
        item.dataset.dropPosition = position;
        item.classList.toggle('drag-over-before', position === 'before');
        item.classList.toggle('drag-over-after', position === 'after');
      }
    });
    item.addEventListener('dragleave', (event) => {
      if (!item.contains(event.relatedTarget)) {
        item.classList.remove('drag-over-before', 'drag-over-after');
        delete item.dataset.dropPosition;
      }
    });
    item.addEventListener('drop', (event) => {
      event.preventDefault();
      internalDragActive = false;
      const fromIndex = Number(event.dataTransfer.getData('text/plain'));
      const toIndex = Number(item.dataset.index);
      const position = item.dataset.dropPosition || 'before';
      if (Number.isNaN(fromIndex) || Number.isNaN(toIndex)) {
        return;
      }
      reorderSlide(fromIndex, toIndex, position);
      clearSlideDragIndicators();
    });

    elements.slides.appendChild(item);
  });
}

function renderSlideEditor() {
  if (state.selectedSection !== 'setlist') {
    [
      elements.showTitle,
      elements.showLyrics,
      elements.showFooter,
      elements.slideLabel,
      elements.titleText,
      elements.lyricsText,
      elements.footerText,
      elements.footerAuto
    ].forEach((input) => {
      input.disabled = true;
    });
    elements.slideLabel.value = '';
    elements.titleText.value = '';
    elements.lyricsText.value = '';
    elements.footerText.value = '';
    elements.footerAuto.checked = false;
    updateSlideVisibility();
    return;
  }
  const slide = getSelectedSlide();
  const disabled = !slide;

  elements.showTitle.checked = slide ? slide.showTitle : false;
  elements.showLyrics.checked = slide ? slide.showLyrics : false;
  elements.showFooter.checked = slide ? slide.showFooter : false;

  elements.slideLabel.value = slide ? slide.label || '' : '';
  elements.titleText.value = slide ? plainFromRichText(slide.titleText) : '';
  elements.lyricsText.value = slide ? plainFromRichText(slide.lyricsText) : '';
  elements.footerAuto.checked = slide ? slide.footerAutoCcli === true : false;
  const song = getSelectedSong();
  const footerValue = slide
    ? slide.footerAutoCcli
      ? buildCcliFooter(song)
      : plainFromRichText(slide.footerText)
    : '';
  elements.footerText.value = footerValue;
  elements.footerText.disabled = !slide || elements.footerAuto.checked;

  updateSlideVisibility();

  [
    elements.showTitle,
    elements.showLyrics,
    elements.showFooter,
    elements.slideLabel,
    elements.titleText,
    elements.lyricsText,
    elements.footerText,
    elements.footerAuto
  ].forEach((input) => {
    input.disabled = disabled;
  });
}

function renderInspector() {
  const section = state.selectedSection || 'setlist';
  const hasSong = section === 'setlist' ? Boolean(getSelectedSong()) : false;
  const hasAnnouncement = section === 'announcements' ? Boolean(getSelectedAnnouncementSlide()) : false;
  const hasTimer = section === 'timer' ? Boolean(getSelectedTimerSlide()) : false;
  const hasSelection = hasSong || hasAnnouncement || hasTimer;

  setInspectorVisible(elements.inspectorEmpty, !hasSelection);
  setInspectorVisible(elements.songStatus, hasSelection && section === 'setlist');

  setInspectorVisible(elements.backgroundInspector, false);
  setInspectorVisible(elements.themeGroup, false);
  setInspectorVisible(elements.ccliGroup, false);
  setInspectorVisible(elements.announcementInspector, false);
  setInspectorVisible(elements.timerInspector, false);

  if (!hasSelection) {
    return;
  }

  if (section === 'announcements') {
    const slide = getSelectedAnnouncementSlide();
    if (elements.songStatus) {
      elements.songStatus.textContent = 'Announcements';
    }
    setInspectorVisible(elements.announcementInspector, Boolean(slide));
    if (!slide) {
      return;
    }
    elements.announcementPath.textContent = slide.mediaPath ? slide.mediaPath : 'No announcement selected';
    elements.announcementAuto.checked = state.project.announcements.autoAdvanceEnabled !== false;
    elements.announcementAdvance.value = state.project.announcements.autoAdvanceSec ?? 15;
    elements.announcementLoop.checked = state.project.announcements.loop === true;
    elements.announcementAdvance.disabled = !elements.announcementAuto.checked;
    elements.openAnnouncementLibrary.disabled = false;
    elements.uploadAnnouncement.disabled = false;
    return;
  }

  if (section === 'timer') {
    const slide = getSelectedTimerSlide();
    if (elements.songStatus) {
      elements.songStatus.textContent = 'Timer';
    }
    setInspectorVisible(elements.timerInspector, Boolean(slide));
    if (!slide) {
      return;
    }
    elements.timerMediaPath.textContent = slide.mediaPath ? slide.mediaPath : 'No timer media selected';
    elements.timerAutoVideo.checked = state.project.timer.autoAdvanceOnVideoEnd !== false;
    elements.timerAutoImages.checked = state.project.timer.autoAdvanceImages === true;
    elements.timerAdvance.value = state.project.timer.autoAdvanceSec ?? 15;
    elements.timerAdvanceRow.style.display = elements.timerAutoImages.checked ? 'flex' : 'none';
    elements.openTimerLibrary.disabled = false;
    elements.uploadTimerMedia.disabled = false;
    return;
  }

  const song = getSelectedSong();
  if (!song) {
    return;
  }

  setInspectorVisible(elements.backgroundInspector, true);
  setInspectorVisible(elements.themeGroup, true);
  setInspectorVisible(elements.ccliGroup, true);
  if (elements.songStatus) {
    elements.songStatus.textContent = 'Editing song';
    elements.songStatus.hidden = false;
  }
  elements.backgroundPath.textContent = song.background.path || 'No background selected';

  elements.themeTarget.value = state.themeTarget || 'lyrics';
  state.themeTarget = elements.themeTarget.value || 'lyrics';
  updateThemeEditorFromSelection(song);
  setThemePositionValue(song.theme.position || 'center');
  elements.themeDim.value = song.theme.dimOpacity ?? 0.25;

  elements.ccliNumber.value = song.ccli.songNumber || '';
  elements.ccliAuthors.value = (song.ccli.authors || []).join(', ');
  elements.ccliPublisher.value = song.ccli.publisher || '';
  elements.ccliCopyright.value = song.ccli.copyright || '';

  updateThemeVisibility();
  toggleInspectorFields(false);
}

function toggleInspectorFields(disabled) {
  [
    elements.openLibrary,
    elements.uploadBackground,
    elements.themeTarget,
    elements.themeFont,
    elements.themeBase,
    elements.themeColor,
    elements.themeStrokeToggle,
    elements.themeStroke,
    elements.themeStrokeColor,
    elements.themeShadowToggle,
    elements.themeShadowX,
    elements.themeShadowY,
    elements.themeShadowBlur,
    elements.themeShadowColor,
    elements.themeDim,
    elements.ccliNumber,
    elements.ccliAuthors,
    elements.ccliPublisher,
    elements.ccliCopyright,
    elements.addSlide
  ].forEach((input) => {
    input.disabled = disabled;
  });

  if (elements.themePosition) {
    elements.themePosition.classList.toggle('disabled', disabled);
    elements.themePosition.querySelectorAll('button').forEach((button) => {
      button.disabled = disabled;
    });
  }
}

function setRowVisible(row, visible) {
  if (!row) {
    return;
  }
  row.style.display = visible ? 'flex' : 'none';
}

function setInspectorVisible(element, visible) {
  if (!element) {
    return;
  }
  element.hidden = !visible;
  element.style.display = visible ? '' : 'none';
}

function updateSlideVisibility() {
  const slide = getSelectedSlide();
  const hasSlide = Boolean(slide);
  const showTitle = hasSlide && elements.showTitle.checked;
  const showLyrics = hasSlide && elements.showLyrics.checked;
  const showFooter = hasSlide && elements.showFooter.checked;

  setRowVisible(elements.titleRow, showTitle);
  setRowVisible(elements.lyricsRow, showLyrics);
  setRowVisible(elements.footerRow, showFooter);
  if (elements.footerText) {
    elements.footerText.disabled = !showFooter || elements.footerAuto.checked;
  }
}

function updateThemeEditorFromSelection(song) {
  if (!song) {
    return;
  }
  const target = elements.themeTarget.value || 'lyrics';
  const style = resolveTextStyle(song.theme, target);
  const fontMatches = Array.from(elements.themeFont.options).some((option) => option.value === style.fontFamily);
  elements.themeFont.value = fontMatches ? style.fontFamily : 'Segoe UI';
  elements.themeBase.value = style.fontPx;
  elements.themeColor.value = style.color || '#ffffff';

  elements.themeStrokeToggle.checked = (style.strokeWidthPx || 0) > 0;
  elements.themeStroke.value = style.strokeWidthPx ?? 1;
  elements.themeStrokeColor.value = style.strokeColor || '#000000';

  const shadowBlur = style.shadow?.blur ?? 0;
  elements.themeShadowToggle.checked = shadowBlur >= 0;
  elements.themeShadowX.value = style.shadow?.dx ?? 0;
  elements.themeShadowY.value = style.shadow?.dy ?? 0;
  elements.themeShadowBlur.value = shadowBlur >= 0 ? shadowBlur : 0;
  elements.themeShadowColor.value = style.shadow?.color || '#000000';
  updateThemeVisibility();
}

function updateThemeVisibility() {
  if (elements.strokeOptions) {
    elements.strokeOptions.hidden = !elements.themeStrokeToggle.checked;
  }
  if (elements.shadowOptions) {
    elements.shadowOptions.hidden = !elements.themeShadowToggle.checked;
  }
}

function setThemePositionValue(value) {
  if (!elements.themePosition) {
    return;
  }
  const buttons = elements.themePosition.querySelectorAll('button[data-value]');
  let matched = false;
  buttons.forEach((button) => {
    const isActive = button.dataset.value === value;
    if (isActive) {
      matched = true;
    }
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
  });
  if (!matched && buttons.length > 0) {
    const fallback = buttons[0];
    fallback.classList.add('active');
    fallback.setAttribute('aria-pressed', 'true');
  }
}

function getThemePositionValue() {
  if (!elements.themePosition) {
    return 'center';
  }
  const active = elements.themePosition.querySelector('button[aria-pressed="true"]');
  if (active && active.dataset.value) {
    return active.dataset.value;
  }
  const fallback = elements.themePosition.querySelector('button[data-value]');
  return fallback ? fallback.dataset.value : 'center';
}

function setThemeSectionCollapsed(collapsed) {
  if (!elements.themeSection || !elements.themeToggle) {
    return;
  }
  elements.themeSection.hidden = collapsed;
  elements.themeToggle.classList.toggle('collapsed', collapsed);
  elements.themeToggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  const chevron = elements.themeToggle.querySelector('.chevron');
  if (chevron) {
    chevron.textContent = collapsed ? '>' : 'v';
  }
}

function setCcliSectionCollapsed(collapsed) {
  if (!elements.ccliSection || !elements.ccliToggle) {
    return;
  }
  elements.ccliSection.hidden = collapsed;
  elements.ccliToggle.classList.toggle('collapsed', collapsed);
  elements.ccliToggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  const chevron = elements.ccliToggle.querySelector('.chevron');
  if (chevron) {
    chevron.textContent = collapsed ? '>' : 'v';
  }
}

function updatePreview() {
  const selection = getPreviewSelection();
  if (!selection) {
    previewRenderer.clear();
    return;
  }
  const renderState = buildRenderState(selection, false, { textImmediate: true });
  if (renderState) {
    previewRenderer.render(renderState);
  }
}

function sendProgramState() {
  const selection = getLiveSelection();
  const renderState = buildRenderState(selection, state.panic);
  if (!renderState) {
    return;
  }
  window.api.sendProgramState(renderState);
}

function updateLiveStatus() {
  const selection = getLiveSelection();
  if (!selection) {
    elements.liveStatus.textContent = 'Live: -';
    return;
  }
  if (selection.section === 'announcements') {
    elements.liveStatus.textContent = `Live: Announcements (Slide ${selection.slideIndex + 1})`;
  } else if (selection.section === 'timer') {
    elements.liveStatus.textContent = `Live: Timer (Slide ${selection.slideIndex + 1})`;
  } else {
    elements.liveStatus.textContent = `Live: ${selection.song.title} (Slide ${selection.slideIndex + 1})`;
  }
}

function clearAnnouncementAdvanceTimer() {
  if (announcementAdvanceTimer) {
    window.clearTimeout(announcementAdvanceTimer);
    announcementAdvanceTimer = null;
  }
}

function clearTimerImageAdvanceTimer() {
  if (timerImageAdvanceTimer) {
    window.clearTimeout(timerImageAdvanceTimer);
    timerImageAdvanceTimer = null;
  }
}

function updateAutoAdvanceTimers() {
  clearAnnouncementAdvanceTimer();
  clearTimerImageAdvanceTimer();

  if (state.live.section === 'announcements') {
    const slides = state.project.announcements.slides || [];
    if (slides.length === 0) {
      return;
    }
    if (state.project.announcements.autoAdvanceEnabled === false) {
      return;
    }
    const intervalSec = Number(state.project.announcements.autoAdvanceSec) || 0;
    if (intervalSec <= 0) {
      return;
    }
    const isLast = state.live.slideIndex >= slides.length - 1;
    if (isLast && !state.project.announcements.loop) {
      return;
    }
    announcementAdvanceTimer = window.setTimeout(() => {
      advanceLive({ source: 'announcements-timer' });
    }, intervalSec * 1000);
    return;
  }

  if (state.live.section === 'timer') {
    const slides = state.project.timer.slides || [];
    if (slides.length === 0) {
      return;
    }
    if (!state.project.timer.autoAdvanceImages) {
      return;
    }
    const current = slides[state.live.slideIndex];
    if (!current || current.mediaType === 'video') {
      return;
    }
    const intervalSec = Number(state.project.timer.autoAdvanceSec) || 0;
    if (intervalSec <= 0) {
      return;
    }
    const isLast = state.live.slideIndex >= slides.length - 1;
    if (isLast) {
      return;
    }
    timerImageAdvanceTimer = window.setTimeout(() => {
      advanceLive({ source: 'timer-image' });
    }, intervalSec * 1000);
  }
}

function selectSong(songId) {
  state.selectedSection = 'setlist';
  state.selectedSongId = songId;
  state.selectedSlideIndex = 0;
  state.preview = { section: 'setlist', songId, slideIndex: 0 };
  refreshAll();
}

function selectSlide(index) {
  state.selectedSection = 'setlist';
  state.selectedSlideIndex = index;
  const song = getSelectedSong();
  if (song) {
    state.preview = { section: 'setlist', songId: song.id, slideIndex: index };
  }
  renderSlides();
  renderSlideEditor();
  updatePreview();
}

function selectAnnouncementSlide(index) {
  state.selectedSection = 'announcements';
  state.selectedAnnouncementIndex = index;
  state.preview = { section: 'announcements', songId: null, slideIndex: index };
  updateEditorVisibility();
  renderLineupButtons();
  renderAnnouncements();
  renderInspector();
  updatePreview();
}

function selectTimerSlide(index) {
  state.selectedSection = 'timer';
  state.selectedTimerIndex = index;
  state.preview = { section: 'timer', songId: null, slideIndex: index };
  updateEditorVisibility();
  renderLineupButtons();
  renderTimerSlides();
  renderInspector();
  updatePreview();
}

function addSong() {
  const song = createSong('New Song');
  song.theme = getDefaultTheme();
  const makeLyricsSlide = (label) => {
    const slide = createSlide({ label, lyricsText: '', footerAutoCcli: false });
    slide.showTitle = false;
    slide.showLyrics = true;
    slide.showFooter = false;
    return slide;
  };

  const firstSlide = createSlide({
    label: 'Verse 1-1',
    titleText: song.title,
    lyricsText: '',
    footerAutoCcli: true
  });
  firstSlide.showTitle = true;
  firstSlide.showLyrics = true;
  firstSlide.showFooter = true;

  const slides = [
    firstSlide,
    makeLyricsSlide('Verse 1-2'),
    makeLyricsSlide('Chorus'),
    makeLyricsSlide('Verse 2-1'),
    makeLyricsSlide('Verse 2-2'),
    makeLyricsSlide('Chorus'),
    makeLyricsSlide('Bridge'),
    makeLyricsSlide('Chorus'),
    makeLyricsSlide('Outro')
  ];
  song.slides.push(...slides);

  state.project.songs[song.id] = song;
  state.project.setlist.push(song.id);
  selectSong(song.id);
  saveProjectIfPossible();
}

function addAnnouncementSlide() {
  const slides = state.project.announcements.slides;
  const label = `Announcement ${slides.length + 1}`;
  slides.push(createMediaSlide({ label, mediaType: 'image' }));
  state.selectedSection = 'announcements';
  state.selectedAnnouncementIndex = slides.length - 1;
  state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
  refreshAll();
  saveProjectIfPossible();
}

function addTimerSlide() {
  const slides = state.project.timer.slides;
  const label = `Timer ${slides.length + 1}`;
  slides.push(createMediaSlide({ label, mediaType: 'image' }));
  state.selectedSection = 'timer';
  state.selectedTimerIndex = slides.length - 1;
  state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
  refreshAll();
  saveProjectIfPossible();
}

function deleteSong(songId) {
  const index = state.project.setlist.indexOf(songId);
  if (index === -1) {
    return;
  }
  state.project.setlist.splice(index, 1);
  delete state.project.songs[songId];

  const remaining = state.project.setlist;
  if (state.selectedSongId === songId) {
    state.selectedSongId = remaining[0] || null;
    state.selectedSlideIndex = 0;
  }
  if (remaining.length === 0) {
    state.selectedSongId = null;
    state.selectedSlideIndex = 0;
    if (state.selectedSection === 'setlist') {
      state.preview = { section: 'setlist', songId: null, slideIndex: 0 };
    }
    if (state.live.section === 'setlist') {
      state.live = { section: null, songId: null, slideIndex: 0 };
    }
  } else {
    if (!state.selectedSongId) {
      state.selectedSongId = remaining[0];
    }
    if (state.preview.section === 'setlist' && state.preview.songId === songId) {
      state.preview = { section: 'setlist', songId: state.selectedSongId, slideIndex: 0 };
    }
    if (state.live.section === 'setlist' && state.live.songId === songId) {
      state.live = { section: 'setlist', songId: state.selectedSongId, slideIndex: 0 };
    }
  }
  refreshAll();
  if (getLiveSelection()) {
    sendProgramState();
  }
  saveProjectIfPossible();
}

function moveSong(index, delta) {
  const newIndex = index + delta;
  if (newIndex < 0 || newIndex >= state.project.setlist.length) {
    return;
  }
  const list = state.project.setlist;
  const [moved] = list.splice(index, 1);
  list.splice(newIndex, 0, moved);
  refreshAll();
  saveProjectIfPossible();
}

function reorderSong(fromIndex, toIndex, position = 'before') {
  const list = state.project.setlist;
  if (fromIndex < 0 || fromIndex >= list.length) {
    return;
  }
  if (toIndex < 0 || toIndex >= list.length) {
    return;
  }
  let insertIndex = toIndex;
  if (position === 'after') {
    insertIndex += 1;
  }
  if (insertIndex > fromIndex) {
    insertIndex -= 1;
  }
  if (insertIndex === fromIndex) {
    hideSongContextMenu();
    clearSongDragIndicators();
    return;
  }
  const [moved] = list.splice(fromIndex, 1);
  const maxIndex = list.length;
  if (insertIndex < 0) {
    insertIndex = 0;
  }
  if (insertIndex > maxIndex) {
    insertIndex = maxIndex;
  }
  list.splice(insertIndex, 0, moved);
  hideSongContextMenu();
  clearSongDragIndicators();
  refreshAll();
  saveProjectIfPossible();
}

function copySong(songId) {
  const song = state.project.songs[songId];
  if (!song) {
    return;
  }
  const copy = JSON.parse(JSON.stringify(song));
  const newId = `song-${crypto.randomUUID()}`;
  copy.id = newId;
  copy.title = song.title ? `${song.title} Copy` : 'Untitled Song';
  copy.slides = (copy.slides || []).map((slide) => ({
    ...slide,
    id: `slide-${crypto.randomUUID()}`
  }));
  state.project.songs[newId] = copy;

  const index = state.project.setlist.indexOf(songId);
  const insertIndex = index === -1 ? state.project.setlist.length : index + 1;
  state.project.setlist.splice(insertIndex, 0, newId);

  if (state.runtimeBackgrounds[songId]) {
    state.runtimeBackgrounds[newId] = state.runtimeBackgrounds[songId];
  }

  selectSong(newId);
  refreshAll();
  saveProjectIfPossible();
}

function showSongContextMenu(x, y, index) {
  if (!elements.songContext) {
    return;
  }
  elements.songContext.dataset.index = String(index);
  elements.songContext.hidden = false;
  const menu = elements.songContext;
  const rect = menu.getBoundingClientRect();
  const maxX = window.innerWidth - rect.width - 10;
  const maxY = window.innerHeight - rect.height - 10;
  const left = Math.max(10, Math.min(x, maxX));
  const top = Math.max(10, Math.min(y, maxY));
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function hideSongContextMenu() {
  if (!elements.songContext) {
    return;
  }
  elements.songContext.hidden = true;
}

function clearSongDragIndicators() {
  if (!elements.setlist) {
    return;
  }
  elements.setlist.querySelectorAll('.list-item').forEach((item) => {
    item.classList.remove('drag-over-before', 'drag-over-after');
    delete item.dataset.dropPosition;
  });
}

function addSlide() {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  const insertIndex = Math.min(state.selectedSlideIndex + 1, song.slides.length);
  const label = `Slide ${song.slides.length + 1}`;
  song.slides.splice(insertIndex, 0, createSlide({ label, lyricsText: '' }));
  state.selectedSlideIndex = insertIndex;
  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex };
  refreshAll();
  saveProjectIfPossible();
}

function deleteSlide(index) {
  const song = getSelectedSong();
  if (!song || song.slides.length <= 1) {
    return;
  }
  song.slides.splice(index, 1);
  state.selectedSlideIndex = Math.max(0, Math.min(state.selectedSlideIndex, song.slides.length - 1));
  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex };
  hideSlideContextMenu();
  refreshAll();
  saveProjectIfPossible();
}

function deleteAnnouncementSlide(index) {
  const slides = state.project.announcements.slides;
  if (!slides || slides.length === 0) {
    return;
  }
  if (index < 0 || index >= slides.length) {
    return;
  }
  slides.splice(index, 1);
  state.selectedAnnouncementIndex = Math.max(0, Math.min(state.selectedAnnouncementIndex, slides.length - 1));
  state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
  if (state.live.section === 'announcements') {
    if (slides.length === 0) {
      state.live = { section: null, songId: null, slideIndex: 0 };
    } else {
      state.live.slideIndex = Math.max(0, Math.min(state.live.slideIndex, slides.length - 1));
    }
  }
  hideMediaContextMenu();
  refreshAll();
  if (getLiveSelection()) {
    sendProgramState();
  }
  saveProjectIfPossible();
}

function deleteTimerSlide(index) {
  const slides = state.project.timer.slides;
  if (!slides || slides.length === 0) {
    return;
  }
  if (index < 0 || index >= slides.length) {
    return;
  }
  slides.splice(index, 1);
  state.selectedTimerIndex = Math.max(0, Math.min(state.selectedTimerIndex, slides.length - 1));
  state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
  if (state.live.section === 'timer') {
    if (slides.length === 0) {
      state.live = { section: null, songId: null, slideIndex: 0 };
    } else {
      state.live.slideIndex = Math.max(0, Math.min(state.live.slideIndex, slides.length - 1));
    }
  }
  hideMediaContextMenu();
  refreshAll();
  if (getLiveSelection()) {
    sendProgramState();
  }
  saveProjectIfPossible();
}

function moveSlide(index, delta) {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  const newIndex = index + delta;
  if (newIndex < 0 || newIndex >= song.slides.length) {
    return;
  }
  const [moved] = song.slides.splice(index, 1);
  song.slides.splice(newIndex, 0, moved);
  state.selectedSlideIndex = newIndex;
  state.preview = { section: 'setlist', songId: song.id, slideIndex: newIndex };
  refreshAll();
  saveProjectIfPossible();
}

function reorderSlide(fromIndex, toIndex, position = 'before') {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  if (fromIndex < 0 || fromIndex >= song.slides.length) {
    return;
  }
  if (toIndex < 0 || toIndex >= song.slides.length) {
    return;
  }
  let insertIndex = toIndex;
  if (position === 'after') {
    insertIndex += 1;
  }
  if (insertIndex > fromIndex) {
    insertIndex -= 1;
  }
  if (insertIndex === fromIndex) {
    hideSlideContextMenu();
    clearSlideDragIndicators();
    return;
  }
  const [moved] = song.slides.splice(fromIndex, 1);
  const maxIndex = song.slides.length;
  if (insertIndex < 0) {
    insertIndex = 0;
  }
  if (insertIndex > maxIndex) {
    insertIndex = maxIndex;
  }
  song.slides.splice(insertIndex, 0, moved);

  if (state.selectedSlideIndex === fromIndex) {
    state.selectedSlideIndex = insertIndex;
  } else if (fromIndex < insertIndex && state.selectedSlideIndex > fromIndex && state.selectedSlideIndex <= insertIndex) {
    state.selectedSlideIndex -= 1;
  } else if (fromIndex > insertIndex && state.selectedSlideIndex >= insertIndex && state.selectedSlideIndex < fromIndex) {
    state.selectedSlideIndex += 1;
  }

  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex };
  hideSlideContextMenu();
  clearSlideDragIndicators();
  refreshAll();
  saveProjectIfPossible();
}

function copySlide(index) {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  const source = song.slides[index];
  if (!source) {
    return;
  }
  const clone = JSON.parse(JSON.stringify(source));
  clone.id = `slide-${crypto.randomUUID()}`;
  song.slides.splice(index + 1, 0, clone);
  state.selectedSlideIndex = index + 1;
  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex };
  hideSlideContextMenu();
  refreshAll();
  saveProjectIfPossible();
}

function showSlideContextMenu(x, y, index) {
  if (!elements.slideContext) {
    return;
  }
  elements.slideContext.dataset.index = String(index);
  elements.slideContext.hidden = false;
  const menu = elements.slideContext;
  const rect = menu.getBoundingClientRect();
  const maxX = window.innerWidth - rect.width - 10;
  const maxY = window.innerHeight - rect.height - 10;
  const left = Math.max(10, Math.min(x, maxX));
  const top = Math.max(10, Math.min(y, maxY));
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function hideSlideContextMenu() {
  if (!elements.slideContext) {
    return;
  }
  elements.slideContext.hidden = true;
}

function showMediaContextMenu(x, y, index, section) {
  if (!elements.mediaContext) {
    return;
  }
  elements.mediaContext.dataset.index = String(index);
  elements.mediaContext.dataset.section = section;
  elements.mediaContext.hidden = false;
  const menu = elements.mediaContext;
  const rect = menu.getBoundingClientRect();
  const maxX = window.innerWidth - rect.width - 10;
  const maxY = window.innerHeight - rect.height - 10;
  const left = Math.max(10, Math.min(x, maxX));
  const top = Math.max(10, Math.min(y, maxY));
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function hideMediaContextMenu() {
  if (!elements.mediaContext) {
    return;
  }
  elements.mediaContext.hidden = true;
}

function clearSlideDragIndicators() {
  if (!elements.slides) {
    return;
  }
  elements.slides.querySelectorAll('.list-item').forEach((item) => {
    item.classList.remove('drag-over-before', 'drag-over-after');
    delete item.dataset.dropPosition;
  });
}

function clearAnnouncementDragIndicators() {
  const lists = [elements.announcements, elements.announcementsPreview].filter(Boolean);
  if (lists.length === 0) {
    return;
  }
  lists.forEach((list) => {
    list.querySelectorAll('.list-item').forEach((item) => {
      item.classList.remove('drag-over-before', 'drag-over-after');
      delete item.dataset.dropPosition;
    });
  });
}

function clearTimerDragIndicators() {
  const lists = [elements.timerList, elements.timerPreviewList].filter(Boolean);
  if (lists.length === 0) {
    return;
  }
  lists.forEach((list) => {
    list.querySelectorAll('.list-item').forEach((item) => {
      item.classList.remove('drag-over-before', 'drag-over-after');
      delete item.dataset.dropPosition;
    });
  });
}

function reorderAnnouncement(fromIndex, toIndex, position = 'before') {
  const slides = state.project.announcements.slides;
  if (fromIndex < 0 || fromIndex >= slides.length) {
    return;
  }
  if (toIndex < 0 || toIndex >= slides.length) {
    return;
  }
  let insertIndex = toIndex;
  if (position === 'after') {
    insertIndex += 1;
  }
  if (insertIndex > fromIndex) {
    insertIndex -= 1;
  }
  if (insertIndex === fromIndex) {
    clearAnnouncementDragIndicators();
    return;
  }
  const [moved] = slides.splice(fromIndex, 1);
  const maxIndex = slides.length;
  if (insertIndex < 0) {
    insertIndex = 0;
  }
  if (insertIndex > maxIndex) {
    insertIndex = maxIndex;
  }
  slides.splice(insertIndex, 0, moved);

  if (state.selectedAnnouncementIndex === fromIndex) {
    state.selectedAnnouncementIndex = insertIndex;
  } else if (
    fromIndex < insertIndex &&
    state.selectedAnnouncementIndex > fromIndex &&
    state.selectedAnnouncementIndex <= insertIndex
  ) {
    state.selectedAnnouncementIndex -= 1;
  } else if (
    fromIndex > insertIndex &&
    state.selectedAnnouncementIndex >= insertIndex &&
    state.selectedAnnouncementIndex < fromIndex
  ) {
    state.selectedAnnouncementIndex += 1;
  }
  state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
  refreshAll();
  saveProjectIfPossible();
}

function reorderTimer(fromIndex, toIndex, position = 'before') {
  const slides = state.project.timer.slides;
  if (fromIndex < 0 || fromIndex >= slides.length) {
    return;
  }
  if (toIndex < 0 || toIndex >= slides.length) {
    return;
  }
  let insertIndex = toIndex;
  if (position === 'after') {
    insertIndex += 1;
  }
  if (insertIndex > fromIndex) {
    insertIndex -= 1;
  }
  if (insertIndex === fromIndex) {
    clearTimerDragIndicators();
    return;
  }
  const [moved] = slides.splice(fromIndex, 1);
  const maxIndex = slides.length;
  if (insertIndex < 0) {
    insertIndex = 0;
  }
  if (insertIndex > maxIndex) {
    insertIndex = maxIndex;
  }
  slides.splice(insertIndex, 0, moved);

  if (state.selectedTimerIndex === fromIndex) {
    state.selectedTimerIndex = insertIndex;
  } else if (
    fromIndex < insertIndex &&
    state.selectedTimerIndex > fromIndex &&
    state.selectedTimerIndex <= insertIndex
  ) {
    state.selectedTimerIndex -= 1;
  } else if (
    fromIndex > insertIndex &&
    state.selectedTimerIndex >= insertIndex &&
    state.selectedTimerIndex < fromIndex
  ) {
    state.selectedTimerIndex += 1;
  }
  state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
  refreshAll();
  saveProjectIfPossible();
}

function updateSlideFromEditor() {
  const slide = getSelectedSlide();
  if (!slide) {
    return;
  }
  updateSlideVisibility();
  const newLabel = elements.slideLabel.value.trim();
  if (newLabel) {
    slide.label = newLabel;
  }
  slide.showTitle = elements.showTitle.checked;
  slide.showLyrics = elements.showLyrics.checked;
  slide.showFooter = elements.showFooter.checked;
  slide.footerAutoCcli = elements.footerAuto.checked;

  slide.titleText = richTextFromPlain(elements.titleText.value);
  slide.lyricsText = richTextFromPlain(elements.lyricsText.value);
  if (slide.footerAutoCcli) {
    const song = getSelectedSong();
    slide.footerText = richTextFromPlain(buildCcliFooter(song));
  } else {
    slide.footerText = richTextFromPlain(elements.footerText.value);
  }

  renderSlides();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
}

function updateSongTitle() {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  song.title = elements.songTitle.value.trim() || 'Untitled Song';
  renderSetlist();
  updateLiveStatus();
  saveProjectIfPossible();
}

function openHtmlFilePicker() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.jpg,.jpeg,.png,.gif,.bmp,.webp,.mp4,.mov,.mkv,.avi,.webm';
    input.style.display = 'none';
    document.body.appendChild(input);

    input.addEventListener(
      'change',
      () => {
        const file = input.files && input.files[0];
        const path = file && file.path ? file.path : null;
        document.body.removeChild(input);
        resolve(path);
      },
      { once: true }
    );

    input.click();
  });
}

function openHtmlImagePicker() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.jpg,.jpeg,.png,.gif,.bmp,.webp';
    input.style.display = 'none';
    document.body.appendChild(input);

    input.addEventListener(
      'change',
      () => {
        const file = input.files && input.files[0];
        const path = file && file.path ? file.path : null;
        document.body.removeChild(input);
        resolve(path);
      },
      { once: true }
    );

    input.click();
  });
}

function openHtmlLyricsPicker() {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.txt';
    input.style.display = 'none';
    document.body.appendChild(input);

    input.addEventListener(
      'change',
      () => {
        const file = input.files && input.files[0];
        const path = file && file.path ? file.path : null;
        document.body.removeChild(input);
        resolve(path);
      },
      { once: true }
    );

    input.click();
  });
}

async function pickMediaPath() {
  if (!window.api || !window.api.pickMedia) {
    return openHtmlFilePicker();
  }

  try {
    const result = await window.api.pickMedia();
    return result || null;
  } catch (error) {
    console.error('IPC media picker failed, falling back', error);
    return openHtmlFilePicker();
  }
}

async function pickLyricsPath() {
  if (!window.api || !window.api.pickLyrics) {
    return openHtmlLyricsPicker();
  }

  try {
    const result = await window.api.pickLyrics();
    return result || null;
  } catch (error) {
    console.error('IPC lyrics picker failed, falling back', error);
    return openHtmlLyricsPicker();
  }
}

async function pickAnnouncementPath() {
  if (!window.api || !window.api.pickAnnouncement) {
    return openHtmlImagePicker();
  }

  try {
    const result = await window.api.pickAnnouncement();
    return result || null;
  } catch (error) {
    console.error('IPC announcement picker failed, falling back', error);
    return openHtmlImagePicker();
  }
}

async function applyBackgroundFromSource(sourcePath, options = {}) {
  const song = getSelectedSong();
  if (!song) {
    window.alert('Select a song before choosing a background.');
    return;
  }
  if (!sourcePath) {
    elements.backgroundPath.textContent = 'No file selected.';
    return;
  }
  const target = options.target || 'project';
  if (target === 'project' && !state.projectFolder) {
    elements.backgroundPath.textContent = 'Project folder is required.';
    return;
  }
  elements.backgroundPath.textContent = 'Importing media...';
  let imported = null;
  try {
    if (target === 'library') {
      imported = await window.api.importLibrary(sourcePath);
    } else {
      imported = await window.api.importMedia(state.projectFolder, sourcePath);
    }
  } catch (error) {
    console.error('Failed to import media', error);
    elements.backgroundPath.textContent = 'Import failed. See console.';
    window.alert('Import failed. Check the console for details.');
    return;
  }
  if (!imported || !imported.relativePath) {
    elements.backgroundPath.textContent = 'Import failed. No file path.';
    return;
  }
  if (imported.absolutePath) {
    const normalized = imported.absolutePath.replace(/\\/g, '/');
    state.runtimeBackgrounds[song.id] = `file:///${encodeURI(normalized)}`;
  }
  const ext = (imported.relativePath.split('.').pop() || '').toLowerCase();
  const videoExts = ['mp4', 'mov', 'mkv', 'avi', 'webm'];
  song.background.type = videoExts.includes(ext) ? 'video' : 'image';
  song.background.path = imported.relativePath;
  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex || 0 };
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

async function pickBackground() {
  elements.backgroundPath.textContent = 'Opening file picker...';
  const sourcePath = await pickMediaPath();
  if (!sourcePath) {
    elements.backgroundPath.textContent = 'No file selected.';
    return;
  }
  await applyBackgroundFromSource(sourcePath, { target: 'library' });
}

async function applyAnnouncementFromSource(sourcePath) {
  const slide = getSelectedAnnouncementSlide();
  if (!slide) {
    window.alert('Select an announcement slide before choosing an image.');
    return;
  }
  if (!sourcePath) {
    elements.announcementPath.textContent = 'No file selected.';
    return;
  }
  elements.announcementPath.textContent = 'Importing image...';
  let imported = null;
  try {
    imported = await window.api.importAnnouncement(sourcePath);
  } catch (error) {
    console.error('Failed to import announcement media', error);
    elements.announcementPath.textContent = 'Import failed. See console.';
    window.alert('Import failed. Check the console for details.');
    return;
  }
  if (!imported || !imported.relativePath) {
    elements.announcementPath.textContent = 'Import failed. No file path.';
    return;
  }
  slide.mediaType = 'image';
  slide.mediaPath = imported.relativePath;
  renderAnnouncements();
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

async function pickAnnouncementMedia() {
  elements.announcementPath.textContent = 'Opening file picker...';
  const sourcePath = await pickAnnouncementPath();
  if (!sourcePath) {
    elements.announcementPath.textContent = 'No file selected.';
    return;
  }
  await applyAnnouncementFromSource(sourcePath);
}

async function applyTimerMediaFromSource(sourcePath) {
  const slide = getSelectedTimerSlide();
  if (!slide) {
    window.alert('Select a timer slide before choosing media.');
    return;
  }
  if (!sourcePath) {
    elements.timerMediaPath.textContent = 'No file selected.';
    return;
  }
  elements.timerMediaPath.textContent = 'Importing media...';
  let imported = null;
  try {
    imported = await window.api.importLibrary(sourcePath);
  } catch (error) {
    console.error('Failed to import timer media', error);
    elements.timerMediaPath.textContent = 'Import failed. See console.';
    window.alert('Import failed. Check the console for details.');
    return;
  }
  if (!imported || !imported.relativePath) {
    elements.timerMediaPath.textContent = 'Import failed. No file path.';
    return;
  }
  const ext = (imported.relativePath.split('.').pop() || '').toLowerCase();
  const videoExts = ['mp4', 'mov', 'mkv', 'avi', 'webm'];
  slide.mediaType = videoExts.includes(ext) ? 'video' : 'image';
  slide.mediaPath = imported.relativePath;
  renderTimerSlides();
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

async function pickTimerMedia() {
  elements.timerMediaPath.textContent = 'Opening file picker...';
  const sourcePath = await pickMediaPath();
  if (!sourcePath) {
    elements.timerMediaPath.textContent = 'No file selected.';
    return;
  }
  await applyTimerMediaFromSource(sourcePath);
}

function openLibrary(scope = 'background', target = null) {
  state.libraryTarget = target || scope;
  if (window.api && window.api.openLibrary) {
    window.api.openLibrary({ scope });
  }
}

function isTextFileName(name = '') {
  return name.toLowerCase().endsWith('.txt');
}

function isFileDrag(event) {
  if (internalDragActive) {
    return false;
  }
  if (!event.dataTransfer || !event.dataTransfer.types) {
    return false;
  }
  const types = Array.from(event.dataTransfer.types);
  if (types.includes('application/x-wp-internal')) {
    return false;
  }
  return types.includes('Files');
}

function setDropOverlayVisible(visible) {
  if (!elements.dropOverlay) {
    return;
  }
  elements.dropOverlay.hidden = !visible;
}

async function handleLyricsDrop(event) {
  const files = Array.from(event.dataTransfer?.files || []);
  if (files.length === 0) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  const textFiles = files.filter((file) => isTextFileName(file.name || ''));
  if (textFiles.length === 0) {
    window.alert('Only .txt lyric files are supported.');
    return;
  }

  for (const file of textFiles) {
    if (file.path) {
      await importLyricsFromFile(file.path, { forceNew: true });
    } else if (file.text) {
      const text = await file.text();
      await importLyricsFromText(text, { forceNew: true });
    }
  }
}

function isMetadataLine(line, hasMeta) {
  if (!line) {
    return false;
  }
  const lower = line.toLowerCase();
  if (lower.includes('ccli')) {
    return true;
  }
  if (lower.includes('songselect') || lower.includes('www.ccli.com')) {
    return true;
  }
  if (/(public domain|copyright|publisher|publishing|admin|words:|music:)/i.test(line)) {
    return true;
  }
  if (hasMeta && line.includes(',') && !SECTION_HEADER_REGEX.test(line)) {
    return true;
  }
  return false;
}

function parseLyricsFile(rawText) {
  const lines = rawText.split(/\r?\n/);
  let firstLineIndex = 0;
  while (firstLineIndex < lines.length && lines[firstLineIndex].trim() === '') {
    firstLineIndex += 1;
  }
  const title = lines[firstLineIndex] ? lines[firstLineIndex].trim() : '';
  const remaining = lines.slice(firstLineIndex + 1);

  let end = remaining.length - 1;
  while (end >= 0 && remaining[end].trim() === '') {
    end -= 1;
  }

  let metadataStart = end + 1;
  let foundMeta = false;
  for (let i = end; i >= 0; i -= 1) {
    const line = remaining[i].trim();
    if (!line) {
      if (foundMeta) {
        metadataStart = i;
      }
      continue;
    }
    if (isMetadataLine(line, foundMeta)) {
      foundMeta = true;
      metadataStart = i;
      continue;
    }
    break;
  }

  const contentLines = remaining.slice(0, metadataStart);
  const metadataLines = foundMeta
    ? remaining.slice(metadataStart, end + 1).map((line) => line.trim()).filter(Boolean)
    : [];

  const sections = [];
  let current = null;
  contentLines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      if (current && current.lines.length > 0) {
        current.lines.push('');
      }
      return;
    }
    if (SECTION_HEADER_REGEX.test(trimmed)) {
      if (current) {
        while (current.lines.length > 0 && current.lines[current.lines.length - 1] === '') {
          current.lines.pop();
        }
        sections.push(current);
      }
      current = { header: trimmed, lines: [] };
      return;
    }
    if (!current) {
      current = { header: 'Lyrics', lines: [] };
    }
    current.lines.push(trimmed);
  });

  if (current) {
    while (current.lines.length > 0 && current.lines[current.lines.length - 1] === '') {
      current.lines.pop();
    }
    sections.push(current);
  }

  return { title, sections, metadataLines };
}

function applyLyricsImport(parsed, options = {}) {
  if (!parsed) {
    return;
  }
  const forceNew = options.forceNew === true;
  let song = forceNew ? null : getSelectedSong();
  if (!song) {
    song = createSong('New Song');
    song.theme = getDefaultTheme();
    song.slides = [];
    state.project.songs[song.id] = song;
    state.project.setlist.push(song.id);
  } else {
    song.slides = [];
  }

  if (parsed.title) {
    song.title = parsed.title;
  }

  const ccliMeta = {
    songNumber: '',
    authors: [],
    publisher: '',
    copyright: ''
  };

  parsed.metadataLines.forEach((line) => {
    const ccliMatch = line.match(/CCLI\s*(Song)?\s*#\s*(\d+)/i);
    if (ccliMatch) {
      ccliMeta.songNumber = ccliMatch[2];
      return;
    }
    if (!ccliMeta.authors.length && line.includes(',') && !/ccli|songselect|www\./i.test(line)) {
      ccliMeta.authors = line.split(',').map((part) => part.trim()).filter(Boolean);
      return;
    }
    if (!ccliMeta.publisher && /publisher|publishing/i.test(line)) {
      ccliMeta.publisher = line;
      return;
    }
    if (!ccliMeta.copyright && /(public domain|copyright|©|\(c\))/i.test(line)) {
      ccliMeta.copyright = line;
    }
  });

  song.ccli = {
    songNumber: ccliMeta.songNumber || song.ccli.songNumber || '',
    authors: ccliMeta.authors.length ? ccliMeta.authors : song.ccli.authors || [],
    publisher: ccliMeta.publisher || song.ccli.publisher || '',
    copyright: ccliMeta.copyright || song.ccli.copyright || ''
  };

  const sections = parsed.sections.length ? parsed.sections : [{ header: 'Lyrics', lines: [] }];
  sections.forEach((section, index) => {
    const lyrics = section.lines.join('\n').trim();
    const label = section.header || `Slide ${index + 1}`;
    const isFirst = index === 0;
    const slide = createSlide({
      label,
      titleText: isFirst ? song.title : '',
      lyricsText: lyrics,
      footerAutoCcli: isFirst
    });
    slide.showTitle = isFirst;
    slide.showLyrics = true;
    slide.showFooter = isFirst;
    if (!isFirst) {
      slide.footerAutoCcli = false;
    }
    song.slides.push(slide);
  });

  const ccliFooter = buildCcliFooter(song);
  song.slides.forEach((slide) => {
    if (slide.footerAutoCcli) {
      slide.footerText = richTextFromPlain(ccliFooter);
    }
  });

  state.selectedSongId = song.id;
  state.selectedSlideIndex = 0;
  state.selectedSection = 'setlist';
  state.preview = { section: 'setlist', songId: song.id, slideIndex: 0 };
  refreshAll();
  saveProjectIfPossible();
}

async function importLyricsFromText(text, options = {}) {
  const parsed = parseLyricsFile(text || '');
  applyLyricsImport(parsed, options);
}

async function importLyricsFromFile(filePath, options = {}) {
  if (!filePath) {
    return;
  }
  try {
    let text = '';
    if (window.api && window.api.readTextFile) {
      text = await window.api.readTextFile(filePath);
    } else if (window.require) {
      const fs = window.require('fs');
      text = fs.readFileSync(filePath, 'utf8');
    }
    await importLyricsFromText(text, options);
  } catch (error) {
    console.error('Failed to import lyrics file', error);
    window.alert('Failed to import lyrics file. Check the console for details.');
  }
}

async function importLyrics() {
  const filePath = await pickLyricsPath();
  if (!filePath) {
    return;
  }
  const song = getSelectedSong();
  let forceNew = false;
  if (song && song.slides && song.slides.length > 0) {
    const replace = window.confirm(
      'Import into the selected song? This will replace its slides.\n\nClick Cancel to create a new song.'
    );
    forceNew = !replace;
  }
  await importLyricsFromFile(filePath, { forceNew });
}

function applyBackgroundFromLibraryPath(relativePath) {
  const song = getSelectedSong();
  if (!song) {
    window.alert('Select a song before choosing a background.');
    return;
  }
  if (!relativePath) {
    elements.backgroundPath.textContent = 'No file selected.';
    return;
  }
  const ext = (relativePath.split('.').pop() || '').toLowerCase();
  const videoExts = ['mp4', 'mov', 'mkv', 'avi', 'webm'];
  song.background.type = videoExts.includes(ext) ? 'video' : 'image';
  song.background.path = relativePath;
  delete state.runtimeBackgrounds[song.id];
  state.preview = { section: 'setlist', songId: song.id, slideIndex: state.selectedSlideIndex || 0 };
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

function applyAnnouncementFromLibraryPath(relativePath) {
  const slide = getSelectedAnnouncementSlide();
  if (!slide) {
    window.alert('Select an announcement slide before choosing an image.');
    return;
  }
  if (!relativePath) {
    elements.announcementPath.textContent = 'No file selected.';
    return;
  }
  slide.mediaType = 'image';
  slide.mediaPath = relativePath;
  renderAnnouncements();
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

function applyTimerFromLibraryPath(relativePath) {
  const slide = getSelectedTimerSlide();
  if (!slide) {
    window.alert('Select a timer slide before choosing media.');
    return;
  }
  if (!relativePath) {
    elements.timerMediaPath.textContent = 'No file selected.';
    return;
  }
  const ext = (relativePath.split('.').pop() || '').toLowerCase();
  const videoExts = ['mp4', 'mov', 'mkv', 'avi', 'webm'];
  slide.mediaType = videoExts.includes(ext) ? 'video' : 'image';
  slide.mediaPath = relativePath;
  renderTimerSlides();
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

function clearBackground() {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  song.background.path = '';
  delete state.runtimeBackgrounds[song.id];
  renderInspector();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
  scheduleAssetCheck();
}

function updateTheme() {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  const target = elements.themeTarget.value || 'lyrics';
  if (!song.theme.textStyles) {
    song.theme.textStyles = {
      title: resolveTextStyle(song.theme, 'title'),
      lyrics: resolveTextStyle(song.theme, 'lyrics'),
      footer: resolveTextStyle(song.theme, 'footer')
    };
  }
  const style = resolveTextStyle(song.theme, target);
  style.fontFamily = elements.themeFont.value || 'Segoe UI';
  style.fontPx = Number(elements.themeBase.value) || style.fontPx;
  style.color = elements.themeColor.value;
  const strokeEnabled = elements.themeStrokeToggle.checked;
  style.strokeWidthPx = strokeEnabled ? Number(elements.themeStroke.value) || 0 : 0;
  style.strokeColor = elements.themeStrokeColor.value;
  const shadowEnabled = elements.themeShadowToggle.checked;
  if (shadowEnabled) {
    style.shadow = {
      dx: Number(elements.themeShadowX.value) || 0,
      dy: Number(elements.themeShadowY.value) || 0,
      blur: Number(elements.themeShadowBlur.value) || 0,
      color: elements.themeShadowColor.value
    };
  } else {
    style.shadow = {
      dx: 0,
      dy: 0,
      blur: -1,
      color: elements.themeShadowColor.value || '#000000'
    };
  }
  song.theme.textStyles[target] = style;
  song.theme.position = getThemePositionValue();
  song.theme.dimOpacity = Number(elements.themeDim.value) || 0;
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
}

function updateCcli() {
  const song = getSelectedSong();
  if (!song) {
    return;
  }
  song.ccli.songNumber = elements.ccliNumber.value;
  song.ccli.authors = elements.ccliAuthors.value
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean);
  song.ccli.publisher = elements.ccliPublisher.value;
  song.ccli.copyright = elements.ccliCopyright.value;
  const ccliFooter = buildCcliFooter(song);
  song.slides.forEach((slide) => {
    if (slide.footerAutoCcli) {
      slide.footerText = richTextFromPlain(ccliFooter);
    }
  });
  renderSlideEditor();
  updatePreview();
  sendProgramStateIfLive();
  saveProjectIfPossible();
}

function updateAnnouncementSettings() {
  state.project.announcements.autoAdvanceEnabled = elements.announcementAuto.checked;
  state.project.announcements.autoAdvanceSec = Number(elements.announcementAdvance.value) || 15;
  state.project.announcements.loop = elements.announcementLoop.checked;
  if (state.live.section === 'announcements') {
    updateAutoAdvanceTimers();
  }
  saveProjectIfPossible();
}

function updateTimerSettings() {
  state.project.timer.autoAdvanceOnVideoEnd = elements.timerAutoVideo.checked;
  state.project.timer.autoAdvanceImages = elements.timerAutoImages.checked;
  state.project.timer.autoAdvanceSec = Number(elements.timerAdvance.value) || 15;
  if (elements.timerAdvanceRow) {
    elements.timerAdvanceRow.style.display = elements.timerAutoImages.checked ? 'flex' : 'none';
  }
  if (state.live.section === 'timer') {
    updateAutoAdvanceTimers();
  }
  saveProjectIfPossible();
}

function advancePreview() {
  const selection = getPreviewSelection();
  if (!selection) {
    return;
  }
  if (selection.section === 'announcements') {
    const slides = state.project.announcements.slides || [];
    if (selection.slideIndex < slides.length - 1) {
      state.selectedSection = 'announcements';
      state.selectedAnnouncementIndex = selection.slideIndex + 1;
      state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
    }
    updateEditorVisibility();
    renderLineupButtons();
    renderAnnouncements();
    renderInspector();
    updatePreview();
    if (state.autoGoLive) {
      goLive({ usePreviewIndex: true });
    }
    return;
  }
  if (selection.section === 'timer') {
    const slides = state.project.timer.slides || [];
    if (selection.slideIndex < slides.length - 1) {
      state.selectedSection = 'timer';
      state.selectedTimerIndex = selection.slideIndex + 1;
      state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
    }
    updateEditorVisibility();
    renderLineupButtons();
    renderTimerSlides();
    renderInspector();
    updatePreview();
    if (state.autoGoLive) {
      goLive({ usePreviewIndex: true });
    }
    return;
  }

  const song = selection.song;
  if (selection.slideIndex < song.slides.length - 1) {
    state.selectedSection = 'setlist';
    state.preview.slideIndex += 1;
  } else {
    const currentIndex = state.project.setlist.indexOf(song.id);
    const nextSongId = state.project.setlist[currentIndex + 1];
    if (nextSongId) {
      state.selectedSection = 'setlist';
      state.preview = { section: 'setlist', songId: nextSongId, slideIndex: 0 };
      state.selectedSongId = nextSongId;
      state.selectedSlideIndex = 0;
    }
  }
  renderSlides();
  updatePreview();
  if (state.autoGoLive) {
    goLive({ usePreviewIndex: true });
  }
}

function previousPreview() {
  const selection = getPreviewSelection();
  if (!selection) {
    return;
  }
  if (selection.section === 'announcements') {
    if (selection.slideIndex > 0) {
      state.selectedSection = 'announcements';
      state.selectedAnnouncementIndex = selection.slideIndex - 1;
      state.preview = { section: 'announcements', songId: null, slideIndex: state.selectedAnnouncementIndex };
    }
    updateEditorVisibility();
    renderLineupButtons();
    renderAnnouncements();
    renderInspector();
    updatePreview();
    if (state.autoGoLive) {
      goLive({ usePreviewIndex: true });
    }
    return;
  }
  if (selection.section === 'timer') {
    if (selection.slideIndex > 0) {
      state.selectedSection = 'timer';
      state.selectedTimerIndex = selection.slideIndex - 1;
      state.preview = { section: 'timer', songId: null, slideIndex: state.selectedTimerIndex };
    }
    updateEditorVisibility();
    renderLineupButtons();
    renderTimerSlides();
    renderInspector();
    updatePreview();
    if (state.autoGoLive) {
      goLive({ usePreviewIndex: true });
    }
    return;
  }

  if (selection.slideIndex > 0) {
    state.selectedSection = 'setlist';
    state.preview.slideIndex -= 1;
  } else {
    const currentIndex = state.project.setlist.indexOf(selection.song.id);
    const previousSongId = state.project.setlist[currentIndex - 1];
    if (previousSongId) {
      const prevSong = state.project.songs[previousSongId];
      state.selectedSection = 'setlist';
      state.preview = { section: 'setlist', songId: previousSongId, slideIndex: prevSong.slides.length - 1 };
      state.selectedSongId = previousSongId;
      state.selectedSlideIndex = prevSong.slides.length - 1;
    }
  }
  renderSlides();
  updatePreview();
  if (state.autoGoLive) {
    goLive({ usePreviewIndex: true });
  }
}

function advanceLive(options = {}) {
  const selection = getLiveSelection();
  if (!selection) {
    advancePreview();
    return;
  }
  if (selection.section === 'announcements') {
    const slides = state.project.announcements.slides || [];
    if (selection.slideIndex < slides.length - 1) {
      state.live.slideIndex += 1;
    } else if (state.project.announcements.loop) {
      state.live.slideIndex = 0;
    }
    if (state.followLive) {
      state.selectedSection = 'announcements';
      state.selectedAnnouncementIndex = state.live.slideIndex;
      state.preview = { section: 'announcements', songId: null, slideIndex: state.live.slideIndex };
      updateEditorVisibility();
      renderLineupButtons();
      renderAnnouncements();
      renderInspector();
      updatePreview();
    }
    updateLiveStatus();
    sendProgramState();
    updateAutoAdvanceTimers();
    return;
  }

  if (selection.section === 'timer') {
    const slides = state.project.timer.slides || [];
    if (selection.slideIndex < slides.length - 1) {
      state.live.slideIndex += 1;
    }
    if (state.followLive) {
      state.selectedSection = 'timer';
      state.selectedTimerIndex = state.live.slideIndex;
      state.preview = { section: 'timer', songId: null, slideIndex: state.live.slideIndex };
      updateEditorVisibility();
      renderLineupButtons();
      renderTimerSlides();
      renderInspector();
      updatePreview();
    }
    updateLiveStatus();
    sendProgramState();
    updateAutoAdvanceTimers();
    return;
  }

  const song = selection.song;
  if (selection.slideIndex < song.slides.length - 1) {
    state.live.slideIndex += 1;
  } else {
    const currentIndex = state.project.setlist.indexOf(song.id);
    const nextSongId = state.project.setlist[currentIndex + 1];
    if (nextSongId) {
      state.live = { section: 'setlist', songId: nextSongId, slideIndex: 0 };
    }
  }

  if (state.followLive) {
    state.selectedSection = 'setlist';
    state.selectedSongId = state.live.songId;
    state.selectedSlideIndex = state.live.slideIndex;
    state.preview = { section: 'setlist', songId: state.live.songId, slideIndex: state.live.slideIndex };
    renderSlides();
    renderSongDetails();
    renderSlideEditor();
    renderInspector();
    updatePreview();
  }
  updateLiveStatus();
  sendProgramState();
  updateAutoAdvanceTimers();
}

function previousLive() {
  const selection = getLiveSelection();
  if (!selection) {
    previousPreview();
    return;
  }
  if (selection.section === 'announcements') {
    if (selection.slideIndex > 0) {
      state.live.slideIndex -= 1;
    } else if (state.project.announcements.loop) {
      const slides = state.project.announcements.slides || [];
      state.live.slideIndex = slides.length > 0 ? slides.length - 1 : 0;
    }
    if (state.followLive) {
      state.selectedSection = 'announcements';
      state.selectedAnnouncementIndex = state.live.slideIndex;
      state.preview = { section: 'announcements', songId: null, slideIndex: state.live.slideIndex };
      updateEditorVisibility();
      renderLineupButtons();
      renderAnnouncements();
      renderInspector();
      updatePreview();
    }
    updateLiveStatus();
    sendProgramState();
    updateAutoAdvanceTimers();
    return;
  }

  if (selection.section === 'timer') {
    if (selection.slideIndex > 0) {
      state.live.slideIndex -= 1;
    }
    if (state.followLive) {
      state.selectedSection = 'timer';
      state.selectedTimerIndex = state.live.slideIndex;
      state.preview = { section: 'timer', songId: null, slideIndex: state.live.slideIndex };
      updateEditorVisibility();
      renderLineupButtons();
      renderTimerSlides();
      renderInspector();
      updatePreview();
    }
    updateLiveStatus();
    sendProgramState();
    updateAutoAdvanceTimers();
    return;
  }

  if (selection.slideIndex > 0) {
    state.live.slideIndex -= 1;
  } else {
    const currentIndex = state.project.setlist.indexOf(selection.song.id);
    const previousSongId = state.project.setlist[currentIndex - 1];
    if (previousSongId) {
      const prevSong = state.project.songs[previousSongId];
      state.live = { section: 'setlist', songId: previousSongId, slideIndex: prevSong.slides.length - 1 };
    }
  }

  if (state.followLive) {
    state.selectedSection = 'setlist';
    state.selectedSongId = state.live.songId;
    state.selectedSlideIndex = state.live.slideIndex;
    state.preview = { section: 'setlist', songId: state.live.songId, slideIndex: state.live.slideIndex };
    renderSlides();
    renderSongDetails();
    renderSlideEditor();
    renderInspector();
    updatePreview();
  }
  updateLiveStatus();
  sendProgramState();
  updateAutoAdvanceTimers();
}

function goLive(options = {}) {
  const usePreviewIndex = options.usePreviewIndex === true;
  const section = state.selectedSection || 'setlist';
  let slideIndex = usePreviewIndex ? state.preview.slideIndex : 0;
  let songId = null;

  if (section === 'setlist') {
    if (!state.selectedSongId) {
      return;
    }
    songId = state.selectedSongId;
    const song = state.project.songs[songId];
    if (!song || song.slides.length === 0) {
      return;
    }
    if (slideIndex >= song.slides.length) {
      slideIndex = 0;
    }
  } else if (section === 'announcements') {
    const slides = state.project.announcements.slides || [];
    if (slides.length === 0) {
      return;
    }
    if (slideIndex < 0) {
      slideIndex = 0;
    }
    if (slideIndex >= slides.length) {
      slideIndex = 0;
    }
  } else if (section === 'timer') {
    const slides = state.project.timer.slides || [];
    if (slides.length === 0) {
      return;
    }
    if (slideIndex < 0) {
      slideIndex = 0;
    }
    if (slideIndex >= slides.length) {
      slideIndex = 0;
    }
  }
  if (state.panic) {
    state.panic = false;
    elements.panic.textContent = 'Panic';
  }
  if (document.activeElement && typeof document.activeElement.blur === 'function') {
    document.activeElement.blur();
  }
  if (window.api && window.api.showProgram) {
    if (state.displayId != null) {
      window.api.showProgram(state.displayId);
    } else {
      window.api.showProgram(null);
    }
    window.setTimeout(sendProgramState, 300);
  }
  state.live = { section, songId, slideIndex };
  state.preview = { section, songId, slideIndex };
  if (section === 'setlist') {
    state.selectedSection = 'setlist';
    state.selectedSongId = songId;
    state.selectedSlideIndex = slideIndex;
    renderSlides();
    renderSongDetails();
    renderSlideEditor();
  } else if (section === 'announcements') {
    state.selectedSection = 'announcements';
    state.selectedAnnouncementIndex = slideIndex;
    renderAnnouncements();
  } else if (section === 'timer') {
    state.selectedSection = 'timer';
    state.selectedTimerIndex = slideIndex;
    renderTimerSlides();
  }
  updateEditorVisibility();
  renderLineupButtons();
  renderInspector();
  updateLiveStatus();
  sendProgramState();
  updateAutoAdvanceTimers();
}

function sendProgramStateIfLive() {
  const selection = getLiveSelection();
  if (!selection) {
    return;
  }
  if (selection.section === 'announcements') {
    if (state.selectedSection === 'announcements' && state.selectedAnnouncementIndex === selection.slideIndex) {
      sendProgramState();
    }
    return;
  }
  if (selection.section === 'timer') {
    if (state.selectedSection === 'timer' && state.selectedTimerIndex === selection.slideIndex) {
      sendProgramState();
    }
    return;
  }
  if (selection.song && selection.song.id === state.selectedSongId) {
    sendProgramState();
  }
}

function togglePanic() {
  state.panic = !state.panic;
  elements.panic.textContent = state.panic ? 'Panic (ON)' : 'Panic';
  sendProgramState();
}

async function refreshDisplays() {
  const displays = await window.api.listDisplays();
  elements.displaySelect.innerHTML = '';
  displays.forEach((display) => {
    const option = document.createElement('option');
    option.value = String(display.id);
    option.textContent = display.label;
    elements.displaySelect.appendChild(option);
  });
  if (displays.length > 0) {
    state.displayId = displays[0].id;
    elements.displaySelect.value = String(state.displayId);
  }
}

function saveProjectIfPossible() {
  if (!state.projectFile || !state.autoSaveEnabled) {
    return;
  }
  if (saveTimer) {
    window.clearTimeout(saveTimer);
  }
  saveTimer = window.setTimeout(() => {
    window.api.saveProjectFile(state.projectFile, state.project, { projectFolder: state.projectFolder });
    saveTimer = null;
  }, 400);
}

async function saveProject() {
  if (!window.api || !window.api.saveProjectFile) {
    return;
  }
  const filePath = await window.api.saveProjectFile(state.projectFile, state.project, { projectFolder: state.projectFolder });
  if (!filePath) {
    return;
  }
  state.projectFile = filePath;
  elements.projectPath.textContent = filePath;
}

async function createProject() {
  const project = createNewProject();
  setProject(project, null, null);
}

async function openProject() {
  if (!window.api || !window.api.openProjectFile) {
    return;
  }
  const payload = await window.api.openProjectFile();
  if (!payload || !payload.project) {
    return;
  }
  setProject(payload.project, payload.filePath || null, null);
}

function handleKeydown(event) {
  const tag = event.target.tagName.toLowerCase();
  const isTyping = tag === 'input' || tag === 'textarea' || event.target.isContentEditable;
  if (isTyping) {
    return;
  }
  if (event.code === 'Space' || event.code === 'ArrowRight' || event.code === 'PageDown') {
    event.preventDefault();
    if (state.live.section) {
      advanceLive();
    } else {
      advancePreview();
    }
  } else if (event.code === 'ArrowLeft' || event.code === 'PageUp') {
    event.preventDefault();
    if (state.live.section) {
      previousLive();
    } else {
      previousPreview();
    }
  } else if (event.code === 'Enter') {
    event.preventDefault();
    goLive();
  } else if (event.code === 'Escape') {
    event.preventDefault();
    togglePanic();
  } else if (event.code === 'Delete') {
    event.preventDefault();
    if (state.selectedSection === 'announcements') {
      if (state.selectedAnnouncementIndex >= 0) {
        deleteAnnouncementSlide(state.selectedAnnouncementIndex);
      }
    } else if (state.selectedSection === 'timer') {
      if (state.selectedTimerIndex >= 0) {
        deleteTimerSlide(state.selectedTimerIndex);
      }
    } else {
      if (state.selectedSongId) {
        deleteSlide(state.selectedSlideIndex);
      }
    }
  } else if (event.code === 'F11') {
    event.preventDefault();
    window.api.toggleProgramFullscreen();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault();
    saveProject();
  }
}

function bindEvents() {
  if (elements.newProject) {
    elements.newProject.addEventListener('click', createProject);
  }
  if (elements.openProject) {
    elements.openProject.addEventListener('click', openProject);
  }
  if (elements.saveProject) {
    elements.saveProject.addEventListener('click', saveProject);
  }
  if (elements.fileMenu) {
    elements.fileMenu.addEventListener('click', (event) => {
      if (event.target && event.target.tagName === 'BUTTON') {
        elements.fileMenu.removeAttribute('open');
      }
    });
  }
  if (elements.addAnnouncement) {
    elements.addAnnouncement.addEventListener('click', addAnnouncementSlide);
  }
  if (elements.addTimerSlide) {
    elements.addTimerSlide.addEventListener('click', addTimerSlide);
  }
  if (elements.announcementsButton) {
    elements.announcementsButton.addEventListener('click', () => {
      selectSection('announcements');
      refreshAll();
    });
  }
  if (elements.timerButton) {
    elements.timerButton.addEventListener('click', () => {
      selectSection('timer');
      refreshAll();
    });
  }
  elements.addSong.addEventListener('click', addSong);
  elements.songTitle.addEventListener('input', updateSongTitle);
  elements.addSlide.addEventListener('click', addSlide);

  [
    elements.showTitle,
    elements.showLyrics,
    elements.showFooter,
    elements.slideLabel,
    elements.titleText,
    elements.lyricsText,
    elements.footerText,
    elements.footerAuto
  ].forEach((input) => input.addEventListener('input', updateSlideFromEditor));

  [elements.showTitle, elements.showLyrics, elements.showFooter].forEach((input) => {
    input.addEventListener('change', updateSlideVisibility);
  });

  elements.footerAuto.addEventListener('change', () => {
    renderSlideEditor();
    updateSlideFromEditor();
  });

  elements.uploadBackground.addEventListener('click', pickBackground);
  elements.openLibrary.addEventListener('click', () => openLibrary('background', 'song'));
  if (elements.openAnnouncementLibrary) {
    elements.openAnnouncementLibrary.addEventListener('click', () => openLibrary('announcements', 'announcements'));
  }
  if (elements.openTimerLibrary) {
    elements.openTimerLibrary.addEventListener('click', () => openLibrary('background', 'timer'));
  }
  if (elements.uploadAnnouncement) {
    elements.uploadAnnouncement.addEventListener('click', pickAnnouncementMedia);
  }
  if (elements.uploadTimerMedia) {
    elements.uploadTimerMedia.addEventListener('click', pickTimerMedia);
  }
  if (elements.importLyrics) {
    elements.importLyrics.addEventListener('click', importLyrics);
  }

  [
    elements.themeFont,
    elements.themeBase,
    elements.themeColor,
    elements.themeStroke,
    elements.themeStrokeColor,
    elements.themeShadowX,
    elements.themeShadowY,
    elements.themeShadowBlur,
    elements.themeShadowColor,
    elements.themeDim
  ].forEach((input) => input.addEventListener('input', updateTheme));

  elements.themeTarget.addEventListener('change', () => {
    state.themeTarget = elements.themeTarget.value || 'lyrics';
    updateThemeEditorFromSelection(getSelectedSong());
  });

  if (elements.themePosition) {
    elements.themePosition.querySelectorAll('button[data-value]').forEach((button) => {
      button.addEventListener('click', () => {
        setThemePositionValue(button.dataset.value || 'center');
        updateTheme();
      });
    });
  }

  if (elements.themeToggle) {
    elements.themeToggle.addEventListener('click', () => {
      const isHidden = elements.themeSection && elements.themeSection.hidden;
      setThemeSectionCollapsed(!isHidden);
    });
  }

  if (elements.ccliToggle) {
    elements.ccliToggle.addEventListener('click', () => {
      const isHidden = elements.ccliSection && elements.ccliSection.hidden;
      setCcliSectionCollapsed(!isHidden);
    });
  }

  elements.themeStrokeToggle.addEventListener('change', () => {
    updateThemeVisibility();
    updateTheme();
  });
  elements.themeShadowToggle.addEventListener('change', () => {
    updateThemeVisibility();
    updateTheme();
  });

  [
    elements.ccliNumber,
    elements.ccliAuthors,
    elements.ccliPublisher,
    elements.ccliCopyright
  ].forEach((input) => input.addEventListener('input', updateCcli));

  if (elements.announcementAuto) {
    elements.announcementAuto.addEventListener('change', updateAnnouncementSettings);
  }
  if (elements.announcementAdvance) {
    elements.announcementAdvance.addEventListener('input', updateAnnouncementSettings);
  }
  if (elements.announcementLoop) {
    elements.announcementLoop.addEventListener('change', updateAnnouncementSettings);
  }
  if (elements.timerAutoVideo) {
    elements.timerAutoVideo.addEventListener('change', updateTimerSettings);
  }
  if (elements.timerAutoImages) {
    elements.timerAutoImages.addEventListener('change', updateTimerSettings);
  }
  if (elements.timerAdvance) {
    elements.timerAdvance.addEventListener('input', updateTimerSettings);
  }

  if (elements.showProgram) {
    elements.showProgram.addEventListener('click', () => {
      if (state.displayId != null) {
        window.api.showProgram(state.displayId);
      } else {
        window.api.showProgram(null);
      }
      sendProgramState();
      window.setTimeout(sendProgramState, 300);
      updateAutoAdvanceTimers();
    });
  }
  if (elements.hideProgram) {
    elements.hideProgram.addEventListener('click', () => {
      if (window.api && window.api.hideProgram) {
        window.api.hideProgram();
      }
      state.live = { section: null, songId: null, slideIndex: 0 };
      updateLiveStatus();
      updateAutoAdvanceTimers();
    });
  }
  elements.goLive.addEventListener('click', goLive);
  elements.prevSlide.addEventListener('click', () => {
    if (state.live.section) {
      previousLive();
    } else {
      previousPreview();
    }
  });
  elements.nextSlide.addEventListener('click', () => {
    if (state.live.section) {
      advanceLive();
    } else {
      advancePreview();
    }
  });
  elements.panic.addEventListener('click', togglePanic);
  elements.autoGoLive.addEventListener('change', (event) => {
    state.autoGoLive = event.target.checked;
  });
  if (elements.followLive) {
    elements.followLive.addEventListener('change', (event) => {
      state.followLive = event.target.checked;
    });
  }
  if (elements.announcementAuto) {
    elements.announcementAuto.addEventListener('change', () => {
      if (elements.announcementAdvance) {
        elements.announcementAdvance.disabled = !elements.announcementAuto.checked;
      }
    });
  }

  elements.displaySelect.addEventListener('change', (event) => {
    state.displayId = Number(event.target.value);
    window.api.setProgramDisplay(state.displayId);
  });

  document.addEventListener('keydown', handleKeydown);
  document.addEventListener('click', () => {
    hideSlideContextMenu();
    hideSongContextMenu();
    hideMediaContextMenu();
  });
  document.addEventListener('dragenter', (event) => {
    if (internalDragActive) {
      return;
    }
    if (!isFileDrag(event)) {
      return;
    }
    dropCounter += 1;
    setDropOverlayVisible(true);
  });
  document.addEventListener('dragend', () => {
    internalDragActive = false;
  });
  document.addEventListener('dragleave', (event) => {
    if (internalDragActive) {
      return;
    }
    if (!isFileDrag(event)) {
      return;
    }
    dropCounter = Math.max(0, dropCounter - 1);
    if (dropCounter === 0) {
      setDropOverlayVisible(false);
    }
  });
  document.addEventListener('dragover', (event) => {
    if (internalDragActive) {
      return;
    }
    if (!isFileDrag(event)) {
      return;
    }
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  });
  document.addEventListener('drop', (event) => {
    if (internalDragActive) {
      internalDragActive = false;
      return;
    }
    if (!isFileDrag(event)) {
      return;
    }
    dropCounter = 0;
    setDropOverlayVisible(false);
    handleLyricsDrop(event);
  });
  document.addEventListener('contextmenu', (event) => {
    if (!event.target.closest('#slide-context')) {
      hideSlideContextMenu();
    }
    if (!event.target.closest('#song-context')) {
      hideSongContextMenu();
    }
    if (!event.target.closest('#media-context')) {
      hideMediaContextMenu();
    }
  });

  if (elements.slideContext) {
    elements.slideContext.addEventListener('click', (event) => {
      const action = event.target.dataset.action;
      const index = Number(elements.slideContext.dataset.index);
      if (!action || Number.isNaN(index)) {
        return;
      }
      if (action === 'copy') {
        copySlide(index);
      } else if (action === 'delete') {
        deleteSlide(index);
      }
      hideSlideContextMenu();
    });
  }

  if (elements.songContext) {
    elements.songContext.addEventListener('click', (event) => {
      const action = event.target.dataset.action;
      const index = Number(elements.songContext.dataset.index);
      if (!action || Number.isNaN(index)) {
        return;
      }
      const songId = state.project.setlist[index];
      if (!songId) {
        return;
      }
      if (action === 'copy') {
        copySong(songId);
      } else if (action === 'delete') {
        deleteSong(songId);
      }
      hideSongContextMenu();
    });
  }

  if (elements.mediaContext) {
    elements.mediaContext.addEventListener('click', (event) => {
      const action = event.target.dataset.action;
      const index = Number(elements.mediaContext.dataset.index);
      const section = elements.mediaContext.dataset.section;
      if (!action || Number.isNaN(index)) {
        return;
      }
      if (action === 'delete') {
        if (section === 'announcements') {
          deleteAnnouncementSlide(index);
        } else if (section === 'timer') {
          deleteTimerSlide(index);
        }
      }
      hideMediaContextMenu();
    });
  }
}

bindEvents();
refreshDisplays();
refreshAll();
setThemeSectionCollapsed(false);
setCcliSectionCollapsed(true);

if (window.api && window.api.onLibrarySelected) {
  window.api.onLibrarySelected((payload) => {
    const scope = payload && payload.scope ? payload.scope : null;
    const relativePath = payload && payload.sourcePath ? payload.sourcePath : payload;
    if (scope === 'announcements' || state.libraryTarget === 'announcements') {
      applyAnnouncementFromLibraryPath(relativePath);
    } else if (state.libraryTarget === 'timer') {
      applyTimerFromLibraryPath(relativePath);
    } else {
      applyBackgroundFromLibraryPath(relativePath);
    }
    state.libraryTarget = null;
  });
}

  if (window.api && window.api.onProgramEvent) {
    window.api.onProgramEvent((payload) => {
      if (!payload) {
        return;
      }
      if (payload.type === 'program-hidden') {
        state.live = { section: null, songId: null, slideIndex: 0 };
        updateLiveStatus();
        updateAutoAdvanceTimers();
        return;
      }
      if (payload.type !== 'video-ended') {
        return;
      }
      if (state.live.section !== 'timer') {
        return;
      }
      if (state.project.timer.autoAdvanceOnVideoEnd === false) {
        return;
      }
      const selection = getLiveSelection();
      if (!selection || selection.section !== 'timer') {
        return;
      }
      if (selection.slide.mediaType !== 'video') {
        return;
      }
      advanceLive({ source: 'timer-video' });
    });
  }

if (window.api && window.api.onMenuAction) {
  window.api.onMenuAction((action) => {
    if (action === 'new-project') {
      createProject();
    } else if (action === 'open-project') {
      openProject();
    } else if (action === 'save-project') {
      saveProject();
    }
  });
}

if (window.ResizeObserver && elements.previewStage) {
  previewObserver = new ResizeObserver(() => {
    if (previewSplitState) {
      setPreviewCardHeight(previewSplitState.height || 0);
    }
    updatePreviewScale();
    updatePreview();
  });
  previewObserver.observe(elements.previewStage);
  if (elements.previewWrap) {
    previewObserver.observe(elements.previewWrap);
  }
}

window.addEventListener('resize', () => {
  updatePreviewScale();
  updatePreview();
});
