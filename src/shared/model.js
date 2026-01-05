const DEFAULT_THEME = {
  fontFamily: 'Segoe UI',
  baseFontPx: 70,
  minFontPx: 36,
  textColor: '#FFFFFF',
  strokeWidthPx: 1,
  strokeColor: '#000000',
  shadow: {
    dx: 2,
    dy: 2,
    blur: 6,
    color: '#000000'
  },
  textStyles: {
    title: {
      fontFamily: 'Segoe UI',
      fontPx: 70,
      color: '#FFFFFF',
      strokeWidthPx: 1,
      strokeColor: '#000000',
      shadow: { dx: 2, dy: 2, blur: 6, color: '#000000' }
    },
    lyrics: {
      fontFamily: 'Segoe UI',
      fontPx: 70,
      color: '#FFFFFF',
      strokeWidthPx: 1,
      strokeColor: '#000000',
      shadow: { dx: 2, dy: 2, blur: 6, color: '#000000' }
    },
    footer: {
      fontFamily: 'Segoe UI',
      fontPx: 25,
      color: '#FFFFFF',
      strokeWidthPx: 1,
      strokeColor: '#000000',
      shadow: { dx: 2, dy: 2, blur: 6, color: '#000000' }
    }
  },
  position: 'center',
  dimOpacity: 0.25
};

export function createNewProject() {
  return {
    schemaVersion: 1,
    appVersion: '0.1.0',
    settings: {
      aspectRatio: '16:9',
      textFadeMs: 200,
      bgCrossfadeMs: 700,
      safeMarginsPct: 5
    },
    announcements: {
      slides: [],
      autoAdvanceSec: 15,
      loop: false,
      autoAdvanceEnabled: true
    },
    timer: {
      slides: [],
      autoAdvanceOnVideoEnd: true,
      autoAdvanceImages: false,
      autoAdvanceSec: 15
    },
    setlist: [],
    songs: {}
  };
}

export function createSong(title = 'New Song') {
  const songId = `song-${crypto.randomUUID()}`;
  const song = {
    id: songId,
    title,
    ccli: {
      songNumber: '',
      authors: [],
      publisher: '',
      copyright: ''
    },
    background: {
      type: 'image',
      path: ''
    },
    theme: { ...DEFAULT_THEME },
    slides: []
  };
  return song;
}

export function createSlide({
  label = 'Slide',
  titleText = '',
  lyricsText = '',
  footerText = '',
  footerAutoCcli = true
} = {}) {
  return {
    id: `slide-${crypto.randomUUID()}`,
    label,
    template: 'TitleLyricsFooter',
    showTitle: Boolean(titleText),
    showLyrics: Boolean(lyricsText),
    showFooter: footerAutoCcli || Boolean(footerText),
    footerAutoCcli,
    titleText: richTextFromPlain(titleText),
    lyricsText: richTextFromPlain(lyricsText),
    footerText: richTextFromPlain(footerText)
  };
}

export function createMediaSlide({ label = 'Slide', mediaPath = '', mediaType = 'image' } = {}) {
  return {
    id: `media-${crypto.randomUUID()}`,
    label,
    mediaPath,
    mediaType
  };
}

export function richTextFromPlain(text) {
  return {
    blocks: [
      {
        paragraphStyle: { align: 'center', lineSpacing: 1.1 },
        runs: [
          { text: text || '' }
        ]
      }
    ]
  };
}

export function plainFromRichText(richText) {
  if (!richText || !richText.blocks) {
    return '';
  }
  return richText.blocks
    .map((block) => (block.runs || []).map((run) => run.text || '').join(''))
    .join('\n');
}

export function getDefaultTheme() {
  return {
    ...DEFAULT_THEME,
    shadow: { ...DEFAULT_THEME.shadow },
    textStyles: {
      title: { ...DEFAULT_THEME.textStyles.title, shadow: { ...DEFAULT_THEME.textStyles.title.shadow } },
      lyrics: { ...DEFAULT_THEME.textStyles.lyrics, shadow: { ...DEFAULT_THEME.textStyles.lyrics.shadow } },
      footer: { ...DEFAULT_THEME.textStyles.footer, shadow: { ...DEFAULT_THEME.textStyles.footer.shadow } }
    }
  };
}
