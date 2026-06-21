const crypto = require('node:crypto');
const { z } = require('zod');

const PROJECT_SCHEMA_VERSION = 2;
const MAX_SONGS = 1000;
const MAX_SLIDES = 512;
const MAX_NOTES_LENGTH = 20000;
const RESERVED_RECORD_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

const unknownRecord = z.record(z.string(), z.unknown());
const richTextSchema = z.object({
  blocks: z.array(z.object({
    paragraphStyle: unknownRecord.optional(),
    runs: z.array(z.object({ text: z.string() }).passthrough()).max(2000)
  }).passthrough()).max(2000)
}).passthrough();

const textSlideSchema = z.object({
  id: z.string().min(1).max(256),
  label: z.string().max(1000),
  template: z.string().min(1).max(128),
  showTitle: z.boolean(),
  showLyrics: z.boolean(),
  showFooter: z.boolean(),
  footerAutoCcli: z.boolean(),
  titleText: richTextSchema,
  lyricsText: richTextSchema,
  footerText: richTextSchema,
  speakerNotes: z.string().max(MAX_NOTES_LENGTH)
}).passthrough();

const mediaSlideSchema = z.object({
  id: z.string().min(1).max(256),
  label: z.string().max(1000),
  mediaPath: z.string().max(32768),
  mediaType: z.enum(['image', 'video']),
  hideDuringLoop: z.boolean(),
  speakerNotes: z.string().max(MAX_NOTES_LENGTH)
}).passthrough();

const librarySourceSchema = z.object({
  id: z.string().min(1).max(128),
  revision: z.number().int().positive()
}).strict();

const songSchema = z.object({
  id: z.string().min(1).max(256),
  title: z.string().max(1000),
  ccli: z.object({
    songNumber: z.string().max(256),
    authors: z.array(z.string().max(1000)).max(100),
    publisher: z.string().max(2000),
    copyright: z.string().max(4000)
  }).passthrough(),
  background: z.object({
    type: z.enum(['image', 'video']),
    path: z.string().max(32768)
  }).passthrough(),
  theme: unknownRecord,
  slides: z.array(textSlideSchema).max(MAX_SLIDES),
  librarySource: librarySourceSchema.nullable()
}).passthrough();

const songsSchema = z.record(z.string(), songSchema)
  .refine(
    (songs) => Object.keys(songs).length <= MAX_SONGS,
    `A project cannot contain more than ${MAX_SONGS} songs.`
  )
  .refine(
    (songs) => Object.keys(songs).every((songId) => !RESERVED_RECORD_KEYS.has(songId)),
    'Project contains a reserved song ID.'
  );

const projectSchema = z.object({
  schemaVersion: z.literal(PROJECT_SCHEMA_VERSION),
  appVersion: z.string().max(64),
  settings: unknownRecord,
  announcements: z.object({
    slides: z.array(mediaSlideSchema).max(MAX_SLIDES),
    autoAdvanceSec: z.number().finite().nonnegative(),
    loop: z.boolean(),
    autoAdvanceEnabled: z.boolean()
  }).passthrough(),
  timer: z.object({
    slides: z.array(mediaSlideSchema).max(MAX_SLIDES),
    autoAdvanceOnVideoEnd: z.boolean(),
    autoAdvanceImages: z.boolean(),
    autoAdvanceSec: z.number().finite().nonnegative()
  }).passthrough(),
  setlist: z.array(z.string().min(1).max(256)).max(MAX_SONGS),
  songs: songsSchema
}).passthrough().superRefine((project, context) => {
  project.setlist.forEach((songId, index) => {
    if (!project.songs[songId]) {
      context.addIssue({
        code: 'custom',
        path: ['setlist', index],
        message: `Setlist references missing song ${songId}.`
      });
    }
  });
});

function richTextFromPlain(text = '') {
  return {
    blocks: [{ paragraphStyle: { align: 'center', lineSpacing: 1.1 }, runs: [{ text }] }]
  };
}

function plainFromRichText(richText) {
  if (!Array.isArray(richText?.blocks)) return '';
  return richText.blocks
    .map((block) => Array.isArray(block?.runs) ? block.runs.map((run) => run?.text || '').join('') : '')
    .join('\n');
}

function normalizeMediaSlide(slide, index) {
  const source = slide && typeof slide === 'object' ? slide : {};
  const mediaPath = source.mediaPath || source.path || source.background?.path || '';
  const mediaType = source.mediaType || source.type || source.background?.type || 'image';
  return {
    ...source,
    id: source.id || `media-${crypto.randomUUID()}`,
    label: source.label || `Slide ${index + 1}`,
    mediaPath: String(mediaPath),
    mediaType: mediaType === 'video' ? 'video' : 'image',
    hideDuringLoop: source.hideDuringLoop === true,
    speakerNotes: typeof source.speakerNotes === 'string' ? source.speakerNotes : ''
  };
}

function normalizeTextSlide(slide, index) {
  const source = slide && typeof slide === 'object' ? slide : {};
  const titleText = source.titleText || richTextFromPlain();
  const lyricsText = source.lyricsText || richTextFromPlain();
  const footerText = source.footerText || richTextFromPlain();
  return {
    ...source,
    id: source.id || `slide-${crypto.randomUUID()}`,
    label: source.label || `Slide ${index + 1}`,
    template: source.template || 'TitleLyricsFooter',
    showTitle: source.showTitle ?? Boolean(plainFromRichText(titleText)),
    showLyrics: source.showLyrics ?? Boolean(plainFromRichText(lyricsText)),
    showFooter: source.showFooter ?? Boolean(plainFromRichText(footerText)),
    footerAutoCcli: source.footerAutoCcli === true,
    titleText,
    lyricsText,
    footerText,
    speakerNotes: typeof source.speakerNotes === 'string' ? source.speakerNotes : ''
  };
}

function normalizeLibrarySource(value) {
  if (!value || typeof value !== 'object') return null;
  if (typeof value.id !== 'string' || !Number.isInteger(value.revision) || value.revision < 1) return null;
  return { id: value.id, revision: value.revision };
}

function migrateProject(project) {
  if (!project || typeof project !== 'object' || Array.isArray(project)) {
    throw new Error('Project data must be an object.');
  }
  const source = structuredClone(project);
  const sourceVersion = source.schemaVersion == null ? 1 : Number(source.schemaVersion);
  if (!Number.isInteger(sourceVersion) || sourceVersion < 1 || sourceVersion > PROJECT_SCHEMA_VERSION) {
    throw new Error(`Unsupported project schema version: ${source.schemaVersion}.`);
  }

  const songs = {};
  Object.entries(source.songs && typeof source.songs === 'object' ? source.songs : {}).forEach(([songId, value]) => {
    if (RESERVED_RECORD_KEYS.has(songId)) throw new Error(`Project contains reserved song ID: ${songId}.`);
    const song = value && typeof value === 'object' ? value : {};
    const ccli = song.ccli && typeof song.ccli === 'object' ? song.ccli : {};
    const background = song.background && typeof song.background === 'object' ? song.background : {};
    songs[songId] = {
      ...song,
      id: song.id || songId,
      title: song.title || 'Untitled Song',
      ccli: {
        ...ccli,
        songNumber: typeof ccli.songNumber === 'string' ? ccli.songNumber : '',
        authors: Array.isArray(ccli.authors) ? ccli.authors.map(String) : [],
        publisher: typeof ccli.publisher === 'string' ? ccli.publisher : '',
        copyright: typeof ccli.copyright === 'string' ? ccli.copyright : ''
      },
      background: {
        ...background,
        type: background.type === 'video' ? 'video' : 'image',
        path: typeof background.path === 'string' ? background.path : ''
      },
      theme: song.theme && typeof song.theme === 'object' ? song.theme : {},
      slides: Array.isArray(song.slides) ? song.slides.map(normalizeTextSlide) : [],
      librarySource: normalizeLibrarySource(song.librarySource)
    };
  });

  const announcementSource = source.announcements && typeof source.announcements === 'object'
    ? source.announcements
    : {};
  const timerSource = source.timer && typeof source.timer === 'object' ? source.timer : {};
  const migrated = {
    ...source,
    schemaVersion: PROJECT_SCHEMA_VERSION,
    appVersion: typeof source.appVersion === 'string' ? source.appVersion : '0.1.5',
    settings: source.settings && typeof source.settings === 'object' ? source.settings : {},
    announcements: {
      ...announcementSource,
      slides: Array.isArray(announcementSource.slides)
        ? announcementSource.slides.map(normalizeMediaSlide)
        : [],
      autoAdvanceSec: Number.isFinite(Number(announcementSource.autoAdvanceSec))
        ? Number(announcementSource.autoAdvanceSec)
        : 15,
      loop: announcementSource.loop !== false,
      autoAdvanceEnabled: typeof announcementSource.autoAdvanceEnabled === 'boolean'
        ? announcementSource.autoAdvanceEnabled
        : announcementSource.loop !== false
    },
    timer: {
      ...timerSource,
      slides: Array.isArray(timerSource.slides) ? timerSource.slides.map(normalizeMediaSlide) : [],
      autoAdvanceOnVideoEnd: timerSource.autoAdvanceOnVideoEnd !== false,
      autoAdvanceImages: timerSource.autoAdvanceImages === true,
      autoAdvanceSec: Number.isFinite(Number(timerSource.autoAdvanceSec)) ? Number(timerSource.autoAdvanceSec) : 15
    },
    setlist: Array.isArray(source.setlist) ? source.setlist.map(String) : [],
    songs
  };
  return projectSchema.parse(migrated);
}

function validateProject(project) {
  return projectSchema.parse(project);
}

module.exports = {
  MAX_NOTES_LENGTH,
  PROJECT_SCHEMA_VERSION,
  migrateProject,
  projectSchema,
  songSchema,
  validateProject
};
