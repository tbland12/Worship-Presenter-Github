const fsp = require('node:fs/promises');
const path = require('node:path');
const { z } = require('zod');
const { AtomicJsonStore } = require('./atomic-json-store');
const { projectSchema, songSchema } = require('./project-v2');

const preferencesSchema = z.object({
  version: z.literal(1),
  stageDisplayTarget: z.string().max(256).nullable(),
  shortcuts: z.record(z.string().max(128), z.array(z.string().max(128)).max(8))
}).strict();

const recoverySchema = z.object({
  version: z.literal(1),
  updatedAt: z.string().datetime({ offset: true }),
  sourceSessionPath: z.string().max(32768).nullable(),
  project: projectSchema
}).strict().nullable();

const historySongSchema = z.object({
  id: z.string().min(1).max(256),
  title: z.string().max(1000),
  ccliSongNumber: z.string().max(256)
}).strict();

const historySchema = z.object({
  version: z.literal(1),
  entries: z.array(z.object({
    id: z.string().min(1).max(256),
    filePath: z.string().min(1).max(32768),
    title: z.string().max(1000),
    lastSavedAt: z.string().datetime({ offset: true }),
    slideCount: z.number().int().nonnegative(),
    songs: z.array(historySongSchema).max(1000)
  }).strict()).max(50)
}).strict();

const songLibraryIndexSchema = z.object({
  version: z.literal(1),
  items: z.array(z.object({
    id: z.string().min(1).max(128),
    revision: z.number().int().positive(),
    title: z.string().max(1000),
    ccliSongNumber: z.string().max(256),
    authors: z.array(z.string().max(1000)).max(100),
    updatedAt: z.string().datetime({ offset: true })
  }).strict()).max(10000)
}).strict();

const songLibraryRecordSchema = z.object({
  version: z.literal(1),
  id: z.string().regex(/^[a-zA-Z0-9-]{1,128}$/),
  revision: z.number().int().positive(),
  updatedAt: z.string().datetime({ offset: true }),
  song: songSchema
}).strict();

function createPersistenceStores(userDataPath) {
  const stateRoot = path.join(userDataPath, 'state');
  const songLibraryRoot = path.join(userDataPath, 'song-library');
  const preferences = new AtomicJsonStore({
    filePath: path.join(stateRoot, 'preferences.json'),
    schema: preferencesSchema,
    defaultValue: { version: 1, stageDisplayTarget: null, shortcuts: {} },
    maxBytes: 256 * 1024
  });
  const recovery = new AtomicJsonStore({
    filePath: path.join(userDataPath, 'recovery', 'latest.json'),
    schema: recoverySchema,
    defaultValue: null,
    maxBytes: 2 * 1024 * 1024
  });
  const history = new AtomicJsonStore({
    filePath: path.join(userDataPath, 'history', 'recent.json'),
    schema: historySchema,
    defaultValue: { version: 1, entries: [] },
    maxBytes: 1024 * 1024
  });
  const songLibraryIndex = new AtomicJsonStore({
    filePath: path.join(songLibraryRoot, 'index.json'),
    schema: songLibraryIndexSchema,
    defaultValue: { version: 1, items: [] },
    maxBytes: 1024 * 1024
  });

  return {
    preferences,
    recovery,
    history,
    songLibraryIndex,
    songRecord(id) {
      if (!/^[a-zA-Z0-9-]{1,128}$/.test(id)) throw new Error('Song library ID is invalid.');
      return new AtomicJsonStore({
        filePath: path.join(songLibraryRoot, 'songs', `${id}.json`),
        schema: songLibraryRecordSchema,
        defaultValue: {
          version: 1,
          id,
          revision: 1,
          updatedAt: new Date(0).toISOString(),
          song: {
            id: `song-${id}`,
            title: '',
            ccli: { songNumber: '', authors: [], publisher: '', copyright: '' },
            background: { type: 'image', path: '' },
            theme: {},
            slides: [],
            librarySource: { id, revision: 1 }
          }
        },
        maxBytes: 1024 * 1024
      });
    },
    async initialize() {
      await Promise.all([
        fsp.mkdir(path.join(songLibraryRoot, 'songs'), { recursive: true }),
        preferences.ensure(),
        recovery.ensure(),
        history.ensure(),
        songLibraryIndex.ensure()
      ]);
    }
  };
}

module.exports = {
  createPersistenceStores,
  historySchema,
  preferencesSchema,
  recoverySchema,
  songLibraryIndexSchema,
  songLibraryRecordSchema
};
