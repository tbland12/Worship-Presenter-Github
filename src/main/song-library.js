const crypto = require('node:crypto');
const { songSchema } = require('./project-v2');

const SONG_LIBRARY_ID = /^[a-zA-Z0-9-]{1,128}$/;
const MAX_LIBRARY_RESULTS = 500;

class SongLibraryConflictError extends Error {
  constructor(currentRevision) {
    super('This library song was updated elsewhere.');
    this.name = 'SongLibraryConflictError';
    this.currentRevision = currentRevision;
  }
}

function indexItemFromRecord(record) {
  return {
    id: record.id,
    revision: record.revision,
    title: record.song.title,
    ccliSongNumber: record.song.ccli.songNumber,
    authors: record.song.ccli.authors,
    updatedAt: record.updatedAt
  };
}

function matchesQuery(item, query) {
  if (!query) return true;
  const searchable = [item.title, item.ccliSongNumber, ...(item.authors || [])].join('\n').toLocaleLowerCase();
  return searchable.includes(query.toLocaleLowerCase());
}

function createSongLibrary(stores) {
  if (!stores?.songLibraryIndex || typeof stores.songRecord !== 'function') {
    throw new Error('Song library requires initialized persistence stores.');
  }

  return {
    async list(query = '') {
      const normalizedQuery = String(query || '').trim().slice(0, 256);
      const index = await stores.songLibraryIndex.read();
      return index.items
        .filter((item) => matchesQuery(item, normalizedQuery))
        .sort((first, second) => first.title.localeCompare(second.title) || second.updatedAt.localeCompare(first.updatedAt))
        .slice(0, MAX_LIBRARY_RESULTS);
    },

    async getItem(id) {
      if (!SONG_LIBRARY_ID.test(id)) throw new Error('Song library ID is invalid.');
      const index = await stores.songLibraryIndex.read();
      const item = index.items.find((entry) => entry.id === id);
      if (!item) throw new Error('Library song was not found.');
      return item;
    },

    async save(song, { force = false } = {}) {
      const parsedSong = songSchema.parse(structuredClone(song));
      const source = parsedSong.librarySource;
      let id = source && SONG_LIBRARY_ID.test(source.id) ? source.id : crypto.randomUUID();
      let revision = 1;
      const existingStore = stores.songRecord(id);
      if (source && await existingStore.exists()) {
        const existing = await existingStore.read();
        if (!force && source.revision !== existing.revision) {
          throw new SongLibraryConflictError(existing.revision);
        }
        revision = existing.revision + 1;
      } else if (source) {
        id = crypto.randomUUID();
      }

      const linkedSong = {
        ...parsedSong,
        librarySource: { id, revision }
      };
      const record = {
        version: 1,
        id,
        revision,
        updatedAt: new Date().toISOString(),
        song: linkedSong
      };
      await stores.songRecord(id).replace(record);
      const item = indexItemFromRecord(record);
      await stores.songLibraryIndex.update((index) => ({
        version: 1,
        items: [item, ...index.items.filter((entry) => entry.id !== id)]
      }));
      return { item, song: structuredClone(linkedSong) };
    },

    async instantiate(id) {
      if (!SONG_LIBRARY_ID.test(id)) throw new Error('Song library ID is invalid.');
      const store = stores.songRecord(id);
      if (!await store.exists()) throw new Error('Library song was not found.');
      const record = await store.read();
      return {
        ...structuredClone(record.song),
        id: `song-${crypto.randomUUID()}`,
        librarySource: { id: record.id, revision: record.revision }
      };
    },

    async remove(id) {
      if (!SONG_LIBRARY_ID.test(id)) throw new Error('Song library ID is invalid.');
      const index = await stores.songLibraryIndex.update((current) => ({
        version: 1,
        items: current.items.filter((item) => item.id !== id)
      }));
      await stores.songRecord(id).remove();
      return index.items;
    }
  };
}

module.exports = {
  MAX_LIBRARY_RESULTS,
  SongLibraryConflictError,
  createSongLibrary,
  indexItemFromRecord,
  matchesQuery
};
