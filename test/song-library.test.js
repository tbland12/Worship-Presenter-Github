const assert = require('node:assert/strict');
const fsp = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { after, before, test } = require('node:test');
const { createPersistenceStores } = require('../src/main/persistence-stores');
const { createSongLibrary, SongLibraryConflictError } = require('../src/main/song-library');
const { migrateProject } = require('../src/main/project-v2');

let root;
let library;

function songFixture(title, authors = []) {
  const project = migrateProject({
    songs: {
      fixture: {
        id: 'fixture',
        title,
        ccli: { songNumber: '123', authors, publisher: '', copyright: '' },
        background: { type: 'image', path: '' },
        theme: {},
        slides: []
      }
    }
  });
  return project.songs.fixture;
}

before(async () => {
  root = await fsp.mkdtemp(path.join(os.tmpdir(), 'worship-presenter-songs-'));
  const stores = createPersistenceStores(root);
  await stores.initialize();
  library = createSongLibrary(stores);
});

after(async () => {
  await fsp.rm(root, { recursive: true, force: true });
});

test('library saves, searches, and instantiates independent project songs', async () => {
  const saved = await library.save(songFixture('Amazing Grace', ['John Newton']));
  assert.equal(saved.item.revision, 1);
  assert.equal((await library.list('newton'))[0].id, saved.item.id);
  assert.equal((await library.list('123'))[0].id, saved.item.id);

  const first = await library.instantiate(saved.item.id);
  const second = await library.instantiate(saved.item.id);
  assert.notEqual(first.id, second.id);
  assert.deepEqual(first.librarySource, { id: saved.item.id, revision: 1 });
});

test('library updates revisions and rejects stale writes unless forced', async () => {
  const initial = await library.save(songFixture('Revision Song'));
  const edited = { ...initial.song, title: 'Revision Song Updated' };
  const updated = await library.save(edited);
  assert.equal(updated.item.revision, 2);

  await assert.rejects(
    library.save({ ...initial.song, title: 'Stale Edit' }),
    (error) => error instanceof SongLibraryConflictError && error.currentRevision === 2
  );
  const forced = await library.save({ ...initial.song, title: 'Forced Edit' }, { force: true });
  assert.equal(forced.item.revision, 3);
  assert.equal(forced.song.title, 'Forced Edit');
});

test('library supports duplicate titles and deletion by opaque ID', async () => {
  const first = await library.save(songFixture('Duplicate'));
  const second = await library.save(songFixture('Duplicate'));
  assert.notEqual(first.item.id, second.item.id);
  assert.equal((await library.list('duplicate')).length, 2);
  await library.remove(first.item.id);
  assert.deepEqual((await library.list('duplicate')).map((item) => item.id), [second.item.id]);
  await assert.rejects(library.instantiate(first.item.id), /not found/);
});
