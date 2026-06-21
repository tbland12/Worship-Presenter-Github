const assert = require('node:assert/strict');
const path = require('node:path');
const test = require('node:test');
const {
  MAX_RECENT_PROJECTS,
  buildHistoryEntry,
  publicHistoryEntries,
  removeHistoryEntry,
  upsertHistory
} = require('../src/main/project-history');

test('history entries summarize project content without exposing paths publicly', () => {
  const project = {
    songs: {
      first: {
        id: 'first',
        title: 'First Song',
        ccli: { songNumber: '123' },
        slides: [{}, {}]
      }
    },
    announcements: { slides: [{}] },
    timer: { slides: [{}, {}] }
  };
  const entry = buildHistoryEntry(path.join('C:\\', 'Services', 'Sunday.wpjson'), project, '2026-06-20T12:00:00.000Z');
  assert.equal(entry.title, 'Sunday');
  assert.equal(entry.slideCount, 5);
  assert.deepEqual(entry.songs, [{ id: 'first', title: 'First Song', ccliSongNumber: '123' }]);
  const [publicEntry] = publicHistoryEntries({ version: 1, entries: [entry] });
  assert.equal(Object.hasOwn(publicEntry, 'filePath'), false);
});

test('history upserts move an existing project to the front and enforce the limit', () => {
  let history = { version: 1, entries: [] };
  for (let index = 0; index < MAX_RECENT_PROJECTS + 2; index += 1) {
    const filePath = path.resolve(`service-${index}.wpjson`);
    history = upsertHistory(history, buildHistoryEntry(filePath, {}, new Date(index).toISOString()));
  }
  assert.equal(history.entries.length, MAX_RECENT_PROJECTS);
  const existing = { ...history.entries[10], title: 'Updated' };
  history = upsertHistory(history, existing);
  assert.equal(history.entries[0].title, 'Updated');
  assert.equal(history.entries.filter((entry) => entry.id === existing.id).length, 1);
  assert.deepEqual(removeHistoryEntry(history, existing.id).entries.includes(existing), false);
});
