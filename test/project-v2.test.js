const assert = require('node:assert/strict');
const test = require('node:test');
const { migrateProject, validateProject } = require('../src/main/project-v2');

function richText(text) {
  return { blocks: [{ paragraphStyle: { align: 'center' }, runs: [{ text }] }] };
}

function legacyProject() {
  return {
    schemaVersion: 1,
    appVersion: '0.1.3',
    settings: { safeMarginsPct: 7 },
    announcements: {
      slides: [{ id: 'announcement-1', label: 'Welcome', mediaPath: 'welcome.png', mediaType: 'image' }],
      autoAdvanceSec: 10,
      loop: true,
      autoAdvanceEnabled: false
    },
    timer: {
      slides: [{ id: 'timer-1', label: 'Countdown', mediaPath: 'timer.mp4', mediaType: 'video' }],
      autoAdvanceOnVideoEnd: true,
      autoAdvanceImages: false,
      autoAdvanceSec: 20
    },
    setlist: ['song-1'],
    songs: {
      'song-1': {
        id: 'song-1',
        title: 'Example Song',
        ccli: { songNumber: '123', authors: ['Writer'], publisher: '', copyright: '' },
        background: { type: 'image', path: 'background.png' },
        theme: { fontFamily: 'Segoe UI' },
        slides: [{
          id: 'slide-1',
          label: 'Verse 1',
          template: 'TitleLyricsFooter',
          footerAutoCcli: true,
          titleText: richText('Example Song'),
          lyricsText: richText('First line'),
          footerText: richText('')
        }]
      }
    }
  };
}

test('v1 projects migrate all slide types to schema v2 without losing settings', () => {
  const migrated = migrateProject(legacyProject());
  const slide = migrated.songs['song-1'].slides[0];

  assert.equal(migrated.schemaVersion, 2);
  assert.equal(migrated.settings.safeMarginsPct, 7);
  assert.equal(migrated.announcements.autoAdvanceEnabled, false);
  assert.equal(migrated.announcements.slides[0].speakerNotes, '');
  assert.equal(migrated.timer.slides[0].speakerNotes, '');
  assert.equal(slide.speakerNotes, '');
  assert.equal(slide.showTitle, true);
  assert.equal(slide.showLyrics, true);
  assert.equal(slide.showFooter, false);
  assert.equal(migrated.songs['song-1'].librarySource, null);
});

test('project migration is idempotent', () => {
  const migrated = migrateProject(legacyProject());
  assert.deepEqual(migrateProject(migrated), migrated);
});

test('future schemas and dangling setlist entries are rejected', () => {
  assert.throws(() => migrateProject({ schemaVersion: 3 }), /Unsupported project schema version/);
  const reservedIdProject = JSON.parse('{"schemaVersion":1,"songs":{"__proto__":{}}}');
  assert.throws(() => migrateProject(reservedIdProject), /reserved song ID/);
  const project = migrateProject(legacyProject());
  project.setlist.push('missing-song');
  assert.throws(() => validateProject(project), /Setlist references missing song/);
});
