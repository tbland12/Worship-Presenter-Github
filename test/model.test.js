const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

async function loadModel() {
  const source = fs.readFileSync(path.resolve(__dirname, '../src/shared/model.js'), 'utf8');
  return import(`data:text/javascript;base64,${Buffer.from(source).toString('base64')}`);
}

test('rich text round-trips plain text', async () => {
  const { plainFromRichText, richTextFromPlain } = await loadModel();
  assert.equal(plainFromRichText(richTextFromPlain('Line one\nLine two')), 'Line one\nLine two');
});

test('new songs do not share nested theme state', async () => {
  const { createSong } = await loadModel();
  const first = createSong('First');
  const second = createSong('Second');
  first.theme.textStyles.lyrics.shadow.blur = 99;
  assert.equal(second.theme.textStyles.lyrics.shadow.blur, 6);
});

test('new model records include v2 persistence fields', async () => {
  const { createMediaSlide, createNewProject, createSlide, createSong } = await loadModel();
  assert.equal(createNewProject().schemaVersion, 2);
  assert.equal(createSong('Song').librarySource, null);
  assert.equal(createSlide({ speakerNotes: 'Introduce the chorus' }).speakerNotes, 'Introduce the chorus');
  assert.equal(createMediaSlide({ speakerNotes: 'Wait for the video' }).speakerNotes, 'Wait for the video');
});
