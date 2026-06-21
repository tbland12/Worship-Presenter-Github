const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { after, before, test } = require('node:test');
const yazl = require('yazl');
const { pipeline } = require('node:stream/promises');
const { describeAsset, readSession, writeSession } = require('../src/main/session-v3');
const { migrateProject } = require('../src/main/project-v2');

let temporaryRoot;

before(async () => {
  temporaryRoot = await fsp.mkdtemp(path.join(os.tmpdir(), 'worship-presenter-test-'));
});

after(async () => {
  await fsp.rm(temporaryRoot, { recursive: true, force: true });
});

test('v3 sessions stream assets and verify their content hash', async () => {
  const sourcePath = path.join(temporaryRoot, 'background.png');
  const sessionPath = path.join(temporaryRoot, 'roundtrip.wpjson');
  await fsp.writeFile(sourcePath, Buffer.from('test-image-content'));
  const descriptor = await describeAsset(sourcePath, 'image');
  const project = { songs: { first: { background: { type: 'image', path: `session/${descriptor.id}` } } } };
  await writeSession({
    targetPath: sessionPath,
    appVersion: 'test',
    project,
    assets: [{ ...descriptor, sourcePath }]
  });
  const opened = await readSession({ filePath: sessionPath, cacheDir: path.join(temporaryRoot, 'cache') });
  assert.deepEqual(opened.project, migrateProject(project));
  assert.equal(opened.project.schemaVersion, 2);
  assert.equal(await fsp.readFile(opened.assetPaths.get(descriptor.id), 'utf8'), 'test-image-content');
});

test('saving replaces an existing session without leaving a backup', async () => {
  const sourcePath = path.join(temporaryRoot, 'replacement.png');
  const sessionPath = path.join(temporaryRoot, 'replacement.wpjson');
  await fsp.writeFile(sourcePath, Buffer.from('replacement-image'));
  const descriptor = await describeAsset(sourcePath, 'image');
  const payload = { targetPath: sessionPath, appVersion: 'test', project: {}, assets: [descriptor] };
  await writeSession(payload);
  await writeSession(payload);
  const opened = await readSession({ filePath: sessionPath, cacheDir: path.join(temporaryRoot, 'replacement-cache') });
  assert.equal(opened.assetPaths.size, 1);
  assert.equal((await fsp.readdir(temporaryRoot)).some((name) => name.endsWith('.bak')), false);
});

test('saving rejects an asset changed after it was described', async () => {
  const sourcePath = path.join(temporaryRoot, 'changed.png');
  const sessionPath = path.join(temporaryRoot, 'changed.wpjson');
  await fsp.writeFile(sourcePath, Buffer.from('first-content'));
  const descriptor = await describeAsset(sourcePath, 'image');
  await fsp.writeFile(sourcePath, Buffer.from('changed-content'));
  await assert.rejects(
    writeSession({ targetPath: sessionPath, appVersion: 'test', project: {}, assets: [descriptor] }),
    /changed while the session was being prepared/
  );
});

test('legacy JSON sessions are rejected', async () => {
  const sessionPath = path.join(temporaryRoot, 'legacy.wpjson');
  await fsp.writeFile(sessionPath, JSON.stringify({ sessionVersion: 2, project: {} }));
  await assert.rejects(
    readSession({ filePath: sessionPath, cacheDir: path.join(temporaryRoot, 'legacy-cache') }),
    /requires a v3 session/
  );
});

test('archives containing undeclared entries are rejected', async () => {
  const sessionPath = path.join(temporaryRoot, 'unexpected.wpjson');
  const manifest = Buffer.from(JSON.stringify({
    format: 'worship-presenter-session',
    sessionVersion: 3,
    appVersion: 'test',
    project: {},
    assets: []
  }));
  const zip = new yazl.ZipFile();
  zip.addBuffer(manifest, 'manifest.json');
  zip.addBuffer(Buffer.from('unexpected'), 'unexpected.txt');
  zip.end();
  await pipeline(zip.outputStream, fs.createWriteStream(sessionPath));
  await assert.rejects(
    readSession({ filePath: sessionPath, cacheDir: path.join(temporaryRoot, 'unexpected-cache') }),
    /unexpected archive entries/
  );
});
