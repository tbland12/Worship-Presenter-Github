const assert = require('node:assert/strict');
const fsp = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { afterEach, test } = require('node:test');
const { z } = require('zod');
const { AtomicJsonStore } = require('../src/main/atomic-json-store');
const { createPersistenceStores } = require('../src/main/persistence-stores');
const { migrateProject } = require('../src/main/project-v2');

const temporaryRoots = [];

async function temporaryRoot() {
  const root = await fsp.mkdtemp(path.join(os.tmpdir(), 'worship-presenter-store-'));
  temporaryRoots.push(root);
  return root;
}

function createStore(filePath, maxBytes = 1024) {
  return new AtomicJsonStore({
    filePath,
    schema: z.object({ version: z.literal(1), count: z.number().int().nonnegative() }).strict(),
    defaultValue: { version: 1, count: 0 },
    maxBytes
  });
}

afterEach(async () => {
  await Promise.all(temporaryRoots.splice(0).map((root) => fsp.rm(root, { recursive: true, force: true })));
});

test('atomic stores initialize, replace, update, and remove values', async () => {
  const root = await temporaryRoot();
  const store = createStore(path.join(root, 'state', 'example.json'));

  assert.deepEqual(await store.ensure(), { version: 1, count: 0 });
  assert.equal(await store.exists(), true);
  await store.replace({ version: 1, count: 4 });
  assert.deepEqual(await store.update((value) => ({ ...value, count: value.count + 1 })), {
    version: 1,
    count: 5
  });
  assert.deepEqual(await store.read(), { version: 1, count: 5 });
  await store.remove();
  assert.equal(await store.exists(), false);
});

test('invalid stores are quarantined and replaced with validated defaults', async () => {
  const root = await temporaryRoot();
  const filePath = path.join(root, 'state.json');
  const store = createStore(filePath);
  await fsp.writeFile(filePath, '{not-json');

  assert.deepEqual(await store.ensure(), { version: 1, count: 0 });
  assert.deepEqual(JSON.parse(await fsp.readFile(filePath, 'utf8')), { version: 1, count: 0 });
  assert.equal((await fsp.readdir(root)).some((name) => name.startsWith('state.json.corrupt-')), true);
});

test('store byte limits reject writes and quarantine oversized files', async () => {
  const root = await temporaryRoot();
  const filePath = path.join(root, 'limited.json');
  const store = createStore(filePath, 40);

  await assert.rejects(store.replace({ version: 1, count: 100000000000 }), /byte limit/);
  await fsp.writeFile(filePath, Buffer.alloc(41, 32));
  assert.deepEqual(await store.read(), { version: 1, count: 0 });
  assert.equal(await store.exists(), false);
});

test('persistence services initialize versioned stores and round-trip song records', async () => {
  const root = await temporaryRoot();
  const stores = createPersistenceStores(root);
  await stores.initialize();

  assert.deepEqual(await stores.preferences.read(), { version: 1, stageDisplayTarget: null, shortcuts: {} });
  await stores.preferences.update((preferences) => ({ ...preferences, stageDisplayTarget: 'display-2' }));
  const reopenedStores = createPersistenceStores(root);
  assert.equal((await reopenedStores.preferences.read()).stageDisplayTarget, 'display-2');
  assert.equal(await stores.recovery.read(), null);
  assert.deepEqual(await stores.history.read(), { version: 1, entries: [] });
  assert.deepEqual(await stores.songLibraryIndex.read(), { version: 1, items: [] });

  const recovery = {
    version: 1,
    updatedAt: new Date().toISOString(),
    sourceSessionPath: 'C:\\Services\\Sunday.wpjson',
    project: migrateProject({})
  };
  await stores.recovery.replace(recovery);
  assert.deepEqual(await stores.recovery.read(), recovery);

  const songStore = stores.songRecord('amazing-grace');
  const record = await songStore.ensure();
  record.song.title = 'Amazing Grace';
  record.updatedAt = new Date().toISOString();
  await songStore.replace(record);
  assert.equal((await songStore.read()).song.title, 'Amazing Grace');
  assert.throws(() => stores.songRecord('../escape'), /invalid/);
});
