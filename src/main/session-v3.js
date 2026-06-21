const crypto = require('node:crypto');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const { pipeline } = require('node:stream/promises');
const yauzl = require('yauzl');
const yazl = require('yazl');
const { z } = require('zod');
const { migrateProject } = require('./project-v2');

const FORMAT = 'worship-presenter-session';
const VERSION = 3;
const MAX_MANIFEST_BYTES = 2 * 1024 * 1024;
const MAX_ASSETS = 2000;
const MAX_ASSET_BYTES = 4 * 1024 * 1024 * 1024;
const MAX_TOTAL_BYTES = 8 * 1024 * 1024 * 1024;
const HASH_PATTERN = /^[a-f0-9]{64}$/;
const EXT_PATTERN = /^\.(jpg|jpeg|png|gif|bmp|webp|mp4|mov|mkv|avi|webm)$/;

const assetSchema = z.object({
  id: z.string().regex(HASH_PATTERN),
  archivePath: z.string(),
  mediaType: z.enum(['image', 'video']),
  originalName: z.string().max(255),
  extension: z.string().regex(EXT_PATTERN),
  byteLength: z.number().int().nonnegative().max(MAX_ASSET_BYTES),
  sha256: z.string().regex(HASH_PATTERN)
}).strict();

const manifestSchema = z.object({
  format: z.literal(FORMAT),
  sessionVersion: z.literal(VERSION),
  appVersion: z.string().max(64),
  project: z.record(z.string(), z.unknown()),
  assets: z.array(assetSchema).max(MAX_ASSETS)
}).strict();

function hashFile(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);
    stream.on('data', (chunk) => hash.update(chunk));
    stream.on('error', reject);
    stream.on('end', () => resolve(hash.digest('hex')));
  });
}

async function describeAsset(sourcePath, mediaType) {
  const stat = await fsp.stat(sourcePath);
  if (!stat.isFile() || stat.size > MAX_ASSET_BYTES) {
    throw new Error('Session asset is too large or is not a regular file.');
  }
  const extension = path.extname(sourcePath).toLowerCase();
  if (!EXT_PATTERN.test(extension)) {
    throw new Error(`Unsupported session asset extension: ${extension || '(none)'}`);
  }
  const sha256 = await hashFile(sourcePath);
  return {
    id: sha256,
    archivePath: `assets/${sha256}${extension}`,
    mediaType: mediaType === 'video' ? 'video' : 'image',
    originalName: path.basename(sourcePath).slice(0, 255),
    extension,
    byteLength: stat.size,
    sha256,
    sourcePath
  };
}

async function writeSession({ targetPath, appVersion, project, assets }) {
  if (!Array.isArray(assets) || assets.length > MAX_ASSETS) {
    throw new Error('Session contains too many assets.');
  }
  const descriptors = [];
  const seen = new Set();
  let totalBytes = 0;
  for (const asset of assets) {
    const descriptor = await describeAsset(asset.sourcePath, asset.mediaType);
    if ((asset.id && asset.id !== descriptor.id) || (asset.sha256 && asset.sha256 !== descriptor.sha256)) {
      throw new Error('Session asset changed while the session was being prepared.');
    }
    if (seen.has(descriptor.id)) continue;
    seen.add(descriptor.id);
    totalBytes += descriptor.byteLength;
    if (totalBytes > MAX_TOTAL_BYTES) throw new Error('Session assets exceed the aggregate size limit.');
    descriptors.push(descriptor);
  }
  const migratedProject = migrateProject(project);
  const manifest = {
    format: FORMAT,
    sessionVersion: VERSION,
    appVersion,
    project: migratedProject,
    assets: descriptors.map(({ sourcePath: _sourcePath, ...descriptor }) => descriptor)
  };
  manifestSchema.parse(manifest);
  const manifestData = Buffer.from(JSON.stringify(manifest));
  if (manifestData.length > MAX_MANIFEST_BYTES) throw new Error('Session manifest is too large.');
  const nonce = crypto.randomUUID();
  const temporaryPath = `${targetPath}.${process.pid}.${nonce}.tmp`;
  const backupPath = `${targetPath}.${process.pid}.${nonce}.bak`;
  const zip = new yazl.ZipFile();
  zip.addBuffer(manifestData, 'manifest.json', { compress: true });
  descriptors.forEach((asset) => zip.addFile(asset.sourcePath, asset.archivePath, { compress: false }));
  zip.end({ forceZip64Format: totalBytes > 0xffffffff });
  try {
    await pipeline(zip.outputStream, fs.createWriteStream(temporaryPath, { flags: 'wx' }));
    const targetExists = await fsp.stat(targetPath).then(() => true, () => false);
    if (targetExists) await fsp.rename(targetPath, backupPath);
    try {
      await fsp.rename(temporaryPath, targetPath);
      if (targetExists) await fsp.rm(backupPath, { force: true });
    } catch (error) {
      if (targetExists) await fsp.rename(backupPath, targetPath).catch(() => {});
      throw error;
    }
  } catch (error) {
    await fsp.rm(temporaryPath, { force: true }).catch(() => {});
    throw error;
  }
  return { targetPath, assets: descriptors };
}

function openZip(filePath) {
  return new Promise((resolve, reject) => {
    yauzl.open(filePath, { autoClose: false, lazyEntries: true, validateEntrySizes: true }, (error, zip) => {
      if (error) reject(new Error('Unsupported session format. Worship Presenter requires a v3 session.'));
      else resolve(zip);
    });
  });
}

function readEntry(zip, entry, limit) {
  return new Promise((resolve, reject) => {
    if (entry.uncompressedSize > limit) return reject(new Error('Session entry exceeds its size limit.'));
    zip.openReadStream(entry, (error, stream) => {
      if (error) return reject(error);
      const chunks = [];
      let size = 0;
      stream.on('data', (chunk) => {
        size += chunk.length;
        if (size > limit) stream.destroy(new Error('Session entry exceeds its size limit.'));
        else chunks.push(chunk);
      });
      stream.on('error', reject);
      stream.on('end', () => resolve(Buffer.concat(chunks)));
    });
  });
}

async function extractAsset(zip, entry, descriptor, cacheDir) {
  const finalPath = path.join(cacheDir, `${descriptor.id}${descriptor.extension}`);
  const existing = await fsp.stat(finalPath).catch(() => null);
  if (existing && existing.size === descriptor.byteLength && await hashFile(finalPath) === descriptor.sha256) return finalPath;
  const temporaryPath = `${finalPath}.${process.pid}.${crypto.randomUUID()}.tmp`;
  await fsp.mkdir(cacheDir, { recursive: true });
  const hash = crypto.createHash('sha256');
  let size = 0;
  try {
    await new Promise((resolve, reject) => {
      zip.openReadStream(entry, async (error, stream) => {
        if (error) return reject(error);
        const output = fs.createWriteStream(temporaryPath, { flags: 'wx' });
        stream.on('data', (chunk) => { size += chunk.length; hash.update(chunk); });
        try { await pipeline(stream, output); resolve(); } catch (streamError) { reject(streamError); }
      });
    });
    if (size !== descriptor.byteLength || hash.digest('hex') !== descriptor.sha256) {
      throw new Error('Session asset failed integrity validation.');
    }
    await fsp.rm(finalPath, { force: true });
    await fsp.rename(temporaryPath, finalPath);
    return finalPath;
  } catch (error) {
    await fsp.rm(temporaryPath, { force: true }).catch(() => {});
    throw error;
  }
}

async function readSession({ filePath, cacheDir }) {
  const zip = await openZip(filePath);
  const entries = new Map();
  let manifest = null;
  try {
    await new Promise((resolve, reject) => {
      zip.on('entry', async (entry) => {
        try {
          if (entries.has(entry.fileName) || entry.fileName.includes('..') || path.posix.isAbsolute(entry.fileName)) {
            throw new Error('Session contains an invalid archive entry.');
          }
          entries.set(entry.fileName, entry);
          if (entry.fileName === 'manifest.json') {
            const data = await readEntry(zip, entry, MAX_MANIFEST_BYTES);
            manifest = manifestSchema.parse(JSON.parse(data.toString('utf8')));
          }
          zip.readEntry();
        } catch (error) { reject(error); }
      });
      zip.on('end', resolve);
      zip.on('error', reject);
      zip.readEntry();
    });
    if (!manifest) throw new Error('Session manifest is missing.');
    let totalBytes = 0;
    const assetPaths = new Map();
    for (const descriptor of manifest.assets) {
      if (assetPaths.has(descriptor.id)) throw new Error('Session contains duplicate asset identifiers.');
      const expectedPath = `assets/${descriptor.id}${descriptor.extension}`;
      if (descriptor.archivePath !== expectedPath) throw new Error('Session asset path is invalid.');
      const entry = entries.get(expectedPath);
      if (!entry || entry.uncompressedSize !== descriptor.byteLength) throw new Error('Session asset is missing or has an invalid size.');
      totalBytes += descriptor.byteLength;
      if (totalBytes > MAX_TOTAL_BYTES) throw new Error('Session assets exceed the aggregate size limit.');
      assetPaths.set(descriptor.id, await extractAsset(zip, entry, descriptor, cacheDir));
    }
    const allowedEntries = new Set(['manifest.json', ...manifest.assets.map((asset) => asset.archivePath)]);
    if ([...entries.keys()].some((entry) => !allowedEntries.has(entry))) throw new Error('Session contains unexpected archive entries.');
    return { project: migrateProject(manifest.project), assetPaths };
  } finally {
    zip.close();
  }
}

module.exports = { FORMAT, VERSION, describeAsset, readSession, writeSession };
