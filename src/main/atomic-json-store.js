const crypto = require('node:crypto');
const fsp = require('node:fs/promises');
const path = require('node:path');

function clone(value) {
  return structuredClone(value);
}

class AtomicJsonStore {
  constructor({ filePath, schema, defaultValue, maxBytes }) {
    if (!filePath || !schema || !Number.isInteger(maxBytes) || maxBytes < 1) {
      throw new Error('AtomicJsonStore requires a path, schema, and positive byte limit.');
    }
    this.filePath = filePath;
    this.schema = schema;
    this.defaultValue = schema.parse(clone(defaultValue));
    this.maxBytes = maxBytes;
  }

  async exists() {
    return fsp.stat(this.filePath).then((stat) => stat.isFile(), () => false);
  }

  async quarantine() {
    const quarantinePath = `${this.filePath}.corrupt-${Date.now()}-${crypto.randomUUID()}`;
    await fsp.rename(this.filePath, quarantinePath).catch(() => {});
    return quarantinePath;
  }

  async read() {
    let data;
    try {
      data = await fsp.readFile(this.filePath);
    } catch (error) {
      if (error.code === 'ENOENT') return clone(this.defaultValue);
      throw error;
    }
    if (data.length > this.maxBytes) {
      await this.quarantine();
      return clone(this.defaultValue);
    }
    try {
      return this.schema.parse(JSON.parse(data.toString('utf8')));
    } catch (error) {
      await this.quarantine();
      return clone(this.defaultValue);
    }
  }

  async replace(value) {
    const parsed = this.schema.parse(clone(value));
    const data = Buffer.from(`${JSON.stringify(parsed, null, 2)}\n`);
    if (data.length > this.maxBytes) {
      throw new Error(`Store exceeds its ${this.maxBytes}-byte limit.`);
    }

    await fsp.mkdir(path.dirname(this.filePath), { recursive: true });
    const nonce = `${process.pid}-${crypto.randomUUID()}`;
    const temporaryPath = `${this.filePath}.${nonce}.tmp`;
    const backupPath = `${this.filePath}.${nonce}.bak`;
    let backedUp = false;
    try {
      const handle = await fsp.open(temporaryPath, 'wx');
      try {
        await handle.writeFile(data);
        await handle.sync();
      } finally {
        await handle.close();
      }
      if (await this.exists()) {
        await fsp.rename(this.filePath, backupPath);
        backedUp = true;
      }
      try {
        await fsp.rename(temporaryPath, this.filePath);
      } catch (error) {
        if (backedUp) await fsp.rename(backupPath, this.filePath).catch(() => {});
        throw error;
      }
      if (backedUp) await fsp.rm(backupPath, { force: true }).catch(() => {});
      return clone(parsed);
    } catch (error) {
      await fsp.rm(temporaryPath, { force: true }).catch(() => {});
      throw error;
    }
  }

  async update(updater) {
    if (typeof updater !== 'function') throw new Error('Store updater must be a function.');
    const current = await this.read();
    const next = await updater(clone(current));
    return this.replace(next);
  }

  async ensure() {
    if (!await this.exists()) return this.replace(this.defaultValue);
    const value = await this.read();
    if (!await this.exists()) await this.replace(value);
    return value;
  }

  async remove() {
    await fsp.rm(this.filePath, { force: true });
  }
}

module.exports = { AtomicJsonStore };
