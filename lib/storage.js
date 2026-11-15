const fs = require('fs');
const path = require('path');

// Local file storage (used by `npm start`)
function fileStorage(dataDir) {
  const file = path.join(dataDir, 'streaks.xlsx');
  return {
    describe: () => file,
    async load() {
      if (!fs.existsSync(file)) return null;
      return fs.promises.readFile(file);
    },
    async save(buffer) {
      if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
      try {
        await fs.promises.writeFile(file, buffer);
      } catch (e) {
        if (e && (e.code === 'EBUSY' || e.code === 'EPERM')) {
          const err = new Error('The Excel file is open in another program. Please close it and try again.');
          err.status = 423;
          throw err;
        }
        throw e;
      }
    },
  };
}

// Netlify Blobs storage (used when deployed on Netlify)
function blobStorage() {
  const { getStore } = require('@netlify/blobs');
  const KEY = 'streaks.xlsx';
  const store = () => getStore({ name: 'streak-app', consistency: 'strong' });
  return {
    describe: () => 'Netlify Blobs (store: streak-app, key: streaks.xlsx)',
    async load() {
      const ab = await store().get(KEY, { type: 'arrayBuffer' });
      return ab ? Buffer.from(ab) : null;
    },
    async save(buffer) {
      const ab = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
      await store().set(KEY, ab);
    },
  };
}

function pickStorage() {
  const onNetlify = !!(process.env.NETLIFY || process.env.NETLIFY_BLOBS_CONTEXT || process.env.NETLIFY_DEV);
  return onNetlify ? blobStorage() : fileStorage(path.join(__dirname, '..', 'data'));
}

module.exports = { fileStorage, blobStorage, pickStorage };
