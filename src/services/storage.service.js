const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const { INSERTOS_DIR, UPLOADS_DIR, OUT_DIR } = require('../utils/paths');

function ensureDirSync(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}
function ensureBaseDirs() {
  ensureDirSync(INSERTOS_DIR);
  ensureDirSync(UPLOADS_DIR);
  ensureDirSync(OUT_DIR);
}

async function listFiles(dir, extFilter = null) {
  const entries = await fsp.readdir(dir, { withFileTypes: true });
  const files = [];
  for (const e of entries) {
    if (!e.isFile()) continue;
    if (extFilter && path.extname(e.name).toLowerCase() !== extFilter) continue;

    const full = path.join(dir, e.name);
    const stat = await fsp.stat(full);
    files.push({ name: e.name, size: stat.size });
  }
  return files;
}

module.exports = {
  ensureBaseDirs,
  listFiles
};
