const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const ASSETS_DIR = path.join(ROOT, 'src', 'assets');

const INSERTOS_DIR = path.join(ASSETS_DIR, 'insertos');
const UPLOADS_DIR  = path.join(ASSETS_DIR, 'uploads');
const OUT_DIR      = path.join(ASSETS_DIR, 'out');

module.exports = {
  ROOT,
  ASSETS_DIR,
  INSERTOS_DIR,
  UPLOADS_DIR,
  OUT_DIR
};