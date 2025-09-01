const { Router } = require('express');
const multer = require('multer');
const path = require('path');
const { inspectDocx } = require('../controllers/inspect.controller');

const uploadMem = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (path.extname(file.originalname).toLowerCase() !== '.docx') {
      return cb(new Error('Solo .docx'));
    }
    cb(null, true);
  }
});

const router = Router();
router.post('/', uploadMem.single('file'), inspectDocx);

module.exports = router;
