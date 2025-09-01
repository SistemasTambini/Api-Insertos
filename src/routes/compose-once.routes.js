const { Router } = require('express');
const multer = require('multer');
const path = require('path');
const { composeOnce } = require('../controllers/composeOnce.controller');

const uploadMem = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
  fileFilter: (_req, file, cb) => {
    if (path.extname(file.originalname).toLowerCase() !== '.docx') {
      return cb(new Error('Solo .docx'));
    }
    cb(null, true);
  }
});

const router = Router();

// POST /api/compose-once  (multipart: file + insertIds)
router.post('/', uploadMem.single('file'), composeOnce);

module.exports = router;
