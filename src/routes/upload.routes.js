const { Router } = require('express');
const multer = require('multer');
const path = require('path');
const { UPLOADS_DIR, INSERTOS_DIR } = require('../utils/paths');
const { handleUploadBase, handleUploadInserto } = require('../controllers/upload.controller');

// Solo .docx
const docxFilter = (_req, file, cb) => {
  if (path.extname(file.originalname).toLowerCase() !== '.docx') {
    return cb(new Error('Solo se permiten archivos .docx'));
  }
  cb(null, true);
};

// Storage para Word base (subido por el usuario)
const storageBase = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOADS_DIR),
  filename: (_req, file, cb) => {
    const stamp = Date.now();
    cb(null, `${stamp}-${file.originalname.replace(/\s+/g, '_')}`);
  }
});

// Storage para “insertos” (biblioteca ya existente, por si quieres cargar nuevos)
const storageInserto = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, INSERTOS_DIR),
  filename: (_req, file, cb) => {
    const base = file.originalname.replace(/\s+/g, '_');
    cb(null, base);
  }
});

const uploadBase = multer({
  storage: storageBase,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: docxFilter
});

const uploadInserto = multer({
  storage: storageInserto,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: docxFilter
});

const router = Router();

// POST /api/upload/base    -> sube el DOCX base del usuario
router.post('/base', uploadBase.single('file'), handleUploadBase);

// POST /api/upload/inserto -> opcional: cargar/actualizar DOCX inserto a la biblioteca
router.post('/inserto', uploadInserto.single('file'), handleUploadInserto);

module.exports = router;
