const { Router } = require('express');
const multer = require('multer');
const path = require('path');
const { insertMinuta } = require('../controllers/minuta.controller');

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
// Se esperan DOS archivos: "minuta" y "destino"
router.post('/', uploadMem.fields([{ name: 'minuta', maxCount: 1 }, { name: 'destino', maxCount: 1 }]), insertMinuta);

module.exports = router;
