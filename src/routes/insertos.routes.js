const { Router } = require('express');
const { listInsertos } = require('../controllers/insertos.controller');

const router = Router();

// GET /api/insertos -> lista los DOCX en src/assets/insertos
router.get('/', listInsertos);

module.exports = router;
