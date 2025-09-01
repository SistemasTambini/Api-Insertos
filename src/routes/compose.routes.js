const { Router } = require('express');
const { composeDoc } = require('../controllers/compose.controller');

const router = Router();

// POST /compose  { baseFilename: "169...-miword.docx", insertIds: [1,5,12] }
router.post('/', composeDoc);

module.exports = router;
