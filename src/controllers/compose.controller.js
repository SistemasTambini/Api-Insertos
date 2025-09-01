const path = require('path');
const fs = require('fs/promises');
const { UPLOADS_DIR, OUT_DIR } = require('../utils/paths');
const { composeWithAltChunk } = require('../services/wordCompose.service');

async function composeDoc(req, res) {
  try {
    const { baseFilename, insertIds } = req.body || {};
    if (!baseFilename || !Array.isArray(insertIds) || insertIds.length === 0) {
      return res.status(400).json({ error: 'Envía baseFilename y insertIds (array no vacío).' });
    }

    const basePath = path.join(UPLOADS_DIR, baseFilename);
    try { await fs.access(basePath); } catch {
      return res.status(404).json({ error: 'Archivo base no encontrado en uploads.' });
    }

    const buffer = await composeWithAltChunk(basePath, insertIds.map(Number).filter(n => !isNaN(n)));

    const outName = `compuesto-${Date.now()}.docx`;
    const outPath = path.join(OUT_DIR, outName);
    await fs.writeFile(outPath, buffer);

    // Descarga directa
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    return res.send(buffer);
  } catch (err) {
    console.error('[compose] Error:', err);
    return res.status(500).json({ error: 'No se pudo componer el documento' });
  }
}

module.exports = { composeDoc };
