const { composeInMemory, countInsertosPlaceholders } = require('../services/composeOnce.service');

function parseIds(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw.flatMap(parseIds);
  if (typeof raw === 'string') {
    const s = raw.trim();
    if (!s) return [];
    if (s.startsWith('[')) { try { return JSON.parse(s); } catch {} }
    return s.split(/[,\s]+/).map(Number).filter(n => !Number.isNaN(n));
  }
  return [];
}

/**
 * POST /api/compose-once
 * form-data:
 *  - file: DOCX base (File)
 *  - insertIds: "1,3,5"  o  "[1,3,5]"  (también acepta "ids" o "insertosIds")
 */
async function composeOnce(req, res) {
  try {
    const file = req.file;
    const body = req.body || {};
    const ids = parseIds(body.insertIds ?? body.ids ?? body.insertosIds);

    console.log('[compose-once] → entrada', {
      file: file ? { name: file.originalname, size: file.size } : null,
      ids
    });

    if (!file) {
      console.warn('[compose-once] faltó file');
      return res.status(400).json({ error: 'Falta el archivo .docx en el campo "file".' });
    }
    if (!ids.length) {
      console.warn('[compose-once] faltaron insertIds');
      return res.status(400).json({ error: 'Envía al menos un ID en insertIds (ej. 1,3,5 o [1,3,5]).' });
    }

    // Validación marcador
    const { exact, mixed } = await countInsertosPlaceholders(file.buffer);
    console.log('[compose-once] marcador [INSERTOS]', { exact, mixed });

    if (exact === 0 && mixed === 0) {
      return res.status(400).json({
        error: 'El documento no contiene el marcador requerido.',
        marker: '[INSERTOS]',
        found: false
      });
    }
    if (mixed > 0) {
      return res.status(400).json({
        error: 'El marcador [INSERTOS] debe estar solo en su propio párrafo (sin texto antes/después).',
        marker: '[INSERTOS]',
        mixed
      });
    }

    const outBuffer = await composeInMemory(file.buffer, ids);

    // Post-check
    const post = await countInsertosPlaceholders(outBuffer);
    console.log('[compose-once] post-check', post);

    if (post.exact > 0) {
      console.error('[compose-once] aún hay placeholders exactos tras componer');
      return res.status(500).json({ error: 'No se pudo insertar en los marcadores [INSERTOS].' });
    }

    res.setHeader('X-Insertos-Processed', ids.length.toString());
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="compuesto_${Date.now()}.docx"`);
    return res.send(outBuffer);
  } catch (err) {
    console.error('[compose-once] ERROR ', err);
    return res.status(500).json({ error: 'No se pudo componer el documento' });
  }
}

module.exports = { composeOnce };
