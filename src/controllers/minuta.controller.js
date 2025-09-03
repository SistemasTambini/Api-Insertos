const { mergeMinutaInMemory, countPlaceholder } = require('../services/minuta.service');

/**
 * POST /api/minuta
 * form-data:
 *  - minuta:   File (.docx)  → el contenido a insertar
 *  - destino:  File (.docx)  → documento donde se inserta
 *  - placeholder (opcional): string, por defecto "[MINUTA]" (case-insensitive)
 *  - pageBreakBefore / pageBreakAfter (opcional): "true"/"false"
 */
async function insertMinuta(req, res) {
  try {
    const files = req.files || {};
    const minuta  = files?.minuta?.[0];
    const destino = files?.destino?.[0];

    if (!minuta || !destino) {
      return res.status(400).json({ error: 'Envía ambos archivos en form-data: "minuta" y "destino" (.docx).' });
    }

    const placeholder = (req.body?.placeholder || '[MINUTA]').trim();
    const before = ['1','true','on','yes'].includes(String(req.body?.pageBreakBefore || '').toLowerCase());
    const after  = ['1','true','on','yes'].includes(String(req.body?.pageBreakAfter  || '').toLowerCase());

    // Validación marcador en el destino
    const { exact, mixed } = countPlaceholder(destino.buffer, placeholder);
    if (exact === 0 && mixed === 0) {
      return res.status(400).json({ error: `El documento destino no contiene el marcador requerido.`, marker: placeholder, found: false });
    }
    if (mixed > 0) {
      return res.status(400).json({ error: `El marcador ${placeholder} debe estar solo en su propio párrafo (sin texto antes/después).`, marker: placeholder, mixed });
    }

    // Merge real en memoria
    const out = mergeMinutaInMemory(destino.buffer, minuta.buffer, {
      placeholder,
      pageBreakBefore: before,
      pageBreakAfter: after
    });

    // Sanidad: marcador ya no debería existir
    const post = countPlaceholder(out, placeholder);
    if (post.exact > 0) {
      return res.status(500).json({ error: `No se pudo insertar en el marcador ${placeholder}.` });
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="destino_con_minuta_${Date.now()}.docx"`);
    return res.send(out);
  } catch (err) {
    console.error('[minuta] ERROR', err);
    return res.status(500).json({ error: 'No se pudo procesar la minuta' });
  }
}

module.exports = { insertMinuta };
