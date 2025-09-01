// Lista los DOCX que ya están en src/assets/insertos
const { listFiles } = require('../services/storage.service');
const { INSERTOS_DIR } = require('../utils/paths');

// Detecta prefijo numérico al inicio (1..4 dígitos) seguido de separador o fin
const ID_PREFIX = /^(\d{1,4})(?=[\s._-]|$)/;

async function listInsertos(_req, res) {
  try {
    const files = await listFiles(INSERTOS_DIR, '.docx');

    const items = files.map(f => {
      const m = f.name.match(ID_PREFIX);
      return {
        id: m ? Number(m[1]) : null,
        filename: f.name,
        size: f.size
      };
    });

    // Orden: primero con id (asc), luego sin id (alfabético)
    items.sort((a, b) => {
      const aHas = a.id !== null, bHas = b.id !== null;
      if (aHas && bHas) return a.id - b.id;
      if (aHas && !bHas) return -1;
      if (!aHas && bHas) return 1;
      return a.filename.localeCompare(b.filename, 'es', { numeric: true, sensitivity: 'base' });
    });

    return res.json({ count: items.length, items });
  } catch (err) {
    console.error('[insertos:list] ', err);
    return res.status(500).json({ error: 'No se pudo listar insertos' });
  }
}

module.exports = { listInsertos };
