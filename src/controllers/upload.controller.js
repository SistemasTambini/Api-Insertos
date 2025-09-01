// Maneja la subida del Word base (y opcionalmente de un inserto nuevo)

function handleUploadBase(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No se recibió archivo' });
  }
  const { originalname, filename, size } = req.file;
  return res.json({
    ok: true,
    type: 'base',
    file: { originalname, filename, size }
  });
}

function handleUploadInserto(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No se recibió archivo' });
  }
  const { originalname, filename, size } = req.file;
  return res.json({
    ok: true,
    type: 'inserto',
    file: { originalname, filename, size }
  });
}

module.exports = { handleUploadBase, handleUploadInserto };
