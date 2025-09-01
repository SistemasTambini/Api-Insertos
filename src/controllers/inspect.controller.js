const PizZip = require('pizzip');
const { XMLParser } = require('fast-xml-parser');

async function inspectDocx(req, res) {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ error: 'Sube un .docx en el campo "file".' });

    const zip = new PizZip(file.buffer);

    const docXml = zip.file('word/document.xml')?.asText() || null;
    const relsXml = zip.file('word/_rels/document.xml.rels')?.asText() || null;
    const ctXml   = zip.file('[Content_Types].xml')?.asText() || null;

    // 1) ¿Cuántos <w:altChunk .../> hay en document.xml?
    const altChunkCount = docXml ? (docXml.match(/<w:altChunk\b/gi) || []).length : 0;

    // 2) ¿Qué relaciones aFChunk hay en document.xml.rels?
    let rels = [];
    if (relsXml) {
      const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
      const relsObj = parser.parse(relsXml);
      const items = relsObj?.Relationships?.Relationship || [];
      const arr = Array.isArray(items) ? items : [items];
      rels = arr.filter(r => (r?.Type || '').includes('/aFChunk')).map(r => ({
        Id: r.Id, Type: r.Type, Target: r.Target, TargetMode: r.TargetMode || 'Internal'
      }));
    }

    // 3) Overrides de [Content_Types] para /word/altchunks/*
    let overrides = [];
    if (ctXml) {
      const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
      const ctObj = parser.parse(ctXml);
      const ov = ctObj?.Types?.Override || [];
      const arr = Array.isArray(ov) ? ov : [ov];
      overrides = arr
        .filter(o => (o?.PartName || '').startsWith('/word/altchunks/'))
        .map(o => ({ PartName: o.PartName, ContentType: o.ContentType }));
    }

    // 4) Lista real de archivos /word/altchunks/* en el ZIP
    const entries = Object.keys(zip.files)
      .filter(k => k.startsWith('word/altchunks/') && k.toLowerCase().endsWith('.docx'))
      .sort();

    return res.json({
      size: file.size,
      altChunkNodesInDocumentXml: altChunkCount,
      afChunkRels: rels,
      contentTypesOverrides: overrides,
      zipAltchunkEntries: entries
    });
  } catch (err) {
    console.error('[inspect] ERROR', err);
    return res.status(500).json({ error: 'No se pudo inspeccionar el DOCX' });
  }
}

module.exports = { inspectDocx };
