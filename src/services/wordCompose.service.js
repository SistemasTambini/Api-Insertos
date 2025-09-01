const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const PizZip = require('pizzip');
const { XMLParser, XMLBuilder } = require('fast-xml-parser');
const { INSERTOS_DIR } = require('../utils/paths');

function findInsertoPathById(id) {
  // Busca un archivo que empiece con "id" (ej: "1.docx", "1 - CODIGO CIVIL.docx")
  const files = fs.readdirSync(INSERTOS_DIR).filter(f => f.toLowerCase().endsWith('.docx'));
  const needle = String(id);
  const match = files.find(f => f.startsWith(needle + '.') || f.startsWith(needle + ' ') || f.startsWith(needle + '-'));
  // Si no encuentra por prefijo, intenta exacto "id.docx"
  if (!match) {
    const exact = files.find(f => f.toLowerCase() === `${needle}.docx`);
    return exact ? path.join(INSERTOS_DIR, exact) : null;
  }
  return path.join(INSERTOS_DIR, match);
}

async function composeWithAltChunk(baseDocxPath, insertIds = []) {
  const baseBuffer = await fsp.readFile(baseDocxPath);
  const zip = new PizZip(baseBuffer);

  // Lee y prepara document.xml
  const docXmlPath = 'word/document.xml';
  const relsPath = 'word/_rels/document.xml.rels';
  const contentTypesPath = '[Content_Types].xml';

  let docXml = zip.file(docXmlPath).asText();
  let relsXml = zip.file(relsPath)?.asText();
  let contentTypesXml = zip.file(contentTypesPath).asText();

  if (!relsXml) {
    // Si el base no tiene rels (raro), crea uno mínimo
    relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  }

  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '', suppressEmptyNode: true });

  const relsObj = parser.parse(relsXml);
  const relsRoot = relsObj.Relationships || { Relationship: [] };
  if (!Array.isArray(relsRoot.Relationship)) relsRoot.Relationship = relsRoot.Relationship ? [relsRoot.Relationship] : [];

  const ctObj = parser.parse(contentTypesXml);
  const overrides = ctObj.Types.Override ? (Array.isArray(ctObj.Types.Override) ? ctObj.Types.Override : [ctObj.Types.Override]) : [];

  // Asegura carpeta para altchunks dentro del ZIP
  const altFolder = 'word/altchunks/';
  // no hace falta crear carpeta explícita en PizZip; basta con usar la ruta al añadir archivos

  // Genera relaciones y altChunks
  let altChunksXml = '';
  let relIdCounter = 9000; // id alto para evitar colisiones
  let chunkCounter = 1;

  for (const id of insertIds) {
    const insertoPath = findInsertoPathById(id);
    if (!insertoPath) continue; // salta faltantes

    const chunkName = `chunk${chunkCounter}.docx`;
    const chunkZipPath = `${altFolder}${chunkName}`;
    const chunkBuffer = await fsp.readFile(insertoPath);
    zip.file(chunkZipPath, chunkBuffer);

    // Relationship para altChunk
    const rId = `rIdAlt${relIdCounter++}`;
    relsRoot.Relationship.push({
      Id: rId,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk',
      Target: `altchunks/${chunkName}`
    });

    // Override para el archivo del chunk en Content_Types
    overrides.push({
      PartName: `/word/altchunks/${chunkName}`,
      ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    // Nodo altChunk que se inserta en document.xml (al final del cuerpo)
    altChunksXml += `\n<w:altChunk r:id="${rId}"/>`;
    chunkCounter++;
  }

  // Si no hay nada que insertar, devuelve el mismo buffer
  if (!altChunksXml) return baseBuffer;

  // Inserta antes de </w:body> (bloque-level)
  if (docXml.includes('</w:body>')) {
    docXml = docXml.replace('</w:body>', `${altChunksXml}\n</w:body>`);
  } else {
    // fallback (muy raro): lo agrega al final
    docXml = docXml.replace('</w:document>', `<w:body>${altChunksXml}</w:body></w:document>`);
  }

  // Re-escribe rels y content types
  relsObj.Relationships = relsRoot;
  const newRelsXml = builder.build(relsObj);

  ctObj.Types.Override = overrides;
  const newCtXml = builder.build(ctObj);

  zip.file(docXmlPath, docXml);
  zip.file(relsPath, newRelsXml);
  zip.file(contentTypesPath, newCtXml);

  // Exporta el ZIP como DOCX
  return zip.generate({ type: 'nodebuffer' });
}

module.exports = { composeWithAltChunk };
