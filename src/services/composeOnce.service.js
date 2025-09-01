const fsp = require('fs/promises');
const path = require('path');
const PizZip = require('pizzip');
const { XMLParser, XMLBuilder } = require('fast-xml-parser');
const { INSERTOS_DIR } = require('../utils/paths');

// Cache simple en memoria para buffers de insertos
const insertoCache = new Map();

/** Busca y carga el buffer del DOCX inserto segun ID (1 - Titulo.docx, 1.docx, 1 - ... ) */
async function getInsertoBufferById(id) {
  if (insertoCache.has(id)) return insertoCache.get(id);

  const files = await fsp.readdir(INSERTOS_DIR);
  const needle = String(id);
  const match = files.find(f =>
    f.toLowerCase().endsWith('.docx') &&
    (f.startsWith(`${needle}.`) || f.startsWith(`${needle} `) || f.startsWith(`${needle}-`) || f.toLowerCase() === `${needle}.docx`)
  );
  if (!match) {
    console.warn(`[merge] inserto id=${id} NO encontrado en ${INSERTOS_DIR}`);
    return null;
  }
  const buf = await fsp.readFile(path.join(INSERTOS_DIR, match));
  console.log(`[merge] inserto id=${id} -> ${match} (${buf.length} bytes)`);
  insertoCache.set(id, buf);
  return buf;
}

/** Concatena texto de un párrafo w:p (estructura preserveOrder) */
function getParagraphText(pNode) {
  let text = '';
  function walk(node) {
    if (!node || typeof node !== 'object') return;
    for (const k of Object.keys(node)) {
      if (k === ':@') continue;
      const v = node[k];
      if (k === 'w:t') {
        if (Array.isArray(v)) {
          for (const it of v) {
            if (typeof it === 'string') text += it;
            else if (it && typeof it === 'object' && '#text' in it) text += String(it['#text'] ?? '');
          }
        } else if (typeof v === 'string') text += v;
        else if (v && typeof v === 'object' && '#text' in v) text += String(v['#text'] ?? '');
      } else if (Array.isArray(v)) v.forEach(walk);
      else if (typeof v === 'object') walk(v);
    }
  }
  walk(pNode);
  return text;
}

/** Analiza el body del document.xml y devuelve indices exactos del placeholder */
function analyzeBodyForInsertos(docXml) {
  const parserPO = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true });
  const docObj = parserPO.parse(docXml);

  const findNode = (arr, tag) => Array.isArray(arr) ? arr.find(n => n[tag]) : undefined;
  const docNode = findNode(docObj, 'w:document');
  if (!docNode) return { ok: false };

  const bodyNode = findNode(docNode['w:document'], 'w:body');
  if (!bodyNode) return { ok: false };

  const bodyChildren = bodyNode['w:body'] || [];
  const PLACEHOLDER = '[INSERTOS]';
  const exactParasIdx = [];
  let mixed = 0;

  bodyChildren.forEach((child, idx) => {
    if (child['w:p']) {
      const t = getParagraphText(child);
      if (t.includes(PLACEHOLDER)) {
        if (t.trim() === PLACEHOLDER) exactParasIdx.push(idx);
        else mixed++;
      }
    }
  });

  return { ok: true, docObj, bodyNode, bodyChildren, exactParasIdx, mixed };
}

/** Cuenta placeholders exactos y mixtos */
async function countInsertosPlaceholders(baseBuffer) {
  const zip = new PizZip(baseBuffer);
  const docXml = zip.file('word/document.xml')?.asText();
  if (!docXml) return { exact: 0, mixed: 0 };
  const a = analyzeBodyForInsertos(docXml);
  if (!a.ok) return { exact: 0, mixed: 0 };
  return { exact: a.exactParasIdx.length, mixed: a.mixed };
}

/** Extrae nodos (párrafos/tablas) del body de un DOCX inserto (sin sectPr) */
function extractBodyNodesFromInserto(insertoBuffer) {
  const inz = new PizZip(insertoBuffer);
  const docXml = inz.file('word/document.xml')?.asText();
  if (!docXml) return [];
  const parserPO = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true });
  const docObj = parserPO.parse(docXml);

  const findNode = (arr, tag) => Array.isArray(arr) ? arr.find(n => n[tag]) : undefined;
  const docNode = findNode(docObj, 'w:document');
  if (!docNode) return [];
  const bodyNode = findNode(docNode['w:document'], 'w:body');
  if (!bodyNode) return [];

  const children = bodyNode['w:body'] || [];
  // filtra sectPr (deja w:p, w:tbl, etc.)
  return children.filter(n => !n['w:sectPr']);
}

/** Crea un párrafo de salto de página opcional entre insertos */
function makePageBreakParagraph() {
  return {
    'w:p': [
      { 'w:r': [ { 'w:br': { ':@': { 'w:type': 'page' } } } ] }
    ]
  };
}

/**
 * MERGE REAL EN MEMORIA:
 *  - Reemplaza cada párrafo EXACTO "[INSERTOS]" por los nodos w:p / w:tbl de los insertos.
 *  - Mantiene w:sectPr al final.
 *  - No copia estilos/num/medios externos; para textos legales suele bastar (se respeta pPr/rPr propios).
 */
async function composeInMemory(baseBuffer, insertIds = [], { pageBreakBetween = false } = {}) {
  const ids = insertIds.map(Number).filter(n => !Number.isNaN(n));
  if (!ids.length) {
    console.warn('[merge] no hay insertIds válidos');
    return baseBuffer;
  }

  const zip = new PizZip(baseBuffer);
  const docXmlPath = 'word/document.xml';
  const docXml = zip.file(docXmlPath)?.asText();
  if (!docXml) {
    console.warn('[merge] document.xml ausente');
    return baseBuffer;
  }

  // Analiza base
  const a = analyzeBodyForInsertos(docXml);
  if (!a.ok) {
    console.warn('[merge] base sin w:document/w:body');
    return baseBuffer;
  }
  const { docObj, bodyNode, bodyChildren, exactParasIdx, mixed } = a;
  console.log('[merge] placeholders', { exact: exactParasIdx.length, mixed });

  if (exactParasIdx.length === 0) return baseBuffer;

  // Prepara lote de nodos a insertar (concatenando todos los insertos en orden)
  const allNodes = [];
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const buf = await getInsertoBufferById(id);
    if (!buf) continue;
    const nodes = extractBodyNodesFromInserto(buf);
    if (!nodes.length) {
      console.warn(`[merge] inserto id=${id} no tiene body nodes; omitido`);
      continue;
    }
    if (i > 0 && pageBreakBetween) allNodes.push(makePageBreakParagraph());
    // deep copy para que no compartan referencia
    nodes.forEach(n => allNodes.push(JSON.parse(JSON.stringify(n))));
  }
  if (!allNodes.length) {
    console.warn('[merge] no hay nodos para insertar');
    return baseBuffer;
  }

  // Construye nuevo body reemplazando cada placeholder exacto por el lote
  const newBodyChildren = [];
  const setIdx = new Set(exactParasIdx);
  for (let i = 0; i < bodyChildren.length; i++) {
    const child = bodyChildren[i];
    if (setIdx.has(i)) {
      // inserta una COPIA del lote para cada placeholder
      allNodes.forEach(n => newBodyChildren.push(JSON.parse(JSON.stringify(n))));
      continue;
    }
    newBodyChildren.push(child);
  }

  // Mantener w:sectPr al final
  const sectNodes = newBodyChildren.filter(n => n['w:sectPr']);
  const others = newBodyChildren.filter(n => !n['w:sectPr']);
  bodyNode['w:body'] = sectNodes.length ? [...others, sectNodes[sectNodes.length - 1]] : others;

  // Serializa de vuelta
  const builderPO = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true, suppressEmptyNode: true });
  const newDocXml = builderPO.build(docObj);
  zip.file(docXmlPath, newDocXml);

  return zip.generate({ type: 'nodebuffer' });
}

module.exports = {
  composeInMemory,
  countInsertosPlaceholders
};
