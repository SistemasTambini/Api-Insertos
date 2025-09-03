const PizZip = require('pizzip');
const { XMLParser, XMLBuilder } = require('fast-xml-parser');

/** ====== Util: texto plano de un w:p (preserveOrder) ====== */
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
            else if (v && typeof v === 'object' && '#text' in it) text += String(it['#text'] ?? '');
          }
        } else if (typeof v === 'string') {
          text += v;
        } else if (v && typeof v === 'object' && '#text' in v) {
          text += String(v['#text'] ?? '');
        }
      } else if (Array.isArray(v)) v.forEach(walk);
      else if (typeof v === 'object') walk(v);
    }
  }
  walk(pNode);
  return text;
}

/** ====== Busca marcador en body (case-insensitive) ====== */
function analyzeBodyForPlaceholder(docXml, placeholder = '[MINUTA]') {
  const PLACE = String(placeholder).trim();
  const parserPO = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true });
  const docObj = parserPO.parse(docXml);

  const findNode = (arr, tag) => Array.isArray(arr) ? arr.find(n => n[tag]) : undefined;
  const docNode = findNode(docObj, 'w:document');
  if (!docNode) return { ok: false };

  const bodyNode = findNode(docNode['w:document'], 'w:body');
  if (!bodyNode) return { ok: false };

  const bodyChildren = bodyNode['w:body'] || [];
  const exact = [];
  let mixed = 0;

  bodyChildren.forEach((child, idx) => {
    if (!child['w:p']) return;
    const t = getParagraphText(child);
    if (!t) return;
    const has = t.toUpperCase().includes(PLACE.toUpperCase());
    if (!has) return;
    if (t.trim().toUpperCase() === PLACE.toUpperCase()) exact.push(idx);
    else mixed++;
  });

  return { ok: true, docObj, bodyNode, bodyChildren, exact, mixed };
}

/** ====== Cuenta placeholders en body ====== */
function countPlaceholder(baseBuffer, placeholder = '[MINUTA]') {
  const zip = new PizZip(baseBuffer);
  const docXml = zip.file('word/document.xml')?.asText();
  if (!docXml) return { exact: 0, mixed: 0 };
  const a = analyzeBodyForPlaceholder(docXml, placeholder);
  if (!a.ok) return { exact: 0, mixed: 0 };
  return { exact: a.exact.length, mixed: a.mixed };
}

/** ====== Extrae nodos del body de un DOCX (sin sectPr) y devuelve también el zip ====== */
function extractBodyNodesFromDoc(buffer) {
  const inz = new PizZip(buffer);
  const docXml = inz.file('word/document.xml')?.asText();
  if (!docXml) return { nodes: [], inz: null };
  const parserPO = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true });
  const docObj = parserPO.parse(docXml);

  const findNode = (arr, tag) => Array.isArray(arr) ? arr.find(n => n[tag]) : undefined;
  const docNode = findNode(docObj, 'w:document');
  if (!docNode) return { nodes: [], inz };
  const bodyNode = findNode(docNode['w:document'], 'w:body');
  if (!bodyNode) return { nodes: [], inz };
  const children = bodyNode['w:body'] || [];
  return { nodes: children.filter(n => !n['w:sectPr']), inz };
}

/** ====== (Opcional) salto de página ====== */
function makePageBreakParagraph() {
  return { 'w:p': [ { 'w:r': [ { 'w:br': { ':@': { 'w:type': 'page' } } } ] } ] };
}

/** ====== RELS / imágenes: helpers ====== */
function getOrInitBaseRelsForDocument(baseZip) {
  const relsPath = 'word/_rels/document.xml.rels';
  let relsXml = baseZip.file(relsPath)?.asText();
  if (!relsXml) {
    relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    baseZip.file(relsPath, relsXml);
  }
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  const relsObj = parser.parse(relsXml);
  const root = relsObj.Relationships || { Relationship: [] };
  if (!Array.isArray(root.Relationship)) root.Relationship = root.Relationship ? [root.Relationship] : [];
  const xmlns = root.xmlns || 'http://schemas.openxmlformats.org/package/2006/relationships';
  return { relsObj, root, xmlns, relsPath };
}

function getInsertoImageRels(inz) {
  const relsXml = inz.file('word/_rels/document.xml.rels')?.asText();
  if (!relsXml) return [];
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' });
  const obj = parser.parse(relsXml);
  let rels = obj?.Relationships?.Relationship || [];
  if (!Array.isArray(rels)) rels = [rels];
  return rels.filter(r => (r?.Type || '').endsWith('/image') && typeof r.Target === 'string');
}

function remapEmbedIdsInNodes(nodes, idMap) {
  function walk(n) {
    if (!n || typeof n !== 'object') return;
    const attrs = n[':@'];
    if (attrs) {
      if (attrs['r:embed'] && idMap[attrs['r:embed']]) attrs['r:embed'] = idMap[attrs['r:embed']];
      if (attrs['r:id'] && idMap[attrs['r:id']])       attrs['r:id']    = idMap[attrs['r:id']];
    }
    for (const k of Object.keys(n)) {
      if (k === ':@') continue;
      const v = n[k];
      if (Array.isArray(v)) v.forEach(walk);
      else if (typeof v === 'object') walk(v);
    }
  }
  nodes.forEach(walk);
}

/**
 * ====== MERGE en memoria (BODY), con imágenes ======
 * Reemplaza cada párrafo EXACTO "[MINUTA]" por los nodos de la minuta.
 * Copia imágenes de la minuta → word/media del destino, crea relaciones y remapea r:embed.
 */
function mergeMinutaInMemory(baseBuffer, minutaBuffer, {
  placeholder = '[MINUTA]',
  pageBreakBefore = false,
  pageBreakAfter  = false
} = {}) {
  const baseZip = new PizZip(baseBuffer);
  const docXmlPath = 'word/document.xml';
  const docXml = baseZip.file(docXmlPath)?.asText();
  if (!docXml) return baseBuffer;

  const analysis = analyzeBodyForPlaceholder(docXml, placeholder);
  if (!analysis.ok) return baseBuffer;
  const { docObj, bodyNode, bodyChildren, exact } = analysis;
  if (!exact.length) return baseBuffer;

  // Nodos de la minuta + ZIP origen
  const { nodes: nodesRaw, inz } = extractBodyNodesFromDoc(minutaBuffer);
  if (!nodesRaw.length || !inz) return baseBuffer;

  // Copiar imágenes de la minuta al documento base y remapear r:embed en los nodos a insertar
  const { relsObj, root: baseRelsRoot, xmlns: baseRelsXmlns, relsPath } = getOrInitBaseRelsForDocument(baseZip);
  const imgRels = getInsertoImageRels(inz);
  const idMap = {};
  let ridCounter = 23000;

  for (const rel of imgRels) {
    const oldId = rel.Id;
    const target = rel.Target; // p.ej. "media/image1.png"
    const file = inz.file(`word/${target}`);
    if (!file) continue;

    const ext = (target.split('.').pop() || 'bin').toLowerCase();
    const bin = file.asUint8Array?.() || file.asArrayBuffer?.();
    if (!bin) continue;

    const newName = `img_m_${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
    baseZip.file(`word/media/${newName}`, Buffer.from(bin));

    const newRid = `rIdM${ridCounter++}`;
    baseRelsRoot.Relationship.push({
      Id: newRid,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      Target: `media/${newName}`
    });

    idMap[oldId] = newRid;
  }

  // Deep copy y remapeo
  const nodes = nodesRaw.map(n => JSON.parse(JSON.stringify(n)));
  if (Object.keys(idMap).length) remapEmbedIdsInNodes(nodes, idMap);

  // Construye nuevo body reemplazando marcadores
  const newBody = [];
  const idxSet = new Set(exact);
  for (let i = 0; i < bodyChildren.length; i++) {
    const child = bodyChildren[i];
    if (!idxSet.has(i)) { newBody.push(child); continue; }

    if (pageBreakBefore) newBody.push(makePageBreakParagraph());
    nodes.forEach(n => newBody.push(JSON.parse(JSON.stringify(n))));
    if (pageBreakAfter) newBody.push(makePageBreakParagraph());
  }

  // sectPr al final
  const sect = newBody.filter(n => n['w:sectPr']);
  const others = newBody.filter(n => !n['w:sectPr']);
  bodyNode['w:body'] = sect.length ? [...others, sect[sect.length - 1]] : others;

  // Serializar document.xml
  const builderPO = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true, suppressEmptyNode: true });
  const newDocXml = builderPO.build(docObj);
  baseZip.file(docXmlPath, newDocXml);

  // Serializar rels de document.xml
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '' });
  relsObj.Relationships = baseRelsRoot;
  relsObj.Relationships.xmlns = baseRelsXmlns;
  const newRelsXml = builder.build(relsObj);
  baseZip.file(relsPath, newRelsXml);

  return baseZip.generate({ type: 'nodebuffer' });
}

module.exports = {
  countPlaceholder,
  mergeMinutaInMemory
};
