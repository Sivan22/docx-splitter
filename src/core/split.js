import JSZip from 'jszip';
import { parseDocumentXml, buildDocumentXml } from './xml.js';
import { countWords } from './words.js';
import { META_PATH, buildMeta, hashString } from './meta.js';

const DOC_PATH = 'word/document.xml';

// Greedily pack children into groups whose word counts stay <= maxWords,
// cutting only at child boundaries. A single oversize child forms its own group.
export function partitionChildren(realChildren, maxWords) {
  const groups = [];
  let cur = [];
  let count = 0;
  for (const c of realChildren) {
    const wc = countWords(c.xml);
    if (cur.length > 0 && count + wc > maxWords) {
      groups.push(cur);
      cur = [];
      count = 0;
    }
    cur.push(c);
    count += wc;
  }
  if (cur.length) groups.push(cur);
  return groups;
}

const baseName = (filename) => filename.replace(/\.docx$/i, '');
const pad = (n, width) => String(n).padStart(width, '0');

// Split a .docx (ArrayBuffer | Uint8Array) into parts.
// Returns { parts: [{ name, bytes }], total, summary: [{ part, words }] }.
export async function splitDocx(data, filename, maxWords) {
  if (!(maxWords > 0)) throw new Error('Max words must be a positive number.');
  const zip = await JSZip.loadAsync(data);
  const docFile = zip.file(DOC_PATH);
  if (!docFile) throw new Error("This file isn't a valid .docx (no document.xml).");
  const documentXml = await docFile.async('string');
  const { prefix, suffix, sectPrXml, realChildren } = parseDocumentXml(documentXml);
  if (realChildren.length === 0) throw new Error('This document has no content to split.');
  const groups = partitionChildren(realChildren, maxWords);
  const total = groups.length;
  const hash = hashString(documentXml);
  const base = baseName(filename);
  const width = Math.max(2, String(total).length);

  // Reuse one in-memory zip across parts: every original entry stays
  // byte-identical; we only overwrite document.xml and the metadata sidecar
  // for each output part (avoids re-parsing the whole zip once per part).
  const parts = [];
  const summary = [];
  for (let i = 0; i < groups.length; i++) {
    const words = groups[i].reduce((a, c) => a + countWords(c.xml), 0);
    const childrenXml = groups[i].map((c) => c.xml).join('');
    zip.file(DOC_PATH, buildDocumentXml(prefix, childrenXml, sectPrXml, suffix));
    zip.file(META_PATH, buildMeta({ origin: filename, part: i + 1, total, hash }));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    parts.push({ name: `${base}_part${pad(i + 1, width)}.docx`, bytes });
    summary.push({ part: i + 1, words });
  }
  return { parts, total, summary };
}

// Bundle parts into one zip for download.
export async function zipParts(parts, zipName) {
  const zip = new JSZip();
  for (const p of parts) zip.file(p.name, p.bytes);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { name: zipName, bytes };
}
