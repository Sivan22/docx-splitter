import JSZip from 'jszip';
import { parseDocumentXml, buildDocumentXml } from './xml.js';
import { META_PATH, CONTENT_TYPES_PATH, readMetaJson, undeclareMetaContentType } from './meta.js';

const DOC_PATH = 'word/document.xml';

// files: Array<{ name, data }>, where data is an ArrayBuffer | Uint8Array of a part.
// Returns { name, bytes } — a single .docx whose internal parts equal the original.
export async function mergeDocx(files) {
  if (!files || files.length === 0) throw new Error('No files to merge.');

  const loaded = [];
  for (const f of files) {
    const zip = await JSZip.loadAsync(f.data);
    const metaFile = zip.file(META_PATH);
    if (!metaFile) throw new Error(`"${f.name}" is not a split part made by this tool.`);
    const meta = readMetaJson(await metaFile.async('string'));
    const docFile = zip.file(DOC_PATH);
    if (!docFile) throw new Error(`"${f.name}" is missing document.xml.`);
    loaded.push({ name: f.name, meta, documentXml: await docFile.async('string'), zip });
  }

  const { origin, hash, total } = loaded[0].meta;
  for (const l of loaded) {
    if (l.meta.hash !== hash || l.meta.origin !== origin) {
      throw new Error(`"${l.name}" belongs to a different document.`);
    }
  }

  const byIndex = new Map();
  for (const l of loaded) {
    if (byIndex.has(l.meta.part)) throw new Error(`Part ${l.meta.part} appears more than once.`);
    byIndex.set(l.meta.part, l);
  }
  for (let i = 1; i <= total; i++) {
    if (!byIndex.has(i)) throw new Error(`Missing part ${i} of ${total}.`);
  }

  const ordered = [];
  for (let i = 1; i <= total; i++) ordered.push(byIndex.get(i));

  // Every part shares the same prefix/sectPr/suffix (splitDocx writes them
  // identically into each part), so we take that envelope from part 1 and only
  // concatenate the content children from every part, in order.
  const template = parseDocumentXml(ordered[0].documentXml);
  let childrenXml = '';
  for (const l of ordered) {
    childrenXml += parseDocumentXml(l.documentXml).realChildren.map((c) => c.xml).join('');
  }
  const mergedDocXml = buildDocumentXml(template.prefix, childrenXml, template.sectPrXml, template.suffix);

  const outZip = ordered[0].zip; // all non-document parts are byte-identical to the original
  outZip.file(DOC_PATH, mergedDocXml);
  outZip.remove(META_PATH);
  // Undo the metadata content-type declaration split.js added, restoring the
  // original [Content_Types].xml exactly.
  const ctFile = outZip.file(CONTENT_TYPES_PATH);
  if (ctFile) outZip.file(CONTENT_TYPES_PATH, undeclareMetaContentType(await ctFile.async('string')));
  const bytes = await outZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { name: origin, bytes };
}
