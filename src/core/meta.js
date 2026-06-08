export const META_PATH = 'docx-splitter-meta.json';
export const CONTENT_TYPES_PATH = '[Content_Types].xml';

// The OPC content-type declaration for the metadata sidecar. Every part in a
// .docx must have a content type or Word reports "unreadable content"; the
// sidecar's .json extension is undeclared in a normal document, so each split
// part adds this targeted Override (by exact part name, so it never collides).
const META_OVERRIDE = `<Override PartName="/${META_PATH}" ContentType="application/json"/>`;

// Add the metadata Override to [Content_Types].xml (idempotent).
export function declareMetaContentType(contentTypesXml) {
  if (contentTypesXml.includes(META_OVERRIDE)) return contentTypesXml;
  const i = contentTypesXml.lastIndexOf('</Types>');
  if (i === -1) throw new Error("This file isn't a valid .docx ([Content_Types].xml is malformed).");
  return contentTypesXml.slice(0, i) + META_OVERRIDE + contentTypesXml.slice(i);
}

// Remove the metadata Override, restoring [Content_Types].xml byte-for-byte.
export function undeclareMetaContentType(contentTypesXml) {
  return contentTypesXml.replace(META_OVERRIDE, '');
}

// Small, dependency-free FNV-1a hash. Used only to group parts of one split
// (not for security), so collision risk is irrelevant here.
export function hashString(s) {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, '0');
}

export function buildMeta({ origin, part, total, hash }) {
  return JSON.stringify({ schema: 1, tool: 'docx-splitter', origin, part, total, hash });
}

export function readMetaJson(str) {
  let obj;
  try {
    obj = JSON.parse(str);
  } catch {
    throw new Error('Invalid split metadata.');
  }
  if (
    !obj ||
    obj.tool !== 'docx-splitter' ||
    typeof obj.origin !== 'string' ||
    typeof obj.part !== 'number' ||
    typeof obj.total !== 'number' ||
    typeof obj.hash !== 'string'
  ) {
    throw new Error('Invalid split metadata.');
  }
  return obj;
}
