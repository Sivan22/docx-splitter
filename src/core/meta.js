export const META_PATH = 'docx-splitter-meta.json';

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
