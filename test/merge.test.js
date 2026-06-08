import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { makeDocx } from './helpers/makeDocx.js';
import { splitDocx } from '../src/core/split.js';
import { mergeDocx } from '../src/core/merge.js';

async function internalParts(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const out = {};
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir) continue;
    out[path] = await zip.file(path).async('string');
  }
  return out;
}

const toFiles = (parts) => parts.map((p) => ({ name: p.name, data: p.bytes }));

describe('mergeDocx round-trip', () => {
  it('reproduces every internal part byte-for-byte', async () => {
    const original = await makeDocx(['alpha beta', 'gamma delta', 'epsilon', 'zeta eta theta']);
    const { parts } = await splitDocx(original, 'Doc.docx', 3);
    const merged = await mergeDocx(toFiles(parts));

    const before = await internalParts(original);
    const after = await internalParts(merged.bytes);
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort()); // meta removed
    for (const path of Object.keys(before)) {
      expect(after[path]).toBe(before[path]);
    }
    expect(merged.name).toBe('Doc.docx');
  });

  it('works regardless of upload order', async () => {
    const original = await makeDocx(['a', 'b', 'c', 'd']);
    const { parts } = await splitDocx(original, 'Doc.docx', 1);
    const merged = await mergeDocx(toFiles([...parts].reverse()));
    const after = await internalParts(merged.bytes);
    const before = await internalParts(original);
    expect(after['word/document.xml']).toBe(before['word/document.xml']);
  });
});

describe('mergeDocx validation', () => {
  it('rejects an incomplete set', async () => {
    const original = await makeDocx(['a', 'b', 'c']);
    const { parts } = await splitDocx(original, 'Doc.docx', 1);
    await expect(mergeDocx(toFiles([parts[0], parts[2]]))).rejects.toThrow(/Missing part 2/);
  });
  it('rejects a duplicate part', async () => {
    const original = await makeDocx(['a', 'b']);
    const { parts } = await splitDocx(original, 'Doc.docx', 1);
    await expect(mergeDocx(toFiles([parts[0], parts[0]]))).rejects.toThrow(/more than once/);
  });
  it('rejects a foreign part from a different document', async () => {
    const a = await splitDocx(await makeDocx(['a', 'b']), 'A.docx', 1);
    const b = await splitDocx(await makeDocx(['c', 'd']), 'B.docx', 1);
    await expect(mergeDocx(toFiles([a.parts[0], b.parts[1]]))).rejects.toThrow(/different document/);
  });
  it('rejects a non-part file', async () => {
    const plain = await makeDocx(['a']);
    await expect(mergeDocx([{ name: 'plain.docx', data: plain }])).rejects.toThrow(/not a split part/);
  });
});
