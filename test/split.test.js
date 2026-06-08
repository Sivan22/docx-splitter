import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { makeDocx } from './helpers/makeDocx.js';
import { partitionChildren, splitDocx } from '../src/core/split.js';
import { countWords } from '../src/core/words.js';
import { META_PATH, readMetaJson } from '../src/core/meta.js';

const child = (text) => ({ tag: 'w:p', xml: `<w:p><w:t>${text}</w:t></w:p>` });

describe('partitionChildren', () => {
  it('packs children up to the word limit, cutting at boundaries', () => {
    const kids = [child('a b'), child('c d'), child('e f')]; // 2 words each
    const groups = partitionChildren(kids, 4);
    expect(groups.map((g) => g.length)).toEqual([2, 1]);
  });
  it('never splits inside a paragraph: an oversize paragraph is its own part', () => {
    const kids = [child('a b c d e'), child('f')]; // 5 words, then 1
    const groups = partitionChildren(kids, 2);
    expect(groups.map((g) => g.length)).toEqual([1, 1]);
    expect(countWords(groups[0][0].xml)).toBe(5);
  });
  it('returns a single group when everything fits', () => {
    const kids = [child('a'), child('b')];
    expect(partitionChildren(kids, 100)).toHaveLength(1);
  });
});

describe('splitDocx', () => {
  it('produces parts that each carry correct metadata', async () => {
    const data = await makeDocx(['a b', 'c d', 'e f']); // 2 words each
    const { parts, total, summary } = await splitDocx(data, 'Doc.docx', 4);
    expect(total).toBe(2);
    expect(parts.map((p) => p.name)).toEqual(['Doc_part01.docx', 'Doc_part02.docx']);
    expect(summary).toEqual([
      { part: 1, words: 4 },
      { part: 2, words: 2 },
    ]);
    const z = await JSZip.loadAsync(parts[0].bytes);
    const meta = readMetaJson(await z.file(META_PATH).async('string'));
    expect(meta).toMatchObject({ origin: 'Doc.docx', part: 1, total: 2 });
  });
  it('keeps the body-level sectPr in every part', async () => {
    const data = await makeDocx(['a', 'b', 'c']);
    const { parts } = await splitDocx(data, 'Doc.docx', 1);
    for (const p of parts) {
      const z = await JSZip.loadAsync(p.bytes);
      const doc = await z.file('word/document.xml').async('string');
      expect(doc).toContain('<w:sectPr>');
    }
  });
  it('rejects a non-positive max', async () => {
    const data = await makeDocx(['a']);
    await expect(splitDocx(data, 'Doc.docx', 0)).rejects.toThrow();
  });
  it('rejects a document with no content to split', async () => {
    const data = await makeDocx([]);
    await expect(splitDocx(data, 'Doc.docx', 5)).rejects.toThrow(/no content/i);
  });
  it('produces a valid OPC package: every part has a declared content type', async () => {
    // Word shows "unreadable content" if any zip part lacks a content type in
    // [Content_Types].xml (OPC rule). The metadata sidecar must be declared.
    const data = await makeDocx(['a b', 'c d']);
    const { parts } = await splitDocx(data, 'Doc.docx', 2);
    const z = await JSZip.loadAsync(parts[0].bytes);
    const ct = await z.file('[Content_Types].xml').async('string');
    const defaults = [...ct.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase());
    const overrides = [...ct.matchAll(/<Override PartName="([^"]+)"/g)].map((m) => m[1]);
    for (const path of Object.keys(z.files)) {
      if (z.files[path].dir || path === '[Content_Types].xml') continue;
      const ext = (path.split('.').pop() || '').toLowerCase();
      const covered = defaults.includes(ext) || overrides.includes('/' + path);
      expect(covered, `part "${path}" must have a declared content type`).toBe(true);
    }
  });
  it('references the metadata part from the package relationships (not orphaned)', async () => {
    // An extra part that no relationship points to is treated by Word as not
    // belonging to the package — it must be referenced from _rels/.rels.
    const data = await makeDocx(['a b', 'c d']);
    const { parts } = await splitDocx(data, 'Doc.docx', 2);
    const z = await JSZip.loadAsync(parts[0].bytes);
    const rels = await z.file('_rels/.rels').async('string');
    expect(rels).toContain('Target="docx-splitter-meta.json"');
  });
});
