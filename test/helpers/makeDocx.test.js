import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { makeDocx } from './makeDocx.js';

describe('makeDocx', () => {
  it('produces a zip containing word/document.xml with the paragraphs', async () => {
    const bytes = await makeDocx(['Hello world', 'Second para']);
    const zip = await JSZip.loadAsync(bytes);
    const doc = await zip.file('word/document.xml').async('string');
    expect(doc).toContain('Hello world');
    expect(doc).toContain('Second para');
    expect(doc).toContain('<w:sectPr>');
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('_rels/.rels')).not.toBeNull();
  });
});
