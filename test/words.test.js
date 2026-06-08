import { describe, it, expect } from 'vitest';
import { countWords } from '../src/core/words.js';

describe('countWords', () => {
  it('counts words across multiple runs in a paragraph', () => {
    const p = '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>brave world</w:t></w:r></w:p>';
    expect(countWords(p)).toBe(3);
  });
  it('returns 0 for an empty paragraph', () => {
    expect(countWords('<w:p/>')).toBe(0);
  });
  it('does not split a single word stored across two runs', () => {
    const p = '<w:p><w:r><w:t>Hel</w:t></w:r><w:r><w:t>lo</w:t></w:r></w:p>';
    expect(countWords(p)).toBe(1);
  });
  it('treats <w:tab/> and <w:br/> as whitespace', () => {
    const p = '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>';
    expect(countWords(p)).toBe(3);
  });
  it('decodes entities without breaking word counts', () => {
    const p = '<w:p><w:t>Tom &amp; Jerry</w:t></w:p>';
    expect(countWords(p)).toBe(3);
  });
  it('counts words inside table cells', () => {
    const t = '<w:tbl><w:tr><w:tc><w:p><w:t>one two</w:t></w:p></w:tc></w:tr></w:tbl>';
    expect(countWords(t)).toBe(2);
  });
});
