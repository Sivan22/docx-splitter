import { describe, it, expect } from 'vitest';
import { readTag } from '../src/core/xml.js';
import { splitTopLevelElements } from '../src/core/xml.js';

describe('readTag', () => {
  it('reads an open tag with attributes', () => {
    const s = '<w:p w:rsidR="00A"><w:r/>';
    const t = readTag(s, 0);
    expect(t).toEqual({ kind: 'open', name: 'w:p', end: 19 });
  });
  it('reads a self-closing tag', () => {
    const s = '<w:br/>';
    const t = readTag(s, 0);
    expect(t).toEqual({ kind: 'selfclose', name: 'w:br', end: 7 });
  });
  it('reads a close tag', () => {
    const s = '</w:p>';
    const t = readTag(s, 0);
    expect(t).toEqual({ kind: 'close', name: 'w:p', end: 6 });
  });
  it('ignores > inside quoted attribute values', () => {
    const s = '<w:t w:val="a>b"/>';
    const t = readTag(s, 0);
    expect(t.kind).toBe('selfclose');
    expect(t.end).toBe(s.length);
  });
  it('reads a comment', () => {
    const s = '<!-- hi -->X';
    const t = readTag(s, 0);
    expect(t.kind).toBe('comment');
    expect(t.end).toBe(11);
  });
});

describe('splitTopLevelElements', () => {
  it('splits two sibling paragraphs', () => {
    const inner = '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:t>b</w:t></w:p>';
    const { leading, children } = splitTopLevelElements(inner);
    expect(leading).toBe('');
    expect(children.map((c) => c.tag)).toEqual(['w:p', 'w:p']);
    expect(leading + children.map((c) => c.xml).join('')).toBe(inner);
  });
  it('handles nested same-name elements (tables containing paragraphs)', () => {
    const inner = '<w:tbl><w:tr><w:tc><w:p><w:t>x</w:t></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const { children } = splitTopLevelElements(inner);
    expect(children.map((c) => c.tag)).toEqual(['w:tbl', 'w:p']);
    expect(children.map((c) => c.xml).join('')).toBe(inner);
  });
  it('folds leading and inter-element whitespace losslessly', () => {
    const inner = '\n  <w:p/>\n  <w:p/>\n';
    const { leading, children } = splitTopLevelElements(inner);
    expect(leading).toBe('\n  ');
    expect(leading + children.map((c) => c.xml).join('')).toBe(inner);
  });
});
