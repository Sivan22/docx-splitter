import { describe, it, expect } from 'vitest';
import { readTag } from '../src/core/xml.js';
import { splitTopLevelElements } from '../src/core/xml.js';
import { parseDocumentXml, buildDocumentXml } from '../src/core/xml.js';

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

const DOC = `<?xml version="1.0"?><w:document xmlns:w="x"><w:body>` +
  `<w:p><w:t>one</w:t></w:p>` +
  `<w:p><w:t>two</w:t></w:p>` +
  `<w:sectPr><w:pgSz/></w:sectPr>` +
  `</w:body></w:document>`;

describe('parseDocumentXml', () => {
  it('separates prefix, real children, body-level sectPr, and suffix', () => {
    const r = parseDocumentXml(DOC);
    expect(r.realChildren.map((c) => c.tag)).toEqual(['w:p', 'w:p']);
    expect(r.sectPrXml).toBe('<w:sectPr><w:pgSz/></w:sectPr>');
    expect(r.prefix.endsWith('<w:body>')).toBe(true);
    expect(r.suffix).toBe('</w:body></w:document>');
  });
  it('satisfies the round-trip invariant exactly', () => {
    const r = parseDocumentXml(DOC);
    const rebuilt = r.prefix + r.realChildren.map((c) => c.xml).join('') + r.sectPrXml + r.suffix;
    expect(rebuilt).toBe(DOC);
  });
  it('treats a document with no body-level sectPr as empty sectPr', () => {
    const doc = `<w:document><w:body><w:p><w:t>x</w:t></w:p></w:body></w:document>`;
    const r = parseDocumentXml(doc);
    expect(r.sectPrXml).toBe('');
    expect(r.realChildren).toHaveLength(1);
    expect(r.prefix + r.realChildren[0].xml + r.suffix).toBe(doc);
  });
  it('does not prefix-match <w:bodyPr> before the real <w:body>', () => {
    const doc = `<w:document><w:bodyPr/><w:body><w:p><w:t>x</w:t></w:p><w:sectPr/></w:body></w:document>`;
    const r = parseDocumentXml(doc);
    expect(r.realChildren).toHaveLength(1);
    expect(r.sectPrXml).toBe('<w:sectPr/>');
    expect(r.prefix + r.realChildren.map((c) => c.xml).join('') + r.sectPrXml + r.suffix).toBe(doc);
  });
  it('handles a literal > inside a <w:body> attribute value', () => {
    const doc = `<w:document><w:body w:x="a>b"><w:p><w:t>y</w:t></w:p><w:sectPr/></w:body></w:document>`;
    const r = parseDocumentXml(doc);
    expect(r.realChildren).toHaveLength(1);
    expect(r.prefix.endsWith('>')).toBe(true);
    expect(r.prefix + r.realChildren.map((c) => c.xml).join('') + r.sectPrXml + r.suffix).toBe(doc);
  });
});

describe('buildDocumentXml', () => {
  it('concatenates the four parts in order', () => {
    expect(buildDocumentXml('A', 'B', 'C', 'D')).toBe('ABCD');
  });
});
