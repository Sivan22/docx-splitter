import { describe, it, expect } from 'vitest';
import { readTag } from '../src/core/xml.js';

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
