import { describe, it, expect } from 'vitest';
import { META_PATH, hashString, buildMeta, readMetaJson } from '../src/core/meta.js';

describe('meta', () => {
  it('exposes a stable sidecar path', () => {
    expect(META_PATH).toBe('docx-splitter-meta.json');
  });
  it('hashes deterministically and differs for different input', () => {
    expect(hashString('abc')).toBe(hashString('abc'));
    expect(hashString('abc')).not.toBe(hashString('abd'));
    expect(hashString('abc')).toMatch(/^[0-9a-f]{8}$/);
  });
  it('round-trips metadata through build/read', () => {
    const json = buildMeta({ origin: 'My Doc.docx', part: 2, total: 5, hash: 'deadbeef' });
    const obj = readMetaJson(json);
    expect(obj).toMatchObject({ origin: 'My Doc.docx', part: 2, total: 5, hash: 'deadbeef' });
  });
  it('rejects non-tool JSON', () => {
    expect(() => readMetaJson('{"foo":1}')).toThrow();
    expect(() => readMetaJson('not json')).toThrow();
  });
});
