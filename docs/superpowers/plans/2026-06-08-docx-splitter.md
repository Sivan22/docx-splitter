# DOCX Splitter & Merger Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A static, GitHub Pages–hosted single-page app that splits a `.docx` by word count at paragraph breaks and losslessly merges the parts back into the original.

**Architecture:** 100% client-side. Core logic operates on the raw XML *string* of `word/document.xml` — it partitions the body's top-level children (paragraphs/tables) as verbatim substrings and never re-serializes untouched XML, so the round-trip is lossless by construction. Parts carry a metadata sidecar (`docx-splitter-meta.json`) inside the zip for foolproof merge ordering. A thin UI layer wires the pure core to drop zones and downloads.

**Tech Stack:** Vanilla JS (ES modules), Vite (build/dev), Vitest (tests), JSZip (zip read/write). No backend, no framework.

**Key invariant (the whole design rests on this):** For `parseDocumentXml(xml)` →
`prefix + realChildren.map(c=>c.xml).join('') + sectPrXml + suffix === xml` exactly.
Because every child's `xml` is an exact original substring, reassembling them reproduces `document.xml` byte-for-byte.

---

## File Structure

```
docx-splitter/
  index.html                 # Vite entry + UI markup (Split/Merge tabs)
  package.json
  vite.config.js             # base: './', vitest config
  src/
    main.js                  # UI wiring only (tabs, drop zones, downloads)
    styles.css
    core/
      xml.js                 # readTag, splitTopLevelElements, parseDocumentXml, buildDocumentXml
      words.js               # countWords
      meta.js                # META_PATH, hashString, buildMeta, readMetaJson
      split.js               # partitionChildren, splitDocx, zipParts
      merge.js               # mergeDocx
  test/
    helpers/makeDocx.js      # build a minimal valid .docx (Uint8Array) for tests
    xml.test.js
    words.test.js
    meta.test.js
    split.test.js
    merge.test.js
  .github/workflows/deploy.yml
  .gitignore                 # already exists (node_modules/, dist/)
```

Responsibilities are split so each file holds one concern: `xml.js` is pure string/XML slicing, `words.js` only counts, `meta.js` only handles the sidecar, `split.js`/`merge.js` orchestrate zip I/O, `main.js` is UI only.

---

## Task 1: Project scaffold

**Files:**
- Create: `package.json`, `vite.config.js`, `index.html`, `src/main.js`, `src/styles.css`, `test/smoke.test.js`

- [ ] **Step 1: Create `package.json`**

```json
{
  "name": "docx-splitter",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "preview": "vite preview",
    "test": "vitest run"
  },
  "dependencies": {
    "jszip": "^3.10.1"
  },
  "devDependencies": {
    "vite": "^5.4.0",
    "vitest": "^2.1.8"
  }
}
```

- [ ] **Step 2: Create `vite.config.js`**

```js
import { defineConfig } from 'vite';

export default defineConfig({
  base: './', // relative paths so it works on any GitHub Pages sub-path
  test: {
    environment: 'node',
    include: ['test/**/*.test.js'],
  },
});
```

- [ ] **Step 3: Create minimal `index.html` and `src/main.js` and `src/styles.css` stubs**

`index.html`:
```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>DOCX Splitter &amp; Merger</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/main.js"></script>
  </body>
</html>
```

`src/main.js`:
```js
import './styles.css';

document.querySelector('#app').textContent = 'DOCX Splitter — scaffold OK';
```

`src/styles.css`:
```css
:root { font-family: system-ui, sans-serif; }
```

- [ ] **Step 4: Create `test/smoke.test.js`**

```js
import { describe, it, expect } from 'vitest';

describe('smoke', () => {
  it('runs the test suite', () => {
    expect(1 + 1).toBe(2);
  });
});
```

- [ ] **Step 5: Install deps and run tests**

Run: `npm install && npm test`
Expected: smoke test PASSES (1 passed).

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "chore: scaffold Vite + Vitest project"
```

---

## Task 2: Test helper — build a minimal `.docx`

**Files:**
- Create: `test/helpers/makeDocx.js`
- Test: `test/helpers/makeDocx.test.js`

- [ ] **Step 1: Write the failing test**

`test/helpers/makeDocx.test.js`:
```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/helpers/makeDocx.test.js`
Expected: FAIL — cannot resolve `./makeDocx.js`.

- [ ] **Step 3: Implement `test/helpers/makeDocx.js`**

```js
import JSZip from 'jszip';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// paragraphs: string[]. opts.sectPr: include body-level sectPr (default true).
export async function makeDocx(paragraphs, { sectPr = true } = {}) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', RELS);
  const body = paragraphs
    .map((p) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(p)}</w:t></w:r></w:p>`)
    .join('');
  const sect = sectPr ? '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' : '';
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}${sect}</w:body></w:document>`;
  zip.file('word/document.xml', doc);
  return zip.generateAsync({ type: 'uint8array' });
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/helpers/makeDocx.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "test: add makeDocx fixture helper"
```

---

## Task 3: `xml.js` — `readTag` tag scanner

**Files:**
- Create: `src/core/xml.js`
- Test: `test/xml.test.js`

- [ ] **Step 1: Write the failing test**

`test/xml.test.js`:
```js
import { describe, it, expect } from 'vitest';
import { readTag } from '../src/core/xml.js';

describe('readTag', () => {
  it('reads an open tag with attributes', () => {
    const s = '<w:p w:rsidR="00A"><w:r/>';
    const t = readTag(s, 0);
    expect(t).toEqual({ kind: 'open', name: 'w:p', end: 18 });
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/xml.test.js`
Expected: FAIL — `readTag` is not exported.

- [ ] **Step 3: Implement `readTag` in `src/core/xml.js`**

```js
// Parse a single markup token starting at s[pos] === '<'.
// Returns { kind, name, end } where end is the index just past the token.
export function readTag(s, pos) {
  if (s.startsWith('<!--', pos)) {
    const i = s.indexOf('-->', pos);
    return { kind: 'comment', name: null, end: i === -1 ? s.length : i + 3 };
  }
  if (s.startsWith('<![CDATA[', pos)) {
    const i = s.indexOf(']]>', pos);
    return { kind: 'cdata', name: null, end: i === -1 ? s.length : i + 3 };
  }
  if (s.startsWith('<?', pos)) {
    const i = s.indexOf('?>', pos);
    return { kind: 'pi', name: null, end: i === -1 ? s.length : i + 2 };
  }
  const isClose = s[pos + 1] === '/';
  let i = pos + (isClose ? 2 : 1);
  const nameStart = i;
  while (i < s.length && !/[\s/>]/.test(s[i])) i++;
  const name = s.slice(nameStart, i);
  let quote = null;
  while (i < s.length) {
    const c = s[i];
    if (quote) {
      if (c === quote) quote = null;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === '>') {
      const selfClose = s[i - 1] === '/';
      return { kind: isClose ? 'close' : selfClose ? 'selfclose' : 'open', name, end: i + 1 };
    }
    i++;
  }
  return { kind: isClose ? 'close' : 'open', name, end: s.length };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/xml.test.js`
Expected: PASS (5 passed).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add readTag XML token scanner"
```

---

## Task 4: `xml.js` — `splitTopLevelElements`

**Files:**
- Modify: `src/core/xml.js`
- Test: `test/xml.test.js`

- [ ] **Step 1: Add failing tests**

Append to `test/xml.test.js`:
```js
import { splitTopLevelElements } from '../src/core/xml.js';

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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/xml.test.js`
Expected: FAIL — `splitTopLevelElements` not exported.

- [ ] **Step 3: Implement `splitTopLevelElements`**

Append to `src/core/xml.js`:
```js
// Split the inner XML of <w:body> into top-level child elements.
// Each child's xml is the exact substring from its '<' up to the next child's
// '<' (trailing-run semantics), so leading + children.join('') === inner.
export function splitTopLevelElements(inner) {
  const len = inner.length;
  const els = []; // { start, end, tag }
  let p = 0;
  while (p < len) {
    const lt = inner.indexOf('<', p);
    if (lt === -1) break;
    let depth = 0;
    let q = lt;
    let topName = null;
    while (q < len) {
      if (inner[q] === '<') {
        const tag = readTag(inner, q);
        if (topName === null && (tag.kind === 'open' || tag.kind === 'selfclose')) {
          topName = tag.name;
        }
        if (tag.kind === 'open') depth++;
        else if (tag.kind === 'close') depth--;
        q = tag.end;
        if (depth === 0) break;
      } else {
        q++;
      }
    }
    els.push({ start: lt, end: q, tag: topName });
    p = q;
  }
  const children = els.map((el, i) => {
    const xmlEnd = i + 1 < els.length ? els[i + 1].start : len;
    return { tag: el.tag, xml: inner.slice(el.start, xmlEnd) };
  });
  const leading = els.length ? inner.slice(0, els[0].start) : inner;
  return { leading, children };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/xml.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: split body inner XML into top-level children"
```

---

## Task 5: `xml.js` — `parseDocumentXml` + `buildDocumentXml`

**Files:**
- Modify: `src/core/xml.js`
- Test: `test/xml.test.js`

- [ ] **Step 1: Add failing tests (the core round-trip invariant)**

Append to `test/xml.test.js`:
```js
import { parseDocumentXml, buildDocumentXml } from '../src/core/xml.js';

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
});

describe('buildDocumentXml', () => {
  it('concatenates the four parts in order', () => {
    expect(buildDocumentXml('A', 'B', 'C', 'D')).toBe('ABCD');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/xml.test.js`
Expected: FAIL — functions not exported.

- [ ] **Step 3: Implement both functions**

Append to `src/core/xml.js`:
```js
// Decompose word/document.xml so the body's top-level children can be sliced.
// Invariant: prefix + realChildren.map(c=>c.xml).join('') + sectPrXml + suffix === xml
export function parseDocumentXml(xml) {
  const bodyOpen = xml.indexOf('<w:body');
  if (bodyOpen === -1) {
    return { prefix: xml, suffix: '', sectPrXml: '', realChildren: [] };
  }
  const openEnd = xml.indexOf('>', bodyOpen);
  if (xml[openEnd - 1] === '/') {
    // self-closing <w:body/> — no children
    return { prefix: xml.slice(0, openEnd + 1), suffix: xml.slice(openEnd + 1), sectPrXml: '', realChildren: [] };
  }
  const closeIdx = xml.lastIndexOf('</w:body>');
  const headerPrefix = xml.slice(0, openEnd + 1);
  const inner = xml.slice(openEnd + 1, closeIdx);
  const suffix = xml.slice(closeIdx);
  const { leading, children } = splitTopLevelElements(inner);
  let sectPrXml = '';
  let realChildren = children;
  if (children.length && children[children.length - 1].tag === 'w:sectPr') {
    sectPrXml = children[children.length - 1].xml;
    realChildren = children.slice(0, -1);
  }
  return { prefix: headerPrefix + leading, suffix, sectPrXml, realChildren };
}

export function buildDocumentXml(prefix, childrenXml, sectPrXml, suffix) {
  return prefix + childrenXml + sectPrXml + suffix;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/xml.test.js`
Expected: PASS (all xml tests).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: parse/build document.xml around body children"
```

---

## Task 6: `words.js` — `countWords`

**Files:**
- Create: `src/core/words.js`
- Test: `test/words.test.js`

- [ ] **Step 1: Write the failing test**

`test/words.test.js`:
```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/words.test.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `src/core/words.js`**

```js
function decodeEntities(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&'); // decode last to avoid double-unescaping
}

// Count words in the visible text of a body child (paragraph or table).
// Run texts are concatenated with no separator because spaces are encoded
// inside <w:t> content; layout breaks (<w:tab/>, <w:br/>, <w:cr/>) become spaces.
export function countWords(xml) {
  const normalized = xml.replace(/<w:(?:tab|br|cr)\b[^>]*?\/?>/g, ' ');
  let text = '';
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(normalized)) !== null) text += decodeEntities(m[1]);
  return text.split(/\s+/).filter(Boolean).length;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/words.test.js`
Expected: PASS (6 passed).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add word counter for paragraphs and tables"
```

---

## Task 7: `meta.js` — sidecar metadata

**Files:**
- Create: `src/core/meta.js`
- Test: `test/meta.test.js`

- [ ] **Step 1: Write the failing test**

`test/meta.test.js`:
```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/meta.test.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `src/core/meta.js`**

```js
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/meta.test.js`
Expected: PASS (4 passed).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add split metadata sidecar (hash, build, read)"
```

---

## Task 8: `split.js` — partition + split into parts

**Files:**
- Create: `src/core/split.js`
- Test: `test/split.test.js`

- [ ] **Step 1: Write the failing test (partition logic first — pure)**

`test/split.test.js`:
```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/split.test.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `src/core/split.js`**

```js
import JSZip from 'jszip';
import { parseDocumentXml, buildDocumentXml } from './xml.js';
import { countWords } from './words.js';
import { META_PATH, buildMeta, hashString } from './meta.js';

const DOC_PATH = 'word/document.xml';

// Greedily pack children into groups whose word counts stay <= maxWords,
// cutting only at child boundaries. A single oversize child forms its own group.
export function partitionChildren(realChildren, maxWords) {
  const groups = [];
  let cur = [];
  let count = 0;
  for (const c of realChildren) {
    const wc = countWords(c.xml);
    if (cur.length > 0 && count + wc > maxWords) {
      groups.push(cur);
      cur = [];
      count = 0;
    }
    cur.push(c);
    count += wc;
  }
  if (cur.length) groups.push(cur);
  return groups;
}

const baseName = (filename) => filename.replace(/\.docx$/i, '');
const pad = (n, width) => String(n).padStart(width, '0');

// Split a .docx (ArrayBuffer | Uint8Array) into parts.
// Returns { parts: [{ name, bytes }], total, summary: [{ part, words }] }.
export async function splitDocx(data, filename, maxWords) {
  if (!(maxWords > 0)) throw new Error('Max words must be a positive number.');
  const source = await JSZip.loadAsync(data);
  const docFile = source.file(DOC_PATH);
  if (!docFile) throw new Error("This file isn't a valid .docx (no document.xml).");
  const documentXml = await docFile.async('string');
  const { prefix, suffix, sectPrXml, realChildren } = parseDocumentXml(documentXml);
  const groups = partitionChildren(realChildren, maxWords);
  const total = groups.length;
  const hash = hashString(documentXml);
  const base = baseName(filename);
  const width = Math.max(2, String(total).length);

  const parts = [];
  const summary = [];
  for (let i = 0; i < groups.length; i++) {
    const childrenXml = groups[i].map((c) => c.xml).join('');
    const partDocXml = buildDocumentXml(prefix, childrenXml, sectPrXml, suffix);
    const z = await JSZip.loadAsync(data); // fresh copy of all original parts
    z.file(DOC_PATH, partDocXml);
    z.file(META_PATH, buildMeta({ origin: filename, part: i + 1, total, hash }));
    const bytes = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    parts.push({ name: `${base}_part${pad(i + 1, width)}.docx`, bytes });
    summary.push({ part: i + 1, words: groups[i].reduce((a, c) => a + countWords(c.xml), 0) });
  }
  return { parts, total, summary };
}

// Bundle parts into one zip for download.
export async function zipParts(parts, zipName) {
  const zip = new JSZip();
  for (const p of parts) zip.file(p.name, p.bytes);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { name: zipName, bytes };
}
```

- [ ] **Step 4: Run partition tests**

Run: `npx vitest run test/split.test.js`
Expected: PASS (3 partition tests).

- [ ] **Step 5: Add failing integration tests for `splitDocx`**

Append to `test/split.test.js`:
```js
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
});
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx vitest run test/split.test.js`
Expected: PASS (all split tests).

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: split .docx into parts with metadata"
```

---

## Task 9: `merge.js` — reassemble parts losslessly

**Files:**
- Create: `src/core/merge.js`
- Test: `test/merge.test.js`

- [ ] **Step 1: Write the failing test (round-trip + validation)**

`test/merge.test.js`:
```js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/merge.test.js`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `src/core/merge.js`**

```js
import JSZip from 'jszip';
import { parseDocumentXml, buildDocumentXml } from './xml.js';
import { META_PATH, readMetaJson } from './meta.js';

const DOC_PATH = 'word/document.xml';

// files: Array<{ name, data }>, where data is an ArrayBuffer | Uint8Array of a part.
// Returns { name, bytes } — a single .docx whose internal parts equal the original.
export async function mergeDocx(files) {
  if (!files || files.length === 0) throw new Error('No files to merge.');

  const loaded = [];
  for (const f of files) {
    const zip = await JSZip.loadAsync(f.data);
    const metaFile = zip.file(META_PATH);
    if (!metaFile) throw new Error(`"${f.name}" is not a split part made by this tool.`);
    const meta = readMetaJson(await metaFile.async('string'));
    const docFile = zip.file(DOC_PATH);
    if (!docFile) throw new Error(`"${f.name}" is missing document.xml.`);
    loaded.push({ name: f.name, meta, documentXml: await docFile.async('string'), zip });
  }

  const { origin, hash, total } = loaded[0].meta;
  for (const l of loaded) {
    if (l.meta.hash !== hash || l.meta.origin !== origin) {
      throw new Error(`"${l.name}" belongs to a different document.`);
    }
  }

  const byIndex = new Map();
  for (const l of loaded) {
    if (byIndex.has(l.meta.part)) throw new Error(`Part ${l.meta.part} appears more than once.`);
    byIndex.set(l.meta.part, l);
  }
  for (let i = 1; i <= total; i++) {
    if (!byIndex.has(i)) throw new Error(`Missing part ${i} of ${total}.`);
  }

  const ordered = [];
  for (let i = 1; i <= total; i++) ordered.push(byIndex.get(i));

  const first = parseDocumentXml(ordered[0].documentXml);
  let childrenXml = '';
  for (const l of ordered) {
    childrenXml += parseDocumentXml(l.documentXml).realChildren.map((c) => c.xml).join('');
  }
  const mergedDocXml = buildDocumentXml(first.prefix, childrenXml, first.sectPrXml, first.suffix);

  const outZip = ordered[0].zip; // all non-document parts are byte-identical to the original
  outZip.file(DOC_PATH, mergedDocXml);
  outZip.remove(META_PATH);
  const bytes = await outZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { name: origin, bytes };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/merge.test.js`
Expected: PASS (all merge tests).

- [ ] **Step 5: Run the full suite**

Run: `npm test`
Expected: ALL tests pass.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: merge split parts back into the original losslessly"
```

---

## Task 10: UI — Split tab

**Files:**
- Modify: `index.html`, `src/main.js`, `src/styles.css`

- [ ] **Step 1: Replace `index.html` body markup**

Replace the `<body>` contents of `index.html` with:
```html
<body>
  <main id="app">
    <header>
      <h1>DOCX Splitter &amp; Merger</h1>
      <p class="privacy">🔒 100% in your browser. No uploads, nothing leaves your device. Works offline.</p>
    </header>
    <nav class="tabs">
      <button id="tab-split" class="tab active">Split</button>
      <button id="tab-merge" class="tab">Merge</button>
    </nav>

    <section id="panel-split" class="panel">
      <label>Max words per part
        <input id="max-words" type="number" min="1" value="1000" />
      </label>
      <div id="drop-split" class="dropzone">Drop a .docx here, or click to choose</div>
      <input id="file-split" type="file" accept=".docx" hidden />
      <p id="split-error" class="error" hidden></p>
      <ul id="split-summary" class="summary"></ul>
    </section>

    <section id="panel-merge" class="panel" hidden>
      <div id="drop-merge" class="dropzone">Drop the split parts here, or click to choose</div>
      <input id="file-merge" type="file" accept=".docx" multiple hidden />
      <p id="merge-error" class="error" hidden></p>
      <p id="merge-status" class="status"></p>
    </section>
  </main>
  <script type="module" src="/src/main.js"></script>
</body>
```

- [ ] **Step 2: Implement `src/main.js` (tabs, shared helpers, split flow)**

```js
import './styles.css';
import { splitDocx, zipParts } from './core/split.js';
import { mergeDocx } from './core/merge.js';

const $ = (id) => document.getElementById(id);

function download(name, bytes, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function showError(el, msg) {
  el.textContent = msg;
  el.hidden = !msg;
}

// --- Tabs ---
function activate(which) {
  $('tab-split').classList.toggle('active', which === 'split');
  $('tab-merge').classList.toggle('active', which === 'merge');
  $('panel-split').hidden = which !== 'split';
  $('panel-merge').hidden = which !== 'merge';
}
$('tab-split').addEventListener('click', () => activate('split'));
$('tab-merge').addEventListener('click', () => activate('merge'));

// --- Drop zone wiring (shared) ---
function wireDrop(zoneId, inputId, onFiles) {
  const zone = $(zoneId);
  const input = $(inputId);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    onFiles([...e.dataTransfer.files]);
  });
  input.addEventListener('change', () => onFiles([...input.files]));
}

// --- Split flow ---
wireDrop('drop-split', 'file-split', async (files) => {
  showError($('split-error'), '');
  $('split-summary').innerHTML = '';
  const file = files[0];
  if (!file) return;
  const maxWords = Number($('max-words').value);
  try {
    const data = await file.arrayBuffer();
    const { parts, total, summary } = await splitDocx(data, file.name, maxWords);
    const base = file.name.replace(/\.docx$/i, '');
    const { name, bytes } = await zipParts(parts, `${base}_parts.zip`);
    download(name, bytes, 'application/zip');
    $('split-summary').innerHTML =
      `<li><strong>${total} part(s)</strong> — downloaded ${name}</li>` +
      summary.map((s) => `<li>Part ${s.part}: ${s.words} words</li>`).join('');
  } catch (err) {
    showError($('split-error'), err.message);
  }
});
```

- [ ] **Step 3: Add `src/styles.css`**

```css
:root { font-family: system-ui, sans-serif; color: #1a1a1a; }
#app { max-width: 640px; margin: 2rem auto; padding: 0 1rem; }
h1 { margin-bottom: 0.25rem; }
.privacy { color: #0a7d3c; font-size: 0.9rem; margin-top: 0; }
.tabs { display: flex; gap: 0.5rem; margin: 1rem 0; }
.tab { padding: 0.5rem 1rem; border: 1px solid #ccc; background: #f5f5f5; cursor: pointer; border-radius: 6px; }
.tab.active { background: #1a1a1a; color: #fff; }
.dropzone { border: 2px dashed #aaa; border-radius: 10px; padding: 2.5rem 1rem; text-align: center; color: #555; cursor: pointer; margin: 1rem 0; }
.dropzone.over { border-color: #1a73e8; background: #eef4ff; }
label { display: block; margin-bottom: 0.5rem; }
input[type="number"] { width: 8rem; padding: 0.3rem; }
.error { color: #c0392b; }
.summary { padding-left: 1.2rem; }
.status { color: #0a7d3c; }
```

- [ ] **Step 4: Manual verification**

Run: `npm run dev`
Open the printed URL. On the Split tab: set max words to 2, drop a real `.docx`, confirm a `_parts.zip` downloads and the summary lists parts with word counts.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: split tab UI"
```

---

## Task 11: UI — Merge tab

**Files:**
- Modify: `src/main.js`

- [ ] **Step 1: Append the merge flow to `src/main.js`**

```js
// --- Merge flow ---
wireDrop('drop-merge', 'file-merge', async (files) => {
  showError($('merge-error'), '');
  $('merge-status').textContent = '';
  if (!files.length) return;
  try {
    const payload = await Promise.all(
      files.map(async (f) => ({ name: f.name, data: await f.arrayBuffer() })),
    );
    const { name, bytes } = await mergeDocx(payload);
    download(
      name,
      bytes,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    );
    $('merge-status').textContent = `Merged ${files.length} part(s) → ${name}`;
  } catch (err) {
    showError($('merge-error'), err.message);
  }
});
```

- [ ] **Step 2: Manual verification — the full round trip**

Run: `npm run dev`
1. Split a real `.docx` into several parts (Task 10).
2. Unzip the downloaded `_parts.zip`.
3. On the Merge tab, drop all the parts. Confirm the original filename downloads and the status shows the part count.
4. Open the merged file in Word and confirm it matches the original (content + formatting).
5. Negative check: drop only some parts → see a "Missing part N" error.

- [ ] **Step 3: Commit**

```bash
git add -A
git commit -m "feat: merge tab UI"
```

---

## Task 12: Deploy to GitHub Pages + README

**Files:**
- Create: `.github/workflows/deploy.yml`, `README.md`

- [ ] **Step 1: Create `.github/workflows/deploy.yml`**

```yaml
name: Deploy to GitHub Pages

on:
  push:
    branches: [main]
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: true

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 20
      - run: npm ci
      - run: npm test
      - run: npm run build
      - uses: actions/upload-pages-artifact@v3
        with:
          path: dist
  deploy:
    needs: build
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - id: deployment
        uses: actions/deploy-pages@v4
```

- [ ] **Step 2: Create `README.md`**

```markdown
# DOCX Splitter & Merger

Split a Word `.docx` by **word count** (cutting only at paragraph breaks) and
**merge the parts back** into a byte-equivalent original. 100% client-side —
files never leave your browser.

## Use it
Hosted on GitHub Pages. Split tab: choose a max word count and drop a `.docx`;
download a zip of the parts. Merge tab: drop the parts back to rebuild the original.

## Develop
```bash
npm install
npm run dev     # local dev server
npm test        # run the test suite
npm run build   # production build to dist/
```

## How it stays lossless
All parts come from one original document, so styles, numbering, and
relationships are already identical — no reconciliation is needed. The splitter
slices the body's paragraphs as verbatim XML substrings and never rewrites
untouched markup, so merging reproduces every internal document part exactly.
```

- [ ] **Step 3: Verify build succeeds locally**

Run: `npm run build`
Expected: `dist/` is created with `index.html` and bundled assets, no errors.

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "ci: deploy to GitHub Pages + README"
```

- [ ] **Step 5: Post-deploy manual step (human)**

In the GitHub repo: Settings → Pages → Source = "GitHub Actions". Push to `main`, wait for the workflow, open the published URL, and run one split→merge round trip in the live app.

---

## Definition of Done

- [ ] `npm test` passes (xml, words, meta, split, merge, helpers, smoke).
- [ ] Round-trip test proves every internal part is byte-identical after merge.
- [ ] Split tab produces a `_parts.zip`; Merge tab rebuilds the original; verified in Word.
- [ ] Validation errors (missing/duplicate/foreign/non-part) shown clearly.
- [ ] Deployed to GitHub Pages via Actions; live round trip works.

## Risk Note (verify during Task 10/11 manual checks)

Parts carry an extra zip entry `docx-splitter-meta.json` that is not declared in
`[Content_Types].xml`. Word ignores undeclared zip entries, so parts open
normally — **confirm by opening a part in Word during manual verification.** If a
strict environment ever rejects it, the fallback is to declare the metadata as a
proper `docProps/custom.xml` part; this is localized to `split.js`/`merge.js` and
does not affect the merged result (the sidecar is removed on merge).
```
