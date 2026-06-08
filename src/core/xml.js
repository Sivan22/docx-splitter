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

// Decompose word/document.xml so the body's top-level children can be sliced.
// Invariant: prefix + realChildren.map(c=>c.xml).join('') + sectPrXml + suffix === xml
export function parseDocumentXml(xml) {
  // Locate the real <w:body> start tag (not a prefix sibling like <w:bodyPr>).
  let bodyOpen = -1;
  for (let from = 0; ; ) {
    const idx = xml.indexOf('<w:body', from);
    if (idx === -1) break;
    const after = xml[idx + '<w:body'.length];
    if (after === undefined || after === '>' || after === '/' || /\s/.test(after)) {
      bodyOpen = idx;
      break;
    }
    from = idx + '<w:body'.length;
  }
  if (bodyOpen === -1) {
    return { prefix: xml, suffix: '', sectPrXml: '', realChildren: [] };
  }
  // Use the quote-aware tag scanner to find the end of the <w:body ...> open tag.
  const bodyTag = readTag(xml, bodyOpen);
  if (bodyTag.kind === 'selfclose') {
    // self-closing <w:body/> — no children
    return { prefix: xml.slice(0, bodyTag.end), suffix: xml.slice(bodyTag.end), sectPrXml: '', realChildren: [] };
  }
  const openEnd = bodyTag.end; // index just past the '>' of <w:body ...>
  const closeIdx = xml.lastIndexOf('</w:body>');
  const headerPrefix = xml.slice(0, openEnd);
  const inner = xml.slice(openEnd, closeIdx);
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
