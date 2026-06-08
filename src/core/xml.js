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
