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
