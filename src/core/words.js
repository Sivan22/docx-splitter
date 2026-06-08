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
  const normalized = xml.replace(/<w:(?:tab|br|cr)\b[^>]*?\/?>/g, '<w:t> </w:t>');
  let text = '';
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(normalized)) !== null) text += decodeEntities(m[1]);
  return text.split(/\s+/).filter(Boolean).length;
}
