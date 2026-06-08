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
