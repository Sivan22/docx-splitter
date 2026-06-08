# DOCX Splitter & Merger — Design

**Date:** 2026-06-08
**Status:** Approved (pending spec review)

## Summary

A static, GitHub Pages–hosted single-page web app that splits a Word (`.docx`)
document into smaller parts by **word count**, cutting only at paragraph breaks,
and can later **merge the parts back into a byte-equivalent copy of the
original**. All processing happens client-side in the browser — files are never
uploaded to any server.

## Goals

- Split one `.docx` into N parts, each at or under a user-defined **max word
  count**, cutting only at the **last paragraph break before the limit**.
- **Lossless:** each part is a valid standalone `.docx`; merging the full set
  reproduces the original document — **every internal zip part** (`document.xml`,
  `styles.xml`, etc.) byte-identical to the original, so it opens identically in
  Word. (The outer zip *container* may differ cosmetically — compression level,
  entry order, timestamps — without affecting document content; see Testing.)
- **Foolproof merge:** parts carry embedded metadata so the app can auto-order
  them and refuse incomplete, duplicated, or foreign sets.
- **Privacy:** 100% client-side, no uploads, works offline.

## Non-Goals (YAGNI)

- Splitting by page count (pages don't exist in `.docx` — Word computes them at
  render time) or by file size (unpredictable due to zip compression).
- Splitting *inside* a paragraph.
- Merging documents that did **not** originate from the same split operation.
- Rich-media remapping work beyond what verbatim part-copying already preserves
  (target documents are "mostly text & paragraphs"; images/tables are preserved
  as-is because untouched XML and shared parts are copied verbatim).
- Any server, account, or backend.

## Key Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Size unit | **Word count** | Deterministic; file-size and page-count are not computable reliably in-browser. |
| Document scope | Mostly text & paragraphs | Simplifies the lossless guarantee; richer content still preserved verbatim. |
| Merge ordering | **Embedded metadata** (origin name, part N of M, content hash) | Auto-orders, detects missing/duplicate/foreign parts; order-independent upload. |
| Split output | **Single ZIP** of all parts | One click; clean for many parts; parts are exactly what Merge expects. |
| Architecture | Client-side, zip-level XML splicing | Only approach that is lossless *by construction* and unit-testable. |
| Split granularity | Top-level body children (`<w:p>`, `<w:tbl>`) | Cut only at paragraph/table boundaries; never rewrite a child's XML. |

## Why This Can Be Lossless Where Generic Mergers Fail

Generic docx mergers (e.g. `docx-merger`) reconcile styles and numbering from
**independently authored** documents, which produces Word numbering errors. Our
parts all descend from **one** original document, so styles, numbering, fonts,
settings, and relationship IDs are already identical across every part. We never
reconcile anything — we partition and reassemble slices of a single document.
That is the foundation of the lossless guarantee.

## Architecture

A static single-page app. No backend. Two tabs (Split / Merge) over a shared set
of pure core functions.

```
┌─────────────────────────────────────────────┐
│  UI layer (index.html + ui.js)               │
│   - Split tab: drop .docx, max-words input   │
│   - Merge tab: drop parts                     │
│   - Privacy banner, file list, errors         │
└──────────────┬──────────────────────────────┘
               │ calls
┌──────────────▼──────────────────────────────┐
│  Core (pure, testable modules)               │
│   docx.js   - open/read/write zip parts       │
│   words.js  - countWords(bodyChild)           │
│   split.js  - splitDocx(file, maxWords)        │
│   merge.js  - mergeDocx(files)                 │
│   meta.js   - read/write embedded part metadata│
└──────────────┬──────────────────────────────┘
               │ uses
        JSZip + browser DOMParser/XMLSerializer
```

## Core Components

### `words.countWords(bodyChildElement) -> number`
Counts words in the concatenated text of all `<w:t>` descendants of a body child
(paragraph or table). Words = runs of non-whitespace separated by whitespace.

### `split.splitDocx(file, maxWords) -> { parts: Blob[], summary }`
1. Unzip; parse `word/document.xml`.
2. Capture the body's closing `<w:sectPr>` (final section properties).
3. Iterate top-level body children in order, accumulating word count:
   - If adding the next child would exceed `maxWords` **and** the current part is
     non-empty, close the current part *before* that child.
   - If a single child alone exceeds `maxWords`, it forms its own over-limit part
     (never split inside it).
4. For each part, build a `.docx` = a verbatim copy of the original zip with
   `word/document.xml`'s body children replaced by that part's slice followed by
   the captured `<w:sectPr>`, plus an embedded metadata part.
5. Bundle all parts into one ZIP named `<origin>_parts.zip`; members named
   `<origin>_partNN.docx` (zero-padded).

### `merge.mergeDocx(files) -> { blob: Blob, origin: string }`
1. Read each file's metadata: origin name, part index, total, content hash.
2. Validate: all share one origin + hash family; indices form a complete
   consecutive `1..M` with no gaps/dupes; reject foreign/incomplete sets with a
   specific error.
3. Order by index. Concatenate the body-children slices (dropping every part's
   appended closing `<w:sectPr>` except restoring the single original one at the
   end), strip metadata, repackage to a single `.docx`.
4. Result has every internal part byte-identical to the original's, so it opens
   identically in Word. (To also make the outer container reproducible, JSZip is
   configured with a fixed compression setting, preserved entry order, and fixed
   timestamps — a best-effort stretch, not a correctness requirement.)

### `meta` (embedded metadata)
Stored as a dedicated zip entry (custom XML part) — **not** inside
`document.xml`, so the document body stays pristine. Fields: `origin` (filename),
`part` (1-based), `total`, `hash` (hash of the original concatenated body /
identifying the split family), `schemaVersion`.

## Data Flow

**Split:** `.docx` → unzip → parse body → partition children by word count →
per-part: clone zip + replace body + add metadata → zip all parts → download.

**Merge:** parts → read metadata → validate completeness/origin → order →
concatenate slices → strip metadata → repackage → download original.

## Error Handling

- Not a `.docx` / corrupt zip → "This file isn't a valid .docx."
- `maxWords` ≤ 0 or non-numeric → inline validation, block split.
- Merge: foreign part (different origin/hash) → name the offending file.
- Merge: missing index (e.g. have 1,2,4 of 4) → "Missing part 3 of 4."
- Merge: duplicate index → "Part 2 appears twice."
- All errors shown inline near the relevant drop zone; nothing throws to console
  silently.

## Testing

Vitest unit tests on the pure core:
- `countWords`: empty, whitespace-only, multi-run paragraphs, tables.
- `splitDocx`: exact-boundary cut, oversize-single-paragraph part, single part
  when under limit, correct metadata on each part.
- `mergeDocx`: happy path; missing/duplicate/foreign-part rejections.
- **Round-trip property:** for `merge(split(doc, k))`, **every internal zip part
  is byte-identical** to the original's, across a set of sample `.docx` fixtures
  and several `k` values. (Outer-container byte-equality is asserted as a
  secondary, best-effort check once reproducible-zip options are in place.)

## Deployment

- Built with Vite (vanilla TS/JS).
- GitHub Actions workflow builds and publishes to GitHub Pages on push to the
  default branch.
- No environment variables, secrets, or runtime services.

## Open Questions

None outstanding. All keystone decisions resolved during brainstorming.
