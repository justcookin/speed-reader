# Speed Reader — Claude Code Session Notes

## Session History

### February 2026 — Initial build
Built from scratch in one session. Started with RSVP-only, then added Guide mode, structured PDF
extraction, multi-word RSVP chunking, and highlight transition polish.

---

## What This App Is

A browser-based speed reader (`index.html` + `style.css`, no build step, no dependencies except CDN scripts).
Supports PDF, PPTX, and TXT files. Two reading modes: **RSVP** and **Guide**.

Intentionally a proof of concept — no framework, no build pipeline, no backend. The goal was to
replicate the core mechanics of apps like Spreeder and Outread to understand where their value
actually lies.

---

## Architecture

### State variables
```js
let words        = [];   // flat word array — drives RSVP and word count
let blocks       = [];   // structured blocks — drives Guide mode rendering
let currentIndex = 0;    // index into words[]
let wpm          = 300;
let intervalId   = null;
let mode         = 'rsvp';   // 'rsvp' | 'guide'
let rsvpChunk    = 1;        // words per tick in RSVP (1–7)
let chunkSize    = 3;        // words per tick in Guide (1–10)
```

### Block structure
Each extraction function returns `Block[]`:
```js
{ type: 'paragraph' | 'heading' | 'bullet', text: string }
```
`processBlocks()` assigns `block.startIndex`, `block.words`, and builds the flat `words[]` array.
The `data-i` indices on guide text spans are sequential across all blocks.

### Timing
```js
function intervalMs() {
  const c = mode === 'rsvp' ? rsvpChunk : chunkSize;
  return (c * 60000) / wpm;
}
```
More words per chunk → longer display time → WPM stays constant.

---

## RSVP Mode

### ORP (Optimal Recognition Point)
```js
function orpIndex(word) {
  return Math.ceil(word.length / 2) - 1;
}
```
The pivot character is highlighted red; letters before it go in `#word-prefix`, after in `#word-suffix`.

### Multi-word chunks (rsvpChunk > 1)
One ORP character total per chunk — from the **centre word**:
```js
const pivotIdx = Math.floor(chunk.length / 2);
```
- 1 word → index 0; 2 words → index 1; 3 words → index 1 (middle); 5 → index 2; 7 → index 3.
- Words before pivot + pivot's prefix → `wordPrefix.textContent`
- Single ORP char → `wordOrp.textContent`
- Pivot's suffix + words after → `wordSuffix.textContent`

**Critical**: use `textContent`, not `innerHTML`, for the three spans. Spaces are embedded in the
string content itself (`before + ' ' + pivotPrefix`, `pivotSuffix + ' ' + after`). `innerHTML`
with `.join(' ')` produces space text-nodes that flex containers can collapse.

**Detached span guard**: if a previous `innerHTML` call on `#word-display` detached the fixed spans,
re-attach them before setting textContent:
```js
if (!wordPrefix.isConnected) {
  wordDisplay.textContent = '';
  wordDisplay.append(wordPrefix, wordOrp, wordSuffix);
}
```

### Font scaling
```js
wordDisplay.style.fontSize =
  rsvpChunk <= 2 ? '2.6rem' : rsvpChunk <= 4 ? '1.9rem' : '1.4rem';
```

---

## Guide Mode

### How it works
Full text is rendered as indexed `<span data-i="N">` elements. A chunk of words is highlighted
at a time; the reader **pages** (snaps `scrollTop`) rather than smoothly scrolling — smooth
scrolling was jarring.

### Paging logic
```js
const spanBottom = firstSpan.offsetTop + firstSpan.offsetHeight;
const viewBottom = readerEl.scrollTop + readerEl.clientHeight;
if (spanBottom > viewBottom) {
  readerEl.scrollTop = firstSpan.offsetTop - 20;
}
```
`offsetTop` works correctly here because `#reader` has `position: relative` and `display: block`
(in guide mode), making it the `offsetParent` for all descendant spans.

### Why `display: block` on `#reader.guide-mode` (critical)
`#reader` is normally a flex container. In flex layout, children have `flex-shrink: 1` by default,
so `#guide-text` was being **shrunk to fit** the 280px box — nothing to scroll. Switching to
`display: block` lets `#guide-text` expand to its full content height.

### Scrollbar hiding
```css
#reader.guide-mode { overflow-y: scroll; scrollbar-width: none; }
#reader.guide-mode::-webkit-scrollbar { display: none; }
```
`overflow-y: scroll` (not `hidden`) is needed so `scrollTop` is reliably writable.

### Highlight transition
```css
#guide-text span {
  transition: color 0.15s ease-out, background 0.15s ease-out, text-shadow 0.15s ease-out;
  text-shadow: 0 0 0 rgba(88, 166, 255, 0);  /* must match syntax of active state */
}
#guide-text span.active-chunk {
  color: var(--text);
  background: rgba(88, 166, 255, 0.18);
  text-shadow: 0 0 12px rgba(88, 166, 255, 0.45);
}
```
Transitioning from `none` to a shadow value doesn't animate — both states need matching shadow
syntax (same number of values, just opacity goes 0 → non-zero).

### Structured block rendering
```js
switch (block.type) {
  case 'heading': return `<div class="guide-heading">${inner}</div>`;
  case 'bullet':  return `<div class="guide-bullet"><span class="bullet-marker">•</span><span class="bullet-body">${inner}</span></div>`;
  default:        return `<p class="guide-para">${inner}</p>`;
}
```

---

## PDF Extraction

PDF.js `getTextContent()` returns items with `transform[4]` (x), `transform[5]` (y), `height`
(font size). PDF y=0 is **bottom** of page; sort descending to get reading order.

### Line grouping
Group items whose y-coordinates are within ±2 units:
```js
const existing = lines.find(l => Math.abs(l.y - y) <= 2);
```

### Paragraph detection
```js
const medianGap     = gaps[Math.floor(gaps.length / 2)] || 12;
const paraThreshold = medianGap * 1.8;  // gap > 1.8× median = new paragraph
```

### Heading detection
```js
const isHeading = line.maxFS > medianFS * 1.2;  // 20% larger than median font size
```

### Bullet detection
```js
const BULLET_RE = /^[\u2022\u2023\u25E6\u25AA\u25CF\u25CB\u2043•·▪▸◦]\s*|^[-*]\s+/;
```
Strip the matched bullet prefix from the text before creating the block.

### Limitations
- Single-column PDFs only (multi-column will mix columns)
- Font size from `item.height || Math.abs(item.transform[3])` — not always accurate
- Scanned/image PDFs return no text items

---

## PPTX Extraction

Uses PML namespace for shape detection:
```js
const PML = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const ph  = sp.getElementsByTagNameNS(PML, 'ph')[0];
const isTitle = ph && (ph.getAttribute('type') === 'title' || ph.getAttribute('type') === 'ctrTitle');
```
Title placeholder → heading block; all other shapes → paragraph block.

---

## TXT Extraction

Split on `\n{2,}` for paragraphs. If every line in a section starts with a bullet char, each
line becomes its own bullet block. Otherwise lines are joined into a single paragraph.

---

## Chunk Controls (shared UI)

Same `#chunk-controls` UI serves both modes. `setChunkSize(delta)` reads `mode` to update
the right variable (`rsvpChunk` vs `chunkSize`) and recalculates the interval if playing.
`updateChunkDisplay()` reads the current mode's chunk size for the label.

---

## Known Gotchas & Decisions

| Gotcha | Resolution |
|--------|-----------|
| Flex container shrinks `#guide-text` to 280px | `display: block` on `#reader.guide-mode` |
| `overflow: hidden` prevents reliable `scrollTop` | Use `overflow-y: scroll` + hide scrollbar via CSS |
| `offsetTop` unreliable in flex layout | Fixed by switching reader to block layout; `offsetParent` is now unambiguously `#reader` |
| CSS shadow transition from `none` doesn't animate | Declare matching `text-shadow: 0 0 0 rgba(…, 0)` on inactive state |
| `innerHTML` on `#word-display` detaches fixed span refs | Guard with `wordPrefix.isConnected`; re-attach with `wordDisplay.append()` |
| `.join(' ')` spaces collapse in flex containers | Use `textContent` with spaces embedded in strings |
| RSVP multi-word pivot on wrong word | `Math.floor(chunk.length / 2)` gives true centre for all sizes |
| PDF y=0 is page bottom | Sort lines descending by y for reading order |

---

## Product & Competitive Learnings

### What this proves
The core reading mechanics of commercial speed readers — RSVP with ORP, Guide mode with pacing,
structured text extraction — fit in ~600 lines of vanilla JS. The engine itself is essentially
a commodity. This was built in a single session to near feature-parity with the core reader in
apps that charge subscription fees.

### Where Spreeder / Outread / Readwise Reader actually earn their keep
These are the hard parts this proof of concept does NOT have:

- **Library & sync** — iCloud/Pocket/Instapaper/Readwise integration, reading position persisted
  across devices. Genuinely hard infrastructure.
- **Mobile native feel** — gesture controls, background fetch, offline support, haptics. A web
  app can approximate but not match native.
- **Persistence** — position remembered per book, per-document WPM settings, reading history,
  streak tracking, stats.
- **Content pipelines** — web article stripping (Readability.js), ePub/MOBI parsing, cleaning
  PDF extraction artifacts far beyond our heuristic approach, handling poorly-structured documents.
- **Training & habit products** — Spreeder in particular sells structured speed reading courses,
  not just the tool. The reader is the delivery mechanism.
- **Polish at scale** — edge cases in real-world PDFs (multi-column, tables, footnotes mixed into
  body, right-to-left text, ligatures) that our heuristic extractor mangles.

### The honest limitation of this build
Our PDF extraction will fail or produce garbled output on:
- Multi-column layouts (columns get interleaved)
- Tables (cells merge into nonsense)
- Footnotes (mixed into paragraph text)
- Heavily styled documents where `item.height` doesn't reflect actual font size

Production apps spend enormous effort on these edge cases. That gap is larger than it looks from
the outside.

### The moat is thin — but real
The reading engine is replicable in a weekend. The defensible value is the ecosystem: sync,
library management, content sourcing, and habit-formation product design built up over years.
Same pattern as most SaaS tools when you strip out the core feature.

---

## Potential Next Features (if continuing this project)

- **Persistence** — `localStorage` for WPM, chunk size, last file position
- **ePub support** — JSZip already in the stack; ePub is just a zip of HTML files
- **Web article mode** — paste a URL, strip with Readability.js, read immediately
- **Reading stats** — words read, time spent, WPM history chart
- **Progress memory** — remember position per file (hash the filename as key)
- **Keyboard shortcuts** — arrow keys for WPM, number keys for chunk size
- **Better PDF extraction** — column detection via x-coordinate clustering, footnote suppression
  by ignoring lines with significantly smaller font size than median
- **Mobile layout** — the current UI is desktop-only
