---
id: SPEC-016
title: Cut, copy and paste
status: done
area: editing
depends_on: [SPEC-011, SPEC-015]
---
## Why
Most content arrives by paste — from Word, from a browser, from an email. Paste
is where formatting fidelity is judged, and where a naive implementation
injects junk markup that poisons the saved file.

## Acceptance criteria
- [ ] Ctrl/Cmd+X, Ctrl/Cmd+C and Ctrl/Cmd+V cut, copy and paste, as do the context menu entries
- [ ] Copying puts three flavours on the clipboard: plain text, HTML, and an internal document-model payload
- [ ] Pasting content copied from this editor restores formatting, styles, lists and structure exactly
- [ ] Pasting HTML from Word Online, Google Docs or a web page maps headings, bold/italic/underline, strikethrough, sub/superscript, font family/size/colour and highlight to the document model, and drops scripts, classes, positioning and unknown markup; until their specs land, list items import as plain paragraphs (SPEC-018), hyperlink runs lose only the URL (SPEC-025), and pasted `<table>`s flatten to tab-separated rows (SPEC-022) — all three remapped when the later spec merges, per Reference
- [ ] Ctrl/Cmd+Shift+V pastes as plain text, taking the formatting of the paragraph at the caret
- [ ] A paste options chip appears after a paste with "Keep source formatting", "Match destination formatting" and "Keep text only"; choosing one re-applies the paste as a single undo step
- [ ] Pasting into a selection replaces it; pasting multiple paragraphs into the middle of a paragraph splits it correctly
- [ ] Cut removes the content and places it on the clipboard in one undo step
- [ ] Pasted content is sanitised: no `javascript:` URLs, event handlers, `<script>`, `<style>` or embedded remote content survives into the model — covered by an explicit security test
- [ ] Pasting an image from the clipboard is accepted once SPEC-024 is merged, and until then shows "Images aren't supported yet"

## Out of scope
Drag-and-drop of selected text within the document. Clipboard history. Paste of
Excel ranges as live-linked tables.

## Reference

### Target document model (inline, normative here)

The paste/import pipeline's output type **is the model in
`src/lib/documentModel.ts`**, copied below. Later specs (listed in the "HTML
mapping" table) extend these nodes; until they merge, an HTML feature listed as
deferred is imported at the "text level" — its characters and character
formatting survive, its structure collapses to plain paragraphs.

```ts
// src/lib/documentModel.ts — canonical. Block = Paragraph until SPEC-022.
export interface RunFormatting {
  bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean;
  fontFamily?: string;             // e.g. "Calibri"
  fontSizePt?: number;             // points
  color?: string;                  // CSS-ready colour, e.g. "#ff0000"
  highlight?: string | false;      // OOXML colour name | "#rrggbb" | false = "No colour"
  shadingFill?: string;            // preservation slot, never set by paste
  vertAlign?: "subscript" | "superscript";   // mutually exclusive single field
}
export interface Run { text: string; formatting: RunFormatting }
export interface ParagraphFormatting {
  alignment?: "left" | "center" | "right" | "justify";
  indentStartCm?: number; indentEndCm?: number; indentFirstLineCm?: number;
  spaceBeforePt?: number; spaceAfterPt?: number;
  lineSpacing?: number; lineSpacingPt?: number; lineSpacingRule?: "exact";
  direction?: "ltr" | "rtl";
}
export interface Paragraph {
  type: "paragraph";
  runs: Run[];                     // may be empty; never zero-length runs
  formatting?: ParagraphFormatting;
  styleId?: string;                // raw style ref when the source names one
  headingLevel?: number;           // 1-based, undefined for body text
}
export type Block = Paragraph;     // tables (SPEC-022) extend this later
```

Editing surface (SPEC-011, `src/lib/textEditing.ts`): mutations take
`(model, caret, …)` and return the new caret — the DOM is never the source of
truth. Paste uses `insertText`, `deleteRange(a, b)`, `splitParagraph` and
`deleteBackward/Forward`. SPEC-015: paste and cut each call `recordEdit` once,
so each is a single undo step; the undo test compares with `equalStructural`.

**There is deliberately nothing in the model for lists, tables or hyperlinks.**
`headingLevel`/`styleId` carry headings; character formatting carries
bold/italic/underline/strike/sub/sup/colour/size/highlight. Lists (SPEC-018),
tables (SPEC-022), links (SPEC-025) and images (SPEC-024) extend these types in
their own PRs — and each such PR must amend this spec's HTML-mapping table
(one normative source, mirroring SPEC-022's note below), rather than leave two.

### Clipboard flavours and how paste detects the internal payload

All three flavours are written in the synchronous `copy`/`cut` handlers via
`ClipboardEvent.clipboardData.setData(...)` and read in the `paste` handler via
`getData(...)`. This is the only API that can carry a private MIME type —
`navigator.clipboard.read()/write()` accept only a fixed allowlist, so it is
never used here.

1. `text/plain` — linearised text: paragraphs joined with `\n`. No formatting.
2. `text/html` — a minimal, self-generated fragment for interop with Word,
   Google Docs and mail clients, plus this editor's own import when the
   internal payload is missing (e.g. the clipboard was written by another app,
   or the browser dropped the custom type): `<p>`, `<h1>`–`<h6>`,
   `<b>/<i>/<u>/<s>/<sub>/<sup>`, `<span>` carrying only
   `color`/`font-family`/`font-size` in `style`, `<ul>/<ol><li>`, `<table>`,
   `<a href>`. No `class` attributes, no layout markup, nothing the model
   cannot map back (table).
3. `application/x-docu-edit-docmodel+json` — the internal payload, checked
   first so editor-to-editor paste is exact. Exact format, labelled with the
   MIME type as the key in `setData`/`getData`:
   ```jsonc
   { "mimeType": "application/x-docu-edit-docmodel+json",
     "version": 1,
     "blocks": [ /* the copied Paragraph[] verbatim from the model */ ] }
   ```
   - **Identification:** paste detects it by calling
     `e.clipboardData.getData("application/x-docu-edit-docmodel+json")`; a
     non-empty string means "produced by this editor". Parse, then validate
     `mimeType` matches the key **and** `version === 1`; anything else falls
     through to `text/html`. The `mimeType` field is defence-in-depth, not the
     detection mechanism.
   - **Content:** copy/cut captures the `Paragraph[]` spanning the selection
     (split boundary runs at the selection edges, run-merge via the model's
     normalisation so the fragment obeys the model invariants). Document-scoped
     data (`id`, `title`, `sections`, `defaultStyles`, opaque parts) is
     excluded.
   - **Insert:** `structuredClone` the fragment out of the clipboard cache,
     delete the selection, then splice the blocks in on a fresh selection that
     spans the paste, driving the existing `splitBlock`/`insertText` machinery
     so caret placement, rendering and undo all work as for typing.

### HTML → model mapping (Word Online, Google Docs, web)

Whitelist import; anything not in this table is dropped (Sanitisation below).

| Source HTML | Model form |
| --- | --- |
| `<h1>`…`<h6>` | `Paragraph` with `headingLevel` 1–6 |
| `<b>`/`<strong>`, `<i>`/`<em>`, `<u>`, `<s>`/`<strike>`/`<del>` | `bold`/`italic`/`underline`/`strikethrough: true` on the run |
| `<sub>`, `<sup>` | `vertAlign: "subscript"` / `"superscript"` |
| `<span style="color / font-family / font-size">`, `<font color face size>` | `color`, `fontFamily`, `fontSizePt` (HTML px/pt converted to Pt) |
| `<p>`, block `<div>` | a `Paragraph` per block; `text-align` → `alignment`, `margin-top/bottom`, `line-height`, `text-indent`, `padding-left/right` → the matching `ParagraphFormatting` field |
| `<ul>`/`<ol>`/`<li>` | **until SPEC-018:** each `<li>` becomes a plain `Paragraph` (runs and character formatting kept, list identity dropped) |
| `<table>` | **until SPEC-022:** one `Paragraph` per row, cell text joined with `\t`, each cell's inline content mapped to runs (see AC) |
| `<a href>` | the anchor's runs with character formatting; the URL is discarded (the model has no link field) — links become full features under SPEC-025 |
| `<br>` | a literal `"\n"` inside the current run (forced line break; `white-space: pre-wrap` renders it) |

### Internal paste and the three chip options — precise meanings

Internal paste performs, by default, *exactly* what "Keep source formatting"
means — the fragment is inserted verbatim, which is why the AC says it
"restores formatting, styles, lists and structure exactly". The chip therefore
re-interprets the same paste rather than leaving the default as an
alternative:

- **Keep source formatting** — the verbatim insert already applied; choosing
  it is a confirmation no-op.
- **Match destination formatting** — delete the just-pasted blocks and
  re-import the same fragment with formatting re-resolved to the destination:
  each block's `formatting`, `styleId` and `headingLevel` are dropped, and it
  takes the destination paragraph's paragraph formatting; each run is stripped
  of character formatting and re-set through the caret's
  `PendingFormat`/run-formatting resolution (SPEC-011/012), i.e. the formatting
  a freshly typed character would have. Paragraph boundaries and run text are
  unchanged; deferred structure (lists/tables, once SPEC-018/022 land)
  collapses to its destination form.
- **Keep text only** — discard blocks and all formatting; feed the
  concatenated run text (paragraph breaks kept) through the plain-text paste
  path, taking the destination paragraph's formatting — identical to
  Ctrl/Cmd+Shift+V.

The paste and the subsequent re-apply land as **one undo step** (SPEC-015's
"one intention"); `equalStructural` is the test relation. The same three
options apply to the external-HTML and plain-text paths with the same meaning
("Keep source formatting" there means the HTML import's own fidelity).

### Sanitisation — what "embedded remote content" means

Every pasted HTML node goes through a whitelist import (only the table above
maps); anything else is dropped, and a security test asserts none of the
following survives into the model:

- **Active content:** `<script>`, `<style>`, `<link rel="stylesheet">`,
  `<iframe>`, `<object>`, `<embed>`, `<video>/<audio>/<source>/<track>`,
  `<form>/<button>/<input>`, and `<svg>`/`<math>` (and all descendants) —
  dropped wholesale.
- **Attributes:** every `on*` handler, `srcdoc`, `background`, and any
  `url(...)`, `javascript:`, `vbscript:` or `data:` value in a style or
  attribute. A `javascript:` link is stripped anywhere (SPEC-025 policy).
- **External images** (`<img src="http(s)://…">`) — yes, these count as
  embedded remote content. The model stores no image at all until SPEC-024, and
  SPEC-024 explicitly excludes linked/remote images, so a remote `<img>` must
  never be preserved as a live reference. Before SPEC-024, every `<img>` (remote
  or `data:`) takes the "Images aren't supported yet" path; after SPEC-024 only
  images already imported as media parts may be pasted.

The security test drives a clipboard payload containing a `<script>`, a
`<style>` with an `onclick`, an `<iframe>`, a remote-URL `<img>` and a
`javascript:` link, and asserts none of them appears in the resulting model.

### Dependency contracts (restated per SPEC-000 rule 2)

- **SPEC-011** (`src/lib/textEditing.ts`): `insertText`, `deleteRange(a, b)`,
  `splitParagraph`, plus the caret coordinate type
  `{ section, block, offset }`. Model is plain JSON; `Paragraph` shape above.
- **SPEC-015** (`src/lib/undoHistory.ts`): paste and cut each call `recordEdit`
  once; the chip's re-apply shares that step; `equalStructural` is the undo
  test relation. `MAX_HISTORY_STEPS` is 200.
- Equality: `equalStructural` from `src/lib/equality.ts`.
