---
id: SPEC-016
title: Cut, copy and paste
status: draft
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
- [ ] Pasting HTML from Word Online, Google Docs or a web page maps headings, bold/italic/underline, lists and links to the document model, and drops scripts, classes, positioning and unknown markup; pasted `<table>`s are flattened to tab-separated rows until SPEC-022 models tables (see Reference)
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

Depend on SPEC-011's command API (`insertText`, `deleteRange`, `splitBlock`,
each returning a new model state — the DOM is never the source of truth) and
SPEC-015 (every paste/cut is one undo step; `equalStructural` is the undo
test). The canonical model and equality relations live in
`src/lib/documentModel.ts` and `src/lib/equality.ts`.

### Clipboard flavours and the internal payload

All three flavours are written in the synchronous `copy`/`cut` handlers via
`ClipboardEvent.clipboardData.setData(...)` and read in the `paste` handler via
`getData(...)`. This is the only API that can carry a private MIME type —
`navigator.clipboard.read()/write()` accept only a fixed allowlist, so it is
never used here. Paste order: internal payload, then `text/html`, then
`text/plain`.

1. `text/plain` — linearised text: paragraphs joined with `\n`, a tab between
   cells of any flattened table cell (see below). No formatting.
2. `text/html` — a minimal, self-generated fragment for interop with Word,
   Google Docs and mail clients, plus this editor's own import when the
   internal payload is absent: `<p>`, `<h1>`–`<h6>`, `<b>/<i>/<u>/<s>`, `<span>`
   carrying only `color`/`font-family`/`font-size` in `style`, `<ul>/<ol><li>`,
   and `<a href>` once SPEC-025 lands. No `class` attributes, no layout
   markup, nothing the model cannot map back.
3. `application/x-docu-edit-docmodel+json` — the internal payload, checked
   first so editor-to-editor paste is exact: a JSON serialisation of the copied
   `Block[]` (a `Paragraph[]`, in current-model shape with runs and formatting
   intact — parse-free, no HTML round-trip). Document-scoped data (`id`,
   `title`, `sections`, `defaultStyles`, opaque parts) is excluded; a
   `Paragraph[]` fragment stands alone. Format:
   ```jsonc
   { "mimeType": "application/x-docu-edit-docmodel+json",
     "version": 1,
     "blocks": [ /* Paragraph[] copied verbatim from the model */ ] }
   ```
   If `getData` returns it and `version` is supported, insert
   `blocks` in place of the selection directly — this is what AC "restores
   formatting, styles, lists and structure exactly" is tested against.

### Tables

The model has no table block until SPEC-022 (`Block = Paragraph` today;
SPEC-008 keeps a source file's `w:tbl`/`table:table` only as an opaque node for
same-format round-trips, which never applies to pasted content — a pasted HTML
`<table>` is not the target format's XML). So until SPEC-022 merges, a pasted
`<table>` is imported losslessly at the text level: one paragraph per row,
cells joined by a tab, each cell's inline content mapped to runs with its
character formatting — the same representation Word and Google Docs put in
`text/plain`. When SPEC-022 lands it must amend this spec's table acceptance
criterion (same pattern as SPEC-024 for images) rather than leave two sources
of truth; "Keep source formatting" for a table behaves as "Keep text only"
until then.

### Sanitisation — what "embedded remote content" means

Every pasted HTML node goes through a whitelist import (only the elements listed
above are mapped); anything else is dropped, and a security test asserts none of
the following survives into the model:

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
