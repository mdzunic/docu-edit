---
id: SPEC-135
title: Notes, comments and hyperlinks on cells
status: draft
area: editing
complexity: normal
story_points: 8
depends_on: [SPEC-038, SPEC-122, SPEC-124, SPEC-125, SPEC-127, SPEC-128, SPEC-129, SPEC-133]
---
## Why
A workbook is reviewed in its margins: the yellow note on the cell nobody
understands, the Excel 365 thread asking "where does this number come from?",
the cell that links to the contract or to the summary sheet. This product
reads every one of them and shows none. `xl/commentsN.xml`, the VML drawing
that positions each note, `xl/threadedComments/*.xml` and `xl/persons/person.xml`
ride in `model.opaque`; an `.ods` `office:annotation` survives in SPEC-122's
`Cell.preservedXml`; an `.xlsx` `<hyperlinks>` element survives verbatim in the
retained worksheet XML (with SPEC-124's textual `@ref` adjustment); and an `.ods`
`text:a` is flattened to its text by SPEC-116's reader — and, because SPEC-122
re-emits `<text:p>` from the formatted value, silently dropped the first time
its table is rewritten.

The vocabulary already exists elsewhere. SPEC-038 gave the word processor
threads, replies, resolve, the settable author identity
(`src/lib/core/commentIdentity.ts`, already core) and the
`"{author} ({createdAt}): {text}"` reply encoding; SPEC-025 gave it the link
dialog, the link chip (`src/components/LinkChip.vue`, which imports nothing
and is reused here as-is) and the dangerous-scheme rules. What is new is the
anchor — a cell, or for links a range — and the wire: Excel has **two**
comment generations in one package (legacy notes and 365 threads, the second
mirrored into the first as a placeholder), and ODF has notes only.

**This file is the normative source for the cell annotation model (notes,
threads, persons, hyperlinks), its wire mappings in both formats, the
thread-placeholder rule, and the rule for which link targets a sheet opens.**

## Acceptance criteria
- [ ] `Sheet` gains `notes?: CellNote[]`, `threads?: CellThread[]`, `hyperlinks?: CellHyperlink[]` and `annotationsDirty?: true`, and `SpreadsheetModel` gains `persons?: ThreadPerson[]`, per "The annotation model": plain JSON, each array sorted by (row, col) of its anchor, at most one note and one thread per cell, ids named `noteId`/`threadId`/`messageId`/`linkId`/`personId` (never `id`) and derived from the wire so re-reading a written file twice yields identical ids; absent and `[]` both mean none; every command replaces the array (copy-on-write, like `rows`), and `notes`, `threads` and `hyperlinks` are appended to SPEC-121's `SHEET_BY_REFERENCE_FIELDS` and the workbook-level `persons` (replaced, never mutated, when a new author is added) to its `WORKBOOK_SNAPSHOT_FIELDS`, so the undo snapshot holds them by reference
- [ ] Product-neutral helpers reach `lib/spreadsheet` only through core: `CommentMessage`, `replyParagraphText`, `parseReplyParagraphText`, `canModifyMessage` and `commentInitials` from `src/lib/core/comments.ts`, and `urlScheme`, `isDangerousUrl`, `sanitizeHref` from `src/lib/core/urls.ts` — if SPEC-101 / SPEC-097 have landed these files exist; otherwise this spec performs exactly their phase-1 moves (same names, same files, old paths re-exporting, every existing text-side test unchanged), and `spreadsheet/boundary.spec.ts` stays green
- [ ] xlsx notes read per "xlsx wire mapping": the worksheet's `comments` relationship yields `xl/commentsN.xml` (`authors`, `comment@ref`, `@authorId`, rich `<r>` runs → `runs` with SPEC-117 fonts), its `vmlDrawing` relationship (named by `<legacyDrawing r:id>`) yields the VML part whose `x:ClientData ObjectType="Note"` shapes match notes by `x:Row`/`x:Column`, giving `visible` (`<x:Visible/>`) and `box` (the shape style's `width`/`height` in pt); a legacy comment that is a thread's placeholder (author `tc={threadId}`) is **not** a note; a part that breaks the subset (two comments on one cell, an unresolvable author, unparseable VML) leaves that sheet with no `notes`/`threads` and every part byte-verbatim — annotations never half-parse
- [ ] xlsx threads read: the worksheet's `threadedComment` relationship yields `threadedComment` elements grouped by `parentId` into `CellThread`s (`dT` → `createdAt`, `personId` → `author` through `xl/persons/person.xml` → `model.persons`, root `done="1"` → `resolved`), message order by document order; a message carrying `<mentions>` keeps that element in `ThreadMessage.mentions` and re-emits it while the message is unedited
- [ ] ods notes read: an `office:annotation` child of a `table:table-cell` becomes a note — `dc:creator` → `author`, `dc:date` → `createdAt`, `text:p` paragraphs joined with `"\n"` → `text`, `office:display="true"` → `visible`, `svg:width`/`svg:height` (via `odfLengthToCm`) → `box` — and is lifted **out of** `Cell.preservedXml` (SPEC-122), so it is modelled once, not twice; reading sets nothing dirty
- [ ] Hyperlinks read: xlsx `<hyperlinks>/<hyperlink ref r:id location display tooltip>` — `r:id` → the worksheet relationship's `Target` (`TargetMode="External"`) as `{ kind: "url", href, subaddress?: location }`, `location` alone → `{ kind: "place", location }` normalised to canonical reference text (SPEC-119); ods — the **first** `text:a` in a cell's paragraphs, `xlink:href="#Sheet2.A1"` (also `#'My Sheet'.A1:B2`, `#$Sheet2.$A$1`, `#MyName`) → `place`, anything else → `url`, `office:title` → `tooltip`; a dangerous scheme is **kept** in the model and the file (round-trip fidelity) and is simply never opened
- [ ] The grid shows annotations without changing a cell's look: a note-bearing cell carries `sheet-note-indicator` (a 6 px red triangle in its top-right corner), a thread-bearing cell `sheet-thread-indicator` (a purple corner marker, dimmed when resolved), a linked cell `data-link="true"` and a pointer cursor over its text; the active cell's note or thread text is exposed to assistive technology through `aria-describedby` on a visually hidden element; nothing is restyled at read (a link's colour and underline come from the cell's own style, as in Excel)
- [ ] Hovering an annotated cell for `NOTE_HOVER_DELAY_MS = 400` shows `sheet-note-popover` (author, then the note text) or `sheet-thread-popover` (opening message, "N replies", a `sheet-thread-open` button that opens the panel on that thread); a note with `visible` — or every note while View ▸ "Show all notes" (`sheet-menu-show-all-notes`, session-only, never written) is on — renders as a `sheet-note-box` beside its cell in `SheetNoteLayer.vue`, at the default placement and its `box` size
- [ ] Note commands (`annotations.ts`): Insert ▸ "Note" / Shift+F2 / the cell context menu's "New note" opens `sheet-note-editor` (a plain `<textarea>`) on the active cell, committing on Escape, on blur and on Ctrl/Cmd+Enter (Excel keeps a note's text on Escape); Shift+F2 on a cell with a note edits it; "Delete note", "Show note"/"Hide note" toggle and edit commit through `addNote`, `editNote`, `deleteNote`, `setNoteVisible`, each one undo step ("Add note", "Edit note", "Delete note", "Show note", "Hide note") through SPEC-121's `record(label, { sheetId })` before mutating and `commit()` after; an emptied note on commit deletes it; the author is `loadCommentAuthor()`
- [ ] Threads (`.xlsx` workbooks): View ▸ "Comments" (`sheet-menu-comments`) or the app bar's `sheet-comments-toggle` toggles the `"comments"` pane of SPEC-129's sheet side panel (`useSheetSidePanel().toggle("comments")`; this spec widens `SheetPaneId` with `"comments"` and adds its `SHEET_PANES` entry), whose body is `SheetCommentsPanel.vue` (`sheet-comments-panel`, `aria-label="Comments"`), threads grouped by sheet in tab order (`sheet-comment-group`, heading = sheet name) then by (row, col), each `sheet-comment-thread-${threadId}` with `sheet-comment-message-${messageId}` rows (author, initials chip from `commentInitials`, `formatTimestamp`, text), `sheet-comment-reply-input`/`-submit`, `sheet-comment-resolve`/`-reopen`, `sheet-comment-edit`/`-delete` (own messages only, `canModifyMessage`), a filter `sheet-comment-filter` (all | unresolved), the identity field `sheet-comment-author-name` and `sheet-comment-empty`; Ctrl+Alt+M / Cmd+Option+M (`sheet.newComment` in SPEC-123's `SHEET_SHORTCUTS`; best effort on macOS, whose Option+Cmd+M minimises the app's windows — the context menu's "New comment" is the guaranteed path) starts a thread on the active cell in `sheet-comment-composer`; activating a thread selects its sheet and cell; labels "Add comment", "Reply to comment", "Edit comment", "Resolve comment", "Reopen comment", "Delete comment"
- [ ] In an `.ods` workbook "New comment" is disabled with the title "OpenDocument spreadsheets support notes only", and the panel lists the workbook's notes (`sheet-comment-note-${noteId}`) instead of threads; in both formats every annotation text renders as plain text — never markup, never `v-html` — with a security test feeding `<img src=x onerror=…>` through the popover, the note box, the panel and the indicator's accessible description
- [ ] Ctrl/Cmd+K (grid focused, no cell editor open), Insert ▸ "Link…" or the context menu opens `SheetLinkDialog.vue` (`sheet-link-dialog`, on `BaseDialog.vue`) with `sheet-link-text` (prefilled with the active cell's display text), the kind radios `sheet-link-kind-web` / `sheet-link-kind-place`, `sheet-link-address`, `sheet-link-sheet` (a `<select>` of visible sheets) + `sheet-link-ref` (validated with `parseRange`; SPEC-133's visible names listed in `sheet-link-names`), `sheet-link-tooltip`, `sheet-link-error`, `sheet-link-save`, `sheet-link-remove`, `sheet-link-cancel`; a dangerous scheme is refused inline with "This address type isn't allowed in links" (`LinkDialog.vue`'s string); saving sets one hyperlink over the selection's active range and, when the text differs from the cell's display, sets the active cell's value to that text through SPEC-122's cell-entry path — one undo entry, "Insert link" / "Edit link"
- [ ] Inserting a link through the dialog onto cells whose font colour is automatic also applies the theme `hlink` colour and a single underline through SPEC-123's `applyStylePatch` in the same undo entry; links read from files are never restyled
- [ ] The active cell's link shows `LinkChip.vue` (reused unchanged: `link-chip`, `link-chip-open`, `link-chip-edit`, `link-chip-copy`, `link-chip-remove`) pinned under the cell; a plain click on a linked cell **selects** it (a deliberate difference from Excel — the chip and Ctrl/Cmd+click follow it; Ctrl/Cmd+click over a linked cell's **text** follows the link instead of SPEC-117's add-a-range, while Ctrl/Cmd+click elsewhere in the cell still adds the range); following goes through `openHyperlinkTarget`: `http:`, `https:` and `mailto:` open with `window.open(href, "_blank", "noopener,noreferrer")`, any other scheme or a relative path is refused with the toast `LINK_NOT_OPENABLE_MESSAGE`, a `place` target activates its sheet, selects its range and scrolls it into view, and an unresolvable place (missing sheet, `#REF!`, unknown name) shows `LINK_DESTINATION_MISSING_MESSAGE`
- [ ] A cell whose formula's top-level node is a `HYPERLINK(…)` call (SPEC-127) behaves as linked — chip, Ctrl/Cmd+click, `data-link` — detected by `hyperlinkFormulaCall(cell)` on the AST SPEC-119's `parseFormulaAt(cell.formula, host)` returns (from its shared LRU — no second parse), with its target computed at follow time by evaluating the call's first argument through SPEC-120's `evaluate(node, ctx)` for that cell and coerced to text; it is never stored in `Sheet.hyperlinks` and never written as a `<hyperlink>`
- [ ] Structure follows the cells: `annotations.ts` registers the two `REFERENCE_ADJUSTERS` entries SPEC-124's registry table names for this spec, `"notes"` (notes and threads) and `"hyperlinks"`, per "Adjusters" — refs shift with inserted/deleted rows and columns, an annotation on a deleted cell is deleted, a hyperlink range shrinks or is deleted, every `place` location on any sheet is adjusted with `adjustFormula`, and a sheet whose annotations an adjuster changed is marked through `markAnnotationsEdited`; the same change deletes this spec's rows from SPEC-124's "Retained references" table (`hyperlink@ref`, `comment@ref`/`threadedComment@ref`, the VML `x:Row`/`x:Column`/`x:Anchor` row); the same two ids are registered in SPEC-125's `MOVE_ADJUSTERS` (`src/lib/spreadsheet/move.ts`) so cut-paste and drag-move carry a cell's note, thread and link with it, as Excel does — per "Adjusters"; `annotations.ts` registers `{ id: "annotations" }` in SPEC-128's `ROW_PERMUTATION_ADJUSTERS`, whose `permute` calls `remapAnnotations(model, sheetId, map)` with the sort's permutation, so a sort carries notes, threads and links with their records; SPEC-121's rename and delete pass every `place` location through `remapSheetReferences` beside `Cell.formula` (a deleted sheet's places become `#REF!`); a sheet's own annotations leave with it when it is deleted
- [ ] Every annotation command sets `sheet.annotationsDirty = true` on `.xlsx` and `sheet.dirty = true` on `.ods` (annotations live inside ODS cells, so the table is re-serialised — the SPEC-101 `markCommentsEdited` precedent) through one `markAnnotationsEdited(model, sheet)`; `stripRetention` drops every `source` and `annotationsDirty`
- [ ] xlsx write per "Writer rules": a sheet that is `annotationsDirty` but not `dirty` keeps `sheetData` and every other element byte-verbatim — only `<hyperlinks>` and `<legacyDrawing>` are spliced in CT_Worksheet order; the comments part is rebuilt (authors append-only), the VML part is **spliced** (non-note shapes such as legacy form controls verbatim), the threaded-comments part rebuilt and `xl/persons/person.xml` spliced; parts, relationships and content types are minted or removed as notes appear and vanish; an unchanged note, comment element, thread message or person re-emits its source bytes
- [ ] Every thread is mirrored into the comments and VML parts as Excel's placeholder per "The thread placeholder", rebuilt from the thread on every write, so an older Excel shows the thread read-only and Excel 365 does not raise a repair prompt
- [ ] ods write: SPEC-122's table writer emits each note as the first child of its `table:table-cell` (`source.odsAnnotation` verbatim while unchanged, otherwise generated per "ods wire mapping"), splits a repeated empty run to give a note on an unstored cell its own cell element, and wraps a linked cell's paragraph content in `<text:a xlink:type="simple" xlink:href=… office:title=…>`
- [ ] Cross-format "Download a copy as…": threads become notes on the way to `.ods` — the opening message followed by one `replyParagraphText` paragraph per reply, `resolved` dropped — and notes and hyperlinks convert field for field, so this spec deletes the "comments" and "hyperlinks" rows from SPEC-121's `UNCONVERTIBLE_CONTENT` table
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade` per "Reader upgrade": clean sheets are re-read; a dirty xlsx sheet reads notes/threads from its rels' parts in `opaque` and hyperlinks from its retained worksheet XML; a dirty ods sheet lifts notes from `Cell.preservedXml` and recovers links from its retained `sourceXml` only where the cell's current value still equals the source's at that address
- [ ] Round trip: a no-edit download is byte-identical part for part (`equalBytes`, `src/lib/core/equality.ts`, after `unzip`); add/edit/reply/resolve/link → save → reopen reproduces every modelled field in both formats (field-level `toEqual`, ids excluded, `equalStructural` from `src/lib/equality.ts` over the annotation arrays); writing the reopened model twice is byte-stable; the sanctioned deltas, each pinned by a test: a moved note's box keeps its size but not its custom position, an ods cell that carried two links keeps the first once rewritten, and a thread message edited after read loses its `<mentions>`
- [ ] Fixtures: `spreadsheet/fixtures/build.ts` gains `spec135-annotations.xlsx` (Excel-shaped, sheets "Summary" and "Data": a rich-text note whose first run is the bold author name, a visible note, a VML part that also holds a form-control checkbox shape, a thread with two replies and its placeholder, a resolved thread, a range hyperlink, external/internal/`mailto:` links, a tooltip, a `javascript:` link) and `spec135-annotations.ods` (LibreOffice-shaped: notes with and without `office:display`, internal and external `text:a`, a cell with two links); and a **human-obtained** Excel 365 workbook with a resolved thread, a reply and a legacy note, `e2e/fixtures/spec135-excel365-threads.xlsx`, is committed with a provenance note **before this spec is labelled ready** (Excel is not installed on the operator's Mac and the factory cannot produce one)
- [ ] e2e (`e2e/SPEC-135.spec.ts`): open `spec135-annotations.xlsx` → `B2` shows a note indicator and hovering it shows the note → Shift+F2 on `D4`, type a note, Escape commits → open the comments panel, reply to the thread as "You", resolve it → select `C7`, Ctrl/Cmd+K, link to place `Data!A10` → the chip's Open activates "Data" with `A10` selected → Ctrl/Cmd+click the `javascript:` link shows `LINK_NOT_OPENABLE_MESSAGE` and opens nothing → insert a row above row 2 → the note now sits on `B3` → download, reopen → every note, the reply, the resolved state and both links come back and the checkbox shape's VML bytes are unchanged; the read half repeats against `spec135-annotations.ods`, and `spec135-excel365-threads.xlsx` opens with its threads, authors and resolved state

## Out of scope
@-mention authoring and notifications (an existing `<mentions>` block is
carried, never created). Rich text editing inside notes (existing runs are kept
while a note is unedited; an edited note is plain text). Dragging or resizing a
note box (shown notes sit at the default placement). Note fill, border and font
formatting. Printing notes (`pageSetup@cellComments`, SPEC-136 carries it
verbatim). Comments on chart sheets and drawings. Copy-and-paste of notes and
links (cut-paste moves them; SPEC-125's copy payload carries cells only; Backlog). Excel's "Convert notes
to comments". Hyperlinks on shapes and images (SPEC-131 keeps them verbatim).
Links to other workbooks' places (`[Book2.xlsx]Sheet1!A1` is a `url` here and is
refused by the open rule). Autolink-as-you-type in cells. Following a link that
`HYPERLINK()` produces from inside a nested expression (only a top-level call
links). Comment sync between users (SPEC-044's territory).

## Reference

### Dependencies restated
- `specs/038-comments.md` — `CommentMessage { messageId; author; text; createdAt }`
  (ISO 8601); `replyParagraphText(m)` → `` `${author} (${createdAt}): ${text}` ``,
  `parseReplyParagraphText(text)` → `{author, createdAt, text} | null`,
  `canModifyMessage(m, author)`; identity in `src/lib/core/commentIdentity.ts`:
  `loadCommentAuthor()`, `saveCommentAuthor(name)`, `DEFAULT_COMMENT_AUTHOR = "You"`.
  `specs/101-…`'s phase 1 hoists the first four plus `commentInitials` into
  `src/lib/core/comments.ts`; `specs/097-…`'s phase 1 hoists `urlScheme`,
  `isDangerousUrl`, `sanitizeHref`, `DANGEROUS_SCHEMES` into `src/lib/core/urls.ts`
  (today both live in `src/lib/comments.ts` / `src/lib/links.ts`, text-product
  modules this layer may not import). `formatTimestamp` is `src/lib/format.ts`.
- `specs/025-hyperlinks-and-bookmarks.md` — `LinkChip.vue` (props `left`, `top`,
  `label`, `external`; emits `open`, `edit`, `copy`, `remove`, `dismiss`; no imports),
  dangerous schemes `javascript:`/`data:`/`vbscript:`.
- `specs/116-…` (normative model): copy-on-write rows, sparseness, the source is the
  splice base, `Sheet.partName`/`sourceXml`, `opaque`, `CellRef`/`GridRange`,
  `parseRange`, `formatA1`, `getCell`, `SPREADSHEET_READER_VERSION`, `ReaderUpgrade
  { from; upgradeDirtySheet(model, sheet, sourceParts) }`, `core/xmlScanner.ts`
  (worksheet parts are scanned, never `DOMParser`ed).
- `specs/117-…` `TextRun { text; font?: CellFont }`, `GridMetrics`, `SheetSelection`;
  `specs/119-…` canonical reference text, `parseFormulaAt(text, host)` → `{ node;
  origin }` (LRU of 50,000 keyed by `relativeFormulaKey`); `specs/121-…`'s
  `remapSheetReferences(canonical, remap)` (`formula/sheetRefs.ts`);
  `specs/120-…` `evaluate(node, ctx: EvalContext): EvalValue`; `specs/127-…`
  `HYPERLINK` (a value function); `specs/121-…` `useSpreadsheetUndo`:
  `record(label, scope: UndoScope)` before mutating, `commit()` after,
  `SHEET_BY_REFERENCE_FIELDS`; `Sheet.dirty`, `stripRetention`,
  `UNCONVERTIBLE_CONTENT`, `writeXlsx`/`writeOds`, `renameSheet`/`deleteSheet`; `specs/122-…` CT_Worksheet order (`hyperlinks` after
  `dataValidations`, `legacyDrawing` after `drawing`), `Cell.preservedXml`, the ODS
  table writer and its repeat compression; `specs/123-…` `applyStylePatch`;
  `specs/124-…` `StructuralEdit`, `adjustFormula(text, hostSheetId, edit, sheetNameOf)`,
  `REFERENCE_ADJUSTERS: { id; adjust(model, edit): void }[]` (adjusters replace,
  never mutate), and its `retained` adjuster's textual rewrite of `hyperlink@ref`,
  `comment@ref`, `threadedComment@ref` and the VML row/column fields — which is
  why a dirty sheet's retained parts are current when the upgrade reads them, and
  which this spec retires. `specs/125-…` `MOVE_ADJUSTERS: MoveAdjuster[]`
  (`{ id; move(model, move: RangeMove): void }`), `repointFormula(text,
  hostSheetId, move, sheetNameOf)` and the move table (inside the source moves,
  inside the destination is overwritten, partial overlap stays).
- `specs/128-…` — `ROW_PERMUTATION_ADJUSTERS` in `sort.ts`
  (`RowPermutationAdjuster { id; permute(model, sheetId, range, permutation:
  readonly { from; to }[]): void }`, run by `sortRange` after it writes the
  sorted records). `specs/129-…` — `useSheetSidePanel()` in
  `src/views/spreadsheet/useSheetSidePanel.ts`: `SheetPaneId` (widened per
  pane), `SHEET_PANES`, `{ activePane; open(id); close(); toggle(id); side;
  widthPx; collapsed; tabs }`, rendered by `SidePanel.vue` with the active
  pane's component in its default slot. `specs/133-…` — `model.names`
  (`DefinedName`, hidden and `_xlnm.*` names never listed) for
  `sheet-link-names` and `#MyName` places. `specs/123-…` — `SHEET_SHORTCUTS`.
  `ods/rangeAddress.ts` (SPEC-121): `odsRangeAddressToCanonical`,
  `canonicalToOdsRangeAddress`. `specs/117-…` — Ctrl/Cmd+click adds a range
  (`addRange`).

### The annotation model (`src/lib/spreadsheet/model.ts`)
```ts
import type { CommentMessage } from "../core/comments";

export interface CellNote {
  noteId: string;                 // xlsx `n-${A1 at read}`; ods office:name, else `n-${A1}`
  ref: CellRef;
  author: string;                 // "" when the wire names none
  createdAt?: string;             // ods dc:date; legacy xlsx notes carry no date
  text: string;                   // plain text; paragraphs joined with "\n"
  runs?: TextRun[];               // xlsx rich runs, kept while the note is unedited
  visible?: true;
  box?: { widthPx: number; heightPx: number };   // absent → NOTE_DEFAULT_BOX (144 × 79)
  source?: { xlsxComment?: string; vmlShape?: string; odsAnnotation?: string };
}
export interface ThreadMessage extends CommentMessage {
  personId: string;
  mentions?: string;              // the retained <mentions> element
}
export interface CellThread {
  threadId: string;               // the root threadedComment@id, "{GUID}" as written
  ref: CellRef;
  messages: ThreadMessage[];      // [0] opens the thread
  resolved: boolean;
}
export interface ThreadPerson { personId: string; displayName: string; userId?: string; providerId?: string }
export type HyperlinkTarget =
  | { kind: "url"; href: string; subaddress?: string }
  | { kind: "place"; location: string };          // canonical reference text, or a defined name
export interface CellHyperlink {
  linkId: string;                 // xr:uid, else `h-${A1 of range start}`
  range: GridRange;
  target: HyperlinkTarget;
  tooltip?: string;
  display?: string;               // xlsx @display, retained
  source?: { relId?: string };
}
// Sheet: notes?, threads?, hyperlinks?, annotationsDirty?: true
// SpreadsheetModel: persons?: ThreadPerson[]
```
A command that changes a note's content builds a new note **without** `source`;
an adjuster or remap that only moves it keeps `source` and the writer patches
the position fields. New GUIDs are `` `{${crypto.randomUUID().toUpperCase()}}` ``.
New persons are `{ displayName: name, userId: name, providerId: "None" }`.
Commands in `src/lib/spreadsheet/annotations.ts` (framework-free): `addNote`,
`editNote`, `deleteNote`, `setNoteVisible`, `addThread`, `replyToThread`,
`editThreadMessage`, `deleteThreadMessage` (deleting `messages[0]` deletes the
thread), `setThreadResolved`, `setHyperlink`, `removeHyperlink`,
`hyperlinkAt(sheet, ref)`, `noteAt`, `threadAt`, `hyperlinkFormulaCall(cell)`,
`remapAnnotations(model, sheetId, map: (ref: CellRef) => CellRef | null)`,
`markAnnotationsEdited(model, sheet)`, `openHyperlinkTarget(target): "external" |
"place" | "refused"` (pure decision; the view acts). Constants:
`NOTE_HOVER_DELAY_MS = 400`, `NOTE_DEFAULT_BOX`, `LINK_NOT_OPENABLE_MESSAGE` =
"This link can't be opened from the browser", `LINK_DESTINATION_MISSING_MESSAGE`
= "This link's destination doesn't exist".

### xlsx wire mapping
| Part | Relationship (from) | Content type |
| --- | --- | --- |
| `xl/comments{n}.xml` | `…/officeDocument/2006/relationships/comments` (worksheet) | `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml` |
| `xl/drawings/vmlDrawing{n}.vml` | `…/officeDocument/2006/relationships/vmlDrawing` (worksheet, named by `legacyDrawing@r:id`) | `Default Extension="vml"` `application/vnd.openxmlformats-officedocument.vmlDrawing` |
| `xl/threadedComments/threadedComment{n}.xml` | `http://schemas.microsoft.com/office/2017/10/relationships/threadedComment` (worksheet) | `application/vnd.ms-excel.threadedcomments+xml` |
| `xl/persons/person.xml` | `http://schemas.microsoft.com/office/2017/10/relationships/person` (workbook) | `application/vnd.ms-excel.person+xml` |

- **Comments:** `<comments><authors><author>…</author></authors><commentList><comment
  ref="B2" authorId="0" shapeId="0" xr:uid="{…}"><text>…</text></comment>`. `<text>`
  is `<t>` or `<r>` runs (`<rPr>` read by SPEC-117's run reader); `text` is the runs'
  concatenation with `_xHHHH_` decoded. Authors are **append-only** on write.
- **VML:** `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="…:office:office"
  xmlns:x="…:office:excel">`, an `o:shapelayout` (`o:idmap data`), the
  `v:shapetype id="_x0000_t202"` and one `v:shape` per note; a note shape is the one
  whose `x:ClientData@ObjectType="Note"`. `x:Row`/`x:Column` are 0-based; `x:Anchor`
  is `col, dx, row, dy, col2, dx2, row2, dy2`; size from the shape's
  `style` `width:…pt;height:…pt`. Generated shapes follow
  `src/lib/spreadsheet/xlsx/vmlNoteGolden.ts` (Excel's default note: `fillcolor="#ffffe1"`,
  `visibility:hidden` unless shown, anchor `{c+1}, 15, {max(r−1,0)}, 10, {c+3}, 15,
  {max(r−1,0)+4}, 4`, id `_x0000_s{idmap·1024+n}` above the highest existing).
- **Threads:** root `<ThreadedComments
  xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`,
  `<threadedComment ref dT personId id [parentId] [done]><text>…</text>[<mentions>]</threadedComment>`;
  persons `<personList><person displayName id userId providerId/></personList>`.
- **Hyperlinks:** `<hyperlinks><hyperlink ref r:id location display tooltip xr:uid/>`;
  external targets are worksheet relationships of type
  `…/officeDocument/2006/relationships/hyperlink` with `TargetMode="External"`.
  `location` may name `Sheet!A1`, `'My Sheet'!A1:B2` or a defined name.

### The thread placeholder
Excel 365 writes every thread twice: in the threaded part, and as a legacy
`<comment ref xr:uid="{threadId}">` whose author is `tc={threadId}` and whose text
is `THREAD_PLACEHOLDER_PREAMBLE + "Comment:\n    " + messages[0].text` followed by
`"\nReply:\n    " + text` per reply, with a hidden VML note shape. The preamble lives
in `src/lib/spreadsheet/xlsx/threadPlaceholder.ts`:
`"[Threaded comment]\n\nYour version of Excel allows you to read this threaded
comment; however, any edits to it will get removed if the file is opened in a
newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n"`.
**This string and the `tc=` author form are pinned from
`spec135-excel365-threads.xlsx` before the spec goes ready — where that file
differs, its bytes win and the constant follows** (SPEC-086's golden-first rule).
The reader recognises a placeholder by its `tc=` author (never by its text, which
an older Excel may have edited); the writer regenerates every placeholder from
its thread, so they cannot drift.

### ods wire mapping
`<office:annotation [office:display="true"] [office:name]
[svg:width svg:height …]><dc:creator>…</dc:creator><dc:date>…</dc:date><text:p>…</text:p>…</office:annotation>`
as the first child of `table:table-cell`, declaring `xmlns:dc` locally the way
`src/lib/odf/writer.ts#annotationStartXml` does. Links: `<text:a xlink:type="simple"
xlink:href office:title>` inside the cell's `text:p`. Internal hrefs are `#` + an ODS
address (`Sheet2.A1`, `$Sheet2.$A$1`, `'My Sheet'.A1:B2`) or `#` + a name; they are
converted with `odsRangeAddressToCanonical` / `canonicalToOdsRangeAddress` in
`src/lib/spreadsheet/ods/rangeAddress.ts` (SPEC-121's module).

### Writer rules
- `writeXlsx`: a sheet is annotation-rebuilt when `dirty || annotationsDirty`.
  `<hyperlinks>` is re-serialised from the model (removed when empty) and
  `<legacyDrawing r:id>` emitted iff the sheet has a VML part; both are spliced into
  the retained worksheet XML in CT_Worksheet order — when only `annotationsDirty`,
  nothing else in the part changes. New external links append relationships after
  the highest `rId`; an unreferenced hyperlink relationship is left in place (the
  SPEC-097 rule). Comments, VML and threaded parts are minted at the next free `n`,
  removed with their relationships and `Override`s when a sheet's last note and
  thread go (the VML part survives while it holds other shapes); `[Content_Types].xml`
  and `.rels` changes are string splices, every other declaration verbatim.
- `writeOds`: annotation edits mark the sheet `dirty`, so SPEC-122's table
  re-serialisation applies, with the two additions in the AC.
- No sheet with untouched annotations changes: a clean workbook still re-emits
  every part byte-identical.

### Adjusters
`"notes"` (notes and threads): for the edited sheet, a ref in an inserted band's wake
shifts by `count`; a ref inside a deleted band is removed. `"hyperlinks"`: ranges
shift, a range straddling a deletion shrinks, a range wholly deleted is removed;
then for **every** sheet, each `place` location that parses as a reference is
passed through `adjustFormula(location, hostSheetId, edit, sheetNameOf)` (a
location left holding `#REF!` stays, and following it reports
`LINK_DESTINATION_MISSING_MESSAGE`). Each adjuster replaces arrays, never mutates
them.

`MoveAdjuster`s with the same ids, `move(model, move: RangeMove)` where
`RangeMove { fromSheetId; from; toSheetId; toTop; toLeft }`, follow SPEC-125's
table: a note, thread or hyperlink range entirely inside `from` moves by the
offset onto `toSheetId`; one entirely inside the destination and not in the
source is deleted (its cells were overwritten); a hyperlink range that only
partly overlaps either stays; every `place` location on every sheet passes
through `repointFormula(location, hostSheetId, move, sheetNameOf)`. Both sheets'
annotations are marked through `markAnnotationsEdited`. `remapAnnotations`
applies an arbitrary cell mapping and serves the `annotations` entry in
SPEC-128's `ROW_PERMUTATION_ADJUSTERS` (`RowPermutationAdjuster { id;
permute(model, sheetId, range, permutation: { from; to }[]) }` — a row
permutation inside the sorted range, mapped column by column within it,
`null` never produced).

### Reader upgrade
The appended `ReaderUpgrade.upgradeDirtySheet`: xlsx — follow the sheet part's
`.rels` in `sourceParts` to the comments, VML and threaded parts and parse them as
above; `model.persons` is read once from the workbook rels if absent; hyperlinks are
read from the retained worksheet XML, which already carries SPEC-124's textual
`@ref` adjustments. ods — move every `office:annotation` out of `Cell.preservedXml`
into `notes`; for each `text:a` in the retained `sourceXml`, attach the link only
when the current cell at that address holds the value the source cell held (a
value or structural edit since read makes the address unreliable, and such a link
is not resurrected — SPEC-122 had already dropped it from any save).

### UI wiring
`src/views/spreadsheet/useCellNotes.ts`, `useSheetComments.ts`, `useCellLinks.ts`
(each ≤ 400 lines, `composableSize.spec.ts`); components
`src/components/spreadsheet/SheetNotePopover.vue`, `SheetNoteLayer.vue`,
`SheetCommentsPanel.vue`, `SheetLinkDialog.vue`, plus `LinkChip.vue` reused. The
panel and layer are dumb `:controller`-taking components with runtime
`defineProps` (espree). `SpreadsheetView.vue` gains only three `use*` wiring calls
and template hookup; `SpreadsheetView.size.spec.ts`'s budget never rises.

### Tests
- `spreadsheet/annotations.spec.ts` — every command, copy-on-write, dirty marking,
  `openHyperlinkTarget` scheme table, `hyperlinkFormulaCall`, `remapAnnotations`;
  `structure.spec.ts` gains one case per adjuster (a row insert above moves it, a
  column delete through it removes it, a `place` location in another sheet follows)
  and `move.spec.ts` one per `MoveAdjuster` (cut-paste carries a note and a link,
  a destination annotation is overwritten, a partial link overlap stays).
- `xlsx/annotations.spec.ts` / `ods/annotations.spec.ts` — read, write, the
  out-of-subset preserve path, placeholder regeneration, VML splice keeping the
  checkbox shape, golden equality for `vmlNoteGolden` and `threadPlaceholder`.
- `annotationsRoundTrip.spec.ts` per the round-trip AC; `open.spec.ts` drives the
  upgrade with synthetic dirty sheets in both formats; `core/comments.spec.ts` /
  `core/urls.spec.ts` if this spec performs the moves.
- `useCellNotes.spec.ts`, `useSheetComments.spec.ts`, `useCellLinks.spec.ts`, the
  plain-text security test, and `e2e/SPEC-135.spec.ts`.

### Specs this changes
- `specs/116-…` — the sparseness invariant is unchanged (annotations live on the
  sheet, not in cells). `specs/122-…` — ODS annotations leave `Cell.preservedXml`;
  the table writer gains the annotation slot and `text:a`. `specs/124-…` — the
  `notes` and `hyperlinks` adjusters join the registry and this spec's four rows
  leave "Retained references". `specs/125-…` — `notes` and `hyperlinks` join
  `MOVE_ADJUSTERS`. `specs/128-…` — `ROW_PERMUTATION_ADJUSTERS` gains
  `annotations`. `specs/129-…` — `SheetPaneId` gains `"comments"`.
  `specs/117-…` — Ctrl/Cmd+click over a linked cell's text follows the link
  instead of adding a range. `specs/123-…` — `SHEET_SHORTCUTS` gains
  `sheet.newComment` and the link dialog's Ctrl/Cmd+K. `specs/121-…` —
  rename/delete remap link locations; `stripRetention` drops annotation
  retention; the "comments" and "hyperlinks" rows leave
  `UNCONVERTIBLE_CONTENT`; three fields join `SHEET_BY_REFERENCE_FIELDS` and
  `persons` joins `WORKBOOK_SNAPSHOT_FIELDS`. (SPEC-136, later, keeps
  `pageSetup@cellComments` verbatim.)
