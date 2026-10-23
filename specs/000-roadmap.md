---
id: SPEC-000
title: Product roadmap — word processor, presentations, media and spreadsheets
status: draft
area: product
---
## Why
This file is the index for the spec set. It is not a work order — no engineer
should implement it. It exists so a business person can see the intended order
of the queue and pick which spec to label `status:ready` next.

The first product: a browser-based word processor that opens `.docx` (OOXML)
and `.odt` (ODF) documents, edits them with the feature depth people expect
from Google Docs and LibreOffice Writer, saves them back in the same format
without destroying what it did not understand, and generates new documents in
both formats from scratch.

The second product (Phase 13): presentations — `.pptx`/`.odp` decks in the
same app and library, sharing `src/lib/core` but never importing the word
processor (see `docs/presentations-architecture.md`). Its preparation
milestone (open, library, outline preview, byte-preserving save) shipped
outside the numbered queue on 2026-08-24; SPEC-079 onward builds the editor
on top of it.

The third product (Phase 15): media — images and PDFs in the same library,
at `/media/:id`, with the assembly workspace that merges them
(SPEC-107–112). It is a third box on the same layering diagram: a
`src/lib/media/` product layer over `src/lib/core`, meeting the other two
only in the app layer. It deliberately does **not** render a PDF page —
there is no rasteriser here and no dependency may be added to build one, so
every surface that would want one says so and offers the browser's own
viewer instead.

Phase 16 is the first experimental phase: features behind a Labs flag, off
by default, including the only place this application may talk to a model.

The fourth product (Phases 17–18): spreadsheets — `.xlsx` and `.ods`
workbooks (and, from SPEC-126, `.csv`/`.tsv`) in the same library, at
`/sheets/:id`. It is a fourth box on the same layering diagram: a
`src/lib/spreadsheet/` product layer over `src/lib/core`, meeting the other
three only in the app layer. Everything a spreadsheet needs that a word
processor would not — the formula language, the calculation engine, the
number-format engine, a forward-only XML scanner for million-cell parts — is
written here; `package.json` gains nothing.

Specs are ordered so each one is shippable on its own and builds on the ones
before it. Later specs assume earlier ones are merged (`depends_on` in each
file's frontmatter).

## Acceptance criteria
- [ ] No implementation. This spec is never labelled `status:ready`.

## Out of scope
Everything. See the individual specs.

---

## Spec conventions
Intake sees **one spec file at a time**, so a spec that leans on a shared
concept must carry that concept with it. Two rules, both learned from intake
rejecting SPEC-007 as ambiguous:

1. **One normative source per shared concept.** The canonical document model,
   its units, and the equality relations are defined in
   `specs/002-create-blank-document.md` and, once merged, in
   `docs/document-model.md`. The I/O API surface is defined in
   `specs/004-open-docx-text-and-paragraphs.md` (readers) and
   `specs/006-save-as-docx.md` (writers). The shared fixture set and its
   `manifest.json` are defined in SPEC-004. The anchored drawing object —
   position, wrap, z-order and the canvas/ARIA-proxy rendering every
   floating object uses — is defined in
   `specs/051-drawing-shapes-and-text-boxes.md`; the drawing **group
   container** that supersedes SPEC-051's flat grouping is defined in
   `specs/057-group-and-ungroup-drawings.md`. The **application's** visual
   vocabulary — design tokens, the icon set and the shared control
   primitives — is defined in
   `specs/058-design-system-and-visual-refresh.md`, and is not to be
   confused with the **document's** theme (colours and fonts saved into the
   file), defined in `specs/077-document-themes.md`. The shape geometry
   catalogue that supersedes SPEC-051's fourteen entries is defined in
   `specs/069-shape-catalogue-expansion.md`. On the presentation side:
   the slide model (frames, effective values, preservation fields) is
   defined in `specs/079-slide-canvas-rendering.md`; the presentation
   writers' rebuild rules and the presentation undo stack in
   `specs/080-create-presentation-and-slide-management.md`; slide-part
   serialisation in `specs/081-slide-text-editing.md`; playback
   navigation semantics in `specs/084-slideshow-playback.md`; the
   transition model and wire mappings in
   `specs/085-slide-transitions.md`; the animation model and canonical
   timing forms in `specs/086-slide-animations.md`. On the media
   side: the media kind, the `MediaModel` and the structural PDF reader
   (`core/pdfDocument.ts`) are defined in
   `specs/107-images-and-pdfs-in-the-library.md`; the sheet — page size,
   margins, fit and the resolution ladder every assembled output lands on —
   in `specs/108-the-assembly-sheet-page-size-and-margins.md`; the assembly
   workspace and its layout engine in
   `specs/109-combine-images-into-one-image.md`; the PDF writer half and
   the merge contract in `specs/111-merge-pdfs-and-organise-pages.md`; the
   image edit stack in `specs/112-edit-an-image.md`. And on the
   experimental side: the Labs flag registry and the **only** module
   permitted to make a model request are defined in
   `specs/113-experimental-features-and-the-model-gateway.md`. And on the
   spreadsheet side: the workbook model, its invariants (copy-on-write rows,
   sparseness, the retained source as splice base, dates as serials), A1
   addressing, date serials, the General display, reader versions and the
   per-sheet row split are defined in
   `specs/116-workbooks-in-the-library.md`; cell styles, colour mapping,
   grid geometry and the selection model in
   `specs/117-sheet-grid-styles-sizes-and-selection.md`; number formats and
   the spreadsheet locale in `specs/118-number-formats-and-dates.md`; the
   formula language — canonical text, tokens, AST, references and both wire
   dialects — in `specs/119-the-formula-language.md`; evaluation semantics,
   recalculation and the function registry in
   `specs/120-calculation-engine-and-core-functions.md`; the spreadsheet
   writers' rebuild rules and undo stack in
   `specs/121-create-workbook-and-sheet-management.md`; sheet-part
   serialisation, the append-only shared-strings and styles writers and
   input parsing in `specs/122-cell-entry-and-the-sheet-writer.md`;
   structural reference adjustment and the adjuster registry every
   reference-bearing structure joins in
   `specs/124-rows-columns-and-merged-cells.md`; its cut-and-move twin
   (`MOVE_ADJUSTERS`) in `specs/125-copy-paste-fill-and-move.md`; the
   row-permutation hook sorting runs (`ROW_PERMUTATION_ADJUSTERS`) in
   `specs/128-sort-and-filter.md`; the spreadsheet side panel
   (`useSheetSidePanel`) in `specs/129-conditional-formatting.md`; and
   defined names (validity, resolution, the reserved `_xlnm.*` names), Excel
   tables, the meaning and spelling of structured references and how a
   table style renders in `specs/133-named-ranges-and-tables.md`. Two
   shared helpers are owned by the first spec that needed them: ODS range
   addresses (`spreadsheet/ods/rangeAddress.ts`) by SPEC-121, row and column
   autofit (`spreadsheet/autofit.ts`) by SPEC-122. Nothing redefines these.

2. **Every dependent spec ends with a `## Reference` section** that names the
   dependency spec files by path, restates the contract it relies on inline
   (model shape, function signatures, fixture layout), and names the equality
   relation each test uses. A reader of one file alone must never have to guess.

3. **Every spec not yet built carries a size in its frontmatter** — one
   `complexity` band and its `story_points` estimate:

   ```yaml
   complexity: complex   # complex | normal | easy
   story_points: 13      # 21, 13 | 8, 5 | 3
   ```

   The band is the normative field and the only one the factory reads; the
   points are the estimate that justifies it, on the Fibonacci scale this
   queue has always used informally. They are written **before** a spec is
   labelled `status:ready`, by whoever writes the spec, because the judgement
   is about the work the spec describes and not about how a run is going.

   | Band | Points | What earns it |
   | --- | --- | --- |
   | `complex` | 13, 21 | New machinery with no seam to hang it on — a decoder or engine written from scratch, a type change that crosses the whole model, a dependency this product does not have. |
   | `normal` | 5, 8 | A new model area, one or two wire mappings and a UI surface, built on seams that already exist and named exactly by the spec. |
   | `easy` | 3 | Mostly pure functions and wiring: one surface, no new format mapping, nothing the codebase cannot already almost do. |

   The band is a claim about the *work*, not about how carefully the spec is
   written — every file here is written to the same normative bar, so length
   and acceptance-criteria count are not the signal. The factory reads the
   band to choose which model opens the implement stage
   (`routing.implementByComplexity` in the factory's `config.json`); a spec
   with no band falls back to the default ladder, so omitting it is degraded,
   never broken. Specs already built carry no band — it would be a
   retrospective guess about work that is finished.

Any spec asserting a round trip must name one of the three relations —
`equalBytes`, `equalStructural`, `equalSemantic` — and use the helpers in
`src/lib/equality.ts`. "Asserts model equality" is not a specification. Note
that `equalSemantic` is specified but not yet implemented; see SPEC-002's
"Implementation status".

---

### Phase 1 — Foundation and file I/O
| Spec | Title |
| --- | --- |
| SPEC-001 | App shell and document workspace |
| SPEC-002 | Create a new blank document |
| SPEC-003 | Local document library |
| SPEC-004 | Open .docx — text and paragraphs |
| SPEC-005 | Open .odt — text and paragraphs |
| SPEC-006 | Save as .docx |
| SPEC-007 | Save as .odt |
| SPEC-008 | Round-trip preservation of unsupported content |
| SPEC-009 | Drag-and-drop and open-from-URL |
| SPEC-010 | Autosave and crash recovery — queue after SPEC-011 |

### Phase 2 — Core editing
| Spec | Title |
| --- | --- |
| SPEC-011 | Rich text editing core |
| SPEC-012 | Character formatting — bold, italic, underline, strikethrough |
| SPEC-013 | Font family, size, text colour and highlight |
| SPEC-014 | Paragraph alignment, indents and spacing |
| SPEC-015 | Undo and redo |
| SPEC-016 | Cut, copy and paste |
| SPEC-017 | Find and replace |
| SPEC-018 | Bulleted and numbered lists |
| SPEC-019 | Paragraph styles and the style picker |
| SPEC-020 | Clear formatting and format painter |
| SPEC-021 | Keyboard shortcuts and shortcut reference |

### Phase 3 — Objects and structure
| Spec | Title |
| --- | --- |
| SPEC-022 | Tables — insert and edit |
| SPEC-023 | Table formatting and cell operations |
| SPEC-024 | Images — insert, resize and wrap |
| SPEC-025 | Hyperlinks and bookmarks |
| SPEC-026 | Special characters and symbol picker |

### Phase 4 — Page layout and long documents
| Spec | Title |
| --- | --- |
| SPEC-027 | Page setup — size, margins, orientation |
| SPEC-028 | Page and section breaks |
| SPEC-029 | Headers, footers and page numbers |
| SPEC-030 | Multi-column layout |
| SPEC-031 | Footnotes and endnotes |
| SPEC-032 | Multilevel lists and numbering control |
| SPEC-033 | Table of contents |
| SPEC-034 | Print and export to PDF |

### Phase 5 — Authoring depth, review and collaboration
| Spec | Title |
| --- | --- |
| SPEC-035 | Spell check with user-supplied dictionaries, and document language |
| SPEC-036 | Word count and document statistics |
| SPEC-037 | Document outline navigator |
| SPEC-038 | Comments |
| SPEC-049 | Menu bar and toolbar consolidation — queue before SPEC-039 |
| SPEC-039 | Track changes |
| SPEC-040 | Version history and restore |
| SPEC-041 | Document properties and metadata |
| SPEC-042 | Custom styles manager |
| SPEC-043 | Accessibility checker |
| SPEC-044 | Real-time collaborative editing |

### Phase 6 — Generation and scale
| Spec | Title |
| --- | --- |
| SPEC-045 | Document generation API |
| SPEC-046 | Template gallery and mail merge |
| SPEC-047 | Large document performance |
| SPEC-048 | Font fallback and embedding |
| SPEC-050 | Markdown — open, edit visually, and convert to and from .docx/.odt |

### Phase 7 — Drawing objects and fields
| Spec | Title |
| --- | --- |
| SPEC-051 | Drawing shapes and text boxes |
| SPEC-052 | Charts — render, insert and edit data |
| SPEC-053 | Fields — date, time, document information and cross-references |

### Phase 8 — Authoring precision
| Spec | Title |
| --- | --- |
| SPEC-054 | Tab stops and the tab ruler — resolves the SPEC-014/SPEC-032 conflict |
| SPEC-055 | Captions, caption cross-references and the table of figures |
| SPEC-056 | Equations |
| SPEC-057 | Group and ungroup drawings — native containers, nesting and editing in place |

### Phase 9 — Interface modernisation
The product's feature depth has outgrown its interface. SPEC-058 comes
first and the rest are written against it: it creates the design tokens,
icon set and control primitives every other spec in this phase spends.
SPEC-060 is independently shippable and fixes the product's most visible
hole — there is no way out of a document.

| Spec | Title |
| --- | --- |
| SPEC-058 | Design system — tokens, icons and the visual refresh |
| SPEC-059 | Dark mode and theme preference |
| SPEC-060 | Close the document and return to the library |
| SPEC-061 | Toolbar redesign — icons, grouping, overflow and contextual controls |
| SPEC-062 | Command palette |
| SPEC-063 | Zoom and view modes |
| SPEC-064 | The vertical ruler and measurement units everywhere |
| SPEC-065 | The unified side panel |
| SPEC-066 | Preferences |
| SPEC-067 | The document library — overview redesign |
| SPEC-068 | Responsive layout, touch and small screens |

### Phase 10 — Drawing depth
SPEC-051 shipped fourteen shapes; SPEC-057 made groups real. This phase
closes the distance to what documents actually contain — flowcharts,
connectors, freeform outlines and styling beyond one flat colour. SPEC-069
is the prerequisite for the rest; SPEC-072 is the one SPEC-057 names in its
own out-of-scope list.

| Spec | Title |
| --- | --- |
| SPEC-069 | Shape catalogue expansion and adjustment handles |
| SPEC-070 | Connectors |
| SPEC-071 | Freeform shapes, curves and edit points |
| SPEC-072 | Align, distribute, snap and the position dialog |
| SPEC-073 | Shape styling depth — gradients, effects, arrowheads and presets |

### Phase 11 — Authoring completeness
Four features earlier specs put out of scope and nothing came back for.
Each is named in an existing spec's "Out of scope" section, and each is
something a real document arrives carrying — so all four are fidelity
fixes before they are features.

| Spec | Title |
| --- | --- |
| SPEC-074 | Paragraph and page borders and shading |
| SPEC-075 | Watermarks and line numbering |
| SPEC-076 | Table styles, header repeat, sort and text↔table conversion |
| SPEC-077 | Document themes — colour and font schemes |

### Phase 12 — Structural fidelity
Debt in the drawing model that feature work keeps landing on. SPEC-051
anchored floating drawings by block-list adjacency; two drawings on one
paragraph make that inference ambiguous, and the ambiguity surfaces as
wrong anchor markers, position jumps on re-anchor, and a file format that
tells Word the wrong anchor paragraph.

| Spec | Title |
| --- | --- |
| SPEC-078 | Explicit anchors for floating drawings |

### Phase 13 — The presentation product
Builds the slides editor on the 2026-08-24 preparation milestone
(read-only preview + byte-preserving save). Order matters more than
usual here: SPEC-079 defines the slide model everything else edits,
SPEC-080/081 build the save machinery every content spec reuses, and
the last three are the product's payoff features — Keynote import needs
the editor to import *into*, and the HTML export replays what SPEC-084
–086 define. SPEC-087 additionally needs a human-generated fixture
(`scripts/make-keynote-fixture.applescript`) committed before it can go
ready.

| Spec | Title |
| --- | --- |
| SPEC-079 | Slide rendering — the slide canvas, placeholders and the filmstrip |
| SPEC-080 | Create a presentation, and slide management |
| SPEC-081 | Slide text editing and formatting |
| SPEC-082 | Slide objects — insert, arrange and style images, shapes and text boxes |
| SPEC-083 | Speaker notes |
| SPEC-084 | Slideshow playback |
| SPEC-085 | Slide transitions |
| SPEC-086 | Slide animations — entrance and exit builds |
| SPEC-087 | Import Apple Keynote presentations (.key) |
| SPEC-088 | Export a presentation to HTML — self-contained, with transitions and animations |

### Phase 14 — Presentation depth
Written 2026-08-26 in one pass, after Phase 13 made the deck editable.
Each closes a hole an earlier presentation spec named in its own "Out of
scope": layouts and masters, deck setup, the two content objects a deck
cannot do without, review and accessibility, and the two export surfaces.
SPEC-089 comes first — the layout model is what the rest assign to.
SPEC-104 additionally moves SPEC-034's PDF layer into `src/lib/core`.

| Spec | Title |
| --- | --- |
| SPEC-089 | Slide layouts — the layout catalogue, picker and Reset slide |
| SPEC-090 | Slide size, orientation and deck setup |
| SPEC-091 | Tables on slides — grid model, cell editing and both wire formats |
| SPEC-092 | Charts on slides — render, insert and edit data |
| SPEC-093 | Slide footers, slide numbers and dates |
| SPEC-094 | Deck themes — colour and font schemes |
| SPEC-095 | Slide media — audio and video frames |
| SPEC-096 | Slide object depth — rotation, handles, multi-select, groups, align and distribute |
| SPEC-097 | Slide text depth — lists, spacing, autofit and links |
| SPEC-098 | Deck find and replace, and spell check on slides |
| SPEC-099 | The slide sorter, deck sections and slide reuse |
| SPEC-100 | Presenter view — notes, next slide and the second window |
| SPEC-101 | Comments on slides — review threads, markers and both wire formats |
| SPEC-102 | Deck accessibility — alt text, reading order and the checker |
| SPEC-103 | Autosave, drafts and version history for presentations |
| SPEC-104 | Export a deck to PDF, images and handouts |
| SPEC-105 | Animation depth — paragraph builds, emphasis effects and motion paths |
| SPEC-106 | Presentation generation API, and Markdown to slides |

### Phase 15 — Images, PDFs and assembly
The third product. A library that holds documents and decks but not the
scan, the screenshot and the PDF beside them sends its user somewhere else
for the most ordinary errands there are — combine these images, make this a
PDF, merge these three, cover that account number.

Order is load-bearing. SPEC-107 adds the kind and the structural PDF reader
everything later parses with. SPEC-108 fixes the geometry once, so three
merge specs spend one definition of "page size" rather than three. SPEC-109
builds the assembly workspace; SPEC-110 adds pages and PDF output to it;
SPEC-111 adds PDF sources and the merge. SPEC-112 is independent of the
assembly chain and can be taken any time after SPEC-107.

| Spec | Title |
| --- | --- |
| SPEC-107 | Images and PDFs in the library |
| SPEC-108 | The assembly sheet — page size, margins and fit |
| SPEC-109 | Combine images into one image |
| SPEC-110 | Assemble images into a PDF |
| SPEC-111 | Merge PDFs, and organise their pages |
| SPEC-112 | Edit an image |

### Phase 16 — Experimental
Two speculative features and the machinery that makes shipping them safe.
SPEC-113 is the prerequisite for both: a Labs flag registry (everything off
by default, an unflagged feature not even in the bundle) and the single
audited module through which a model request may be made, with the consent
and credential rules it obeys. Neither feature is worth shipping without
it, and SPEC-114 uses only its flag half — library search makes no network
request at all.

| Spec | Title |
| --- | --- |
| SPEC-113 | Experimental features — the flag registry and the model gateway |
| SPEC-114 | Search inside the library (experimental) |
| SPEC-115 | Design a presentation from a description (experimental) |

### Phase 17 — The spreadsheet product
The fourth product. A library that holds the report but not the numbers
behind it sends its user somewhere else the moment the numbers change.

Order is load-bearing, and stricter than Phase 13's. SPEC-116 adds the kind,
both readers, a byte-preserving download and a read-only grid; its model is
what every later spec edits. SPEC-117–120 then teach the readers everything a
faithful *view* needs — styles and geometry, number formats, the formulas
themselves and their evaluation — **before anything can be edited**, so each
of them upgrades already-stored workbooks by simply re-reading them (SPEC-116's
reader-version rule) instead of migrating user edits. SPEC-121/122 build the
save machinery every editing spec reuses: workbook-level rebuild rules and
the undo stack, then the sheet writer and the cell editor. SPEC-123–127 are
the staples: formatting, rows and columns, the clipboard and fill, CSV, and
the rest of the function library.

Three specs here need something only a person can produce, committed before
they are labelled ready (the factory container runs no office suite):
SPEC-116 an Apple Numbers export (`e2e/fixtures/spec116-numbers.xlsx`, via
`scripts/make-numbers-fixture.applescript`) and a check of its Numbers
detection markers; SPEC-120 a Numbers-calculated answer key
(`spec120-numbers.xlsx`); SPEC-125 four real clipboard captures — Excel for
the web, Google Sheets, LibreOffice and Numbers — with a note of how each was
taken.

| Spec | Title |
| --- | --- |
| SPEC-116 | Workbooks in the library — open .xlsx and .ods, the workbook model and a byte-preserving save |
| SPEC-117 | The sheet grid — styles, sizes, merges, frozen panes and selection |
| SPEC-118 | Number formats and dates |
| SPEC-119 | The formula language — parse, print and translate formulas between both formats |
| SPEC-120 | Calculation — the engine, recalculation and the core functions |
| SPEC-121 | Create a workbook, and sheet management |
| SPEC-122 | Cell entry and the sheet writer |
| SPEC-123 | Cell formatting — fonts, fills, borders, alignment and number formats |
| SPEC-124 | Rows, columns and merged cells |
| SPEC-125 | Copy, paste, fill and move |
| SPEC-126 | CSV and TSV — open, save and export |
| SPEC-127 | The function library and the function picker |

### Phase 18 — Spreadsheet depth
What real workbooks arrive carrying, and what their owners do with them
next. Each spec here models something SPEC-116 preserved verbatim, so each
bumps the reader version, registers its structure with SPEC-124's reference
adjusters (a row inserted above it must move it), and keeps whatever it does
not model byte-verbatim. SPEC-131 comes before SPEC-132 — a chart is a
drawing anchored to cells first. SPEC-136 needs SPEC-104's move of the PDF
writer into `src/lib/core`; SPEC-138 needs SPEC-103's generalisation of
drafts beyond text and SPEC-137's chunked row storage. SPEC-139 is the
largest single spec in the phase and deliberately last. Two specs need a
person first: SPEC-135 an Excel 365 workbook with threaded comments
(`spec135-excel365-threads.xlsx` — Excel for the web can save one), and
SPEC-137 a quiet factory-container baseline run to replace its projected
budgets before its perf gate is enforced.

| Spec | Title |
| --- | --- |
| SPEC-128 | Sort and filter |
| SPEC-129 | Conditional formatting |
| SPEC-130 | Data validation and in-cell dropdowns |
| SPEC-131 | Images and shapes on sheets |
| SPEC-132 | Charts in workbooks |
| SPEC-133 | Named ranges and tables |
| SPEC-134 | Find and replace, and Go To |
| SPEC-135 | Notes, comments and hyperlinks on cells |
| SPEC-136 | Page setup, printing and PDF export for sheets |
| SPEC-137 | Large workbooks — performance |
| SPEC-138 | Autosave, drafts and version history for workbooks |
| SPEC-139 | Pivot tables |
| SPEC-140 | Workbook generation API |

---

## Backlog — identified, not yet specced
Gaps found in the same review that produced Phases 9–11, listed so they are
not rediscovered from scratch. Each names where it was deferred. None is a
work order; a business person picking the next spec writes one of these
first.

**The presentation rows below were written before Phases 14–16 existed.**
The slide-layout picker and masters, tables and charts on slides, slide size
setup, presenter view, the slide sorter, per-paragraph builds, rotation and
multi-select on slides, embedded audio and video, and deck export to PDF are
now specced (SPEC-089–106); they are left here so the trail from "deferred
by" to "specced by" stays readable, not because they are still open. New
gaps this queue defers: rasterising a PDF page and everything that needs it
(SPEC-107), OCR over a scanned page (SPEC-110), merging `/AcroForm` field
trees and tagged-PDF structure trees (SPEC-111), image-editor depth beyond
crop / rotate / adjust / annotate (SPEC-112), and any model-backed feature
for the word processor (SPEC-115).

| Gap | Deferred by |
| --- | --- |
| Find and replace by formatting or style; wildcards; `$1` capture groups in replacements | SPEC-017's out of scope, "noted as a follow-up" |
| Page-thumbnail navigation pane, and Go to page | SPEC-037's out of scope |
| User-remappable keyboard shortcuts | SPEC-021's out of scope |
| Text wrapping around a floating table | SPEC-023's out of scope; SPEC-076 leaves it |
| Formulas in table cells (`=SUM(ABOVE)`) | SPEC-023's out of scope; SPEC-076 leaves it |
| Citations, bibliography and an index | SPEC-031/033/055 each touch the edge |
| Content controls and fillable form fields (`w:sdt`) | SPEC-053 preserves them, models none |
| Compare and merge two documents | Implied by SPEC-039, never specced |
| RTF, HTML, plain-text and EPUB import/export | SPEC-050 did Markdown only |
| Drop caps | Never specced |
| Automatic hyphenation and justification quality | SPEC-035/036 mention it, neither owns it |
| AutoText / building blocks / quick parts, and a cover-page gallery | Never specced |
| Style sets — alternative formatting for the built-in styles | SPEC-019's out of scope; SPEC-077 resolves only theme switching |
| SmartArt, and Word's drawing canvas (`wpc:wpc`) | SPEC-051 and SPEC-057 both preserve, never create |
| Boolean shape operations, text on a path, WordArt | SPEC-071/073's out of scope |
| Glow, reflection, bevel and 3-D shape effects | SPEC-073's out of scope |
| Marquee / rubber-band multi-select for drawings | SPEC-057's out of scope |
| Forced-colours / high-contrast mode | SPEC-059's out of scope |
| Folders or tags in the library | SPEC-067's out of scope |
| Document protection, read-only enforcement, encryption | Never specced |
| Slide layouts picker and master editing | SPEC-080's out of scope (new slides are title+body only) |
| Tables and charts on slides (beyond foreign-frame preservation) | SPEC-079 renders placeholders only |
| ~~Slide sorter / grid view, copy-paste of slides across decks~~ (shipped by SPEC-099) | SPEC-080's out of scope |
| Transition/build import from Keynote onto SPEC-085/086 models | SPEC-087 drops and counts them |
| ~~Rotation editing, adjustment handles, multi-select on slides~~ (shipped by SPEC-096) | SPEC-082's out of scope |
| Embedded audio/video on slides | SPEC-088 keeps placeholders |
| Keynote ’09 (`index.apxl`), .pages/.numbers products | SPEC-087 rejects with pointers |
| Deck export to PDF; HTML export via the generation API | SPEC-034 is text-only; SPEC-088 names the API follow-up |

**Spreadsheet gaps (Phases 17–18).** Deferred by the spreadsheet queue
itself; none is a work order.

| Gap | Deferred by |
| --- | --- |
| `.xls` (BIFF8), `.xlsb`, `.numbers` and flat `.fods` import; opening password-protected workbooks | SPEC-116 refuses each with a message |
| Running, inspecting or editing macros (VBA is carried verbatim) | SPEC-116 |
| Dynamic arrays and spill ranges — FILTER, SORT, UNIQUE, SEQUENCE, LET, LAMBDA evaluation | SPEC-119/120 show cached spill values only; SPEC-127 |
| Iterative calculation; manual calculation mode beyond what SPEC-120 pins | SPEC-120 |
| External workbook links (refresh, edit links) | SPEC-119 keeps them as opaque references |
| Locale-specific formula syntax (`;` separators, localised function names) | SPEC-118 — formulas are always en-US |
| Partial rich-text formatting inside one cell | SPEC-123 |
| Insert/delete cells with shift (not whole rows/columns); outline grouping of rows/columns | SPEC-124 |
| Paste-special operations (add, subtract, multiply, divide) | SPEC-125 |
| Sort by colour or icon; advanced filter with a criteria range | SPEC-128 |
| Sheet and workbook protection — set, change and remove passwords | SPEC-122 only honours protection |
| Sparklines, slicers and timelines | Never specced |
| What-If analysis — Goal Seek, scenarios, data tables; Solver | Never specced |
| Data connections, Power Query, the data model | Never specced |
| Real-time collaboration, spell check and an accessibility checker for sheets | Never specced |
| Small-screen and touch layout for the grid; right-to-left sheets | SPEC-116/117 |
| The command palette for the spreadsheet view | SPEC-062 is text-only |
| Names, tables, charts, validation and conditional formatting in the generation API | SPEC-140 |
