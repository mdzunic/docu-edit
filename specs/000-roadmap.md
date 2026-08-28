---
id: SPEC-000
title: Product roadmap — word processor, presentations and media
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
   `specs/113-experimental-features-and-the-model-gateway.md`. Nothing
   redefines these.

2. **Every dependent spec ends with a `## Reference` section** that names the
   dependency spec files by path, restates the contract it relies on inline
   (model shape, function signatures, fixture layout), and names the equality
   relation each test uses. A reader of one file alone must never have to guess.

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
| Presenter view — notes, next slide, timer | SPEC-083/084 both name it |
| Slide sorter / grid view, copy-paste of slides across decks | SPEC-080's out of scope |
| Per-paragraph text builds, emphasis effects, motion paths | SPEC-086's out of scope |
| Transition/build import from Keynote onto SPEC-085/086 models | SPEC-087 drops and counts them |
| Rotation editing, adjustment handles, multi-select on slides | SPEC-082's out of scope |
| Embedded audio/video on slides | SPEC-088 keeps placeholders |
| Keynote ’09 (`index.apxl`), .pages/.numbers products | SPEC-087 rejects with pointers |
| Deck export to PDF; HTML export via the generation API | SPEC-034 is text-only; SPEC-088 names the API follow-up |
