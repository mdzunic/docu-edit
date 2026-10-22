# Presentations: architecture of the second product

Written 2026-08-24, when the repo was prepared for a presentations product
(`.pptx`/`.odp`). This file records the layering that preparation
established, what already ships, and where future presentation specs plug
in. The behavioural source of truth remains the spec set; this is the map.
The illustrated engineering note (layer diagram, refactor steps,
verification record) is `presentations-engineering-note.html` beside this
file — open it in a browser.

## What ships today (the preparation milestone)

Open a `.pptx` or `.odp` from any open path (library picker, drag-drop,
`/open?url=`), see it in the library with its own badge, filter chip and
slide count, open a **read-only preview** at `/slides/:id` (title + text
outline per slide), and download it back — with **every part of the source
package re-emitted byte-identical** (SPEC-008's retention rule applied
wholesale, `.odp`'s `mimetype` entry first and STOREd).

Since **SPEC-080** the deck is no longer read-only: create a blank deck from
the landing page ("New presentation", `createBlankPresentation`, `pptx` from
birth), and add / duplicate / delete / reorder slides from the filmstrip, all
undoable (a per-view stack capped at 100, `usePresentationUndo`) and persisted
write-through. The writers gained two paths beside verbatim re-emission — a
**from-scratch** minimal-package build (a deck with no retained `opaque`) and a
**part-substitution** rebuild that rewrites only the structural parts
(`presentation.xml`/its rels/`[Content_Types].xml` for pptx; `content.xml`
spliced from the page envelope for odp) while every untouched slide part /
`draw:page` re-emits byte-verbatim (`presentation/slides.ts` commands are the
normative source; SPEC-080 the normative source for the rebuild rules and the
undo stack). Text and object editing inside a slide are still to come
(SPEC-081/082).

## Layering

```
            ┌───────────────────────────────────────────────────────────┐
            │ src/lib/core — product-neutral (boundary.spec.ts)         │
            │ zip/zipWriter · xml (parse/serialise/escape/OPC rels)     │
            │ packageDetect · packageParts (OpaquePart + base64)        │
            │ documentProperties · opcProperties · odfMeta · odfXml     │
            │ triggerDownload · media · units · shapes/charts/tables …  │
            └───────────────┬───────────────────────┬───────────────────┘
        ┌───────────────────┴─────────┐   ┌─────────┴──────────────────┐
        │ TEXT product                │   │ PRESENTATION product       │
        │ documentModel.ts, commands, │   │ src/lib/presentation/      │
        │ textEditing, pagination …   │   │   model.ts (PresentationModel,
        │ formats: docx/ odf/ md/     │   │   SlidePreview, ParsedPresentation)
        │                             │   │   pptx/reader+writer       │
        │                             │   │   odp/reader+writer        │
        │                             │   │   open.ts · download.ts    │
        │                             │   │   fixtures/build.ts        │
        └───────────────────┬─────────┘   └─────────┬──────────────────┘
                            │      app layer        │
        ┌───────────────────┴───────────────────────┴──────────────────┐
        │ documentKinds.ts (kind registry: format→kind, routes, accept) │
        │ openDocument.ts (openAnyDocumentFile funnel, single-unzip     │
        │   detection) · openDocumentFromUrl.ts                        │
        │ documentStore.ts (StoredModel union, kind-narrowed accessors) │
        │ libraryView/libraryPrefs/libraryMenu · format.ts (badges)     │
        └───────────────────┬───────────────────────┬──────────────────┘
                    ┌───────┴───────┐       ┌───────┴────────────┐
                    │ /doc/:id      │       │ /slides/:id        │
                    │ DocumentView  │       │ PresentationView   │
                    └───────┬───────┘       └───────┬────────────┘
                            └───── HomeView ────────┘
                              + src/views/library/* (the library module)
```

### Boundary rules (all enforced as tests)

- **core** imports only sibling core modules — `src/lib/core/boundary.spec.ts`.
- **presentation** imports only presentation siblings and core — never
  `documentModel.ts`, commands, `docx/`, `odf/`, the store, or Vue —
  `src/lib/presentation/boundary.spec.ts`.
- **docx/ and odf/** never import each other **nor presentation/**; the
  text model layer imports no format module — the format-layer block in
  `core/boundary.spec.ts`.
- The two products meet only in the app layer: `documentKinds.ts`,
  `openDocument.ts`, `documentStore.ts`, the router and the views. Those
  are the composition points; they may import both.

### What moved to core in this preparation (and why)

Everything a `.pptx`/`.odp` needs that the word processor already had, none
of it word-shaped:

| Module | From | What it is |
| --- | --- | --- |
| `core/documentProperties.ts` | `documentModel.ts` | DocumentProperties/CustomProperty/… metadata types (re-exported from documentModel, so text-side imports are unchanged) |
| `core/packageParts.ts` | `documentModel.ts` + `opaque.ts` | `OpaquePart` + the base64 codecs + `decodedPartBytes` |
| `core/opcProperties.ts` | `docx/docProps.ts` | `docProps/core.xml`/`custom.xml` parse+build — OPC-level parts, identical in `.pptx` |
| `core/odfMeta.ts` | `odf/meta.ts` | `meta.xml` parse+build — ODF-package-level, identical in `.odp` |
| `core/odfXml.ts` | `odf/xml.ts` | ODF namespaces + attr readers (+ `PRESENTATION_NAMESPACE`) |
| `core/packageDetect.ts` | new | container classification: ODF `mimetype` / OPC main part + `[Content_Types].xml` |
| `core/triggerDownload.ts` | `downloadDocument.ts` | the object-URL download mechanism |

Already in core before this (the 2026-08-15 extraction): zip/zipWriter,
xml/OPC relationships, media content-types, shape/chart/table/list
geometry + painters, units, schedulers, idb. Presentations reuse all of it.

## The kind registry (`src/lib/documentKinds.ts`)

A **kind** is a product (`"text" | "presentation"`); a **format** is a file
type within one. `kindOfFormat` and `editorPathFor` are the single place
"pptx means the presentation product at /slides/:id" is encoded. Everything
kind-dispatched goes through it: the router guards (each editor route
redirects the other kind's ids), `HomeView`/`AppShell`/`OpenFromUrlView`
routing, `useLibrary.downloadDocuments`, the pickers' accept list
(`OPEN_FILE_ACCEPT`). Adding a future kind = extend this module + a new
product layer.

## The store (`documentStore.ts`)

One store, one record shape, a **model union**:
`StoredModel = DocumentModel | PresentationModel`, discriminated by
`PresentationModel.kind` (a text model has no `kind` field). The persistence
machinery (write-through, media split/rejoin latches, trash, drafts,
storage accounting) is shared by construction — a presentation model
carries `id`/`title`/`media?` with the same conventions. Kind safety:

- `getDocument` returns text models only; `getPresentation` the reverse;
  `documentKindOf` answers the routing guards.
- `getOrCreateDocument` throws on a presentation id (unreachable through
  the guarded views; failing loudly beats silently overwriting a deck).
- Drafts/autosave are text-editor features; `saveDraft` throws on a
  presentation record and `useAutosave.beginSession` bails out early.

`LibraryEntry` carries `kind`, so the library UI routes and renders
per-kind without re-deriving anything.

## The open funnel (`openDocument.ts`)

`openAnyDocumentFile` unzips once, classifies via `core/packageDetect`, and
dispatches: ODF-text mimetype → `parseOdt`; ODF-presentation mimetype →
`parseOdp`; OPC PresentationML main part → `parsePptx`; any other ZIP → the
`.docx` path (which owns its own rejections, legacy `.doc` included);
non-ZIP → flat-ODF rejection / Markdown by extension (SPEC-050's
documented exception) / the `.docx` path. Detection never trusts
extensions. `openDocumentFile` keeps its text-only contract for template
and style import — a presentation handed there gets a specific error, never
a misparse. Unknown ODF kinds (e.g. `.ods`) now get an honest message
instead of the generic docx one.

## The presentation model (deliberately minimal)

`PresentationModel` = `kind` + `id`/`title`/`format` + `slides:
SlidePreview[]` (read-only outline: optional title + text lines) +
`opaque: OpaquePart[]` (**the whole package, verbatim, in order**) +
optional `properties`/`media`. The writers re-emit `opaque`; round-trip is
part-level byte equality (the ZIP container is rebuilt, so container bytes
may differ — the tests compare parts, same as the docx/odf suites).

Editing specs will grow this the way SPEC-011+ grew the text model: model
real slide content, and shrink `opaque` to the parts that stay unmodelled.
Big reuse candidates when that happens, deliberately **not** moved yet
(they need a design pass, not a move): the DrawingML shape/chart parsing in
`docx/drawing|shape|chart.ts` (pptx shapes are the same DrawingML), the
`text:` namespace paragraph parsing in `odf/reader.ts` (odp text boxes use
it), `textMetrics`' wrap engine, and a first-class OPC relationship writer.

## The landing page module (`src/views/library/`)

`HomeView.vue` (1,476 → 651 lines, ratcheted by `HomeView.size.spec.ts`)
is now a wiring hub over product-neutral components: `LibraryControls`
(header/search/sort/chips; product actions arrive via a slot),
`LibraryList`, `LibraryTrash`, `LibrarySelectionBar`,
`LibraryStoragePanel`, `LibraryEmptyState` — plus the pre-existing
`components/library/*` cards. Shared logic: `useLibrary` (now also
`kindOf`, `rowTabindex`, kind-dispatched downloads) and `lib/libraryMenu`
(row/trash menu items). The components render entries and emit intents;
only HomeView knows what can be created and where each kind routes.

## Testing

- `src/lib/presentation/roundTrip.spec.ts` — readers (outline, metadata,
  retention, cross-format rejection), writers (part-identical round trip,
  mimetype-first+STOREd), funnel classification, JSON-serialisability.
- `src/lib/presentation/boundary.spec.ts` + the extended
  `core/boundary.spec.ts` — the layering, as tests.
- `e2e/presentations.spec.ts` — open → library badge → /slides preview →
  kind-guard redirect → byte-identical download, over the checked-in
  `e2e/fixtures/mini-deck.pptx|odp` (regenerate with
  `npx vite-node scripts/build-presentation-fixtures.ts`; the fixtures and
  the unit suite share one builder, `presentation/fixtures/build.ts`).
- `src/views/HomeView.size.spec.ts` — the landing page's regrowth ratchet.

## Updated behavioural strings (tests updated with them)

The picker `accept` is now `.docx,.odt,.md,.markdown,.pptx,.odp`
(SPEC-004/005 e2e), and the drop error reads "We can only open .docx,
.odt, .md, .pptx and .odp files" (SPEC-009/050 e2e, AppShell spec).
