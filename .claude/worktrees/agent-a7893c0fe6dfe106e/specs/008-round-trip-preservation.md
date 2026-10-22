---
id: SPEC-008
title: Round-trip preservation of unsupported content
status: done
area: io
depends_on: [SPEC-006, SPEC-007]
---
## Why
Every editor that rewrites a file it only partly understands eventually
destroys someone's work — a chart, a content control, a custom XML part gone
because we did not have a feature for it. The rule for this product: what we
cannot edit, we carry through untouched, and we tell the user when we cannot.

## Acceptance criteria
- [ ] Opening a document retains the original package bytes for the lifetime of the document
- [ ] Package parts the reader does not model (e.g. `customXml/`, embedded objects, charts, themes, `word/webSettings.xml`) are written back byte-identical on save to the same format
- [ ] XML elements inside `document.xml` / `content.xml` that the reader does not model are retained as opaque nodes in the document model, anchored to their position, and re-emitted in place on save
- [ ] Opening a file and saving it to the same format without editing produces `equalStructural(openedModel, reopenedModel)` and an `equalBytes` result for every preserved opaque part — see "Equality" below
- [ ] Editing one paragraph and saving does not alter unrelated opaque content elsewhere in the document
- [ ] Both writers return `WriteResult` — `{ bytes, dropped }` — instead of a bare package, where `dropped` is a `DropRecord[]` naming every item the writer could not represent in the target format — see "Writer API" below
- [ ] Content that cannot be preserved when converting between formats is listed in a "Conversion report" dialog shown once after a cross-format save, naming each dropped item type and its approximate location (e.g. "Chart on page 3 — not converted"); the dialog renders `dropped` and derives no part of its contents by comparing models
- [ ] The conversion report has a "Don't show again for this document" checkbox
- [ ] A test opens a `.docx` containing a chart, a content control and a custom XML part, edits one paragraph, saves, and asserts all three are still present and intact in the output
- [ ] The cross-format test asserts `dropped` deep-equals the exclusion list committed for that direction, then asserts `equalSemantic` between the source model with those items removed and the re-opened model — so the report and the test read the same array and cannot drift apart

## Out of scope
Rendering the preserved content (a chart still shows as a placeholder box).
Editing preserved content. Macro (`vbaProject.bin`) preservation is included by
the part rule above but macro execution is never supported.

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model"), and in `docs/document-model.md` once SPEC-002 is merged.
**Read one of those before starting.** Preserved content is held in three
places, and **none of them is a union member** — each is a side-list, so
`blocks` stays a renderable union and `runs` stays a plain `Run[]`:
```ts
interface DocumentModel { …; opaque?: OpaquePart[]; sourceXml?: SourceXml }
interface Section   { …; opaqueBlocks?: OpaqueBlock[] }
interface Paragraph { …; opaqueInlines?: OpaqueInline[] }

interface OpaqueBlock  { xml: string; blockIndex: number }  // before blocks[blockIndex]
interface OpaqueInline { xml: string; runIndex: number }    // before runs[runIndex]
interface OpaquePart   { name: string; data: string }       // base64 part bytes
```
An index equal to the list's length means "after the last one". `OpaquePart.data`
is base64 rather than a `Uint8Array` because the model must survive
`JSON.stringify`. Opaque XML is compared with `canonicalizeXml`, so namespace
prefixes, attribute order and insignificant whitespace cannot cause false
failures — which is why `xml` is on SPEC-002's XML-valued key list.

### Existing I/O surface
```ts
// src/lib/docx/reader.ts  parseDocx(data): Promise<ParsedDocxDocument>        — SPEC-004
// src/lib/odf/reader.ts   parseOdt(data):  Promise<ParsedOdtDocument>         — SPEC-005
// src/lib/docx/writer.ts  writeDocx(model, options?): Promise<WriteResult>    — SPEC-006, widened here
// src/lib/odf/writer.ts   writeOdt(model, options?):  Promise<WriteResult>    — SPEC-007, widened here
// src/lib/opaque.ts       WriteResult, DropRecord, DropKind — defined below
```
This spec makes preserving opaque content the default and unconditional
behaviour rather than an option flag — neither `WriteDocxOptions` nor
`WriteOdtOptions` carries a `preserveOpaque` switch. It extends both readers to
produce opaque parts and nodes instead of dropping content (`ParsedDocxDocument
.opaque` and `.sourceXml`), and widens both writers' return type from a bare
package to `WriteResult`. SPEC-006 and SPEC-007 ship the bare-package form;
migrating both writers and their existing call sites is part of this spec's
work. Read `specs/006-save-as-docx.md` and `specs/007-save-as-odt.md` first.

### Writer API — where the exclusion list comes from
`equalSemantic` is a boolean relation (`equalStructural(resolve(a), resolve(b))`)
and produces no diff, so it cannot be the source of the conversion report. Nor
could any post-hoc model diff be: once a chart has been dropped it is simply
absent, and a diff over the models can only report a missing node at a path, not
"a chart, on page 3". The writer performing the conversion is the only component
that knows both facts at the moment it decides an item cannot be represented, so
it is the single source both consumers read:
```ts
// src/lib/opaque.ts
export type DropKind =
  | 'chart' | 'contentControl' | 'customXml' | 'embeddedObject'
  | 'macro' | 'opaqueBlock' | 'opaqueInline' | 'opaquePart'
export interface DropRecord {
  kind: DropKind
  location: string   // human-readable, e.g. "page 3", 'package part "customXml/item1.xml"'
  reason: string     // why the target format cannot carry it
}
export interface WriteResult { bytes: Uint8Array; dropped: DropRecord[] }
```
`bytes` rather than a `Blob` because the writers carry no DOM dependency
(SPEC-006); wrapping them for download is the caller's job.

**`DropKind` is an open list, widened by the spec that needs a new member** —
never reused loosely because an existing member is close enough. Later specs
have already added `runShading` (SPEC-013), `svgImage` (SPEC-024),
`trackedPropertyChange` (SPEC-039) and `shapeStyle` (SPEC-073 — a shape
styling property one format cannot express); each widening belongs in the PR of the spec
that introduces it, and `DROP_KIND_LABELS` in the same module gains the
matching user-facing label so the conversion report needs no change.
`dropped` is empty for a same-format save — opaque content round-trips
untouched, so there is nothing to report. The conversion report dialog renders
`dropped` directly; the cross-format test compares `dropped` against the
committed exclusion list for that direction and uses it to mask the source model
before calling `equalSemantic`. Neither side re-derives the list.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). This
spec adds `fixtures/docx/word-charts-and-controls.docx` — containing a chart, a
content control and a custom XML part — and `fixtures/odt/lo-embedded.odt`,
with manifest entries listing the parts that must survive.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"),
implemented in `src/lib/equality.ts`, whose XML canonicalization is exported as
`canonicalizeXml`. All three relations are used here — note that
`equalSemantic` is specified by SPEC-002 but **not yet implemented**, so the
first spec to need it builds it (see SPEC-002's "Implementation status"):
- **`equalBytes`** — for each preserved opaque part, before and after a save.
  This is the only place byte equality is required.
- **`equalStructural`** — for the no-edit same-format round trip, and for
  asserting that editing one paragraph left every other node untouched.
- **`equalSemantic`** — for cross-format saves, paired with the exclusion list
  this spec surfaces to the user as the conversion report. `equalSemantic`
  returns a boolean and generates no list of its own; the list is the writer's
  `dropped: DropRecord[]` (see "Writer API" above). The report renders that
  array and the test asserts against it, so the two cannot drift apart.
  Apply the exclusions as a transform on the source model —
  `equalSemantic(withoutDropped(source, dropped), reopened)` — never as a
  hand-rolled subset comparison.

### Determinism — pin `modified` whenever two saves are compared
Both writers stamp the modification timestamp (`dcterms:created`/
`dcterms:modified` in `docProps/core.xml`, `dc:date` in `meta.xml`) from
`options.modified`, which **defaults to `new Date()`**. A save is therefore not
a pure function of the model, and anything that compares two saves — or that
commits a save's bytes — must pass a fixed `modified`:

- **Round-trip and byte-equality tests**: pass the same fixed `Date` to every
  `writeDocx`/`writeOdt` call being compared. Without it a save→reopen→save
  assertion can pass all day and fail when the two saves land either side of
  midnight, or compare unequal bytes purely because the clock moved between
  them. `equalStructural` ignores the stamp, but `equalBytes` and any
  whole-part XML comparison do not.
- **Fixture generators** (`scripts/build-spec*-fixtures.*`): pin it too, so
  re-running a generator against an unchanged model rewrites the same content
  instead of churning the committed bytes with the day it was re-run — a diff
  that says nothing about the change being made and hides the ones that do.

Related: a generator is a script, run deliberately. Never regenerate a
committed fixture as a side effect of `npm test` — the e2e suite loads those
same files, so a unit run and an e2e run in one checkout can collide over a
half-written package. A spec that wants to guard a fixture should read the
committed bytes and compare (see `src/lib/fixtures.fields.spec.ts`), leaving
the writing to `scripts/`.
