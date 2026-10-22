---
id: SPEC-055
title: Captions, caption cross-references and the table of figures
status: draft
area: structure
depends_on: [SPEC-033, SPEC-053]
---
## Why
"As shown in Figure 3" is a promise that Figure 3 is still Figure 3 after
someone inserts a screenshot on page 2. Typing the numbers by hand means it
usually is not, which is why every long document that has been through more
than one draft has a wrong figure reference in it somewhere.

SPEC-032 and SPEC-033 both put this out of scope, and SPEC-053 stopped at
cross-references to headings and bookmarks for the same reason: captions need
their own numbering sequence, and a sequence is only useful once something
renumbers it. This spec adds that sequence, the captions built on it, the
cross-references that point at them, and the table of figures they generate —
the same relationship SPEC-033's table of contents has with heading styles.

## Acceptance criteria
- [ ] "Insert → Caption" with an image, table, chart or equation selected inserts a caption paragraph in the Caption style, positioned above or below per the dialog, containing the label, a number, a separator and an editable text field
- [ ] The caption dialog offers the labels Figure, Table and Equation, a "New label…" action for custom labels, a position (above/below), a separator, an "Exclude label from caption" option, and a number format from the list SPEC-029 fixes (1, 2, 3 / i, ii, iii / a, b, c)
- [ ] The number is a `seq` field run (SPEC-053) scoped to its label, and only its number is computed — the label and the caption text are ordinary runs the user can edit freely
- [ ] Inserting, deleting or moving a captioned object renumbers every caption of that label per "Numbering" below, in one undo step with the edit that caused it
- [ ] Two labels number independently: adding a figure never changes a table's number
- [ ] A setting, "Automatically caption new images and tables", inserts a caption on insertion of each object type it is enabled for, with the caret placed in the caption text
- [ ] Deleting a captioned object asks whether to delete its caption too, defaulting to yes, and deleting a caption never deletes its object
- [ ] SPEC-053's "Insert → Cross-reference" dialog gains a category per label, listing that label's captions by their full text, and offers the reference formats in "Reference formats" below
- [ ] A caption cross-reference renumbers with its target: adding a figure above Figure 3 makes every reference to it say Figure 4 on the next field update
- [ ] "Insert → Table of figures" inserts a block listing every caption of one chosen label, with its text, a dot leader and its page number, built and updated exactly the way SPEC-033's table of contents is — generated, not editable, with an update control when it is stale, and offering "update page numbers only" or "update entire table"
- [ ] A document may hold more than one table of figures, one per label, each independent
- [ ] Captions in headers, footers and footnotes are excluded from numbering and from the table of figures — see "Numbering"
- [ ] Captions round-trip through `.docx` (a Caption-styled paragraph containing a `SEQ Figure \* ARABIC` field; the table of figures as a `TOC \c "Figure"` field with cached result text) and `.odt` (`text:sequence-decl` plus inline `text:sequence`; the index as `text:illustration-index`/`text:table-index` with an index body)
- [ ] Caption cross-references round-trip as `.docx` `REF` fields against a bookmark spanning the caption, and as `.odt` `text:sequence-ref` with the matching `text:reference-format`
- [ ] Opening the fixtures below renders every caption at its saved number, resolves each cross-reference, and regenerates a Word-authored table of figures with the same entries

## Out of scope
An index — marking entries and generating one is a separate spec, and shares
none of this machinery beyond the block shape. A bibliography and citations,
still out of scope from SPEC-033. Chapter-based caption numbers ("Figure 2-1"),
which need the chapter numbering SPEC-029 also excludes. Captions on anything
other than the four object types above. Caption styles beyond the built-in
Caption style. Reordering captions independently of their objects. Localised
default labels — the three built-in labels are English, and a user who needs
"Abbildung" creates it with "New label…".

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** The facts this spec rests on: `DocumentModel` holds
`sections: Section[]` whose blocks are
`Block = Paragraph | Table | Image | BreakBlock | TableOfContents`; a
`Paragraph` holds `runs: Run[]` with no `Inline` union; lengths are centimetres
and type sizes points; and equality drops every key named `id`, so a reference
that must survive a round-trip test is never called `id`.

### Fields
Defined normatively in `specs/053-fields-date-time-and-cross-references.md`,
built on the shipped mechanism in `src/lib/documentModel.ts`: a field is a
**run** whose `text` is `FIELD_PLACEHOLDER_TEXT` (one character, so it is a
single caret atom) carrying `field?: FieldSpec`, a union discriminated on
`kind`. Members hold a `cached` string — what other applications display —
which `equalStructural` compares, and fields update only when asked.

This spec adds one member to that union:

```ts
| { kind: "seq"; label: string; format: NumberFormat; cached: string }
```

| Field | Shows | `.docx` | `.odt` |
| --- | --- | --- | --- |
| `seq` | the item's number in its label's sequence | `SEQ <label> \* <format>` | `text:sequence` |

and widens the `ref` member's `show` with the caption formats below. It adds no
other kind and no new update path — "Update all fields" already covers these.

### A caption is a paragraph, not a node
A caption is an ordinary paragraph carrying the built-in Caption style
(SPEC-019), whose children are: a run holding the label text, a `seq` field
run, a run holding the separator, and the user's runs. There is no `Caption` block
kind.

This matters because it is what both formats actually store, and because it
keeps every existing feature working on captions for free — a caption is
findable (SPEC-017), counted (SPEC-036), spell-checked (SPEC-035), commentable
(SPEC-038) and styleable (SPEC-042) with no special cases. "Exclude label from
caption" simply omits the label run.

### Caption-to-object binding is positional
Neither `.docx` nor `.odt` stores a link from a caption to the object it
describes. A model field holding one would have to be invented on read and
discarded on write, and would be wrong the moment the file came back from Word.

So the binding is **computed, not stored**: a caption belongs to the object in
the adjacent paragraph — the paragraph immediately before it for a caption
positioned below, immediately after it for one positioned above — when that
paragraph's only content is a single `Drawing` or is a table, and the caption
paragraph carries the Caption style and a `seq` field run.

The consequences are stated rather than hidden. Typing a sentence between an
image and its caption breaks the binding, and the caption becomes an ordinary
Caption-styled paragraph that still numbers and still appears in the table of
figures — it just stops being offered for deletion with the object. That is the
honest behaviour, and it is what Word does too.

### Numbering
`seq` fields are numbered per label, in document order, starting at 1, when
"Update all fields" runs or when an edit renumbers them.

**Document order excludes headers, footers and footnotes.** A header repeats on
every page, so a caption in one has no single position in the sequence and no
single page number for the table of figures. Captions there keep their
`cached` value and are skipped by the counter — a rule worth naming because
the alternative fails silently and produces a document with two Figure 4s.

Renumbering runs over the whole document rather than incrementally from the
edit point. A 300-caption document renumbers within SPEC-033's one-second TOC
budget, and the simplicity is worth more than the saved milliseconds.

### Reference formats
A caption cross-reference stores its target in `target` and its format in
`show`:

| `show` | Renders | `.docx` | `.odt` `text:reference-format` |
| --- | --- | --- | --- |
| `full` | "Figure 3: Revenue by quarter" | `REF` | `text` |
| `labelAndNumber` | "Figure 3" | `REF … \r`-style caption subset | `category-and-value` |
| `numberOnly` | "3" | same, number subset | `value` |
| `textOnly` | "Revenue by quarter" | same, text subset | `caption` |
| `page` | the page it is on | `PAGEREF` | `page` |
| `aboveBelow` | "above" / "below" | `REF … \p` | `direction` |

A reference whose target has been deleted renders "Error! Reference source not
found." and is listed by the update, exactly as SPEC-053 already specifies for
heading and bookmark references. Referencing a caption inserts a hidden
bookmark spanning it if it has none, the same mechanism SPEC-053 uses for
headings.

### The table of figures
A block kind of its own, mirroring the table of contents rather than reusing
it — the two select different content and a shared block would need a
discriminator anyway:

```ts
// src/lib/documentModel.ts — mirrors the shipped TableOfContents
export type Block =
  | Paragraph | Table | Image | BreakBlock | TableOfContents
  | TableOfFigures                          // added by this spec

export interface TableOfFigures {
  type: "tof"
  /** "Figure", "Table", or a custom label. */
  label: string
  options: { pageNumbers: boolean; dotLeader: boolean; hyperlinks: boolean }
  /** The generated result, cached the way a field's value is. */
  entries: TofEntry[]
  stale: boolean
}
export interface TofEntry { text: string; page: number; targetBookmark: string }

export function isTof(block: Block | undefined): block is TableOfFigures {
  return block?.type === "tof";
}
```
The shipped `TableOfContents` is `{ type: "toc"; options: TocOptions; entries:
TocEntry[] }`, and `isToc` is its narrowing helper; this mirrors both,
including the `isX` helper that `.vue` scripts need because espree cannot parse
a type predicate.

Its behaviour is SPEC-033's, restated so a reader of this file alone does not
have to guess: generated from document content and not directly editable;
typing inside it shows the update hint; it shows an update control when the
content it summarises has changed since generation; updating offers "page
numbers only" or "entire table"; each entry is clickable and scrolls to its
target; deleting it leaves the captions untouched; and it writes with cached
result text so Word and Writer display it without prompting for an update.

Dot leaders come from the same place SPEC-033's do. If
`specs/054-tab-stops-and-the-tab-ruler.md` has merged, that is a right tab stop
with a `dot` leader on the entry paragraph, which is what both formats store.

### Specs this changes
- `specs/053-fields-date-time-and-cross-references.md` — its field catalogue
  gains `seq`, and its cross-reference dialog gains the caption categories and
  the `show` values above. No change to its update rules, its determinism
  rule or its unsupported-field handling.
- `specs/033-table-of-contents.md` — no change. The table of figures is a
  separate block that copies its behaviour; SPEC-033's out-of-scope line
  "Table of figures, index, and bibliography" is resolved for the first of the
  three by this spec, and the other two stay out of scope.
- `specs/032-multilevel-lists-and-numbering.md` — its out-of-scope line
  "Numbering of captions, figures and tables" is resolved by this spec.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the Insert menu gains
  "Caption…" and "Table of figures…".

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic`, **which is not implemented**; this spec does not need it.
- **`equalStructural`** — for the same-format round trip. It covers the `seq`
  fields' `cached` values and the `TableOfFigures` block's `entries`, so a
  writer that emits a stale or empty generated result fails rather than passing
  on a technicality. SPEC-053's rule that opening a document never updates a
  field is what makes this hold.
- For cross-format saves, assert `equalStructural` between the source model
  masked by the writer's `dropped` records and the re-opened model. Masking is
  needed because the reference formats are not symmetric: ODF's `category-and-value` has no single
  `.docx` field switch, and a custom label maps to a `text:sequence-decl` that
  Word represents only as the `SEQ` argument.

`equalBytes` is not used by this spec.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). Adds
`fixtures/docx/word-captions.docx` and `fixtures/odt/lo-captions.odt`, each
with three figure captions and two table captions interleaved, one caption
above its object and one below, one custom label, one caption separated from
its object by a paragraph, one cross-reference of each format, one broken
cross-reference, and a generated table of figures. Manifest entries name each
caption's label, number and text, and the table of figures' expected entries.
