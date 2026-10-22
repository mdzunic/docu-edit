---
id: SPEC-002
title: Create a new blank document — canonical document model
status: draft
area: core
depends_on: [SPEC-001]
---
## Why
The cheapest way to prove the document model works end to end is to make one
from nothing. This spec is also the **normative definition of the canonical
document model, its units, and the three equality relations** used by every
round-trip test in the product. Later specs cite this file rather than
redefining these; if this spec and another spec disagree, this spec wins.

It describes `src/lib/documentModel.ts` and `src/lib/equality.ts` as they are
built. Between this spec and that code there is no winner to declare — they are
meant to say the same thing, and a difference is a bug in one of them. Changing
the model means changing both in the same PR.

## Acceptance criteria
- [ ] Clicking "New document" in the empty state creates a document and navigates to `/doc/:id`
- [ ] A new document contains exactly one empty paragraph and renders a page-shaped canvas (white page, visible margins) centred in the canvas area
- [ ] The default title is "Untitled document" and appears in the app bar
- [ ] Clicking the title in the app bar makes it editable; typing a new title and pressing Enter or blurring updates it, and Escape reverts it
- [ ] A title of only whitespace reverts to "Untitled document"
- [ ] The model is implemented in `src/lib/documentModel.ts` exactly as specified in "Canonical document model" below, including every node kind, field name and unit
- [ ] `docs/document-model.md` is committed, documents the model, the unit rules and the equality relations, and is the file later specs link to
- [ ] `src/lib/equality.ts` exports `equalBytes`, `equalStructural`, `firstStructuralDifference` and `canonicalizeXml` with the semantics defined in "Equality" below, and `resolve` and `equalSemantic` when the first spec needing them lands — see "Implementation status"
- [ ] `firstStructuralDifference` reports the path of the first difference (e.g. `sections[0].blocks[3].runs[1].formatting.bold`) and `equalStructural` is defined as that returning `null`, so failing round-trip tests say where they broke
- [ ] Unit tests cover: creating a document, reading/writing its title, and JSON serialise → deserialise producing an `equalStructural` document
- [ ] Each equality rule listed below has its own unit test asserting that rule alone
- [ ] Unit tests assert `equalStructural` is reflexive, symmetric and transitive over hand-built models, and that a difference in any single model field is reported as unequal

## Out of scope
Typing into the document (SPEC-011). Reading or writing files (SPEC-004 to
SPEC-007). Persistence across reloads (SPEC-003). Templates (SPEC-046).
Block kinds beyond paragraph — tables (SPEC-022), images (SPEC-024) and
footnotes (SPEC-031) each extend the model, and each must update
`docs/document-model.md` in the same PR, plus the XML-valued key list in
"Equality" below if the new field holds XML.

## Reference

### Canonical document model
Implemented in `src/lib/documentModel.ts`; this section and that file say the
same thing, and both change together. A plain-JSON tree — no class instances,
no DOM nodes, no cyclic references — that must survive `structuredClone` **and
`JSON.stringify`**, which is why binary payloads are base64 strings rather than
`Uint8Array`.

Fields marked with a later spec number are absent until that spec lands, so a
blank document has only the first five. The listing is the shape, not the whole
field set: per-feature fields each spec adds (SPEC-039's `changes`, SPEC-023's
cell properties, and so on) are defined in those specs and documented in
`docs/document-model.md`.

```ts
interface DocumentModel {
  id: string                    // runtime only, excluded from equality
  title: string
  sections: Section[]           // at least one
  defaultStyles: DefaultStyles
  format?: DocumentFormat       // "docx" | "odt"; absent for a never-saved document
  // Every field below is added by a later spec and absent until it lands:
  styles?: NamedStyle[]         // SPEC-019   listDefinitions?: ListDefinition[]  // SPEC-032
  notes?: NoteContent[]         // SPEC-031   comments?: CommentThread[]          // SPEC-038
  media?: MediaPart[]           // SPEC-024   properties?: DocumentProperties     // SPEC-041
  opaque?: OpaquePart[]         // SPEC-008   sourceXml?: SourceXml               // SPEC-008
}

interface Section {
  blocks: Block[]
  opaqueBlocks?: OpaqueBlock[]  // SPEC-008 — a side-list, not a Block member
  pageSize?: PageSize           // absent falls back to defaultStyles
  margins?: Margins
  // SPEC-028/029/030 add: start, columns, headers, footers, pageNumberStart, …
}

type Block = Paragraph | Table | Image | BreakBlock | TableOfContents

interface Paragraph {
  type: "paragraph"
  runs: Run[]                   // [] means an empty paragraph
  formatting?: ParagraphFormatting
  styleId?: string              // raw style reference, kept beside the baked formatting
  headingLevel?: number         // 1-based, resolved from the style chain
  list?: ListFormatting         // SPEC-018
  opaqueInlines?: OpaqueInline[]     // SPEC-008 — a side-list, not a run kind
  bookmarks?: BookmarkAnchor[]       // SPEC-025 — likewise, by character offset
}

interface Run {
  text: string                  // tabs are "\t", forced line breaks are "\n"
  formatting: RunFormatting     // required, but may be {}
  field?: FieldSpec             // SPEC-029 — a computed value, see below
  note?: NoteRef                // SPEC-031 — mutually exclusive with `field`
  styleId?: string              // character style (SPEC-042)
}

interface RunFormatting {       // every field optional; absent = inherit
  bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean
  fontFamily?: string           // as authored, e.g. "Calibri"
  fontSizePt?: number
  color?: string                // CSS-ready "#rrggbb"
  highlight?: string | false    // palette name, "#rrggbb", or false = explicit none
  vertAlign?: "subscript" | "superscript"
}

interface ParagraphFormatting {
  alignment?: "left" | "center" | "right" | "justify"
  indentStartCm?: number; indentEndCm?: number; indentFirstLineCm?: number
  spaceBeforePt?: number; spaceAfterPt?: number
  lineSpacing?: number           // a multiplier
  lineSpacingPt?: number         // an absolute height; wins over lineSpacing
  lineSpacingRule?: "exact"      // absent means "at least"
  direction?: "ltr" | "rtl"
  // SPEC-028 adds keepNext, keepLines, pageBreakBefore, widowControl;
  // SPEC-032 adds tabStops.
}

interface DefaultStyles {
  pageSize: PageSize            // { widthCm, heightCm }
  margins: Margins              // { topCm, rightCm, bottomCm, leftCm }
  fontFamily: string
  fontSizePt: number
  language?: string             // SPEC-035
}
```

**Units — normative, no exceptions.**
- **Lengths are centimetres**, as a `number`, and the field name says so:
  `indentStartCm`, `widthCm`, `columnWidthsCm`, `paddingCm`. **Type sizes and
  vertical spacing are points**, likewise named: `fontSizePt`, `spaceBeforePt`,
  `lineSpacingPt`. A length or size field whose name carries no unit suffix is
  a bug. Readers convert from twips, EMU and ODF's unit-suffixed strings on the
  way in; writers convert back on the way out. Neither unit ever reaches the
  model.
- Fractional values are expected — `1.27` cm is the indent step — so no
  rounding is applied on the way in. Round-trip tests use the 0.01 cm tolerance
  SPEC-014 defines rather than exact equality on converted lengths.
- Colours are CSS-ready lowercase `"#rrggbb"` strings, so a renderer can use
  one directly. `highlight` is the one exception and carries three cases: an
  OOXML palette name (`"yellow"`), a `"#rrggbb"` for anything outside that
  palette, or `false` for an explicit "No colour" — distinct from absent.
- Binary payloads (`MediaPart.data`, `OpaquePart.data`) are **base64 strings**,
  which is what keeps `JSON.stringify` viable.
- A blank document is A4 (`21 × 29.7` cm), 2.54 cm margins, portrait, product
  default body font.

**Two structural decisions worth stating**, because both differ from the
obvious design and later specs depend on them:

- **Preserved content is a side-list, not a union member.** `OpaqueBlock` sits
  in `Section.opaqueBlocks` with a `blockIndex`, and `OpaqueInline` in
  `Paragraph.opaqueInlines` with a `runIndex` — each meaning "immediately
  before this index", with `length` meaning "after the last one". So `blocks`
  stays a renderable union and `runs` stays a plain `Run[]`: no renderer,
  editing primitive or block walk has to narrow past content it cannot draw.
  `BookmarkAnchor` and `CommentAnchor` follow the same pattern by character
  offset.
- **There is no `LineBreak` node and no `Inline` union.** A run's `text` is the
  text, with `"\t"` for a tab and `"\n"` for a forced line break, rendered with
  `white-space: pre-wrap`. Caret offsets are UTF-16 code-unit indices into that
  string (SPEC-026), which is only simple because there is nothing else in it.
  A field or note mark is a run whose `text` is a single-character placeholder
  (`FIELD_PLACEHOLDER_TEXT`) carrying `field`/`note`, so it occupies exactly one
  caret position and every offset calculation stays arithmetic.

**Invariants.** A `DocumentModel` has ≥1 section; a section has ≥1 block; a
paragraph's `runs` may be empty. `Run.formatting` is always present, though it
may be `{}`. `id` is regenerated freely and never compared — and any *other*
reference that must survive a round-trip test is deliberately **not** named
`id` (`ListFormatting.defId`, `NoteRef.noteId`, `BookmarkAnchor.name`), because
equality drops that key by name. Zero-length runs and adjacent runs with equal
formatting are removed by the editing primitives in `src/lib/textEditing.ts`,
not by equality — see below.

### Equality
Three relations, referred to **by name** in later specs. No round-trip
assertion may be written without naming one of them.

1. **`equalBytes(a, b)`** — the two saved packages are byte-identical. Used
   only for preserved opaque parts (SPEC-008).

2. **`equalStructural(a, b)`** — deep equality with the normalisations below
   applied during the walk rather than as a separate `canonicalize` pass, so
   there is no intermediate copy and the first difference can be reported with
   its path. Implemented as `firstStructuralDifference(a, b) === null`. The
   rules, each of which is a unit test of its own:

   - **Ignored keys.** Four keys are skipped wherever they appear: `id`
     (runtime-only, regenerated freely) and `suppressConversionReport`,
     `measurementUnit`, `showRuler` — UI state that no writer ever puts in a
     file, so a same-format round trip cannot carry it either way.
   - **Emptiness is one value.** `undefined`, `[]` and `{}` all mean "nothing
     to say" and compare equal to each other and to an absent key. The test is
     recursive: an object whose every key is ignored or itself empty is empty.
     This is what lets a reader that writes `formatting: {}` and one that omits
     the field describe the same document.
   - **XML-valued fields are compared canonically**, not as strings, via
     `canonicalizeXml`: expanded `{ns}local` names, attributes sorted by
     expanded name, namespace declarations and comments dropped, text verbatim.
     The fields are `xml` (opaque blocks and inlines), `docxSectPr`,
     `docxTblPr`, `docxTrPr`, `docxTcPr`, `odtFontFaceDecls`, and the string
     array `odtAutomaticStyles`. **A field added later that holds XML must be
     added to that list in the same PR**, or it will be compared as a raw
     string and a round trip will fail on a namespace prefix.
   - **Unparseable XML falls back to string identity** rather than throwing, so
     a malformed preserved payload fails the comparison instead of the test run.
   - **`NaN` equals `NaN`**, which plain `===` does not give.
   - Arrays compare by length first, then element-wise; objects compare over
     the union of both key sets in sorted order.

   Everything else is `===`. In particular equality does **not** merge adjacent
   runs, drop zero-length runs, drop props matching a default, or normalise
   colour case or numeric precision. Those are the editing primitives' job
   (`src/lib/textEditing.ts` keeps runs coalesced as it mutates them) or the
   readers' (both normalise colour to lowercase `"#rrggbb"` on the way in).
   Equality compares what the model actually holds — a writer that emits
   `"#FF0000"` where the reader produces `"#ff0000"` has a real bug, and hiding
   it inside the comparison would mean shipping it.

   This is the relation for **same-format round trips** (open → save → open),
   for **undo/redo** (SPEC-015), and for **serialise → deserialise**. It is
   strict about content and blind only to representation choices that carry no
   meaning.

3. **`equalSemantic(a, b)`** — `equalStructural(resolve(a), resolve(b))`, where
   `resolve` flattens each paragraph's style chain into effective props, so a
   directly-bolded run and a run bolded by its style compare equal. This is the
   relation for **cross-format conversion** (`.docx` → save `.odt` → open), and
   it must always be paired with an explicit exclusion list naming what that
   conversion is permitted to lose (see SPEC-008's conversion report).

A test that cannot pass `equalStructural` asserts `equalSemantic` plus a named
exclusion list — never a hand-rolled subset comparison.

### Implementation status
The model and `equalStructural` above describe the code as it stands. One piece
of the equality API is still owed:

| Required here | Status |
| --- | --- |
| `equalBytes` | shipped |
| `equalStructural` | shipped |
| `firstStructuralDifference` | shipped |
| `canonicalizeXml` | shipped, and what the XML-valued rule above uses |
| `resolve` (style-chain flattening) | **not implemented** |
| `equalSemantic` | **not implemented** |

Every spec asserting a cross-format round trip names `equalSemantic`, so the
first one built has to build it. Until then those specs assert
`equalStructural` between the source model masked by the writer's `dropped`
records and the re-opened model — the same comparison without the
style-resolution step, which is sound today because both readers bake a
paragraph's resolved formatting onto its runs and keep `styleId` alongside it.
`resolve` becomes necessary the moment a writer emits a style reference *in
place of* baked formatting rather than as well as it.

There is no `canonicalize` function and no need for one: the rules above are
applied during the comparison walk, which is how the first-difference path is
produced. A spec asking for "the canonicalized model" is asking for something
this codebase deliberately does not build.
