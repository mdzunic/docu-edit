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
redefining these; if this spec and another disagree, this spec wins.

## Acceptance criteria
- [ ] Clicking "New document" in the empty state creates a document and navigates to `/doc/:id`
- [ ] A new document contains exactly one empty paragraph and renders a page-shaped canvas (white page, visible margins) centred in the canvas area
- [ ] The default title is "Untitled document" and appears in the app bar
- [ ] Clicking the title in the app bar makes it editable; typing a new title and pressing Enter or blurring updates it, and Escape reverts it
- [ ] A title of only whitespace reverts to "Untitled document"
- [ ] The model is implemented in `src/model/types.ts` exactly as specified in "Canonical document model" below, including every node kind, field name and unit
- [ ] `docs/document-model.md` is committed, documents the model, the unit rules and the equality relations, and is the file later specs link to
- [ ] `src/model/equality.ts` exports `canonicalize`, `resolve`, `equalStructural` and `equalSemantic` with the semantics defined in "Equality" below
- [ ] `equalStructural` reports the path of the first difference (e.g. `sections[0].blocks[3].children[1].props.bold`) so failing round-trip tests are debuggable
- [ ] Unit tests cover: creating a document, reading/writing its title, and JSON serialise → deserialise producing an `equalStructural` document
- [ ] Each canonicalization rule listed below has its own unit test asserting that rule alone
- [ ] Unit tests assert `equalStructural` is reflexive, symmetric and transitive over hand-built models, and that a difference in any single model field is reported as unequal

## Out of scope
Typing into the document (SPEC-011). Reading or writing files (SPEC-004 to
SPEC-007). Persistence across reloads (SPEC-003). Templates (SPEC-046).
Block kinds beyond paragraph — tables (SPEC-022), images (SPEC-024) and
footnotes (SPEC-031) each extend the model, and each must update
`docs/document-model.md` and the canonicalization rules in the same PR.

## Reference

### Canonical document model
A plain-JSON tree. No class instances, no DOM nodes, no cyclic references — it
must survive `structuredClone` and `JSON.stringify`.

```ts
type Doc = {
  id: string                    // runtime only, excluded from equality
  title: string
  sections: Section[]           // at least one
  styles: StyleTable            // styleId -> StyleDef; {} until SPEC-019
  defaults: Defaults
  metadata: Metadata            // { title, author, created, modified }; SPEC-041 extends
  opaque: OpaquePart[]          // [] until SPEC-008 populates it
}

type Section   = { id: string; props: SectionProps; blocks: Block[] }
type Block     = Paragraph | OpaqueBlock       // later specs add Table, etc.
type Paragraph = {
  kind: 'paragraph'
  id: string                    // runtime only
  styleId: string | null        // null = Normal
  props: ParaProps              // direct formatting only, never resolved
  children: Inline[]            // [] means an empty paragraph
}
type Inline    = Run | LineBreak | OpaqueInline
type Run       = { kind: 'run'; props: RunProps; text: string }
type LineBreak = { kind: 'break' }

type RunProps = {               // every field optional; absent = inherit
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean
  fontFamily?: string           // as authored, e.g. "Calibri"
  fontSizeHalfPoints?: number   // integer; 22 = 11pt
  color?: Hex; highlight?: Hex
  vertAlign?: 'super' | 'sub'
}
type ParaProps = {
  align?: 'left' | 'center' | 'right' | 'justify'
  indentLeft?: Emu; indentRight?: Emu; indentFirstLine?: Emu
  spaceBefore?: Emu; spaceAfter?: Emu
  lineSpacing?: { mode: 'multiple' | 'exact' | 'atLeast'; value: number }
  direction?: 'ltr' | 'rtl'
}
type SectionProps = {
  pageWidth: Emu; pageHeight: Emu
  marginTop: Emu; marginRight: Emu; marginBottom: Emu; marginLeft: Emu
  orientation: 'portrait' | 'landscape'
}
type Defaults = { section: SectionProps; run: RunProps }
```

**Units — normative, no exceptions.**
- `Emu` is an integer in English Metric Units: 914400 per inch, 360000 per cm,
  635 per twip, 12700 per point. **Every length in the model is an `Emu`.**
  Readers convert on the way in, writers convert on the way out. No floats, no
  unit-suffixed strings, no CSS values in the model.
- Font size is an integer in half-points (Word's native unit), so 11pt is `22`.
  `.odt`'s points are converted on read and write.
- `Hex` is a lowercase `#rrggbb` string. No named colours, no `auto`.
- A blank document is A4 (`pageWidth: 7560000`, `pageHeight: 10692000`),
  2.54cm margins (`914400` each), portrait, product default body font.

**Invariants.** A `Doc` has ≥1 section; a section has ≥1 block; a paragraph's
`children` may be empty but never contains zero-length runs; adjacent runs
never have deep-equal `props`; `id` fields are regenerated freely and are never
compared.

### Equality
Three relations, referred to **by name** in later specs. No round-trip
assertion may be written without naming one of them.

1. **`equalBytes(a, b)`** — the two saved packages are byte-identical. Used
   only for preserved opaque parts (SPEC-008).

2. **`equalStructural(a, b)`** — `deepEqual(canonicalize(a), canonicalize(b))`.
   `canonicalize` applies, in this order:
   - drop all `id` fields
   - drop props whose value equals the inherited default in `Doc.defaults`
   - drop `undefined`-valued props and empty prop objects
   - merge adjacent runs with deep-equal `props`, concatenating `text`
   - drop zero-length runs
   - normalise `Hex` to lowercase and every length to an integer `Emu`
   - normalise opaque nodes with XML canonicalization (C14N)
   - sort object keys

   This is the relation for **same-format round trips** (open → save → open),
   for **undo/redo** (SPEC-015), and for **serialise → deserialise**. It is
   strict about content and deliberately blind to representation choices a
   writer is free to make.

3. **`equalSemantic(a, b)`** — `equalStructural(resolve(a), resolve(b))`, where
   `resolve` flattens each paragraph's style chain into effective props, so a
   directly-bolded run and a run bolded by its style compare equal. This is the
   relation for **cross-format conversion** (`.docx` → save `.odt` → open), and
   it must always be paired with an explicit exclusion list naming what that
   conversion is permitted to lose (see SPEC-008's conversion report).

A test that cannot pass `equalStructural` asserts `equalSemantic` plus a named
exclusion list — never a hand-rolled subset comparison.
