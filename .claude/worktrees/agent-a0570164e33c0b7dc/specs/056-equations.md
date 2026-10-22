---
id: SPEC-056
title: Equations
status: draft
area: structure
depends_on: [SPEC-026, SPEC-052, SPEC-055]
---
## Why
A document with mathematics in it is either written in a word processor that
can set mathematics or it is not written in a word processor at all. SPEC-026
put formula editing out of scope while building the symbol catalogue this
depends on; both target formats carry maths as first-class structured content,
not as pictures; and a `.docx` full of equations currently opens as a page of
placeholder boxes.

Two constraints shape the whole spec, and both are worth stating before the
criteria rather than after. `package.json` is protected, so there is no KaTeX
and no MathJax — the layout engine is ours. That is only tractable against a
**closed construct set**, which is why one is defined here and why anything
outside it is preserved and shown read-only rather than approximated. And ODF
stores a formula as an embedded object sub-document, which is the machinery
SPEC-052 builds for charts; this queues after it to use that path rather than
building a second one.

## Acceptance criteria
- [ ] "Insert → Equation" inserts an empty inline equation at the caret with the caret inside it; "Insert → Display equation" inserts one centred on its own paragraph
- [ ] Alt+= inserts an inline equation, registered in the SPEC-021 shortcut registry like every other binding
- [ ] Typing inside an equation uses the language in "Linear input" below and renders live: `\frac{a}{b}` becomes a fraction as the closing brace is typed, not on blur
- [ ] A template palette on the equation toolbar inserts every construct in "Construct set" below with empty slots, and Tab and Shift+Tab move between a construct's slots
- [ ] Equations render per "Layout rules" below — identifiers italic, numbers and function names upright, and the operator spacing in that section
- [ ] Arrow keys walk the structure rather than the source: Right from the numerator's end enters the denominator, and Backspace at the start of a construct's first slot unwraps the construct and keeps its content
- [ ] SPEC-026's character picker inserts from its Maths and Greek categories into an equation, and `\alpha`-style names in the linear input resolve against that same catalogue with no second table
- [ ] A display equation can carry a number: a caption with the Equation label (SPEC-055) placed on the equation's own paragraph at a right tab stop, and it is a valid cross-reference target
- [ ] Every equation carries alt text, defaulting to its linear form and editable, exposed to assistive technology and picked up by SPEC-043's existing alt-text detection with no change to that spec
- [ ] An equation using a construct outside the set is preserved verbatim per SPEC-008, rendered as its linear form in a read-only box with "We can't edit this equation yet — it's kept exactly as it was", and written back unchanged
- [ ] No dependency is added to `package.json`: the layout engine, the linear parser and both serialisers are written for this spec, and a test asserts the manifest is unchanged
- [ ] Equations render with the glyphs and rules described in "Fonts and rules" below, with no math font and no external asset
- [ ] Equations print and export to PDF (SPEC-034) as selectable text, not as an image
- [ ] Equations round-trip through `.docx` (`m:oMath` inline, `m:oMathPara` for display, inside `w:p`) and `.odt` (a `draw:frame`/`draw:object` referencing an `application/vnd.oasis.opendocument.formula` sub-document whose `content.xml` holds MathML, anchored `as-char` for inline equations)
- [ ] A cross-format save reports a `DropRecord` for every construct the target cannot express, per "Format mapping" below
- [ ] Opening the fixtures below renders one equation of every construct in the set, one inline and one display equation, one numbered equation, and one unsupported construct as the read-only box

## Out of scope
Handwriting and image recognition of equations. Computing or evaluating an
equation — nothing here has a value, only a form. Chemistry notation, commutative
diagrams, and long division layouts. Equation arrays with per-row alignment
points; a matrix is the only multi-row construct. Colour and per-symbol
formatting inside an equation. Importing LaTeX documents — the linear input is
an input method, not a file format (SPEC-050 owns format import). Equation
numbering schemes beyond SPEC-055's, including chapter-based numbers. A math
font, bundled or embedded, still out of scope from SPEC-048.

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** The facts this spec rests on: `DocumentModel` holds
`sections: Section[]` whose blocks are
`Block = Paragraph | Table | Image | BreakBlock | TableOfContents`; a
`Paragraph` holds `runs: Run[]` and there is no `Inline` union; lengths are
centimetres and type sizes points; and equality drops every key named `id`, so
a reference that must survive a round-trip test is never called `id`.

### Symbol catalogue
Defined normatively in `specs/026-special-characters-and-symbols.md`, exported
from `src/lib/specialCharacters.ts` as plain data:

```ts
interface SpecialCharacter { cp: number; name: string; keywords: string[] }
interface CharacterCategory { id: CategoryId; label: string; chars: SpecialCharacter[] }
type CategoryId = 'punctuation' | 'currency' | 'maths' | 'arrows'
                | 'latin' | 'greek' | 'symbols' | 'emoji'
export const CHARACTER_CATEGORIES: CharacterCategory[]
```
Its `maths` category is 64 characters and its `greek` category is 52, both
listed by code point in that spec. **This spec ships no symbol table.** Every
`\alpha`-style name in the linear input, every `nary` operator character and
every stretchy delimiter resolves against `CHARACTER_CATEGORIES`; a symbol this
spec needs that the catalogue lacks is added there, in that spec's table, not
duplicated here.

SPEC-026 also fixes two facts this spec inherits: `Run.text` is a sequence of
UTF-16 code units with non-BMP characters stored as surrogate pairs, and
`CaretPosition.offset` is a UTF-16 code-unit index. An equation is one atomic
inline for the purposes of that offset — the caret is either before it, after
it, or inside its own structure.

### Equation numbering
Defined normatively in `specs/055-captions-and-table-of-figures.md`. A caption
is a paragraph in the Caption style holding a label run, a `Seq` field scoped
to that label, a separator and the user's text; numbering runs per label in
document order; and captions are valid cross-reference targets through
SPEC-053's `Ref` field. A numbered display equation uses that mechanism with
the built-in Equation label and adds nothing to it — no second sequence, no
equation-specific numbering path. The right tab stop that pushes the number to
the margin is `specs/054-tab-stops-and-the-tab-ruler.md`'s.

### An equation is a run, not a drawing block
`specs/051-drawing-shapes-and-text-boxes.md` defines `DrawingPlacement` for
blocks that float, and an equation is deliberately not one. An inline equation
sits *in* a line of text with a baseline the surrounding runs align to; it
reflows, breaks lines around it, and inherits the paragraph's font size. A
floating block is positioned against the page, the margin or a paragraph, and
none of that applies.

Instead it follows the convention `Run.field` (SPEC-029) and `Run.note`
(SPEC-031) already established: a marker on a run whose `text` is the
one-character `FIELD_PLACEHOLDER_TEXT`, so the equation is a single caret atom
and every offset calculation in `src/lib/textEditing.ts` is unchanged, with the
renderer substituting the laid-out equation at paint time.

```ts
// src/lib/documentModel.ts
export interface Run {
  text: string                  // FIELD_PLACEHOLDER_TEXT when `equation` is set
  formatting: RunFormatting
  field?: FieldSpec             // SPEC-029
  note?: NoteRef                // SPEC-031
  equation?: EquationSpec       // this spec — mutually exclusive with both
  styleId?: string
  changes?: RunTrackedChanges
}

export interface EquationSpec {
  display: "inline" | "block"
  body: MathNode[]
  /** Defaults to the linear form; "" means none. */
  altText?: string
}
```
A display equation is a paragraph whose only run carries an `EquationSpec` with
`display: "block"`; the paragraph centres it. That keeps display and inline
equations one mechanism rather than two.

`.docx` agrees with the run model — `m:oMath` is inline XML in the paragraph,
with no drawing wrapper. `.odt` disagrees, wrapping the formula in a
`draw:frame`; the writer emits that frame with `text:anchor-type="as-char"` and
`style:vertical-rel="baseline"`, which is ODF's way of spelling the same thing,
and the reader unwraps it. The frame is a serialisation detail of one format
and does not reach the model.

### Construct set
The closed set. `MathNode` has exactly these kinds:

```ts
type MathNode =
  | { kind: 'sym';    text: string; variant: 'italic' | 'upright' }
  | { kind: 'frac';   num: MathNode[]; den: MathNode[]; bar: boolean }
  | { kind: 'script'; base: MathNode[]; sup?: MathNode[]; sub?: MathNode[] }
  | { kind: 'radical'; radicand: MathNode[]; index?: MathNode[] }
  | { kind: 'nary';   op: string; lower?: MathNode[]; upper?: MathNode[]; body: MathNode[] }
  | { kind: 'delim';  open: string; close: string; body: MathNode[] }
  | { kind: 'matrix'; rows: MathNode[][][] }        // rows → cells → content
  | { kind: 'accent'; accent: string; base: MathNode[] }
  | { kind: 'func';   name: string; body: MathNode[] }
  | { kind: 'text';   text: string }
  | { kind: 'opaqueMath'; source: string; linear: string }
```

`opaqueMath` is the preservation case: `source` holds the original OMML or
MathML subtree verbatim and `linear` the best plain-text rendering of it, which
is what the read-only box shows. It is SPEC-008's opaque-node rule applied
inside an equation, compared with the same XML canonicalization.

`nary` covers ∑ ∏ ∫ ∬ ∮ ⋃ ⋂ — the operator character comes from SPEC-026's
Maths category, so this spec ships no symbol table of its own. `func` covers
the upright-set function names: sin cos tan sec csc cot, sinh cosh tanh, arcsin
arccos arctan, exp ln log lim max min sup inf det gcd deg.

### Linear input
The input language is a **LaTeX subset**, not Word's linear format. Word's is
undocumented, context-sensitive and ambiguous (`a/b/c` has no defined
grouping); the LaTeX subset is a published notation a large share of the people
who need this already type, and its braces make grouping explicit.

| Input | Produces |
| --- | --- |
| `x`, `12`, `+` | `sym` — see "Layout rules" for the italic/upright rule |
| `\alpha` … `\Omega` | `sym` resolved against SPEC-026's Greek category |
| `\frac{a}{b}`, `\binom{n}{k}` | `frac` with `bar: true` / `bar: false` |
| `x^2`, `x_i`, `x_i^2` | `script` |
| `\sqrt{x}`, `\sqrt[3]{x}` | `radical` |
| `\sum_{i=1}^{n}`, `\int_a^b`, `\prod`, `\bigcup` | `nary` |
| `\left( … \right)`, and `[ ] \{ \} \| \|\|` | `delim` |
| `\matrix{a & b \\ c & d}` | `matrix` |
| `\hat{x} \bar{x} \vec{x} \dot{x} \tilde{x}` | `accent` |
| `\sin`, `\lim`, `\log` … | `func` |
| `\text{if } x > 0` | `text`, spaces preserved |

**The linear form is input, never storage.** The model is the `MathNode` tree;
the file formats hold OMML and MathML. A linear form is regenerated from the
tree for editing and for `altText`, so an equation read from a `.docx` — which
carries no LaTeX at all — is editable in exactly the same way as one typed
here. Input outside the table is inserted as literal `sym` nodes rather than
raising an error, so a stray backslash never blocks typing.

### Layout rules
Normative, because "it looks right" is not a test.

**Variant.** A single-letter Latin or Greek identifier is italic. A number, an
operator, a delimiter, a function name and anything inside `text` is upright.
This is the one rule that makes output look like mathematics rather than like
prose, and it is applied at render time from `sym.variant`, which the parser
sets and the reader maps from OMML's `m:sty`/MathML's `mathvariant`.

**Spacing**, in eighteenths of an em, between adjacent nodes:

| Left → right | Space |
| --- | --- |
| ordinary → ordinary | 0 |
| ordinary → binary operator (`+ − ± × ÷`) | 4/18 |
| binary operator → ordinary | 4/18 |
| ordinary → relation (`= ≠ < > ≤ ≥ ≈ ≡`) | 5/18 |
| relation → ordinary | 5/18 |
| after a comma or semicolon | 3/18 |
| `func` name → its argument | 3/18 |
| either side of a `nary` body | 3/18 |

A binary operator with nothing to its left is a sign, not an operator, and
takes no space — the rule that makes `−1` set correctly.

**Sizes.** A script is 71% of its base's size; a script of a script is 50%,
and does not shrink further. A fraction inside a fraction's numerator or
denominator is 71%. Nothing renders below 5 pt.

**Rules.** The fraction bar and the radical bar are 1/18 em thick with 1/6 em
clearance above and below their content. A display equation centres on its
paragraph; an inline equation's baseline is its own — for a fraction, the
vertical centre of the bar sits on the surrounding text's maths axis.

### Fonts and rules
SPEC-048 bundles Carlito, Caladea and the Liberation family, and no math font;
adding one is out of scope there and adding a dependency is forbidden by
`CLAUDE.md`. So every glyph comes from SPEC-026's catalogue — Maths, Greek,
Arrows and Latin — rendered in the document's serif fallback, and every
non-glyph part is drawn as a rule.

Rendering is DOM, not canvas. `CLAUDE.md` reaches for canvas where DOM is a
poor fit, and math layout is the opposite case: it is nested boxes and
horizontal rules, which is what CSS is for, and DOM gives selection, text
search, screen-reader access, PDF text extraction and e2e assertions for free.
Fraction and radical bars are borders.

Stretchy delimiters scale by CSS transform up to 3×; beyond that they are
assembled from Unicode's own bracket pieces (U+239B–U+23AD), which is why no
math font is needed for a tall matrix.

### Format mapping
| Construct | `.docx` OMML | `.odt` MathML |
| --- | --- | --- |
| `sym` | `m:r`/`m:t` | `mi` / `mn` / `mo` |
| `frac` | `m:f` (`m:type val="noBar"` when `bar: false`) | `mfrac` (`linethickness="0"`) |
| `script` | `m:sSup` / `m:sSub` / `m:sSubSup` | `msup` / `msub` / `msubsup` |
| `radical` | `m:rad` | `msqrt` / `mroot` |
| `nary` | `m:nary` | `munderover` over an `mo` |
| `delim` | `m:d` | `mrow` with stretchy `mo` |
| `matrix` | `m:m` | `mtable` |
| `accent` | `m:acc` | `mover accent="true"` |
| `func` | `m:func` | `mi mathvariant="normal"` |
| `text` | `m:r` with `m:nor` | `mtext` |

The asymmetries that produce `DropRecord`s (`src/lib/opaque.ts`), reusing the
existing `kind: "opaqueInline"` with a `reason` naming the construct: an `opaqueMath`
node read from one format has no expression in the other and drops on a
cross-format save; OMML's alignment and break hints inside `m:oMathPara` have no
MathML equivalent; and MathML `mtable` attributes beyond the plain grid are not
modelled. Same-format saves drop nothing.

On `.odt`, the reader takes the MathML in the formula sub-document and ignores
any `annotation` element; the writer emits MathML and no StarMath annotation.
LibreOffice regenerates its own annotation from the MathML on open, so writing
one would mean implementing a second notation to produce something the reading
application throws away.

### Existing surfaces this reuses
```ts
// src/lib/docx/reader.ts   parseDocx(data): Promise<ParsedDocxDocument>   — SPEC-004
// src/lib/odf/reader.ts    parseOdt(data):  Promise<ParsedOdtDocument>    — SPEC-005
// src/lib/docx/writer.ts   writeDocx(…): Promise<WriteResult>            — SPEC-006/008
// src/lib/odf/writer.ts    the .odt writer, same WriteResult             — SPEC-007/008
// src/lib/opaque.ts        WriteResult = { bytes: Uint8Array; dropped: DropRecord[] }
```
The `.odt` embedded-object path — writing a sub-document with its own
`mimetype` and manifest entry, and resolving `draw:object/@xlink:href` on read —
is built by `specs/052-charts.md` for chart objects and reused here unchanged.
Per `CLAUDE.md`, the linear parser, the layout engine and both serialisers live
in `src/lib/` as plain TypeScript with no Vue imports, unit-testable without
mounting a component; only the renderer touches components.

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic`, **which is not implemented**; this spec does not need it.
- **`equalStructural`** — for the same-format round trip of the fixtures below.
  `opaqueMath.source` is compared with `canonicalizeXml`, the same C14N pass
  SPEC-008 applies to preserved XML, so attribute order and insignificant
  whitespace in a preserved subtree do not fail a test.
- For cross-format saves, assert `equalStructural` between the source model
  masked by the writer's `dropped` records and the re-opened model — never a
  hand-rolled subset comparison.

`equalBytes` is not used by this spec. A round trip is asserted on the model,
not on the rendered layout: layout is checked separately by unit tests over the
spacing and size rules above, each rule tested on its own as SPEC-002 requires
of canonicalization rules.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). Adds
`fixtures/docx/word-equations.docx` and `fixtures/odt/lo-equations.odt`, each
containing one equation per construct, the quadratic formula as an integration
case (fraction, radical, scripts, delimiters and signs in one expression), one
inline equation inside a sentence, one numbered display equation, a 3×3 matrix
with stretched delimiters, and one construct outside the set. Manifest entries
name each equation's constructs and its expected linear form.
