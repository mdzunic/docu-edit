---
id: SPEC-074
title: Paragraph and page borders and shading
status: draft
area: formatting
complexity: normal
story_points: 8
depends_on: [SPEC-014, SPEC-019, SPEC-023, SPEC-027]
---
## Why
Two specs put this out of scope on the same reasoning — not now — and
nothing has come back for it. SPEC-014's out-of-scope list names
"Paragraph borders and shading"; SPEC-027's names "Page borders and
background colour". Between them that is every border in a document except
a table's.

The absence shows up as three different visible failures. A pull-quote, a
warning box, a code block and a note callout are all *paragraphs with a
border and a background* — the most common piece of document design there
is, and one this product can neither create nor display. A horizontal rule
typed as `---` in Word becomes a bottom border on the preceding paragraph,
so a very ordinary document opens with the rule missing. And a document
with a page border — a certificate, a title page, a form — opens with the
border gone and saves without it.

The model is already most of the way there. `src/lib/core/tables.ts`
defines `BorderSide` (style, width in points, colour) and `CellBorders`
(four optional sides) for SPEC-023, and `TableCell.shadingFill` is a
CSS-ready `"#rrggbb"`. Paragraph borders are the same four sides plus a
between-paragraphs edge, and paragraph shading is the same fill. What is
missing is a home for the types outside the table module, a paragraph-level
field, a page-level field, and the layout rule for what a bordered
paragraph does at a page break.

## Acceptance criteria
- [x] `BorderSide` and `BorderLineStyle` move to `src/lib/core/borders.ts`, product-neutral, re-exported from `src/lib/core/tables.ts` so no existing import changes, and gain the additional line styles in "Border styles"
- [x] `ParagraphFormatting` gains `borders?: ParagraphBorders` and `shadingFill?: string`, per "The paragraph model"
- [x] A Borders and shading dialog (`Format → Borders and shading…`) sets each of the five paragraph edges independently — top, bottom, left, right and between — with a style, width, colour, and a preset row (none, box, all, top and bottom, custom)
- [x] Padding between the border and the text is settable per edge, in the current measurement unit (SPEC-064), and defaults to 4pt
- [x] The dialog's paragraph shading control sets a background fill, with "no fill" as a distinct value from white
- [x] Borders and shading apply to every paragraph the selection touches, using SPEC-014's "touched by the selection" rule
- [x] **Consecutive paragraphs with identical borders and shading merge into one bordered block**: the outer edges draw once around the run, and the `between` edge draws between them — the behaviour both formats specify and the reason `between` exists as a separate edge
- [x] The shading region spans the paragraph's full content width including its padding — not just behind the glyphs — and follows the paragraph's indents
- [x] A bordered paragraph split across a page break draws its side borders on both fragments, its top border only on the first and its bottom only on the last, per "Pagination"
- [x] Borders and shading are inherited from the paragraph style (SPEC-019) until the paragraph sets its own; the style editor (SPEC-042) can set them on a style
- [x] Page borders are settable per section from Page setup — style, width, colour, an art-free line only, plus which edges and whether the border measures from the page edge or the text — and render on every page of that section
- [x] A page background colour is settable per section, renders behind the whole page, and is excluded from printing unless "print background" is chosen, matching both formats' own distinction
- [x] Formatting marks (SPEC-028) and the accessibility checker (SPEC-043) treat a shaded paragraph as a contrast surface: text on a shaded paragraph is checked against the shading, not against the page
- [x] Paragraph borders and shading round-trip through `.docx` (`w:pBdr` with `w:top/bottom/left/right/between`, `@w:val`, `@w:sz` in eighths of a point, `@w:space` in points, `@w:color`; `w:shd` with `@w:fill`) and `.odt` (`fo:border*`, `fo:padding*`, `fo:background-color`, plus `style:join-border` for the merge behaviour)
- [x] Page borders and background round-trip through `.docx` (`w:pgBorders` in `w:sectPr`, `w:background` in the document) and `.odt` (master-page and page-layout border and background properties)
- [x] A border style one format lacks maps to its nearest equivalent with a `DropRecord`, per "Border styles"
- [x] Fixtures with a boxed note paragraph, a bottom-rule heading, three consecutive identical bordered paragraphs, a shaded code block and a page border open, render, save and reopen `equalStructural`

## Out of scope
Word's decorative "art" page borders (the clip-art border set) — preserved
per SPEC-008, never created. Character-level borders and shading
(`w:bdr` on a run) — highlight (SPEC-013) is the run-level fill this
product has. Gradient or picture paragraph shading; solid colour only.
Borders on images and drawings — SPEC-051's outline is theirs. Table
borders (SPEC-023 owns them, unchanged). Drop caps. Text boxes as an
alternative to bordered paragraphs. Border corner styles beyond a square
join.

## Reference

### The types this reuses
```ts
// src/lib/core/tables.ts — as shipped, moving to core/borders.ts
export type BorderLineStyle = "none" | "single" | "double" | "dashed" | "dotted";
export interface BorderSide { style: BorderLineStyle; widthPt: number; color: string; }
export interface CellBorders { top?: BorderSide; bottom?: BorderSide; left?: BorderSide; right?: BorderSide; }
```

`color` is a CSS-ready `"#rrggbb"` or the string `"auto"` for the source's
automatic colour — a convention SPEC-023 established and this spec keeps,
because paragraph borders in both formats have the same "auto" value.

The move is a re-export, not a rename: `src/lib/core/tables.ts` continues
to export all three names, so every shipped import and every SPEC-023 test
compiles untouched. `src/lib/core` is the product-neutral layer whose
boundary `boundary.spec.ts` enforces, and a border type belongs there
rather than inside the table module — that is the whole reason for the
move.

### Border styles
The shipped five (`none`, `single`, `double`, `dashed`, `dotted`) gain
`thick`, `dashDot` and `dashDotDot` — enough for what documents actually
use, without reproducing OOXML's ~180-value `ST_Border` enumeration.

Mapping: `.docx` `w:val` takes the same names (`single`, `double`,
`dashed`, `dotted`, `thick`, `dashDotStroked`, …) and `.odt` expresses
style, width and colour in one `fo:border` shorthand
(`0.5pt solid #000000`), so ODF has no `double`-vs-`thick` ambiguity but
also no `dashDot`. A style ODF cannot express writes as its nearest
equivalent — `dashDot` and `dashDotDot` become `dashed` — with a
`DropRecord` of the existing `"paragraphFormatting"` kind, so the
conversion report names it. `.odt` `fo:border` values ODF has and OOXML
does not are read to the nearest of the eight.

### The paragraph model
```ts
// src/lib/documentModel.ts
export interface ParagraphBorders {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
  /** Drawn between consecutive paragraphs that share identical borders —
   *  see "Merging". */
  between?: BorderSide;
  /** Space between each border and the text, in points. Absent means the
   *  4pt default. */
  paddingPt?: { top?: number; bottom?: number; left?: number; right?: number };
}

export interface ParagraphFormatting {
  …                                  // SPEC-014/028/032 fields, unchanged
  borders?: ParagraphBorders;
  /** Paragraph background as a CSS-ready "#rrggbb". Absent means no fill,
   *  which is distinct from "#ffffff". */
  shadingFill?: string;
}
```

Note the naming: `shadingFill` matches `TableCell.shadingFill` exactly,
because it is the same concept and a second name for it would be a trap.

### Merging
Both formats specify that consecutive paragraphs with **identical** border
and shading settings render as one bordered region. OOXML does it
implicitly (that is what `w:between` is for); ODF does it explicitly with
`style:join-border`, default `true`.

"Identical" means structurally equal `borders` and equal `shadingFill`, and
the run is broken by any paragraph that differs, by a different indent, and
by any non-paragraph block. Within a run:

- the top border draws above the first paragraph only;
- the bottom border draws below the last only;
- left and right draw down the whole run;
- the `between` border draws at each internal boundary, and where
  `between` is absent, nothing is drawn there.

The merge computation is pure and lives beside the layout code that needs
it:

```ts
// src/lib/paragraphBorders.ts
/** Groups consecutive blocks into border runs. Returns, per paragraph
 *  index, which edges it should paint. */
export function borderRuns(blocks: Block[]): Array<{ top: boolean; bottom: boolean; sides: boolean; between: boolean }>;
```

### Pagination
A bordered paragraph that splits across pages is the case that is wrong in
most implementations, so it is stated: the side borders draw on both
fragments; the top border draws only on the fragment containing the
paragraph's first line; the bottom border only on the fragment containing
its last. Shading fills each fragment's own region. This is what both
Word and Writer do, and it falls out of painting borders per *line box run*
rather than per paragraph — the same place SPEC-047's pagination already
knows where a paragraph's lines landed.

A border run split across pages behaves the same way at run level: the
top edge on the first page, the bottom on the last.

### Page borders and background
Per section, alongside `pageSize`, `margins` and `columns`:

```ts
export interface Section {
  …
  /** Page border for this section; absent means none. */
  pageBorders?: {
    top?: BorderSide; bottom?: BorderSide; left?: BorderSide; right?: BorderSide;
    /** Distance from the reference edge, in points. */
    offsetPt?: number;
    /** What offsetPt measures from. */
    measureFrom?: "page" | "text";
    /** Which pages of the section get it. */
    appliesTo?: "allPages" | "firstPage" | "exceptFirstPage";
  };
  /** Page background as "#rrggbb"; absent means none (paper white). */
  pageBackground?: string;
}
```

`pageBackground` is a *section* field in the model even though `.docx`
stores it document-wide (`w:background`) and `.odt` stores it per master
page: the writers reconcile — `.docx` writes the first section's value and
reports a `DropRecord` if sections disagree; `.odt` writes each section's
master page directly. Modelling per section and losing on write is better
than modelling per document and being unable to read a valid `.odt`.

Printing a page background is opt-in, exactly as it is in both source
applications: SPEC-034's print path ignores `pageBackground` unless the
print dialog's "print background colours" is set. Page *borders* always
print.

### Specs this changes
- `specs/014-paragraph-alignment-indents-spacing.md` — "Paragraph borders
  and shading" leaves its out-of-scope list. Its alignment, indent, spacing
  and direction rules are unchanged, and its "touched by the selection"
  rule governs how these apply.
- `specs/027-page-setup.md` — "Page borders and background colour" leaves
  its out-of-scope list; Page setup gains a Borders tab. Page size, margins
  and orientation are unchanged.
- `specs/023-table-formatting-and-cell-operations.md` — no behaviour
  change; its border types move file and are re-exported.
- `specs/019-paragraph-styles-and-style-picker.md` /
  `specs/042-custom-styles-manager.md` — borders and shading join the
  properties a paragraph style carries and the style editor edits.
- `specs/043-accessibility-checker.md` — its contrast check reads a
  paragraph's `shadingFill` as the background where one is set.
