---
id: SPEC-076
title: Table styles, header repeat, sort and text↔table conversion
status: draft
area: structure
complexity: normal
story_points: 8
depends_on: [SPEC-022, SPEC-023, SPEC-042]
---
## Why
SPEC-023's out-of-scope list is four items long and every one of them is
something people do with tables constantly: "Table styles / banded row
themes. Repeat header row on page breaks. Text wrapping around floating
tables. Sorting and formulas." Three of the four are in this spec; the
fourth (floating tables) and formulas are left out for reasons stated
below.

The header-repeat gap is the sharpest, because the model already knows the
answer and nothing acts on it. `TableRow.isHeader` exists — its doc comment
describes `w:tblHeader` and ODF's `table:table-header-rows` — and it is
read by the accessibility checker and by nothing else. `isHeader` appears
nowhere in `src/lib/pagination.ts`. So a document whose table spans four
pages shows its column headings on page one and three pages of unlabelled
numbers after it, having faithfully preserved the flag that says not to.

Table styles are the formatting gap. `NamedStyle` already carries
`tableBorders` and `tableShading` for SPEC-042's table styles, but there is
no banding — no "every other row shaded", no distinct header row, no first-
column emphasis. Banding is what makes a table readable, it is one
`w:tblLook`/`w:tblStylePr` mechanism in `.docx` and one set of ODF cell
style assignments, and its absence means every table imported from a real
document arrives visually flat.

Sorting and text↔table conversion are the two operations people reach for
when a table is *almost* right: data pasted in the wrong order, or a
tab-separated list that should have been a table all along.

## Acceptance criteria
- [ ] A row marked `isHeader` **repeats at the top of every page the table continues onto**, rendered identically to its first appearance, and the repeat is a layout artefact — the model holds one copy
- [ ] A repeated header is not editable in its repeat position: clicking it places the caret in the real header row, and it is skipped by caret navigation, find, spell-check and word count
- [ ] Multiple contiguous leading header rows repeat together; a header row that is not part of the leading run is written to `.docx` as `w:tblHeader` and reported for `.odt`, per "Header rows in each format"
- [ ] `Table → Repeat header row` toggles `isHeader` on the selected row, and the first row of a new table created from the Insert dialog can be marked as a header at creation
- [ ] A table style gains banding: `bandedRows`, `bandedColumns`, `firstRow`, `lastRow`, `firstColumn` and `lastColumn` emphasis, each independently toggleable, per "The banding model"
- [ ] A table style gallery offers a set of built-in looks (plain, grid, banded, header-emphasised, and a coloured variant), applied in one click and one undo entry, and shown as live previews of the selected table
- [ ] Banding resolves at render time from the style plus the table's toggles — it is **not** baked into per-cell shading — so toggling banded rows off restores whatever per-cell shading the cells actually carry
- [ ] A cell's own `shadingFill` and `borders` always win over the style's banding, matching both formats' precedence, and the precedence order is asserted by a unit test
- [ ] `Table → Sort…` sorts the table's rows by up to three columns, each ascending or descending, with a data type of text, number or date per column, and a "my table has a header row" toggle defaulting to the table's own `isHeader`
- [ ] Sorting moves whole rows, preserves per-cell formatting with its row, refuses tables containing merged cells that span the sort range with a stated reason, and records one undo entry
- [ ] `Table → Convert text to table…` converts the selected paragraphs, splitting on tabs, commas, semicolons or a custom character, with a preview of the resulting column count; empty trailing fields produce empty cells rather than fewer columns
- [ ] `Table → Convert table to text…` converts the selected table to paragraphs, joining cells with a chosen separator, one paragraph per row, and nested-table content flattened in document order
- [ ] Both conversions are round-trip stable for the simple case: text → table → text with the same separator returns the original paragraphs, asserted as a unit test
- [ ] Banding round-trips through `.docx` (`w:tblLook` flags with `w:tblStylePr` conditional formatting on the referenced style) and `.odt` (the table's own style plus the per-row/column cell style assignments ODF uses), with a `DropRecord` for a conditional-formatting kind the model does not carry
- [ ] Table styles and banding are preserved for tables whose style we do not fully model, per SPEC-008's rule and the existing `docxTblPr` mechanism
- [ ] Fixtures with a four-page table with a repeating header, a banded table, a table sorted by two columns and a tab-separated list open, render, save and reopen `equalStructural`

## Out of scope
Text wrapping around a floating table — the remaining SPEC-023 item, which
needs the wrap machinery drawings use and deserves its own spec. Formulas
in table cells (`=SUM(ABOVE)`) — a small expression language plus a
recalculation model, which is a spec of its own and is better placed after
SPEC-053's field infrastructure. Sorting by a custom key or by cell colour.
Sorting a selection of rows rather than the table. Nested tables as a
conversion target. Importing CSV as a table — SPEC-009's open path owns
file import. Table styles shared across documents (SPEC-042 already puts a
shared style library out of scope).

## Reference

### What ships today
```ts
// src/lib/documentModel.ts
export interface Table {
  type: "table";
  rows: TableRow[];
  columnWidthsCm: number[];
  alignment?: TableAlignment;
  widthMode?: TableWidthMode;         // "auto" | "fixed" | "percent"
  widthPercent?: number;
  /** Verbatim w:tblPr from a .docx source; re-emitted on a same-format
   *  save for whatever SPEC-023 does not model. Cleared by any command
   *  that edits the modelled fields. */
  docxTblPr?: string;
  odtName?: string;
  odtStyleName?: string;
  …                                   // custom table style ref (SPEC-042)
}

export interface TableRow {
  cells: TableCell[];
  minHeightCm?: number;
  /** w:tblHeader (.docx) or membership in a leading
   *  table:table-header-rows wrapper (.odt). */
  isHeader?: boolean;
  docxTrPr?: string;
  odtStyleName?: string;
  changes?: …;
}

export interface TableCell {
  blocks: Paragraph[];
  colSpan?: number; rowSpan?: number; covered?: boolean;
  borders?: CellBorders;
  shadingFill?: string;               // "#rrggbb"
  …
}
```

`NamedStyle` (SPEC-042) carries `tableBorders` and `tableShading` for
table-type styles. `src/lib/pagination.ts` splits tables across pages and
does not read `isHeader`.

### Header rows in each format
The two formats disagree in a way that must be handled rather than
averaged:

- **`.docx`** — `w:tblHeader` is a **per-row** flag in `w:trPr`. Any row
  can carry it, though only a leading run has a defined rendering.
- **`.odt`** — header rows are rows inside a leading
  `table:table-header-rows` wrapper. Only one contiguous leading run can
  exist.

So the model's per-row `isHeader` (already the shipped shape, and the
right one) writes losslessly to `.docx`, and to `.odt` writes the leading
contiguous run into the wrapper. A non-leading header row saved to `.odt`
reports a `DropRecord` of the existing `"tableFormatting"` kind. Reading
`.odt` marks every row in the wrapper.

Repeating is a **layout** behaviour, computed by pagination: when a table
continues onto a new page, the leading header run's laid-out rows are drawn
again at the top of the continuation, and the continuation's content starts
below them. The repeated copy is a render artefact with no model presence
and no editable DOM — it carries `aria-hidden="true"` and
`data-repeated-header="true"`, and every text-walking consumer (caret
navigation, find, spell-check, word count, the accessibility checker)
already walks the *model*, so they need no change; the guard is a test
asserting the word count of a four-page table is unchanged by the repeat.

### The banding model
```ts
// on NamedStyle, for type: "table"
export interface TableStyleBanding {
  /** Shading applied to alternate row bands. */
  bandedRows?: { size: number; shading: string };
  bandedColumns?: { size: number; shading: string };
  /** Conditional emphasis for the edge rows and columns. */
  firstRow?: TableConditionalFormat;
  lastRow?: TableConditionalFormat;
  firstColumn?: TableConditionalFormat;
  lastColumn?: TableConditionalFormat;
}

export interface TableConditionalFormat {
  shading?: string;
  bold?: boolean;
  borders?: CellBorders;
}

// on Table — which parts of the style's banding this table uses
export interface TableLook {
  firstRow?: boolean; lastRow?: boolean;
  firstColumn?: boolean; lastColumn?: boolean;
  bandedRows?: boolean; bandedColumns?: boolean;
}
```

`Table.look?: TableLook` is the direct model of `w:tblLook`, which is why
it is a separate object from the style: the style says what a banded row
looks like, the table says whether to use it. Both formats make the same
split.

**Resolution order** for a cell's effective shading and borders, lowest
first:

1. the table style's base `tableBorders` / `tableShading`;
2. row/column banding from the style, where the table's `look` enables it;
3. conditional formats for first/last row and column, where enabled;
4. the cell's own `borders` / `shadingFill`.

Resolution is pure and unit-tested, in `src/lib/tableStyles.ts`, taking a
table, its resolved style and a cell position and returning the effective
formatting. Nothing writes banding into cells — which is what makes the
"toggling banding off restores per-cell shading" AC true by construction
rather than by cleanup.

### Sort
```ts
// src/lib/tableSort.ts — pure
export interface SortKey {
  columnIndex: number;
  direction: "asc" | "desc";
  type: "text" | "number" | "date";
}
export function sortTableRows(table: Table, keys: SortKey[], hasHeader: boolean): Table;
```

Comparison: `text` uses the document language's locale collation
(SPEC-035's language tag, via `Intl.Collator` — a platform API, not a
dependency); `number` parses the cell's text leniently, treating an
unparseable cell as less than any number; `date` parses ISO and the
locale's common formats, likewise. Ties fall through to the next key, then
to the original row order — a stable sort, so re-sorting by the same key is
idempotent.

Merged cells: a table whose sort range contains a cell with `rowSpan > 1`
is refused with the reason "sorting is not available for tables with
merged rows", shown in the dialog rather than as a failed action.
`colSpan` is fine.

### Conversion
```ts
// src/lib/tableConvert.ts — pure
export function textToTable(paragraphs: Paragraph[], separator: string): Table;
export function tableToText(table: Table, separator: string): Paragraph[];
```

`textToTable` gives the table as many columns as the row with the most
fields, padding shorter rows with empty cells, and keeps each field's run
formatting. Column widths divide the content width equally.
`tableToText` emits one paragraph per row, cells joined by the separator,
taking each cell's first paragraph's formatting for the result and
appending any further paragraphs' text inline. A cell containing a nested
table contributes its own flattened text.

Both operate on the model, are unit-tested directly, and are wired to
commands that record one undo entry each.

### Specs this changes
- `specs/023-table-formatting-and-cell-operations.md` — three of its four
  out-of-scope items are resolved here. Its per-cell border and shading
  model, its cell operations and its round-tripping are unchanged, and
  cell-level formatting keeps precedence over everything a style says.
- `specs/022-tables-insert-and-edit.md` — the Insert dialog gains a
  header-row option; table creation is otherwise unchanged.
- `specs/042-custom-styles-manager.md` — "Table styles with banding"
  leaves its out-of-scope list; `NamedStyle`'s table fields gain the
  banding structure, and the style editor edits it.
- `specs/047-large-document-performance.md` — repeated header rows are
  extra laid-out rows per page break; the baseline must not move, so the
  repeat reuses the already-measured header row's layout rather than
  re-measuring it per page.
- `specs/043-accessibility-checker.md` — its "table without a header row"
  check is unchanged and gains meaning: marking a header row now has a
  visible effect, so the check's advice becomes actionable.
