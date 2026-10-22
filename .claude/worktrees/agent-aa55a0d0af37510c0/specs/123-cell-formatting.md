---
id: SPEC-123
title: Cell formatting — fonts, fills, borders, alignment and number formats
status: draft
area: formatting
complexity: complex
story_points: 13
depends_on: [SPEC-061, SPEC-077, SPEC-118, SPEC-122]
---
## Why
After SPEC-122 a workbook can be typed into, and SPEC-117 already draws every
font, fill, border and alignment a file brings. Nothing lets the user change
one. A budget whose totals cannot be bolded, a price list whose prices
cannot be shown as currency, a header row that cannot be given a fill: the
grid is editable and still reads as a viewer.

The model is ready for it. `CellStyle` (SPEC-117) holds every formatting
fact both formats carry; `internStyle` and the append-only style writers
(SPEC-122) turn a new style into an index without disturbing a single byte
of the source's existing style tables; SPEC-118 renders any format code in
six locales. What is missing is the command layer that changes styles over
**ranges** — including whole columns a million rows tall, which must cost
what Excel's own file costs (one `<col style>`), not a million cells — and
the surfaces that drive it: a formatting strip, the Format Cells dialog,
the format painter, Clear, and the cell styles gallery.

**This file is the normative source for formatting commands over ranges
(`StylePatch`, `patchStyle`, `applyStylePatch` in
`spreadsheet/formatting.ts`), whole-row/column/sheet formatting without
materialising empty cells, the sheet formatting toolbar, the Format Cells
dialog and its number-code builders, the number-format presets and decimal
stepping, the format painter, Clear, and named cell styles (their model,
the built-in definitions and both wire mappings).** Merge buttons sit in
this toolbar's alignment band but belong to SPEC-124.

## Acceptance criteria
- [ ] `src/lib/spreadsheet/formatting.ts` exports `StylePatch`, `patchStyle(style, patch): CellStyle` and `applyStylePatch(model, sheetId, ranges, patch): FormatResult` per "Patches"; `patchStyle` is pure, a `null` removes a field or a whole aspect, empty sub-objects are dropped, and a result that differs from its input never carries `sourceXf`/`sourceStyleName`; a patch that changes nothing returns the input object itself
- [ ] `applyStylePatch` honours the model invariants: copy-on-write (a new `rows` array, new `Row`/`Cell` objects only for what changed, every untouched row keeping its identity), sparseness (an empty position is stored only when its patched style differs from what the position resolves to without a cell), interning memoised per base style id within one command (a 200,000-cell range sharing one base style calls `internStyle` once), and `sheet.dirty = true` on the edited sheet only
- [ ] Whole rows, whole columns and the whole sheet are formatted per "Range classes" without materialising empty cells: a full column selection splits/creates `ColumnSpan`s carrying the patched `styleId`, a full row selection sets `Row.styleId`, the whole sheet patches every span, every styled row and every stored cell; cells are created only at the intersections the table names — formatting column C of a 1,048,576-row sheet with 40 stored rows creates at most 40 cells
- [ ] A command that would materialise more than `MAX_FORMAT_MATERIALISED = 1_000_000` cells and rows is refused before touching the model with `FORMAT_TOO_LARGE_MESSAGE` ("That's too many cells to format one by one — select whole columns or rows instead"); a range on a protected sheet that contains locked cells is refused with SPEC-122's protected-cell message unless the sheet's protection permits formatting (`sheetProtection@formatCells="0"`, `formatColumns="0"`, `formatRows="0"` for the matching class)
- [ ] A font patch on a cell with rich-text `runs` is applied to every run's `font` as well, so the change is visible; Clear ▸ Formats flattens `runs` to plain text (the value is unchanged); formatting part of a cell's text is backlog
- [ ] The formatting strip (`SheetFormatToolbar.vue`, `role="toolbar"`, `aria-label="Sheet toolbar"`, `sheet-format-toolbar`) lays out the eight bands of "The toolbar" in fixed order with SPEC-061's chrome — `AppIcon` icons, tooltips from `shortcutLabel`, one roving tab stop, 1px band rules — and collapses bands 8→5 through the existing `useToolbarOverflow`/`planCollapse` unchanged; at a 1440px viewport no band collapses
- [ ] Toggle and value controls reflect the **active cell's** effective style (resolution cell → row → column → 0 per SPEC-117), never a scan of the selection; a toggle sets the property on the whole selection when the active cell lacks it and removes it when the active cell has it (Bold on a selection whose active cell is bold un-bolds every cell) — a stated deviation from SPEC-061's mixed state, because a spreadsheet selection can be a million cells and both Excel and Sheets behave this way
- [ ] Font controls reuse `FontFamilySelect` (`font-family-select`; `documentFonts` = the families used by `model.styles`), `FontSizeControl` (`font-size-*`, 1–409 pt) and `TextColorPicker` (`text-color-*`) with `:theme="model.theme ?? DEFAULT_THEME"`, so a theme swatch stores a `ThemeColorRef`; B/I/U/S are `sheet-bold`, `sheet-italic`, `sheet-underline` (single), `sheet-strikethrough` with `aria-pressed`
- [ ] Fill colour (`SheetFillPicker.vue`: `sheet-fill-button` applies the last-used fill, `sheet-fill-menu-button` opens `sheet-fill-menu` with a "No fill" item `sheet-fill-none`, the theme rows through `ThemeSwatchRow` with testid prefix `sheet-fill-theme`, the `TEXT_COLOR_PALETTE` swatches `sheet-fill-swatch` and SPEC-013's recent colours) applies `{ kind: "solid", color }` or removes the fill; a theme swatch stores the `ThemeColorRef` (`lumMod`/`lumOff` from `THEME_SWATCH_VARIANTS`), which SPEC-117's tint inversion writes as `theme`+`tint`
- [ ] Borders (`sheet-borders-button` re-applies the last preset; `sheet-borders-menu-button` → `sheet-borders-menu`) offer the presets Bottom, Top, Left, Right, All borders, Outside borders, Thick outside borders and No border (`sheet-border-bottom|top|left|right|all|outside|thick-outside|none`) plus "Line colour" (`sheet-border-line-color`) and "Line style" (`sheet-border-line-style`, the 13 `SheetBorderStyle`s) that persist for the view session; each preset is computed by `borderPresetPatches(range, preset, edge)` per "Borders" and applied through `applyStylePatch`, including the neighbour-edge clearing rule
- [ ] Alignment controls: `sheet-align-left|center|right` (clicking the active state returns to `general`), `sheet-valign-top|middle|bottom`, `sheet-wrap`, `sheet-indent-decrease|increase` (indent 0–250; increasing on a `general`/`center` cell also sets `horizontal: "left"`), and `sheet-rotation` (`sheet-rotation-menu`: Angle counterclockwise 45, Angle clockwise 135, Vertical text 255, Rotate text up 90, Rotate text down 180, "Format cell alignment…" — choosing the current value resets to 0); band 6 reserves the slot SPEC-124's `sheet-merge` fills
- [ ] The number-format dropdown (`sheet-number-format`, showing the active cell's category name; `sheet-number-format-menu`) lists General, Number, Currency, Accounting, Short date, Long date, Time, Percentage, Fraction, Scientific, Text and "More formats…" (`sheet-number-format-<id>` per "Number-format presets"), each item previewing the active cell's value in that format through SPEC-118's `formatValue`; "More formats…" opens Format Cells on its Number tab
- [ ] `sheet-currency` (Accounting), `sheet-percent` (`0%`) and `sheet-comma` (Accounting without a symbol) apply the locale's codes from "Number-format presets"; `sheet-decimals-increase`/`sheet-decimals-decrease` apply `stepDecimals(code, value, ±1)` computed from the **active cell** to the whole selection, with every example in "Decimal stepping" a unit test
- [ ] Format Cells (`FormatCellsDialog.vue` on `BaseDialog`, `format-cells-dialog`; Ctrl/Cmd+1, the Format menu and the cell context menu) has the tabs Number, Alignment, Font, Border, Fill and Protection (`format-cells-tab-number|alignment|font|border|fill|protection`) with the controls of "Format Cells", opens on the active cell's effective style, renders a live preview `format-cells-preview` of the active cell's value with the pending style, and on OK applies **only the fields the user changed** as one `StylePatch` (untouched fields keep each cell's own value — a dialog opened on a mixed selection never flattens it)
- [ ] The Number tab's code builders `buildNumberCode(spec, locale)` / `describeNumberCode(code, locale)` round-trip every category's option set exactly (a code `describeNumberCode` cannot express opens on Custom); the Custom category edits the code in `format-cells-custom-code`, lists the built-in codes plus every custom code in the workbook, previews through `formatValue` in `format-cells-sample` as the user types, and a code SPEC-118's `parseFormatCode` rejects shows "This number format isn't valid" (`format-cells-custom-error`) and disables OK
- [ ] The format painter (`sheet-format-painter`): a click arms it for one application, a double-click locks it (`aria-pressed="true"`, `data-locked="true"`) until Escape or another click on it; while armed the grid carries `data-format-painter="once|locked"` and the copy cursor; the next completed selection (pointer-up, or Enter after a keyboard selection) receives the source's style pattern per "Format painter", one undo step per application, across sheets if the user switches tabs
- [ ] Clear (`sheet-clear` → `sheet-clear-menu`: "Clear all" `sheet-clear-all`, "Clear formats" `sheet-clear-formats`, "Clear contents" `sheet-clear-contents`) runs `clearRange(model, sheetId, ranges, "all" | "formats" | "contents")` per "Clear"; Clear contents is the same operation as SPEC-122's Delete key
- [ ] Named cell styles: `SpreadsheetModel.namedStyles?: NamedCellStyle[]` and `BUILTIN_CELL_STYLES` live in `spreadsheet/styles.ts` per "Named cell styles"; the gallery (`sheet-cell-styles` → `CellStylesGallery.vue`, `sheet-cell-styles-gallery`, one `sheet-cell-style` per style with `data-style="<name>"`, each drawn in its own style) groups the 18 built-ins as Excel does plus a "Custom" group for workbook-defined names, and `applyNamedStyle(model, sheetId, ranges, name)` replaces exactly the aspects the style includes and sets `namedStyle`
- [ ] A built-in the workbook already defines (xlsx by `builtinId` or name, ods by display name) renders and applies with the **workbook's** definition; one it lacks uses the pinned rendering and is appended to `namedStyles` on first use, then written per the wire mapping — xlsx a `cellStyleXfs` `xf` + `cellStyle` (with `builtinId`), ods a common `table-cell` style in `styles.xml`
- [ ] Both readers populate `namedStyles` and each `CellStyle.namedStyle` per the wire mapping (xlsx `cellXfs/xf@xfId` → `cellStyles`, `applyX="0"` exclusions; ods `style:parent-style-name` → the common style's display name, "Default" = Normal); this spec bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`, whose dirty-sheet step fills the workbook-level fields from the retained styles part idempotently
- [ ] Shortcuts per "Shortcuts", declared in `SHEET_SHORTCUTS` (`src/views/spreadsheet/sheetShortcuts.ts`) and matched through `src/lib/shortcuts.ts`'s `findShortcut`/`shortcutLabel`, which gain an optional `definitions` parameter (default `SHORTCUTS`) and two binding fields — `code` (match `KeyboardEvent.code`) and `ctrl` (the Control key on every platform) — with `shortcuts.spec.ts` asserting no `SHEET_SHORTCUTS` binding claims `RESERVED_BROWSER_KEYS` and the text registry's behaviour is unchanged
- [ ] Every command is one undo step through SPEC-121's `record(label, scope)` (labels "Bold", "Fill colour", "Borders", "Number format", "Format cells", "Cell style", "Format painter", "Clear formats", "Clear contents", "Clear all") and persists write-through; undo restores the previous `rows`, column spans and `styles` length exactly (a named style appended to `namedStyles` stays — an unused definition is harmless, and Excel keeps them too)
- [ ] Round trip: formatting `spec123-styles.xlsx`/`.ods` per the e2e flow, downloading and reopening reproduces every changed `CellStyle` field and `namedStyle` (field-level `toEqual` on the re-read resolved styles, ignoring `sourceXf`/`sourceStyleName`), every untouched part is `equalBytes` with the source, and every source `numFmts`/`fonts`/`fills`/`borders`/`cellStyleXfs`/`cellXfs`/`cellStyles` entry keeps its index and bytes (append-only)
- [ ] `SpreadsheetView.vue` gains only `useFormatting({ … })` and `useFormatPainter({ … })` wiring calls plus template hookup; logic lives in `src/views/spreadsheet/useFormatting.ts`, `useFormatPainter.ts` and `useFormatCellsDialog.ts` (each ≤ 400 lines per `composableSize.spec.ts`), and `SpreadsheetView.size.spec.ts`'s budget is not raised (extract into a component if the hookup does not fit)
- [ ] Fixtures: `spreadsheet/fixtures/build.ts` builds `e2e/fixtures/spec123-styles.xlsx` (Excel-shaped: `cellStyleXfs` with Normal, Good, Heading 1 and a custom "Invoice total", `cellStyles` with `builtinId`s and one `hidden="1"`, `cellXfs` with `xfId` and `applyX` attributes, theme colours with tints, custom `numFmt` 164+, a `<col style>`, a `<row s customFormat="1">`, a `dxfs` block and an x14 `extLst` that must survive untouched) and `spec123-styles.ods` (LibreOffice-shaped: `office:styles` with `Default`, `Good`, `Heading_20_1`; automatic `ce` styles with `style:parent-style-name`; `N1xx` number styles with `style:map`; a column `table:default-cell-style-name`)
- [ ] Tests: `spreadsheet/formatting.spec.ts` (patch semantics, every range class and its materialisation bound, the cap, protection, runs, borders and neighbour clearing, presets per locale, `stepDecimals`, `buildNumberCode`/`describeNumberCode`, `clearRange`, `applyNamedStyle`, painter patterns), `styles.spec.ts` (built-in table), `xlsx/styles.spec.ts` and `ods/styles.spec.ts` (named-style read/write), `formattingRoundTrip.spec.ts`, `open.spec.ts` (the upgrade step), `shortcuts.spec.ts`, `useFormatting.spec.ts`, `useFormatPainter.spec.ts`, `FormatCellsDialog.spec.ts`, `SheetFormatToolbar.spec.ts`
- [ ] e2e (`e2e/SPEC-123.spec.ts`): open `spec123-styles.xlsx` → select B2:C3 → Bold, fill Accent 1 lighter 40%, All borders → Currency from the dropdown (B2 reads "$1,234.50") → select D2 (0.125) → Ctrl/Cmd+1 → Custom `0.0%` previews "12.5%" → OK → apply "Good" from the gallery → format painter from B2 onto E2:E4 → select column F → Bold (no `sheet-cell` beyond the used rows gains a stored cell) → Clear formats on C3 → undo twice → download → reopen → the same formatting and named styles; then the same core flow on `spec123-styles.ods`

## Out of scope
Merge, unmerge and Merge & Centre (SPEC-124 — its button sits in band 6).
Formatting part of a cell's text (rich-text runs as an editing surface) —
backlog. Creating, modifying, duplicating or deleting named styles (the
gallery applies them; the Style dialog is backlog), "Merge styles" from
another workbook, and the Comma/Currency/Percent named styles as gallery
entries (read and preserved, not offered). Gradient fills as something the
user authors (they render and survive via SPEC-117; the Fill tab sets solid
and pattern fills). Choosing a theme font
reference from the font list. Conditional formatting (SPEC-129), table
styles (SPEC-133), sheet and workbook protection editing (whole-queue
backlog; SPEC-122 only honours it), deleting custom number formats from the
workbook, and "Shrink to fit"'s rendering beyond SPEC-117's.

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — `SpreadsheetModel { sheets, styles
  (117), theme? (117), opaque, … }`, `Sheet { id, name, rows: Row[] }`,
  `Row { index; cells: Cell[] }`, `Cell { col; value; formula?; styleId?;
  runs?: TextRun[] }`, `MAX_ROWS = 1_048_576`, `MAX_COLS = 16_384`;
  invariants **copy-on-write rows**, **sparse**, **the source is the splice
  base**, plain JSON. `address.ts`: `CellRef { row; col }`, `GridRange { top;
  left; bottom; right }` (inclusive, normalised), `findRow`, `getCell`,
  `usedRange`. `open.ts`: `SPREADSHEET_READER_VERSION`, `ReaderUpgrade { from;
  upgradeDirtySheet(model, sheet, sourceParts) }`, `READER_UPGRADES`.
- `specs/117-sheet-grid-styles-sizes-and-selection.md` — `CellStyle { font?;
  fill?; borders?: { top?; right?; bottom?; left?; diagonal? }; alignment?;
  numberFormat?; protection?: { locked?: false; hidden?: true }; namedStyle?;
  sourceXf?; sourceStyleName? }`, `CellFont { family?: ModelFont; sizePt?;
  bold?: true; italic?: true; underline?: "single" | "double" |
  "singleAccounting" | "doubleAccounting"; strike?: true; color?: ModelColor;
  vertAlign? }`, `CellFill` (`solid` | `pattern` | `gradient`),
  `CellBorderEdge { style: SheetBorderStyle; color? }`, `CellAlignment {
  horizontal?; vertical?; wrap?: true; indent?; textRotation? /* 0..180, 255 */;
  shrinkToFit?: true }`. `model.styles[0]` is "Normal"; resolution cell → row
  (`Row.styleId`, present only for `customFormat` rows) → column
  (`ColumnSpan.styleId`) → 0, through `effectiveStyleId(sheet, row, col)` and
  `cellStyle(model, id)` in `styles.ts`. `Sheet.columns: ColumnSpan[] { min; max;
  widthPx?; hidden?; styleId?; sourceWidthChars? }` sorted, non-overlapping.
  xlsx `theme` index swap and tint ↔ `lumMod`/`lumOff` mapping (t>0 →
  `lumMod 1−t, lumOff t`; t<0 → `lumMod 1+t`). `SheetSelection { active;
  ranges; activeRangeIndex }`. Primary modifier via `core/platform.ts`.
- `specs/118-number-formats-and-dates.md` — `formatValue(value, code, ctx:
  FormatContext): FormattedValue`, `parseFormatCode(code)` (cached; rejects
  invalid codes), `BUILTIN_NUMBER_FORMATS` (ids 0–49), `SpreadsheetLocale`
  (`en-US` default, `en-GB`, `de-DE`, `fr-FR`, `es-ES`, `cs-CZ`, a per-user
  preference). Format codes are stored in the invariant syntax (`.` decimal,
  `,` group); the locale changes only rendering.
- `specs/121-create-workbook-and-sheet-management.md` —
  `useSpreadsheetUndo.ts`: `record(label, scope)`, by-reference snapshots of
  `rows`, sheet metadata (`columns`, merges…) and `styles` length;
  `Sheet.dirty?: true`; write-through persistence.
- `specs/122-cell-entry-and-the-sheet-writer.md` — `internStyle(model, style):
  number` (structural dedup, append-only), the xlsx style writer (append-only
  `cellXfs`/`fonts`/`fills`/`borders`/`numFmts`; built-in id reuse for a code
  equal to `BUILTIN_NUMBER_FORMATS[id]`), the ods writer's generated `de-ce1…`
  automatic styles, protection honour and its message, the Delete key.
- `specs/061-…` — bands, overflow (`src/lib/toolbarOverflow.ts`
  `COLLAPSE_ORDER = [8, 7, 6, 5]`, `planCollapse`; `src/composables/
  useToolbarOverflow.ts`), roving tab stop, tooltips via `shortcutLabel`,
  `aria-pressed`. `specs/077-…`/SPEC-094 — `core/theme.ts`: `ModelColor`,
  `ThemeColorRef { theme; lumMod?; lumOff?; … }`, `DEFAULT_THEME`,
  `resolveColor`, `THEME_SWATCH_SLOTS`, `THEME_SWATCH_VARIANTS`;
  `core/colors.ts`: `TEXT_COLOR_PALETTE`, recents; components
  `TextColorPicker.vue` (`theme` prop), `ThemeSwatchRow.vue` (`testid`
  prefix), `FontFamilySelect.vue`, `FontSizeControl.vue`, `BaseDialog.vue`,
  `ui/UiMenuSurface.vue`, `AppIcon.vue` + `icons/paths.ts`.
- `specs/020-…` — the painter vocabulary (click once, double-click locks,
  Escape disarms); `specs/021-…` — `src/lib/shortcuts.ts` (`ShortcutBinding
  { key; mod?; shift?; alt?; shiftOptional? }`, `findShortcut`,
  `shortcutLabel`, `formatBinding`, `RESERVED_BROWSER_KEYS`).

### Patches
```ts
// src/lib/spreadsheet/formatting.ts
type Nullable<T> = { [K in keyof T]?: T[K] | null };
export interface StylePatch {
  font?: Nullable<CellFont> | null;          // object: merge key by key; null: remove the aspect
  fill?: CellFill | null;
  borders?: { top?: CellBorderEdge | null; right?: CellBorderEdge | null; bottom?: CellBorderEdge | null;
              left?: CellBorderEdge | null; diagonal?: (CellBorderEdge & { up?: true; down?: true }) | null } | null;
  alignment?: Nullable<CellAlignment> | null;
  numberFormat?: string | null;              // null = General
  protection?: Nullable<{ locked: false; hidden: true }> | null;
  namedStyle?: string | null;
}
export type FormatResult = { ok: true } | { ok: false; message: string };
export function patchStyle(style: CellStyle, patch: StylePatch): CellStyle;
// a position's current style is SPEC-117's effectiveStyleId(sheet, row, col) — no second resolver here
export function applyStylePatch(model: SpreadsheetModel, sheetId: string, ranges: GridRange[], patch: StylePatch): FormatResult;
export function clearRange(model, sheetId, ranges, what: "all" | "formats" | "contents"): FormatResult;
export function applyNamedStyle(model, sheetId, ranges, name: string): FormatResult;
export const MAX_FORMAT_MATERIALISED = 1_000_000;
export const FORMAT_TOO_LARGE_MESSAGE = "That's too many cells to format one by one — select whole columns or rows instead";
```
A stored cell this spec writes carries its `styleId` explicitly whenever it
is non-zero, and carries `styleId: 0` when it is 0 but its row/column
fallback is not — so the cell means the same thing under any reading of an
absent id. Overlapping ranges are harmless: patches are idempotent.

### Range classes
| Class | Test | What changes | Cells materialised |
| --- | --- | --- | --- |
| Ordinary | not full rows/columns | every position in the range | an empty position whose patched style ≠ its fallback (≤ area) |
| Full rows | `left = 0`, `right = MAX_COLS−1` | each row's `styleId` = patch(row's style ?? 0); every stored cell in it | when the row had **no** `styleId` before: one cell per column of every styled `ColumnSpan` (patch of the column's style), because the winner for those positions changes from column to row |
| Full columns | `top = 0`, `bottom = MAX_ROWS−1` | spans split at the range edges, gaps created; each span's `styleId` = patch(span's style ?? 0); every stored cell in them | one cell per (styled row × target column) — the row still wins there, so the patch must land on a cell |
| Whole sheet | both | every span (gaps filled to cover `0..MAX_COLS−1`), every row with a `styleId`, every stored cell | none — no position's winner changes |
New `Row` objects and materialised cells both count toward
`MAX_FORMAT_MATERIALISED`; the count is computed before any change. Merged
areas are formatted position by position like any range (covered cells
keep their own styles, which is how Excel draws borders around a merge).

### Borders
`borderPresetPatches(range, preset, edge): { range: GridRange; patch: StylePatch }[]`
(at most nine entries — interior, four edges, four corners — plus the
neighbour entries), `edge` = the session's line style (default `thin`) and
colour (default automatic = no `color`):

| Preset | Edges set inside the range |
| --- | --- |
| Bottom / Top / Left / Right | that outer edge of the boundary cells |
| All borders | every edge of every cell |
| Outside borders | the four outer edges |
| Thick outside borders | the four outer edges with style `medium` (what Excel's "Thick Box Border" writes) |
| No border | every edge of every cell removed, diagonals kept |

When a preset sets or removes an outer edge, the facing edge of the stored
neighbour just outside the range (e.g. `top` of the cell below for
Bottom) is removed, so the shared line shows exactly what was chosen;
neighbours are never materialised for this. Full-row/column ranges use the
same patches through the range classes (Outside on column C: `left`/
`right` on the span, `top` on C1, `bottom` on C1048576).

### Number-format presets
Codes are invariant syntax; ids are `BUILTIN_NUMBER_FORMATS` entries.
`LOCALE_CURRENCY` (in `formatting.ts`; SPEC-118's `SpreadsheetLocale` is not
widened): en-US `$` before, negatives in parentheses; en-GB `£`
before; de-DE, fr-FR, es-ES `€` after; cs-CZ `Kč` after (minus negatives
outside en-US).

| Entry (`sheet-number-format-<id>`) | en-US | other locales |
| --- | --- | --- |
| `general` General | absent | absent |
| `number` Number | id 2 `0.00` | same |
| `currency` Currency | `"$"#,##0.00` | `"£"#,##0.00` / `#,##0.00 "€"` / `#,##0.00 "Kč"` |
| `accounting` Accounting (also `sheet-currency`) | id 44 | before: `_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-`; after: `_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-` |
| `short-date` Short date | id 14 (rendered as the locale's short date) | same |
| `long-date` Long date | `dddd, mmmm d, yyyy` | en-GB `dddd, d mmmm yyyy`; de-DE `dddd, d. mmmm yyyy`; fr-FR `dddd d mmmm yyyy`; es-ES `dddd, d "de" mmmm "de" yyyy`; cs-CZ `dddd d. mmmm yyyy` |
| `time` Time | id 19 `h:mm:ss AM/PM` | id 21 `h:mm:ss` |
| `percentage` Percentage | id 10 `0.00%` | same |
| `fraction` Fraction | id 12 `# ?/?` | same |
| `scientific` Scientific | id 11 `0.00E+00` | same |
| `text` Text | id 49 `@` | same |

`sheet-percent` = id 9 `0%`; `sheet-comma` = id 43 (en-US) or the
locale's accounting template without the symbol and its space.

**Decimal stepping** — `stepDecimals(code, value, dir)`: on General, `d` =
the decimals in `generalText(value)`; +1 → `0.` + `d+1` zeros, −1 → `d−1`
zeros (`0` at none). Otherwise every number section (tokenised by SPEC-118's
grammar; text `@` and date/time sections untouched) adds or removes one `0`
after its last decimal placeholder, inserting/removing the point:
`0`→`0.0`, `0.0`→(−1)`0`, `#,##0.00`→`#,##0.000`, `0%`→`0.0%`,
`0.00E+00`→`0.000E+00`, `"$"#,##0.00_);("$"#,##0.00)`→both sections,
`0`→(−1)`0` (no-op), `yyyy-mm-dd`→no-op.

**Code builders** — `buildNumberCode({ category, decimals, thousands,
negative, symbol, type }, locale)`. Number: `0`/`#,##0` + `.`+zeros;
negatives `minus` (`c`), `red` (`c;[Red]c`), `parens` (`c_);(c)`),
`redParens` (`c_);[Red](c)`). Currency: Number with the symbol quoted
before (`"$"c`) or after (`c "€"`) per symbol position; symbols None, the
locale's, `$`, `€`, `£`, `¥`, `Kč`, `CHF`. Accounting: the template above
with `.00`/`??` resized to `decimals`. Percentage `0`+decimals+`%`;
Scientific `0.`+zeros+`E+00`; Fraction `# ?/?`, `# ??/??`, `# ???/???`,
`# ?/2`, `# ?/4`, `# ?/8`, `# ??/16`, `# ?/10`, `# ??/100`; Date ids 14, 15,
16, 17, 22 + the long date + `yyyy-mm-dd` + `mmmm yyyy`; Time ids 18–21,
45, 46, 47; Text `@`.

### Named cell styles
```ts
// src/lib/spreadsheet/styles.ts (additions)
export type StyleAspect = "numberFormat" | "alignment" | "font" | "borders" | "fill" | "protection";
export interface NamedCellStyle {
  name: string;               // display name ("Good", "Heading 1", "Invoice total")
  builtinId?: number;         // xlsx cellStyle@builtinId
  style: CellStyle;           // the definition: only included aspects; never namedStyle/sourceXf
  includes: StyleAspect[];
  hidden?: true;              // xlsx cellStyle@hidden — kept, not shown in the gallery
  sourceXfId?: number;        // xlsx cellStyleXfs index read from
  sourceStyleName?: string;   // ods common style:name ("Heading_20_1")
}
export const BUILTIN_CELL_STYLES: readonly { name: string; builtinId: number; includes: StyleAspect[]; style: CellStyle }[];
// SpreadsheetModel gains: namedStyles?: NamedCellStyle[]  (file order)
```
Pinned renderings (this product's definitions for a workbook that lacks the
style; Office-theme colours as `ThemeColorRef`s so they follow the theme):

| Name | builtinId | Includes | Definition |
| --- | --- | --- | --- |
| Normal | 0 | all | `model.styles[0]` |
| Good / Bad / Neutral | 26 / 27 / 28 | font, fill | font `#006100`/`#9C0006`/`#9C5700`; fill `#C6EFCE`/`#FFC7CE`/`#FFEB9C` |
| Title | 15 | font | `{ theme: "major" }` 18 pt, colour dk2 |
| Heading 1 | 16 | font, borders | 15 pt bold dk2; bottom `thick` accent1 |
| Heading 2 | 17 | font, borders | 13 pt bold dk2; bottom `thick` accent1 lumMod .5 lumOff .5 |
| Heading 3 | 18 | font, borders | bold dk2; bottom `medium` accent1 lumMod .6 lumOff .4 |
| Heading 4 | 19 | font | bold dk2 |
| Total | 25 | font, borders | bold dk1; top `thin` accent1, bottom `double` accent1 |
| Input | 20 | font, fill, borders | `#3F3F76`; `#FFCC99`; all `thin` `#7F7F7F` |
| Output | 21 | font, fill, borders | bold `#3F3F3F`; `#F2F2F2`; all `thin` `#3F3F3F` |
| Calculation | 22 | font, fill, borders | bold `#FA7D00`; `#F2F2F2`; all `thin` `#7F7F7F` |
| Check Cell | 23 | font, fill, borders | bold lt1; `#A5A5A5`; all `double` `#3F3F3F` |
| Note | 10 | fill, borders | `#FFFFCC`; all `thin` `#B2B2B2` |
| Warning Text | 11 | font | `#FF0000` |
| Explanatory Text | 53 | font | italic `#7F7F7F` |
| Linked Cell | 24 | font, borders | `#FA7D00`; bottom `double` `#FF8001` |

Gallery groups: "Good, Bad and Neutral" (Normal, Bad, Good, Neutral), "Data
and model" (Calculation, Check Cell, Explanatory Text, Input, Linked Cell,
Note, Output, Warning Text), "Titles and headings" (Heading 1–4, Title,
Total), "Custom". **Applying** sets, for each included aspect, exactly the
definition's value (a font is replaced wholesale — keys it lacks fall back
to Normal's), leaves excluded aspects alone and sets `namedStyle`; Normal
equals Clear formats. Later direct formatting keeps `namedStyle`.

**xlsx.** Read: `cellStyleXfs/xf[i]` decoded by SPEC-117's xf decoder is the
definition; an absent `applyX` on a style xf means included, `applyX="0"`
excluded; `cellStyles/cellStyle@name|xfId|builtinId|hidden` names it (the
first `cellStyle` per `xfId` wins); `cellXfs/xf@xfId` → that `CellStyle`'s
`namedStyle` (xfId of Normal → absent). Write (SPEC-122's writer, extended):
a style's `xf@xfId` is its named style's `sourceXfId`, else a `cellStyleXfs`
`xf` is appended (with `applyX="0"` for excluded aspects) plus a `cellStyle`
(`builtinId` only for the table's names; `customBuiltin` never), counts
updated; a source without `cellStyles` gains one holding Normal first.

**ods.** Read: a cell's automatic style's `style:parent-style-name` P (≠
`Default`) → `namedStyle` = P's `style:display-name`, else P with `_20_`
decoded to a space; the definition is the `office:styles` common
`table-cell` style's properties (SPEC-117's decoder), `includes` the aspects
it sets. Write: generated `de-ce` styles carry `style:parent-style-name` =
the named style's `sourceStyleName` (or `Default`); a named style without one
is appended to `styles.xml`'s `office:styles` (string splice before
`</office:styles>`) as `<style:style style:name="Heading_20_1"
style:display-name="Heading 1" style:family="table-cell"
style:parent-style-name="Default">` + properties (non-NCName characters
encoded `_xx_` in hex, as LibreOffice does); `styles.xml` is otherwise
byte-identical.

### The toolbar
| # | Band | Controls |
| --- | --- | --- |
| 1 | History | `sheet-format-painter` (undo/redo stay in the app bar, where SPEC-121 put `spreadsheet-undo`/`spreadsheet-redo`) |
| 2 | Type | `FontFamilySelect`, `FontSizeControl` |
| 3 | Character | bold, italic, underline, strikethrough |
| 4 | Colour | `TextColorPicker`, `SheetFillPicker`, `SheetBordersMenu` |
| 5 | Number | number-format dropdown, currency, percent, comma, decimals −/+ |
| 6 | Alignment | left/centre/right, top/middle/bottom, wrap, (SPEC-124 merge) |
| 7 | Indent & orientation | indent −/+, rotation |
| 8 | Styles | cell styles gallery, Clear |
New `ICON_PATHS` glyphs (stroke-only on the 20×20 grid): `fill-color`,
`borders`, `align-top`, `align-middle`, `align-bottom`, `wrap-text`,
`text-rotation`, `currency`, `percent`, `comma-style`, `decimals-increase`,
`decimals-decrease`, `cell-styles`, `merge-cells` (drawn here, used by 124).
A cell context menu `SheetCellContextMenu.vue` (`sheet-cell-menu`; right-
click, Shift+F10, the Menu key) opens with "Format cells…" and "Clear
contents" — this spec creates it; SPEC-124/125 and later specs append their
entries. The `SpreadsheetMenuBar` gains a Format menu (this spec adds it)
with Format cells… (`sheet-menu-format-cells`), Number format ▸ (the presets), Cell styles ▸, and
Clear ▸.

### Format Cells
| Tab | Controls (testid suffix after `format-cells-`) |
| --- | --- |
| Number | `category` listbox (General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Custom), `decimals` (0–30), `thousands`, `negative`, `symbol`, `type`, `custom-code`, `custom-list`, `sample` |
| Alignment | `horizontal` (General, Left (indent), Centre, Right (indent), Fill, Justify, Centre across selection, Distributed (indent)), `vertical` (Top, Centre, Bottom, Justify, Distributed), `indent` (0–250), `wrap`, `shrink`, `rotation` (−90…90 → 0…180 per xlsx; `vertical-text` → 255) |
| Font | `font-family`, `font-style` (Regular, Italic, Bold, Bold italic), `font-size`, `underline` (None, Single, Double, Single accounting, Double accounting), `font-color`, `strike`, `superscript`, `subscript` |
| Border | `border-style` (the 13 styles + None), `border-color`, presets `border-preset-none|outline|inside`, edges `border-edge-top|inside-horizontal|bottom|left|inside-vertical|right|diagonal-up|diagonal-down`, a preview box |
| Fill | `fill-color` (theme + palette + `no-fill`), `pattern-style` (`solid` and the 17 xlsx hatch patterns), `pattern-color`, a sample |
| Protection | `locked` (default checked), `hidden`, and the note "Locking cells or hiding formulas has no effect until the sheet is protected." |
Plus `format-cells-ok`, `format-cells-cancel`. The inside edges apply only
to multi-cell selections and become per-position patches as in
"Borders".

### Format painter and Clear
`capturePattern(sheet, range): StylePattern` records each position's
`styleId` (for a whole-row/column source, the rows'/spans' ids);
`applyPattern(model, sheetId, target, pattern)` tiles it from the target's
top-left — a single-cell target grows to the source's size, a larger target
repeats the pattern, a smaller one takes what fits. A multi-range source
disables the painter; a source over 1,000,000 positions is refused with
`FORMAT_TOO_LARGE_MESSAGE`. `clearRange`: **contents** removes `value`,
`formula`, `runs` (a partial array range is refused per SPEC-122), keeps
styles, drops cells left with nothing; **formats** removes `styleId` (and
row/span ids for full rows/columns, everything for the whole sheet) and
flattens `runs`; **all** does both.

### Shortcuts
| Id (`SHEET_SHORTCUTS`) | Windows / Linux | macOS |
| --- | --- | --- |
| `sheet.formatCells` | Ctrl+1 | Cmd+1 |
| `sheet.bold` / `.italic` / `.underline` | Ctrl+B / I / U | Cmd+B / I / U |
| `sheet.strikethrough` | Ctrl+5 | Cmd+Shift+X |
| `sheet.formatGeneral` | Ctrl+Shift+~ (`Backquote`) | Control+Shift+~ |
| `sheet.formatNumber` (id 4 `#,##0.00`) | Ctrl+Shift+! (`Digit1`) | Control+Shift+! |
| `sheet.formatTime` (id 18 `h:mm AM/PM`) | Ctrl+Shift+@ (`Digit2`) | Control+Shift+@ |
| `sheet.formatDate` (id 15 `d-mmm-yy`) | Ctrl+Shift+# (`Digit3`) | Control+Shift+# |
| `sheet.formatCurrency` (en-US id 7; else the locale's Currency) | Ctrl+Shift+$ (`Digit4`) | Control+Shift+$ |
| `sheet.formatPercent` (id 9 `0%`) | Ctrl+Shift+% (`Digit5`) | Control+Shift+% |
| `sheet.formatScientific` (id 11) | Ctrl+Shift+^ (`Digit6`) | Control+Shift+^ |
The number-format rows bind `{ code, ctrl: true, shift: true }` — physical
keys, as Excel binds them, so they work on every layout; macOS uses Control
because Cmd+Shift+3/4/5 are the system screenshot keys. `formatBinding`
shows `ctrl` as "Ctrl" (Windows/Linux) or "Control" (macOS). Bindings act
only when focus is in the grid (not the cell editor, the formula bar or a
dialog). Chrome and Firefox keep Ctrl/Cmd+1 … 8 for switching browser tabs,
so Ctrl/Cmd+1 (Format Cells) and Windows/Linux Ctrl+5 (strikethrough) are
best effort: the Format menu and `sheet-strikethrough` are the guaranteed
paths.

### Tests
Unit files per the Tests AC; e2e per the last AC. Round trips name their
relation: part-level `equalBytes` (`src/lib/core/equality.ts`) for untouched
parts, field-level `toEqual` on re-read styles.

### Specs this changes
- SPEC-117's `CellStyle.namedStyle` is populated from here; its readers gain
  `cellStyleXfs`/`cellStyles` and ods parent styles.
- SPEC-122's style writers gain `xfId`/`cellStyleXfs`/`cellStyles` and the
  ods `style:parent-style-name` + common-style append.
- SPEC-021's `src/lib/shortcuts.ts` gains `code`/`ctrl` binding fields and
  an optional `definitions` parameter; existing behaviour unchanged.
- SPEC-124 adds Merge to band 6, merges to the painter's pattern and
  unmerging to Clear formats; SPEC-129 adds conditional formats to Clear.
