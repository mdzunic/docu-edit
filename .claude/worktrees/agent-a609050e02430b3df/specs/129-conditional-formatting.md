---
id: SPEC-129
title: Conditional formatting
status: draft
area: formatting
complexity: complex
story_points: 13
depends_on: [SPEC-120, SPEC-123, SPEC-124, SPEC-125]
---
## Why
Open any real dashboard, grade book or budget and half of what it says is
said by conditional formatting: the overdue rows are red, the bar behind each
amount shows its share, the arrow says the number went up. After SPEC-123 a
workbook here can be formatted by hand, but the rules files carry — Excel's
`<conditionalFormatting>`/`<cfRule>` with their differential styles in
`styles.xml`'s `<dxfs>` and their Excel 2010 extensions in the worksheet's
`<extLst>`, LibreOffice's `calcext:conditional-formats` — are bytes this
product carries and never shows. Every highlighted cell opens plain, every
data bar is missing, and there is no way to add a rule.

The machinery to evaluate rules is already here: SPEC-120 evaluates formulas
and recalculates incrementally, SPEC-117 resolves a cell's style and renders
it, SPEC-123 edits styles over ranges. What is missing is the rule model, the
**differential style** (a partial style layered over the cell's own), an
evaluator that stays cheap on a scrolling grid whose rules need whole-range
statistics, renderers for data bars and icons, the wire mappings, and a
place to manage rules.

**This file is the normative source for the conditional-format rule model
(`ConditionalFormat`, `CfRule` and its kinds, global priority and
`stopIfTrue`), the differential style table (`Dxf`, `SpreadsheetModel.dxfs`,
`internDxf`) and how it layers over a cell's style, rule evaluation and its
statistics, data-bar and icon rendering and the icon artwork, and the xlsx
(`conditionalFormatting`, `dxfs`, x14 extensions) and ODS
(`calcext:conditional-formats`, legacy `style:map`) wire mappings.** Workbook-defined
table styles (which SPEC-133 preserves verbatim) index into the same `dxfs`
table — one more reason it is append-only.

## Acceptance criteria
- [ ] `src/lib/spreadsheet/conditionalFormats.ts` exports the model per "The rule model": `Sheet.conditionalFormats: ConditionalFormat[]` (`id`, `ranges: GridRange[]`, `rules: CfRule[]`, optional `pivot`/`sourceXml`), and `CfRule` with a sheet-global `priority` (1 = highest, unique, renumbered densely after every command), optional `stopIfTrue`, `dxfId`, `x14`/`x14Id`, and the kinds `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`, `aboveAverage`, `duplicates`, `text`, `blanks`, `errors`, `timePeriod` and `preserved`; readers, `createBlankWorkbook` and `addSheet` (SPEC-121) initialise `[]`
- [ ] `SpreadsheetModel.dxfs: Dxf[]` is the differential-style table per "Differential styles": a `Dxf` carries a tri-state `DxfFont` (colour, bold, italic, underline incl. `"none"`, strike), a `CellFill`, border edges and a number format; `layerDxf(base: CellStyle, dxf: Dxf): CellStyle` replaces exactly the properties the dxf sets (a `bold: false` dxf un-bolds a bold cell); `internDxf(model, dxf)` in `styles.ts` dedups by structural equality and only ever appends, so every `dxfId` an untouched sheet references stays valid; readers fill it and `createBlankWorkbook` starts it as `[]`
- [ ] Rule evaluation follows "Evaluation semantics" exactly — including a blank cell compared as `0` by `cellIs` against numbers (Excel formats empty cells for "less than 5"), `between` with bounds in either order, `expression`/`cellIs` formulas stored relative to the top-left of the format's **first** range (the origin), parsed once with SPEC-119's `parseFormulaAt(text, origin)` and evaluated at each cell by SPEC-120's `evaluate` with `host` = that cell and `origin` = the origin (never re-shifted text per cell), an expression true only for `TRUE` or a non-zero number (an error is false), duplicates compared case-insensitively with blanks never matching, text rules over the value's text, and `timePeriod` against "today" from `core/clock.ts`'s `now()` with weeks starting on Sunday
- [ ] Statistics-based kinds resolve per "Value objects and statistics": `min`/`max`/`autoMin`/`autoMax`/`num`/`percent`/`percentile`/`formula` thresholds over the **numeric** cells of all of the format's ranges; colour scales interpolate linearly per sRGB channel between the bracketing stops (2 or 3) and clamp outside them; data bars produce a fraction between `minLengthPct` and `maxLengthPct`, with negative values drawn from an axis per `axis` (`automatic` at |min| / (|min| + max), `middle`, `none`); icon sets pick the highest threshold the value meets (`gte` → `≥`, else `>`), `reverse` flipping the order; `top10` includes ties
- [ ] Rules combine per "Priority and layering": the rules covering a cell are walked by ascending `priority`; a true rule with `stopIfTrue` ends the walk; each true formatting rule contributes only the dxf properties no higher-priority rule has set; one fill (a dxf fill or a colour-scale colour, whichever ranks higher), at most one data bar and at most one icon survive; the merged dxf is layered over the cell's resolved style, a dxf `numberFormat` re-formats the display text through SPEC-118's `formatValue`, and `stopIfTrue` on a colour scale, data bar or icon set is preserved but ignored (Excel)
- [ ] `createCfEvaluator(model, sheetId, deps)` is viewport-driven per "The evaluator": `resultFor(row, col)` evaluates only cells the grid asks for, per-rule statistics are computed lazily once and cached until `invalidate` reports a changed cell inside that format's ranges (or `"all"`), each rule formula is parsed once per origin (`parseFormulaAt`'s LRU), `timePeriod` rules and formulas calling volatile functions re-evaluate on every recalculation pass — and a unit test proves 1,000 `resultFor` calls over a 100,000-cell colour scale compute its statistics exactly once
- [ ] `src/views/spreadsheet/useConditionalFormats.ts` owns one evaluator per open sheet and invalidates it from every `RecalcResult` (the changed keys), and with `"all"` on undo/redo, structural edits, rule commands and sheet switches; `SheetCellLayer.vue` applies `layerDxf` to the cell's `cellStyle(model, effectiveStyleId(sheet, row, col))` before SPEC-117's `cellCss(style, theme)`, so a highlighted cell's background, text colour, weight, borders and number format are the ordinary cell rendering with the layer applied; a cell with any result carries `data-cf="true"`
- [ ] A data bar renders as a `sheet-cf-databar` element (`aria-hidden="true"`, `data-width-pct`) absolutely positioned **behind** the cell's text inside its content box (2 px inset top and bottom): solid, or a gradient from the bar colour to the colour mixed 90% toward white computed in TypeScript (no CSS colour functions), an optional 1 px border, negative bars in the negative colour and a dashed 1 px axis line; `showValue: false` hides the text and keeps the bar
- [ ] An icon renders as a `sheet-cf-icon` inline `<svg>` (`role="img"`, `aria-label` such as "Green up arrow", `data-icon="3Arrows:2"`), 16 × 16 px at 100% zoom and scaled with the zoom, at the cell's left edge with the text aligned in the remaining width; the artwork is **this repo's own** — the glyph paths and colours in `src/lib/spreadsheet/cfIcons.ts` per "Icon artwork", never copied from Office's or LibreOffice's image resources — and covers every Excel set name, the x14 sets and LibreOffice's `3Smilies`/`3ColorSmilies`
- [ ] xlsx read per "xlsx wire mapping": every `<conditionalFormatting sqref [pivot]>` and `<cfRule>` type in the table, formulas through `fromXlsxFormula`; `xl/styles.xml` `<dxfs>` into `dxfs` by index with Excel's quirks (a dxf `patternFill` with no `patternType` is solid and its colour is `bgColor`; `<b val="0"/>` is explicit not-bold; theme colours with `tint` via SPEC-117's mapping); a rule kind or child this spec cannot model becomes a `preserved` rule holding its XML verbatim with its `sqref` still modelled
- [ ] x14 extensions per "x14": a base rule's `<extLst>` `x14:id` GUID pairs it with the `x14:cfRule` of the same id in the worksheet's `{78C0D931-6437-407d-A8EE-F0AAD7539E65}` ext, whose data-bar settings overlay the base rule's (`minLength`/`maxLength`, `gradient`, `border`, `direction`, negative colours, axis, `autoMin`/`autoMax`); x14-only rules (`3Stars`, `3Triangles`, `5Boxes`, `NoIcons`, and cell/expression rules whose `xm:f` references another sheet, with their inline `x14:dxf`) are modelled with `x14: true`; custom icon sets and anything else stay `preserved` (`where: "x14"`); every other `<ext>` in the worksheet's `<extLst>` is untouched
- [ ] xlsx write: on a dirty sheet the run of `<conditionalFormatting>` elements is replaced from the model at its CT_Worksheet position (after `phoneticPr`, before `dataValidations`), each block emitted from its `sourceXml` while unchanged and serialised otherwise; only the `{78C0D931-…}` ext is rewritten (created in Excel's ext order if absent, namespaces declared on the `<ext>` itself); a data bar is always written as a base rule plus an x14 rule sharing a GUID (fresh upper-case `{…}` GUIDs from `crypto.randomUUID()` for new rules); new dxfs are appended to `<dxfs>` with `count` updated and solid fills written as `bgColor`; the text, blank, error and time-period rules' `<formula>` is regenerated from the model with Excel's formula for the kind (Excel evaluates the formula, not the attribute)
- [ ] ODS read per "ODS wire mapping": `calcext:conditional-formats` inside each `table:table` — `calcext:condition` with every `calcext:value` form in the table, `calcext:color-scale`, `calcext:data-bar`, `calcext:icon-set`, `calcext:date-is` — with `calcext:apply-style-name` resolved to a `Dxf` (`sourceStyleName` kept) from the named cell style in `styles.xml` (its own properties and its ancestors' up to, not including, `Default`) and formulas rebased from `calcext:base-cell-address` onto the first range's top-left; a table with no `calcext:` rules reads legacy `style:map` conditions on its automatic cell styles as `cellIs`/`expression` rules over the cells (and column defaults) using that style, without materialising repeated empty cells
- [ ] ODS write: `calcext:conditional-formats` only (never `style:map`), in priority order; each dxf without a still-valid `sourceStyleName` becomes a generated named style `de-cf<n>` (`style:display-name` "Conditional <n>", parent `Default`) spliced before `</office:styles>` in `styles.xml`, which is otherwise byte-verbatim; `xmlns:calcext` is spliced into the content root when absent; on a rewritten table, every cell style whose source automatic style carries a `style:map` is re-emitted as a generated `de-ce` style without it, so a deleted rule cannot come back in LibreOffice
- [ ] Round trip: an untouched workbook downloads byte-identical part for part (`equalBytes`); a dirty sheet whose rules were not edited re-emits each block's `sourceXml` verbatim; add/edit → save → reopen reproduces `conditionalFormats` and the referenced `dxfs` under `toEqual` minus minted ids and `sourceXml`, in both formats; "Download a copy as…" carries every modelled kind across formats per the tables, SPEC-121's `UNCONVERTIBLE_CONTENT` loses its "conditional formatting" row and gains "some conditional formatting" (marker: a `preserved` rule, or reversed icon order going to `.ods`)
- [ ] Format ▸ "Conditional formatting…" (`sheet-menu-conditional-formatting`) opens the rules manager (`cf-manager`) as the `conditionalFormats` tab of the spreadsheet side panel per "The side panel": `cf-scope` ("Current selection" / "This sheet"), `cf-rule-list` (`role="listbox"`) of `cf-rule` options in priority order, each with `cf-rule-preview`, `cf-rule-description` (strings per "Rule descriptions"), an editable `cf-rule-applies-to` and, for formatting kinds, `cf-rule-stop-if-true`; actions `cf-add`, `cf-edit`, `cf-delete`, `cf-move-up`, `cf-move-down`, `cf-clear-selection` ("Clear rules from selected cells"), `cf-clear-sheet` ("Clear rules from entire sheet")
- [ ] The rule editor (`cf-editor`, `cf-editor-back`) offers Excel's six rule types in `cf-rule-type` and each type's fields per "The rule editor" (operators, `cf-value-1`/`cf-value-2` accepting locale values via `parseInput` or `=` formulas, colour-scale stops, data-bar options, `cf-icon-set` with editable thresholds and "Reverse icon order" — hidden for ODS workbooks, where it has no wire form), `cf-format` with Excel's six presets plus "Custom format…" opening SPEC-123's `FormatCellsDialog.vue` (`format-cells-dialog`) in a differential mode showing only its Number, Font (colour, bold, italic, underline, strike — no family or size), Border and Fill tabs, a live `cf-preview`, `cf-save` and `cf-cancel`; an unparseable formula or range is an inline error and saves nothing
- [ ] `cf-add` opens the quick presets per "Presets" — Greater than…, Less than…, Between…, Equal to…, Text that contains…, A date occurring…, Duplicate values…, Top 10 items…, Bottom 10 items…, Above average, Below average, Data bars (six colours, gradient and solid), Colour scales (six), Icon sets (every Excel set) and New rule… — each applying to the current selection; a new rule takes priority 1 (Excel puts it first)
- [ ] Commands in `conditionalFormats.ts` — `addConditionalFormat`, `updateRule`, `deleteRule`, `moveRule`, `setFormatRanges`, `clearRulesInRanges` (subtracting ranges with `subtractRanges(ranges, cut): GridRange[]`, which this spec adds to `address.ts` — SPEC-130 states the identical contract; a block left with no range is removed) — replace the arrays and objects they touch (copy-on-write), drop the touched block's `sourceXml`, renumber priorities, set `Sheet.dirty`, and record one undo entry labelled "Conditional formatting"
- [ ] Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends its `ReaderUpgrade`: a clean sheet is re-read; for a dirty sheet the step reads `dxfs` once (when absent) from the retained `xl/styles.xml` or ODS `styles.xml` and parses only the sheet's conditional formats from its retained part or `sourceXml` — the retained source has carried SPEC-124's and SPEC-125's retained-reference rewrites since it was read, so the parsed positions already match the current model — tested with a synthetic previous-version record holding a dirty and a clean sheet
- [ ] Registers `{ id: "conditionalFormats", adjust }` in `REFERENCE_ADJUSTERS` per "Structural edits" and, in the same change, retires its rows from SPEC-124's "Retained references" tables (`conditionalFormatting@sqref`, `cfRule/formula`, `x14:conditionalFormatting/xm:sqref` and its `xm:f`; `calcext:conditional-format@calcext:target-range-address`, `calcext:condition@calcext:value`/`@calcext:base-cell-address`, `calcext:formatting-entry@calcext:value`): inserting two rows above a format on `B5:B20` moves it to `B7:B22` and its expression `B5>$Z$1` to `B7>$Z$1`; deleting the format's first row re-anchors the formula onto the surviving top-left with `adjustAnchoredFormula` instead of turning it into `#REF!`; a fully deleted range is dropped and an empty format removed — the test joins `structure.spec.ts`'s registry contract test; sheet rename and delete run SPEC-121's `remapSheetReferences` over every rule formula and value object
- [ ] Registers `{ id: "conditionalFormats", move }` in SPEC-125's `MOVE_ADJUSTERS` per "Moves": cutting `B5:B6` out of a format on `B5:B20` and pasting it at `F5` leaves the format on `B7:B20` with its expression re-anchored to `B7>$Z$1` and creates one on `F5:F6` carrying the same rules with `F5>$Z$1` (the reference followed its moved cell; `$Z$1` did not move); a format wholly inside the source keeps its id and moves; the destination rectangle's previous rules are removed — the test joins `move.spec.ts`'s registry contract test
- [ ] Fixtures from `spreadsheet/fixtures/build.ts`: `e2e/fixtures/spec129-cf.xlsx` (Excel-shaped per "Fixtures": every kind, x14 data bar pairs with GUIDs, an x14-only `3Stars` rule, a custom-icon x14 rule, a sparkline ext that must stay verbatim, theme colours with tints, bgColor-only dxfs) and `spec129-cf.ods` (LibreOffice-shaped `calcext:` rules over named styles) plus `spec129-legacy.ods` (`style:map` only)
- [ ] Tests: `spreadsheet/conditionalFormats.spec.ts` (every kind's semantics, value objects, interpolation, bar geometry, icon selection, priority/stopIfTrue/layering, commands), `conditionalFormats.evaluator.spec.ts` (caching, invalidation, volatility under `setClock`), `cfIcons.spec.ts` (every set name resolves, glyph counts per set), `xlsx/conditionalFormats.spec.ts`, `ods/conditionalFormats.spec.ts`, `styles.spec.ts` (`internDxf`, dxf quirks), the adjuster test in `structure.spec.ts`, the move adjuster test in `move.spec.ts`, the upgrade step in `open.spec.ts`, `roundTrip.spec.ts` additions, `views/spreadsheet/useConditionalFormats.spec.ts` and the manager component spec
- [ ] e2e (`e2e/SPEC-129.spec.ts`): open `spec129-cf.xlsx` → `B3` (value 140, rule "> 100") has background `rgb(255, 199, 206)` and text `rgb(156, 0, 6)` → the `C` column's `sheet-cf-databar` widths increase with their values and the negative cell's bar sits left of the axis → `D2` shows a `sheet-cf-icon` with `data-icon="3Arrows:2"` → Format ▸ Conditional formatting… → select `B2:B13` → Add ▸ Greater than… → 120 → "Green fill with dark green text" → Save → `B3` turns green (the new rule ranks first) → `cf-move-down` → `B3` is red again → type 130 into `B4` → it turns red → Ctrl/Cmd+Z → plain again → download → reopen → the manager lists the same rules in the same order; then open `spec129-cf.ods` → its `calcext:condition` cells show their named style's fill

## Out of scope
Copying rules with cells — paste-all and the format painter carrying
conditional formats (Excel does both) are Backlog; copy-pasted values land
under whatever rules cover the destination (a cut-paste or drag-move does
carry its cells' rules, per "Moves"). Choosing custom icons per threshold
(x14 `custom="1"` with `x14:cfIcon`) — read, honoured as `preserved`,
editable only by deletion. Data bars and icons in right-to-left sheets
(no RTL sheets in this queue). "Show rules for" another sheet (the manager
scopes to the active sheet). Applying conditional formats to a pivot table's
fields (`pivot="1"` is preserved; its ranges evaluate as plain ranges until
SPEC-139). Sparklines (preserved verbatim, Backlog). Excel's "Apply rule to"
whole-column performance beyond the viewport-driven bound here (SPEC-137).
Filtering or sorting by conditional-format colour (SPEC-128 keeps those
criteria `preserved`).

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — the plain-JSON model,
  copy-on-write rows, sparse storage, the retained source as the splice base
  (xlsx parts in `opaque`, ODS `sourceXml` and `odsContentEnvelope`);
  `GridRange { top; left; bottom; right }` (0-based, inclusive), `CellRef`,
  `parseSqref`, `formatRange`, `getCell`; `open.ts`'s `ReaderUpgrade`
  `{ from; upgradeDirtySheet(model, sheet, sourceParts) }` and
  `READER_UPGRADES`.
- SPEC-117 — `CellStyle { font?: CellFont; fill?: CellFill; borders?: { top?;
  right?; bottom?; left?: CellBorderEdge; diagonal? }; alignment?;
  numberFormat?; protection?; namedStyle?; sourceXf?; sourceStyleName? }`,
  `CellFont { family?; sizePt?; bold?: true; italic?: true; underline?:
  "single" | "double" | "singleAccounting" | "doubleAccounting"; strike?:
  true; color?: ModelColor; vertAlign? }`, `CellFill` (`solid` | `pattern` |
  `gradient`), `CellBorderEdge { style: SheetBorderStyle; color? }`;
  `effectiveStyleId(sheet, row, col)` and `cellStyle(model, id)` in
  `styles.ts`; xlsx colours: `theme` n → slot with Excel's swap (0→lt1,
  1→dk1, 2→lt2, 3→dk2, 4..9→accent1..6), `tint` t<0 → `lumMod: 1+t`, t>0 →
  `lumMod: 1−t, lumOff: t` (`SheetColorRef` keeps the wire `xlsxTint`),
  `indexed` legacy palette, ARGB `FFrrggbb` → `#rrggbb`;
  `SpreadsheetModel.theme?: DocumentTheme`; rendering: `cellCss(style,
  theme)` in `cellRender.ts`, drawn by `SheetCellLayer.vue` inside
  `SheetGrid.vue` (`sheet-cell`, `data-ref`), only the viewport's cells
  mounted. Core: `ModelColor`, `resolveColor(color, theme)`,
  `DEFAULT_THEME` in `src/lib/core/theme.ts`.
- SPEC-118 — `formatValue(value, code, ctx)`; ODS `number:*` styles ↔ format
  codes (`numberFormat/odsNumberStyles.ts`).
- SPEC-119 — canonical formula text; `parseFormula`; `parseFormulaAt(text,
  host): { node; origin } | FormulaSyntaxError` (LRU by
  `relativeFormulaKey`); `shiftFormula(text, dRow, dCol)`;
  `fromXlsxFormula`, `toXlsxFormula`, `fromOdsFormula`, `toOdsFormula`
  (token-level).
- SPEC-120 — `evaluate(node, ctx: EvalContext): EvalValue` resolving each
  relative part as `part + (host − origin)`; `EvalContext { model; host;
  origin; arrayMode; settings; now(); random(); valueAt(sheetId, row, col);
  recordDependency(area); resolveName? }`; comparison semantics (numbers <
  text, case-insensitive text, blank = 0 or `""`, 15-significant-digit
  numeric comparison), text coercion, `FunctionSpec` with `volatile?: true`,
  `recalculate(model, changed: CellKey[]): RecalcResult`, `PERCENTILE.INC`'s
  interpolation (reused for `percentile`).
- SPEC-121 — `useSpreadsheetUndo` `{ record(label, scope) /* before */,
  commit(alsoRecalc?) /* after */, … }`: sheets snapshotted as `{ ...sheet }`
  with object fields `structuredClone`d except `SHEET_BY_REFERENCE_FIELDS`
  (`rows`, `sourceXml`), plus `WORKBOOK_SNAPSHOT_FIELDS` (`styles` by
  reference — `internStyle` replaces the array when it appends);
  `Sheet.dirty`; `writeXlsx`/`writeOds`; `stripRetention`;
  `remapSheetReferences(canonical, remap)` for sheet rename/delete; the
  `UNCONVERTIBLE_CONTENT` table behind the "Not included in the .ods copy:
  …" toast.
- SPEC-122 — the dirty-sheet splice in CT_Worksheet order (… `mergeCells`,
  `phoneticPr`, **`conditionalFormatting`**, `dataValidations`, `hyperlinks`
  … `extLst`); `xl/styles.xml` append-only; ODS dirty tables re-serialised
  with generated `de-` automatic styles; `internStyle(model, style)`;
  `parseInput(text, locale, dateSystem)`.
- SPEC-123 — `applyStylePatch`, `FormatCellsDialog.vue`
  (`format-cells-dialog`, tabs `format-cells-tab-number|alignment|font|
  border|fill|protection`) and its colour pickers, the `SpreadsheetMenuBar`
  Format menu.
- SPEC-124 — `StructuralEdit`, `adjustIndex`, `adjustRange(range, edit):
  GridRange | null`, `adjustFormula(text, hostSheetId, edit, sheetNameOf)`,
  `REFERENCE_ADJUSTERS: { id; adjust(model, edit): void }[]` in
  `structure.ts` (adjusters replace, never mutate) and
  `rewriteRetainedReferences`, whose conditional-format rows this spec
  retires.
- SPEC-125 — `src/lib/spreadsheet/move.ts`: `RangeMove { fromSheetId; from:
  GridRange; toSheetId; toTop; toLeft }`, `MoveAdjuster { id; move(model,
  move): void }`, `MOVE_ADJUSTERS` (registered: `cells`, `formulas`,
  `merges`, `retained`), `moveRange`, `repointFormula(text, hostSheetId,
  move, sheetNameOf)` (a reference entirely inside the source follows the
  move; one entirely inside the destination becomes `#REF!`; anything else
  is unchanged) and `moveMapper`; `move.spec.ts` holds the contract test
  later move adjusters join.
- `src/lib/spreadsheet/ods/rangeAddress.ts` (SPEC-121) —
  `odsRangeAddressToCanonical(text): string | null` /
  `canonicalToOdsRangeAddress(canonical): string | null` (`"Sheet1.A1:Sheet1.A10"`
  ↔ `"Sheet1!A1:A10"`, space-separated lists ↔ parenthesised unions).
- `specs/065-unified-side-panel.md` — `src/components/SidePanel.vue` is
  presentational: props `side`, `widthPx`, `collapsed`, `tabs` (`{ id;
  title; icon; badge }[]`), `activePane`; emits `activate`, `close`,
  `toggle-side`, `toggle-collapse`, `update:widthPx`, `commit-width`,
  `nudge-width`; testids `side-panel`, `side-panel-tab-<id>`,
  `side-panel-close`. Its state composable `src/views/document/useSidePanel.ts`
  extends the text editor's `EditorContext` and a closed seven-pane `PaneId`;
  `src/lib/sidePanel.ts` exports `MIN_PANEL_WIDTH`, `DEFAULT_PANEL_WIDTH`,
  `clampPanelWidth`.
- `src/lib/core/clock.ts` — `now()`/`setClock()`; tests never use real time.

### The rule model (`src/lib/spreadsheet/conditionalFormats.ts`)
```ts
export interface ConditionalFormat {
  id: string;                     // minted at read; never written
  ranges: GridRange[];            // sqref / calcext:target-range-address, source order
  rules: CfRule[];
  pivot?: true;                   // xlsx @pivot, preserved
  sourceXml?: string;             // verbatim source element (+ preceding whitespace) while unchanged
}
interface CfRuleBase {
  id: string;
  priority: number;               // sheet-global; 1 = highest; unique and dense
  stopIfTrue?: true;
  dxfId?: number;                 // into model.dxfs, for the formatting kinds
  x14?: true;                     // written into the x14 extLst (xlsx)
  x14Id?: string;                 // "{GUID}" pairing a base rule with its x14 extension
}
export type CfOperator = "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual"
  | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";
export type CfTimePeriod = "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek"
  | "nextWeek" | "thisMonth" | "lastMonth" | "nextMonth" | "thisYear" | "lastYear" | "nextYear";
export interface CfValue { type: "min" | "max" | "autoMin" | "autoMax" | "num" | "percent" | "percentile" | "formula";
  value?: string }                // number text or canonical formula (absolute, evaluated at the origin)
export interface CfDataBar { min: CfValue; max: CfValue; color: ModelColor; gradient: boolean;
  border?: ModelColor; negativeColor?: ModelColor; negativeBorderColor?: ModelColor;
  axis: "automatic" | "middle" | "none"; axisColor?: ModelColor;
  direction: "context" | "leftToRight" | "rightToLeft"; minLengthPct: number; maxLengthPct: number; showValue: boolean }
export interface CfIconSet { set: IconSetName; thresholds: { at: CfValue; gte: boolean }[];  // one per icon, lowest first
  reverse: boolean; showValue: boolean }
export type CfRule = CfRuleBase & (
  | { kind: "cellIs"; operator: CfOperator; formulas: string[] }        // 1, or 2 for (not)between
  | { kind: "expression"; formula: string }
  | { kind: "colorScale"; stops: { at: CfValue; color: ModelColor }[] }  // 2 or 3
  | { kind: "dataBar"; bar: CfDataBar }
  | { kind: "iconSet"; icons: CfIconSet }
  | { kind: "top10"; bottom: boolean; percent: boolean; rank: number }
  | { kind: "aboveAverage"; below: boolean; equalAverage: boolean; stdDev?: 1 | 2 | 3 }
  | { kind: "duplicates"; unique: boolean }
  | { kind: "text"; operator: "contains" | "notContains" | "beginsWith" | "endsWith"; text: string }
  | { kind: "blanks"; not: boolean }
  | { kind: "errors"; not: boolean }
  | { kind: "timePeriod"; period: CfTimePeriod }
  | { kind: "preserved"; xml: string; where: "base" | "x14" | "ods" });
export type IconSetName = "3Arrows" | "3ArrowsGray" | "3Flags" | "3TrafficLights1" | "3TrafficLights2"
  | "3Signs" | "3Symbols" | "3Symbols2" | "3Stars" | "3Triangles" | "3Smilies" | "3ColorSmilies"
  | "4Arrows" | "4ArrowsGray" | "4RedToBlack" | "4Rating" | "4TrafficLights"
  | "5Arrows" | "5ArrowsGray" | "5Rating" | "5Quarters" | "5Boxes" | "NoIcons";
```
The **origin** of a format is `ranges[0]`'s top-left: every `cellIs` and
`expression` formula is stored relative to it, parsed once by
`parseFormulaAt(formula, origin)` and evaluated at cell `(r, c)` with
`evaluate(node, { ...base, host: { sheetId, row: r, col: c }, origin,
arrayMode: false })` — the engine's own `part + (host − origin)` rule, so
no text is re-shifted per cell. Value-object formulas are absolute and
evaluated with `host` = `origin`. Default icon
thresholds: 3 icons `percent` 0/33/67, 4 icons 0/25/50/75, 5 icons
0/20/40/60/80, all `gte`.

### Differential styles
```ts
export interface DxfFont { color?: ModelColor; bold?: boolean; italic?: boolean;
  underline?: CellFont["underline"] | "none"; strike?: boolean }
export interface Dxf {
  font?: DxfFont;
  fill?: CellFill;
  borders?: { top?: CellBorderEdge; right?: CellBorderEdge; bottom?: CellBorderEdge; left?: CellBorderEdge };
  numberFormat?: string;
  sourceXml?: string;             // xlsx <dxf> as read — keeps alignment/protection children
  sourceStyleName?: string;       // ods named style it came from
}
// SpreadsheetModel gains: dxfs: Dxf[]   — index = xlsx dxfId; append-only
export function layerDxf(base: CellStyle, dxf: Dxf): CellStyle;          // conditionalFormats.ts
export function internDxf(model: SpreadsheetModel, dxf: Dxf): number;   // styles.ts, beside internStyle
```
`internDxf` compares the modelled fields (not `sourceXml`/`sourceStyleName`)
and returns an existing index when equal; when it appends it **replaces**
the `dxfs` array (as `internStyle` does `styles`), and `dxfs` joins SPEC-121's
`WORKBOOK_SNAPSHOT_FIELDS` by reference, so an undo past a new dxf restores
the shorter array. `conditionalFormats` joins `SHEET_BY_REFERENCE_FIELDS`:
every command replaces the array and the formats it touches.

### Evaluation semantics
| Kind | True when (cell value `v`) |
| --- | --- |
| `cellIs` | SPEC-120 comparison of `v` with the formula results; blank is `0` against a number and `""` against text; `between` is inclusive with bounds in either order |
| `expression` | the shifted formula evaluates to `TRUE` or a non-zero number |
| `top10` | `v` numeric and ≥ (bottom: ≤) the `rank`-th largest (smallest) numeric value, ties included; `percent`: rank `max(1, floor(n × rank / 100))` |
| `aboveAverage` | `v` numeric and > mean (≥ with `equalAverage`; below mirrored); `stdDev` k: `v > mean + k·σ` (`v < mean − k·σ`), σ the population standard deviation |
| `duplicates` | `v` non-blank and its key occurs more than once (`unique`: exactly once) across the format's ranges; key = number, or lower-cased text, booleans and errors by their text |
| `text` | the value's text (SPEC-120 coercion) contains / does not contain / begins / ends with `text`, case-insensitive; `notContains` is true for a blank |
| `blanks` | the value's text trimmed is `""` (`not`: is not) |
| `errors` | `v` is a `CellError` (`not`: is not) |
| `timePeriod` | `v` numeric and `floor(v)` falls in the period around today's serial (`now()`, local time, `dateToSerial` in the workbook's system); weeks start Sunday; `last7Days` = today − 6 … today |
| `colorScale`, `dataBar`, `iconSet` | `v` is a number (dates included; text, booleans and blanks never) |
| `preserved` | never — it is carried, not evaluated |

### Value objects and statistics
Statistics are taken over the numeric values (not booleans, not numeric
text) in **all** of the format's ranges: count, min, max, mean, population
σ, the sorted values (for `percentile` and `top10`) and the duplicate-key
counts. `min`/`max` → those; `autoMin` = min(0, min), `autoMax` = max(0,
max); `num` → the number, or the formula's result when `value` starts with a
formula; `percent` p → min + (max − min)·p/100; `percentile` p → the
`PERCENTILE.INC` of the values at p/100; `formula` → evaluated at the origin.
- **Colour scale.** `v` ≤ the first stop → its colour; ≥ the last → its
  colour; between two stops → each sRGB channel interpolated linearly by
  `(v − s₀)/(s₁ − s₀)` and rounded; coincident stops resolve to the lower
  stop's colour.
- **Data bar.** With lo, hi the resolved min and max: if lo ≥ 0 or `axis`
  is `none`, width = minLen + (maxLen − minLen)·(v − lo)/(hi − lo) (clamped;
  hi = lo → maxLen), drawn from the leading edge. Otherwise the axis sits at
  `a` = |lo|/(|lo| + hi) of the bar area (`middle`: 0.5): a positive `v`
  draws `(v/hi)·(1 − a)` right of the axis, a negative `v` draws
  `(|v|/|lo|)·a` left of it in the negative colour. `rightToLeft` mirrors;
  `context` is left to right.
- **Icon set.** Index = the highest `i` whose threshold `v` meets (`≥` when
  `gte`, else `>`), the first threshold being the floor; `reverse` maps `i`
  to `n − 1 − i`; `NoIcons` draws nothing.

### Priority and layering
For a cell: collect the rules of every format whose ranges contain it, sort
by `priority`, and walk. A true formatting rule (every kind except colour
scale, data bar, icon set) merges its dxf property by property — font
colour, bold, italic, underline, strike, fill, each border edge, number
format — taking only properties still unset; if it has `stopIfTrue` the walk
ends. A colour scale sets the fill if still unset; the first data bar and the
first icon win. The result:
```ts
export interface CfCellResult {
  dxf?: Dxf;                       // merged
  scaleColor?: string;             // "#rrggbb", when a colour scale won the fill
  dataBar?: { startPct: number; widthPct: number; negative: boolean; axisPct?: number;
              color: string; endColor?: string; borderColor?: string; axisColor?: string };
  icon?: { set: IconSetName; index: number };
  hideValue?: true;                // the winning bar or icon has showValue false
}
```

### The evaluator
```ts
export interface CfEvaluator {
  resultFor(row: number, col: number): CfCellResult | undefined;
  invalidate(changed: CellKey[] | "all"): void;
}
export function createCfEvaluator(model: SpreadsheetModel, sheetId: string, deps: {
  base: Omit<EvalContext, "host" | "origin" | "arrayMode">;  // valueAt reads stored values; recordDependency a no-op
  theme: DocumentTheme;                                      // model.theme ?? DEFAULT_THEME
  onStatistics?(formatId: string): void;                     // test hook
}): CfEvaluator;
```
Runtime caches (`Map`s) live in the evaluator, never in the model. A
formula that does not parse (`FormulaSyntaxError`) makes its rule false
everywhere; an `Uncomputable` result is false. A
per-cell result cache is cleared by every `invalidate`; statistics for a
format are cleared when a changed key lies in its ranges; formula-valued
thresholds are re-evaluated whenever the per-cell cache is cleared; each
format keeps a bounding box so `resultFor` skips formats that cannot cover
the cell.

### Icon artwork (`src/lib/spreadsheet/cfIcons.ts`)
```ts
export type CfGlyph = "arrowUp" | "arrowUpRight" | "arrowRight" | "arrowDownRight" | "arrowDown"
  | "circle" | "circleRimmed" | "triangleUp" | "triangleDown" | "dash" | "diamond" | "flag"
  | "check" | "cross" | "exclaim" | "checkCircle" | "crossCircle" | "exclaimCircle"
  | "star" | "starHalf" | "starEmpty" | "bars0" | "bars1" | "bars2" | "bars3" | "bars4"
  | "quarter0" | "quarter1" | "quarter2" | "quarter3" | "quarter4"
  | "box0" | "box1" | "box2" | "box3" | "box4" | "smileHappy" | "smileNeutral" | "smileSad";
export interface CfIconDef { glyph: CfGlyph; color: string; label: string }
export const CF_ICON_COLORS = { green: "#2E9E4F", amber: "#E8A317", red: "#D64541", grey: "#808080",
  black: "#333333", pink: "#E99A9A", blue: "#3B78C3", gold: "#F2B01E" } as const;
export const CF_ICON_SETS: Record<IconSetName, CfIconDef[]>;   // lowest → highest
export function cfGlyphPaths(glyph: CfGlyph): { d: string; paint: "icon" | "white" | "none" }[];  // 16-unit viewBox
```
Authored here from geometry (the `core/chartIcons.ts` precedent): arrows,
circles, triangles, flags, ticks, crosses, stars, signal bars, quarter pies,
boxes, faces. Sets, lowest → highest: `3Arrows` red down, amber right, green
up (`3ArrowsGray` all grey); `3Flags` red, amber, green flag;
`3TrafficLights1` red, amber, green circle (`…2` rimmed); `3Signs` red
diamond, amber triangle, green circle; `3Symbols` red cross, amber
exclamation, green tick, each circled (`3Symbols2` uncircled); `3Stars` empty,
half, full gold star; `3Triangles` red down, amber dash, green up; `3Smilies`
black sad, neutral, happy (`3ColorSmilies` red, amber, green); `4Arrows` red
down, amber down-right, amber up-right, green up (`4ArrowsGray` grey);
`4RedToBlack` black, grey, pink, red circle; `4Rating` bars1–bars4 in blue;
`4TrafficLights` black, red, amber, green circle; `5Arrows` red down, amber
down-right, amber right, amber up-right, green up (`5ArrowsGray` grey);
`5Rating` bars0–bars4; `5Quarters` quarter0–quarter4 in black; `5Boxes`
box0–box4 in blue; `NoIcons` none. Labels read "<colour> <shape>" ("Green up
arrow", "Half star", "3 of 4 bars").

### xlsx wire mapping
| `cfRule@type` | Kind | Carried by |
| --- | --- | --- |
| `cellIs` | `cellIs` | `@operator`, `<formula>` ×1–2 |
| `expression` | `expression` | `<formula>` |
| `colorScale` | `colorScale` | `<colorScale>`: 2–3 `<cfvo type val>` then as many `<color>` |
| `dataBar` | `dataBar` | `<dataBar minLength="10" maxLength="90" showValue>`, 2 `<cfvo>`, `<color>`; + x14 |
| `iconSet` | `iconSet` | `<iconSet iconSet="3TrafficLights1" showValue reverse>`, n `<cfvo type val gte>` |
| `top10` | `top10` | `@rank`, `@percent`, `@bottom` |
| `aboveAverage` | `aboveAverage` | `@aboveAverage` (default 1), `@equalAverage`, `@stdDev` |
| `duplicateValues` / `uniqueValues` | `duplicates` | — |
| `containsText` / `notContainsText` / `beginsWith` / `endsWith` | `text` | `@operator`, `@text`; `<formula>` regenerated |
| `containsBlanks` / `notContainsBlanks` | `blanks` | `<formula>` regenerated |
| `containsErrors` / `notContainsErrors` | `errors` | `<formula>` regenerated |
| `timePeriod` | `timePeriod` | `@timePeriod`; `<formula>` regenerated |
| `dataBar`/`iconSet` sets with no 2007 form, unknown types | x14 or `preserved` | see x14 |

Regenerated formulas, with `A1` = the origin: `NOT(ISERROR(SEARCH("t",A1)))`,
`ISERROR(SEARCH("t",A1))`, `LEFT(A1,LEN("t"))="t"`, `RIGHT(A1,LEN("t"))="t"`,
`LEN(TRIM(A1))=0` / `>0`, `ISERROR(A1)` / `NOT(ISERROR(A1))`, and for time
periods Excel's own: `FLOOR(A1,1)=TODAY()` (±1 for yesterday/tomorrow),
`AND(TODAY()-FLOOR(A1,1)<=6,FLOOR(A1,1)<=TODAY())`,
`AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))`,
`AND(MONTH(A1)=MONTH(EDATE(TODAY(),0-1)),YEAR(A1)=YEAR(EDATE(TODAY(),0-1)))`
(`0+1` for next month), and the `WEEKDAY(TODAY())` forms for the weeks; the
LibreOffice-only year periods are written as `expression` rules
(`YEAR(A1)=YEAR(TODAY())`, ±1). The evaluator uses the model's semantics;
the formula is what Excel reads.

**dxfs.** `<dxf>` children in order `font`, `numFmt`, `fill`, `alignment`,
`protection`, `border`. Font: `<b/>`/`<b val="1"/>` → true, `val="0"` →
false (likewise `i`, `strike`); `<u/>` → `single`, `<u val="none"/>` →
`"none"`; `<color>` per SPEC-117. Fill: `patternFill` with no `patternType`
or `solid` → `{ kind: "solid", color: bgColor ?? fgColor }`; other patterns
and `gradientFill` per SPEC-117. `numFmt@formatCode` → `numberFormat` (on
write, the id is the one `<numFmts>` already gives that code, else the next
free id ≥ 164). `<dxfs>` sits after `cellStyles`, before `tableStyles`; when
absent it is created there.

### x14
- The worksheet `<extLst>` ext `uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"`
  holds `x14:conditionalFormattings` › `x14:conditionalFormatting`
  (`xmlns:xm`) › `x14:cfRule @type @priority @id` + `xm:sqref`. A base rule's
  `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{GUID}`
  pairs with the `x14:cfRule` of that `@id`.
- Paired data bar: `x14:dataBar @minLength @maxLength @gradient @border
  @direction @negativeBarColorSameAsPositive @negativeBarBorderColorSameAsPositive
  @axisPosition`, `x14:cfvo` (incl. `autoMin`/`autoMax`, values in `xm:f`),
  `x14:borderColor`, `x14:negativeFillColor`, `x14:negativeBorderColor`,
  `x14:axisColor` overlay the base `dataBar`; the base rule's `priority`
  counts. Defaults written for a new bar: base `minLength`/`maxLength`
  omitted, x14 `minLength="0" maxLength="100"`, `autoMin`/`autoMax`,
  negative fill `FFFF0000`, axis `FF000000`.
- Standalone x14 rules (no base twin) take their `@priority` into the same
  global order; their `x14:dxf` is inline (interned into `dxfs`, written
  inline again). Written to x14: icon sets `3Stars`, `3Triangles`, `5Boxes`,
  `NoIcons`; `cellIs`/`expression` rules whose formulas reference another
  sheet; the x14 half of every data bar. Everything else is base-only.
- New exts go in Excel's order — conditional formattings `{78C0D931…}`, data
  validations `{CCE6A557…}` (SPEC-130), then the rest as found; each spec
  rewrites only its own ext.

### ODS wire mapping
`calcext:conditional-formats` is a child of `table:table` (re-emitted where
the source had it; new ones as the table's last child). Each
`calcext:conditional-format calcext:target-range-address` (space-separated
`Sheet1.A1:Sheet1.A10` addresses, read and written through
`odsRangeAddressToCanonical`/`canonicalToOdsRangeAddress` and
`parseRange`) holds its entries in priority order; priorities are assigned
by document order across the table.

| Model | `calcext:value` (on `calcext:condition`) |
| --- | --- |
| `cellIs` | `=f`, `!=f`, `>f`, `>=f`, `<f`, `<=f`, `between(f1,f2)`, `not-between(f1,f2)` |
| `expression` | `formula-is(f)` |
| `duplicates` | `duplicate` / `unique` |
| `top10` | `top-elements(n)`, `bottom-elements(n)`, `top-percent(n)`, `bottom-percent(n)` |
| `aboveAverage` | `above-average`, `below-average`, `above-equal-average`, `below-equal-average` (a `stdDev` rule → `formula-is`) |
| `text` | `begins-with("t")`, `ends-with("t")`, `contains-text("t")`, `not-contains-text("t")` |
| `errors` | `is-error` / `is-no-error` |
| `blanks` | `formula-is(LEN(TRIM(A1))=0)` (read back as `expression`) |
| `timePeriod` | `calcext:date-is calcext:date="today|yesterday|tomorrow|last-7-days|this-week|last-week|next-week|this-month|last-month|next-month|this-year|last-year|next-year"` + `calcext:style` |
| `colorScale` | `calcext:color-scale` › `calcext:color-scale-entry @calcext:value @calcext:type @calcext:color` |
| `dataBar` | `calcext:data-bar @calcext:positive-color @calcext:negative-color @calcext:axis-color @calcext:gradient @calcext:axis-position @calcext:show-value @calcext:min-length @calcext:max-length` › 2 `calcext:formatting-entry` |
| `iconSet` | `calcext:icon-set @calcext:icon-set-type @calcext:show-value` › n `calcext:formatting-entry` (no `gte` attribute = `gte`) |

Entry types: `minimum`↔`min`, `maximum`↔`max`, `auto-minimum`↔`autoMin`,
`auto-maximum`↔`autoMax`, `number`↔`num`, `percent`, `percentile`,
`formula`. Operand formulas are OpenFormula as in `table:formula` minus the
`of:=` prefix — references bracketed (`[.B1]`, `[$Sheet2.$A$1]`), as
SPEC-124's retained-reference rule also assumes — read with
`fromOdsFormula` and written with `toOdsFormula`; each is rebased from
`calcext:base-cell-address` to the origin with `shiftFormula`, and written
with the origin as `calcext:base-cell-address`. Legacy
`style:map style:condition` forms read: `cell-content()<op>v`,
`cell-content-is-between(a,b)`, `cell-content-is-not-between(a,b)`,
`is-true-formula(f)`, with `style:apply-style-name` and
`style:base-cell-address`.

### Structural edits
`adjustAnchoredFormula(text, origin, firstRange, hostSheetId, edit,
sheetNameOf): string` joins `formula/adjust.ts` (this spec adds it; SPEC-130
states the identical contract): when the edit deletes the origin's row or
column, the formula is first rebased with `shiftFormula` onto the first
surviving cell of the first surviving range (in pre-edit coordinates), then
`adjustFormula` applies. The adjuster maps every range with `adjustRange`,
re-anchors `cellIs`/`expression` formulas, runs `adjustFormula` over
value-object formulas, and over every `<formula>`/`<xm:f>` text of a
`preserved` rule (through `fromXlsxFormula`/`toXlsxFormula`) — the retained
rows it retires no longer do that; it replaces what it changes and drops
`sourceXml`. Duplicating a sheet (SPEC-121) copies its formats with fresh
ids; renaming or deleting a sheet runs `remapSheetReferences` over every
rule formula (a deleted sheet's references become `#REF!`).

### Moves
The `conditionalFormats` move adjuster (cut-paste and drag-move — formats travel
with their cells, as in Excel), for a `RangeMove` with offset `(dr, dc) =
(toTop − from.top, toLeft − from.left)`:
1. Every format on `toSheetId` loses the destination rectangle
   (`subtractRanges`) — the overwritten cells' rules go.
2. For every format on `fromSheetId`, the part of its ranges inside
   `move.from` is cut out. A format wholly inside moves: ranges offset onto
   `toSheetId`, same id. Otherwise the cut-out pieces become a new format
   on `toSheetId` (fresh ids), whose formulas are the source's first rebased
   onto the pieces' top-left with `shiftFormula` (so each piece means what
   it meant in place) — and a source left without its first range's
   top-left is re-anchored the same way onto its new origin. A new format's
   rules take the priorities right after its source's on the same sheet,
   or the lowest on another sheet; priorities are then renumbered.
3. Every formula of every format on every sheet then goes through
   `repointFormula`, exactly as cell formulas do — references to moved
   cells follow them, nothing else shifts. Across sheets, a moved
   format's unqualified references are first qualified with the source
   sheet's name, so they keep meaning the same cells.
Touched formats are replaced (never mutated) and lose `sourceXml`.

### The side panel
This spec creates the spreadsheet view's one side panel, and **every later
spreadsheet pane registers in it** rather than mounting a panel of its own
(SPEC-135's comments, SPEC-138's version history, SPEC-139's pivot field
list). `src/views/spreadsheet/useSheetSidePanel.ts`:
```ts
export type SheetPaneId = "conditionalFormats";   // later specs widen the union
export const SHEET_PANES: Record<SheetPaneId, { title: string; icon: string }>;  // + one entry per widening
export interface SheetSidePanel {
  activePane: Ref<SheetPaneId | null>;            // null = closed
  open(id: SheetPaneId): void;                    // shows that pane (a tab per pane opened this session)
  close(): void;
  toggle(id: SheetPaneId): void;                  // open, or close when it is the active pane
  side: Ref<"left" | "right">; widthPx: Ref<number>; collapsed: Ref<boolean>;
  tabs: ComputedRef<{ id: SheetPaneId; title: string; icon: string; badge?: string }[]>;  // SidePanel.vue's `tabs`
}
export function useSheetSidePanel(): SheetSidePanel;
```
State (active pane, side, width, collapse) persists under
`docu-edit:sheet-side-panel` with `src/lib/sidePanel.ts`'s width bounds;
`SpreadsheetView.vue` renders `SidePanel.vue` from it, the active pane's
component in its default slot. `useSidePanel` is not reused: it is wired to
the text editor's `EditorContext` and its closed pane list, the reason
SPEC-085 gave the deck a sibling panel. This spec's pane is
`src/components/spreadsheet/ConditionalFormatsPane.vue`; its state is
`src/views/spreadsheet/useConditionalFormatManager.ts` (≤ 400 lines);
`AppIcon` gains a `conditional-format` icon; `SpreadsheetView.size.spec.ts`'s
budget does not rise.

### Rule descriptions
"Cell value > 100" (operators `=`, `≠`, `>`, `≥`, `<`, `≤`; "Cell value
between 1 and 10", "… not between …"), "Formula: =$C2>AVERAGE($C$2:$C$21)",
"Graded colour scale", "Data bar", "Icon set (3 Arrows)", "Top 10", "Bottom
10%", "Above average" (+ " (1 std dev)"), "Below average", "Duplicate
values", "Unique values", "Cell value contains "abc"" (and "does not
contain", "begins with", "ends with"), "Blanks", "No blanks", "Errors", "No
errors", "Date occurring: Last 7 days", "Rule kept from the file" for
`preserved`.

### The rule editor
`cf-rule-type`: "Format all cells based on their values" (colour scale,
data bar, icon set: `cf-style` select, per-stop `cf-stop-type-<i>`,
`cf-stop-value-<i>`, `cf-stop-color-<i>`; bar `cf-bar-color`,
`cf-bar-gradient`, `cf-bar-negative-color`, `cf-bar-axis`; icons
`cf-icon-set`, `cf-icon-threshold-<i>`, `cf-icon-reverse`, `cf-show-value`),
"Format only cells that contain" (`cf-contain-kind`: Cell value, Specific
text, Dates occurring, Blanks, No blanks, Errors, No errors; then
`cf-operator`, `cf-value-1`, `cf-value-2`), "Format only top or bottom ranked
values", "Format only values that are above or below average", "Format only
unique or duplicate values", "Use a formula to determine which cells to
format" (`cf-value-1` takes the formula, relative to the active cell, stored
relative to the origin).

### Presets
Formats: "Light red fill with dark red text" (fill `#FFC7CE`, text
`#9C0006`), "Yellow fill with dark yellow text" (`#FFEB9C`, `#9C5700`),
"Green fill with dark green text" (`#C6EFCE`, `#006100`), "Light red fill",
"Red text", "Red border" (thin `#9C0006` on four edges). Data bars: blue
`#638EC6`, green `#63BE7B`, red `#FF555A`, orange `#FFB628`, light blue
`#008AEF`, purple `#D6007B`, gradient and solid. Colour scales (min,
50th percentile, max): Green–Yellow–Red (`#63BE7B`, `#FFEB84`, `#F8696B`),
Red–Yellow–Green, Green–White–Red (`#FCFCFF` middle), Red–White–Green,
White–Red and Green–White (two stops). Testids `cf-preset-greater-than`,
`-less-than`, `-between`, `-equal`, `-text-contains`, `-date`,
`-duplicates`, `-top10`, `-bottom10`, `-above-average`, `-below-average`,
`-data-bar-<colour>-<gradient|solid>`, `-color-scale-<n>`,
`-icon-set-<name>`, `-new-rule`.

### Fixtures
`spec129-cf.xlsx`, sheet "Scores" (Excel-shaped: `mc:Ignorable="x14ac xr
xr2 xr3"`, `xr:uid` on blocks, three `<dxfs>` with bgColor-only fills):
`B2:B13` `cellIs greaterThan 100` (dxf 0; `B3` = 140, `B4` = 90); `A2:H13`
expression `$H2="late"` with `stopIfTrue` (rows 3 and 4 are not late); `C2:C13` a data bar with an x14 twin (GUID,
negative values, `gradient="0"`, axis automatic); `D2:D13` `3Arrows` (`D2` the column's largest value) with
`gte="0"` thresholds; `E2:E13` a 3-colour scale with theme colours and tints;
`F2:F13` an x14-only `3Stars`; `G2:G13` a custom-icon x14 rule; plus
`top10 percent`, `aboveAverage stdDev="1"`, `duplicateValues`,
`containsText` with Excel's formula, `containsBlanks` and `timePeriod
last7Days` on sheet "More"; and a sparkline ext in the same `<extLst>`.
`spec129-cf.ods` carries the LibreOffice equivalents with named styles
"Bad", "Good" and "Accent 3" (encoded `Accent_20_3`) in `styles.xml`;
`spec129-legacy.ods` uses only `style:map` on automatic styles over a
repeated-cell run.

### Tests
Per the Tests AC. Time-dependent cases use `setClock`; colour assertions
compare resolved `#rrggbb` strings; the round-trip suite asserts
`equalBytes` per part for untouched fixtures and `toEqual` on the model
after edit → save → reopen.

### Specs this changes
- SPEC-117 — `SheetCellLayer.vue` applies the CF layer (dxf before
  `cellCss`, scale fill, bar, icon) per mounted cell.
- SPEC-120 — conditional formats evaluate through `evaluate`/`EvalContext`
  with their own `host`/`origin`; their volatility follows
  `FunctionSpec.volatile`.
- SPEC-121 — `SHEET_BY_REFERENCE_FIELDS` gains `conditionalFormats`,
  `WORKBOOK_SNAPSHOT_FIELDS` gains `dxfs`; duplicate/delete/rename of sheets
  cover rule formulas; `stripRetention` drops `sourceXml`, `x14Id`,
  `sourceStyleName`; `UNCONVERTIBLE_CONTENT` per the round-trip AC.
- SPEC-122 — `conditionalFormatting` joins the rebuilt set; `<dxfs>` is
  append-only like `cellXfs`; the ODS table writer emits
  `calcext:conditional-formats` and drops `style:map` from re-emitted cell
  styles.
- SPEC-123 — `FormatCellsDialog.vue` gains a differential mode (font colour
  and style only — no family or size — plus border, fill, number); the
  `SpreadsheetMenuBar` Format menu gains "Conditional formatting…".
- SPEC-125 — `MOVE_ADJUSTERS` gains `conditionalFormats`.
- SPEC-116/SPEC-117 — `address.ts` gains `subtractRanges`; the view gains
  the sheet side panel (`useSheetSidePanel`) later panes join.
- SPEC-124 — `formula/adjust.ts` gains `adjustAnchoredFormula`.
- SPEC-124 — the conditional-format rows of "Retained references" are
  retired (its e2e still holds: the modelled path writes the same shifted
  `sqref` and `calcext:target-range-address`).
- SPEC-128 — colour filters and colour sorts stay `preserved`, though `dxfs`
  now make evaluating them possible (Backlog).
- SPEC-133 (later) — its preserved workbook-defined table styles keep
  referencing `model.dxfs` indices, which append-only interning keeps valid.
