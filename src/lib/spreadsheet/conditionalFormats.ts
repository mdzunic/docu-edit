// Conditional formatting (SPEC-129) — THIS FILE is the normative source for
// the rule model (`ConditionalFormat`, `CfRule` and its kinds, the
// sheet-global priority and `stopIfTrue`), the differential style (`Dxf`)
// and how it layers over a cell's style (`layerDxf`), the per-cell result the
// renderers draw (`CfCellResult`), and the rule commands.
//
// The model:
// - `Sheet.conditionalFormats` holds a sheet's formats in source order (new
//   ones appended — commands never reorder the array; precedence is
//   `priority`). A format is a list of ranges and its rules; its ORIGIN is
//   `ranges[0]`'s top-left, which every `cellIs` and `expression` formula is
//   stored relative to (the engine's own `part + (host − origin)` rule
//   shifts it per cell, so no text is re-shifted). Value-object formulas
//   (`CfValue.value` starting a formula) are absolute and evaluate at the
//   origin.
// - `priority` is sheet-global: 1 is the highest, every rule's is unique,
//   and the numbers are dense — renumbered (renumberPriorities) after every
//   command, and at read (keeping each block's `sourceXml`: the writer
//   re-emits an unedited block as the file wrote it).
// - `SpreadsheetModel.dxfs` is the workbook's differential-style table —
//   index = the xlsx `dxfId` — and is append-only (styles.ts's internDxf), so
//   every dxfId an untouched sheet (or SPEC-133's preserved table styles)
//   references stays valid.
// - A `preserved` rule is one this product cannot model (an xlsx custom icon
//   set, an unknown type …): its XML verbatim, never evaluated, editable only
//   by deletion; its format's ranges stay modelled.
//
// Commands (record → command → commit through SPEC-121's undo stack, one
// entry labelled CONDITIONAL_FORMATTING_LABEL): addConditionalFormat,
// updateRule, deleteRule, moveRule, setFormatRanges and clearRulesInRanges.
// Each replaces what it touches — the sheet, its `conditionalFormats` array,
// the formats and rules it changes (copy-on-write: `conditionalFormats` is a
// SHEET_BY_REFERENCE_FIELDS entry, shared by reference with undo snapshots)
// — drops the `sourceXml` of every format it changes (one whose rules were
// renumbered included), renumbers the priorities densely and sets
// `Sheet.dirty`. A format left with no rule or no range is removed.
//
// Layering rule (boundary.spec.ts): framework-free, imports only spreadsheet
// siblings and src/lib/core.
import type { ModelColor } from "../core/theme";
import { subtractRanges, type CellRef, type GridRange } from "./address";
import { tokenize } from "./formula/tokens";
import { shiftFormula } from "./formula/references";
import type { CellBorderEdge, CellFill, CellFont, CellStyle, Sheet, SpreadsheetModel } from "./model";

// --- the rule model ----------------------------------------------------------------------------

export interface ConditionalFormat {
  id: string; // minted at read; never written
  ranges: GridRange[]; // sqref / calcext:target-range-address, source order
  rules: CfRule[];
  pivot?: true; // xlsx @pivot, preserved
  sourceXml?: string; // verbatim source element (+ preceding whitespace) while unchanged
}

export interface CfRuleBase {
  id: string;
  priority: number; // sheet-global; 1 = highest; unique and dense
  stopIfTrue?: true;
  dxfId?: number; // into model.dxfs, for the formatting kinds
  x14?: true; // written into the x14 extLst (xlsx)
  x14Id?: string; // "{GUID}" pairing a base rule with its x14 extension
}

export type CfOperator = "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";

export type CfTimePeriod =
  | "today"
  | "yesterday"
  | "tomorrow"
  | "last7Days"
  | "thisWeek"
  | "lastWeek"
  | "nextWeek"
  | "thisMonth"
  | "lastMonth"
  | "nextMonth"
  | "thisYear"
  | "lastYear"
  | "nextYear";

export interface CfValue {
  type: "min" | "max" | "autoMin" | "autoMax" | "num" | "percent" | "percentile" | "formula";
  value?: string; // number text or canonical formula (absolute, evaluated at the origin)
}

export interface CfDataBar {
  min: CfValue;
  max: CfValue;
  color: ModelColor;
  gradient: boolean;
  border?: ModelColor;
  negativeColor?: ModelColor;
  negativeBorderColor?: ModelColor;
  axis: "automatic" | "middle" | "none";
  axisColor?: ModelColor;
  direction: "context" | "leftToRight" | "rightToLeft";
  minLengthPct: number;
  maxLengthPct: number;
  showValue: boolean;
}

export interface CfIconSet {
  set: IconSetName;
  thresholds: { at: CfValue; gte: boolean }[]; // one per icon, lowest first
  reverse: boolean;
  showValue: boolean;
}

export type CfRuleBody =
  | { kind: "cellIs"; operator: CfOperator; formulas: string[] } // 1, or 2 for (not)between
  | { kind: "expression"; formula: string }
  | { kind: "colorScale"; stops: { at: CfValue; color: ModelColor }[] } // 2 or 3
  | { kind: "dataBar"; bar: CfDataBar }
  | { kind: "iconSet"; icons: CfIconSet }
  | { kind: "top10"; bottom: boolean; percent: boolean; rank: number }
  | { kind: "aboveAverage"; below: boolean; equalAverage: boolean; stdDev?: 1 | 2 | 3 }
  | { kind: "duplicates"; unique: boolean }
  | { kind: "text"; operator: "contains" | "notContains" | "beginsWith" | "endsWith"; text: string }
  | { kind: "blanks"; not: boolean }
  | { kind: "errors"; not: boolean }
  | { kind: "timePeriod"; period: CfTimePeriod }
  | { kind: "preserved"; xml: string; where: "base" | "x14" | "ods" };

export type CfRule = CfRuleBase & CfRuleBody;

export type CfRuleKind = CfRule["kind"];

/** A rule as an editor builds it: everything but the minted `id` and the
 * `priority` the commands assign. */
export type CfRuleDraft = Omit<CfRuleBase, "id" | "priority"> & CfRuleBody;

export type IconSetName =
  | "3Arrows"
  | "3ArrowsGray"
  | "3Flags"
  | "3TrafficLights1"
  | "3TrafficLights2"
  | "3Signs"
  | "3Symbols"
  | "3Symbols2"
  | "3Stars"
  | "3Triangles"
  | "3Smilies"
  | "3ColorSmilies"
  | "4Arrows"
  | "4ArrowsGray"
  | "4RedToBlack"
  | "4Rating"
  | "4TrafficLights"
  | "5Arrows"
  | "5ArrowsGray"
  | "5Rating"
  | "5Quarters"
  | "5Boxes"
  | "NoIcons";

/** Every icon set name, Excel's sets first in its gallery order, then the
 * x14 sets and LibreOffice's smilies. */
export const ICON_SET_NAMES: readonly IconSetName[] = [
  "3Arrows",
  "3ArrowsGray",
  "3Flags",
  "3TrafficLights1",
  "3TrafficLights2",
  "3Signs",
  "3Symbols",
  "3Symbols2",
  "4Arrows",
  "4ArrowsGray",
  "4RedToBlack",
  "4Rating",
  "4TrafficLights",
  "5Arrows",
  "5ArrowsGray",
  "5Rating",
  "5Quarters",
  "3Stars",
  "3Triangles",
  "5Boxes",
  "NoIcons",
  "3Smilies",
  "3ColorSmilies",
];

/** The sets with no Excel 2007 form: xlsx writes them to the x14 ext. */
export const X14_ICON_SETS: ReadonlySet<IconSetName> = new Set(["3Stars", "3Triangles", "5Boxes", "NoIcons"]);

/** The sets only LibreOffice writes (no xlsx form at all). */
export const ODS_ONLY_ICON_SETS: ReadonlySet<IconSetName> = new Set(["3Smilies", "3ColorSmilies"]);

/** How many thresholds (icons) a set has — `NoIcons` is a 3-threshold set
 * drawing nothing. */
export function iconSetSize(set: IconSetName): 3 | 4 | 5 {
  return set.startsWith("5") ? 5 : set.startsWith("4") ? 4 : 3;
}

/** The default thresholds of an n-icon set: `percent` 0/33/67, 0/25/50/75
 * or 0/20/40/60/80, all `gte`. */
export function defaultIconThresholds(count: 3 | 4 | 5): { at: CfValue; gte: boolean }[] {
  const steps = count === 3 ? [0, 33, 67] : count === 4 ? [0, 25, 50, 75] : [0, 20, 40, 60, 80];
  return steps.map((step) => ({ at: { type: "percent", value: String(step) }, gte: true }));
}

/** The kinds that format through a dxf (everything but the colour scale,
 * the data bar, the icon set and a preserved rule). */
export function isFormattingRule(rule: Pick<CfRule, "kind">): boolean {
  return rule.kind !== "colorScale" && rule.kind !== "dataBar" && rule.kind !== "iconSet" && rule.kind !== "preserved";
}

/** The undo entry every rule command records. */
export const CONDITIONAL_FORMATTING_LABEL = "Conditional formatting";

// --- differential styles ------------------------------------------------------------------------

export interface DxfFont {
  color?: ModelColor;
  bold?: boolean;
  italic?: boolean;
  underline?: CellFont["underline"] | "none";
  strike?: boolean;
}

export interface Dxf {
  font?: DxfFont;
  fill?: CellFill;
  borders?: { top?: CellBorderEdge; right?: CellBorderEdge; bottom?: CellBorderEdge; left?: CellBorderEdge };
  numberFormat?: string;
  sourceXml?: string; // xlsx <dxf> as read — keeps alignment/protection children
  sourceStyleName?: string; // ods named style it came from
}

/** The border edges a dxf may set, in the order the merge walks them. */
export const DXF_BORDER_EDGES = ["top", "right", "bottom", "left"] as const;

/** `base` with exactly the properties `dxf` sets replaced (SPEC-129
 * "Differential styles"): a dxf `bold: false` un-bolds a bold cell, an
 * `underline: "none"` removes one; a fill, each border edge and the number
 * format replace the base's. Pure — `base` is never mutated. */
export function layerDxf(base: CellStyle, dxf: Dxf): CellStyle {
  const style: CellStyle = { ...base };
  const font = dxf.font;
  if (font && Object.values(font).some((value) => value !== undefined)) {
    const next: CellFont = { ...base.font };
    if (font.color !== undefined) {
      next.color = font.color;
    }
    setFlag(next, "bold", font.bold);
    setFlag(next, "italic", font.italic);
    setFlag(next, "strike", font.strike);
    if (font.underline === "none") {
      delete next.underline;
    } else if (font.underline !== undefined) {
      next.underline = font.underline;
    }
    style.font = next;
  }
  if (dxf.fill !== undefined) {
    style.fill = dxf.fill;
  }
  if (dxf.borders && DXF_BORDER_EDGES.some((edge) => dxf.borders![edge] !== undefined)) {
    const borders = { ...base.borders };
    for (const edge of DXF_BORDER_EDGES) {
      const value = dxf.borders[edge];
      if (value !== undefined) {
        borders[edge] = value;
      }
    }
    style.borders = borders;
  }
  if (dxf.numberFormat !== undefined) {
    style.numberFormat = dxf.numberFormat;
    delete style.builtinFormatId;
  }
  return style;
}

function setFlag(font: CellFont, key: "bold" | "italic" | "strike", value: boolean | undefined): void {
  if (value === true) {
    font[key] = true;
  } else if (value === false) {
    delete font[key];
  }
}

/** A dxf without its source fields — what interning and comparison look at. */
export function modelledDxf(dxf: Dxf): Dxf {
  const { sourceXml: _xml, sourceStyleName: _name, ...rest } = dxf;
  return rest;
}

// --- the per-cell result ------------------------------------------------------------------------

export interface CfCellResult {
  dxf?: Dxf; // merged
  scaleColor?: string; // "#rrggbb", when a colour scale won the fill
  dataBar?: {
    startPct: number;
    widthPct: number;
    negative: boolean;
    axisPct?: number;
    color: string;
    endColor?: string;
    borderColor?: string;
    axisColor?: string;
  };
  icon?: { set: IconSetName; index: number };
  hideValue?: true; // the winning bar or icon has showValue false
}

// --- helpers ------------------------------------------------------------------------------------

/** A fresh id for a format or a rule (never written to a file). */
export function mintCfId(): string {
  return crypto.randomUUID();
}

/** A format's origin: its first range's top-left. */
export function formatOrigin(format: Pick<ConditionalFormat, "ranges">): CellRef {
  const first = format.ranges[0];
  return first ? { row: first.top, col: first.left } : { row: 0, col: 0 };
}

/** Whether any of `ranges` contains (row, col). */
export function rangesContain(ranges: readonly GridRange[], row: number, col: number): boolean {
  return ranges.some((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
}

/** The bounding box of `ranges`, or null for none. */
export function boundingBox(ranges: readonly GridRange[]): GridRange | null {
  if (ranges.length === 0) {
    return null;
  }
  let { top, left, bottom, right } = ranges[0]!;
  for (const range of ranges) {
    top = Math.min(top, range.top);
    left = Math.min(left, range.left);
    bottom = Math.max(bottom, range.bottom);
    right = Math.max(right, range.right);
  }
  return { top, left, bottom, right };
}

/** A sheet's formats, `[]` for a sheet stored before SPEC-129 that no reader
 * upgrade has reached (a re-stamped csv record). */
export function sheetFormats(sheet: Pick<Sheet, "conditionalFormats">): ConditionalFormat[] {
  return sheet.conditionalFormats ?? [];
}

/** Every rule of `formats` with its format, by ascending priority (ties —
 * never present after a renumber — in array order). */
export function rulesByPriority(formats: readonly ConditionalFormat[]): { format: ConditionalFormat; rule: CfRule }[] {
  const all: { format: ConditionalFormat; rule: CfRule; order: number }[] = [];
  for (const format of formats) {
    for (const rule of format.rules) {
      all.push({ format, rule, order: all.length });
    }
  }
  all.sort((a, b) => a.rule.priority - b.rule.priority || a.order - b.order);
  return all.map(({ format, rule }) => ({ format, rule }));
}

/** The formats with priorities renumbered 1..n in their current order
 * (ascending priority, ties in array order) and each format's rules sorted
 * by priority — the same array when nothing changes. A format whose rules
 * change is replaced, without its `sourceXml` unless `keepSource` (the
 * readers', whose blocks are still the file's). */
export function renumberPriorities(formats: readonly ConditionalFormat[], options: { keepSource?: boolean } = {}): ConditionalFormat[] {
  const next = new Map<string, number>();
  rulesByPriority(formats).forEach(({ rule }, index) => next.set(rule.id, index + 1));
  let out: ConditionalFormat[] | null = null;
  formats.forEach((format, at) => {
    const rules = format.rules.map((rule) => (rule.priority === next.get(rule.id) ? rule : { ...rule, priority: next.get(rule.id)! }));
    rules.sort((a, b) => a.priority - b.priority);
    if (rules.every((rule, index) => rule === format.rules[index])) {
      return;
    }
    out ??= formats.slice();
    out[at] = options.keepSource ? { ...format, rules } : withoutSource({ ...format, rules });
  });
  return out ?? (formats as ConditionalFormat[]);
}

function withoutSource(format: ConditionalFormat): ConditionalFormat {
  const { sourceXml: _source, ...rest } = format;
  return rest;
}

/** Whether a formula references a sheet other than `hostSheetName` (an
 * unqualified reference is the host's). */
export function referencesOtherSheet(formula: string, hostSheetName: string): boolean {
  const host = hostSheetName.toUpperCase();
  return tokenize(formula).some((token) => {
    const sheet = token.ref?.sheet;
    return sheet !== undefined && (sheet.toUpperCase() !== host || token.ref?.sheetEnd !== undefined);
  });
}

/** Whether an xlsx writer puts `rule` in the x14 ext (SPEC-129 "x14"): an
 * icon set with no 2007 form, a `cellIs`/`expression` rule whose formulas
 * reference another sheet, or a rule read from there that is still
 * preserved there. (A data bar is written to both: base plus x14 twin.) */
export function writesToX14(rule: CfRule, hostSheetName: string): boolean {
  switch (rule.kind) {
    case "iconSet":
      return X14_ICON_SETS.has(rule.icons.set);
    case "cellIs":
      return rule.formulas.some((formula) => referencesOtherSheet(formula, hostSheetName));
    case "expression":
      return referencesOtherSheet(rule.formula, hostSheetName);
    case "preserved":
      return rule.where === "x14";
    default:
      return false;
  }
}

/** `rule` with its `x14` flag set as `format`'s writer will place it (only
 * xlsx workbooks have an x14 ext), so a save → reopen reads the same flag. */
export function withX14Placement<R extends CfRule>(rule: R, format: SpreadsheetModel["format"], hostSheetName: string): R {
  const x14 = format === "xlsx" && rule.kind !== "preserved" && writesToX14(rule, hostSheetName);
  if (x14 === (rule.x14 === true)) {
    return rule;
  }
  if (x14) {
    return { ...rule, x14: true };
  }
  const { x14: _x14, ...rest } = rule;
  return rest as R;
}

/** `rule` re-anchored from origin `from` to origin `to`: its relative
 * `cellIs`/`expression` formulas shifted by (to − from), so every covered
 * cell keeps its meaning. Value-object formulas are absolute and stay. */
export function reanchorRule(rule: CfRule, from: CellRef, to: CellRef): CfRule {
  const dRow = to.row - from.row;
  const dCol = to.col - from.col;
  if (dRow === 0 && dCol === 0) {
    return rule;
  }
  if (rule.kind === "cellIs") {
    return { ...rule, formulas: rule.formulas.map((formula) => shiftFormula(formula, dRow, dCol)) };
  }
  if (rule.kind === "expression") {
    return { ...rule, formula: shiftFormula(rule.formula, dRow, dCol) };
  }
  return rule;
}

// --- commands ------------------------------------------------------------------------------------

function sheetById(model: SpreadsheetModel, sheetId: string): Sheet {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet) {
    throw new Error(`No sheet ${sheetId}`);
  }
  return sheet;
}

/** Installs `formats` (renumbered) on the sheet, replacing it, marked dirty. */
function commitFormats(model: SpreadsheetModel, sheetId: string, formats: ConditionalFormat[]): void {
  const renumbered = renumberPriorities(formats);
  model.sheets = model.sheets.map((sheet) => (sheet.id === sheetId ? { ...sheet, conditionalFormats: renumbered, dirty: true } : sheet));
}

/** A rule built from a draft: minted id, the given priority, `x14` set per
 * the model's format. */
function ruleFromDraft(model: SpreadsheetModel, sheet: Sheet, draft: CfRuleDraft, id: string, priority: number): CfRule {
  return withX14Placement({ ...structuredClone(draft), id, priority } as CfRule, model.format, sheet.name);
}

/** Adds a format over `ranges` holding one rule built from `draft`, at
 * priority 1 (Excel puts a new rule first). Returns the new ids. */
export function addConditionalFormat(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], draft: CfRuleDraft): { formatId: string; ruleId: string } {
  const sheet = sheetById(model, sheetId);
  if (ranges.length === 0) {
    throw new Error("A conditional format needs a range");
  }
  const rule = ruleFromDraft(model, sheet, draft, mintCfId(), 0);
  const format: ConditionalFormat = { id: mintCfId(), ranges: ranges.map((range) => ({ ...range })), rules: [rule] };
  commitFormats(model, sheetId, [...sheetFormats(sheet), format]);
  return { formatId: format.id, ruleId: rule.id };
}

/** The format holding rule `ruleId` and its index in the sheet's formats. */
function locateRule(sheet: Sheet, ruleId: string): { format: ConditionalFormat; at: number; rule: CfRule } {
  const formats = sheetFormats(sheet);
  for (const [at, format] of formats.entries()) {
    const rule = format.rules.find((each) => each.id === ruleId);
    if (rule) {
      return { format, at, rule };
    }
  }
  throw new Error(`No conditional formatting rule ${ruleId}`);
}

/** Replaces rule `ruleId` with one built from `draft` — same id and
 * priority; a data bar keeps its x14 GUID while it stays a data bar. */
export function updateRule(model: SpreadsheetModel, sheetId: string, ruleId: string, draft: CfRuleDraft): void {
  const sheet = sheetById(model, sheetId);
  const { format, at, rule } = locateRule(sheet, ruleId);
  const { x14Id: _id, x14: _x14, ...fields } = draft as CfRuleDraft & { x14Id?: string };
  let next = ruleFromDraft(model, sheet, fields as CfRuleDraft, rule.id, rule.priority);
  if (next.kind === "dataBar" && rule.kind === "dataBar" && rule.x14Id !== undefined) {
    next = { ...next, x14Id: rule.x14Id };
  }
  const formats = sheetFormats(sheet).slice();
  formats[at] = withoutSource({ ...format, rules: format.rules.map((each) => (each.id === ruleId ? next : each)) });
  commitFormats(model, sheetId, formats);
}

/** Sets or clears a rule's `stopIfTrue` (the manager's per-rule checkbox). */
export function setStopIfTrue(model: SpreadsheetModel, sheetId: string, ruleId: string, stop: boolean): void {
  const sheet = sheetById(model, sheetId);
  const { format, at, rule } = locateRule(sheet, ruleId);
  if ((rule.stopIfTrue === true) === stop) {
    return;
  }
  const { stopIfTrue: _stop, ...rest } = rule;
  const next = (stop ? { ...rest, stopIfTrue: true } : rest) as CfRule;
  const formats = sheetFormats(sheet).slice();
  formats[at] = withoutSource({ ...format, rules: format.rules.map((each) => (each.id === ruleId ? next : each)) });
  commitFormats(model, sheetId, formats);
}

/** Deletes rule `ruleId`; a format left with no rule is removed. */
export function deleteRule(model: SpreadsheetModel, sheetId: string, ruleId: string): void {
  const sheet = sheetById(model, sheetId);
  const { format, at } = locateRule(sheet, ruleId);
  const rules = format.rules.filter((each) => each.id !== ruleId);
  const formats = sheetFormats(sheet).slice();
  if (rules.length === 0) {
    formats.splice(at, 1);
  } else {
    formats[at] = withoutSource({ ...format, rules });
  }
  commitFormats(model, sheetId, formats);
}

/** Moves rule `ruleId` one place up (towards priority 1) or down the
 * sheet's priority order, swapping it with its neighbour; at either end it
 * changes nothing. */
export function moveRule(model: SpreadsheetModel, sheetId: string, ruleId: string, direction: "up" | "down"): void {
  const sheet = sheetById(model, sheetId);
  const ordered = rulesByPriority(sheetFormats(sheet));
  const index = ordered.findIndex((entry) => entry.rule.id === ruleId);
  if (index < 0) {
    throw new Error(`No conditional formatting rule ${ruleId}`);
  }
  const other = direction === "up" ? index - 1 : index + 1;
  if (other < 0 || other >= ordered.length) {
    return;
  }
  const a = ordered[index]!.rule;
  const b = ordered[other]!.rule;
  const swapped = new Map<string, number>([
    [a.id, b.priority],
    [b.id, a.priority],
  ]);
  const formats = sheetFormats(sheet).map((format) =>
    format.rules.some((rule) => swapped.has(rule.id))
      ? withoutSource({ ...format, rules: format.rules.map((rule) => (swapped.has(rule.id) ? { ...rule, priority: swapped.get(rule.id)! } : rule)) })
      : format,
  );
  commitFormats(model, sheetId, formats);
}

/** Sets a format's ranges (the manager's "Applies to"). The stored formula
 * text is kept — relative to the new origin, as Excel keeps it. An empty
 * list is refused. */
export function setFormatRanges(model: SpreadsheetModel, sheetId: string, formatId: string, ranges: readonly GridRange[]): void {
  if (ranges.length === 0) {
    throw new Error("A conditional format needs a range");
  }
  const sheet = sheetById(model, sheetId);
  const formats = sheetFormats(sheet);
  const at = formats.findIndex((format) => format.id === formatId);
  if (at < 0) {
    throw new Error(`No conditional format ${formatId}`);
  }
  const next = formats.slice();
  next[at] = withoutSource({ ...formats[at]!, ranges: ranges.map((range) => ({ ...range })) });
  commitFormats(model, sheetId, next);
}

/** Removes `cut` from every format's ranges ("Clear rules from selected
 * cells"; the whole sheet clears every format). A format left with no range
 * is removed; one that lost its origin is re-anchored onto its new one.
 * Returns whether anything changed (nothing is recorded as dirty if not). */
export function clearRulesInRanges(model: SpreadsheetModel, sheetId: string, cut: readonly GridRange[]): boolean {
  const sheet = sheetById(model, sheetId);
  const formats = sheetFormats(sheet);
  let changed = false;
  const next: ConditionalFormat[] = [];
  for (const format of formats) {
    const ranges = subtractRanges(format.ranges, cut);
    if (sameRanges(ranges, format.ranges)) {
      next.push(format);
      continue;
    }
    changed = true;
    if (ranges.length === 0) {
      continue;
    }
    const from = formatOrigin(format);
    const to = formatOrigin({ ranges });
    next.push(withoutSource({ ...format, ranges, rules: format.rules.map((rule) => reanchorRule(rule, from, to)) }));
  }
  if (changed) {
    commitFormats(model, sheetId, next);
  }
  return changed;
}

function sameRanges(a: readonly GridRange[], b: readonly GridRange[]): boolean {
  return (
    a.length === b.length &&
    a.every((range, i) => range.top === b[i]!.top && range.left === b[i]!.left && range.bottom === b[i]!.bottom && range.right === b[i]!.right)
  );
}
