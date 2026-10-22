// The xlsx wire mapping of SPEC-129's conditional formats ("xlsx wire
// mapping" and "x14", normative there; conditionalFormats.ts is the model's
// source): what xlsx/sheetXml.ts's reader makes of a worksheet's
// `<conditionalFormatting>` blocks and the x14 conditional formattings in its
// `<extLst>`, and what its writer puts back.
//
// Reading (readConditionalFormatSource — the reader hands over the root start
// tag, each block with the whitespace before it and the root `<extLst>`):
// - Each `<conditionalFormatting sqref [pivot]>` is a format (its sqref
//   parsed as a list of ranges — whole columns/rows too; a block whose sqref
//   names nothing is not modelled) keeping the block verbatim as
//   `sourceXml`. Each `<cfRule>` is a rule of the kind its `@type` names per
//   the spec's table, formulas through fromXlsxFormula; text, blank, error
//   and time-period rules are modelled from their attributes and their
//   `<formula>` ignored (the writer regenerates it). An unknown type, a
//   child the kind does not carry, an operator, set name or value object the
//   model cannot hold → a `preserved` rule (`where: "base"`) holding the
//   cfRule verbatim. Base-only data bars read as Excel 2007 draws them:
//   gradient, no axis.
// - The worksheet extLst's `{78C0D931-…}` ext › `x14:conditionalFormattings`
//   › `x14:conditionalFormatting` (`xm:sqref`) › `x14:cfRule`. A base data
//   bar whose `<extLst>` `{B025F937-…}` ext holds `<x14:id>` pairs with the
//   x14 dataBar rule of that `@id`: its x14 settings overlay the base's
//   (minLength/maxLength, gradient, border + borderColor, direction,
//   negativeFillColor or negativeBarColorSameAsPositive, negativeBorderColor
//   when negativeBarBorderColorSameAsPositive="0", axisPosition, axisColor,
//   fillColor, x14:cfvo incl. autoMin/autoMax with `xm:f` values — each x14
//   default applying when absent) and the GUID is kept as `x14Id`. Every
//   other x14 rule is the rule of a format over its block's `xm:sqref`:
//   `3Stars`, `3Triangles`, `5Boxes` and `NoIcons` icon sets (not custom)
//   and `cellIs`/`expression` rules whose `xm:f` references another sheet
//   (their inline `x14:dxf` interned into `dxfs`) are modelled with `x14:
//   true` and their `@id` as `x14Id`; anything else — a custom icon set
//   (`custom="1"`, `x14:cfIcon`) among them — is `preserved` (`where:
//   "x14"`). Such a format keeps its `x14:conditionalFormatting` verbatim as
//   `sourceXml` unless the block also held a twin. Formats: the base blocks
//   in order, then the x14 ones; priorities (base and x14 one order, a
//   missing one last) renumbered keeping the sources.
//
// Writing (writeConditionalFormats; sheetXml.ts splices the result into a
// DIRTY sheet — a clean part is never touched — and leaves the run and the
// ext as they are while the model holds the source's formats unedited):
// - The base run, in the model's order: a format still holding its
//   `sourceXml` (and a retained source to fit into) is emitted verbatim —
//   through the writer's sheet-name pass, its source text naming sheets as
//   the source did; any other format is serialised with the part's prefix:
//   `<conditionalFormatting sqref [pivot="1"]>` and a `<cfRule type [dxfId]
//   (through the style writer, which appends the dxfs the part lacks)
//   priority [stopIfTrue] [aboveAverage] [percent] [bottom] [operator]
//   [text] [timePeriod] [rank] [stdDev] [equalAverage]>` per base rule, with
//   its children (`formula`, `colorScale`, `dataBar` + the `<extLst>` naming
//   its x14 twin, `iconSet`). The text, blank, error and time-period rules'
//   `<formula>` is Excel's for the kind with A1 = the origin (the
//   LibreOffice-only year periods become `expression` rules); `3Smilies`/
//   `3ColorSmilies` are written as LibreOffice's own export writes them.
// - Priorities: serialised rules carry the model's; a verbatim block keeps
//   its written ones while they still rank in the model's order, else ONLY
//   its `priority` attributes are patched (a preserved rule inside a
//   serialised block likewise).
// - The x14 ext's blocks, in the model's order: a verbatim x14 format's
//   `sourceXml`; else one `<x14:conditionalFormatting>` per format holding
//   the x14 twin of each data bar (a base rule plus an x14 rule sharing a
//   GUID — `x14Id`, else a fresh upper-case `{…}` from crypto.randomUUID();
//   `minLength`/`maxLength` the model's, `gradient="0"`, `border="1"` +
//   borderColor, direction, negativeBarColorSameAsPositive or a
//   negativeFillColor (FFFF0000 by default), a negativeBorderColor, the axis
//   position and axisColor (FF000000 by default)), the rules writesToX14
//   places there (`priority`, `id`, `xm:f` formulas, an inline `x14:dxf`) and
//   the preserved x14 rules verbatim, then `<xm:sqref>`.
// A model with no retained source (a new sheet, a stripRetention copy) is
// serialised whole; a rule preserved from an ODS file has no xlsx form.
import type { DocumentTheme, ModelColor } from "../../core/theme";
import { sameModelColor } from "../../core/theme";
import { formatA1, formatRange, lettersToCol, parseRange, type GridRange } from "../address";
import {
  formatOrigin,
  iconSetSize,
  ICON_SET_NAMES,
  mintCfId,
  referencesOtherSheet,
  renumberPriorities,
  writesToX14,
  X14_ICON_SETS,
  type CfDataBar,
  type CfOperator,
  type CfRule,
  type CfRuleBody,
  type CfTimePeriod,
  type CfValue,
  type ConditionalFormat,
  type Dxf,
  type IconSetName,
} from "../conditionalFormats";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import { MAX_COLS, MAX_ROWS, type SpreadsheetModel } from "../model";
import { internDxf } from "../styles";
import { applySplices, escapeAttributeValue, escapeElementText, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { childNodes, elementTree, nodeAttr, whitespaceBefore, type XmlNode } from "./elementTree";
import { dxfXml, isSpreadsheetMl, readDxfNode } from "./styles";
import { modelColorToXlsxAttrs, xlsxColorFrom } from "./themeColor";

/** The Excel 2010 SpreadsheetML namespace (`x14:`). */
export const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
/** The Excel formula namespace (`xm:f`, `xm:sqref`). */
export const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
/** The worksheet ext holding the x14 conditional formattings. */
export const CF_EXT_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
/** The cfRule ext naming a base rule's x14 twin. */
export const RULE_ID_EXT_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

/** What the reader cuts out of a worksheet part for this module. */
export interface WorksheetCfSource {
  /** The worksheet's root start tag — the namespaces in scope. */
  rootTag: string;
  /** Each root-level `<conditionalFormatting>`, the whitespace before it included. */
  blocks: string[];
  /** The root-level `<extLst>`, when the part has one. */
  extLst?: string;
}

const inMain = (ns: string | undefined): boolean => isSpreadsheetMl(ns);
const inX14 = (ns: string | undefined): boolean => ns === X14_NS;
const inXm = (ns: string | undefined): boolean => ns === XM_NS;

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function isOff(value: string | undefined): boolean {
  return value === "0" || value === "false";
}

function parseNumber(text: string | undefined): number | undefined {
  if (text === undefined || text.trim() === "") {
    return undefined;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

function parseIndex(text: string | undefined): number | undefined {
  const value = parseNumber(text);
  return value !== undefined && Number.isInteger(value) && value >= 0 ? value : undefined;
}

/** A rule without a usable `priority` sorts after every one that has one. */
const NO_PRIORITY = Number.MAX_SAFE_INTEGER;

const NUMBER = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;

/** An sqref (or `xm:sqref`): its ranges — A1 ranges, whole columns `A:C`
 * and whole rows `2:5` — in order; anything else is dropped. */
export function parseCfSqref(text: string): GridRange[] {
  const ranges: GridRange[] = [];
  for (const entry of text.trim().split(/\s+/)) {
    if (entry === "") {
      continue;
    }
    const range = parseRange(entry);
    if (range) {
      ranges.push(range);
      continue;
    }
    const columns = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(entry);
    const rows = /^\$?(\d+):\$?(\d+)$/.exec(entry);
    if (columns) {
      const a = lettersToCol(columns[1]!);
      const b = lettersToCol(columns[2]!);
      if (a !== null && b !== null && a < MAX_COLS && b < MAX_COLS) {
        ranges.push({ top: 0, bottom: MAX_ROWS - 1, left: Math.min(a, b), right: Math.max(a, b) });
      }
    } else if (rows) {
      const a = Number(rows[1]) - 1;
      const b = Number(rows[2]) - 1;
      if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) {
        ranges.push({ top: Math.min(a, b), bottom: Math.max(a, b), left: 0, right: MAX_COLS - 1 });
      }
    }
  }
  return ranges;
}

const sqrefText = (ranges: readonly GridRange[]): string => ranges.map(formatRange).join(" ");

// --- reading ------------------------------------------------------------------------------

const OPERATORS: ReadonlySet<string> = new Set<CfOperator>([
  "equal",
  "notEqual",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "between",
  "notBetween",
]);

const TIME_PERIODS: ReadonlySet<string> = new Set<CfTimePeriod>([
  "today",
  "yesterday",
  "tomorrow",
  "last7Days",
  "thisWeek",
  "lastWeek",
  "nextWeek",
  "thisMonth",
  "lastMonth",
  "nextMonth",
  "thisYear",
  "lastYear",
  "nextYear",
]);

const VALUE_TYPES: ReadonlySet<string> = new Set<CfValue["type"]>(["min", "max", "autoMin", "autoMax", "num", "percent", "percentile", "formula"]);

const ICON_SETS: ReadonlySet<string> = new Set<string>(ICON_SET_NAMES);

/** What one read needs at hand. */
interface ReadContext {
  sheetName: string;
  holder: Pick<SpreadsheetModel, "dxfs">;
  indexedColors: readonly string[] | undefined;
  /** Names each format and rule (mintCfId; SPEC-137's worksheet scan passes
   * placeholders the main thread re-mints, so the scan never needs crypto). */
  mint: () => string;
}

function colorOf(node: XmlNode | undefined, ctx: ReadContext): ModelColor | undefined {
  return node ? xlsxColorFrom((name) => nodeAttr(node, name), ctx.indexedColors) : undefined;
}

/** A value object: `type`, and for the valued types the number text or the
 * formula (canonical) `wire` holds; undefined when the model cannot hold it. */
function readValue(node: XmlNode, wire: string | undefined): CfValue | undefined {
  const type = nodeAttr(node, "type");
  if (type === undefined || !VALUE_TYPES.has(type)) {
    return undefined;
  }
  const value: CfValue = { type: type as CfValue["type"] };
  if (type === "min" || type === "max" || type === "autoMin" || type === "autoMax") {
    return value;
  }
  const text = wire?.trim();
  if (text === undefined || text === "") {
    return undefined;
  }
  value.value = type !== "formula" && NUMBER.test(text) ? text : fromXlsxFormula(text);
  return value;
}

/** A base `<cfvo type val>`. */
const baseValue = (node: XmlNode): CfValue | undefined => readValue(node, nodeAttr(node, "val"));

/** An `<x14:cfvo type>` with its `<xm:f>` value. */
const x14Value = (node: XmlNode): CfValue | undefined => readValue(node, childNodes(node, "f", inXm)[0]?.text);

/** Whether every child of `node` is a main-namespace element named in `allowed`. */
function onlyChildren(node: XmlNode, allowed: readonly string[], inNs: (ns: string | undefined) => boolean = inMain): boolean {
  return node.children.every((child) => inNs(child.ns) && allowed.includes(child.local));
}

/** A base data bar's twin GUID from its `<extLst>` (undefined: no extLst),
 * or null when the extLst holds anything else. */
function twinGuid(rule: XmlNode): string | undefined | null {
  const extLst = childNodes(rule, "extLst", inMain)[0];
  if (!extLst) {
    return undefined;
  }
  const exts = extLst.children;
  const ext = exts[0];
  if (exts.length !== 1 || !ext || ext.local !== "ext" || (nodeAttr(ext, "uri") ?? "").toUpperCase() !== RULE_ID_EXT_URI) {
    return null;
  }
  const id = childNodes(ext, "id", inX14)[0]?.text.trim();
  return id ? id : null;
}

function readIconSet(node: XmlNode, valueOf: (cfvo: XmlNode) => CfValue | undefined, cfvoNs: (ns: string | undefined) => boolean): CfRuleBody | null {
  const set = nodeAttr(node, "iconSet") ?? "3TrafficLights1";
  if (!ICON_SETS.has(set)) {
    return null;
  }
  const cfvos = childNodes(node, "cfvo", cfvoNs);
  if (cfvos.length !== iconSetSize(set as IconSetName) || cfvos.length !== node.children.length) {
    return null;
  }
  const thresholds: { at: CfValue; gte: boolean }[] = [];
  for (const cfvo of cfvos) {
    const at = valueOf(cfvo);
    if (!at) {
      return null;
    }
    thresholds.push({ at, gte: !isOff(nodeAttr(cfvo, "gte")) });
  }
  return { kind: "iconSet", icons: { set: set as IconSetName, thresholds, reverse: isOn(nodeAttr(node, "reverse")), showValue: !isOff(nodeAttr(node, "showValue")) } };
}

const TEXT_TYPES: Readonly<Record<string, Extract<CfRuleBody, { kind: "text" }>["operator"]>> = {
  containsText: "contains",
  notContainsText: "notContains",
  beginsWith: "beginsWith",
  endsWith: "endsWith",
};

/** A base cfRule's kind and fields, or null when the model cannot hold it. */
function baseBody(node: XmlNode, ctx: ReadContext): (CfRuleBody & { x14Id?: string }) | null {
  const type = nodeAttr(node, "type") ?? "";
  const formulas = childNodes(node, "formula", inMain).map((formula) => fromXlsxFormula(formula.text.trim()));
  switch (type) {
    case "cellIs": {
      const operator = nodeAttr(node, "operator") ?? "";
      const count = operator === "between" || operator === "notBetween" ? 2 : 1;
      if (!OPERATORS.has(operator) || formulas.length < count || !onlyChildren(node, ["formula"])) {
        return null;
      }
      return { kind: "cellIs", operator: operator as CfOperator, formulas: formulas.slice(0, count) };
    }
    case "expression":
      return formulas.length >= 1 && onlyChildren(node, ["formula"]) ? { kind: "expression", formula: formulas[0]! } : null;
    case "colorScale": {
      const scale = childNodes(node, "colorScale", inMain)[0];
      if (!scale || !onlyChildren(node, ["colorScale"]) || !onlyChildren(scale, ["cfvo", "color"])) {
        return null;
      }
      const cfvos = childNodes(scale, "cfvo", inMain);
      const colors = childNodes(scale, "color", inMain);
      if ((cfvos.length !== 2 && cfvos.length !== 3) || colors.length !== cfvos.length) {
        return null;
      }
      const stops: { at: CfValue; color: ModelColor }[] = [];
      for (const [index, cfvo] of cfvos.entries()) {
        const at = baseValue(cfvo);
        const color = colorOf(colors[index], ctx);
        if (!at || color === undefined) {
          return null;
        }
        stops.push({ at, color });
      }
      return { kind: "colorScale", stops };
    }
    case "dataBar": {
      const bar = childNodes(node, "dataBar", inMain)[0];
      const guid = twinGuid(node);
      if (!bar || guid === null || !onlyChildren(node, ["dataBar", "extLst"]) || !onlyChildren(bar, ["cfvo", "color"])) {
        return null;
      }
      const cfvos = childNodes(bar, "cfvo", inMain);
      const min = cfvos[0] ? baseValue(cfvos[0]) : undefined;
      const max = cfvos[1] ? baseValue(cfvos[1]) : undefined;
      const color = colorOf(childNodes(bar, "color", inMain)[0], ctx);
      if (cfvos.length !== 2 || !min || !max || color === undefined) {
        return null;
      }
      const dataBar: CfDataBar = {
        min,
        max,
        color,
        gradient: true,
        axis: "none",
        direction: "context",
        minLengthPct: parseNumber(nodeAttr(bar, "minLength")) ?? 10,
        maxLengthPct: parseNumber(nodeAttr(bar, "maxLength")) ?? 90,
        showValue: !isOff(nodeAttr(bar, "showValue")),
      };
      return guid === undefined ? { kind: "dataBar", bar: dataBar } : { kind: "dataBar", bar: dataBar, x14Id: guid };
    }
    case "iconSet": {
      const icons = childNodes(node, "iconSet", inMain)[0];
      return icons && onlyChildren(node, ["iconSet"]) ? readIconSet(icons, baseValue, inMain) : null;
    }
    case "top10":
      return onlyChildren(node, ["formula"])
        ? { kind: "top10", bottom: isOn(nodeAttr(node, "bottom")), percent: isOn(nodeAttr(node, "percent")), rank: parseIndex(nodeAttr(node, "rank")) ?? 10 }
        : null;
    case "aboveAverage": {
      const stdDev = parseIndex(nodeAttr(node, "stdDev")) ?? 0;
      if (stdDev > 3 || !onlyChildren(node, ["formula"])) {
        return null;
      }
      const body: CfRuleBody = { kind: "aboveAverage", below: isOff(nodeAttr(node, "aboveAverage")), equalAverage: isOn(nodeAttr(node, "equalAverage")) };
      return stdDev === 0 ? body : { ...body, stdDev: stdDev as 1 | 2 | 3 };
    }
    case "duplicateValues":
    case "uniqueValues":
      return onlyChildren(node, ["formula"]) ? { kind: "duplicates", unique: type === "uniqueValues" } : null;
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith": {
      const text = nodeAttr(node, "text");
      return text !== undefined && onlyChildren(node, ["formula"]) ? { kind: "text", operator: TEXT_TYPES[type]!, text } : null;
    }
    case "containsBlanks":
    case "notContainsBlanks":
      return onlyChildren(node, ["formula"]) ? { kind: "blanks", not: type === "notContainsBlanks" } : null;
    case "containsErrors":
    case "notContainsErrors":
      return onlyChildren(node, ["formula"]) ? { kind: "errors", not: type === "notContainsErrors" } : null;
    case "timePeriod": {
      const period = nodeAttr(node, "timePeriod") ?? "";
      return TIME_PERIODS.has(period) && onlyChildren(node, ["formula"]) ? { kind: "timePeriod", period: period as CfTimePeriod } : null;
    }
    default:
      return null;
  }
}

/** The fields every rule shares: a minted id, its written priority (to be
 * renumbered), `stopIfTrue`. */
function ruleBase(node: XmlNode, ctx: ReadContext): { id: string; priority: number; stopIfTrue?: true } {
  const priority = parseNumber(nodeAttr(node, "priority"));
  const base: { id: string; priority: number; stopIfTrue?: true } = { id: ctx.mint(), priority: priority ?? NO_PRIORITY };
  if (isOn(nodeAttr(node, "stopIfTrue"))) {
    base.stopIfTrue = true;
  }
  return base;
}

function readBaseRule(node: XmlNode, text: string, ctx: ReadContext): CfRule {
  const base = ruleBase(node, ctx);
  const body = baseBody(node, ctx);
  if (body === null) {
    return { id: base.id, priority: base.priority, kind: "preserved", xml: text.slice(node.start, node.end), where: "base" };
  }
  const rule = { ...base, ...body } as CfRule;
  const dxfId = parseIndex(nodeAttr(node, "dxfId"));
  if (dxfId !== undefined) {
    rule.dxfId = dxfId;
  }
  return rule;
}

/** An x14 data bar's settings over its base twin's (see the module header). */
function overlayDataBar(bar: CfDataBar, node: XmlNode, ctx: ReadContext): void {
  bar.minLengthPct = parseNumber(nodeAttr(node, "minLength")) ?? 10;
  bar.maxLengthPct = parseNumber(nodeAttr(node, "maxLength")) ?? 90;
  bar.gradient = !isOff(nodeAttr(node, "gradient"));
  const direction = nodeAttr(node, "direction");
  bar.direction = direction === "leftToRight" || direction === "rightToLeft" ? direction : "context";
  const axis = nodeAttr(node, "axisPosition");
  bar.axis = axis === "middle" || axis === "none" ? axis : "automatic";
  const cfvos = childNodes(node, "cfvo", inX14);
  const min = cfvos[0] ? x14Value(cfvos[0]) : undefined;
  const max = cfvos[1] ? x14Value(cfvos[1]) : undefined;
  if (cfvos.length === 2 && min && max) {
    bar.min = min;
    bar.max = max;
  }
  const color = (local: string): ModelColor | undefined => colorOf(childNodes(node, local, inX14)[0], ctx);
  bar.color = color("fillColor") ?? bar.color;
  delete bar.border;
  delete bar.negativeColor;
  delete bar.negativeBorderColor;
  delete bar.axisColor;
  if (isOn(nodeAttr(node, "border"))) {
    bar.border = color("borderColor") ?? bar.color;
  }
  const negative = isOn(nodeAttr(node, "negativeBarColorSameAsPositive")) ? bar.color : color("negativeFillColor");
  if (negative !== undefined) {
    bar.negativeColor = negative;
  }
  const negativeBorder = isOff(nodeAttr(node, "negativeBarBorderColorSameAsPositive")) ? color("negativeBorderColor") : undefined;
  if (negativeBorder !== undefined) {
    bar.negativeBorderColor = negativeBorder;
  }
  const axisColor = color("axisColor");
  if (axisColor !== undefined) {
    bar.axisColor = axisColor;
  }
}

/** An x14 rule with no base twin (see the module header). */
function readX14Rule(node: XmlNode, text: string, ctx: ReadContext): CfRule {
  const base = ruleBase(node, ctx);
  const id = nodeAttr(node, "id");
  const preserved: CfRule = { id: base.id, priority: base.priority, kind: "preserved", xml: text.slice(node.start, node.end), where: "x14" };
  const type = nodeAttr(node, "type");
  let body: CfRuleBody | null = null;
  let dxfId: number | undefined;
  if (type === "iconSet") {
    const icons = childNodes(node, "iconSet", inX14)[0];
    const set = icons ? (nodeAttr(icons, "iconSet") ?? "3TrafficLights1") : "";
    if (icons && onlyChildren(node, ["iconSet"], inX14) && X14_ICON_SETS.has(set as IconSetName) && !isOn(nodeAttr(icons, "custom"))) {
      body = readIconSet(icons, x14Value, inX14);
    }
  } else if (type === "cellIs" || type === "expression") {
    const allowed = node.children.every((child) => (inXm(child.ns) && child.local === "f") || (inX14(child.ns) && child.local === "dxf"));
    const formulas = childNodes(node, "f", inXm).map((formula) => fromXlsxFormula(formula.text.trim()));
    const operator = nodeAttr(node, "operator") ?? "";
    const count = type === "expression" ? 1 : operator === "between" || operator === "notBetween" ? 2 : 1;
    const used = formulas.slice(0, count);
    if (allowed && used.length === count && (type === "expression" || OPERATORS.has(operator)) && used.some((formula) => referencesOtherSheet(formula, ctx.sheetName))) {
      body = type === "expression" ? { kind: "expression", formula: used[0]! } : { kind: "cellIs", operator: operator as CfOperator, formulas: used };
      const dxfNode = childNodes(node, "dxf", inX14)[0];
      const dxf = dxfNode ? readDxfNode(dxfNode, ctx.indexedColors) : undefined;
      if (dxf !== undefined && Object.keys(dxf).length > 0) {
        dxfId = internDxf(ctx.holder, dxf);
      }
    }
  }
  if (body === null) {
    return preserved;
  }
  const rule = { ...base, ...body, x14: true } as CfRule;
  if (id !== undefined && id !== "") {
    rule.x14Id = id;
  }
  if (dxfId !== undefined) {
    rule.dxfId = dxfId;
  }
  return rule;
}

/** A worksheet's conditional formats from what the reader cut out of its
 * part (see the module header); x14 inline dxfs are interned into
 * `holder.dxfs`. `mint` names each format and rule (default mintCfId). */
export function readConditionalFormatSource(
  source: WorksheetCfSource,
  sheetName: string,
  holder: Pick<SpreadsheetModel, "dxfs">,
  indexedColors?: readonly string[],
  mint: () => string = mintCfId,
): ConditionalFormat[] {
  const ctx: ReadContext = { sheetName, holder, indexedColors, mint };
  const context = [source.rootTag];
  const formats: ConditionalFormat[] = [];
  const twins = new Map<string, Extract<CfRule, { kind: "dataBar" }>>();
  for (const block of source.blocks) {
    const node = elementTree(block, context)?.find((each) => each.local === "conditionalFormatting" && inMain(each.ns));
    const ranges = node ? parseCfSqref(nodeAttr(node, "sqref") ?? "") : [];
    const rules = node ? childNodes(node, "cfRule", inMain).map((rule) => readBaseRule(rule, block, ctx)) : [];
    if (!node || ranges.length === 0 || rules.length === 0) {
      continue;
    }
    for (const rule of rules) {
      if (rule.kind === "dataBar" && rule.x14Id !== undefined && !twins.has(rule.x14Id.toUpperCase())) {
        twins.set(rule.x14Id.toUpperCase(), rule);
      }
    }
    const format: ConditionalFormat = { id: mint(), ranges, rules, sourceXml: block };
    if (isOn(nodeAttr(node, "pivot"))) {
      format.pivot = true;
    }
    formats.push(format);
  }
  const extLst = source.extLst === undefined ? undefined : elementTree(source.extLst, context)?.[0];
  const extText = source.extLst ?? "";
  const exts = extLst ? childNodes(extLst, "ext", inMain).filter((ext) => (nodeAttr(ext, "uri") ?? "").toUpperCase() === CF_EXT_URI.toUpperCase()) : [];
  for (const ext of exts) {
    for (const list of childNodes(ext, "conditionalFormattings", inX14)) {
      for (const block of childNodes(list, "conditionalFormatting", inX14)) {
        const ranges = parseCfSqref(childNodes(block, "sqref", inXm)[0]?.text ?? "");
        const rules: CfRule[] = [];
        let paired = false;
        for (const node of childNodes(block, "cfRule", inX14)) {
          const twin = twins.get((nodeAttr(node, "id") ?? "").toUpperCase());
          const x14Bar = childNodes(node, "dataBar", inX14)[0];
          if (twin && nodeAttr(node, "type") === "dataBar" && x14Bar) {
            overlayDataBar(twin.bar, x14Bar, ctx);
            twins.delete(twin.x14Id!.toUpperCase());
            paired = true;
            continue;
          }
          rules.push(readX14Rule(node, extText, ctx));
        }
        if (ranges.length === 0 || rules.length === 0) {
          continue;
        }
        const format: ConditionalFormat = { id: mint(), ranges, rules };
        if (isOn(nodeAttr(block, "pivot"))) {
          format.pivot = true;
        }
        if (!paired) {
          format.sourceXml = extText.slice(whitespaceBefore(extText, block.start), block.end);
        }
        formats.push(format);
      }
    }
  }
  return renumberPriorities(formats, { keepSource: true });
}

// --- writing ------------------------------------------------------------------------------

/** What the writer needs from the save around it. */
export interface CfWriteContext {
  /** The part's main-namespace prefix ("" or "x:"). */
  prefix: string;
  theme: DocumentTheme | undefined;
  /** The workbook's differential styles (an x14 rule's inline dxf). */
  dxfs: readonly Dxf[];
  /** The id a dxf number format is written with (the save's style writer). */
  numFmtId(code: string): number;
  /** The dxfId a base rule writes for a model dxf (the save's style writer,
   * which appends the model's dxfs the part lacks up to it). */
  dxfId?(id: number): number;
  /** Set when the part has a retained source: source pieces may be
   * re-emitted — through `pass`, the sheet-name pass — in `context` (the
   * root start tag). */
  verbatim?: { pass(xml: string): string; context: readonly string[] };
}

/** The writer's output: the base run, and the x14 ext's
 * `<x14:conditionalFormatting>` elements. */
export interface CfWriteOutput {
  run: string;
  x14: string[];
}

/** A fresh x14 GUID, as Excel writes them. */
export function mintX14Guid(): string {
  return `{${crypto.randomUUID().toUpperCase()}}`;
}

/** The written form of a block: verbatim (with the rule start tags whose
 * priority it carries) or serialised. */
interface Piece {
  text: string;
  /** Verbatim pieces only: each rule's start tag offset and written priority. */
  slots?: { rule: CfRule; at: number; written: number }[];
}

const X14_CONTEXT = `<_ xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}">`;

/** A verbatim piece's rule slots: its cfRule elements (base or x14) mapped
 * to the format's rules by written priority; undefined when they do not
 * line up. */
function slotsOf(text: string, format: ConditionalFormat, x14: boolean, context: readonly string[]): Piece["slots"] {
  const node = elementTree(text, x14 ? [...context, X14_CONTEXT] : context)?.find((each) => each.local === "conditionalFormatting");
  if (!node) {
    return undefined;
  }
  const rules = childNodes(node, "cfRule", x14 ? inX14 : inMain).map((rule, order) => ({ at: rule.start, written: parseNumber(nodeAttr(rule, "priority")) ?? NO_PRIORITY, order }));
  if (rules.length !== format.rules.length) {
    return undefined;
  }
  rules.sort((a, b) => a.written - b.written || a.order - b.order);
  const ordered = format.rules.toSorted((a, b) => a.priority - b.priority);
  return rules.map((slot, index) => ({ rule: ordered[index]!, at: slot.at, written: slot.written }));
}

/** `xml` (a rule element) with its `priority` attribute set — or left
 * alone when it carries none and `always` is false. */
function withPriority(xml: string, priority: number, always: boolean): string {
  const tag = startTagAt(xml, 0);
  if (!always && !/\spriority\s*=/.test(tag)) {
    return xml;
  }
  return setAttribute(tag, "priority", String(priority)) + xml.slice(tag.length);
}

/** Patches the pieces whose written priorities no longer rank in the
 * model's order (see the module header). */
function patchPriorities(pieces: readonly Piece[], written: ReadonlyMap<string, number>): void {
  for (;;) {
    const entries: { piece: Piece | undefined; model: number; written: number }[] = [];
    const slotted = new Map<string, Piece>();
    for (const piece of pieces) {
      for (const slot of piece.slots ?? []) {
        slotted.set(slot.rule.id, piece);
        entries.push({ piece, model: slot.rule.priority, written: slot.written });
      }
    }
    for (const [id, priority] of written) {
      if (!slotted.has(id)) {
        entries.push({ piece: undefined, model: priority, written: priority });
      }
    }
    entries.sort((a, b) => a.model - b.model);
    const conflicted = new Set<Piece>();
    for (let i = 0; i < entries.length; i += 1) {
      for (let j = i + 1; j < entries.length; j += 1) {
        if (entries[i]!.written >= entries[j]!.written) {
          for (const piece of [entries[i]!.piece, entries[j]!.piece]) {
            if (piece !== undefined && piece.slots!.some((slot) => slot.written !== slot.rule.priority)) {
              conflicted.add(piece);
            }
          }
        }
      }
    }
    if (conflicted.size === 0) {
      return;
    }
    for (const piece of conflicted) {
      const splices: Splice[] = [];
      let shift = 0;
      for (const slot of piece.slots!.toSorted((a, b) => a.at - b.at)) {
        const tag = startTagAt(piece.text, slot.at);
        const next = setAttribute(tag, "priority", String(slot.rule.priority));
        splices.push({ start: slot.at, end: slot.at + tag.length, text: next });
        slot.at += shift;
        shift += next.length - tag.length;
        slot.written = slot.rule.priority;
      }
      piece.text = applySplices(piece.text, splices);
    }
  }
}

function attr(name: string, value: string | number | undefined): string {
  return value === undefined ? "" : ` ${name}="${typeof value === "number" ? String(value) : escapeAttributeValue(value)}"`;
}

function colorXml(p: string, local: string, color: ModelColor, theme: DocumentTheme | undefined): string {
  return `<${p}${local}${modelColorToXlsxAttrs(color, theme)}/>`;
}

/** A value object's wire text: a number as written, a formula through toXlsxFormula. */
function valueWire(value: CfValue): string | undefined {
  if (value.value === undefined) {
    return undefined;
  }
  return value.type !== "formula" && NUMBER.test(value.value) ? value.value : toXlsxFormula(value.value);
}

/** A base `<cfvo>` (no auto types before Excel 2010: min/max stand for them). */
function baseCfvo(p: string, value: CfValue, gte = true): string {
  const type = value.type === "autoMin" ? "min" : value.type === "autoMax" ? "max" : value.type;
  const wire = type === "min" || type === "max" ? undefined : valueWire(value);
  return `<${p}cfvo${attr("type", type)}${attr("val", wire)}${gte ? "" : ` gte="0"`}/>`;
}

/** An `<x14:cfvo>` with its value in `<xm:f>`. */
function x14Cfvo(value: CfValue, gte = true): string {
  const wire = value.type === "min" || value.type === "max" || value.type === "autoMin" || value.type === "autoMax" ? undefined : valueWire(value);
  const open = `<x14:cfvo${attr("type", value.type)}${gte ? "" : ` gte="0"`}`;
  return wire === undefined ? `${open}/>` : `${open}><xm:f>${escapeElementText(wire)}</xm:f></x14:cfvo>`;
}

/** A string literal in an xlsx formula. */
const literal = (text: string): string => `"${text.replace(/"/g, '""')}"`;

/** Excel's formula for a time period, `a1` the origin. */
function timePeriodFormula(period: CfTimePeriod, a1: string): string {
  switch (period) {
    case "today":
      return `FLOOR(${a1},1)=TODAY()`;
    case "yesterday":
      return `FLOOR(${a1},1)=TODAY()-1`;
    case "tomorrow":
      return `FLOOR(${a1},1)=TODAY()+1`;
    case "last7Days":
      return `AND(TODAY()-FLOOR(${a1},1)<=6,FLOOR(${a1},1)<=TODAY())`;
    case "thisWeek":
      return `AND(TODAY()-ROUNDDOWN(${a1},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a1},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case "lastWeek":
      return `AND(TODAY()-ROUNDDOWN(${a1},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a1},0)<(WEEKDAY(TODAY())+7))`;
    case "nextWeek":
      return `AND(ROUNDDOWN(${a1},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a1},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case "thisMonth":
      return `AND(MONTH(${a1})=MONTH(TODAY()),YEAR(${a1})=YEAR(TODAY()))`;
    case "lastMonth":
      return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a1})=YEAR(EDATE(TODAY(),0-1)))`;
    case "nextMonth":
      return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a1})=YEAR(EDATE(TODAY(),0+1)))`;
    case "thisYear":
      return `YEAR(${a1})=YEAR(TODAY())`;
    case "lastYear":
      return `YEAR(${a1})=YEAR(TODAY())-1`;
    case "nextYear":
      return `YEAR(${a1})=YEAR(TODAY())+1`;
  }
}

/** The `<formula>` Excel evaluates for a text, blank, error or time-period
 * rule (see the module header), `a1` the origin — wire text. */
export function regeneratedFormula(rule: CfRule, a1: string): string | undefined {
  switch (rule.kind) {
    case "text": {
      const t = literal(rule.text);
      switch (rule.operator) {
        case "contains":
          return `NOT(ISERROR(SEARCH(${t},${a1})))`;
        case "notContains":
          return `ISERROR(SEARCH(${t},${a1}))`;
        case "beginsWith":
          return `LEFT(${a1},LEN(${t}))=${t}`;
        case "endsWith":
          return `RIGHT(${a1},LEN(${t}))=${t}`;
      }
      return undefined;
    }
    case "blanks":
      return rule.not ? `LEN(TRIM(${a1}))>0` : `LEN(TRIM(${a1}))=0`;
    case "errors":
      return rule.not ? `NOT(ISERROR(${a1}))` : `ISERROR(${a1})`;
    case "timePeriod":
      return timePeriodFormula(rule.period, a1);
    default:
      return undefined;
  }
}

const TEXT_WIRE: Readonly<Record<Extract<CfRule, { kind: "text" }>["operator"], { type: string; operator: string }>> = {
  contains: { type: "containsText", operator: "containsText" },
  notContains: { type: "notContainsText", operator: "notContains" },
  beginsWith: { type: "beginsWith", operator: "beginsWith" },
  endsWith: { type: "endsWith", operator: "endsWith" },
};

/** A base `<cfRule>` serialised (see the module header); `guid` names a
 * data bar's x14 twin. */
function baseRuleXml(rule: CfRule, format: ConditionalFormat, ctx: CfWriteContext, guid: string | undefined): string {
  const p = ctx.prefix;
  const theme = ctx.theme;
  const formula = (text: string): string => `<${p}formula>${escapeElementText(text)}</${p}formula>`;
  const a1 = formatA1(formatOrigin(format));
  let type = "";
  let attrs = "";
  let children = "";
  switch (rule.kind) {
    case "cellIs":
      type = "cellIs";
      attrs = attr("operator", rule.operator);
      children = rule.formulas.map((each) => formula(toXlsxFormula(each))).join("");
      break;
    case "expression":
      type = "expression";
      children = formula(toXlsxFormula(rule.formula));
      break;
    case "colorScale":
      type = "colorScale";
      children = `<${p}colorScale>${rule.stops.map((stop) => baseCfvo(p, stop.at)).join("")}${rule.stops.map((stop) => colorXml(p, "color", stop.color, theme)).join("")}</${p}colorScale>`;
      break;
    case "dataBar":
      type = "dataBar";
      children =
        `<${p}dataBar${rule.bar.showValue ? "" : ` showValue="0"`}>${baseCfvo(p, rule.bar.min)}${baseCfvo(p, rule.bar.max)}${colorXml(p, "color", rule.bar.color, theme)}</${p}dataBar>` +
        (guid === undefined ? "" : `<${p}extLst><${p}ext uri="${RULE_ID_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>${escapeElementText(guid)}</x14:id></${p}ext></${p}extLst>`);
      break;
    case "iconSet": {
      const icons = rule.icons;
      type = "iconSet";
      children =
        `<${p}iconSet${icons.set === "3TrafficLights1" ? "" : attr("iconSet", icons.set)}${icons.showValue ? "" : ` showValue="0"`}${icons.reverse ? ` reverse="1"` : ""}>` +
        `${icons.thresholds.map((threshold) => baseCfvo(p, threshold.at, threshold.gte)).join("")}</${p}iconSet>`;
      break;
    }
    case "top10":
      type = "top10";
      attrs = (rule.percent ? ` percent="1"` : "") + (rule.bottom ? ` bottom="1"` : "") + attr("rank", rule.rank);
      break;
    case "aboveAverage":
      type = "aboveAverage";
      attrs = (rule.below ? ` aboveAverage="0"` : "") + attr("stdDev", rule.stdDev) + (rule.equalAverage ? ` equalAverage="1"` : "");
      break;
    case "duplicates":
      type = rule.unique ? "uniqueValues" : "duplicateValues";
      break;
    case "text":
      type = TEXT_WIRE[rule.operator].type;
      attrs = attr("operator", TEXT_WIRE[rule.operator].operator) + attr("text", rule.text);
      children = formula(regeneratedFormula(rule, a1)!);
      break;
    case "blanks":
      type = rule.not ? "notContainsBlanks" : "containsBlanks";
      children = formula(regeneratedFormula(rule, a1)!);
      break;
    case "errors":
      type = rule.not ? "notContainsErrors" : "containsErrors";
      children = formula(regeneratedFormula(rule, a1)!);
      break;
    case "timePeriod":
      if (rule.period === "thisYear" || rule.period === "lastYear" || rule.period === "nextYear") {
        type = "expression";
      } else {
        type = "timePeriod";
        attrs = attr("timePeriod", rule.period);
      }
      children = formula(regeneratedFormula(rule, a1)!);
      break;
    case "preserved":
      return withPriority(rule.xml, rule.priority, true);
  }
  const dxfId = rule.dxfId === undefined ? undefined : (ctx.dxfId?.(rule.dxfId) ?? rule.dxfId);
  const head = `<${p}cfRule${attr("type", type)}${attr("dxfId", dxfId)}${attr("priority", rule.priority)}${rule.stopIfTrue ? ` stopIfTrue="1"` : ""}${attrs}`;
  return children === "" ? `${head}/>` : `${head}>${children}</${p}cfRule>`;
}

/** The x14 twin of a data bar (see the module header). */
function x14DataBarXml(bar: CfDataBar, guid: string, theme: DocumentTheme | undefined): string {
  const sameNegative = bar.negativeColor !== undefined && sameModelColor(bar.negativeColor, bar.color);
  let attrs = attr("minLength", bar.minLengthPct) + attr("maxLength", bar.maxLengthPct);
  attrs += (bar.border !== undefined ? ` border="1"` : "") + (bar.gradient ? "" : ` gradient="0"`);
  attrs += bar.direction === "context" ? "" : attr("direction", bar.direction);
  attrs += sameNegative ? ` negativeBarColorSameAsPositive="1"` : "";
  attrs += bar.negativeBorderColor !== undefined ? ` negativeBarBorderColorSameAsPositive="0"` : "";
  attrs += bar.axis === "automatic" ? "" : attr("axisPosition", bar.axis);
  let children = x14Cfvo(bar.min) + x14Cfvo(bar.max);
  if (bar.border !== undefined) {
    children += colorXml("x14:", "borderColor", bar.border, theme);
  }
  if (!sameNegative) {
    children += colorXml("x14:", "negativeFillColor", bar.negativeColor ?? "#ff0000", theme);
  }
  if (bar.negativeBorderColor !== undefined) {
    children += colorXml("x14:", "negativeBorderColor", bar.negativeBorderColor, theme);
  }
  children += colorXml("x14:", "axisColor", bar.axisColor ?? "#000000", theme);
  return `<x14:cfRule type="dataBar"${attr("id", guid)}><x14:dataBar${attrs}>${children}</x14:dataBar></x14:cfRule>`;
}

/** A canonical formula as an x14 rule's `<xm:f>`. */
function xmFormula(canonical: string): string {
  return `<xm:f>${escapeElementText(toXlsxFormula(canonical))}</xm:f>`;
}

/** A rule writesToX14 places in the x14 ext (or a preserved x14 one). */
function x14RuleXml(rule: CfRule, ctx: CfWriteContext): string {
  if (rule.kind === "preserved") {
    return withPriority(rule.xml, rule.priority, false);
  }
  const head = (type: string, extra = ""): string =>
    `<x14:cfRule${attr("type", type)}${attr("priority", rule.priority)}${rule.stopIfTrue ? ` stopIfTrue="1"` : ""}${extra}${attr("id", rule.x14Id ?? mintX14Guid())}>`;
  const dxf = rule.dxfId === undefined ? undefined : ctx.dxfs[rule.dxfId];
  const dxfPart = dxf === undefined ? "" : `<x14:dxf>${dxfXml(dxf, ctx.theme, ctx.prefix, (code) => ctx.numFmtId(code))}</x14:dxf>`;
  switch (rule.kind) {
    case "iconSet": {
      const icons = rule.icons;
      return (
        head("iconSet") +
        `<x14:iconSet${attr("iconSet", icons.set)}${icons.showValue ? "" : ` showValue="0"`}${icons.reverse ? ` reverse="1"` : ""}>` +
        `${icons.thresholds.map((threshold) => x14Cfvo(threshold.at, threshold.gte)).join("")}</x14:iconSet></x14:cfRule>`
      );
    }
    case "cellIs":
      return `${head("cellIs", attr("operator", rule.operator))}${rule.formulas.map(xmFormula).join("")}${dxfPart}</x14:cfRule>`;
    case "expression":
      return `${head("expression")}${xmFormula(rule.formula)}${dxfPart}</x14:cfRule>`;
    default:
      return "";
  }
}

/** Whether a format holding a `sourceXml` came from the x14 ext. */
function isX14Format(format: ConditionalFormat): boolean {
  return format.rules.length > 0 && format.rules.every((rule) => (rule.kind === "preserved" ? rule.where === "x14" : rule.x14 === true));
}

/** The model's formats as the base run and the x14 ext's blocks (see the
 * module header). */
export function writeConditionalFormats(formats: readonly ConditionalFormat[], sheetName: string, ctx: CfWriteContext): CfWriteOutput {
  const p = ctx.prefix;
  const base: Piece[] = [];
  const x14: Piece[] = [];
  const serialisedPriorities = new Map<string, number>();
  for (const format of formats) {
    const rules = format.rules.filter((rule) => rule.kind !== "preserved" || rule.where !== "ods").toSorted((a, b) => a.priority - b.priority);
    const source = ctx.verbatim && format.sourceXml !== undefined && /^\s*<([\w.-]+:)?conditionalFormatting[\s>/]/.test(format.sourceXml) ? format.sourceXml : undefined;
    const fromX14 = isX14Format(format);
    if (source !== undefined && ctx.verbatim) {
      const text = ctx.verbatim.pass(source);
      const slots = slotsOf(text, format, fromX14, ctx.verbatim.context);
      if (slots !== undefined) {
        (fromX14 ? x14 : base).push({ text, slots });
        if (!fromX14) {
          // The data bars' x14 twins are rewritten with the ext.
          const twins = rules.flatMap((rule) => (rule.kind === "dataBar" && rule.x14Id !== undefined ? [x14DataBarXml(rule.bar, rule.x14Id, ctx.theme)] : []));
          if (twins.length > 0) {
            x14.push({ text: `<x14:conditionalFormatting${format.pivot ? ` pivot="1"` : ""}>${twins.join("")}<xm:sqref>${sqrefText(format.ranges)}</xm:sqref></x14:conditionalFormatting>` });
          }
        }
        continue;
      }
    }
    if (format.ranges.length === 0 || rules.length === 0) {
      continue;
    }
    const baseRules: string[] = [];
    const x14Rules: string[] = [];
    for (const rule of rules) {
      serialisedPriorities.set(rule.id, rule.priority);
      if (rule.kind === "dataBar") {
        const guid = rule.x14Id ?? mintX14Guid();
        baseRules.push(baseRuleXml(rule, format, ctx, guid));
        x14Rules.push(x14DataBarXml(rule.bar, guid, ctx.theme));
      } else if (rule.kind === "preserved" ? rule.where === "x14" : writesToX14(rule, sheetName)) {
        x14Rules.push(x14RuleXml(rule, ctx));
      } else {
        baseRules.push(baseRuleXml(rule, format, ctx, undefined));
      }
    }
    const pivot = format.pivot ? ` pivot="1"` : "";
    if (baseRules.length > 0) {
      base.push({ text: `<${p}conditionalFormatting${attr("sqref", sqrefText(format.ranges))}${pivot}>${baseRules.join("")}</${p}conditionalFormatting>` });
    }
    if (x14Rules.length > 0) {
      x14.push({ text: `<x14:conditionalFormatting${pivot}>${x14Rules.join("")}<xm:sqref>${sqrefText(format.ranges)}</xm:sqref></x14:conditionalFormatting>` });
    }
  }
  patchPriorities([...base, ...x14], serialisedPriorities);
  return { run: base.map((piece) => piece.text).join(""), x14: x14.map((piece) => piece.text) };
}

/** The `{78C0D931-…}` ext holding `blocks`: the source ext's start tag (or a
 * new one) with `xmlns:x14` and `xmlns:xm` declared on it. */
export function conditionalFormattingsExtXml(blocks: readonly string[], p: string, source?: { tag: string; name: string }): string {
  let tag = source?.tag ?? `<${p}ext uri="${CF_EXT_URI}">`;
  if (tag.endsWith("/>")) {
    tag = `${tag.slice(0, -2).trimEnd()}>`;
  }
  if (!/\sxmlns:x14\s*=/.test(tag)) {
    tag = setAttribute(tag, "xmlns:x14", X14_NS);
  }
  if (!/\sxmlns:xm\s*=/.test(tag)) {
    tag = setAttribute(tag, "xmlns:xm", XM_NS);
  }
  return `${tag}<x14:conditionalFormattings>${blocks.join("")}</x14:conditionalFormattings></${source?.name ?? `${p}ext`}>`;
}
