// The formulas a conditional-format rule carries, and how a reference change
// reaches them (SPEC-129 "Structural edits" and "Moves", SPEC-121's sheet
// rename and delete) — THIS FILE is the normative source for which texts
// those are and how they are mapped; structure.ts's and move.ts's
// `conditionalFormats` adjusters and sheets.ts's commands all go through it.
//
// | Rule | Formulas | Mapped by |
// | `cellIs` | `formulas` (1, or 2 for (not)between), relative to the format's origin | `relative` |
// | `expression` | `formula`, relative to the origin | `relative` |
// | `colorScale`, `dataBar`, `iconSet` | each value object's formula — `stops[].at`, `bar.min`/`bar.max`, `icons.thresholds[].at` — absolute | `absolute` |
// | `preserved` (xlsx, `base`/`x14`) | the text of every `<formula>` and `<xm:f>` element (any prefix), through fromXlsxFormula/toXlsxFormula | `absolute` |
// | `preserved` (`ods`) and every other kind | none | — |
//
// - A value object's `value` is a formula exactly when cfEvaluator.ts's
//   resolveCfValue treats it as one: a `formula` value always; a
//   `num`/`percent`/`percentile` value when it is neither blank nor a plain
//   number (the evaluator's PLAIN_NUMBER); a `min`/`max`/`autoMin`/`autoMax`
//   value never (it is not read). A leading "=" (the evaluator strips one)
//   stays out of the mapped text and is kept.
// - A preserved xlsx rule's XML is scanned once; only the formula texts the
//   map changes are spliced (re-escaped), every other byte stays. XML that
//   does not scan is left as it is.
// - Copy-on-write: a value object, rule or format is replaced only when one
//   of its formulas changes — the format then without its `sourceXml` (a
//   fresh element is serialised) — and comes back as the same object
//   otherwise, so callers keep the identity of everything untouched
//   (`conditionalFormats` is shared by reference with SPEC-121's snapshots).
//   A changed `cellIs`/`expression` rule is re-placed with withX14Placement
//   for its host: a formula that came to reference another sheet (a move
//   across sheets) or stopped to (a deleted sheet's `#REF!`) changes where
//   an xlsx writer puts it.
import { scanXml, XmlScanError } from "../core/xmlScanner";
import { withX14Placement, type CfRule, type CfValue, type ConditionalFormat } from "./conditionalFormats";
import { fromXlsxFormula, toXlsxFormula } from "./formula/dialect";
import type { SpreadsheetFormat } from "./model";
import { applySplices, decodeXmlText, escapeElementText, type Splice } from "./xmlSplice";

/** One formula's text (canonical, no "=") mapped. */
export type FormulaMap = (formula: string) => string;

/** The maps a rule's formulas go through (see the module header's table). */
export interface RuleFormulaMaps {
  /** `cellIs`/`expression` formulas, stored relative to the origin. */
  relative: FormulaMap;
  /** Value-object formulas and a preserved xlsx rule's formula texts. */
  absolute: FormulaMap;
}

/** Where a rule's format lives: the workbook's format and the host sheet's
 * name (for withX14Placement). */
export interface RuleHost {
  format: SpreadsheetFormat;
  sheetName: string;
}

/** The evaluator's plain-number test (cfEvaluator.ts's PLAIN_NUMBER). */
const PLAIN_NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?\s*$/i;
const LEADING_EQUALS = /^\s*=/;

/** Whether `value.value` is formula text (see the module header). */
export function isCfValueFormula(value: CfValue): boolean {
  const text = value.value;
  if (text === undefined) {
    return false;
  }
  switch (value.type) {
    case "formula":
      return true;
    case "num":
    case "percent":
    case "percentile":
      return text.trim() !== "" && !PLAIN_NUMBER.test(text);
    default:
      return false;
  }
}

/** A value object with its formula mapped; the same object when it holds
 * none or the map changes nothing. */
function mapCfValue(value: CfValue, map: FormulaMap): CfValue {
  if (!isCfValueFormula(value)) {
    return value;
  }
  const text = value.value!;
  const lead = LEADING_EQUALS.exec(text)?.[0] ?? "";
  const formula = text.slice(lead.length);
  const next = map(formula);
  return next === formula ? value : { ...value, value: lead + next };
}

/** Every `<formula>` / `<xm:f>` text of a preserved xlsx rule's XML mapped
 * (see the module header); the identical string when none changes. */
export function mapXlsxPreservedFormulas(xml: string, map: FormulaMap): string {
  const splices: Splice[] = [];
  let contentStart = -1;
  try {
    scanXml(xml, (event) => {
      if (event.type === "text" || (event.local !== "formula" && event.local !== "f")) {
        return;
      }
      if (event.type === "open") {
        contentStart = event.selfClosing ? -1 : event.end;
        return;
      }
      if (contentStart < 0) {
        return;
      }
      const raw = xml.slice(contentStart, event.start);
      contentStart = -1;
      if (raw.trim() === "") {
        return;
      }
      const canonical = fromXlsxFormula(decodeXmlText(raw));
      const next = map(canonical);
      if (next !== canonical) {
        splices.push({ start: event.start - raw.length, end: event.start, text: escapeElementText(toXlsxFormula(next)) });
      }
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return xml;
    }
    throw error;
  }
  return splices.length === 0 ? xml : applySplices(xml, splices);
}

/** `rule` with every formula it carries mapped (see the module header); the
 * same object when none changes. */
export function mapRuleFormulas(rule: CfRule, maps: RuleFormulaMaps, host: RuleHost): CfRule {
  switch (rule.kind) {
    case "cellIs": {
      const formulas = rule.formulas.map((formula) => maps.relative(formula));
      return formulas.every((formula, i) => formula === rule.formulas[i]) ? rule : withX14Placement({ ...rule, formulas }, host.format, host.sheetName);
    }
    case "expression": {
      const formula = maps.relative(rule.formula);
      return formula === rule.formula ? rule : withX14Placement({ ...rule, formula }, host.format, host.sheetName);
    }
    case "colorScale": {
      const stops = rule.stops.map((stop) => {
        const at = mapCfValue(stop.at, maps.absolute);
        return at === stop.at ? stop : { ...stop, at };
      });
      return stops.every((stop, i) => stop === rule.stops[i]) ? rule : { ...rule, stops };
    }
    case "dataBar": {
      const min = mapCfValue(rule.bar.min, maps.absolute);
      const max = mapCfValue(rule.bar.max, maps.absolute);
      return min === rule.bar.min && max === rule.bar.max ? rule : { ...rule, bar: { ...rule.bar, min, max } };
    }
    case "iconSet": {
      const thresholds = rule.icons.thresholds.map((threshold) => {
        const at = mapCfValue(threshold.at, maps.absolute);
        return at === threshold.at ? threshold : { ...threshold, at };
      });
      return thresholds.every((threshold, i) => threshold === rule.icons.thresholds[i]) ? rule : { ...rule, icons: { ...rule.icons, thresholds } };
    }
    case "preserved": {
      if (rule.where === "ods") {
        return rule;
      }
      const xml = mapXlsxPreservedFormulas(rule.xml, maps.absolute);
      return xml === rule.xml ? rule : { ...rule, xml };
    }
    default:
      return rule;
  }
}

/** A format without its `sourceXml` (something about it changed). */
export function withoutFormatSource(format: ConditionalFormat): ConditionalFormat {
  if (format.sourceXml === undefined) {
    return format;
  }
  const { sourceXml: _source, ...rest } = format;
  return rest;
}

/** `format` with every rule's formulas mapped: the same object when none
 * changes, else a new format without its `sourceXml`. */
export function mapFormatFormulas(format: ConditionalFormat, maps: RuleFormulaMaps, host: RuleHost): ConditionalFormat {
  let changed = false;
  const rules = format.rules.map((rule) => {
    const next = mapRuleFormulas(rule, maps, host);
    changed ||= next !== rule;
    return next;
  });
  return changed ? withoutFormatSource({ ...format, rules }) : format;
}

/** `formats` with every formula mapped by one map (mapFormatFormulas): the
 * same array when nothing changes. */
export function mapFormatsFormulas(formats: readonly ConditionalFormat[], map: FormulaMap, host: RuleHost): ConditionalFormat[] {
  const maps: RuleFormulaMaps = { relative: map, absolute: map };
  let out: ConditionalFormat[] | null = null;
  formats.forEach((format, at) => {
    const next = mapFormatFormulas(format, maps, host);
    if (next !== format) {
      out ??= formats.slice();
      out[at] = next;
    }
  });
  return out ?? (formats as ConditionalFormat[]);
}
