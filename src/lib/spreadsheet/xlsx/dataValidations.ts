// The xlsx wire mapping of SPEC-130's data validations ("xlsx wire mapping"
// and "x14", normative there; validation.ts is the model's source): what
// xlsx/sheetXml.ts's reader makes of a worksheet's `<dataValidations>` and
// the x14 data validations in its `<extLst>`, and what its writer puts back.
//
// Reading (readValidationSource — the reader hands over the root start tag,
// the root-level `<dataValidations>` and the root `<extLst>`; or
// readWorksheetValidations over a whole worksheet part, which a reader
// upgrade runs on a retained part):
// - Each `<dataValidation>` is a rule: `type` (`none`, the default, ↔ `any`;
//   a type the model does not know reads as `any`), `operator` (default
//   `between`), `allowBlank`, `showDropDown` — "1" HIDES the arrow, so the
//   model's `showDropdown` is its negation —, `showInputMessage` +
//   `promptTitle` + `prompt` → `input`, `showErrorMessage` + `errorStyle`
//   (default `stop`) + `errorTitle` + `error` → `error` (attribute text as
//   the scanner decodes it: `&#10;` is a line break), `<formula1>` and
//   `<formula2>` through fromXlsxFormula, `sqref` → `ranges` (A1 ranges,
//   whole columns and rows — conditionalFormats.ts's parseCfSqref). The
//   element itself, verbatim, is the rule's `sourceXml`: `imeMode` and
//   `xr:uid` live only there. A rule whose sqref names nothing is dropped.
// - The worksheet extLst's `{CCE6A557-…}` ext › `x14:dataValidations` ›
//   `x14:dataValidation`: the same attributes, the formulas in
//   `x14:formula1`/`x14:formula2` › `xm:f`, the ranges in `xm:sqref`; read
//   with `x14: true`. Rules: the base ones in order, then the x14 ones.
// - `disablePrompts="1"` on either container → `promptsDisabled`
//   (`Sheet.validationPromptsDisabled`).
//
// Writing (writeDataValidations; sheetXml.ts splices the block into a DIRTY
// sheet's part — a clean part is never touched — at its CT_Worksheet place,
// after `conditionalFormatting`, before `hyperlinks`, and the ext into the
// extLst after SPEC-129's `{78C0D931-…}`, every other ext as found):
// - Placement: a rule still holding a `sourceXml` that fits the part keeps
//   the location it was read from (its `x14` flag) and is emitted verbatim
//   — through the writer's sheet-name pass, its text naming sheets as the
//   source did; any other rule goes to the x14 ext iff its formulas
//   reference another sheet (validation.ts writesToX14, as Excel does), else
//   to the base block, serialised: attributes in Excel's (the schema's)
//   order — type, errorStyle, operator, allowBlank, showDropDown,
//   showInputMessage, showErrorMessage, errorTitle, error, promptTitle,
//   prompt, sqref — each omitted at its default as Excel omits it, text
//   XML-escaped with `&#10;` for a line break; formulas through
//   toXlsxFormula.
// - The block is `<dataValidations count="n" [disablePrompts="1"]>`: the
//   source's start tag kept (its xWindow, yWindow …) with `count` and
//   `disablePrompts` set; none when no rule is placed there. The ext is
//   `<ext uri="{CCE6A557-…}" xmlns:x14>` › `<x14:dataValidations count
//   xmlns:xm>` — the source's start tags kept the same way — or nothing.
// - Each location is left exactly as the source has it while the rules the
//   model places there are the source's own, unedited and in order (and the
//   prompts flag agrees), so a dirty sheet whose validations were not
//   touched keeps their bytes.
// A model with no retained source (a new sheet, a stripRetention copy) is
// serialised whole; a source text from the other format has no xlsx form.
import { scanXml } from "../../core/xmlScanner";
import { formatRange, type GridRange } from "../address";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import type { Sheet } from "../model";
import {
  mintValidationId,
  writesToX14,
  type DataValidation,
  type ValidationErrorStyle,
  type ValidationOperator,
  type ValidationType,
} from "../validation";
import { attributeValue, escapeAttributeValue, escapeElementText, setAttribute, startTagAt } from "../xmlSplice";
import { parseCfSqref, X14_NS, XM_NS } from "./conditionalFormats";
import { childNodes, elementTree, nodeAttr, type XmlNode } from "./elementTree";
import { isSpreadsheetMl } from "./styles";

/** The worksheet ext holding the x14 data validations. */
export const DV_EXT_URI = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

/** What the reader cuts out of a worksheet part for this module. */
export interface WorksheetDvSource {
  /** The worksheet's root start tag — the namespaces in scope. */
  rootTag: string;
  /** The root-level `<dataValidations>`, when the part has one. */
  block?: string;
  /** The root-level `<extLst>`, when the part has one. */
  extLst?: string;
}

/** A worksheet's validations as read. */
export interface WorksheetValidations {
  /** The base block's rules, then the x14 ext's. */
  validations: DataValidation[];
  /** Either container's `disablePrompts`. */
  promptsDisabled: boolean;
}

const inMain = (ns: string | undefined): boolean => isSpreadsheetMl(ns);
const inX14 = (ns: string | undefined): boolean => ns === X14_NS;
const inXm = (ns: string | undefined): boolean => ns === XM_NS;

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

const TYPES: ReadonlySet<string> = new Set<ValidationType>(["whole", "decimal", "list", "date", "time", "textLength", "custom"]);
const OPERATORS: ReadonlySet<string> = new Set<ValidationOperator>([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
]);

/** Whether an ext element is the data-validation one. */
function isDvExt(node: XmlNode): boolean {
  return node.local === "ext" && inMain(node.ns) && (nodeAttr(node, "uri") ?? "").toUpperCase() === DV_EXT_URI.toUpperCase();
}

// --- reading ------------------------------------------------------------------------------

/** A formula element's canonical text, or undefined when absent or empty. */
function formulaText(node: XmlNode | undefined): string | undefined {
  const text = node?.text.trim();
  return text === undefined || text === "" ? undefined : fromXlsxFormula(text);
}

/** A rule from one `dataValidation` element (base or x14) and its parts;
 * null when its sqref names nothing. */
function readRule(node: XmlNode, sourceXml: string, ranges: GridRange[], formula1: string | undefined, formula2: string | undefined, x14: boolean, mint: () => string): DataValidation | null {
  if (ranges.length === 0) {
    return null;
  }
  const type = nodeAttr(node, "type") ?? "none";
  const operator = nodeAttr(node, "operator") ?? "between";
  const style = nodeAttr(node, "errorStyle");
  const rule: DataValidation = {
    id: mint(),
    ranges,
    type: TYPES.has(type) ? (type as ValidationType) : "any",
    operator: OPERATORS.has(operator) ? (operator as ValidationOperator) : "between",
    allowBlank: isOn(nodeAttr(node, "allowBlank")),
    showDropdown: !isOn(nodeAttr(node, "showDropDown")),
    input: { show: isOn(nodeAttr(node, "showInputMessage")), title: nodeAttr(node, "promptTitle") ?? "", text: nodeAttr(node, "prompt") ?? "" },
    error: {
      show: isOn(nodeAttr(node, "showErrorMessage")),
      style: style === "warning" || style === "information" ? (style as ValidationErrorStyle) : "stop",
      title: nodeAttr(node, "errorTitle") ?? "",
      text: nodeAttr(node, "error") ?? "",
    },
    sourceXml,
  };
  if (formula1 !== undefined) {
    rule.formula1 = formula1;
  }
  if (formula2 !== undefined) {
    rule.formula2 = formula2;
  }
  if (x14) {
    rule.x14 = true;
  }
  return rule;
}

/** A base `<dataValidation>`. */
function readBaseRule(node: XmlNode, text: string, mint: () => string = mintValidationId): DataValidation | null {
  const formula = (local: string): string | undefined => formulaText(childNodes(node, local, inMain)[0]);
  return readRule(node, text.slice(node.start, node.end), parseCfSqref(nodeAttr(node, "sqref") ?? ""), formula("formula1"), formula("formula2"), false, mint);
}

/** An `<x14:dataValidation>`. */
function readX14Rule(node: XmlNode, text: string, mint: () => string = mintValidationId): DataValidation | null {
  const formula = (local: string): string | undefined => {
    const holder = childNodes(node, local, inX14)[0];
    return formulaText(holder ? childNodes(holder, "f", inXm)[0] : undefined);
  };
  const sqref = childNodes(node, "sqref", inXm)[0]?.text ?? "";
  return readRule(node, text.slice(node.start, node.end), parseCfSqref(sqref), formula("formula1"), formula("formula2"), true, mint);
}

/** The x14 `dataValidations` lists of an `<extLst>` (parsed in `context`). */
function x14Lists(extLst: string, context: readonly string[]): XmlNode[] {
  const list = elementTree(extLst, context)?.[0];
  if (!list) {
    return [];
  }
  return childNodes(list, "ext", inMain)
    .filter(isDvExt)
    .flatMap((ext) => childNodes(ext, "dataValidations", inX14));
}

/** A worksheet's validations from what the reader cut out of its part (see
 * the module header). `mint` names each rule (default mintValidationId;
 * SPEC-137's worksheet scan passes placeholders the main thread re-mints). */
export function readValidationSource(source: WorksheetDvSource, mint: () => string = mintValidationId): WorksheetValidations {
  const context = [source.rootTag];
  const validations: DataValidation[] = [];
  let promptsDisabled = false;
  if (source.block !== undefined) {
    const block = elementTree(source.block, context)?.find((node) => node.local === "dataValidations" && inMain(node.ns));
    if (block) {
      promptsDisabled = isOn(nodeAttr(block, "disablePrompts"));
      for (const node of childNodes(block, "dataValidation", inMain)) {
        const rule = readBaseRule(node, source.block, mint);
        if (rule) {
          validations.push(rule);
        }
      }
    }
  }
  if (source.extLst !== undefined) {
    for (const list of x14Lists(source.extLst, context)) {
      promptsDisabled ||= isOn(nodeAttr(list, "disablePrompts"));
      for (const node of childNodes(list, "dataValidation", inX14)) {
        const rule = readX14Rule(node, source.extLst, mint);
        if (rule) {
          validations.push(rule);
        }
      }
    }
  }
  return { validations, promptsDisabled };
}

/** The root start tag, the first root-level `<dataValidations>` and the
 * root `<extLst>` of a whole worksheet part, cut out in one scanner pass. */
export function worksheetDvSource(worksheetXml: string): WorksheetDvSource {
  const source: WorksheetDvSource = { rootTag: "" };
  let depth = 0;
  let open: { local: "dataValidations" | "extLst"; start: number } | null = null;
  const take = (local: "dataValidations" | "extLst", start: number, end: number): void => {
    if (local === "dataValidations") {
      source.block ??= worksheetXml.slice(start, end);
    } else {
      source.extLst ??= worksheetXml.slice(start, end);
    }
  };
  scanXml(worksheetXml, (event) => {
    if (event.type === "open") {
      const level = depth + 1;
      if (level === 1) {
        source.rootTag = worksheetXml.slice(event.start, event.end);
      } else if (level === 2 && inMain(event.ns) && (event.local === "dataValidations" || event.local === "extLst")) {
        if (event.selfClosing) {
          take(event.local, event.start, event.end);
        } else {
          open = { local: event.local, start: event.start };
        }
      }
      if (!event.selfClosing) {
        depth = level;
      }
    } else if (event.type === "close") {
      if (depth === 2 && open !== null) {
        take(open.local, open.start, event.end);
        open = null;
      }
      depth -= 1;
    }
  });
  return source;
}

/** A whole worksheet part's validations (see the module header) — what a
 * reader upgrade parses from a dirty sheet's retained part. */
export function readWorksheetValidations(worksheetXml: string): WorksheetValidations {
  return readValidationSource(worksheetDvSource(worksheetXml));
}

// --- writing ------------------------------------------------------------------------------

/** What the writer needs from the save around it. */
export interface DvWriteContext {
  /** The part's main-namespace prefix ("" or "x:"). */
  prefix: string;
  /** Set when the part has a retained source: a rule's `sourceXml` may be
   * re-emitted — through `pass`, the sheet-name pass — in `context` (the
   * root start tag). */
  verbatim?: { pass(xml: string): string; context: readonly string[] };
}

/** The data-validation pieces of the part being written into. */
export interface DvPartSource {
  /** The root-level `<dataValidations>` as the part holds it. */
  block?: string;
  /** The extLst's `{CCE6A557-…}` ext as the part holds it. */
  ext?: string;
}

/** The writer's output. */
export interface DvWriteOutput {
  /** The part's `<dataValidations>` — the source's own text while its rules
   * are unchanged; "" for none. */
  block: string;
  /** The `{CCE6A557-…}` ext: undefined leaves the part's as it is (or
   * none); "" drops it; else the element. */
  ext: string | undefined;
}

/** An attribute value as written: XML-escaped, line breaks and tabs as
 * character references (a literal one would be normalised to a space). */
function attrText(value: string): string {
  return escapeAttributeValue(value).replace(/\r/g, "&#13;").replace(/\n/g, "&#10;").replace(/\t/g, "&#9;");
}

const attr = (name: string, value: string): string => (value === "" ? "" : ` ${name}="${attrText(value)}"`);
const flag = (name: string, on: boolean): string => (on ? ` ${name}="1"` : "");

/** A rule's attributes but sqref, in Excel's order, defaults omitted. */
function ruleAttributes(rule: DataValidation): string {
  return (
    (rule.type === "any" ? "" : attr("type", rule.type)) +
    (rule.error.style === "stop" ? "" : attr("errorStyle", rule.error.style)) +
    (rule.operator === "between" ? "" : attr("operator", rule.operator)) +
    flag("allowBlank", rule.allowBlank) +
    flag("showDropDown", !rule.showDropdown) +
    flag("showInputMessage", rule.input.show) +
    flag("showErrorMessage", rule.error.show) +
    attr("errorTitle", rule.error.title) +
    attr("error", rule.error.text) +
    attr("promptTitle", rule.input.title) +
    attr("prompt", rule.input.text)
  );
}

const sqrefText = (ranges: readonly GridRange[]): string => ranges.map(formatRange).join(" ");

/** A rule as a base `<dataValidation>` (self-closing without formulas, as
 * Excel writes one). */
export function baseValidationXml(rule: DataValidation, p: string): string {
  const formula = (local: string, text: string | undefined): string =>
    text === undefined ? "" : `<${p}${local}>${escapeElementText(toXlsxFormula(text))}</${p}${local}>`;
  const children = formula("formula1", rule.formula1) + formula("formula2", rule.formula2);
  const head = `<${p}dataValidation${ruleAttributes(rule)}${attr("sqref", sqrefText(rule.ranges))}`;
  return children === "" ? `${head}/>` : `${head}>${children}</${p}dataValidation>`;
}

/** An x14 rule's `<x14:formula1|2>`: the canonical text as `<xm:f>`. */
function x14Formula(local: string, text: string | undefined): string {
  return text === undefined ? "" : `<x14:${local}><xm:f>${escapeElementText(toXlsxFormula(text))}</xm:f></x14:${local}>`;
}

/** A rule as an `<x14:dataValidation>`. */
export function x14ValidationXml(rule: DataValidation): string {
  return (
    `<x14:dataValidation${ruleAttributes(rule)}>${x14Formula("formula1", rule.formula1)}${x14Formula("formula2", rule.formula2)}` +
    `<xm:sqref>${sqrefText(rule.ranges)}</xm:sqref></x14:dataValidation>`
  );
}

/** A start tag made an open tag (`<a/>` → `<a>`). */
function opened(tag: string): string {
  return tag.endsWith("/>") ? `${tag.slice(0, -2).trimEnd()}>` : tag;
}

/** `tag` with `disablePrompts` following `want`: left as written when it
 * already says so. */
function withPromptsFlag(tag: string, want: boolean): string {
  return isOn(attributeValue(tag, "disablePrompts")) === want ? tag : setAttribute(tag, "disablePrompts", want ? "1" : null);
}

/** An element's name and start tag (the first element of `xml`, read in
 * `context`). */
interface SourceElement {
  node: XmlNode;
  tag: string;
}

function sourceElement(xml: string | undefined, context: readonly string[], local: string, inNs: (ns: string | undefined) => boolean): SourceElement | undefined {
  if (xml === undefined) {
    return undefined;
  }
  const node = elementTree(xml, context)?.find((each) => each.local === local && inNs(each.ns));
  return node ? { node, tag: startTagAt(xml, node.start) } : undefined;
}

/** Declarations the written ext puts in scope. */
const X14_DECLARATIONS = `<_ xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}">`;

/** Whether `text` is one `dataValidation` element of the namespace `inNs`
 * accepts, read in `context` — a source text that fits where it goes. */
function fits(text: string, context: readonly string[], inNs: (ns: string | undefined) => boolean): boolean {
  const nodes = elementTree(text, context);
  const node = nodes?.[0];
  return nodes?.length === 1 && node !== undefined && node.local === "dataValidation" && inNs(node.ns) && node.start === 0 && node.end === text.length;
}

/** One rule as it will be written. */
interface Placed {
  x14: boolean;
  xml: string;
  /** Set when the text is the rule's source, re-emitted. */
  verbatim: boolean;
}

/** Whether `placed` are exactly `source`'s rules: each re-emitted verbatim,
 * in order. */
function sameRules(placed: readonly Placed[], source: readonly string[]): boolean {
  return placed.length === source.length && placed.every((piece, index) => piece.verbatim && piece.xml === source[index]);
}

/** The model's validations as the part's `<dataValidations>` block and
 * `{CCE6A557-…}` ext (see the module header). `part` is undefined for a part
 * written from scratch. */
export function writeDataValidations(
  sheet: Pick<Sheet, "name" | "validations" | "validationPromptsDisabled">,
  part: DvPartSource | undefined,
  ctx: DvWriteContext,
): DvWriteOutput {
  const p = ctx.prefix;
  const rules = sheet.validations as DataValidation[] | undefined;
  if (rules === undefined) {
    // A sheet stored before SPEC-130 that no reader upgrade has reached.
    return { block: part?.block ?? "", ext: undefined };
  }
  const disabled = sheet.validationPromptsDisabled === true;
  const verbatim = ctx.verbatim;
  const rootContext = verbatim?.context ?? [];

  // The source pieces: the block's and the ext's start tags (the context
  // their rules' text is read in) and the rules they hold.
  const block = verbatim ? sourceElement(part?.block, rootContext, "dataValidations", inMain) : undefined;
  const ext = verbatim ? sourceElement(part?.ext, rootContext, "ext", inMain) : undefined;
  const list = ext ? ext.node.children.find((child) => child.local === "dataValidations" && inX14(child.ns)) : undefined;
  const listTag = list && part?.ext !== undefined ? startTagAt(part.ext, list.start) : undefined;
  const baseContext = block ? [...rootContext, opened(block.tag)] : rootContext;
  const x14Context = [...rootContext, X14_DECLARATIONS, ...(ext ? [opened(ext.tag)] : []), ...(listTag ? [opened(listTag)] : [])];
  const sourceBase = block && part?.block !== undefined ? childNodes(block.node, "dataValidation", inMain).map((node) => part.block!.slice(node.start, node.end)) : [];
  const sourceX14 = list && part?.ext !== undefined ? childNodes(list, "dataValidation", inX14).map((node) => part.ext!.slice(node.start, node.end)) : [];
  // Only the rules the reader keeps count: one whose sqref names nothing
  // stays where it is while nothing changes.
  const readable = (texts: string[], x14: boolean, context: readonly string[]): string[] =>
    texts.filter((text) => {
      const node = elementTree(text, context)?.[0];
      return node !== undefined && (x14 ? readX14Rule(node, text) : readBaseRule(node, text)) !== null;
    });

  const placed: Placed[] = [];
  for (const rule of rules) {
    if (rule.ranges.length === 0) {
      continue;
    }
    if (verbatim && rule.sourceXml !== undefined) {
      const x14 = rule.x14 === true;
      const text = verbatim.pass(rule.sourceXml);
      if (fits(text, x14 ? x14Context : baseContext, x14 ? inX14 : inMain)) {
        placed.push({ x14, xml: text, verbatim: true });
        continue;
      }
    }
    const x14 = writesToX14(rule, sheet.name);
    placed.push({ x14, xml: x14 ? x14ValidationXml(rule) : baseValidationXml(rule, p), verbatim: false });
  }
  const base = placed.filter((piece) => !piece.x14);
  const x14 = placed.filter((piece) => piece.x14);

  // The block.
  let blockXml: string;
  const sourceBlockOn = block ? isOn(nodeAttr(block.node, "disablePrompts")) : false;
  if (part?.block !== undefined && block && sameRules(base, readable(sourceBase, false, baseContext)) && (base.length === 0 || sourceBlockOn === disabled)) {
    blockXml = part.block;
  } else if (base.length === 0) {
    blockXml = "";
  } else {
    const tag = withPromptsFlag(setAttribute(opened(block?.tag ?? `<${p}dataValidations>`), "count", String(base.length)), disabled);
    blockXml = `${tag}${base.map((piece) => piece.xml).join("")}</${block?.node.name ?? `${p}dataValidations`}>`;
  }

  // The ext.
  let extXml: string | undefined;
  const sourceListOn = list ? isOn(nodeAttr(list, "disablePrompts")) : false;
  if (x14.length === 0 && part?.ext === undefined) {
    extXml = undefined;
  } else if (ext && sameRules(x14, readable(sourceX14, true, x14Context)) && (x14.length === 0 || sourceListOn === disabled)) {
    extXml = undefined;
  } else if (x14.length === 0) {
    extXml = "";
  } else {
    let extTag = opened(ext?.tag ?? `<${p}ext uri="${DV_EXT_URI}">`);
    if (!/\sxmlns:x14\s*=/.test(extTag)) {
      extTag = setAttribute(extTag, "xmlns:x14", X14_NS);
    }
    let tag = opened(listTag ?? `<x14:dataValidations>`);
    tag = withPromptsFlag(setAttribute(tag, "count", String(x14.length)), disabled);
    if (!/\sxmlns:xm\s*=/.test(extTag) && !/\sxmlns:xm\s*=/.test(tag)) {
      tag = setAttribute(tag, "xmlns:xm", XM_NS);
    }
    extXml = `${extTag}${tag}${x14.map((piece) => piece.xml).join("")}</${list?.name ?? "x14:dataValidations"}></${ext?.node.name ?? `${p}ext`}>`;
  }
  return { block: blockXml, ext: extXml };
}
