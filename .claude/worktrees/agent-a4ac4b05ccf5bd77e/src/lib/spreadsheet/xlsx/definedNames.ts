// The xlsx wire form of defined names (SPEC-133 "Names on the wire",
// "Reserved names"; names.ts is the model's normative source).
//
// Reading `workbook.xml` `definedNames/definedName` → `DefinedName[]`, in
// source order:
// - `localSheetId` is the 0-based position in `<sheets>` (hidden and chart
//   sheets counted) → `scopeSheetId`, the model sheet at that source
//   position. One naming no sheet (out of range, or not a number) keeps the
//   attribute verbatim in `sourceAttrs`, the name hidden; it never resolves
//   and is written back as read.
// - `hidden` "1"/"true" → `hidden`; `comment` → `comment`; every other
//   attribute verbatim, with its leading whitespace, in source order, in
//   `sourceAttrs`.
// - The text (XML-decoded) through fromXlsxFormula then structuredFromXlsx,
//   checked with parseFormulaAt relative to A1; text that does not parse is
//   kept verbatim with `formulaUnparsed`. `_xlnm.*` names are kept spelled as
//   in the file; `_xlfn.`/`_xlpm.` placeholders are read like any name (their
//   text — `#NAME?` — parses) and written back as read.
//
// Writing `<definedNames>` from `model.names` (writer.ts, whenever the
// workbook part is rebuilt): model order — source order, new names
// appended, except where the model places them as SPEC-121 and SPEC-128
// always wrote them (names.ts setReservedName puts a new
// `_xlnm._FilterDatabase` first; sheets.ts duplicateSheet puts a copy's
// local name right after its origin's); a name scoped to a sheet no longer in the model is not written;
// zero names write no `<definedNames>` (one the source has is left empty, as
// SPEC-128 leaves it). Each name is matched to the source
// element it was read from (same name, case-insensitively, and the same
// scope — a scoped name through its sheet's source position, an orphan by
// its raw `localSheetId`); a matched name keeps the source start tag — so
// the source's attribute order and every unmodelled attribute — with only
// `name`, `localSheetId` (recomputed from the scope sheet's CURRENT index),
// `hidden` (an orphan's left as read) and `comment` set where the model
// differs, and keeps the source text while the model's formula is still what
// that text reads as (an unparsed one: the same text, after the save's sheet
// remap). An unmatched name — new, a copy's, or one whose source is gone —
// is written `name`, `comment`, `localSheetId`, `hidden`, then its
// `sourceAttrs`, the text toXlsxFormula(structuredToXlsx(formula)) (an
// unparsed one verbatim), escaped. Each matched element keeps the whitespace
// before it, and the block its trailing whitespace, so an unchanged name
// list re-emits the source bytes.
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import { FormulaSyntaxError, parseFormulaAt } from "../formula/parse";
import { isEmptyRemap, remapSheetReferences, type SheetRemap } from "../formula/sheetRefs";
import type { DefinedName, Sheet } from "../model";
import { isOrphanName } from "../names";
import { copyRemap } from "../sourceNames";
import { structuredFromXlsx, structuredToXlsx } from "../tables";
import { attributeValue, escapeAttributeValue, escapeElementText, setAttribute, type Splice } from "../xmlSplice";
import { partTextOf, readSourceWorkbook, type PartLookup } from "./package";
import { remapXlsxFormula } from "./sheetNamePass";

/** One source `<definedName>`. */
export interface SourceDefinedNameElement {
  /** Offset of the whitespace run before the element. */
  wsStart: number;
  start: number;
  end: number;
  /** The start tag as written. */
  tag: string;
  /** The raw text between the tags ("" when self-closing). */
  raw: string;
  /** The decoded text, as the scanner reports it. */
  text: string;
  /** The close tag ("" when self-closing). */
  closeTag: string;
  name: string | undefined;
  /** `localSheetId` trimmed, when present. */
  localSheetId: string | undefined;
}

/** Where the workbook part keeps its names. */
export interface DefinedNamesLayout {
  /** The root's element prefix (`x:` or ""). */
  prefix: string;
  /** The first `<definedNames>`: its range, start tag end and close tag start. */
  container: { start: number; end: number; tagEnd: number; closeStart: number; selfClosing: boolean } | null;
  elements: SourceDefinedNameElement[];
  /** Where a new `<definedNames>` goes: before `calcPr` or a later CT_Workbook child, else the root's end. */
  insertAt: number;
}

/** CT_Workbook's children after `definedNames`. */
const AFTER_DEFINED_NAMES = new Set(["calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"]);

/** The workbook part's `<definedNames>` and its elements, located. */
export function scanDefinedNames(xml: string): DefinedNamesLayout {
  const layout: DefinedNamesLayout = { prefix: "", container: null, elements: [], insertAt: -1 };
  let depth = 0;
  let rootClose = xml.length;
  let inContainer = false;
  let open: { start: number; tagEnd: number; tag: string; text: string; attrs: Parameters<typeof scanAttr>[0] } | null = null;
  let previousEnd = 0;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      depth += 1;
      if (depth === 1) {
        layout.prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
      } else if (depth === 2) {
        if (event.local === "definedNames" && layout.container === null) {
          layout.container = { start: event.start, end: event.end, tagEnd: event.end, closeStart: event.end, selfClosing: event.selfClosing };
          inContainer = !event.selfClosing;
          previousEnd = event.end;
        } else if (AFTER_DEFINED_NAMES.has(event.local) && layout.insertAt < 0) {
          layout.insertAt = event.start;
        }
      } else if (depth === 3 && inContainer && event.local === "definedName") {
        const tag = xml.slice(event.start, event.end);
        if (event.selfClosing) {
          push({ start: event.start, tagEnd: event.end, tag, text: "", attrs: event.attrs }, event.end, "");
        } else {
          open = { start: event.start, tagEnd: event.end, tag, text: "", attrs: event.attrs };
        }
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "text") {
      if (open !== null && depth === 3) {
        open.text += event.text;
      }
    } else {
      if (depth === 3 && open !== null && event.local === "definedName") {
        push(open, event.end, xml.slice(event.start, event.end));
        open = null;
      } else if (depth === 2 && inContainer && event.local === "definedNames") {
        layout.container!.end = event.end;
        layout.container!.closeStart = event.start;
        inContainer = false;
      } else if (depth === 1) {
        rootClose = event.start;
      }
      depth -= 1;
    }
  });
  if (layout.insertAt < 0) {
    layout.insertAt = rootClose;
  }
  return layout;

  function push(element: { start: number; tagEnd: number; tag: string; text: string; attrs: Parameters<typeof scanAttr>[0] }, end: number, closeTag: string): void {
    layout.elements.push({
      wsStart: previousEnd <= element.start && /^\s*$/.test(xml.slice(previousEnd, element.start)) ? previousEnd : element.start,
      start: element.start,
      end,
      tag: element.tag,
      raw: closeTag === "" ? "" : xml.slice(element.tagEnd, end - closeTag.length),
      text: element.text,
      closeTag,
      name: scanAttr(element.attrs, "name"),
      localSheetId: scanAttr(element.attrs, "localSheetId")?.trim(),
    });
    previousEnd = end;
  }
}

const MODELLED_ATTRIBUTES = new Set(["name", "localSheetId", "hidden", "comment"]);
const ATTRIBUTE = /(\s+)([^\s=/>]+)(\s*=\s*)("[^"]*"|'[^']*')/g;

/** A start tag's attributes, each verbatim with its leading whitespace. */
function rawAttributes(tag: string): { name: string; text: string }[] {
  const body = tag.replace(/^<[^\s/>]+/, "");
  return [...body.matchAll(ATTRIBUTE)].map((match) => ({ name: match[2]!, text: match[0] }));
}

const isTrue = (value: string | undefined): boolean => value === "1" || value === "true";

/** A name's wire text read as the model stores it: canonical when it parses
 * (relative to A1), else the text verbatim with `formulaUnparsed`. */
export function readNameFormula(wire: string): Pick<DefinedName, "formula" | "formulaUnparsed"> {
  try {
    const canonical = structuredFromXlsx(fromXlsxFormula(wire));
    if (!(parseFormulaAt(canonical, { row: 0, col: 0 }) instanceof FormulaSyntaxError)) {
      return { formula: canonical };
    }
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) {
      throw error;
    }
  }
  return { formula: wire, formulaUnparsed: true };
}

/** One source element as a DefinedName; `scopeOf(localSheetId)` is the
 * model sheet id at that source position (undefined: none). */
function nameFrom(element: SourceDefinedNameElement, scopeOf: (index: number) => string | undefined): DefinedName | undefined {
  if (element.name === undefined) {
    return undefined;
  }
  const out: DefinedName = { name: element.name, ...readNameFormula(element.text) };
  let sourceAttrs = "";
  let orphan = false;
  for (const attr of rawAttributes(element.tag)) {
    if (attr.name === "localSheetId") {
      const local = element.localSheetId ?? "";
      const scope = /^\d+$/.test(local) ? scopeOf(Number(local)) : undefined;
      if (scope === undefined) {
        sourceAttrs += attr.text;
        orphan = true;
      } else {
        out.scopeSheetId = scope;
      }
    } else if (!MODELLED_ATTRIBUTES.has(attr.name)) {
      sourceAttrs += attr.text;
    }
  }
  const comment = attributeValue(element.tag, "comment");
  if (comment !== undefined) {
    out.comment = comment;
  }
  if (orphan || isTrue(attributeValue(element.tag, "hidden"))) {
    out.hidden = true;
  }
  if (sourceAttrs !== "") {
    out.sourceAttrs = sourceAttrs;
  }
  return out;
}

/** The names of a workbook part; `sheetIds[i]` is the id of the model sheet
 * read from the i-th `<sheet>`. */
export function readXlsxDefinedNames(workbookXml: string, sheetIds: readonly (string | undefined)[]): DefinedName[] {
  if (!workbookXml.includes("definedName")) {
    return [];
  }
  return scanDefinedNames(workbookXml).elements.flatMap((element) => nameFrom(element, (index) => sheetIds[index]) ?? []);
}

/** The model sheet each source `<sheet>` is: an original holding its part. */
function sourceSheetIds(parts: PartLookup, sheets: readonly Sheet[]): (string | undefined)[] | undefined {
  const workbook = readSourceWorkbook(parts);
  if (!workbook) {
    return undefined;
  }
  return workbook.sheets.map((source) => sheets.find((sheet) => sheet.copiedFrom === undefined && sheet.partName !== undefined && sheet.partName === source.partName)?.id);
}

/** A parsed name's formula through `remap`. */
function remappedName(name: DefinedName, remap: SheetRemap): DefinedName {
  return !isEmptyRemap(remap) && !name.formulaUnparsed ? { ...name, formula: remapSheetReferences(name.formula, remap) } : name;
}

/** The retained workbook part's names mapped onto the model (SPEC-133
 * "Reader upgrade"): `localSheetId` through the SOURCE sheet order to the
 * model sheet holding that part — a name scoped to a sheet no longer present
 * dropped — and every parsed formula through `remap` (sourceSheetRemap). A
 * copy (SPEC-121's `copiedFrom`, the source sheet it copies) gets each name
 * that sheet scopes, right after the origin's (in its place when the origin
 * is gone), through the copy's own remap: the names SPEC-121's writer
 * duplicated at write time before this version, which the retained part
 * never held. For a step that adopts names from the retained source. */
export function readSourceDefinedNames(parts: PartLookup, sheets: readonly Sheet[], remap: SheetRemap = { renamed: {}, deleted: [], orderBefore: [] }): DefinedName[] {
  const workbook = readSourceWorkbook(parts);
  const ids = sourceSheetIds(parts, sheets);
  const xml = workbook === undefined ? undefined : partTextOf(parts, workbook.path);
  if (workbook === undefined || ids === undefined || xml === undefined || !xml.includes("definedName")) {
    return [];
  }
  const out: DefinedName[] = [];
  for (const element of scanDefinedNames(xml).elements) {
    const local = element.localSheetId;
    const index = local !== undefined && /^\d+$/.test(local) ? Number(local) : undefined;
    // A name of a deleted sheet is dropped; its copies still get theirs.
    const name = index !== undefined && index < ids.length && ids[index] === undefined ? undefined : nameFrom(element, (at) => ids[at]);
    if (name) {
      out.push(remappedName(name, remap));
    }
    const origin = index === undefined ? undefined : workbook.sheets[index]?.name.toUpperCase();
    for (const copy of origin === undefined ? [] : sheets.filter((sheet) => sheet.copiedFrom?.toUpperCase() === origin)) {
      const clone = nameFrom(element, () => copy.id);
      if (clone) {
        out.push(remappedName(clone, copyRemap(remap, copy)));
      }
    }
  }
  return out;
}

// --- writing --------------------------------------------------------------------------------

export interface DefinedNamesWriteContext {
  /** The written `<sheets>` position of a model sheet (undefined: not written). */
  indexOf(sheetId: string): number | undefined;
  /** The source position of the sheet a scoped name's source element named
   * (an original's), for matching; undefined for a sheet with no source. */
  sourceIndexOf?(sheetId: string): number | undefined;
  /** The save's sheet remap, applied to unparsed text (SPEC-121). */
  remap?: SheetRemap;
}

/** The wire text of a model name. */
function wireText(name: DefinedName, remap: SheetRemap | undefined): string {
  if (name.formulaUnparsed) {
    if (remap === undefined || isEmptyRemap(remap)) {
      return name.formula;
    }
    try {
      return remapXlsxFormula(name.formula, remap);
    } catch {
      return name.formula;
    }
  }
  return toXlsxFormula(structuredToXlsx(name.formula));
}

/** `sourceAttrs` safe for a part whose root binds no extra prefix: the
 * unprefixed attributes only. */
function plainAttributes(sourceAttrs: string | undefined): string {
  return sourceAttrs === undefined ? "" : rawAttributes(`<x${sourceAttrs}`).filter((attr) => !attr.name.includes(":")).map((attr) => attr.text).join("");
}

/** A `<definedName>` written from the model (see the module header). */
export function definedNameXml(name: DefinedName, index: number | undefined, prefix: string, options: { remap?: SheetRemap; prefixedAttrs?: boolean } = {}): string {
  const orphan = isOrphanName(name);
  const attrs =
    ` name="${escapeAttributeValue(name.name)}"` +
    (name.comment === undefined ? "" : ` comment="${escapeAttributeValue(name.comment).replace(/\n/g, "&#10;").replace(/\t/g, "&#9;")}"`) +
    (index === undefined || orphan ? "" : ` localSheetId="${index}"`) +
    (name.hidden ? ` hidden="1"` : "") +
    (options.prefixedAttrs === false ? plainAttributes(name.sourceAttrs) : (name.sourceAttrs ?? ""));
  return `<${prefix}definedName${attrs}>${escapeElementText(wireText(name, options.remap))}</${prefix}definedName>`;
}

/** A matched source element brought in line with the model name. */
function patchedElement(element: SourceDefinedNameElement, name: DefinedName, index: number | undefined, prefix: string, remap: SheetRemap | undefined): string {
  let tag = element.tag;
  if (element.name !== name.name) {
    tag = setAttribute(tag, "name", name.name);
  }
  if (!isOrphanName(name) && index !== undefined && element.localSheetId !== String(index)) {
    tag = setAttribute(tag, "localSheetId", String(index));
  }
  if (!isOrphanName(name) && isTrue(attributeValue(tag, "hidden")) !== (name.hidden === true)) {
    tag = setAttribute(tag, "hidden", name.hidden ? "1" : null);
  }
  if (attributeValue(tag, "comment") !== name.comment) {
    tag = setAttribute(tag, "comment", name.comment ?? null);
  }
  const sourceReads = name.formulaUnparsed ? element.text : readNameFormula(element.text);
  const same = name.formulaUnparsed ? sourceReads === name.formula : typeof sourceReads !== "string" && !sourceReads.formulaUnparsed && sourceReads.formula === name.formula;
  let body: string;
  if (same && !(name.formulaUnparsed && remap !== undefined && !isEmptyRemap(remap))) {
    body = element.raw;
  } else {
    body = escapeElementText(wireText(name, remap));
  }
  if (body === element.raw && tag === element.tag) {
    return element.closeTag === "" ? element.tag : element.tag + element.raw + element.closeTag;
  }
  const open = tag.endsWith("/>") ? tag.replace(/\s*\/>$/, ">") : tag;
  const close = element.closeTag === "" ? `</${/^<([^\s/>]+)/.exec(tag)?.[1] ?? `${prefix}definedName`}>` : element.closeTag;
  return `${open}${body}${close}`;
}

/** A name compared case-insensitively. */
const keyOf = (text: string | undefined): string => (text ?? "").toUpperCase();

/** Whether a source element's `localSheetId` is `local`: as numbers when
 * both are, else as written (an orphan's `x` matches only `x`). */
function sameLocalSheetId(source: string | undefined, local: string): boolean {
  if (source === undefined) {
    return false;
  }
  const numeric = /^\d+$/;
  return numeric.test(local) ? numeric.test(source) && Number(source) === Number(local) : source === local;
}

/** The splices bringing a workbook part's `<definedNames>` in line with
 * `names` (see the module header). */
export function definedNamesSplices(xml: string, layout: DefinedNamesLayout, names: readonly DefinedName[], ctx: DefinedNamesWriteContext): Splice[] {
  const p = layout.prefix;
  const claimed = new Set<SourceDefinedNameElement>();
  const findSource = (name: DefinedName): SourceDefinedNameElement | undefined => {
    let local: string | undefined;
    if (isOrphanName(name)) {
      local = /(?:^|\s)localSheetId\s*=\s*["']\s*([^"']*?)\s*["']/.exec(name.sourceAttrs ?? "")?.[1];
    } else if (name.scopeSheetId !== undefined) {
      const at = ctx.sourceIndexOf?.(name.scopeSheetId);
      if (at === undefined) {
        return undefined;
      }
      local = String(at);
    }
    return layout.elements.find((element) => !claimed.has(element) && keyOf(element.name) === keyOf(name.name) && (local === undefined ? element.localSheetId === undefined : sameLocalSheetId(element.localSheetId, local)));
  };
  const firstWs = layout.elements[0] ? xml.slice(layout.elements[0].wsStart, layout.elements[0].start) : "";
  let content = "";
  let written = 0;
  for (const name of names) {
    const index = name.scopeSheetId === undefined ? undefined : ctx.indexOf(name.scopeSheetId);
    if (name.scopeSheetId !== undefined && index === undefined) {
      continue; // its sheet is gone
    }
    const source = findSource(name);
    written += 1;
    if (source) {
      claimed.add(source);
      content += xml.slice(source.wsStart, source.start) + patchedElement(source, name, index, p, ctx.remap);
    } else {
      content += firstWs + definedNameXml(name, index, p, ctx.remap === undefined ? {} : { remap: ctx.remap });
    }
  }
  const container = layout.container;
  if (container === null) {
    return written === 0 ? [] : [{ start: layout.insertAt, end: layout.insertAt, text: `<${p}definedNames>${content}</${p}definedNames>` }];
  }
  if (written === 0 && container.selfClosing) {
    return [];
  }
  if (container.selfClosing) {
    const tag = xml.slice(container.start, container.end);
    const elementName = /^<([^\s/>]+)/.exec(tag)![1];
    return [{ start: container.start, end: container.end, text: `${tag.replace(/\s*\/>$/, ">")}${content}</${elementName}>` }];
  }
  const lastEnd = layout.elements.at(-1)?.end ?? container.tagEnd;
  const trailing = /^\s*$/.test(xml.slice(lastEnd, container.closeStart)) ? xml.slice(lastEnd, container.closeStart) : "";
  const next = content + trailing;
  return xml.slice(container.tagEnd, container.closeStart) === next ? [] : [{ start: container.tagEnd, end: container.closeStart, text: next }];
}

/** A from-scratch workbook's `<definedNames>` ("" for none). */
export function definedNamesXml(names: readonly DefinedName[], indexOf: (sheetId: string) => number | undefined): string {
  const elements = names.flatMap((name) => {
    const index = name.scopeSheetId === undefined ? undefined : indexOf(name.scopeSheetId);
    return name.scopeSheetId !== undefined && index === undefined ? [] : [definedNameXml(name, index, "", { prefixedAttrs: false })];
  });
  return elements.length === 0 ? "" : `<definedNames>${elements.join("")}</definedNames>`;
}
