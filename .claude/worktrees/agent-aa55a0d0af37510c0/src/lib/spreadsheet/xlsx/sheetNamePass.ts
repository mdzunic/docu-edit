// The retained-XML sheet-reference pass for .xlsx parts (SPEC-121 "Sheet
// references in formulas", retained side): rewrites renamed and deleted
// sheet names in exactly the table's locations, by offset splices located
// with core/xmlScanner.ts, every other byte verbatim.
//
// | Part | Location |
// | worksheets | `c/f` text; `cfRule/formula`; `dataValidation/formula1`, `formula2`; `x14:*/xm:f`; `hyperlink@location` |
// | xl/workbook.xml | `definedName` text (the writer's workbook rewrite calls remapXlsxFormula) |
// | xl/charts/*.xml, chartex | `c:f`, `cx:f` text |
// | xl/tables/*.xml | `calculatedColumnFormula`, `totalsRowFormula` |
// | xl/pivotCache/pivotCacheDefinition*.xml | `worksheetSource@sheet` (renames only) |
//
// Formula text goes through fromXlsxFormula → remapSheetReferences →
// toXlsxFormula; a location whose remapped canonical text equals its
// original is not touched. An empty remap skips the pass, and a part
// containing no case-insensitive occurrence of any remapped name is emitted
// without being scanned — or decoded: mentionsAnyName searches the bytes.
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import { isEmptyRemap, remapSheetReferences, type SheetRemap } from "../formula/sheetRefs";
import { applySplices, decodeXmlText, escapeElementText, findAttribute, startTagAt, type Splice } from "../xmlSplice";

/** The names a remap rewrites: its old names and deleted names. */
export function remappedNames(remap: SheetRemap): string[] {
  return [...Object.keys(remap.renamed), ...remap.deleted];
}

/** The longest stretch of a name that XML escaping and formula quoting can
 * never alter — what must appear verbatim wherever the name is written. */
function verbatimRun(name: string): string {
  return name.split(/['"&<>]/).reduce((longest, run) => (run.length > longest.length ? run : longest), "");
}

const utf8Encoder = new TextEncoder();
const utf8Decoder = new TextDecoder("utf-8");

/** An ASCII byte lower-cased. */
function lower(byte: number): number {
  return byte >= 0x41 && byte <= 0x5a ? byte + 32 : byte;
}

/** ASCII case-insensitive search of `needle` in `bytes`. */
function includesAsciiFolded(bytes: Uint8Array, needle: Uint8Array): boolean {
  const first = lower(needle[0]!);
  outer: for (let i = 0; i + needle.length <= bytes.length; i += 1) {
    if (lower(bytes[i]!) !== first) {
      continue;
    }
    for (let j = 1; j < needle.length; j += 1) {
      if (lower(bytes[i + j]!) !== lower(needle[j]!)) {
        continue outer;
      }
    }
    return true;
  }
  return false;
}

/** Whether a part may mention any of `names` (case-insensitively). Pure
 * ASCII names are searched in the bytes; a name with other letters
 * decodes the part (case folding needs the characters). */
export function mentionsAnyName(bytes: Uint8Array, names: readonly string[]): boolean {
  let text: string | undefined;
  for (const name of names) {
    const run = verbatimRun(name);
    if (run === "") {
      return true;
    }
    if (/^[\x20-\x7e]+$/.test(run)) {
      if (includesAsciiFolded(bytes, utf8Encoder.encode(run))) {
        return true;
      }
    } else {
      text ??= utf8Decoder.decode(bytes).toUpperCase();
      if (text.includes(run.toUpperCase())) {
        return true;
      }
    }
  }
  return false;
}

/** The same check over text already in memory. */
export function textMentionsAnyName(text: string, names: readonly string[]): boolean {
  return names.some((name) => {
    const run = verbatimRun(name);
    return run === "" || new RegExp(run.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i").test(text);
  });
}

/** An xlsx formula text remapped; the identical string when untouched. */
export function remapXlsxFormula(wire: string, remap: SheetRemap): string {
  const canonical = fromXlsxFormula(wire);
  const next = remapSheetReferences(canonical, remap);
  return next === canonical ? wire : toXlsxFormula(next);
}

/** Splices for every element named in `locals`: its text through
 * remapXlsxFormula. */
function formulaElementSplices(xml: string, locals: ReadonlySet<string>, remap: SheetRemap): Splice[] {
  const splices: Splice[] = [];
  let open: { local: string; end: number } | null = null;
  scanXml(xml, (event) => {
    if (event.type === "open" && locals.has(event.local) && !event.selfClosing) {
      open = { local: event.local, end: event.end };
    } else if (event.type === "close" && open && event.local === open.local) {
      const raw = xml.slice(open.end, event.start);
      open = null;
      if (raw.includes("<")) {
        return;
      }
      const wire = decodeXmlText(raw);
      const next = remapXlsxFormula(wire, remap);
      if (next !== wire) {
        splices.push({ start: event.start - raw.length, end: event.start, text: escapeElementText(next) });
      }
    }
  });
  return splices;
}

/** A splice of one attribute's value in the start tag at `start`. */
function attributeSplice(xml: string, start: number, qname: string, value: string): Splice | undefined {
  const tag = startTagAt(xml, start);
  const span = findAttribute(tag, qname);
  if (!span) {
    return undefined;
  }
  const escaped = value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  return { start: start + span.valueStart, end: start + span.valueEnd, text: escaped };
}

const WORKSHEET_FORMULAS: ReadonlySet<string> = new Set(["f", "formula", "formula1", "formula2"]);
const CHART_FORMULAS: ReadonlySet<string> = new Set(["f"]);
const TABLE_FORMULAS: ReadonlySet<string> = new Set(["calculatedColumnFormula", "totalsRowFormula"]);

/** A worksheet part with the pass applied. */
export function passWorksheet(xml: string, remap: SheetRemap): string {
  if (isEmptyRemap(remap)) {
    return xml;
  }
  const splices = formulaElementSplices(xml, WORKSHEET_FORMULAS, remap);
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === "hyperlink") {
      const location = scanAttr(event.attrs, "location");
      if (location !== undefined) {
        const next = remapSheetReferences(location, remap);
        const splice = next === location ? undefined : attributeSplice(xml, event.start, attributeQName(xml, event.start, "location"), next);
        if (splice) {
          splices.push(splice);
        }
      }
    }
  });
  return applySplices(xml, splices);
}

/** The qualified name an unprefixed attribute is written with. */
function attributeQName(_xml: string, _start: number, local: string): string {
  return local;
}

/** A chart (or chartex) part with the pass applied. */
export function passChart(xml: string, remap: SheetRemap): string {
  return isEmptyRemap(remap) ? xml : applySplices(xml, formulaElementSplices(xml, CHART_FORMULAS, remap));
}

/** A table part with the pass applied. */
export function passTable(xml: string, remap: SheetRemap): string {
  return isEmptyRemap(remap) ? xml : applySplices(xml, formulaElementSplices(xml, TABLE_FORMULAS, remap));
}

/** A pivot cache definition with `worksheetSource@sheet` renamed (a deleted
 * source sheet is left as it is). */
export function passPivotCache(xml: string, remap: SheetRemap): string {
  const renamed = new Map(Object.entries(remap.renamed).map(([from, to]) => [from.toUpperCase(), to] as const));
  if (renamed.size === 0) {
    return xml;
  }
  const splices: Splice[] = [];
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === "worksheetSource") {
      const sheet = scanAttr(event.attrs, "sheet");
      const next = sheet === undefined ? undefined : renamed.get(sheet.toUpperCase());
      const splice = next === undefined ? undefined : attributeSplice(xml, event.start, "sheet", next);
      if (splice) {
        splices.push(splice);
      }
    }
  });
  return applySplices(xml, splices);
}

/** Which pass a retained part takes, by its name; undefined for a part the
 * pass never touches. */
export function passFor(partName: string): ((xml: string, remap: SheetRemap) => string) | undefined {
  if (/^xl\/(?:worksheets|dialogsheets|macrosheets)\/[^/]+\.xml$/.test(partName)) {
    return passWorksheet;
  }
  if (/^xl\/charts\/[^/]+\.xml$/.test(partName) && !/\/(?:style|colors)\d*\.xml$/.test(partName)) {
    return passChart;
  }
  if (/^xl\/tables\/[^/]+\.xml$/.test(partName)) {
    return passTable;
  }
  if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(partName)) {
    return passPivotCache;
  }
  return undefined;
}
