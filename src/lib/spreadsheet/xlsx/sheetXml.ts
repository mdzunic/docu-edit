// The worksheet reader (SPEC-116 "Reading .xlsx", extended by SPEC-117
// "Reading xlsx sheet layout"): one worksheet part → its sparse rows, column
// spans, default sizes, merges and view settings, in ONE forward-only
// scanner pass (core/xmlScanner.ts) — a million-cell sheet is never scanned
// twice and never becomes a DOM. Also the string items both cell sources
// share: the shared-strings part and inline `<is>` strings.
//
// Cells (SPEC-116): rows and cells without `r` are placed sequentially; `t`
// absent/n, s, str, inlineStr, b, e, d; a formula's cached `<v>` is the
// value. SPEC-117: `s` → `styleId` (as written, even
// 0), and a `<c>` carrying a style but no value is stored as a style-only
// cell, counted against MAX_WORKBOOK_CELLS like any other. Shared-string and
// inline `<r>` runs carry their `<rPr>` as `font` (xlsx/styles.ts's font
// mapping); a plain `<t>` item stays `{ text }`. SPEC-118: a `t="d"` value
// whose style is General points at a derived style with the implied
// `yyyy-mm-dd` (+ ` hh:mm:ss`) format (impliedFormats.ts).
//
// Formulas (SPEC-119 "Reading formulas from .xlsx"): a `<f>` gives the cell
// `formula` = fromXlsxFormula(text), checked with parseFormulaAt — a formula
// that does not parse keeps its `<f>` text verbatim, `formulaUnparsed`. A
// formula cell is stored even with no `<v>` (`value: null`, no cached value).
// - `t="shared"`: the master (the `<f>` with text) reads as above; each
//   follower (`<f t="shared" si="…"/>`) is shiftFormula of the master's
//   canonical text by its offset from the master — an out-of-grid shift
//   yields `#REF!`, as Excel's own expansion does — and a follower met before
//   its master is resolved at the end of the sheet. A follower whose `si` is
//   never defined keeps its value and no formula. (The followers of a master
//   that does not parse get its wire text shifted, still unparsed.)
// - `t="array"`: the anchor's `formula` plus `arrayRange` (its `ref`; one
//   cell when absent); members keep their cached values only. With `c@cm`
//   (Excel's dynamic-array metadata) it is a dynamic array: `dynamicArray`.
// - `t="dataTable"`: the anchor's `dataTableXml` — the `<f>` element's
//   source text, sliced by the scanner's offsets. No formula.
// - `c@cm` → `cm` on any cell. `ca`, `aca`, `bx` and xl/calcChain.xml are
//   not read; the retained source keeps them.
//
// Layout and view (SPEC-117, only the root's own children are read):
// - `sheetFormatPr`: `defaultRowHeight` (absent element → 15 pt);
//   `defaultColWidth` → px through the MDW formula, else `baseColWidth`
//   (default 8) through baseColWidthToPx (Calibri 11: 64 px);
// - `cols/col`: 1-based `min`/`max` → 0-based, `width` → `widthPx` (with
//   `sourceWidthChars`), `hidden`, `style` → `styleId`; every other attribute
//   (`bestFit`, `collapsed`, `outlineLevel`, `phonetic`, …) verbatim in
//   `sourceAttrs` (`customWidth` is implied by a width). A later `col`
//   overlapping an earlier one is clipped to the gap. A span is stored only
//   with a width other than the default, `hidden` or a `styleId`;
// - `row`: `ht` → `heightPt` (honoured with or without `customHeight`),
//   `customHeight`, `hidden`, `s` → `styleId` only with `customFormat="1"`;
//   every other attribute (`spans`, `x14ac:dyDescent`, `outlineLevel`,
//   `collapsed`, `thickTop`, `thickBot`, …) verbatim in `sourceAttrs`. A row
//   with no cells is stored only with a height, `hidden` or a `styleId`;
// - `sheetPr/tabColor` → `tabColor`; `mergeCells/mergeCell@ref` → `merges`
//   (merges.ts acceptMerges: two or more cells, no overlap with an earlier
//   one);
// - `sheetViews/sheetView[0]`: `showGridLines`/`showZeros` "0" → false,
//   `zoomScale` (clamped 10–400; 100 → absent), `topLeftCell`; a `pane` whose
//   `state` is `frozen` or `frozenSplit` → `frozen` (ySplit rows, xSplit
//   columns) and its `topLeftCell`; a `split` pane (or one with no state) is
//   not a freeze. The `selection` of the active pane (`pane@activePane`,
//   default topLeft) → `selection`, normalised so the active cell lies in
//   its active range. SPEC-119: `showFormulas="1"` → `showFormulas`.
// Everything else — the split panes, the other panes' selections, outline
// levels, `zeroHeight` — stays in the retained part, untouched.
//
// SPEC-122 captures, for the writer below to re-emit: a `t="s"` cell's
// `sstIndex`; every `<c>` attribute but r, s, t and cm verbatim in
// `preservedAttrs` (` vm="2"`, ` ph="1"` — each with its leading space);
// every child but f, v and is verbatim in `preservedXml`, in source order. A
// cell holding only such content is stored, and counted against the ceiling
// like any other. `sheetProtection@sheet` "1"/"true" → `protection`.
//
// The worksheet writer (SPEC-122 "The worksheet writer", normative there):
// serializeWorksheet(sheet, base, ctx) rebuilds `dimension`, `sheetViews`,
// `sheetFormatPr`, `cols`, `sheetData` and `mergeCells` from the model and
// splices each over its source counterpart's exact range, located with the
// scanner (an absent one is inserted before the first source child that
// comes later in CT_Worksheet order — an `mc:AlternateContent` ranks as the
// element it wraps — and an empty `cols`/`mergeCells` is removed); the root
// start tag and every other element stay byte-verbatim. A new sheet (no
// base) starts from SPEC-121's empty-sheet template, reproduced byte for
// byte for an empty sheet.
// - dimension: `ref` = the used range, "A1" when empty.
// - sheetViews: the first sheetView's tabSelected, showGridLines,
//   showFormulas, zoomScale and topLeftCell follow the model (an attribute
//   already saying what the model says is left as written); its pane — a
//   freeze generated (xSplit/ySplit/topLeftCell/activePane/state="frozen"),
//   a source split pane kept, a source freeze the model no longer has
//   dropped — and the active pane's selection are regenerated, the other
//   panes' selections kept while their pane exists; pivotSelection, extLst
//   and further sheetViews verbatim.
// - sheetFormatPr: defaultRowHeight (+ customHeight="1") spliced when the
//   model's default differs from the source's; everything else verbatim.
// - cols: one `col` per ColumnSpan — 1-based min/max, `width`
//   (sourceWidthChars when present, else the smallest multiple of 1/256 the
//   reader's px formula maps to widthPx), `style`, `hidden`,
//   `customWidth="1"` with a width, then the span's sourceAttrs.
// - sheetData: `<row r [s customFormat="1"] [ht [customHeight="1"]]
//   [hidden="1"] {sourceAttrs less spans and x14ac:dyDescent}>` per stored
//   row; `<c r [s] [t] [cm] {preservedAttrs}>` per cell — `s` omitted for xf
//   0 (unless the row or column would otherwise lend the cell its style),
//   self-closing when childless, children `<f>`, `<v>`, then preservedXml.
//   number → no t; string → t="s" through the save's SharedStringsWriter
//   (inline strings become shared ones); a formula's or an array member's
//   string result → t="str" (encodeXlsxText); boolean → t="b" 1/0; error →
//   t="e". `<f>` is toXlsxFormula(formula) — the stored wire text when
//   formulaUnparsed; an array anchor writes `<f t="array" ref>`, its members
//   their cached `<v>` only; a dynamic array keeps its `cm`; a data table's
//   dataTableXml is its `<f>`, verbatim. Shared formulas are never written.
//   Numbers: `Object.is(n, -0) ? "0" : String(n)`.
// - mergeCells: `count` + one `mergeCell ref` per merge.
import type { DocumentTheme, ModelColor } from "../../core/theme";
import { scanAttr, scanXml, type XmlScanAttribute, type XmlScanEvent } from "../../core/xmlScanner";
import { colToLetters, formatA1, formatRange, parseA1, parseRange, parseSqref, usedRange, type CellRef, type GridRange } from "../address";
import { baseColWidthToPx, DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT, pxToXlsxWidth, xlsxWidthToPx } from "../geometry";
import { acceptMerges } from "../merges";
import {
  MAX_COLS,
  MAX_ROWS,
  toCellError,
  type Cell,
  type CellFont,
  type CellValue,
  type ColumnSpan,
  type Row,
  type Sheet,
  type SpreadsheetModel,
  type TextRun,
} from "../model";
import { impliedDateCode, type ImpliedFormatStyles } from "../impliedFormats";
import { readXlsxFormula, toXlsxFormula, type FormulaFields } from "../formula/dialect";
import { FormulaSyntaxError, parseFormulaAt } from "../formula/parse";
import { shiftFormula } from "../formula/references";
import { addressText, normaliseCells, normaliseRows, outOfGridError, parseError, type CellBudget } from "../parsed";
import type { SheetSelection } from "../selection";
import { odsDateValueToSerial } from "../serial";
import { columnSpanAt } from "../styles";
import { applySplices, attributeValue, escapeElementText, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { decodeXmlEscapes, encodeXlsxText, type SharedStringsWriter } from "./sharedStrings";
import { isSpreadsheetMl, XlsxFontReader, type XlsxStyleWriter } from "./styles";
import { modelColorToXlsxAttrs, scannedAttributes, xlsxColorFrom } from "./themeColor";

/** A shared string: its text, plus its runs when it was rich. */
export interface SharedString {
  text: string;
  runs?: TextRun[];
}

export { decodeXmlEscapes };

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
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

/** Collects the text of a string item — `<si>` or `<is>` — from the events
 * inside it: plain `<t>` text, or `<r>` runs with their `<rPr>` fonts;
 * phonetic runs are skipped. */
class StringItemReader {
  private plain = "";
  private runs: TextRun[] = [];
  private runText: string | null = null;
  private runFont: CellFont | undefined;
  /** Set while inside a run's `<rPr>`. */
  private fontReader: XlsxFontReader | null = null;
  private inText = false;
  private skipDepth = 0;

  constructor(private readonly indexedColors: readonly string[] | undefined) {}

  /** Feeds one event from inside the item (not the item's own tags). */
  feed(event: XmlScanEvent): void {
    if (this.skipDepth > 0) {
      if (event.type === "open" && !event.selfClosing) {
        this.skipDepth += 1;
      } else if (event.type === "close") {
        this.skipDepth -= 1;
      }
      return;
    }
    if (this.fontReader !== null) {
      if (event.type === "open") {
        if (isSpreadsheetMl(event.ns)) {
          this.fontReader.add(event.local, scannedAttributes(event.attrs));
        }
        if (!event.selfClosing) {
          this.skipDepth = 1; // a property's content is not modelled
        }
      } else if (event.type === "close") {
        const font = this.fontReader.build();
        this.fontReader = null;
        this.runFont = Object.keys(font).length > 0 ? font : undefined;
      }
      return;
    }
    if (event.type === "text") {
      if (this.inText) {
        if (this.runText !== null) {
          this.runText += event.text;
        } else {
          this.plain += event.text;
        }
      }
      return;
    }
    if (!isSpreadsheetMl(event.ns)) {
      if (event.type === "open" && !event.selfClosing) {
        this.skipDepth = 1;
      }
      return;
    }
    if (event.type === "open") {
      if (event.local === "rPh" || event.local === "phoneticPr") {
        if (!event.selfClosing) {
          this.skipDepth = 1;
        }
      } else if (event.local === "rPr") {
        if (!event.selfClosing) {
          if (this.runText !== null) {
            this.fontReader = new XlsxFontReader(this.indexedColors);
          } else {
            this.skipDepth = 1;
          }
        }
      } else if (event.local === "r") {
        this.runText = event.selfClosing ? null : "";
        this.runFont = undefined;
      } else if (event.local === "t") {
        this.inText = !event.selfClosing;
      }
      return;
    }
    if (event.local === "t") {
      this.inText = false;
    } else if (event.local === "r" && this.runText !== null) {
      const run: TextRun = { text: decodeXmlEscapes(this.runText) };
      if (this.runFont !== undefined) {
        run.font = this.runFont;
      }
      this.runs.push(run);
      this.runText = null;
      this.runFont = undefined;
    }
  }

  /** The item's string, with its runs when it had any. */
  finish(): SharedString {
    if (this.runs.length > 0) {
      return { text: this.runs.map((run) => run.text).join(""), runs: this.runs };
    }
    return { text: decodeXmlEscapes(this.plain) };
  }
}

/** Reads `xl/sharedStrings.xml` with the scanner. `indexedColors` is the
 * workbook's override palette, for run colours. */
export function readSharedStrings(xml: string, indexedColors?: readonly string[]): SharedString[] {
  const strings: SharedString[] = [];
  let item: StringItemReader | null = null;
  let depth = 0;
  let itemDepth = -1;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      if (item === null && isSpreadsheetMl(event.ns) && event.local === "si") {
        if (event.selfClosing) {
          strings.push({ text: "" });
          return;
        }
        item = new StringItemReader(indexedColors);
        itemDepth = depth + 1;
      } else if (item !== null) {
        item.feed(event);
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return;
    }
    if (event.type === "close") {
      if (item !== null && depth === itemDepth) {
        strings.push(item.finish());
        item = null;
      } else if (item !== null) {
        item.feed(event);
      }
      depth -= 1;
      return;
    }
    item?.feed(event);
  });
  return strings;
}

// --- cells ----------------------------------------------------------------------

/** One `<c>` being read. */
interface PendingCell {
  col: number;
  type: string | undefined;
  styleId: number | undefined;
  /** The `<v>` text — null when the cell had no `<v>`. */
  v: string | null;
  inlineString: StringItemReader | null;
  /** Nesting inside the cell: >0 while inside `<v>`, `<is>`, `<f>`, … */
  depth: number;
  inV: boolean;
  /** `c@cm` (SPEC-119). */
  cm: number | undefined;
  /** The cell's `<f>`: its attributes, text and source offsets. */
  f: PendingFormula | null;
  inF: boolean;
  /** SPEC-122: the unmodelled `<c>` attributes, verbatim with leading spaces. */
  preservedAttrs: string | undefined;
  /** SPEC-122: the unmodelled children so far, verbatim, and the start of
   * the one being read (-1 when none). */
  preservedXml: string;
  preservedStart: number;
}

/** One `<f>` being read (SPEC-119). */
interface PendingFormula {
  attrs: readonly XmlScanAttribute[];
  text: string;
  start: number;
  end: number;
}

/** A shared formula's master, for its followers (SPEC-119). */
interface SharedMaster {
  wire: string;
  read: FormulaFields;
  row: number;
  col: number;
}

/** A follower's formula: the master's canonical text shifted by the
 * follower's offset (parse-checked); an unparsed master's wire text shifted,
 * unparsed. */
function followerFormula(master: SharedMaster, row: number, col: number): FormulaFields {
  const dRow = row - master.row;
  const dCol = col - master.col;
  if (!master.read.formulaUnparsed) {
    const formula = shiftFormula(master.read.formula, dRow, dCol);
    if (!(parseFormulaAt(formula, { row, col }) instanceof FormulaSyntaxError)) {
      return { formula };
    }
  }
  return { formula: shiftFormula(master.wire, dRow, dCol), formulaUnparsed: true };
}

/** A cell value from its type and cached text; null for "no value".
 * `where` names the cell for an error — only built when one is thrown. */
function cellValue(
  pending: PendingCell,
  sharedStrings: readonly SharedString[],
  dateSystem: "1900" | "1904",
  where: () => string,
): { value: CellValue; runs?: TextRun[] } {
  const v = pending.v;
  switch (pending.type) {
    case undefined:
    case "n": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      const n = Number(v);
      if (!Number.isFinite(n)) {
        throw parseError(where(), `"${v}" is not a number`);
      }
      return { value: n };
    }
    case "s": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      const index = Number(v);
      const shared = Number.isInteger(index) ? sharedStrings[index] : undefined;
      if (!shared) {
        throw parseError(where(), `shared string ${v} does not exist`);
      }
      return shared.runs ? { value: shared.text, runs: shared.runs } : { value: shared.text };
    }
    case "str":
      return { value: v === null ? null : decodeXmlEscapes(v) };
    case "inlineStr": {
      if (pending.inlineString === null) {
        return { value: v === null ? null : decodeXmlEscapes(v) };
      }
      const inline = pending.inlineString.finish();
      return inline.runs ? { value: inline.text, runs: inline.runs } : { value: inline.text };
    }
    case "b": {
      const flag = v?.trim();
      if (flag === undefined || flag === "") {
        return { value: null };
      }
      if (flag === "1" || flag === "true") {
        return { value: true };
      }
      if (flag === "0" || flag === "false") {
        return { value: false };
      }
      throw parseError(where(), `"${v}" is not a boolean`);
    }
    case "e":
      return { value: v === null || v.trim() === "" ? null : toCellError(v) };
    case "d": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      try {
        return { value: odsDateValueToSerial(v, dateSystem) };
      } catch {
        throw parseError(where(), `"${v}" is not an ISO 8601 date`);
      }
    }
    default:
      throw parseError(where(), `cell type "${pending.type}" is not a SpreadsheetML type`);
  }
}

const RAW_REFERENCE = /^\$?([A-Za-z]+)\$?([0-9]+)$/;

/** The 0-based column of a `c@r`, failing loudly — out-of-grid addresses
 * name themselves, anything unreadable is malformed. */
function columnOfReference(reference: string, sheetName: string): number {
  const ref = parseA1(reference);
  if (ref) {
    return ref.col;
  }
  if (RAW_REFERENCE.test(reference)) {
    throw outOfGridError(sheetName, reference);
  }
  throw parseError(`"${sheetName}"`, `"${reference}" is not a cell reference`);
}

// --- verbatim attributes ----------------------------------------------------------

function isTagSpace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

/** The attributes of a start tag that `modelled` does not name, verbatim
 * (`name="value"` exactly as the source quotes and escapes the value), in
 * source order, joined by single spaces; undefined when there are none. Only
 * the tag's own text is read — `start`/`end` are the scanner's offsets of a
 * well-formed tag, so every value is quoted. `previous`, when it spells the
 * same text, is returned instead: consecutive rows share one string. */
function unmodelledAttributes(
  xml: string,
  event: { attrs: readonly XmlScanAttribute[]; start: number; end: number },
  modelled: ReadonlySet<string>,
  previous?: string,
): string | undefined {
  if (event.attrs.every((attr) => modelled.has(attr.name))) {
    return undefined;
  }
  let text = "";
  let i = event.start + 1;
  const end = event.end;
  while (i < end && !isTagSpace(xml.charCodeAt(i)) && xml.charCodeAt(i) !== 0x2f && xml.charCodeAt(i) !== 0x3e) {
    i += 1; // the element name
  }
  for (;;) {
    while (i < end && isTagSpace(xml.charCodeAt(i))) {
      i += 1;
    }
    const nameStart = i;
    while (i < end && xml.charCodeAt(i) !== 0x3d /* = */ && !isTagSpace(xml.charCodeAt(i)) && xml.charCodeAt(i) !== 0x2f && xml.charCodeAt(i) !== 0x3e) {
      i += 1;
    }
    if (i === nameStart) {
      break; // `/>` or `>`
    }
    const name = xml.slice(nameStart, i);
    while (i < end && xml.charCodeAt(i) !== 0x22 && xml.charCodeAt(i) !== 0x27) {
      i += 1; // whitespace and `=`
    }
    const valueEnd = i < end ? xml.indexOf(xml[i]!, i + 1) : -1;
    if (valueEnd < 0 || valueEnd >= end) {
      break; // not a quoted value — the scanner never reports such a tag
    }
    if (!modelled.has(name)) {
      text += `${text === "" ? "" : " "}${name}=${xml.slice(i, valueEnd + 1)}`;
    }
    i = valueEnd + 1;
  }
  if (text === "") {
    return undefined;
  }
  return text === previous ? previous : text;
}

/** The `<c>` attributes the model carries; the rest go to `preservedAttrs`
 * (SPEC-122). */
const CELL_MODELLED: ReadonlySet<string> = new Set(["r", "s", "t", "cm"]);
/** The `<row>` attributes the model carries; the rest go to `sourceAttrs`. */
const ROW_MODELLED: ReadonlySet<string> = new Set(["r", "ht", "customHeight", "hidden", "s", "customFormat"]);
/** The `<col>` attributes the model carries (`customWidth` is implied by a
 * width); the rest go to `sourceAttrs`. */
const COL_MODELLED: ReadonlySet<string> = new Set(["min", "max", "width", "customWidth", "hidden", "style"]);

// --- columns ------------------------------------------------------------------------

/** One `<col>` as read, before clipping and the default-width test. */
interface ColumnCandidate {
  min: number; // 0-based inclusive
  max: number;
  width: number | undefined; // characters
  hidden: boolean;
  styleId: number | undefined;
  sourceAttrs: string | undefined;
}

/** The parts of [min, max] no interval of `claimed` (sorted, disjoint)
 * covers. */
function uncovered(claimed: readonly [number, number][], min: number, max: number): [number, number][] {
  const pieces: [number, number][] = [];
  let cursor = min;
  for (const [start, end] of claimed) {
    if (end < cursor) {
      continue;
    }
    if (start > max) {
      break;
    }
    if (start > cursor) {
      pieces.push([cursor, start - 1]);
    }
    cursor = end + 1;
    if (cursor > max) {
      break;
    }
  }
  if (cursor <= max) {
    pieces.push([cursor, max]);
  }
  return pieces;
}

/** Adds [min, max] to `claimed`, keeping it sorted and disjoint. */
function claim(claimed: [number, number][], min: number, max: number): void {
  const merged: [number, number][] = [];
  let start = min;
  let end = max;
  let placed = false;
  for (const interval of claimed) {
    if (interval[1] < start - 1) {
      merged.push(interval);
    } else if (interval[0] > end + 1) {
      if (!placed) {
        merged.push([start, end]);
        placed = true;
      }
      merged.push(interval);
    } else {
      start = Math.min(start, interval[0]);
      end = Math.max(end, interval[1]);
    }
  }
  if (!placed) {
    merged.push([start, end]);
  }
  claimed.splice(0, claimed.length, ...merged);
}

/** The sheet's column spans: each `<col>` clipped to the columns no earlier
 * one claimed, kept only when it differs from the default (a width other
 * than `defaultColWidthPx`, `hidden` or a `styleId`); sorted by `min`. */
function columnSpans(candidates: readonly ColumnCandidate[], defaultColWidthPx: number, mdw: number): ColumnSpan[] {
  const spans: ColumnSpan[] = [];
  const claimed: [number, number][] = [];
  let sorted = true;
  for (const candidate of candidates) {
    const last = claimed.at(-1);
    let pieces: [number, number][];
    if (last === undefined || candidate.min > last[1]) {
      // The usual case — Excel writes cols ascending and disjoint.
      pieces = [[candidate.min, candidate.max]];
      if (last !== undefined && candidate.min === last[1] + 1) {
        last[1] = candidate.max;
      } else {
        claimed.push([candidate.min, candidate.max]);
      }
    } else {
      pieces = uncovered(claimed, candidate.min, candidate.max);
      claim(claimed, candidate.min, candidate.max);
      sorted = false;
    }
    const widthPx = candidate.width === undefined ? undefined : xlsxWidthToPx(candidate.width, mdw);
    if (!(widthPx !== undefined && widthPx !== defaultColWidthPx) && !candidate.hidden && candidate.styleId === undefined) {
      continue;
    }
    for (const [min, max] of pieces) {
      const span: ColumnSpan = { min, max };
      if (widthPx !== undefined) {
        span.widthPx = widthPx;
      }
      if (candidate.hidden) {
        span.hidden = true;
      }
      if (candidate.styleId !== undefined) {
        span.styleId = candidate.styleId;
      }
      if (candidate.width !== undefined) {
        span.sourceWidthChars = candidate.width;
      }
      if (candidate.sourceAttrs !== undefined) {
        span.sourceAttrs = candidate.sourceAttrs;
      }
      spans.push(span);
    }
  }
  return sorted ? spans : spans.toSorted((a, b) => a.min - b.min);
}

// --- the view -----------------------------------------------------------------------

/** One `<selection>` of the first sheetView, as written. */
interface RawSelection {
  pane: string;
  activeCell: string | undefined;
  activeCellId: string | undefined;
  sqref: string | undefined;
}

function cellRefOf(text: string | undefined): CellRef | undefined {
  const ref = text === undefined ? null : parseA1(text.trim());
  return ref ? { row: ref.row, col: ref.col } : undefined;
}

function contains(range: GridRange, ref: CellRef): boolean {
  return ref.row >= range.top && ref.row <= range.bottom && ref.col >= range.left && ref.col <= range.right;
}

function singleCell(ref: CellRef): SheetSelection {
  return { active: ref, ranges: [{ top: ref.row, left: ref.col, bottom: ref.row, right: ref.col }], activeRangeIndex: 0 };
}

/** A `<selection>` → SheetSelection: `activeCell`, `sqref` (or the active
 * cell alone), `activeCellId` → `activeRangeIndex`, normalised so the active
 * cell lies inside its range — another range containing it, else the active
 * cell alone. */
function selectionFrom(raw: RawSelection): SheetSelection {
  const ranges = raw.sqref === undefined ? [] : parseSqref(raw.sqref);
  const requested = parseIndex(raw.activeCellId) ?? 0;
  // No activeCell (the schema has no default): the requested range's corner.
  const fallback = ranges[Math.min(requested, ranges.length - 1)];
  const active = cellRefOf(raw.activeCell) ?? (fallback ? { row: fallback.top, col: fallback.left } : { row: 0, col: 0 });
  if (ranges.length === 0) {
    return singleCell(active);
  }
  if (requested < ranges.length && contains(ranges[requested]!, active)) {
    return { active, ranges, activeRangeIndex: requested };
  }
  const containing = ranges.findIndex((range) => contains(range, active));
  return containing >= 0 ? { active, ranges, activeRangeIndex: containing } : singleCell(active);
}

/** The first `<sheetView>`'s settings, as written. */
interface ViewState {
  showGridLines?: false;
  showZeros?: false;
  showFormulas?: true;
  zoomPercent?: number;
  topLeftCell?: CellRef;
  pane?: { state: string | undefined; xSplit: number | undefined; ySplit: number | undefined; topLeftCell: CellRef | undefined; activePane: string };
  selections: RawSelection[];
}

const ZOOM_MIN = 10;
const ZOOM_MAX = 400;

function readSheetView(attrs: readonly XmlScanAttribute[], view: ViewState): void {
  const gridLines = scanAttr(attrs, "showGridLines");
  if (gridLines === "0" || gridLines === "false") {
    view.showGridLines = false;
  }
  const zeros = scanAttr(attrs, "showZeros");
  if (zeros === "0" || zeros === "false") {
    view.showZeros = false;
  }
  if (isOn(scanAttr(attrs, "showFormulas"))) {
    view.showFormulas = true;
  }
  // A zero (or negative) scale is not a zoom a producer meant — it reads as
  // the default; anything else is clamped into Excel's 10–400%.
  const zoom = parseNumber(scanAttr(attrs, "zoomScale"));
  if (zoom !== undefined && zoom > 0) {
    const percent = Math.round(Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom)));
    if (percent !== 100) {
      view.zoomPercent = percent;
    }
  }
  const topLeft = cellRefOf(scanAttr(attrs, "topLeftCell"));
  if (topLeft) {
    view.topLeftCell = topLeft;
  }
}

/** The frozen split counts of a pane: whole rows/columns, clamped to the
 * grid. */
function splitCount(value: number | undefined, limit: number): number {
  return value === undefined ? 0 : Math.min(limit, Math.max(0, Math.trunc(value)));
}

/** Resolves the shared-formula followers met before their master, at the
 * end of the sheet: each gets its shifted formula; one whose `si` no master
 * defined keeps its value and no formula — and is dropped when that leaves
 * it empty (as is its row, then). */
function resolveFollowers(
  rows: Row[],
  waiting: ReadonlyArray<{ si: string; row: Row; cell: Cell }>,
  masters: ReadonlyMap<string, SharedMaster>,
): void {
  const emptied = new Set<Row>();
  for (const { si, row, cell } of waiting) {
    const master = masters.get(si);
    if (master) {
      Object.assign(cell, followerFormula(master, row.index, cell.col));
    } else if (cell.value === null && cell.styleId === undefined && cell.cm === undefined && cell.preservedAttrs === undefined && cell.preservedXml === undefined) {
      row.cells.splice(row.cells.indexOf(cell), 1);
      emptied.add(row);
    }
  }
  for (const row of emptied) {
    if (row.cells.length === 0 && row.heightPt === undefined && !row.hidden && row.styleId === undefined) {
      rows.splice(rows.indexOf(row), 1);
    }
  }
}

// --- the worksheet --------------------------------------------------------------------

/** What readWorksheet extracts from a worksheet part: the sheet's rows and
 * SPEC-117 layout and view fields (optional ones only when the part has
 * them). */
export type WorksheetContent = Pick<Sheet, "rows" | "columns" | "defaultColWidthPx" | "defaultRowHeightPt" | "merges"> &
  Partial<Pick<Sheet, "frozen" | "showGridLines" | "showZeros" | "zoomPercent" | "tabColor" | "selection" | "topLeftCell" | "showFormulas" | "protection">>;

export interface WorksheetReadOptions {
  sharedStrings: readonly SharedString[];
  dateSystem: "1900" | "1904";
  /** The workbook-wide cell ceiling (parsed.ts). */
  budget: CellBudget;
  /** The workbook default font's maximum digit width (geometry.ts). */
  mdw: number;
  /** The workbook's `<indexedColors>` override palette. */
  indexedColors?: readonly string[];
  /** The workbook's implied-format styles (SPEC-118), for `t="d"` values
   * with a General style. */
  implied?: ImpliedFormatStyles;
}

/** The root's children whose content this reader reads. */
type Section = "sheetPr" | "sheetViews" | "cols" | "sheetData" | "mergeCells";

const SECTIONS: ReadonlySet<string> = new Set<Section>(["sheetPr", "sheetViews", "cols", "sheetData", "mergeCells"]);

/** Reads one worksheet part (see the module header) in a single scanner
 * pass. Returns null when the cell ceiling stopped the scan. */
export function readWorksheet(xml: string, sheetName: string, options: WorksheetReadOptions): WorksheetContent | null {
  const { sharedStrings, dateSystem, budget, mdw, indexedColors, implied } = options;
  const rows: Row[] = [];
  /** Open elements enclosing the next event (the root is level 1). */
  let depth = 0;
  let section: Section | null = null;

  // sheetData
  let lastRow = -1;
  let row: Row | null = null;
  let lastCol = -1;
  let cellsOrdered = true;
  let rowsOrdered = true;
  let pending: PendingCell | null = null;
  let tooLarge = false;
  /** The last row's sourceAttrs, shared by the rows that spell the same. */
  let lastRowAttrs: string | undefined;
  /** SPEC-119: shared-formula masters by `si`, and the followers met before
   * theirs (resolved when the sheet ends). */
  const sharedMasters = new Map<string, SharedMaster>();
  const waitingFollowers: Array<{ si: string; row: Row; cell: Cell }> = [];

  // layout and view
  let defaultRowHeightPt = DEFAULT_ROW_HEIGHT_PT;
  let defaultColWidth: number | undefined;
  let baseColWidth = 8;
  const columnCandidates: ColumnCandidate[] = [];
  const mergeCandidates: GridRange[] = [];
  let tabColor: ModelColor | undefined;
  let protectedSheet = false;
  let sheetViews = 0;
  let inFirstView = false;
  const view: ViewState = { selections: [] };

  /** The style a column's `<col>` gives it (the first `<col>` claiming it). */
  const columnStyleAt = (col: number): number | undefined =>
    columnCandidates.find((candidate) => candidate.min <= col && col <= candidate.max)?.styleId;

  /** What a `<c>`'s `<f>` gives the stored cell (see the module header);
   * "wait" for a follower whose master is still to come. */
  const formulaOf = (cell: PendingCell, rowIndex: number): Partial<Cell> | "wait" | null => {
    const f = cell.f!;
    const type = scanAttr(f.attrs, "t");
    const host = { row: rowIndex, col: cell.col };
    if (type === "dataTable") {
      return { dataTableXml: xml.slice(f.start, f.end) };
    }
    if (type === "shared") {
      const si = scanAttr(f.attrs, "si");
      if (f.text.length > 0) {
        const read = readXlsxFormula(f.text, host);
        if (si !== undefined) {
          sharedMasters.set(si, { wire: f.text, read, row: rowIndex, col: cell.col });
        }
        return read;
      }
      const master = si === undefined ? undefined : sharedMasters.get(si);
      if (master) {
        return followerFormula(master, rowIndex, cell.col);
      }
      return si === undefined ? null : "wait";
    }
    if (f.text.length === 0) {
      return null;
    }
    const read: Partial<Cell> = readXlsxFormula(f.text, host);
    if (type === "array") {
      read.arrayRange = parseRange(scanAttr(f.attrs, "ref")?.trim() ?? "") ?? { top: rowIndex, left: cell.col, bottom: rowIndex, right: cell.col };
      if (cell.cm !== undefined) {
        read.dynamicArray = true;
      }
    }
    return read;
  };

  const finishCell = (cell: PendingCell): "stop" | undefined => {
    const current = row!;
    const { value, runs } = cellValue(cell, sharedStrings, dateSystem, () => `"${sheetName}"!${addressText(current.index, cell.col)}`);
    const formula = cell.f === null ? null : formulaOf(cell, current.index);
    let styleId = cell.styleId;
    if (cell.type === "d" && typeof value === "number" && implied) {
      // A date its producer left General shows as a date (SPEC-118).
      const base = styleId ?? current.styleId ?? columnStyleAt(cell.col) ?? 0;
      const shown = implied.styleFor(base, impliedDateCode(value));
      if (shown !== base) {
        styleId = shown;
      }
    }
    const preservedXml = cell.preservedXml === "" ? undefined : cell.preservedXml;
    if (value === null && styleId === undefined && formula === null && cell.cm === undefined && cell.preservedAttrs === undefined && preservedXml === undefined) {
      return undefined;
    }
    if (!budget.take(1)) {
      tooLarge = true;
      return "stop";
    }
    // One literal per shape — the common cases never grow an object.
    let stored: Cell;
    if (runs) {
      stored = styleId === undefined ? { col: cell.col, value, runs } : { col: cell.col, value, runs, styleId };
    } else {
      stored = styleId === undefined ? { col: cell.col, value } : { col: cell.col, value, styleId };
    }
    if (formula === "wait") {
      waitingFollowers.push({ si: scanAttr(cell.f!.attrs, "si")!, row: current, cell: stored });
    } else if (formula !== null) {
      Object.assign(stored, formula);
    }
    if (cell.cm !== undefined) {
      stored.cm = cell.cm;
    }
    if (cell.type === "s" && value !== null) {
      stored.sstIndex = Number(cell.v);
    }
    if (cell.preservedAttrs !== undefined) {
      stored.preservedAttrs = cell.preservedAttrs;
    }
    if (preservedXml !== undefined) {
      stored.preservedXml = preservedXml;
    }
    const cells = current.cells;
    if (cells.length > 0 && cells[cells.length - 1]!.col >= cell.col) {
      cellsOrdered = false;
    }
    cells.push(stored);
    return undefined;
  };

  const finishRow = (): void => {
    const current = row!;
    if (current.cells.length > 0 || current.heightPt !== undefined || current.hidden || current.styleId !== undefined) {
      if (!cellsOrdered) {
        normaliseCells(current.cells);
      }
      if (rows.length > 0 && rows[rows.length - 1]!.index >= current.index) {
        rowsOrdered = false;
      }
      rows.push(current);
    }
    lastRow = current.index;
    row = null;
  };

  const openRow = (event: Extract<XmlScanEvent, { type: "open" }>): void => {
    const attrs = event.attrs;
    const r = scanAttr(attrs, "r");
    const number = r === undefined ? lastRow + 2 : Number(r);
    if (!Number.isInteger(number) || number < 1) {
      throw parseError(`"${sheetName}"`, `"${r}" is not a row number`);
    }
    if (number > MAX_ROWS) {
      throw outOfGridError(sheetName, `row ${number}`);
    }
    const opened: Row = { index: number - 1, cells: [] };
    const height = parseNumber(scanAttr(attrs, "ht"));
    if (height !== undefined && height >= 0) {
      opened.heightPt = height;
    }
    if (isOn(scanAttr(attrs, "customHeight"))) {
      opened.customHeight = true;
    }
    if (isOn(scanAttr(attrs, "hidden"))) {
      opened.hidden = true;
    }
    if (isOn(scanAttr(attrs, "customFormat"))) {
      const styleId = parseIndex(scanAttr(attrs, "s"));
      if (styleId !== undefined) {
        opened.styleId = styleId;
      }
    }
    const sourceAttrs = unmodelledAttributes(xml, event, ROW_MODELLED, lastRowAttrs);
    if (sourceAttrs !== undefined) {
      opened.sourceAttrs = sourceAttrs;
      lastRowAttrs = sourceAttrs;
    }
    row = opened;
    lastCol = -1;
    cellsOrdered = true;
    if (event.selfClosing) {
      finishRow();
    }
  };

  const openCell = (event: Extract<XmlScanEvent, { type: "open" }>): "stop" | undefined => {
    const r = scanAttr(event.attrs, "r");
    const col = r === undefined ? lastCol + 1 : columnOfReference(r, sheetName);
    if (col >= MAX_COLS) {
      throw outOfGridError(sheetName, addressText(row!.index, col));
    }
    const cell: PendingCell = {
      col,
      type: scanAttr(event.attrs, "t"),
      styleId: parseIndex(scanAttr(event.attrs, "s")),
      v: null,
      inlineString: null,
      depth: 0,
      inV: false,
      cm: parseIndex(scanAttr(event.attrs, "cm")),
      f: null,
      inF: false,
      preservedAttrs: undefined,
      preservedXml: "",
      preservedStart: -1,
    };
    const unmodelled = unmodelledAttributes(xml, event, CELL_MODELLED);
    if (unmodelled !== undefined) {
      cell.preservedAttrs = ` ${unmodelled}`;
    }
    if (event.selfClosing) {
      lastCol = col;
      return finishCell(cell);
    }
    pending = cell;
    return undefined;
  };

  const openCol = (event: Extract<XmlScanEvent, { type: "open" }>): void => {
    const min = parseIndex(scanAttr(event.attrs, "min"));
    const max = parseIndex(scanAttr(event.attrs, "max"));
    if (min === undefined || max === undefined || min < 1 || max < min || min > MAX_COLS) {
      return;
    }
    const width = parseNumber(scanAttr(event.attrs, "width"));
    columnCandidates.push({
      min: min - 1,
      max: Math.min(max, MAX_COLS) - 1,
      width: width !== undefined && width >= 0 ? width : undefined,
      hidden: isOn(scanAttr(event.attrs, "hidden")),
      styleId: parseIndex(scanAttr(event.attrs, "style")),
      sourceAttrs: unmodelledAttributes(xml, event, COL_MODELLED),
    });
  };

  const openInView = (event: Extract<XmlScanEvent, { type: "open" }>): void => {
    if (event.local === "pane" && view.pane === undefined) {
      view.pane = {
        state: scanAttr(event.attrs, "state"),
        xSplit: parseNumber(scanAttr(event.attrs, "xSplit")),
        ySplit: parseNumber(scanAttr(event.attrs, "ySplit")),
        topLeftCell: cellRefOf(scanAttr(event.attrs, "topLeftCell")),
        activePane: scanAttr(event.attrs, "activePane") ?? "topLeft",
      };
    } else if (event.local === "selection") {
      view.selections.push({
        pane: scanAttr(event.attrs, "pane") ?? "topLeft",
        activeCell: scanAttr(event.attrs, "activeCell"),
        activeCellId: scanAttr(event.attrs, "activeCellId"),
        sqref: scanAttr(event.attrs, "sqref"),
      });
    }
  };

  /** An open tag inside a `<c>`: route it to the pending cell. */
  const openInCell = (cell: PendingCell, event: Extract<XmlScanEvent, { type: "open" }>): void => {
    if (cell.depth === 0 && isSpreadsheetMl(event.ns) && (event.local === "v" || event.local === "is" || event.local === "f")) {
      if (event.local === "v") {
        cell.v = "";
        cell.inV = !event.selfClosing;
      } else if (event.local === "is") {
        cell.inlineString = new StringItemReader(indexedColors);
      } else {
        cell.f = { attrs: event.attrs, text: "", start: event.start, end: event.end };
        cell.inF = !event.selfClosing;
      }
    } else if (cell.depth === 0) {
      // SPEC-122: an unmodelled child (`<extLst>`, …), kept verbatim.
      if (event.selfClosing) {
        cell.preservedXml += xml.slice(event.start, event.end);
      } else {
        cell.preservedStart = event.start;
      }
    } else if (cell.inlineString !== null && cell.depth > 0) {
      cell.inlineString.feed(event);
    }
    if (!event.selfClosing) {
      cell.depth += 1;
    }
  };

  const onOpen = (event: Extract<XmlScanEvent, { type: "open" }>, level: number): "stop" | undefined => {
    if (pending !== null) {
      openInCell(pending, event);
      return undefined;
    }
    if (!isSpreadsheetMl(event.ns)) {
      return undefined;
    }
    if (level === 2) {
      if (event.local === "sheetFormatPr") {
        const height = parseNumber(scanAttr(event.attrs, "defaultRowHeight"));
        if (height !== undefined && height > 0) {
          defaultRowHeightPt = height;
        }
        const width = parseNumber(scanAttr(event.attrs, "defaultColWidth"));
        defaultColWidth = width !== undefined && width >= 0 ? width : undefined;
        const base = parseNumber(scanAttr(event.attrs, "baseColWidth"));
        if (base !== undefined && base >= 0) {
          baseColWidth = base;
        }
      } else if (event.local === "sheetProtection" && isOn(scanAttr(event.attrs, "sheet"))) {
        protectedSheet = true;
      }
      section = !event.selfClosing && SECTIONS.has(event.local) ? (event.local as Section) : null;
      return undefined;
    }
    switch (section) {
      case "sheetData":
        if (level === 3 && event.local === "row") {
          openRow(event);
        } else if (level === 4 && row !== null && event.local === "c") {
          return openCell(event);
        }
        return undefined;
      case "cols":
        if (level === 3 && event.local === "col") {
          openCol(event);
        }
        return undefined;
      case "sheetViews":
        if (level === 3 && event.local === "sheetView") {
          sheetViews += 1;
          if (sheetViews === 1) {
            readSheetView(event.attrs, view);
            inFirstView = !event.selfClosing;
          }
        } else if (level === 4 && inFirstView) {
          openInView(event);
        }
        return undefined;
      case "sheetPr":
        if (level === 3 && event.local === "tabColor") {
          tabColor = xlsxColorFrom(scannedAttributes(event.attrs), indexedColors);
        }
        return undefined;
      case "mergeCells":
        if (level === 3 && event.local === "mergeCell") {
          const range = parseRange(scanAttr(event.attrs, "ref")?.trim() ?? "");
          if (range) {
            mergeCandidates.push(range);
          }
        }
        return undefined;
      default:
        return undefined;
    }
  };

  const onClose = (event: Extract<XmlScanEvent, { type: "close" }>, level: number): "stop" | undefined => {
    if (pending !== null) {
      if (pending.depth === 0) {
        // The </c>.
        const cell = pending;
        pending = null;
        lastCol = cell.col;
        return finishCell(cell);
      }
      pending.depth -= 1;
      if (pending.depth === 0) {
        pending.inV = false;
        if (pending.inF) {
          pending.f!.end = event.end;
          pending.inF = false;
        }
        if (pending.preservedStart >= 0) {
          pending.preservedXml += xml.slice(pending.preservedStart, event.end);
          pending.preservedStart = -1;
        }
      } else if (pending.inlineString !== null) {
        pending.inlineString.feed(event);
      }
      return undefined;
    }
    if (level === 2) {
      section = null;
    } else if (section === "sheetData" && level === 3 && row !== null) {
      finishRow();
    } else if (section === "sheetViews" && level === 3) {
      inFirstView = false;
    }
    return undefined;
  };

  scanXml(xml, (event) => {
    if (event.type === "open") {
      const level = depth + 1;
      if (!event.selfClosing) {
        depth = level;
      }
      return onOpen(event, level);
    }
    if (event.type === "close") {
      const level = depth;
      depth -= 1;
      return onClose(event, level);
    }
    if (pending !== null) {
      if (pending.inV) {
        pending.v += event.text;
      } else if (pending.inF) {
        pending.f!.text += event.text;
      } else if (pending.inlineString !== null && pending.depth > 1) {
        pending.inlineString.feed(event);
      }
    }
    return undefined;
  });

  if (tooLarge) {
    return null;
  }
  resolveFollowers(rows, waitingFollowers, sharedMasters);
  if (!rowsOrdered) {
    normaliseRows(rows);
  }
  const defaultColWidthPx = defaultColWidth !== undefined ? xlsxWidthToPx(defaultColWidth, mdw) : baseColWidthToPx(baseColWidth, mdw);
  const content: WorksheetContent = {
    rows,
    columns: columnSpans(columnCandidates, defaultColWidthPx, mdw),
    defaultColWidthPx,
    defaultRowHeightPt,
    merges: acceptMerges(mergeCandidates),
  };
  const pane = view.pane;
  if (pane && (pane.state === "frozen" || pane.state === "frozenSplit")) {
    const frozen = { rows: splitCount(pane.ySplit, MAX_ROWS - 1), cols: splitCount(pane.xSplit, MAX_COLS - 1) };
    if (frozen.rows > 0 || frozen.cols > 0) {
      content.frozen = frozen;
    }
  }
  if (view.showGridLines === false) {
    content.showGridLines = false;
  }
  if (view.showZeros === false) {
    content.showZeros = false;
  }
  if (view.showFormulas) {
    content.showFormulas = true;
  }
  if (view.zoomPercent !== undefined) {
    content.zoomPercent = view.zoomPercent;
  }
  if (tabColor !== undefined) {
    content.tabColor = tabColor;
  }
  if (protectedSheet) {
    content.protection = { sheet: true };
  }
  const activePane = pane?.activePane ?? "topLeft";
  const selection = view.selections.find((candidate) => candidate.pane === activePane);
  if (selection) {
    content.selection = selectionFrom(selection);
  }
  const frozenTopLeft = pane && (pane.state === "frozen" || pane.state === "frozenSplit") ? pane.topLeftCell : undefined;
  const topLeftCell = frozenTopLeft ?? view.topLeftCell;
  if (topLeftCell) {
    content.topLeftCell = topLeftCell;
  }
  return content;
}

// --- the worksheet writer (SPEC-122) ------------------------------------------------

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** An empty worksheet part (SPEC-121 "From scratch"): the tab colour, the
 * selected flag on the active sheet, the default sizes. serializeWorksheet
 * starts a new sheet from it. */
export function emptyWorksheetXml(sheet: Sheet, active: boolean, model: Pick<SpreadsheetModel, "theme">): string {
  const tabColor = sheet.tabColor === undefined ? "" : `<sheetPr><tabColor${modelColorToXlsxAttrs(sheet.tabColor, model.theme)}/></sheetPr>`;
  const colWidth =
    sheet.defaultColWidthPx === DEFAULT_COL_WIDTH_PX ? "" : ` defaultColWidth="${pxToXlsxWidth(sheet.defaultColWidthPx, 7)}"`;
  const rowHeight = Number.isFinite(sheet.defaultRowHeightPt) ? sheet.defaultRowHeightPt : DEFAULT_ROW_HEIGHT_PT;
  return (
    `${XML_DECLARATION}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${tabColor}<dimension ref="A1"/>` +
    `<sheetViews><sheetView workbookViewId="0"${active ? ` tabSelected="1"` : ""}/></sheetViews>` +
    `<sheetFormatPr${colWidth} defaultRowHeight="${rowHeight}"/><sheetData/>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  );
}

/** CT_Worksheet's children, in schema order. */
export const WORKSHEET_ORDER: readonly string[] = [
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

/** What the worksheet writer needs from the save around it. */
export interface XlsxSheetContext {
  /** The sheet is the workbook's active one: its first sheetView is `tabSelected`. */
  active: boolean;
  /** The workbook theme (a new sheet's tab colour). */
  theme: DocumentTheme | undefined;
  /** The workbook default font's maximum digit width (geometry.ts), for column widths. */
  mdw: number;
  /** The save's shared-string table. */
  strings: Pick<SharedStringsWriter, "indexFor">;
  /** The save's style table: a model style id → its cellXfs index. */
  styles: Pick<XlsxStyleWriter, "xfFor">;
}

/** A child of the worksheet root: its local name ("" outside SpreadsheetML),
 * its CT_Worksheet rank (-1 unknown) and its source range. */
interface RootChild {
  local: string;
  order: number;
  mc: boolean;
  start: number;
  end: number;
}

/** A child of the first sheetView. */
interface ViewChild {
  local: string;
  attrs: XmlScanAttribute[];
  start: number;
  end: number;
}

interface WorksheetLayout {
  prefix: string;
  children: RootChild[];
  rootClose: number;
  /** The first sheetView of the first sheetViews. */
  view: { start: number; tagEnd: number; end: number; selfClosing: boolean; name: string; children: ViewChild[] } | null;
}

/** The text with `<sheetData>…</sheetData>` collapsed to `<sheetData/>`:
 * the writer replaces it whole, so its cells — nearly all of a large part —
 * are never scanned. */
function collapseSheetData(xml: string): string {
  let found = null as { start: number; end: number; selfClosing: boolean; name: string } | null;
  let depth = 0;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      if (depth === 1 && isSpreadsheetMl(event.ns) && event.local === "sheetData") {
        found = { start: event.start, end: event.end, selfClosing: event.selfClosing, name: event.name };
        return "stop";
      }
      if (!event.selfClosing) {
        depth += 1;
      }
    } else if (event.type === "close") {
      depth -= 1;
    }
    return undefined;
  });
  if (found === null || found.selfClosing) {
    return xml;
  }
  const close = xml.indexOf(`</${found.name}`, found.end);
  if (close < 0) {
    return xml;
  }
  return `${xml.slice(0, found.start)}<${found.name}/>${xml.slice(xml.indexOf(">", close) + 1)}`;
}

/** The root's children and the first sheetView, located with the scanner. */
function scanLayout(xml: string): WorksheetLayout {
  const layout: WorksheetLayout = { prefix: "", children: [], rootClose: xml.length, view: null };
  let depth = 0;
  let child: RootChild | null = null;
  let viewOpen = false;
  let viewChild: ViewChild | null = null;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      const level = depth + 1;
      if (level === 1) {
        layout.prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
      } else if (level === 2) {
        const main = isSpreadsheetMl(event.ns);
        const entry: RootChild = {
          local: main ? event.local : "",
          order: main ? WORKSHEET_ORDER.indexOf(event.local) : -1,
          mc: event.ns === MC_NS && event.local === "AlternateContent",
          start: event.start,
          end: event.end,
        };
        layout.children.push(entry);
        child = event.selfClosing ? null : entry;
      } else if (level === 3 && child?.local === "sheetViews" && layout.view === null && isSpreadsheetMl(event.ns) && event.local === "sheetView") {
        layout.view = { start: event.start, tagEnd: event.end, end: event.end, selfClosing: event.selfClosing, name: event.name, children: [] };
        viewOpen = !event.selfClosing;
      } else if (level === 4 && viewOpen && layout.view !== null) {
        const entry: ViewChild = { local: isSpreadsheetMl(event.ns) ? event.local : "", attrs: event.attrs, start: event.start, end: event.end };
        layout.view.children.push(entry);
        viewChild = event.selfClosing ? null : entry;
      } else if (level === 4 && child !== null && child.mc && child.order < 0 && isSpreadsheetMl(event.ns)) {
        child.order = WORKSHEET_ORDER.indexOf(event.local); // the element an mc:Choice wraps
      }
      if (!event.selfClosing) {
        depth = level;
      }
      return;
    }
    if (event.type === "close") {
      if (depth === 1) {
        layout.rootClose = event.start;
      } else if (depth === 2 && child !== null) {
        child.end = event.end;
        child = null;
      } else if (depth === 3 && viewOpen && layout.view !== null) {
        layout.view.end = event.end;
        viewOpen = false;
      } else if (depth === 4 && viewChild !== null) {
        viewChild.end = event.end;
        viewChild = null;
      }
      depth -= 1;
    }
  });
  return layout;
}

/** A number as XML text: the shortest round-trip form, -0 as 0. */
function numberText(value: number): string {
  return Object.is(value, -0) ? "0" : String(value);
}

/** `tag` with a boolean attribute following `want` (default `fallback`):
 * left as written when it already says so, else set to `onText`/"0" or
 * removed when `want` is the default. */
function flagAttribute(tag: string, name: string, want: boolean, fallback: boolean): string {
  const raw = attributeValue(tag, name);
  const current = raw === undefined ? fallback : fallback ? !(raw === "0" || raw === "false") : isOn(raw);
  if (current === want) {
    return tag;
  }
  return setAttribute(tag, name, want === fallback ? null : want ? "1" : "0");
}

function isFreeze(state: string | undefined): boolean {
  return state === "frozen" || state === "frozenSplit";
}

/** One `<selection>` for `pane`. */
function selectionXml(selection: SheetSelection, pane: string, p: string): string {
  const sqref = selection.ranges.length > 0 ? selection.ranges.map(formatRange).join(" ") : formatA1(selection.active);
  return (
    `<${p}selection${pane === "topLeft" ? "" : ` pane="${pane}"`} activeCell="${formatA1(selection.active)}"` +
    `${selection.activeRangeIndex > 0 ? ` activeCellId="${selection.activeRangeIndex}"` : ""} sqref="${sqref}"/>`
  );
}

/** The first sheetView's start tag with the model's view attributes. */
function viewTag(tag: string, sheet: Sheet, ctx: XlsxSheetContext): string {
  let next = flagAttribute(tag, "tabSelected", ctx.active, false);
  next = flagAttribute(next, "showGridLines", sheet.showGridLines !== false, true);
  next = flagAttribute(next, "showFormulas", sheet.showFormulas === true, false);
  const zoom = sheet.zoomPercent ?? 100;
  const raw = parseNumber(attributeValue(next, "zoomScale"));
  const current = raw !== undefined && raw > 0 ? Math.round(Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, raw))) : 100;
  if (current !== zoom) {
    next = setAttribute(next, "zoomScale", zoom === 100 ? null : String(zoom));
  }
  if (!isFrozen(sheet)) {
    const want = sheet.topLeftCell;
    const have = cellRefOf(attributeValue(next, "topLeftCell"));
    if (want?.row !== have?.row || want?.col !== have?.col) {
      next = setAttribute(next, "topLeftCell", want === undefined ? null : formatA1(want));
    }
  }
  return next;
}

function isFrozen(sheet: Sheet): boolean {
  return sheet.frozen !== undefined && (sheet.frozen.rows > 0 || sheet.frozen.cols > 0);
}

/** The first sheetView's regenerated pane and selections. */
function paneAndSelections(xml: string, children: readonly ViewChild[], sheet: Sheet, p: string): string {
  let pane = "";
  let active = "topLeft";
  let panes: readonly string[] = ["topLeft"];
  const sourcePane = children.find((child) => child.local === "pane");
  if (isFrozen(sheet)) {
    const { rows, cols } = sheet.frozen!;
    active = rows > 0 && cols > 0 ? "bottomRight" : rows > 0 ? "bottomLeft" : "topRight";
    const topLeft = { row: Math.max(rows, sheet.topLeftCell?.row ?? rows), col: Math.max(cols, sheet.topLeftCell?.col ?? cols) };
    pane =
      `<${p}pane${cols > 0 ? ` xSplit="${cols}"` : ""}${rows > 0 ? ` ySplit="${rows}"` : ""}` +
      ` topLeftCell="${formatA1(topLeft)}" activePane="${active}" state="frozen"/>`;
    panes = rows > 0 && cols > 0 ? ["topRight", "bottomLeft", "bottomRight"] : [active];
  } else if (sourcePane !== undefined && !isFreeze(scanAttr(sourcePane.attrs, "state"))) {
    pane = xml.slice(sourcePane.start, sourcePane.end); // a split: not modelled, kept
    active = scanAttr(sourcePane.attrs, "activePane") ?? "topLeft";
    panes = ["topLeft", "topRight", "bottomLeft", "bottomRight"];
  }
  const kept = children
    .filter((child) => {
      const of = scanAttr(child.attrs, "pane") ?? "topLeft";
      return child.local === "selection" && of !== active && panes.includes(of);
    })
    .map((child) => xml.slice(child.start, child.end));
  return pane + kept.join("") + (sheet.selection ? selectionXml(sheet.selection, active, p) : "");
}

/** The rebuilt `<sheetViews>` (see the module header). A part without one
 * gains a generated element; one whose `<sheetViews>` holds no `<sheetView>`
 * gets the generated view first inside it, the rest kept. */
function sheetViewsXml(xml: string, element: RootChild | undefined, layout: WorksheetLayout, sheet: Sheet, ctx: XlsxSheetContext): string {
  const p = layout.prefix;
  const view = layout.view;
  if (element === undefined || view === null || view.start < element.start || view.end > element.end) {
    const tag = viewTag(`<${p}sheetView workbookViewId="0"/>`, sheet, ctx);
    const block = paneAndSelections(xml, [], sheet, p);
    const generated = block === "" ? tag : `${tag.slice(0, -2)}>${block}</${p}sheetView>`;
    if (element === undefined) {
      return `<${p}sheetViews>${generated}</${p}sheetViews>`;
    }
    const open = startTagAt(xml, element.start);
    if (open.endsWith("/>")) {
      const name = /^<([^\s/>]+)/.exec(open)![1]!;
      return `${open.replace(/\s*\/>$/, ">")}${generated}</${name}>`;
    }
    return open + generated + xml.slice(element.start + open.length, element.end);
  }
  const tag = viewTag(startTagAt(xml, view.start), sheet, ctx);
  const block = paneAndSelections(xml, view.children, sheet, p);
  let viewText: string;
  if (view.selfClosing) {
    viewText = block === "" ? tag : `${tag.replace(/\s*\/>$/, ">")}${block}</${view.name}>`;
  } else {
    const replaced = view.children.filter((child) => child.local === "pane" || child.local === "selection");
    const at = replaced[0]?.start ?? view.tagEnd;
    const splices: Splice[] = [{ start: 0, end: view.tagEnd - view.start, text: tag }];
    splices.push(...replaced.map((child) => ({ start: child.start - view.start, end: child.end - view.start, text: "" })));
    splices.push({ start: at - view.start, end: at - view.start, text: block });
    viewText = applySplices(xml.slice(view.start, view.end), splices);
  }
  return xml.slice(element.start, view.start) + viewText + xml.slice(view.end, element.end);
}

/** The rebuilt `<sheetFormatPr>` (see the module header). */
function sheetFormatPrXml(xml: string, element: RootChild | undefined, sheet: Sheet, ctx: XlsxSheetContext, p: string): string {
  const height = Number.isFinite(sheet.defaultRowHeightPt) && sheet.defaultRowHeightPt > 0 ? sheet.defaultRowHeightPt : DEFAULT_ROW_HEIGHT_PT;
  if (element !== undefined) {
    const text = xml.slice(element.start, element.end);
    const tag = startTagAt(text, 0);
    const raw = parseNumber(attributeValue(tag, "defaultRowHeight"));
    const current = raw !== undefined && raw > 0 ? raw : DEFAULT_ROW_HEIGHT_PT;
    if (current === height) {
      return text;
    }
    return setAttribute(setAttribute(tag, "defaultRowHeight", numberText(height)), "customHeight", "1") + text.slice(tag.length);
  }
  const colWidth =
    sheet.defaultColWidthPx === baseColWidthToPx(8, ctx.mdw) ? "" : ` defaultColWidth="${pxToXlsxWidth(sheet.defaultColWidthPx, ctx.mdw)}"`;
  return `<${p}sheetFormatPr${colWidth} defaultRowHeight="${numberText(height)}"/>`;
}

/** The smallest multiple of 1/256 character the reader's formula
 * (geometry.ts xlsxWidthToPx) maps to `px` (or past it, for a fractional
 * px). */
export function xlsxWidthForPx(px: number, mdw: number): number {
  const offset = Math.trunc(128 / mdw);
  const toPx = (k: number): number => Math.trunc(((k + offset) / 256) * mdw);
  let k = Math.max(0, Math.ceil((px * 256) / mdw) - offset - 2);
  while (toPx(k) < px) {
    k += 1;
  }
  while (k > 0 && toPx(k - 1) >= px) {
    k -= 1;
  }
  return k / 256;
}

function colsXml(sheet: Sheet, ctx: XlsxSheetContext, p: string): string {
  if (sheet.columns.length === 0) {
    return "";
  }
  const cols = sheet.columns.map((span) => {
    let attrs = ` min="${span.min + 1}" max="${span.max + 1}"`;
    if (span.widthPx !== undefined) {
      attrs += ` width="${numberText(span.sourceWidthChars ?? xlsxWidthForPx(span.widthPx, ctx.mdw))}"`;
    }
    if (span.styleId !== undefined) {
      attrs += ` style="${ctx.styles.xfFor(span.styleId)}"`;
    }
    if (span.hidden) {
      attrs += ` hidden="1"`;
    }
    if (span.widthPx !== undefined) {
      attrs += ` customWidth="1"`;
    }
    if (span.sourceAttrs !== undefined) {
      attrs += ` ${span.sourceAttrs}`;
    }
    return `<${p}col${attrs}/>`;
  });
  return `<${p}cols>${cols.join("")}</${p}cols>`;
}

/** Row attributes the writer drops: Excel's layout caches. */
const ROW_CACHE_ATTRIBUTES = /^(?:spans|(?:[\w.-]+:)?dyDescent)$/;

/** A row's sourceAttrs less `spans` and `x14ac:dyDescent`, with a leading
 * space ("" when nothing is left); memoised — rows share the strings. */
function rowSourceAttrs(attrs: string, memo: Map<string, string>): string {
  let out = memo.get(attrs);
  if (out === undefined) {
    out = "";
    for (const match of attrs.matchAll(/([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/g)) {
      if (!ROW_CACHE_ATTRIBUTES.test(match[1]!)) {
        out += ` ${match[1]}=${match[2]}`;
      }
    }
    memo.set(attrs, out);
  }
  return out;
}

function inAnyRange(ranges: readonly GridRange[], row: number, col: number): boolean {
  return ranges.some((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
}

/** One `<c>` (see the module header). */
function cellXml(cell: Cell, row: Row, sheet: Sheet, ctx: XlsxSheetContext, p: string, arrays: readonly GridRange[]): string {
  let attrs = ` r="${colToLetters(cell.col)}${row.index + 1}"`;
  if (cell.styleId !== undefined) {
    const xf = ctx.styles.xfFor(cell.styleId);
    if (xf !== 0 || (row.styleId ?? columnSpanAt(sheet.columns, cell.col)?.styleId ?? 0) !== 0) {
      attrs += ` s="${xf}"`;
    }
  }
  const value = cell.value;
  const computed = cell.formula !== undefined || cell.dataTableXml !== undefined;
  let type = "";
  let v = "";
  if (typeof value === "number") {
    v = numberText(value);
  } else if (typeof value === "string") {
    if (computed || (arrays.length > 0 && inAnyRange(arrays, row.index, cell.col))) {
      type = "str";
      v = encodeXlsxText(value);
    } else {
      type = "s";
      v = String(ctx.strings.indexFor(cell));
    }
  } else if (typeof value === "boolean") {
    type = "b";
    v = value ? "1" : "0";
  } else if (value !== null) {
    type = "e";
    v = escapeElementText(value.error);
  }
  if (type !== "") {
    attrs += ` t="${type}"`;
  }
  if (cell.cm !== undefined) {
    attrs += ` cm="${cell.cm}"`;
  }
  if (cell.preservedAttrs !== undefined) {
    attrs += cell.preservedAttrs;
  }
  let children = "";
  if (cell.dataTableXml !== undefined) {
    children += cell.dataTableXml;
  } else if (cell.formula !== undefined) {
    const text = escapeElementText(cell.formulaUnparsed ? cell.formula : toXlsxFormula(cell.formula));
    children += cell.arrayRange ? `<${p}f t="array" ref="${formatRange(cell.arrayRange)}">${text}</${p}f>` : `<${p}f>${text}</${p}f>`;
  }
  if (value !== null) {
    children += `<${p}v>${v}</${p}v>`;
  }
  if (cell.preservedXml !== undefined) {
    children += cell.preservedXml;
  }
  return children === "" ? `<${p}c${attrs}/>` : `<${p}c${attrs}>${children}</${p}c>`;
}

function sheetDataXml(sheet: Sheet, ctx: XlsxSheetContext, p: string): string {
  if (sheet.rows.length === 0) {
    return `<${p}sheetData/>`;
  }
  const arrays: GridRange[] = [];
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.arrayRange !== undefined) {
        arrays.push(cell.arrayRange);
      }
    }
  }
  const memo = new Map<string, string>();
  const out: string[] = [`<${p}sheetData>`];
  for (const row of sheet.rows) {
    let attrs = ` r="${row.index + 1}"`;
    if (row.styleId !== undefined) {
      attrs += ` s="${ctx.styles.xfFor(row.styleId)}" customFormat="1"`;
    }
    if (row.heightPt !== undefined) {
      attrs += ` ht="${numberText(row.heightPt)}"${row.customHeight ? ` customHeight="1"` : ""}`;
    }
    if (row.hidden) {
      attrs += ` hidden="1"`;
    }
    if (row.sourceAttrs !== undefined) {
      attrs += rowSourceAttrs(row.sourceAttrs, memo);
    }
    if (row.cells.length === 0) {
      out.push(`<${p}row${attrs}/>`);
      continue;
    }
    out.push(`<${p}row${attrs}>`);
    for (const cell of row.cells) {
      out.push(cellXml(cell, row, sheet, ctx, p, arrays));
    }
    out.push(`</${p}row>`);
  }
  out.push(`</${p}sheetData>`);
  return out.join("");
}

function mergeCellsXml(sheet: Sheet, p: string): string {
  if (sheet.merges.length === 0) {
    return "";
  }
  return `<${p}mergeCells count="${sheet.merges.length}">${sheet.merges.map((merge) => `<${p}mergeCell ref="${formatRange(merge)}"/>`).join("")}</${p}mergeCells>`;
}

/** A worksheet part for `sheet` (see the module header): the model's
 * dimension, views, default sizes, columns, cells and merges spliced into
 * `base` — the retained part with the writer's sheet-reference pass and tab
 * colour already applied — or into SPEC-121's empty-sheet template when the
 * sheet has none. */
export function serializeWorksheet(sheet: Sheet, base: string | undefined, ctx: XlsxSheetContext): string {
  const xml = collapseSheetData(base ?? emptyWorksheetXml(sheet, ctx.active, ctx));
  const layout = scanLayout(xml);
  const p = layout.prefix;
  const first = (local: string): RootChild | undefined => layout.children.find((child) => child.local === local);
  const range = usedRange(sheet);
  const rebuilt: [string, string][] = [
    ["dimension", `<${p}dimension ref="${range ? formatRange(range) : "A1"}"/>`],
    ["sheetViews", sheetViewsXml(xml, first("sheetViews"), layout, sheet, ctx)],
    ["sheetFormatPr", sheetFormatPrXml(xml, first("sheetFormatPr"), sheet, ctx, p)],
    ["cols", colsXml(sheet, ctx, p)],
    ["sheetData", sheetDataXml(sheet, ctx, p)],
    ["mergeCells", mergeCellsXml(sheet, p)],
  ];
  const splices: Splice[] = [];
  for (const [local, text] of rebuilt) {
    const found = first(local);
    if (found !== undefined) {
      if (xml.slice(found.start, found.end) !== text) {
        splices.push({ start: found.start, end: found.end, text });
      }
    } else if (text !== "") {
      const rank = WORKSHEET_ORDER.indexOf(local);
      const at = layout.children.find((child) => child.order > rank)?.start ?? layout.rootClose;
      splices.push({ start: at, end: at, text });
    }
  }
  return applySplices(xml, splices);
}
