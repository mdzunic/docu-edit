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
import type { ModelColor } from "../../core/theme";
import { scanAttr, scanXml, type XmlScanAttribute, type XmlScanEvent } from "../../core/xmlScanner";
import { parseA1, parseRange, parseSqref, type CellRef, type GridRange } from "../address";
import { baseColWidthToPx, DEFAULT_ROW_HEIGHT_PT, xlsxWidthToPx } from "../geometry";
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
  type TextRun,
} from "../model";
import { impliedDateCode, type ImpliedFormatStyles } from "../impliedFormats";
import { readXlsxFormula, type FormulaFields } from "../formula/dialect";
import { FormulaSyntaxError, parseFormulaAt } from "../formula/parse";
import { shiftFormula } from "../formula/references";
import { addressText, normaliseCells, normaliseRows, outOfGridError, parseError, type CellBudget } from "../parsed";
import type { SheetSelection } from "../selection";
import { odsDateValueToSerial } from "../serial";
import { isSpreadsheetMl, XlsxFontReader } from "./styles";
import { scannedAttributes, xlsxColorFrom } from "./themeColor";

/** A shared string: its text, plus its runs when it was rich. */
export interface SharedString {
  text: string;
  runs?: TextRun[];
}

/** `_xHHHH_` escapes decoded to their code unit. `_x005F_` is a literal
 * underscore, so `_x005F_x0041_` reads as `_x0041_` (the match after an
 * escape resumes past it). */
export function decodeXmlEscapes(text: string): string {
  if (!text.includes("_x")) {
    return text;
  }
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

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
    } else if (cell.value === null && cell.styleId === undefined && cell.cm === undefined) {
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
  Partial<Pick<Sheet, "frozen" | "showGridLines" | "showZeros" | "zoomPercent" | "tabColor" | "selection" | "topLeftCell" | "showFormulas">>;

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
    if (value === null && styleId === undefined && formula === null && cell.cm === undefined) {
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
    };
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
    if (cell.depth === 0 && isSpreadsheetMl(event.ns)) {
      if (event.local === "v") {
        cell.v = "";
        cell.inV = !event.selfClosing;
      } else if (event.local === "is") {
        cell.inlineString = new StringItemReader(indexedColors);
      } else if (event.local === "f") {
        cell.f = { attrs: event.attrs, text: "", start: event.start, end: event.end };
        cell.inF = !event.selfClosing;
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
