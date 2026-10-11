// The .xlsx half of SPEC-124's retained references (see
// ../retainedReferences.ts, whose header is the mapper contract): the row and
// column numbers inside the reference-bearing xlsx XML this product does not
// model yet, moved through a RetainedReferenceMapper by offset splices
// located with core/xmlScanner.ts — the changed attribute values and text
// nodes, whole elements for removals, every other byte verbatim. A row leaves
// this table in the change that models its element (structure.ts).
//
// | Part | Element / attribute | Kind |
// | sheet | `conditionalFormatting@sqref`; `cfRule/formula` | sqref; formula |
// | sheet | `dataValidation@sqref`; `formula1`, `formula2` | sqref; formula |
// | sheet | `customSheetView`'s `autoFilter@ref`, `sortState@ref`, `sortCondition@ref` | range |
// | sheet | `hyperlink@ref` | range |
// | sheet | `rowBreaks/brk@id` (rows), `colBreaks/brk@id` (columns) | index |
// | sheet | `protectedRange@sqref`, `ignoredError@sqref`, `dataRef@ref`, `inputCells@r`, `cellWatch@r` | sqref |
// | sheet `extLst` | `x14:conditionalFormatting/xm:sqref`, its `xm:f`; `x14:dataValidation/xm:sqref`, `x14:formula1|2/xm:f`; `x14:sparkline/xm:sqref`, `xm:f` (and its group's date-axis `xm:f`) | sqref; formula |
// | workbook | `definedName` text (but `_xlnm._FilterDatabase`'s) | formula |
// | tables | `table@ref`, `autoFilter@ref`, `sortState@ref`, `sortCondition@ref` | range |
// | charts (not style/colors parts) | every `c:f` (any `f` element: `cx:f`, `c15:f` too) | formula |
// | pivot caches | `worksheetSource@ref` on `@sheet` (one naming `name` or `r:id` is left alone) | range |
// | pivot tables | `location@ref` | range |
// | drawings | `xdr:from`/`xdr:to` `xdr:row`, `xdr:col` — a twoCellAnchor's unless `editAs="absolute"`, a oneCellAnchor's `from` | index |
// | comments, threaded comments | `comment@ref`, `threadedComment@ref` | range |
// | VML | `x:ClientData` `x:Row`, `x:Column`, `x:Anchor` numbers 1, 3, 5 and 7 (LeftCol, TopRow, RightCol, BottomRow) | index |
//
// - Retired by SPEC-128 (its `autofilter` adjuster moves the model, the
//   writer serialises it): the worksheet's root-level `autoFilter@ref` and
//   `sortState@ref` with every `sortState`/`sortCondition` inside them, and
//   the `_xlnm._FilterDatabase` definedName text (writer.ts reconciles it
//   with the model) — so nothing is adjusted twice. A table part's
//   autoFilter and sortState (SPEC-133's) and a customSheetView's stay here.
// - Hosts. Every model sheet with a `partName` (grid and foreign) hosts its
//   part and what the part's relationships reach — tables, drawings and the
//   charts they relate, comments, threaded comments, VML, pivot tables; the
//   returned set holds the sheets whose own parts changed. A chart no
//   sheet's drawing relates has host "", as has a definedName without
//   `localSheetId`; one with it is hosted by the model sheet holding the part
//   the SOURCE workbook's `<sheet>` at that index names. Pivot caches and the
//   workbook part belong to no sheet.
// - Sheet names. `range` and `index` get the host's model name — or, for
//   `dataRef@sheet` and `worksheetSource@sheet` (source-time names), the
//   name of the model sheet resolveRetainedSheetName finds (nothing found,
//   nothing moves; a `dataRef` naming no sheet is on its host). `formula`
//   gets canonical text naming sheets as found, and the host id. No sheet
//   name is ever rewritten.
// - Range text is a whitespace-separated A1 list: each entry through
//   `range`, a moved one printed with its own `$` markers and shape (one
//   cell or two), an untouched or unparseable one and every separator kept
//   as bytes. Where an element may go (conditionalFormatting, dataValidation,
//   hyperlink, protectedRange, ignoredError, cellWatch, comment,
//   threadedComment, x14:conditionalFormatting, x14:dataValidation,
//   x14:sparkline) an entry mapped to null is dropped, with the separator
//   before it, and an element left with no entry is removed; anywhere else a
//   null leaves the entry as it was.
// - Removal takes the whole element and what an empty container cannot be
//   left as: `<dataValidations>`, `<hyperlinks>`, `<protectedRanges>`,
//   `<ignoredErrors>` and `<cellWatches>` go with their last item, a
//   `dataValidations@count` (x14's too) is set to the items left, a
//   sparklineGroup goes with its last sparkline, and a legacy comment takes
//   its VML `v:shape` — the one whose `x:Row`/`x:Column` (ObjectType Note)
//   name the comment's ORIGINAL cell in the host's VML parts.
// - Formula text: decoded, fromXlsxFormula, `formula`, toXlsxFormula,
//   escaped — written only when the canonical text changed. Index text: the
//   integer through `index` (a null leaves it).
// - Copy-on-write: a changed part becomes a NEW OpaquePart in a NEW `opaque`
//   array; every other part keeps its identity, and `opaque` itself when
//   nothing changes. A part that is not well-formed XML or not UTF-8 is left
//   as it is. A byte-order mark survives.
// - Cost. A worksheet part is split around its `<sheetData>` without
//   decoding the cells: only base64 near the part's head and tail is
//   decoded, the tail is scanned (behind `<sheetData/>` standing in for the
//   cells, every element this table names follows them) only when it
//   mentions one of them, and a changed tail is re-encoded onto the untouched
//   base64 before it. Nothing here touches `<sheetData>`, `<mergeCells>`,
//   `<cols>`, `<sheetViews>`, `<dimension>` or the root-level `<autoFilter>`
//   and `<sortState>` — the writer rebuilds them.
import { base64ToBytes, bytesToBase64, type OpaquePart } from "../../core/packageParts";
import { scanAttr, scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { colToLetters, parseA1, type GridRange } from "../address";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import type { Sheet, SpreadsheetModel } from "../model";
import type { RetainedReferenceMapper } from "../retainedReferences";
import { resolveRetainedSheetName } from "../sourceNames";
import { applySplices, decodeXmlText, escapeElementText, findAttribute, startTagAt, type Splice } from "../xmlSplice";
import { FILTER_DATABASE_NAME } from "./autoFilter";
import {
  encodeUtf8,
  LazyParts,
  readSourceWorkbook,
  relationshipKind,
  relationshipsOf,
  targetPart,
  type PartLookup,
  type RelEntry,
  type SourceWorkbook,
} from "./package";

const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
const RELATIONSHIP_ID_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

type Axis = "row" | "col";

/** The bytes as UTF-8 text (a byte-order mark kept), or undefined when they
 * are not UTF-8 — such a part is never re-encoded. */
const STRICT_UTF8 = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true });

function decodeStrict(bytes: Uint8Array): string | undefined {
  try {
    return STRICT_UTF8.decode(bytes);
  } catch {
    return undefined;
  }
}

// --- the text kinds ---------------------------------------------------------------------

interface AbsoluteMarks {
  rowAbs: boolean;
  colAbs: boolean;
}

/** One A1 range-list entry as written: its range and the `$` markers of its
 * one or two cells. */
interface RangeEntry {
  range: GridRange;
  first: AbsoluteMarks;
  second?: AbsoluteMarks;
}

function parseEntry(text: string): RangeEntry | null {
  const colon = text.indexOf(":");
  const first = parseA1(colon < 0 ? text : text.slice(0, colon));
  if (!first) {
    return null;
  }
  if (colon < 0) {
    return { range: { top: first.row, left: first.col, bottom: first.row, right: first.col }, first };
  }
  const second = parseA1(text.slice(colon + 1));
  if (!second) {
    return null;
  }
  return {
    range: {
      top: Math.min(first.row, second.row),
      left: Math.min(first.col, second.col),
      bottom: Math.max(first.row, second.row),
      right: Math.max(first.col, second.col),
    },
    first,
    second,
  };
}

function formatCell(row: number, col: number, marks: AbsoluteMarks): string {
  return `${marks.colAbs ? "$" : ""}${colToLetters(col)}${marks.rowAbs ? "$" : ""}${row + 1}`;
}

/** A moved entry, in the shape it was written in. */
function formatEntry(range: GridRange, entry: RangeEntry): string {
  const start = formatCell(range.top, range.left, entry.first);
  if (entry.second === undefined && range.top === range.bottom && range.left === range.right) {
    return start;
  }
  return `${start}:${formatCell(range.bottom, range.right, entry.second ?? entry.first)}`;
}

function sameRange(a: GridRange, b: GridRange): boolean {
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

/** A range list's raw text moved on `sheetName` (see the module header):
 * the new text, null when `removable` and no entry is left, undefined when
 * nothing changes. */
function adjustRangeList(raw: string, sheetName: string, mapper: RetainedReferenceMapper, removable: boolean): string | null | undefined {
  const pattern = /\S+/g;
  let out = "";
  let leading: string | undefined;
  let previousEnd = 0;
  let kept = 0;
  let changed = false;
  for (let match = pattern.exec(raw); match !== null; match = pattern.exec(raw)) {
    const separator = raw.slice(previousEnd, match.index);
    previousEnd = match.index + match[0].length;
    leading ??= separator;
    let text = match[0];
    const entry = parseEntry(text);
    if (entry) {
      const moved = mapper.range(sheetName, entry.range);
      if (moved === null) {
        if (removable) {
          changed = true;
          continue;
        }
      } else if (!sameRange(moved, entry.range)) {
        text = formatEntry(moved, entry);
        changed = true;
      }
    }
    out += kept === 0 ? text : separator + text;
    kept += 1;
  }
  if (!changed) {
    return undefined;
  }
  return kept === 0 ? null : (leading ?? "") + out + raw.slice(previousEnd);
}

/** An integer's raw text moved along `axis`: the index as written and, when
 * it moves, the new text. */
function adjustIndexText(
  raw: string,
  sheetName: string,
  axis: Axis,
  mapper: RetainedReferenceMapper,
): { index: number; text?: string } | undefined {
  const match = /^(\s*)(\d{1,9})(\s*)$/.exec(raw);
  if (!match) {
    return undefined;
  }
  const index = Number(match[2]);
  const moved = mapper.index(sheetName, axis, index);
  return moved === null || moved === index ? { index } : { index, text: `${match[1]}${moved}${match[3]}` };
}

/** The axes of a VML `x:Anchor`'s eight numbers (the offsets never move). */
const ANCHOR_AXES: readonly (Axis | undefined)[] = ["col", undefined, "row", undefined, "col", undefined, "row", undefined];

function adjustAnchorText(raw: string, sheetName: string, mapper: RetainedReferenceMapper): string | undefined {
  let position = 0;
  let changed = false;
  const text = raw.replace(/\d+/g, (digits) => {
    const axis = ANCHOR_AXES[position];
    position += 1;
    if (axis === undefined) {
      return digits;
    }
    const index = Number(digits);
    const moved = mapper.index(sheetName, axis, index);
    if (moved === null || moved === index) {
      return digits;
    }
    changed = true;
    return String(moved);
  });
  return changed ? text : undefined;
}

/** Formula element text moved: the new raw text, undefined when the
 * canonical text does not change. */
function adjustFormulaText(raw: string, hostSheetId: string, mapper: RetainedReferenceMapper): string | undefined {
  if (raw.trim() === "") {
    return undefined;
  }
  const canonical = fromXlsxFormula(decodeXmlText(raw));
  const next = mapper.formula(canonical, hostSheetId);
  return next === canonical ? undefined : escapeElementText(toXlsxFormula(next));
}

// --- scanning ---------------------------------------------------------------------------

/** An open element, with what the passes note on it. */
interface Frame {
  local: string;
  ns: string | undefined;
  attrs: readonly XmlScanAttribute[];
  /** Offset of the start tag's `<`, and just past its `>`. */
  start: number;
  openEnd: number;
  parent: Frame | undefined;
  /** Removed whole when it closes. */
  remove?: true;
  /** An item counted in its parent container's `items`. */
  item?: true;
  items?: number;
  removedItems?: number;
  /** VML: a shape's ClientData ObjectType and note cell as written. */
  objectType?: string;
  noteRow?: number;
  noteCol?: number;
}

/** Scans `xml` keeping the open elements as a chain of frames: `onOpen` as
 * each element opens, `onClose` as it closes (at once for a self-closing
 * one) with the end of its content and of the element. */
function walk(
  xml: string,
  onOpen: (frame: Frame) => void,
  onClose: (frame: Frame, contentEnd: number, end: number) => void,
): void {
  let top: Frame | undefined;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      const frame: Frame = { local: event.local, ns: event.ns, attrs: event.attrs, start: event.start, openEnd: event.end, parent: top };
      onOpen(frame);
      if (event.selfClosing) {
        onClose(frame, event.end, event.end);
      } else {
        top = frame;
      }
    } else if (event.type === "close" && top) {
      const frame = top;
      top = frame.parent;
      onClose(frame, event.start, event.end);
    }
  });
}

const ignore = (): void => {};

function hasAncestor(frame: Frame, test: (ancestor: Frame) => boolean): boolean {
  for (let at = frame.parent; at; at = at.parent) {
    if (test(at)) {
      return true;
    }
  }
  return false;
}

/** Counts `frame` as an item of its parent container. */
function countItem(frame: Frame): void {
  if (frame.parent) {
    frame.item = true;
    frame.parent.items = (frame.parent.items ?? 0) + 1;
  }
}

function hasRelationshipId(attrs: readonly XmlScanAttribute[]): boolean {
  return attrs.some((attr) => attr.local === "id" && attr.ns !== undefined && RELATIONSHIP_ID_NAMESPACES.has(attr.ns));
}

/** One part's edits: value splices, and whole elements removed (a removal
 * swallows every splice inside it, an enclosing removal every removal). */
class PartEdits {
  private readonly xml: string;
  private readonly mapper: RetainedReferenceMapper;
  private readonly values: Splice[] = [];
  private readonly removals: { start: number; end: number }[] = [];

  constructor(xml: string, mapper: RetainedReferenceMapper) {
    this.xml = xml;
    this.mapper = mapper;
  }

  /** A start-tag attribute's raw value and its offsets in the part. */
  private attribute(frame: Frame, qname: string): { raw: string; start: number; end: number } | undefined {
    const span = findAttribute(startTagAt(this.xml, frame.start), qname);
    return span && { raw: span.raw, start: frame.start + span.valueStart, end: frame.start + span.valueEnd };
  }

  /** An element's raw text content; undefined when it holds markup. */
  private content(frame: Frame, contentEnd: number): string | undefined {
    const raw = this.xml.slice(frame.openEnd, contentEnd);
    return raw.includes("<") ? undefined : raw;
  }

  private range(raw: string, start: number, end: number, sheetName: string, removable: boolean): "gone" | undefined {
    const next = adjustRangeList(raw, sheetName, this.mapper, removable);
    if (next === null) {
      return "gone";
    }
    if (next !== undefined) {
      this.values.push({ start, end, text: next });
    }
    return undefined;
  }

  /** A range-list attribute moved; "gone" when `removable` and emptied. */
  rangeAttribute(frame: Frame, qname: string, sheetName: string, removable: boolean): "gone" | undefined {
    const value = this.attribute(frame, qname);
    return value && this.range(value.raw, value.start, value.end, sheetName, removable);
  }

  /** Range-list element text moved; "gone" when `removable` and emptied. */
  rangeText(frame: Frame, contentEnd: number, sheetName: string, removable: boolean): "gone" | undefined {
    const raw = this.content(frame, contentEnd);
    return raw === undefined ? undefined : this.range(raw, frame.openEnd, contentEnd, sheetName, removable);
  }

  indexAttribute(frame: Frame, qname: string, sheetName: string, axis: Axis): void {
    const value = this.attribute(frame, qname);
    const next = value && adjustIndexText(value.raw, sheetName, axis, this.mapper);
    if (value && next?.text !== undefined) {
      this.values.push({ start: value.start, end: value.end, text: next.text });
    }
  }

  /** Integer element text moved; returns the index as written. */
  indexText(frame: Frame, contentEnd: number, sheetName: string, axis: Axis): number | undefined {
    const raw = this.content(frame, contentEnd);
    const next = raw === undefined ? undefined : adjustIndexText(raw, sheetName, axis, this.mapper);
    if (next?.text !== undefined) {
      this.values.push({ start: frame.openEnd, end: contentEnd, text: next.text });
    }
    return next?.index;
  }

  anchorText(frame: Frame, contentEnd: number, sheetName: string): void {
    const raw = this.content(frame, contentEnd);
    const next = raw === undefined ? undefined : adjustAnchorText(raw, sheetName, this.mapper);
    if (next !== undefined) {
      this.values.push({ start: frame.openEnd, end: contentEnd, text: next });
    }
  }

  formulaText(frame: Frame, contentEnd: number, hostSheetId: string): void {
    const raw = this.content(frame, contentEnd);
    const next = raw === undefined ? undefined : adjustFormulaText(raw, hostSheetId, this.mapper);
    if (next !== undefined) {
      this.values.push({ start: frame.openEnd, end: contentEnd, text: next });
    }
  }

  /** A container's `count` set to the items it keeps. */
  setCount(frame: Frame, qname: string, count: number): void {
    const value = this.attribute(frame, qname);
    if (value && value.raw !== String(count)) {
      this.values.push({ start: value.start, end: value.end, text: String(count) });
    }
  }

  /** Every element's close: one marked `remove` goes, counted out of its
   * container. */
  close(frame: Frame, end: number): void {
    if (!frame.remove) {
      return;
    }
    this.removals.push({ start: frame.start, end });
    if (frame.item && frame.parent) {
      frame.parent.removedItems = (frame.parent.removedItems ?? 0) + 1;
    }
  }

  splices(): Splice[] {
    const outer: { start: number; end: number }[] = [];
    for (const removal of this.removals.toSorted((a, b) => a.start - b.start || b.end - a.end)) {
      const last = outer.at(-1);
      if (!last || removal.start >= last.end) {
        outer.push(removal); // elements nest or are disjoint
      }
    }
    const out: Splice[] = outer.map(({ start, end }) => ({ start, end, text: "" }));
    let r = 0;
    for (const value of this.values.toSorted((a, b) => a.start - b.start)) {
      while (r < outer.length && outer[r]!.end <= value.start) {
        r += 1;
      }
      const removal = outer[r];
      if (!removal || value.start < removal.start) {
        out.push(value);
      }
    }
    return out;
  }
}

// --- the passes ---------------------------------------------------------------------------

/** What the passes read with. */
interface Pass {
  readonly model: SpreadsheetModel;
  readonly parts: PartLookup;
  readonly mapper: RetainedReferenceMapper;
}

/** Sheet elements an edit can empty, with their range attribute and the
 * container they are items of. */
const REMOVABLE_SHEET_ELEMENTS: ReadonlyMap<string, { attribute: string; container?: string }> = new Map([
  ["conditionalFormatting", { attribute: "sqref" }],
  ["dataValidation", { attribute: "sqref", container: "dataValidations" }],
  ["hyperlink", { attribute: "ref", container: "hyperlinks" }],
  ["protectedRange", { attribute: "sqref", container: "protectedRanges" }],
  ["ignoredError", { attribute: "sqref", container: "ignoredErrors" }],
  ["cellWatch", { attribute: "r", container: "cellWatches" }],
]);
/** Containers that may not be left empty. */
const EMPTY_CONTAINERS_REMOVED: ReadonlySet<string> = new Set(["dataValidations", "hyperlinks", "protectedRanges", "ignoredErrors", "cellWatches"]);
/** Sheet elements whose range attribute moves but never empties them. */
const MOVED_SHEET_ATTRIBUTES: ReadonlyMap<string, string> = new Map([
  ["autoFilter", "ref"],
  ["sortState", "ref"],
  ["sortCondition", "ref"],
  ["inputCells", "r"],
]);

/** SPEC-128: the worksheet's root-level autoFilter and sortState (and
 * everything inside them) are modelled — the autofilter adjuster moves
 * them; a customSheetView's are still this table's. */
function insideModelledFilter(frame: Frame): boolean {
  for (let at: Frame | undefined = frame; at; at = at.parent) {
    if ((at.local === "autoFilter" || at.local === "sortState") && at.parent !== undefined && at.parent.parent === undefined) {
      return true;
    }
  }
  return false;
}
/** x14 elements whose `xm:sqref` child is their range, and the x14 elements
 * whose `xm:f` descendants are formulas. */
const X14_SQREF_OWNERS: ReadonlySet<string> = new Set(["conditionalFormatting", "dataValidation", "sparkline"]);
const X14_FORMULA_SCOPES: ReadonlySet<string> = new Set(["conditionalFormatting", "dataValidation", "sparklineGroup"]);

/** Whether a worksheet's text may hold an element of the table. */
const WORKSHEET_ELEMENTS =
  /conditionalFormatting|dataValidation|autoFilter|sortState|hyperlink|rowBreaks|colBreaks|protectedRange|ignoredError|dataRef|inputCells|cellWatch|sparkline/;

function worksheetEdits(xml: string, host: Sheet, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      const parent = frame.parent;
      if (frame.ns === X14_NS) {
        if ((frame.local === "dataValidation" && parent?.local === "dataValidations") || (frame.local === "sparkline" && parent?.local === "sparklines")) {
          countItem(frame);
        }
        return;
      }
      if (frame.ns === XM_NS) {
        return;
      }
      const removable = REMOVABLE_SHEET_ELEMENTS.get(frame.local);
      if (removable) {
        if (removable.container !== undefined && parent?.local === removable.container) {
          countItem(frame);
        }
        if (edits.rangeAttribute(frame, removable.attribute, host.name, true) === "gone") {
          frame.remove = true;
        }
        return;
      }
      const moved = MOVED_SHEET_ATTRIBUTES.get(frame.local);
      if (moved !== undefined) {
        if (frame.local === "inputCells" || !insideModelledFilter(frame)) {
          edits.rangeAttribute(frame, moved, host.name, false);
        }
      } else if (frame.local === "brk" && (parent?.local === "rowBreaks" || parent?.local === "colBreaks")) {
        edits.indexAttribute(frame, "id", host.name, parent.local === "rowBreaks" ? "row" : "col");
      } else if (frame.local === "dataRef" && scanAttr(frame.attrs, "ref") !== undefined && !hasRelationshipId(frame.attrs)) {
        const named = scanAttr(frame.attrs, "sheet");
        const on = named === undefined ? host : resolveRetainedSheetName(pass.model, pass.parts, named, host.id);
        if (on) {
          edits.rangeAttribute(frame, "ref", on.name, false);
        }
      }
    },
    (frame, contentEnd, end) => {
      const parent = frame.parent;
      if (frame.ns === XM_NS) {
        if (frame.local === "f" && hasAncestor(frame, (at) => at.ns === X14_NS && X14_FORMULA_SCOPES.has(at.local))) {
          edits.formulaText(frame, contentEnd, host.id);
        } else if (frame.local === "sqref" && parent?.ns === X14_NS && X14_SQREF_OWNERS.has(parent.local)) {
          if (edits.rangeText(frame, contentEnd, host.name, true) === "gone") {
            parent.remove = true;
          }
        }
      } else if (frame.ns === X14_NS) {
        if (frame.local === "sparklines" && frame.items && frame.removedItems === frame.items && parent?.local === "sparklineGroup") {
          parent.remove = true;
        }
      } else if (frame.local === "formula" && hasAncestor(frame, (at) => at.local === "cfRule" && at.ns !== X14_NS)) {
        edits.formulaText(frame, contentEnd, host.id);
      } else if ((frame.local === "formula1" || frame.local === "formula2") && hasAncestor(frame, (at) => at.local === "dataValidation" && at.ns !== X14_NS)) {
        edits.formulaText(frame, contentEnd, host.id);
      } else if (EMPTY_CONTAINERS_REMOVED.has(frame.local) && frame.items && frame.removedItems === frame.items) {
        frame.remove = true;
      }
      if (frame.local === "dataValidations" && !frame.remove && frame.items && frame.removedItems) {
        edits.setCount(frame, "count", frame.items - frame.removedItems);
      }
      edits.close(frame, end);
    },
  );
  return edits;
}

const TABLE_RANGE_ELEMENTS: ReadonlySet<string> = new Set(["table", "autoFilter", "sortState", "sortCondition"]);

function tableEdits(xml: string, host: Sheet, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      if (TABLE_RANGE_ELEMENTS.has(frame.local)) {
        edits.rangeAttribute(frame, "ref", host.name, false);
      }
    },
    ignore,
  );
  return edits;
}

function pivotTableEdits(xml: string, host: Sheet, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      if (frame.local === "location" && frame.parent?.local === "pivotTableDefinition") {
        edits.rangeAttribute(frame, "ref", host.name, false);
      }
    },
    ignore,
  );
  return edits;
}

function pivotCacheEdits(xml: string, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      if (frame.local !== "worksheetSource" || scanAttr(frame.attrs, "name") !== undefined || hasRelationshipId(frame.attrs)) {
        return;
      }
      const named = scanAttr(frame.attrs, "sheet");
      const on = named === undefined ? undefined : resolveRetainedSheetName(pass.model, pass.parts, named);
      if (on) {
        edits.rangeAttribute(frame, "ref", on.name, false);
      }
    },
    ignore,
  );
  return edits;
}

/** Every `f` element of a chart (c:f, cx:f, c15:f) as a formula. */
function chartEdits(xml: string, hostSheetId: string, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(xml, ignore, (frame, contentEnd) => {
    if (frame.local === "f") {
      edits.formulaText(frame, contentEnd, hostSheetId);
    }
  });
  return edits;
}

/** `hosts[i]`: the id of the model sheet the source workbook's i-th sheet
 * is ("" for none). */
function definedNameEdits(xml: string, hosts: readonly string[], pass: Pass): PartEdits | undefined {
  if (!xml.includes("definedName")) {
    return undefined;
  }
  const edits = new PartEdits(xml, pass.mapper);
  walk(xml, ignore, (frame, contentEnd) => {
    // SPEC-128: _xlnm._FilterDatabase follows the model's autofilter (writer.ts).
    if (frame.local !== "definedName" || scanAttr(frame.attrs, "name")?.toLowerCase() === FILTER_DATABASE_NAME.toLowerCase()) {
      return;
    }
    const local = scanAttr(frame.attrs, "localSheetId")?.trim();
    const host = local !== undefined && /^\d+$/.test(local) ? (hosts[Number(local)] ?? "") : "";
    edits.formulaText(frame, contentEnd, host);
  });
  return edits;
}

function drawingEdits(xml: string, host: Sheet, pass: Pass): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(xml, ignore, (frame, contentEnd) => {
    if (frame.local !== "row" && frame.local !== "col") {
      return;
    }
    const corner = frame.parent;
    const anchor = corner?.parent;
    if (!corner || !anchor || (corner.local !== "from" && corner.local !== "to")) {
      return;
    }
    const moves =
      anchor.local === "twoCellAnchor" ? scanAttr(anchor.attrs, "editAs") !== "absolute" : anchor.local === "oneCellAnchor" && corner.local === "from";
    if (moves) {
      edits.indexText(frame, contentEnd, host.name, frame.local === "row" ? "row" : "col");
    }
  });
  return edits;
}

const cellKey = (row: number, col: number): string => `${row},${col}`;

/** A comments (`comment`) or threaded comments (`threadedComment`) part; the
 * original cells of removed legacy comments go into `removedCells`. */
function noteEdits(xml: string, host: Sheet, pass: Pass, local: "comment" | "threadedComment", removedCells?: Set<string>): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      if (frame.local !== local || edits.rangeAttribute(frame, "ref", host.name, true) !== "gone") {
        return;
      }
      frame.remove = true;
      const cell = parseEntry(scanAttr(frame.attrs, "ref")?.trim() ?? "");
      if (cell && removedCells) {
        removedCells.add(cellKey(cell.range.top, cell.range.left));
      }
    },
    (frame, _contentEnd, end) => edits.close(frame, end),
  );
  return edits;
}

function vmlEdits(xml: string, host: Sheet, pass: Pass, removedCells: ReadonlySet<string>): PartEdits {
  const edits = new PartEdits(xml, pass.mapper);
  walk(
    xml,
    (frame) => {
      if (frame.local === "ClientData" && frame.parent) {
        frame.parent.objectType = scanAttr(frame.attrs, "ObjectType") ?? "";
      }
    },
    (frame, contentEnd, end) => {
      const data = frame.parent?.local === "ClientData" ? frame.parent : undefined;
      if (data && (frame.local === "Row" || frame.local === "Column")) {
        const index = edits.indexText(frame, contentEnd, host.name, frame.local === "Row" ? "row" : "col");
        const shape = data.parent;
        if (shape && index !== undefined) {
          if (frame.local === "Row") {
            shape.noteRow = index;
          } else {
            shape.noteCol = index;
          }
        }
      } else if (data && frame.local === "Anchor") {
        edits.anchorText(frame, contentEnd, host.name);
      } else if (
        frame.local === "shape" &&
        frame.noteRow !== undefined &&
        frame.noteCol !== undefined &&
        (frame.objectType === "Note" || frame.objectType === "") &&
        removedCells.has(cellKey(frame.noteRow, frame.noteCol))
      ) {
        frame.remove = true;
      }
      edits.close(frame, end);
    },
  );
  return edits;
}

// --- worksheet parts without their cells ----------------------------------------------------

const SHEET_DATA_OPEN = encodeUtf8("<sheetData");
const SHEET_DATA_CLOSE = encodeUtf8("</sheetData>");
/** What stands in for the cells while a worksheet's tail is scanned. */
const EMPTY_SHEET_DATA = "<sheetData/>";
/** base64 characters first decoded at a worksheet part's head and tail
 * (multiples of 4), doubled until `<sheetData` / `</sheetData>` is found. */
const HEAD_PROBE = 16_384;
const TAIL_PROBE = 65_536;

function matchesAt(bytes: Uint8Array, needle: Uint8Array, at: number): boolean {
  for (let j = 0; j < needle.length; j += 1) {
    if (bytes[at + j] !== needle[j]) {
      return false;
    }
  }
  return true;
}

function lastIndexOfBytes(bytes: Uint8Array, needle: Uint8Array): number {
  for (let i = bytes.length - needle.length; i >= 0; i -= 1) {
    if (matchesAt(bytes, needle, i)) {
      return i;
    }
  }
  return -1;
}

/** The offset of the first `<sheetData` start tag (its next byte in
 * `bytes`), or -1. */
function sheetDataStart(bytes: Uint8Array): number {
  for (let i = 0; i + SHEET_DATA_OPEN.length < bytes.length; i += 1) {
    if (bytes[i] === 0x3c && matchesAt(bytes, SHEET_DATA_OPEN, i)) {
      const next = bytes[i + SHEET_DATA_OPEN.length];
      if (next === 0x3e || next === 0x2f || next === 0x20 || next === 0x09 || next === 0x0a || next === 0x0d) {
        return i;
      }
    }
  }
  return -1;
}

/** A worksheet part's text before `<sheetData` and after `</sheetData>`,
 * and its base64 up to the tail: `keep` (whole 3-byte groups) then the 0–2
 * bytes `carry`. */
interface WorksheetSplit {
  head: string;
  tail: string;
  keep: string;
  carry: Uint8Array;
}

function splitWorksheet(data: string): WorksheetSplit | undefined {
  if (data.length % 4 !== 0) {
    return undefined;
  }
  let suffix: { bytes: Uint8Array; start: number; close: number } | undefined;
  for (let chars = Math.min(TAIL_PROBE, data.length); ; chars = Math.min(chars * 2, data.length)) {
    const from = data.length - chars;
    const bytes = base64ToBytes(data.slice(from));
    const close = lastIndexOfBytes(bytes, SHEET_DATA_CLOSE);
    if (close >= 0) {
      suffix = { bytes, start: (from / 4) * 3, close };
      break;
    }
    if (from === 0) {
      return undefined;
    }
  }
  let prefix: { bytes: Uint8Array; open: number } | undefined;
  for (let chars = Math.min(HEAD_PROBE, data.length); ; chars = Math.min(chars * 2, data.length)) {
    const bytes = base64ToBytes(data.slice(0, chars));
    const open = sheetDataStart(bytes);
    if (open >= 0) {
      prefix = { bytes, open };
      break;
    }
    if (chars === data.length) {
      return undefined;
    }
  }
  if (prefix.open >= suffix.start + suffix.close) {
    return undefined;
  }
  const tailStart = suffix.start + suffix.close + SHEET_DATA_CLOSE.length;
  const groups = Math.floor(tailStart / 3);
  const head = decodeStrict(prefix.bytes.subarray(0, prefix.open));
  const tail = decodeStrict(suffix.bytes.subarray(suffix.close + SHEET_DATA_CLOSE.length));
  if (head === undefined || tail === undefined) {
    return undefined;
  }
  return { head, tail, keep: data.slice(0, groups * 4), carry: suffix.bytes.slice(groups * 3 - suffix.start, tailStart - suffix.start) };
}

function concatBytes(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}

// --- the run ------------------------------------------------------------------------------

/** One adjustment over the model's ORIGINAL parts: each part rewritten at
 * most once, the results committed copy-on-write. */
class Rewrite implements Pass {
  readonly model: SpreadsheetModel;
  readonly mapper: RetainedReferenceMapper;
  readonly parts: LazyParts;
  private readonly sources: ReadonlyMap<string, OpaquePart>;
  /** Part → whether it changed, once handled. */
  private readonly outcomes = new Map<string, boolean>();
  /** Part → its new base64. */
  private readonly rewritten = new Map<string, string>();

  constructor(model: SpreadsheetModel, mapper: RetainedReferenceMapper) {
    this.model = model;
    this.mapper = mapper;
    this.parts = new LazyParts(model.opaque);
    this.sources = new Map(model.opaque.map((part) => [part.name, part] as const));
  }

  has(name: string): boolean {
    return this.sources.has(name);
  }

  names(): Iterable<string> {
    return this.sources.keys();
  }

  relationships(owner: string): RelEntry[] {
    try {
      return relationshipsOf(this.parts, owner);
    } catch (error) {
      if (error instanceof XmlScanError) {
        return [];
      }
      throw error;
    }
  }

  sourceWorkbook(): SourceWorkbook | undefined {
    try {
      return readSourceWorkbook(this.parts);
    } catch (error) {
      if (error instanceof XmlScanError) {
        return undefined;
      }
      throw error;
    }
  }

  /** A part through `build`'s edits, once; whether it changed. */
  rewrite(name: string, build: (xml: string) => PartEdits | undefined): boolean {
    const known = this.outcomes.get(name);
    if (known !== undefined) {
      return known;
    }
    this.outcomes.set(name, false);
    const bytes = this.parts.get(name);
    const xml = bytes === undefined ? undefined : decodeStrict(bytes);
    if (xml === undefined) {
      return false;
    }
    let next = xml;
    try {
      const edits = build(xml);
      next = edits ? applySplices(xml, edits.splices()) : xml;
    } catch (error) {
      if (!(error instanceof XmlScanError)) {
        throw error;
      }
    }
    if (next === xml) {
      return false;
    }
    this.rewritten.set(name, bytesToBase64(encodeUtf8(next)));
    this.outcomes.set(name, true);
    return true;
  }

  /** A sheet's own part: its tail alone when the part splits around
   * `<sheetData>` (see the module header), else the whole text. */
  rewriteWorksheet(name: string, host: Sheet): boolean {
    const known = this.outcomes.get(name);
    if (known !== undefined) {
      return known;
    }
    const source = this.sources.get(name);
    let split: WorksheetSplit | undefined;
    try {
      split = source && splitWorksheet(source.data);
    } catch {
      split = undefined; // not canonical base64: the whole-part path decides
    }
    if (split) {
      if (!WORKSHEET_ELEMENTS.test(split.tail)) {
        this.outcomes.set(name, false);
        return false;
      }
      const offset = split.head.length + EMPTY_SHEET_DATA.length;
      let splices: Splice[] | undefined;
      try {
        splices = worksheetEdits(split.head + EMPTY_SHEET_DATA + split.tail, host, this).splices();
      } catch (error) {
        if (!(error instanceof XmlScanError)) {
          throw error;
        }
      }
      if (splices && splices.every((splice) => splice.start >= offset)) {
        const tail = applySplices(
          split.tail,
          splices.map((splice) => ({ start: splice.start - offset, end: splice.end - offset, text: splice.text })),
        );
        const changed = tail !== split.tail;
        if (changed) {
          this.rewritten.set(name, split.keep + bytesToBase64(concatBytes(split.carry, encodeUtf8(tail))));
        }
        this.outcomes.set(name, changed);
        return changed;
      }
    }
    return this.rewrite(name, (xml) => (WORKSHEET_ELEMENTS.test(xml) ? worksheetEdits(xml, host, this) : undefined));
  }

  /** `model.opaque` with the rewritten parts, copy-on-write. */
  commit(): void {
    if (this.rewritten.size === 0) {
      return;
    }
    this.model.opaque = this.model.opaque.map((part) => {
      const data = this.rewritten.get(part.name);
      return data === undefined ? part : { name: part.name, data };
    });
  }
}

const CHART_PART = /^xl\/charts\/[^/]+\.xml$/;
const CHART_STYLE_PART = /\/(?:style|colors)\d*\.xml$/;
const PIVOT_CACHE_PART = /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/;

/** Adjusts every retained xlsx reference through `mapper` (see the module
 * header); returns the ids of the sheets whose own retained text changed. */
export function rewriteXlsxRetainedReferences(model: SpreadsheetModel, mapper: RetainedReferenceMapper): Set<string> {
  const changed = new Set<string>();
  if (model.opaque.length === 0) {
    return changed;
  }
  const run = new Rewrite(model, mapper);
  const chartHosts = new Map<string, string>();
  for (const sheet of model.sheets) {
    const part = sheet.partName;
    if (part === undefined || !run.has(part)) {
      continue;
    }
    let own = run.rewriteWorksheet(part, sheet);
    const removedCells = new Set<string>();
    const vml: string[] = [];
    for (const rel of run.relationships(part)) {
      const target = targetPart(part, rel);
      const kind = relationshipKind(rel.type);
      if (target === undefined) {
        continue;
      }
      if (kind === "table") {
        own = run.rewrite(target, (xml) => tableEdits(xml, sheet, run)) || own;
      } else if (kind === "pivotTable") {
        own = run.rewrite(target, (xml) => pivotTableEdits(xml, sheet, run)) || own;
      } else if (kind === "comments") {
        own = run.rewrite(target, (xml) => noteEdits(xml, sheet, run, "comment", removedCells)) || own;
      } else if (kind === "threadedComment") {
        own = run.rewrite(target, (xml) => noteEdits(xml, sheet, run, "threadedComment")) || own;
      } else if (kind === "vmlDrawing") {
        vml.push(target); // after the comments, whose removals it follows
      } else if (kind === "drawing") {
        own = run.rewrite(target, (xml) => drawingEdits(xml, sheet, run)) || own;
        for (const chartRel of run.relationships(target)) {
          const chartKind = relationshipKind(chartRel.type);
          const chart = chartKind === "chart" || chartKind === "chartEx" ? targetPart(target, chartRel) : undefined;
          if (chart !== undefined) {
            const host = chartHosts.get(chart) ?? sheet.id;
            chartHosts.set(chart, host);
            own = run.rewrite(chart, (xml) => chartEdits(xml, host, run)) || own;
          }
        }
      }
    }
    for (const target of vml) {
      own = run.rewrite(target, (xml) => vmlEdits(xml, sheet, run, removedCells)) || own;
    }
    if (own) {
      changed.add(sheet.id);
    }
  }

  const caches = new Set<string>();
  const workbook = run.sourceWorkbook();
  if (workbook) {
    const hosts = workbook.sheets.map((source) =>
      source.partName === "" ? "" : (model.sheets.find((sheet) => sheet.partName === source.partName)?.id ?? ""),
    );
    run.rewrite(workbook.path, (xml) => definedNameEdits(xml, hosts, run));
    for (const rel of workbook.rels) {
      const target = relationshipKind(rel.type) === "pivotCacheDefinition" ? targetPart(workbook.path, rel) : undefined;
      if (target !== undefined) {
        caches.add(target);
      }
    }
  }
  for (const name of run.names()) {
    if (PIVOT_CACHE_PART.test(name)) {
      caches.add(name);
    } else if (CHART_PART.test(name) && !CHART_STYLE_PART.test(name) && !chartHosts.has(name)) {
      run.rewrite(name, (xml) => chartEdits(xml, "", run));
    }
  }
  for (const cache of caches) {
    run.rewrite(cache, (xml) => pivotCacheEdits(xml, run));
  }
  run.commit();
  return changed;
}
