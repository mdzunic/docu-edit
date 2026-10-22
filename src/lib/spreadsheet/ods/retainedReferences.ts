// The .ods half of SPEC-124's retained references (see
// ../retainedReferences.ts for the mapper contract): every reference-bearing
// attribute of the table below, in retained text this product does not model
// yet, moved through the RetainedReferenceMapper by exact offsets — only the
// changed attribute values are spliced (and the elements and attributes the
// rules remove); every other byte stays verbatim.
//
// | Where | Element@attribute | Kind |
// | envelope prefix | `table:label-range@table:label-cell-range-address`, `@table:data-cell-range-address` | list, emptied → element removed |
// | suffix, not a database range the model owns | `table:database-range@table:target-range-address` (emptied → removed); `table:filter@table:condition-source-range-address` | list |
// | suffix | `table:data-pilot-table@table:target-range-address`; `table:source-cell-range@table:cell-range-address` | list |
// | suffix | `table:consolidation@table:source-cell-range-addresses`, `@table:target-cell-address` | list; anchor |
// | inside tables, `Cell.preservedXml`, `Sheet.drawings` source XML | `draw:object@draw:notify-on-update-of-ranges` | list |
// | inside tables | `table:scenario@table:scenario-ranges` | list |
// | `Object N/content.xml` | `chart:plot-area@table:cell-range-address`, `chart:series@chart:values-cell-range-address`, `@chart:label-cell-address`, `chart:domain@table:cell-range-address`, `chart:categories@table:cell-range-address` | list |
//
// - A list is rangeAddress.ts's grammar, split keeping its space runs. Each
//   entry is parsed; its sheet as written (source time) resolved through
//   resolveRetainedSheetName — a sheet-less `.A1` is on the host — and moved
//   by `mapper.range(<model name>, range)`. A 3-D entry, one that does not
//   parse or names no model sheet, and every entry the mapper does not move
//   keep their exact bytes. A moved entry has only its cell parts rewritten:
//   the sheet text as written, each part's `$` kept, letters upper-case, a
//   part that did not move verbatim; a whole column ignores the row axis and
//   a whole row the column axis. An entry mapped to null is dropped (with
//   one space run); a list emptied that way removes its element or attribute
//   where the table says so and keeps its original text everywhere else.
// - An anchor is one cell: when its cell is gone it becomes the cell the band
//   now starts at — `mapper.index` per axis.
// - Hosts. Inside a table, the table's sheet: for content.xml's source
//   tables the sheet resolveRetainedSheetName gives the table's name, for a
//   sheet's own sourceXml and its cells' preservedXml that sheet, its id
//   selecting a copy's override. In the envelope, the sheet of the element's
//   own base-cell-address, else none (a sheet-less entry is then left
//   alone). An `Object N`: the sheet whose sourceXml shows it
//   (`objectsReferenced`), a copy's override only for a copy that alone
//   shows it (the writer's rule).
// - SPEC-128: a `table:database-range` that is a sheet's autofilter (the
//   claim of ods/databaseRanges.ts, decided on the retained text alone) is
//   the model's — `Sheet.autoFilter`, which the `autofilter` reference
//   adjuster moves and the writer serialises — so nothing inside it (its
//   target, its filter's criteria range) is adjusted here: nothing is
//   adjusted twice. Every other database range still is. The filter of a
//   claimed range that does not decompose is the model's verbatim
//   `autoFilter.odsUnmappedFilter`; its `table:condition-source-range-address`
//   (a criteria range, on any sheet) is moved by the `autofilter` adjusters
//   of structure.ts and move.ts through adjustOdsFilterCriteriaRanges, by
//   this module's list rules, on the model's copy: the autofilter replaced
//   without its sourceXml (the writer then serialises that filter), its
//   sheet marked dirty.
// - Retired by SPEC-129 (its `conditionalFormats` adjusters, structure.ts's
//   and move.ts's, move the model; the writer re-serialises a dirty sheet's
//   `calcext:conditional-formats` from it):
//   `calcext:conditional-format@calcext:target-range-address`,
//   `calcext:condition@calcext:value` and `@calcext:base-cell-address`, and
//   `calcext:formatting-entry@calcext:value` — left verbatim, an emptied
//   format and an emptied `calcext:conditional-formats` never removed, so
//   nothing is adjusted twice.
// - Retired by SPEC-130 (the `validations` adjusters of structure.ts and
//   move.ts move Sheet.validations; the writer regenerates the prefix's
//   `table:content-validations` from the model, a changed rule serialised,
//   an unchanged one from its sourceXml):
//   `table:content-validation@table:base-cell-address` and
//   `@table:condition` — left verbatim, so nothing is adjusted twice.
// - Retired by SPEC-131 (its `drawings` reference adjuster moves
//   Sheet.drawings' anchors; the table writer re-emits a dirty table's
//   drawings from the model, an imported one's anchor patched):
//   `draw:*@table:end-cell-address` — left verbatim, so nothing is adjusted
//   twice. A drawing's own source XML (an imported `draw:frame` > `draw:object`)
//   still has its `draw:notify-on-update-of-ranges` moved, the drawing
//   replaced copy-on-write.
// - Narrowed by SPEC-132 (the `charts` adjusters of structure.ts and move.ts
//   move a modelled chart's `refs`, syncChartNotifyRanges its frame; the
//   writer patches or regenerates its object from the model): the
//   `Object N/content.xml` rows and the `draw:object@draw:notify-on-update-
//   of-ranges` row apply only to UNMODELLED charts — an object some drawing's
//   modelled chart content (isModelledChart) names as its sourcePart is
//   skipped, as is the notify list of a `draw:object` showing one (in tables,
//   preservedXml) and every drawing whose content is a modelled chart — so
//   nothing is adjusted twice. Out-of-catalogue charts (`chart:surface`) are
//   still rewritten here. SPEC-125's move path runs this same rewrite.
// - Retired by SPEC-136 (print areas are SPEC-133's `_xlnm.Print_Area`
//   names, moved by the `names` adjusters; the writer writes the table's
//   `table:print-ranges` from the model's name — ods/pageSetup.ts):
//   `table:table@table:print-ranges` — left verbatim, so nothing is
//   adjusted twice.
// - Retired by SPEC-133 (the `names` adjusters of structure.ts and move.ts
//   move the model's defined names, which the writer reconciles with the
//   retained `table:named-expressions` of the suffix and of each table —
//   ods/namedExpressions.ts — writing the model's whenever they differ):
//   `table:named-range@table:cell-range-address` and
//   `@table:base-cell-address`, `table:named-expression@table:expression`
//   and `@table:base-cell-address` — left verbatim, an emptied named range
//   never removed, so nothing is adjusted twice. With them went the only
//   formula-valued location: every row left is a list or an anchor. The
//   `table:database-range` row stays (SPEC-128's claim above).
// - `table:tracked-changes` is history: nothing inside it is touched.
// - Elements and attributes are matched by their qualified names as written
//   (a fragment's prefixes are unbound), as sheetNamePass.ts does.
// - Where it lives, replaced copy-on-write: content.xml is the envelope
//   prefix + the source tables + the suffix — the envelope scanned as one
//   document (the tables between are whole elements), then each table that
//   names a location; it gets a new OpaquePart, and a new
//   `odsContentEnvelope` when its prefix or suffix changed, so content.xml
//   is always exactly prefix + tables + suffix. A sheet gets a new object
//   for a new sourceXml (a copy's is not in content.xml) or rows — only the
//   rows and cells whose preservedXml changed are replaced. An `Object N`
//   gets a new OpaquePart; `opaque` a new array when any part changed. Text
//   naming none of the table's attributes is never scanned.
import { bytesToBase64, type OpaquePart } from "../../core/packageParts";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import { colToLetters, type GridRange } from "../address";
import type { RefPart, ReferenceNode } from "../formula/references";
import { scanQuoted } from "../formula/tokens";
import { isOriginal, MAX_COLS, MAX_ROWS, type Cell, type Row, type Sheet, type SheetDrawing, type SpreadsheetModel } from "../model";
import type { RetainedReferenceMapper } from "../retainedReferences";
import { resolveRetainedSheetName } from "../sourceNames";
import { applySplices, findAttribute, type Splice } from "../xmlSplice";
import { decodeUtf8, encodeUtf8, LazyParts } from "../xlsx/package";
import { odsOwnedDatabaseRangeSpans } from "./databaseRanges";
import { isModelledChartDrawing } from "./drawingXml";
import { objectsReferenced, sourceTables } from "./package";
import { parseOdsRangeAddress, splitOdsRangeAddressList } from "./rangeAddress";

// --- the location table -----------------------------------------------------------------

type Kind = "list" | "anchor";

interface Location {
  attribute: string;
  kind: Kind;
  /** What an emptied list removes; absent: the list keeps its text. */
  emptied?: "element" | "attribute";
}

const list = (attribute: string, emptied?: "element" | "attribute"): Location =>
  emptied === undefined ? { attribute, kind: "list" } : { attribute, kind: "list", emptied };
const anchor = (attribute: string): Location => ({ attribute, kind: "anchor" });


const CONTENT_LOCATIONS: ReadonlyMap<string, readonly Location[]> = new Map([
  ["table:label-range", [list("table:label-cell-range-address", "element"), list("table:data-cell-range-address", "element")]],
  ["table:database-range", [list("table:target-range-address", "element")]],
  ["table:filter", [list("table:condition-source-range-address")]],
  ["table:data-pilot-table", [list("table:target-range-address")]],
  ["table:source-cell-range", [list("table:cell-range-address")]],
  ["table:consolidation", [list("table:source-cell-range-addresses"), anchor("table:target-cell-address")]],
  ["table:scenario", [list("table:scenario-ranges")]],
  ["draw:object", [list("draw:notify-on-update-of-ranges")]],
]);

const CHART_LOCATIONS: ReadonlyMap<string, readonly Location[]> = new Map([
  ["chart:plot-area", [list("table:cell-range-address")]],
  ["chart:series", [list("chart:values-cell-range-address"), list("chart:label-cell-address")]],
  ["chart:domain", [list("table:cell-range-address")]],
  ["chart:categories", [list("table:cell-range-address")]],
]);

type LocationTable = (element: string) => readonly Location[] | undefined;

const contentLocations: LocationTable = (element) => CONTENT_LOCATIONS.get(element);
/** A verbatim ods filter (`autoFilter.odsUnmappedFilter`): its criteria range. */
const filterLocations: LocationTable = (element) => (element === "table:filter" ? CONTENT_LOCATIONS.get(element) : undefined);
const CRITERIA_SOURCE = /table:condition-source-range-address\s*=/;
const chartLocations: LocationTable = (element) => CHART_LOCATIONS.get(element);

/** The prefilter: any of a table's attribute names followed by its `=` —
 * one pass (a bare name would also match a longer one it begins, such as
 * a filter's `table:condition-source` for `table:condition`). */
function attributePattern(table: ReadonlyMap<string, readonly Location[]>): RegExp {
  const names = new Set([...table.values()].flat().map((location) => location.attribute));
  return new RegExp(`(?:${[...names].map((name) => name.replace(/[.-]/g, "\\$&")).join("|")})\\s*=`);
}

const CONTENT_ATTRIBUTES = attributePattern(CONTENT_LOCATIONS);
const CHART_ATTRIBUTES = attributePattern(CHART_LOCATIONS);

/** Whether `text` holds an attribute the pattern names (never scanned when not). */
function mentions(text: string, attributes: RegExp): boolean {
  return attributes.test(text);
}

/** A part's text and the byte-order mark it began with: the decoder drops
 * one (as the reader's does), and a rewrite writes it back. */
function decodedPart(bytes: Uint8Array): { bom: string; xml: string } {
  return { bom: bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? "\uFEFF" : "", xml: decodeUtf8(bytes) };
}

const TRACKED_CHANGES = "table:tracked-changes";
const CONTENT_XML = "content.xml";
const OBJECT_CONTENT = /^(Object [^/]+)\/content\.xml$/;

// --- raw attribute text -----------------------------------------------------------------

/** An attribute value as an XML parser reads it — entities decoded, a
 * literal tab or line end a space — and the raw offset each decoded index
 * starts at (`at(text.length)` is the raw length), so a rewrite can keep
 * every raw byte it does not replace. */
interface Decoded {
  text: string;
  at(index: number): number;
}

const ENTITY = /&(?:#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/y;
const NAMED_ENTITIES: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

/** One entity's text, undefined when it names nothing. */
function entityText(entity: string): string | undefined {
  const body = entity.slice(1, -1);
  if (!body.startsWith("#")) {
    return Object.hasOwn(NAMED_ENTITIES, body) ? NAMED_ENTITIES[body] : undefined;
  }
  const code = body.startsWith("#x") ? Number.parseInt(body.slice(2), 16) : Number.parseInt(body.slice(1), 10);
  return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : undefined;
}

/** A raw (still-escaped) attribute value decoded (see Decoded). */
function decodeRaw(raw: string): Decoded {
  if (!/[&\t\n\r]/.test(raw)) {
    return { text: raw, at: (index) => index };
  }
  let text = "";
  const offsets: number[] = [];
  let i = 0;
  while (i < raw.length) {
    const code = raw.charCodeAt(i);
    if (code === 0x26) {
      ENTITY.lastIndex = i;
      const entity = ENTITY.exec(raw)?.[0];
      const decoded = entity === undefined ? undefined : entityText(entity);
      if (entity !== undefined && decoded !== undefined) {
        for (let unit = 0; unit < decoded.length; unit += 1) {
          offsets.push(i);
        }
        text += decoded;
        i += entity.length;
        continue;
      }
    }
    offsets.push(i);
    text += code === 0x09 || code === 0x0a || code === 0x0d ? " " : raw[i];
    i += 1;
  }
  offsets.push(raw.length);
  return { text, at: (index) => offsets[index]! };
}

// --- one address ------------------------------------------------------------------------

/** A raw slice of an address by decoded offsets into it. */
type RawOf = (from: number, to: number) => string;

/** Where a moved address's parts go. */
interface Moved {
  start: RefPart;
  end?: RefPart;
}

/** Where the CELL part of the endpoint at `i` sits: after its sheet's `.`
 * (the address has parsed, so the shape is known), up to a `:` or the end. */
function cellSpan(address: string, i: number): { start: number; end: number } {
  let j = i;
  if (address.charCodeAt(j) !== 0x2e) {
    if (address.charCodeAt(j) === 0x24) {
      j += 1;
    }
    if (address.charCodeAt(j) === 0x27) {
      j = scanQuoted(address, j)?.end ?? address.length;
    } else {
      while (j < address.length && address.charCodeAt(j) !== 0x2e) {
        j += 1;
      }
    }
  }
  const start = j + 1;
  let end = start;
  while (end < address.length && address.charCodeAt(end) !== 0x3a) {
    end += 1;
  }
  return { start, end };
}

function partText(part: RefPart): string {
  const col = part.col === undefined ? "" : `${part.colAbs ? "$" : ""}${colToLetters(part.col)}`;
  const row = part.row === undefined ? "" : `${part.rowAbs ? "$" : ""}${part.row + 1}`;
  return col + row;
}

const samePosition = (a: RefPart, b: RefPart | undefined): boolean => b !== undefined && a.row === b.row && a.col === b.col;

/** The grid area an address covers (a whole column or row spans its axis). */
function gridOf(ref: ReferenceNode): GridRange {
  const end = ref.end ?? ref.start;
  const rows = ref.start.row === undefined || end.row === undefined ? [0, MAX_ROWS - 1] : [ref.start.row, end.row];
  const cols = ref.start.col === undefined || end.col === undefined ? [0, MAX_COLS - 1] : [ref.start.col, end.col];
  return { top: Math.min(rows[0]!, rows[1]!), bottom: Math.max(rows[0]!, rows[1]!), left: Math.min(cols[0]!, cols[1]!), right: Math.max(cols[0]!, cols[1]!) };
}

/** An address's parts placed on `to`, each axis keeping its orientation; a
 * part with no row (a whole column) or no column (a whole row) gains none. */
function placed(ref: ReferenceNode, to: GridRange): Moved {
  const start: RefPart = { ...ref.start };
  const end: RefPart | undefined = ref.end && { ...ref.end };
  if (ref.start.row !== undefined) {
    const reversed = ref.end?.row !== undefined && ref.start.row > ref.end.row;
    start.row = reversed ? to.bottom : to.top;
    if (end?.row !== undefined) {
      end.row = reversed ? to.top : to.bottom;
    }
  }
  if (ref.start.col !== undefined) {
    const reversed = ref.end?.col !== undefined && ref.start.col > ref.end.col;
    start.col = reversed ? to.right : to.left;
    if (end?.col !== undefined) {
      end.col = reversed ? to.left : to.right;
    }
  }
  return end ? { start, end } : { start };
}

/** The address with its moved cell parts rewritten — its sheet text, the
 * `:` and every part that did not move as written. */
function rewriteParts(address: string, rawOf: RawOf, ref: ReferenceNode, moved: Moved): string {
  const part = (before: RefPart | undefined, after: RefPart | undefined, span: { start: number; end: number }): string =>
    before && after && !samePosition(before, after) ? partText(after) : rawOf(span.start, span.end);
  const first = cellSpan(address, 0);
  let out = rawOf(0, first.start) + part(ref.start, moved.start, first);
  if (first.end < address.length) {
    const second = cellSpan(address, first.end + 1);
    out += rawOf(first.end, second.start) + part(ref.end, moved.end, second) + rawOf(second.end, address.length);
  }
  return out;
}

/** A list's pieces rejoined without the dropped entries (null): each kept
 * entry after the first preceded by the run written before it, a leading and
 * a trailing run kept; null when no entry is left. */
function joinList(pieces: readonly string[], out: readonly (string | null)[]): string | null {
  const last = pieces.length - 1;
  let text = pieces[0] === "" && last > 0 ? out[1]! : "";
  let kept = 0;
  for (let k = 0; k <= last; k += 2) {
    const entry = out[k];
    if (pieces[k] === "" || entry === null || entry === undefined) {
      continue;
    }
    if (kept > 0) {
      text += out[k - 1]!;
    }
    text += entry;
    kept += 1;
  }
  if (kept === 0) {
    return null;
  }
  return last > 0 && pieces[last] === "" ? text + out[last - 1]! : text;
}

// --- the rewrite --------------------------------------------------------------------------

/** What a sheet-less address is on, and how source-time names resolve. */
interface Host {
  sheet?: Sheet;
  /** The id resolveRetainedSheetName selects a copy's override by. */
  id?: string;
  /** Envelope text: each element's host is the sheet of its own
   * base-cell-address. */
  byBaseCell?: true;
}

const ENVELOPE_HOST: Host = { byBaseCell: true };

interface Frame {
  start: number;
  /** The element is removed when it closes. */
  remove: boolean;
  /** Nothing inside is adjusted (tracked changes, a removed element). */
  quiet: boolean;
}

class OdsRetainedRewrite {
  private readonly model: SpreadsheetModel;
  private readonly mapper: RetainedReferenceMapper;
  private readonly parts: LazyParts;
  private readonly resolved = new Map<string, Sheet | null>();
  private readonly shown = new Map<Sheet, string[]>();

  constructor(model: SpreadsheetModel, mapper: RetainedReferenceMapper) {
    this.model = model;
    this.mapper = mapper;
    this.parts = new LazyParts(model.opaque);
    for (const sheet of model.sheets) {
      for (const drawing of sheet.drawings ?? []) {
        if (isModelledChartDrawing(drawing) && drawing.content.sourcePart !== undefined) {
          this.modelledParts.add(drawing.content.sourcePart);
        }
      }
    }
  }

  /** SPEC-132: the `Object N/content.xml` parts some modelled chart drawing
   * names as its sourcePart — their addresses, and the notify list of a
   * frame showing one, are the `charts` adjusters' (never adjusted here). */
  private readonly modelledParts = new Set<string>();

  /** Whether a `draw:object` start tag shows a modelled chart's object. */
  private showsModelledChart(tag: string): boolean {
    if (this.modelledParts.size === 0) {
      return false;
    }
    const object = objectsReferenced(tag)[0];
    return object !== undefined && this.modelledParts.has(`${object}/content.xml`);
  }

  run(): Set<string> {
    const model = this.model;
    const changed = new Set<string>();
    const ownSplices = new Map<Sheet, Splice[]>();
    let sheetsChanged = false;
    const sheets = model.sheets.map((sheet) => {
      const next = this.sheet(sheet, ownSplices);
      if (next !== sheet) {
        sheetsChanged = true;
        changed.add(sheet.id);
      }
      return next;
    });
    const replaced = this.objects(changed);
    const content = this.content(ownSplices);
    if (content) {
      replaced.set(CONTENT_XML, { name: CONTENT_XML, data: bytesToBase64(encodeUtf8(content.bom + content.xml)) });
      if (content.envelope) {
        model.odsContentEnvelope = content.envelope;
      }
    }
    if (replaced.size > 0) {
      model.opaque = model.opaque.map((part) => replaced.get(part.name) ?? part);
    }
    if (sheetsChanged) {
      model.sheets = sheets;
    }
    return changed;
  }

  /** Each verbatim ods filter whose criteria range moves, moved (see
   * adjustOdsFilterCriteriaRanges). Sheet-less entries stay, as in the
   * suffix; names resolve with the sheet's id (a copy's override). */
  filterSources(): void {
    let sheets: Sheet[] | undefined;
    this.model.sheets.forEach((sheet, i) => {
      const filter = sheet.autoFilter;
      const xml = filter?.odsUnmappedFilter;
      if (!filter || xml === undefined || !mentions(xml, CRITERIA_SOURCE)) {
        return;
      }
      const odsUnmappedFilter = applySplices(xml, this.splices(xml, { id: sheet.id }, filterLocations));
      if (odsUnmappedFilter === xml) {
        return;
      }
      const { sourceXml: _source, ...rest } = filter;
      sheets ??= this.model.sheets.slice();
      sheets[i] = { ...sheet, autoFilter: { ...rest, odsUnmappedFilter }, dirty: true };
    });
    if (sheets) {
      this.model.sheets = sheets;
    }
  }

  /** A sheet with its sourceXml and its cells' preservedXml adjusted (the
   * same object when neither changes); records the sourceXml's splices. */
  private sheet(sheet: Sheet, ownSplices: Map<Sheet, Splice[]>): Sheet {
    const host: Host = { sheet, id: sheet.id };
    let next = sheet;
    if (sheet.sourceXml !== undefined && mentions(sheet.sourceXml, CONTENT_ATTRIBUTES)) {
      const splices = this.splices(sheet.sourceXml, host, contentLocations);
      ownSplices.set(sheet, splices);
      const sourceXml = applySplices(sheet.sourceXml, splices);
      if (sourceXml !== sheet.sourceXml) {
        next = { ...next, sourceXml };
      }
    }
    const rows = this.rows(sheet.rows, host);
    if (rows !== sheet.rows) {
      next = { ...next, rows };
    }
    const drawings = this.drawings(sheet.drawings, host);
    return drawings === sheet.drawings ? next : { ...next, drawings };
  }

  /** SPEC-131: the drawings with each one's source XML adjusted,
   * copy-on-write (the same array when none changes). */
  private drawings(drawings: SheetDrawing[] | undefined, host: Host): SheetDrawing[] | undefined {
    let out: SheetDrawing[] | undefined;
    drawings?.forEach((drawing, i) => {
      const xml = drawing.sourceXml;
      if (xml === undefined || isModelledChartDrawing(drawing) || !mentions(xml, CONTENT_ATTRIBUTES)) {
        return;
      }
      const splices = this.splices(`<w>${xml}</w>`, host, contentLocations).map((splice) => ({ ...splice, start: splice.start - 3, end: splice.end - 3 }));
      const sourceXml = applySplices(xml, splices);
      if (sourceXml !== xml) {
        out ??= drawings.slice();
        out[i] = { ...drawing, sourceXml };
      }
    });
    return out ?? drawings;
  }

  /** Rows with every cell's preservedXml adjusted, copy-on-write. */
  private rows(rows: Row[], host: Host): Row[] {
    let out: Row[] | undefined;
    for (let r = 0; r < rows.length; r += 1) {
      const row = rows[r]!;
      let cells: Cell[] | undefined;
      for (let c = 0; c < row.cells.length; c += 1) {
        const cell = row.cells[c]!;
        const xml = cell.preservedXml;
        if (xml === undefined || !mentions(xml, CONTENT_ATTRIBUTES)) {
          continue;
        }
        // A fragment: scanned inside a dummy root.
        const splices = this.splices(`<w>${xml}</w>`, host, contentLocations).map((splice) => ({ ...splice, start: splice.start - 3, end: splice.end - 3 }));
        const preservedXml = applySplices(xml, splices);
        if (preservedXml !== xml) {
          cells ??= row.cells.slice();
          cells[c] = { ...cell, preservedXml };
        }
      }
      if (cells) {
        out ??= rows.slice();
        out[r] = { ...row, cells };
      }
    }
    return out ?? rows;
  }

  /** The adjusted content.xml and, when its prefix or suffix changed, the
   * new envelope; undefined when nothing changed (or the part does not hold
   * the envelope, which is then not touched either). */
  private content(ownSplices: ReadonlyMap<Sheet, Splice[]>): { bom: string; xml: string; envelope?: { prefix: string; suffix: string } } | undefined {
    const envelope = this.model.odsContentEnvelope;
    const bytes = envelope && this.parts.get(CONTENT_XML);
    if (!envelope || !bytes) {
      return undefined;
    }
    const { bom, xml } = decodedPart(bytes);
    const { prefix, suffix } = envelope;
    const tail = xml.length - suffix.length;
    if (tail < prefix.length || !xml.startsWith(prefix) || !xml.endsWith(suffix)) {
      return undefined;
    }
    const prefixSplices: Splice[] = [];
    const suffixSplices: Splice[] = [];
    const outside = prefix + suffix;
    const tables = sourceTables(xml, prefix.length, suffix.length);
    if (mentions(outside, CONTENT_ATTRIBUTES)) {
      // SPEC-128: the database ranges the model owns are left alone.
      const owned = odsOwnedDatabaseRangeSpans(suffix, tables.map((table) => table.name));
      for (const splice of this.splices(outside, ENVELOPE_HOST, contentLocations)) {
        // One straddling the two would be an ancestor of the tables: never a location.
        if (splice.end <= prefix.length) {
          prefixSplices.push(splice);
        } else if (splice.start >= prefix.length) {
          const inSuffix = { ...splice, start: splice.start - prefix.length, end: splice.end - prefix.length };
          if (!owned.some((span) => inSuffix.start >= span.start && inSuffix.end <= span.end)) {
            suffixSplices.push(inSuffix);
          }
        }
      }
    }
    const tableSplices: Splice[] = [];
    for (const table of tables) {
      const text = xml.slice(table.start, table.end);
      if (!mentions(text, CONTENT_ATTRIBUTES)) {
        continue;
      }
      const sheet = this.sheetNamed(table.name, undefined);
      // An original's sourceXml is this text, already adjusted with the same host.
      const splices = (sheet?.sourceXml === text ? ownSplices.get(sheet) : undefined) ?? this.splices(text, sheet ? { sheet, id: sheet.id } : {}, contentLocations);
      const shift = table.start - prefix.length;
      for (const splice of splices) {
        tableSplices.push({ ...splice, start: splice.start + shift, end: splice.end + shift });
      }
    }
    if (prefixSplices.length + suffixSplices.length + tableSplices.length === 0) {
      return undefined;
    }
    const nextPrefix = applySplices(prefix, prefixSplices);
    const nextSuffix = applySplices(suffix, suffixSplices);
    const next = nextPrefix + applySplices(xml.slice(prefix.length, tail), tableSplices) + nextSuffix;
    if (next === xml) {
      return undefined;
    }
    return nextPrefix === prefix && nextSuffix === suffix ? { bom, xml: next } : { bom, xml: next, envelope: { prefix: nextPrefix, suffix: nextSuffix } };
  }

  /** Each embedded object's chart ranges adjusted: the replaced parts, and
   * the sheets showing a changed one added to `changed`. */
  private objects(changed: Set<string>): Map<string, OpaquePart> {
    const replaced = new Map<string, OpaquePart>();
    for (const part of this.model.opaque) {
      const object = OBJECT_CONTENT.exec(part.name)?.[1];
      const bytes = object === undefined ? undefined : this.parts.get(part.name);
      if (object === undefined || !bytes || this.modelledParts.has(part.name)) {
        continue;
      }
      const { bom, xml } = decodedPart(bytes);
      if (!mentions(xml, CHART_ATTRIBUTES)) {
        continue;
      }
      const showing = this.model.sheets.filter((sheet) => this.objectsOf(sheet).includes(object));
      const only = showing.length === 1 ? showing[0] : undefined;
      const shared = showing.find(isOriginal) ?? showing[0];
      const host: Host = only ? { sheet: only, id: only.id } : shared ? { sheet: shared } : {};
      const next = applySplices(xml, this.splices(xml, host, chartLocations));
      if (next === xml) {
        continue;
      }
      replaced.set(part.name, { name: part.name, data: bytesToBase64(encodeUtf8(bom + next)) });
      for (const sheet of showing) {
        changed.add(sheet.id);
      }
    }
    return replaced;
  }

  private objectsOf(sheet: Sheet): string[] {
    let objects = this.shown.get(sheet);
    if (!objects) {
      objects = sheet.sourceXml === undefined ? [] : objectsReferenced(sheet.sourceXml);
      this.shown.set(sheet, objects);
    }
    return objects;
  }

  /** The model sheet a source-time name in retained text means. */
  private sheetNamed(name: string, hostId: string | undefined): Sheet | undefined {
    const key = `${hostId ?? ""}\u0000${name.toUpperCase()}`;
    let sheet = this.resolved.get(key);
    if (sheet === undefined) {
      sheet = resolveRetainedSheetName(this.model, this.parts, name, hostId) ?? null;
      this.resolved.set(key, sheet);
    }
    return sheet ?? undefined;
  }

  /** The splices adjusting every location of `xml` (a document, scanned
   * once), in its offsets; none when it does not scan. */
  private splices(xml: string, host: Host, table: LocationTable): Splice[] {
    const splices: Splice[] = [];
    const stack: Frame[] = [];
    let quiet = 0;
    const removeElement = (start: number, end: number): void => {
      // Everything recorded since the element opened lies inside it.
      while (splices.length > 0 && splices[splices.length - 1]!.start >= start) {
        splices.pop();
      }
      splices.push({ start, end, text: "" });
    };
    try {
      scanXml(xml, (event) => {
        if (event.type === "text") {
          return undefined;
        }
        if (event.type === "close") {
          const frame = stack.pop();
          if (!frame) {
            return undefined;
          }
          if (frame.quiet) {
            quiet -= 1;
          }
          if (frame.remove) {
            removeElement(frame.start, event.end);
          }
          return undefined;
        }
        const tracked = event.name === TRACKED_CHANGES;
        let remove = false;
        // SPEC-132: a modelled chart's frame notify list is the charts adjusters'.
        const modelledFrame = event.name === "draw:object" && this.showsModelledChart(xml.slice(event.start, event.end));
        const locations = quiet > 0 || tracked || modelledFrame ? undefined : table(event.name);
        if (locations) {
          const result = this.element(xml.slice(event.start, event.end), event.start, locations, host);
          if (result === "remove") {
            remove = true;
          } else {
            splices.push(...result);
          }
        }
        if (event.selfClosing) {
          if (remove) {
            removeElement(event.start, event.end);
          }
          return undefined;
        }
        const frame: Frame = { start: event.start, remove, quiet: tracked || remove };
        if (frame.quiet) {
          quiet += 1;
        }
        stack.push(frame);
        return undefined;
      });
    } catch (error) {
      if (error instanceof XmlScanError) {
        return [];
      }
      throw error;
    }
    return splices;
  }

  /** One location element's start tag (at `tagStart`): the splices of its
   * changed attributes, or "remove" when an emptied list removes it. */
  private element(tag: string, tagStart: number, locations: readonly Location[], host: Host): Splice[] | "remove" {
    const own = host.byBaseCell ? this.baseCellHost(tag) : host;
    const splices: Splice[] = [];
    for (const location of locations) {
      const span = findAttribute(tag, location.attribute);
      if (!span) {
        continue;
      }
      const next = this.list(span.raw, own, location.kind);
      if (next === undefined) {
        continue;
      }
      if (next === null) {
        if (location.emptied === "element") {
          return "remove";
        }
        if (location.emptied === "attribute") {
          splices.push({ start: tagStart + span.start, end: tagStart + span.end, text: "" });
        }
        continue;
      }
      splices.push({ start: tagStart + span.valueStart, end: tagStart + span.valueEnd, text: next });
    }
    return splices;
  }

  /** An envelope element's host: the sheet its base-cell-address names. */
  private baseCellHost(tag: string): Host {
    const span = findAttribute(tag, "table:base-cell-address");
    if (!span) {
      return {};
    }
    const first = splitOdsRangeAddressList(decodeRaw(span.raw).text).find((piece, k) => k % 2 === 0 && piece !== "");
    const ref = first === undefined ? null : parseOdsRangeAddress(first);
    const sheet = ref?.sheet === undefined ? undefined : this.sheetNamed(ref.sheet, undefined);
    return sheet ? { sheet } : {};
  }

  /** A raw address list adjusted: undefined when nothing moved, null when
   * every entry is gone. */
  private list(raw: string, host: Host, kind: Kind): string | null | undefined {
    const value = decodeRaw(raw);
    const pieces = splitOdsRangeAddressList(value.text);
    const out: (string | null)[] = [];
    let changed = false;
    let pos = 0;
    for (let k = 0; k < pieces.length; k += 1) {
      const piece = pieces[k]!;
      const start = pos;
      pos += piece.length;
      const next = k % 2 === 1 || piece === "" ? undefined : this.address(piece, (from, to) => raw.slice(value.at(start + from), value.at(start + to)), host, kind);
      if (next === undefined) {
        out.push(raw.slice(value.at(start), value.at(pos)));
      } else {
        changed = true;
        out.push(next);
      }
    }
    return changed ? joinList(pieces, out) : undefined;
  }

  /** One address's raw replacement: undefined when it stays as written;
   * null when its area is gone. */
  private address(text: string, rawOf: RawOf, host: Host, kind: Kind): string | null | undefined {
    const ref = parseOdsRangeAddress(text);
    if (!ref || ref.sheetEnd !== undefined) {
      return undefined;
    }
    const sheet = ref.sheet === undefined ? host.sheet : this.sheetNamed(ref.sheet, host.id);
    if (!sheet) {
      return undefined;
    }
    const moved = this.move(sheet.name, ref, kind === "anchor");
    if (moved === undefined) {
      return undefined;
    }
    if (moved === null) {
      return null;
    }
    return rewriteParts(text, rawOf, ref, moved);
  }

  /** Where an address's parts go: undefined when nowhere new, null when its
   * area is gone (an anchor's cell then goes where its band now starts). */
  private move(sheetName: string, ref: ReferenceNode, isAnchor: boolean): Moved | null | undefined {
    const range = gridOf(ref);
    const mapped = this.mapper.range(sheetName, range);
    if (mapped === range) {
      return undefined;
    }
    let moved: Moved;
    if (mapped !== null) {
      moved = placed(ref, mapped);
    } else if (!isAnchor) {
      return null;
    } else {
      const start = this.indexed(sheetName, ref.start);
      const end = ref.end && this.indexed(sheetName, ref.end);
      if (!start || end === null) {
        return undefined;
      }
      moved = end ? { start, end } : { start };
    }
    const same = samePosition(ref.start, moved.start) && (ref.end === undefined || samePosition(ref.end, moved.end));
    return same ? undefined : moved;
  }

  /** A part moved axis by axis through `mapper.index`; null when an axis
   * has nowhere to go. */
  private indexed(sheetName: string, part: RefPart): RefPart | null {
    const out: RefPart = { ...part };
    if (part.row !== undefined) {
      const row = this.mapper.index(sheetName, "row", part.row);
      if (row === null) {
        return null;
      }
      out.row = row;
    }
    if (part.col !== undefined) {
      const col = this.mapper.index(sheetName, "col", part.col);
      if (col === null) {
        return null;
      }
      out.col = col;
    }
    return out;
  }
}

/** Adjusts every retained ods reference through `mapper` (see the module
 * header); returns the ids of the sheets whose own retained text changed —
 * their sourceXml, a cell's preservedXml, or a chart object their table
 * shows. */
export function rewriteOdsRetainedReferences(model: SpreadsheetModel, mapper: RetainedReferenceMapper): Set<string> {
  return new OdsRetainedRewrite(model, mapper).run();
}

/** SPEC-128: every sheet's verbatim ods filter (`autoFilter.odsUnmappedFilter`)
 * with its `table:condition-source-range-address` moved through `mapper` —
 * the `autofilter` adjusters' share of the database range the model owns
 * (see the module header). Each autofilter whose filter changed is
 * replaced, without its sourceXml, and its sheet marked dirty. */
export function adjustOdsFilterCriteriaRanges(model: SpreadsheetModel, mapper: RetainedReferenceMapper): void {
  if (model.format !== "ods" || !model.sheets.some((sheet) => sheet.autoFilter?.odsUnmappedFilter !== undefined)) {
    return;
  }
  new OdsRetainedRewrite(model, mapper).filterSources();
}
