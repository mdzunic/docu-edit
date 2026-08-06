// Orchestrates the .docx -> DocumentModel pipeline: unzip the OPC package,
// walk the package/document relationships to find the real part names (per
// AC-3, never assume "word/document.xml" is where it is), map w:p/w:r onto
// the document model, resolve w:pStyle against styles.xml, and apply the
// section's page setup.
import type {
  Block,
  BorderLineStyle,
  BorderSide,
  CellBorders,
  CellVerticalAlign,
  DefaultStyles,
  NamedStyle,
  OpaqueBlock,
  OpaqueInline,
  OpaquePart,
  Paragraph,
  Run,
  RunFormatting,
  Section,
  SourceXml,
  Table,
  TableAlignment,
  TableCell,
  TableRow,
  TableWidthMode,
} from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { bytesToBase64 } from "../opaque";
import { ZipError, unzip } from "./zip";
import {
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  serializeElement,
  wordAttr,
  WORD_NAMESPACE,
  XmlParseError,
} from "./xml";
import { catalogToNamedStyles, parseStylesCatalog, resolveStyle, type StyleCatalog } from "./styles";
import {
  EMPTY_NUMBERING_CATALOG,
  NUMBERING_RELATIONSHIP,
  parseNumberingCatalog,
  type NumberingCatalog,
} from "./numbering";
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseOutlineLevel,
  parseParagraphProperties,
  parseRunProperties,
} from "./properties";
import { eighthPointsToPt, ooxmlColorToCss, twipsToCm } from "./units";

export class DocxOpenError extends Error {}
export class LegacyDocError extends Error {}

export const INVALID_DOCX_MESSAGE = "This file isn't a Word document we can open";
export const LEGACY_DOC_MESSAGE = "Old .doc files aren't supported — save as .docx first";

const OFFICE_DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

// Transient markup Word itself regenerates on every open — spell/grammar
// error ranges and rendered pagination markers. Preserving these as opaque
// nodes would freeze stale derived data into the file, so they are the one
// deliberate exception to SPEC-008's "carry everything through" rule.
const TRANSIENT_PARAGRAPH_CHILDREN = new Set(["proofErr"]);
const MODELLED_RUN_CHILDREN = new Set(["rPr", "t", "tab", "br", "cr", "lastRenderedPageBreak"]);

/** True when a run carries content the model cannot express (w:drawing,
 * w:object, w:pict, field chars, …) — such a run is preserved whole as an
 * opaque inline node rather than flattened to text. */
function runHasUnmodelledContent(rEl: Element): boolean {
  return Array.from(rEl.children).some(
    (child) => child.namespaceURI !== WORD_NAMESPACE || !MODELLED_RUN_CHILDREN.has(child.localName),
  );
}

// The Compound File Binary Format signature every legacy .doc (Word 97-2003)
// file starts with — used to reject them with a specific, actionable error
// instead of the generic "not a ZIP" message a naive open would produce.
const DOC_MAGIC_BYTES = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function isLegacyDoc(bytes: Uint8Array): boolean {
  if (bytes.length < DOC_MAGIC_BYTES.length) {
    return false;
  }
  return DOC_MAGIC_BYTES.every((byte, index) => bytes[index] === byte);
}

export interface ParsedDocxDocument {
  sections: Section[];
  defaultStyles: DefaultStyles;
  /** The paragraph styles of styles.xml as the model's style catalogue
   * (SPEC-019); empty when the package has no styles part. */
  styles: NamedStyle[];
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .docx. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from modelled parts (the body-level w:sectPr). */
  sourceXml: SourceXml;
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

function parseRun(rEl: Element, inheritedFormatting: RunFormatting): Run | undefined {
  const rPr = firstChildElement(rEl, WORD_NAMESPACE, "rPr");
  const formatting = mergeRunFormatting(inheritedFormatting, parseRunProperties(rPr));

  let text = "";
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "t") {
      text += child.textContent ?? "";
    } else if (child.localName === "tab") {
      text += "\t";
    } else if (child.localName === "br" || child.localName === "cr") {
      text += "\n";
    }
  }

  if (text.length === 0) {
    return undefined;
  }
  return { text, formatting };
}

function collectRuns(
  container: Element,
  inheritedFormatting: RunFormatting,
  runs: Run[],
  opaqueInlines: OpaqueInline[],
): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pPr") {
      continue; // Modelled separately by parseParagraph.
    }
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "r") {
      if (runHasUnmodelledContent(child)) {
        // A drawing/object/field run is preserved whole — splitting it into
        // "the text we understood" plus "the rest" could not round-trip.
        opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
        continue;
      }
      const run = parseRun(child, inheritedFormatting);
      if (run) {
        runs.push(run);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "hyperlink") {
      // Hyperlinks (SPEC-025) wrap w:r children; not modelling the link
      // itself yet, but its text must still reach the document per the
      // round-trip/no-content-loss principle.
      collectRuns(child, inheritedFormatting, runs, opaqueInlines);
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      TRANSIENT_PARAGRAPH_CHILDREN.has(child.localName)
    ) {
      // Regenerated by Word on open; deliberately not preserved.
    } else {
      // Anything else (w:sdt, w:fldSimple, bookmarks, other namespaces) is
      // an unmodelled inline element: preserved verbatim, anchored to its
      // position between the runs (SPEC-008).
      opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
    }
  }
}

function parseParagraph(
  pEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): Paragraph {
  const pPr = firstChildElement(pEl, WORD_NAMESPACE, "pPr");
  const pStyleEl = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "pStyle") : undefined;
  const styleId = pStyleEl ? (wordAttr(pStyleEl, "val") ?? undefined) : undefined;

  const resolved = resolveStyle(styleCatalog, styleId);
  const formatting = mergeParagraphFormatting(resolved.paragraphFormatting, parseParagraphProperties(pPr));
  // "ltr" exists only to let a direct w:bidi="0" override a style's RTL
  // during the merge above; the model itself holds "rtl" or nothing.
  if (formatting.direction === "ltr") {
    delete formatting.direction;
  }
  const headingLevel = parseOutlineLevel(pPr) ?? resolved.headingLevel;

  const runs: Run[] = [];
  const opaqueInlines: OpaqueInline[] = [];
  collectRuns(pEl, resolved.runFormatting, runs, opaqueInlines);

  const paragraph: Paragraph = { type: "paragraph", runs, formatting };
  // w:numPr resolves against the numbering catalog (SPEC-018 AC-20): the
  // referenced level definition decides type, marker, start and indent.
  const numPr = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "numPr") : undefined;
  if (numPr) {
    const numIdEl = firstChildElement(numPr, WORD_NAMESPACE, "numId");
    const ilvlEl = firstChildElement(numPr, WORD_NAMESPACE, "ilvl");
    const numId = numIdEl ? wordAttr(numIdEl, "val") : null;
    const ilvl = Number((ilvlEl ? wordAttr(ilvlEl, "val") : null) ?? "0");
    if (numId !== null && Number.isInteger(ilvl) && ilvl >= 0) {
      const list = numberingCatalog.listFormattingFor(numId, ilvl);
      if (list) {
        paragraph.list = list;
      }
    }
  }
  if (styleId) {
    paragraph.styleId = styleId;
  }
  if (headingLevel !== undefined) {
    paragraph.headingLevel = headingLevel;
  }
  if (opaqueInlines.length > 0) {
    paragraph.opaqueInlines = opaqueInlines;
  }
  return paragraph;
}

/** OOXML `w:val`s for `w:tcBorders`' children collapse to the model's small
 * enum (SPEC-023): the handful of dashed variants all read as "dashed",
 * "nil"/"none" as "none", and anything else (including "thick", which the
 * model has no separate weight class for) as a plain single line. */
function mapBorderLineStyle(val: string): BorderLineStyle {
  const v = val.toLowerCase();
  if (v === "nil" || v === "none") {
    return "none";
  }
  if (v === "double" || v === "triple") {
    return "double";
  }
  if (v === "dotted") {
    return "dotted";
  }
  if (v.startsWith("dash") || v.includes("dotdash") || v.includes("dotdotdash")) {
    return "dashed";
  }
  return "single";
}

function parseBorderSide(el: Element | undefined): BorderSide | undefined {
  if (!el) {
    return undefined;
  }
  const val = wordAttr(el, "val");
  if (!val) {
    return undefined;
  }
  const style = mapBorderLineStyle(val);
  const sz = Number(wordAttr(el, "sz") ?? "");
  const widthPt = style !== "none" && Number.isFinite(sz) && sz > 0 ? eighthPointsToPt(sz) : 0;
  const color = ooxmlColorToCss(wordAttr(el, "color")) ?? "auto";
  return { style, widthPt, color };
}

/** A cell's `w:tcBorders` (SPEC-023 AC-10..12, AC-24). OOXML's logical
 * `w:start`/`w:end` are read (falling back to the legacy `w:left`/`w:right`
 * some producers still emit) and stored as the model's physical left/right —
 * table RTL is out of scope, so the two coincide. */
function parseCellBorders(tcPr: Element): CellBorders | undefined {
  const bordersEl = firstChildElement(tcPr, WORD_NAMESPACE, "tcBorders");
  if (!bordersEl) {
    return undefined;
  }
  const top = parseBorderSide(firstChildElement(bordersEl, WORD_NAMESPACE, "top"));
  const bottom = parseBorderSide(firstChildElement(bordersEl, WORD_NAMESPACE, "bottom"));
  const left = parseBorderSide(
    firstChildElement(bordersEl, WORD_NAMESPACE, "start") ??
      firstChildElement(bordersEl, WORD_NAMESPACE, "left"),
  );
  const right = parseBorderSide(
    firstChildElement(bordersEl, WORD_NAMESPACE, "end") ?? firstChildElement(bordersEl, WORD_NAMESPACE, "right"),
  );
  if (!top && !bottom && !left && !right) {
    return undefined;
  }
  const borders: CellBorders = {};
  if (top) {
    borders.top = top;
  }
  if (bottom) {
    borders.bottom = bottom;
  }
  if (left) {
    borders.left = left;
  }
  if (right) {
    borders.right = right;
  }
  return borders;
}

/** A cell's `w:shd` fill as CSS (SPEC-023 AC-15, AC-26): `w:fill` is the
 * background colour; `w:color` is the pattern colour on top of it, which the
 * model has no slot for and drops (the same simplification run shading
 * makes, SPEC-013). */
function parseCellShading(tcPr: Element): string | undefined {
  const shd = firstChildElement(tcPr, WORD_NAMESPACE, "shd");
  return shd ? ooxmlColorToCss(wordAttr(shd, "fill")) : undefined;
}

/** A cell's `w:tcMar` (SPEC-023 AC-16): the model holds one uniform padding,
 * so the top margin stands in for all four — same-format round trips are
 * exact because the writer re-emits all four equal to it. */
function parseCellPadding(tcPr: Element): number | undefined {
  const mar = firstChildElement(tcPr, WORD_NAMESPACE, "tcMar");
  const top = mar ? firstChildElement(mar, WORD_NAMESPACE, "top") : undefined;
  const w = top ? Number(wordAttr(top, "w") ?? "") : NaN;
  return Number.isFinite(w) && w >= 0 ? twipsToCm(w) : undefined;
}

function parseCellVerticalAlign(tcPr: Element): CellVerticalAlign | undefined {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "vAlign");
  const val = el ? wordAttr(el, "val") : null;
  return val === "top" || val === "center" || val === "bottom" ? val : undefined;
}

function parseGridSpan(tcPr: Element): number | undefined {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "gridSpan");
  const val = el ? Number(wordAttr(el, "val")) : NaN;
  return Number.isInteger(val) && val > 1 ? val : undefined;
}

/** A cell's `w:vMerge`: absent (not vertically merged), "restart" (the
 * origin of a new vertical merge — carries its own content) or "continue"
 * (covered by the origin above; an explicit `w:val` is required for
 * "restart", its absence defaults to "continue" per the schema). */
type VMergeKind = "restart" | "continue" | null;

function parseVMerge(tcPr: Element): VMergeKind {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "vMerge");
  if (!el) {
    return null;
  }
  return wordAttr(el, "val") === "restart" ? "restart" : "continue";
}

/** One `w:tc`'s parse result: the cell itself plus the two attributes
 * (SPEC-023) that decide how many grid slots it and its followers occupy —
 * kept out of the TableCell union so pass 2 below can resolve vMerge chains
 * before the row/table shape is finalised. */
interface ParsedCellSlot {
  cell: TableCell;
  colSpan: number;
  vMerge: VMergeKind;
}

function parseTableCell(
  tcEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): ParsedCellSlot | null {
  const blocks: Paragraph[] = [];
  let tcPrEl: Element | undefined;
  let tcPrXml: string | undefined;
  for (const child of Array.from(tcEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "tcPr") {
      tcPrEl = child;
      tcPrXml = serializeElement(child);
    } else if (child.localName === "p") {
      blocks.push(parseParagraph(child, styleCatalog, numberingCatalog));
    } else {
      // Nested tables, block-level sdt, … — out of the model's table shape.
      return null;
    }
  }
  if (blocks.length === 0) {
    // OOXML requires a w:p per cell; tolerate its absence with an empty one
    // so every cell can hold a caret.
    blocks.push({ type: "paragraph", runs: [], formatting: {} });
  }
  const cell: TableCell = { blocks };
  const colSpan = tcPrEl ? parseGridSpan(tcPrEl) : undefined;
  if (colSpan !== undefined) {
    cell.colSpan = colSpan;
  }
  const borders = tcPrEl ? parseCellBorders(tcPrEl) : undefined;
  if (borders) {
    cell.borders = borders;
  }
  const shadingFill = tcPrEl ? parseCellShading(tcPrEl) : undefined;
  if (shadingFill !== undefined) {
    cell.shadingFill = shadingFill;
  }
  const paddingCm = tcPrEl ? parseCellPadding(tcPrEl) : undefined;
  if (paddingCm !== undefined) {
    cell.paddingCm = paddingCm;
  }
  const verticalAlign = tcPrEl ? parseCellVerticalAlign(tcPrEl) : undefined;
  if (verticalAlign !== undefined) {
    cell.verticalAlign = verticalAlign;
  }
  if (tcPrXml !== undefined) {
    cell.docxTcPr = tcPrXml;
  }
  return { cell, colSpan: colSpan ?? 1, vMerge: tcPrEl ? parseVMerge(tcPrEl) : null };
}

/** A `w:tc` occupies one grid column normally, or `colSpan` of them when it
 * carries `w:gridSpan` — the extra columns have no `w:tc` of their own, so
 * they become covered placeholder slots here to keep the model's per-row
 * cell count equal to the table's grid column count (SPEC-023). */
type RowSlot = { kind: "cell"; parsed: ParsedCellSlot } | { kind: "covered" };

function parseTableRowSlots(
  trEl: Element,
  gridColumns: number,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): { slots: RowSlot[]; trPrEl?: Element; trPr?: string } | null {
  const slots: RowSlot[] = [];
  let trPrEl: Element | undefined;
  let trPr: string | undefined;
  for (const child of Array.from(trEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "trPr") {
      trPrEl = child;
      trPr = serializeElement(child);
    } else if (child.localName === "tc") {
      const parsed = parseTableCell(child, styleCatalog, numberingCatalog);
      if (!parsed) {
        return null;
      }
      slots.push({ kind: "cell", parsed });
      for (let i = 1; i < parsed.colSpan; i++) {
        slots.push({ kind: "covered" });
      }
    } else {
      // tblPrEx, bookmarks between cells, … — keep the table verbatim.
      return null;
    }
  }
  if (slots.length === 0 || slots.length !== gridColumns) {
    return null;
  }
  return { slots, trPrEl, trPr };
}

/** A row's minimum height from `w:trHeight` (SPEC-023 AC-4, AC-5): both
 * `hRule="atLeast"` (an explicit minimum) and its absence (Word's own
 * default meaning) map to the model's `minHeightCm`, which is always a
 * floor. `hRule="exact"` (a fixed height, not modelled) is read the same
 * way — an approximation, but a closer one than dropping the height. */
function parseRowMinHeight(trPr: Element | undefined): number | undefined {
  const heightEl = trPr ? firstChildElement(trPr, WORD_NAMESPACE, "trHeight") : undefined;
  const val = heightEl ? Number(wordAttr(heightEl, "val") ?? "") : NaN;
  return Number.isFinite(val) && val > 0 ? twipsToCm(val) : undefined;
}

/** A table's `w:tblW` (SPEC-023 AC-23, AC-27): `pct` is fiftieths of a
 * percent (5000 = 100%), `auto` sizes to content, anything else (including
 * no `w:tblW` at all) is a fixed width — the `columnWidthsCm` sum. */
function parseTableWidth(tblPr: Element | undefined): { widthMode?: TableWidthMode; widthPercent?: number } {
  const tblW = tblPr ? firstChildElement(tblPr, WORD_NAMESPACE, "tblW") : undefined;
  if (!tblW) {
    return {};
  }
  const type = wordAttr(tblW, "type");
  if (type === "pct") {
    const w = Number(wordAttr(tblW, "w") ?? "");
    if (Number.isFinite(w) && w > 0) {
      return { widthMode: "percent", widthPercent: Math.min(100, Math.max(1, Math.round(w / 50))) };
    }
  }
  if (type === "auto") {
    return { widthMode: "auto" };
  }
  return { widthMode: "fixed" };
}

/** A table's `w:jc` (SPEC-023 AC-22): "start" is the logical form of
 * "left" — table RTL is out of scope, so both map to the model's "left". */
function parseTableAlignment(tblPr: Element | undefined): TableAlignment | undefined {
  const jc = tblPr ? firstChildElement(tblPr, WORD_NAMESPACE, "jc") : undefined;
  const val = jc ? wordAttr(jc, "val") : null;
  if (val === "center" || val === "right") {
    return val;
  }
  if (val === "left" || val === "start") {
    return "left";
  }
  return undefined;
}

/** Parses a `w:tbl` into a modelled Table (SPEC-022; merged cells and
 * formatting SPEC-023), or null when the table uses features outside the
 * model's regular grid — nested tables, a grid that doesn't match the rows —
 * in which case the caller preserves the element verbatim as an opaque
 * block instead (SPEC-008: fidelity over feature count).
 *
 * Merged cells keep the grid regular for indexing (every row still has
 * `columnWidthsCm.length` cells): a `w:gridSpan` cell's extra columns and a
 * `w:vMerge="continue"` cell both become `covered: true` placeholders, and
 * the real cell they are covered by carries `colSpan`/`rowSpan`. Resolving
 * `vMerge` needs a second pass over the whole table (a "continue" cell's
 * `rowSpan` accumulates on whichever earlier row most recently started a
 * chain at that column), so rows are parsed into slots first and only
 * assembled into the model's `TableRow`s afterwards. */
function parseTable(
  tblEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): Table | null {
  const columnWidthsCm: number[] = [];
  let tblPrEl: Element | undefined;
  let tblPrXml: string | undefined;
  let sawGrid = false;
  const rowSlots: Array<{ slots: RowSlot[]; trPrEl?: Element; trPr?: string }> = [];
  for (const child of Array.from(tblEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "tblPr") {
      tblPrEl = child;
      tblPrXml = serializeElement(child);
    } else if (child.localName === "tblGrid") {
      sawGrid = true;
      for (const col of Array.from(child.children)) {
        if (col.namespaceURI !== WORD_NAMESPACE || col.localName !== "gridCol") {
          return null;
        }
        const w = Number(wordAttr(col, "w") ?? "");
        if (!Number.isFinite(w) || w <= 0) {
          return null;
        }
        columnWidthsCm.push(twipsToCm(w));
      }
    } else if (child.localName === "tr") {
      if (!sawGrid) {
        return null;
      }
      const parsedRow = parseTableRowSlots(child, columnWidthsCm.length, styleCatalog, numberingCatalog);
      if (!parsedRow) {
        return null;
      }
      rowSlots.push(parsedRow);
    } else {
      return null;
    }
  }
  if (!sawGrid || rowSlots.length === 0 || columnWidthsCm.length === 0) {
    return null;
  }

  const rows: TableRow[] = rowSlots.map(({ slots, trPrEl, trPr }) => {
    const cells: TableCell[] = slots.map((slot) =>
      slot.kind === "cell" ? slot.parsed.cell : { blocks: [], covered: true },
    );
    const row: TableRow = { cells };
    if (trPr !== undefined) {
      row.docxTrPr = trPr;
    }
    const minHeightCm = parseRowMinHeight(trPrEl);
    if (minHeightCm !== undefined) {
      row.minHeightCm = minHeightCm;
    }
    return row;
  });

  // Pass 2: a "continue" cell's rowSpan accumulates onto whichever row most
  // recently started (or continued) a merge at that column.
  const openOrigin: Array<number | null> = Array.from({ length: columnWidthsCm.length }, () => null);
  for (let r = 0; r < rowSlots.length; r++) {
    const slots = rowSlots[r].slots;
    for (let col = 0; col < slots.length; col++) {
      const slot = slots[col];
      if (slot.kind !== "cell") {
        continue;
      }
      if (slot.parsed.vMerge === "continue" && openOrigin[col] !== null) {
        const originRow = openOrigin[col] as number;
        const originCell = rows[originRow].cells[col];
        originCell.rowSpan = (originCell.rowSpan ?? 1) + 1;
        rows[r].cells[col] = { blocks: [], covered: true };
      } else {
        openOrigin[col] = r;
      }
    }
  }

  const table: Table = { type: "table", rows, columnWidthsCm };
  if (tblPrXml !== undefined) {
    table.docxTblPr = tblPrXml;
  }
  const width = parseTableWidth(tblPrEl);
  if (width.widthMode !== undefined) {
    table.widthMode = width.widthMode;
  }
  if (width.widthPercent !== undefined) {
    table.widthPercent = width.widthPercent;
  }
  const alignment = parseTableAlignment(tblPrEl);
  if (alignment !== undefined) {
    table.alignment = alignment;
  }
  return table;
}

function applySectionProperties(sectPr: Element | undefined, base: DefaultStyles): DefaultStyles {
  if (!sectPr) {
    return base;
  }

  let pageSize = base.pageSize;
  const pgSz = firstChildElement(sectPr, WORD_NAMESPACE, "pgSz");
  if (pgSz) {
    const w = wordAttr(pgSz, "w");
    const h = wordAttr(pgSz, "h");
    if (w && h) {
      let widthTwips = Number(w);
      let heightTwips = Number(h);
      const orient = wordAttr(pgSz, "orient")?.toLowerCase();
      if (orient === "landscape" && widthTwips < heightTwips) {
        [widthTwips, heightTwips] = [heightTwips, widthTwips];
      }
      pageSize = { widthCm: twipsToCm(widthTwips), heightCm: twipsToCm(heightTwips) };
    }
  }

  let margins = base.margins;
  const pgMar = firstChildElement(sectPr, WORD_NAMESPACE, "pgMar");
  if (pgMar) {
    const top = wordAttr(pgMar, "top");
    const right = wordAttr(pgMar, "right");
    const bottom = wordAttr(pgMar, "bottom");
    const left = wordAttr(pgMar, "left");
    if (top && right && bottom && left) {
      margins = {
        topCm: twipsToCm(Number(top)),
        rightCm: twipsToCm(Number(right)),
        bottomCm: twipsToCm(Number(bottom)),
        leftCm: twipsToCm(Number(left)),
      };
    }
  }

  return { ...base, pageSize, margins };
}

/** Parses a `.docx` file's bytes into the sections/defaultStyles this
 * project's document model needs. Throws `LegacyDocError` for a binary
 * Word 97-2003 `.doc` file, or `DocxOpenError` for anything else that
 * isn't a valid, openable `.docx` package. */
export async function parseDocx(data: ArrayBuffer | Uint8Array): Promise<ParsedDocxDocument> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  if (isLegacyDoc(bytes)) {
    throw new LegacyDocError(LEGACY_DOC_MESSAGE);
  }

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      throw new DocxOpenError(INVALID_DOCX_MESSAGE);
    }
    throw error;
  }

  const readPart = (path: string): string | undefined => {
    const raw = parts.get(path);
    return raw ? decodeUtf8(raw) : undefined;
  };

  let mainPartPath = "word/document.xml";
  let corePartPath: string | undefined;
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const officeDocRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP,
    );
    if (officeDocRel) {
      mainPartPath = resolvePartPath("", officeDocRel.target);
    }
    // Core properties are rebuilt on save (they carry the edited title), so
    // the original core part is the one part besides the main document that
    // is not retained verbatim.
    const coreRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP,
    );
    if (coreRel) {
      corePartPath = resolvePartPath("", coreRel.target);
    }
  }

  const documentXmlText = readPart(mainPartPath);
  if (!documentXmlText) {
    throw new DocxOpenError(INVALID_DOCX_MESSAGE);
  }

  let documentXml: Document;
  try {
    documentXml = parseXml(documentXmlText);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new DocxOpenError(INVALID_DOCX_MESSAGE);
    }
    throw error;
  }

  const body = firstChildElement(documentXml.documentElement, WORD_NAMESPACE, "body");
  if (!body) {
    throw new DocxOpenError(INVALID_DOCX_MESSAGE);
  }

  let stylesDoc: Document | undefined;
  let numberingCatalog: NumberingCatalog = EMPTY_NUMBERING_CATALOG;
  const mainRelsText = readPart(relsPathFor(mainPartPath));
  if (mainRelsText) {
    const mainRels = parseRelationships(mainRelsText);
    const stylesRel = Array.from(mainRels.values()).find((rel) => rel.type === STYLES_RELATIONSHIP);
    if (stylesRel) {
      const stylesText = readPart(resolvePartPath(mainPartPath, stylesRel.target));
      if (stylesText) {
        stylesDoc = parseXml(stylesText);
      }
    }
    const numberingRel = Array.from(mainRels.values()).find(
      (rel) => rel.type === NUMBERING_RELATIONSHIP,
    );
    if (numberingRel) {
      const numberingText = readPart(resolvePartPath(mainPartPath, numberingRel.target));
      if (numberingText) {
        numberingCatalog = parseNumberingCatalog(parseXml(numberingText));
      }
    }
  }
  const styleCatalog = parseStylesCatalog(stylesDoc);

  const blocks: Block[] = [];
  const opaqueBlocks: OpaqueBlock[] = [];
  let sectPr: Element | undefined;
  for (const child of Array.from(body.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      blocks.push(parseParagraph(child, styleCatalog, numberingCatalog));
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "tbl") {
      // Regular-grid tables are modelled (SPEC-022); a table the model
      // cannot express (merged cells, nested tables, …) falls through to
      // the opaque path below.
      const table = parseTable(child, styleCatalog, numberingCatalog);
      if (table) {
        blocks.push(table);
      } else {
        opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "sectPr") {
      sectPr = child;
    } else {
      // Block-level content controls, bookmarks, anything from another
      // namespace: unmodelled body children are preserved verbatim,
      // anchored to their position among the paragraphs (SPEC-008).
      opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
    }
  }

  const fallbackStyles = createDefaultStyles();
  const defaultStyles = applySectionProperties(sectPr, {
    ...fallbackStyles,
    fontFamily: styleCatalog.docDefaultRunFormatting.fontFamily ?? fallbackStyles.fontFamily,
    fontSizePt: styleCatalog.docDefaultRunFormatting.fontSizePt ?? fallbackStyles.fontSizePt,
  });

  const section: Section = { blocks };
  if (opaqueBlocks.length > 0) {
    section.opaqueBlocks = opaqueBlocks;
  }

  // Everything the pipeline above did not consume-and-rebuild is retained
  // verbatim for the lifetime of the document (SPEC-008 AC): styles.xml and
  // the relationship/content-type parts included, since a same-format save
  // reuses them as-is.
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (name !== mainPartPath && name !== corePartPath) {
      opaque.push({ name, data: bytesToBase64(data) });
    }
  }

  const sourceXml: SourceXml = {};
  if (sectPr) {
    sourceXml.docxSectPr = serializeElement(sectPr);
  }

  return {
    sections: [section],
    defaultStyles,
    styles: catalogToNamedStyles(styleCatalog),
    opaque,
    sourceXml,
  };
}
