// Orchestrates the .odt -> DocumentModel pipeline: unzip the ODF package,
// detect the format from the package's `mimetype` entry (never the file
// extension), parse content.xml and styles.xml, map text:p/text:h/text:span
// onto the document model with ODF whitespace handling (text:s, text:tab,
// text:line-break), resolve style chains, and apply the master page's layout.
import type {
  Block,
  DefaultStyles,
  ListFormatting,
  OpaqueBlock,
  OpaqueInline,
  OpaquePart,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunFormatting,
  Section,
  SourceXml,
  Table,
  TableCell,
  TableRow,
} from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { bytesToBase64 } from "../opaque";
import { ZipError, unzip } from "../docx/zip";
import { cmToTwips } from "../docx/units";
import { firstChildElement, parseXml, serializeElement, XmlParseError } from "../docx/xml";
import { clampListLevel, defaultMarkerStyle } from "../lists";
import { odfLengthToCm } from "./units";
import {
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  styleAttr,
  tableAttr,
  textAttr,
} from "./xml";
import {
  parseOdfStyleCatalog,
  resolveListLevel,
  resolveParagraphStyle,
  resolveTextStyle,
  type OdfListLevelDef,
  type OdfStyleCatalog,
} from "./styles";

export class OdtOpenError extends Error {}
export class FlatOdfError extends Error {}

export const INVALID_ODT_MESSAGE = "This file isn't an OpenDocument text file we can open";
export const FLAT_ODF_MESSAGE = "Flat ODF files aren't supported yet";

/** The package mimetype identifying an OpenDocument text document. */
export const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";

export interface ParsedOdtDocument {
  sections: Section[];
  defaultStyles: DefaultStyles;
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .odt. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from content.xml that the model cannot express:
   * font-face declarations and non-paragraph/text automatic styles that
   * preserved opaque nodes may reference. */
  sourceXml: SourceXml;
}

// Transient markup LibreOffice regenerates on every open (rendered
// pagination) — the one deliberate exception to "carry everything through";
// preserving it would freeze stale derived data into the file.
const TRANSIENT_INLINE_ELEMENTS = new Set(["soft-page-break"]);

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** A flat ODF file (.fodt) is the whole document as one XML file: its root
 * element is `office:document`. Recognised only to reject it with a specific
 * message rather than a generic "can't open". */
function isFlatOdf(bytes: Uint8Array): boolean {
  const text = decodeUtf8(bytes);
  if (!text.trimStart().startsWith("<")) {
    return false;
  }
  let doc: Document;
  try {
    doc = parseXml(text);
  } catch {
    return false;
  }
  const root = doc.documentElement;
  return root.namespaceURI === OFFICE_NAMESPACE && root.localName === "document";
}

/** Classifies a file's bytes by content, per AC "format is detected from the
 * mimetype entry rather than the extension": "package" for a ZIP whose
 * `mimetype` entry declares OpenDocument text, "flat" for single-XML ODF
 * (.fodt), "none" for anything else (the .docx path handles those). */
export async function detectOdf(data: ArrayBuffer | Uint8Array): Promise<"package" | "flat" | "none"> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      return isFlatOdf(bytes) ? "flat" : "none";
    }
    throw error;
  }
  const mimetype = parts.get("mimetype");
  return mimetype && decodeUtf8(mimetype).trim() === ODT_MIMETYPE ? "package" : "none";
}

// --- inline content --------------------------------------------------------

/** A contiguous stretch of same-formatted text while walking a paragraph.
 * Pieces reference the shared resolved formatting object, so adjacent
 * content under the same formatting merges by identity; each piece becomes
 * an independent Run (with its own formatting copy) once the paragraph is
 * done. */
interface TextPiece {
  text: string;
  formatting: RunFormatting;
}

interface InlineAccumulator {
  pieces: TextPiece[];
  /** Last character emitted into the paragraph ("" at paragraph start) —
   * what literal XML whitespace collapses against. Opaque nodes are
   * transparent to it, matching ODF whitespace handling (elements do not
   * interrupt collapsing). */
  lastChar: string;
  /** Unmodelled inline elements, anchored by the piece index they precede. */
  opaqueInlines: OpaqueInline[];
  /** Set after an opaque node so the following text starts a new piece —
   * merging across it would break the node's anchor. */
  mergeBarrier: boolean;
}

function appendText(acc: InlineAccumulator, text: string, formatting: RunFormatting): void {
  if (text.length === 0) {
    return;
  }
  const last = acc.pieces[acc.pieces.length - 1];
  if (!acc.mergeBarrier && last && last.formatting === formatting) {
    last.text += text;
  } else {
    acc.pieces.push({ text, formatting });
  }
  acc.mergeBarrier = false;
  acc.lastChar = text[text.length - 1];
}

const WHITESPACE_RUN = /[ \t\r\n]+/g;

/** Appends a literal text node with ODF whitespace handling (6.1.2): runs of
 * XML whitespace collapse to a single space, and a space at the start of a
 * paragraph or after other whitespace (including explicit text:s/text:tab
 * content) is dropped. Explicit whitespace elements bypass this entirely —
 * that split is what makes the result match what LibreOffice shows. */
function appendLiteralText(acc: InlineAccumulator, raw: string, formatting: RunFormatting): void {
  let text = raw.replace(WHITESPACE_RUN, " ");
  if (text.startsWith(" ") && (acc.lastChar === "" || " \t\n".includes(acc.lastChar))) {
    text = text.slice(1);
  }
  appendText(acc, text, formatting);
}

function collectInline(
  container: Element,
  formatting: RunFormatting,
  catalog: OdfStyleCatalog,
  acc: InlineAccumulator,
): void {
  for (const node of Array.from(container.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.CDATA_SECTION_NODE) {
      appendLiteralText(acc, (node as CharacterData).data, formatting);
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const el = node as Element;
    if (el.namespaceURI !== TEXT_NAMESPACE) {
      // Frames, annotations, change marks, … — unmodelled inline elements
      // from other namespaces are preserved verbatim, anchored to their
      // position between the runs (SPEC-008).
      acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
      acc.mergeBarrier = true;
      continue;
    }
    switch (el.localName) {
      case "s": {
        // <text:s text:c="3"/> is three explicit (non-collapsing) spaces.
        const count = Number(textAttr(el, "c") ?? "1");
        appendText(acc, " ".repeat(Number.isInteger(count) && count > 0 ? count : 1), formatting);
        break;
      }
      case "tab":
        appendText(acc, "\t", formatting);
        break;
      case "line-break":
        appendText(acc, "\n", formatting);
        break;
      case "span":
        collectInline(el, resolveTextStyle(catalog, textAttr(el, "style-name") ?? undefined, formatting), catalog, acc);
        break;
      case "a":
        // Hyperlinks wrap inline content; the link itself isn't modelled yet
        // (SPEC-025) but its text must reach the document, matching the
        // .docx reader's w:hyperlink handling.
        collectInline(el, formatting, catalog, acc);
        break;
      default:
        if (TRANSIENT_INLINE_ELEMENTS.has(el.localName)) {
          // Regenerated by the producer on open; deliberately not preserved.
          break;
        }
        // Bookmarks, fields, notes, … — unmodelled text-namespace elements
        // are preserved verbatim at their anchored position (SPEC-008).
        acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
        acc.mergeBarrier = true;
        break;
    }
  }
}

/** ODF's fo:margin-left/right are physical, so the property parser fills
 * the logical slots as if the paragraph were LTR. Once the style chain is
 * resolved, an effective RTL writing mode means the physical left edge is
 * the logical END — so the two slots swap; an explicit "ltr" (recorded only
 * to win the chain merge) is dropped, leaving "rtl" the model's one stored
 * direction (SPEC-014). */
function normalizeResolvedDirection(formatting: ParagraphFormatting): void {
  if (formatting.direction === "rtl") {
    const start = formatting.indentStartCm;
    const end = formatting.indentEndCm;
    if (end !== undefined) {
      formatting.indentStartCm = end;
    } else {
      delete formatting.indentStartCm;
    }
    if (start !== undefined) {
      formatting.indentEndCm = start;
    } else {
      delete formatting.indentEndCm;
    }
  } else if (formatting.direction === "ltr") {
    delete formatting.direction;
  }
}

function parseTextBlock(el: Element, catalog: OdfStyleCatalog, isHeading: boolean): Paragraph {
  const styleName = textAttr(el, "style-name") ?? undefined;
  const resolved = resolveParagraphStyle(catalog, styleName);

  const acc: InlineAccumulator = { pieces: [], lastChar: "", opaqueInlines: [], mergeBarrier: false };
  collectInline(el, resolved.runFormatting, catalog, acc);
  // Each run gets its own formatting object — runs must stay independently
  // mutable for the editing specs even where they shared a resolved style.
  const runs: Run[] = acc.pieces
    .filter((piece) => piece.text.length > 0)
    .map((piece) => ({ text: piece.text, formatting: { ...piece.formatting } }));

  // Fresh object per paragraph — the resolver may hand back the catalog's
  // shared default when the style chain is empty.
  const formatting: ParagraphFormatting = { ...resolved.paragraphFormatting };
  normalizeResolvedDirection(formatting);
  const paragraph: Paragraph = { type: "paragraph", runs, formatting };
  if (styleName) {
    paragraph.styleId = styleName;
  }
  if (acc.opaqueInlines.length > 0) {
    paragraph.opaqueInlines = acc.opaqueInlines;
  }
  if (isHeading) {
    // text:outline-level is 1-based; per ODF a text:h without one is level 1
    // unless its style chain says otherwise (style:default-outline-level).
    const attrValue = Number(textAttr(el, "outline-level") ?? "");
    const level = Number.isInteger(attrValue) && attrValue > 0 ? attrValue : resolved.headingLevel;
    paragraph.headingLevel = level ?? 1;
  }
  return paragraph;
}

// --- lists (SPEC-018) ------------------------------------------------------

/** An item's model list formatting from its resolved level definition, with
 * default-elision so an editor-default list re-imports without noise: the
 * marker is stored only off the AC-5 cycle, startAt only when not 1, and
 * the indent only when it differs from level × 1.27 cm (compared in twips,
 * the shared cross-format grid, so float noise cannot leak in). */
function buildListFormatting(def: OdfListLevelDef | undefined, level: number): ListFormatting {
  const type = def?.type ?? "bullet";
  const list: ListFormatting = { type, level };
  if (!def) {
    return list;
  }
  if (def.marker !== defaultMarkerStyle(type, level)) {
    list.marker = def.marker;
  }
  if (type === "number" && def.startAt !== 1) {
    list.startAt = def.startAt;
  }
  if (def.indentCm !== undefined && cmToTwips(def.indentCm) !== level * 720) {
    list.indentCm = def.indentCm;
  }
  return list;
}

/** Parses a `text:list` (and its nested sub-lists) into consecutive list
 * paragraphs (AC-21): nesting depth is the model's level, and the list's
 * `text:list-style` levels supply type, marker, start and indent. Nested
 * lists without their own style name inherit the ancestor's, per ODF. */
function parseListElement(
  listEl: Element,
  catalog: OdfStyleCatalog,
  depth: number,
  inheritedStyleName: string | undefined,
  blocks: Block[],
  opaqueBlocks: OpaqueBlock[],
): void {
  const styleName = textAttr(listEl, "style-name") ?? inheritedStyleName;
  const level = clampListLevel(depth);
  for (const item of Array.from(listEl.children)) {
    const isItem =
      item.namespaceURI === TEXT_NAMESPACE &&
      (item.localName === "list-item" || item.localName === "list-header");
    if (!isItem) {
      opaqueBlocks.push({ xml: serializeElement(item), blockIndex: blocks.length });
      continue;
    }
    for (const child of Array.from(item.children)) {
      if (
        child.namespaceURI === TEXT_NAMESPACE &&
        (child.localName === "p" || child.localName === "h")
      ) {
        const paragraph = parseTextBlock(child, catalog, child.localName === "h");
        paragraph.list = buildListFormatting(
          resolveListLevel(styleName ? catalog.listStyles.get(styleName) : undefined, level),
          level,
        );
        blocks.push(paragraph);
      } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
        parseListElement(child, catalog, depth + 1, styleName, blocks, opaqueBlocks);
      } else {
        // Unmodelled list-item content (tables, frames, …) is preserved at
        // the body level, anchored where the item sat (SPEC-008).
        opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
      }
    }
  }
}

// --- tables (SPEC-022) -----------------------------------------------------

/** Column widths declared by `table-column`-family styles, from both parts'
 * automatic styles: style name -> `style:column-width` in cm. */
function collectTableColumnWidths(contentDoc: Document, stylesDoc: Document | undefined): Map<string, number> {
  const widths = new Map<string, number>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-column") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-column-properties");
      const width = props ? odfLengthToCm(styleAttr(props, "column-width")) : undefined;
      if (name && width !== undefined) {
        widths.set(name, width);
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return widths;
}

// Repeat expansion cap: `table:number-columns-repeated` is also how ODF
// producers pad to the sheet-like maximum; anything past this is not a text
// document's table and stays opaque.
const MAX_TABLE_REPEAT = 512;

function repeatCount(el: Element, attr: string): number | null {
  const raw = tableAttr(el, attr);
  if (raw === null) {
    return 1;
  }
  const count = Number(raw);
  if (!Number.isInteger(count) || count < 1 || count > MAX_TABLE_REPEAT) {
    return null;
  }
  return count;
}

function parseTableCellElement(cellEl: Element, catalog: OdfStyleCatalog): TableCell | null {
  const spanned = tableAttr(cellEl, "number-columns-spanned") ?? tableAttr(cellEl, "number-rows-spanned");
  if (spanned !== null && Number(spanned) !== 1) {
    // Merged cells (SPEC-023) break the regular grid; the whole table stays
    // opaque.
    return null;
  }
  const blocks: Block[] = [];
  const strayOpaque: OpaqueBlock[] = [];
  for (const child of Array.from(cellEl.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "p") {
      blocks.push(parseTextBlock(child, catalog, false));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "h") {
      blocks.push(parseTextBlock(child, catalog, true));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
      parseListElement(child, catalog, 1, undefined, blocks, strayOpaque);
    } else {
      // Nested tables, frames, … — outside the model's cell shape.
      return null;
    }
  }
  if (strayOpaque.length > 0) {
    // A cell whose list carried unmodelled content has no body level to
    // anchor it to; preserve the whole table verbatim instead.
    return null;
  }
  const paragraphs: Paragraph[] = [];
  for (const block of blocks) {
    if (block.type !== "paragraph") {
      return null;
    }
    paragraphs.push(block);
  }
  if (paragraphs.length === 0) {
    paragraphs.push({ type: "paragraph", runs: [], formatting: {} });
  }
  const cell: TableCell = { blocks: paragraphs };
  const styleName = tableAttr(cellEl, "style-name");
  if (styleName !== null) {
    cell.odtStyleName = styleName;
  }
  return cell;
}

/** Parses a `table:table` into a modelled Table (SPEC-022), or null when it
 * uses features outside the model's regular grid (merged/covered cells,
 * header rows, nested tables, sub-tables, …) — the caller then preserves
 * the element verbatim as an opaque block (SPEC-008). */
function parseTableElement(
  tableEl: Element,
  catalog: OdfStyleCatalog,
  columnWidths: Map<string, number>,
): Table | null {
  const columnStyleNames: Array<string | undefined> = [];
  const rows: TableRow[] = [];
  for (const child of Array.from(tableEl.children)) {
    if (child.namespaceURI !== TABLE_NAMESPACE) {
      return null;
    }
    if (child.localName === "table-column") {
      const repeat = repeatCount(child, "number-columns-repeated");
      if (repeat === null) {
        return null;
      }
      const styleName = tableAttr(child, "style-name") ?? undefined;
      for (let i = 0; i < repeat; i++) {
        columnStyleNames.push(styleName);
      }
    } else if (child.localName === "table-row") {
      const repeat = repeatCount(child, "number-rows-repeated");
      if (repeat === null) {
        return null;
      }
      const cells: TableCell[] = [];
      for (const cellEl of Array.from(child.children)) {
        if (cellEl.namespaceURI !== TABLE_NAMESPACE || cellEl.localName !== "table-cell") {
          return null;
        }
        const cellRepeat = repeatCount(cellEl, "number-columns-repeated");
        if (cellRepeat === null) {
          return null;
        }
        const cell = parseTableCellElement(cellEl, catalog);
        if (!cell) {
          return null;
        }
        cells.push(cell);
        for (let i = 1; i < cellRepeat; i++) {
          cells.push(structuredClone(cell));
        }
      }
      if (cells.length === 0) {
        return null;
      }
      const rowStyleName = tableAttr(child, "style-name");
      for (let i = 0; i < repeat; i++) {
        const row: TableRow = { cells: i === 0 ? cells : structuredClone(cells) };
        if (rowStyleName !== null) {
          row.odtStyleName = rowStyleName;
        }
        rows.push(row);
      }
    } else {
      // table-header-rows, covered cells, sub-tables, title/desc, … — the
      // model cannot express these; keep the table verbatim.
      return null;
    }
  }
  if (rows.length === 0 || columnStyleNames.length === 0) {
    return null;
  }
  if (rows.some((row) => row.cells.length !== columnStyleNames.length)) {
    return null;
  }
  // Column widths from the referenced column styles; when any is missing,
  // fall back to an equal split of the text area (or of the declared
  // widths' own total when only some are missing there is no better
  // anchor).
  const fallback = createDefaultStyles();
  const pageSize = catalog.pageSize ?? fallback.pageSize;
  const margins = catalog.margins ?? fallback.margins;
  const textAreaCm = Math.max(1, pageSize.widthCm - margins.leftCm - margins.rightCm);
  const declared = columnStyleNames.map((name) => (name ? columnWidths.get(name) : undefined));
  const columnWidthsCm = declared.every((width) => width !== undefined && width > 0)
    ? (declared as number[])
    : columnStyleNames.map(() => textAreaCm / columnStyleNames.length);

  const table: Table = { type: "table", rows, columnWidthsCm };
  const name = tableAttr(tableEl, "name");
  if (name !== null) {
    table.odtName = name;
  }
  const styleName = tableAttr(tableEl, "style-name");
  if (styleName !== null) {
    table.odtStyleName = styleName;
  }
  if (columnStyleNames.some((columnName) => columnName !== undefined)) {
    table.odtColumnStyleNames = columnStyleNames;
  }
  return table;
}

// --- package ---------------------------------------------------------------

function parsePart(text: string): Document {
  try {
    return parseXml(text);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new OdtOpenError(INVALID_ODT_MESSAGE);
    }
    throw error;
  }
}

/** Parses a `.odt` file's bytes into the sections/defaultStyles this
 * project's document model needs. Throws `FlatOdfError` for a flat XML ODF
 * (.fodt) file, or `OdtOpenError` for anything else that isn't a valid,
 * openable OpenDocument text package. */
export async function parseOdt(data: ArrayBuffer | Uint8Array): Promise<ParsedOdtDocument> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      if (isFlatOdf(bytes)) {
        throw new FlatOdfError(FLAT_ODF_MESSAGE);
      }
      throw new OdtOpenError(INVALID_ODT_MESSAGE);
    }
    throw error;
  }

  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }
  const contentDoc = parsePart(decodeUtf8(contentBytes));

  const stylesBytes = parts.get("styles.xml");
  const stylesDoc = stylesBytes ? parsePart(decodeUtf8(stylesBytes)) : undefined;

  const catalog = parseOdfStyleCatalog(contentDoc, stylesDoc);
  const tableColumnWidths = collectTableColumnWidths(contentDoc, stylesDoc);

  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const textRoot = body ? firstChildElement(body, OFFICE_NAMESPACE, "text") : undefined;
  if (!textRoot) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }

  const blocks: Block[] = [];
  const opaqueBlocks: OpaqueBlock[] = [];
  for (const child of Array.from(textRoot.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "p") {
      blocks.push(parseTextBlock(child, catalog, false));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "h") {
      blocks.push(parseTextBlock(child, catalog, true));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
      // Bulleted/numbered lists (SPEC-018 AC-21).
      parseListElement(child, catalog, 1, undefined, blocks, opaqueBlocks);
    } else if (child.namespaceURI === TABLE_NAMESPACE && child.localName === "table") {
      // Regular-grid tables are modelled (SPEC-022); anything the model
      // cannot express falls through to the opaque path below.
      const table = parseTableElement(child, catalog, tableColumnWidths);
      if (table) {
        blocks.push(table);
      } else {
        opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
      }
    } else {
      // Sequence declarations, frames, ... — unmodelled body children are
      // preserved verbatim, anchored to their position among the
      // paragraphs (SPEC-008).
      opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
    }
  }

  const fallback = createDefaultStyles();
  const defaultStyles: DefaultStyles = {
    pageSize: catalog.pageSize ?? fallback.pageSize,
    margins: catalog.margins ?? fallback.margins,
    fontFamily: catalog.defaultRunFormatting.fontFamily ?? fallback.fontFamily,
    fontSizePt: catalog.defaultRunFormatting.fontSizePt ?? fallback.fontSizePt,
  };

  const section: Section = { blocks };
  if (opaqueBlocks.length > 0) {
    section.opaqueBlocks = opaqueBlocks;
  }

  // content.xml and meta.xml are rebuilt on save (body/title/timestamp);
  // everything else — mimetype, styles.xml, the manifest, settings,
  // images, embedded objects — is retained verbatim for the lifetime of
  // the document (SPEC-008).
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (name !== "content.xml" && name !== "meta.xml") {
      opaque.push({ name, data: bytesToBase64(data) });
    }
  }

  const sourceXml: SourceXml = {};
  const fontFaceDecls = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "font-face-decls");
  if (fontFaceDecls) {
    sourceXml.odtFontFaceDecls = serializeElement(fontFaceDecls);
  }
  const automaticStyles = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "automatic-styles");
  if (automaticStyles) {
    // Paragraph/text-family styles are resolved into the model and
    // regenerated flat on save; every other automatic style (graphic,
    // table, list, … families) may be referenced by preserved opaque nodes
    // and is kept verbatim. Families never share a name scope in ODF, so
    // the regenerated and preserved definitions cannot collide.
    const kept: string[] = [];
    for (const def of Array.from(automaticStyles.children)) {
      const family =
        def.namespaceURI === STYLE_NAMESPACE && def.localName === "style"
          ? def.getAttributeNS(STYLE_NAMESPACE, "family")
          : null;
      if (family !== "paragraph" && family !== "text") {
        kept.push(serializeElement(def));
      }
    }
    if (kept.length > 0) {
      sourceXml.odtAutomaticStyles = kept;
    }
  }

  return { sections: [section], defaultStyles, opaque, sourceXml };
}
