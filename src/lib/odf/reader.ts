// Orchestrates the .odt -> DocumentModel pipeline: unzip the ODF package,
// detect the format from the package's `mimetype` entry (never the file
// extension), parse content.xml and styles.xml, map text:p/text:h/text:span
// onto the document model with ODF whitespace handling (text:s, text:tab,
// text:line-break), resolve style chains, and apply the master page's layout.
import type {
  Block,
  BookmarkAnchor,
  BorderLineStyle,
  BorderSide,
  CellBorders,
  CellVerticalAlign,
  DefaultStyles,
  ImageWrapMode,
  ListFormatting,
  MediaPart,
  NamedStyle,
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
  TableAlignment,
  TableCell,
  TableRow,
  TableWidthMode,
} from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { bytesToBase64 } from "../opaque";
import { ZipError, unzip } from "../docx/zip";
import { cmToTwips } from "../docx/units";
import { firstChildElement, parseXml, serializeElement, XmlParseError } from "../docx/xml";
import { sanitizeHref } from "../links";
import { clampListLevel, defaultMarkerStyle } from "../lists";
import { collectGraphicWrapStyles, parseDrawFrame } from "./drawing";
import { odfColorToCss, odfLengthToCm, odfLengthToPt, odfPercentage } from "./units";
import {
  DRAW_NAMESPACE,
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  foAttr,
  styleAttr,
  tableAttr,
  textAttr,
  xlinkAttr,
} from "./xml";
import {
  catalogToNamedStyles,
  paragraphAutoStyleSignals,
  parseOdfStyleCatalog,
  resolveListLevel,
  resolveParagraphStyle,
  resolveTextStyle,
  type OdfListLevelDef,
  type OdfMasterPageLayout,
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
  /** The *named* paragraph styles of styles.xml as the model's style
   * catalogue (SPEC-019); empty when the package defines none. Automatic
   * styles are direct formatting and stay baked (see catalogToNamedStyles). */
  styles: NamedStyle[];
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .odt. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from content.xml that the model cannot express:
   * font-face declarations and non-paragraph/text automatic styles that
   * preserved opaque nodes may reference. */
  sourceXml: SourceXml;
  /** Media parts modelled Image blocks reference (SPEC-024), one per
   * distinct source media part. */
  media: MediaPart[];
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
  /** Point bookmarks (SPEC-025), anchored at their character offset. */
  bookmarks: BookmarkAnchor[];
}

function accumulatedTextLength(acc: InlineAccumulator): number {
  return acc.pieces.reduce((sum, piece) => sum + piece.text.length, 0);
}

/** The run formatting a `text:a`'s children inherit: its xlink:href folded
 * in as linkAnchor (a "#fragment" target) or linkHref (external, sanitized
 * — a dangerous scheme leaves the text unlinked, AC-17). A fresh object, so
 * pieces inside the link never identity-merge with pieces outside it. */
function linkFormatting(el: Element, inherited: RunFormatting): RunFormatting {
  const href = xlinkAttr(el, "href");
  if (href === null || href.length === 0) {
    return inherited;
  }
  if (href.startsWith("#")) {
    return { ...inherited, linkAnchor: href.slice(1) };
  }
  const sanitized = sanitizeHref(href);
  return sanitized !== null ? { ...inherited, linkHref: sanitized } : inherited;
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
        // Hyperlinks (SPEC-025): the wrapper's target becomes link
        // formatting on its inline content, matching the .docx reader's
        // w:hyperlink handling; the element itself is not preserved.
        collectInline(el, linkFormatting(el, formatting), catalog, acc);
        break;
      case "bookmark":
      case "bookmark-start": {
        // Point bookmarks (SPEC-025); a ranged text:bookmark-start anchors
        // at its start position and its matching -end is dropped — the
        // writer re-emits a point text:bookmark.
        const name = textAttr(el, "name");
        if (name) {
          acc.bookmarks.push({ name, offset: accumulatedTextLength(acc) });
        }
        break;
      }
      case "bookmark-end":
        break;
      default:
        if (TRANSIENT_INLINE_ELEMENTS.has(el.localName)) {
          // Regenerated by the producer on open; deliberately not preserved.
          break;
        }
        // Fields, notes, … — unmodelled text-namespace elements are
        // preserved verbatim at their anchored position (SPEC-008).
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

  const acc: InlineAccumulator = { pieces: [], lastChar: "", opaqueInlines: [], mergeBarrier: false, bookmarks: [] };
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
  if (acc.bookmarks.length > 0) {
    paragraph.bookmarks = acc.bookmarks;
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

/** Context a body-level text-block parse needs to resolve a `draw:frame`
 * into actual image bytes (SPEC-024). */
export interface OdfDrawingContext {
  parts: Map<string, Uint8Array>;
  graphicWrapStyles: Map<string, ImageWrapMode>;
}

/** collectInline run over a detached clone of `nodes` — reused so a
 * mid-paragraph image split (parseBodyTextBlockOrSplit) gets exactly the
 * same whitespace-collapsing/span-resolution behaviour as an ordinary
 * paragraph, without re-implementing collectInline's state machine for a
 * node subset. Cloning is safe (the source tree is never mutated) and cheap
 * at this scale (paragraph-sized fragments only). */
function collectInlineSegment(
  nodes: Node[],
  formatting: RunFormatting,
  catalog: OdfStyleCatalog,
  ownerDocument: Document,
): InlineAccumulator {
  const acc: InlineAccumulator = { pieces: [], lastChar: "", opaqueInlines: [], mergeBarrier: false, bookmarks: [] };
  const temp = ownerDocument.createElementNS(TEXT_NAMESPACE, "text:tmp");
  for (const node of nodes) {
    temp.appendChild(node.cloneNode(true));
  }
  collectInline(temp, formatting, catalog, acc);
  return acc;
}

interface BodyTextBlockResult {
  blocks: Block[];
  media: MediaPart[];
}

/** Parses one body-level `text:p`/`text:h` into one or more blocks
 * (SPEC-024): a paragraph containing a modelled `draw:frame` picture splits
 * around each image, since the document model addresses an image as its own
 * top-level block (see documentModel.ts's Image doc comment) — mirroring
 * how a `table:table` already breaks the block sequence, just with the split
 * happening here instead of being structural in the source XML. Every
 * resulting paragraph piece keeps the source block's own style/heading
 * level. A frame nested inside a `text:span`/`text:a` is not detected (real
 * producers place image frames as direct paragraph children, never inside a
 * run) — such a paragraph parses exactly as parseTextBlock always did, so
 * this function is a pure superset, never a behaviour change for the common
 * case. Used only for section-level body content, not list items or table
 * cells — images are not modelled inside either (out of scope, SPEC-024). */
function parseBodyTextBlockOrSplit(
  el: Element,
  catalog: OdfStyleCatalog,
  isHeading: boolean,
  ctx: OdfDrawingContext,
): BodyTextBlockResult {
  const frameChildren = Array.from(el.children).filter(
    (child) => child.namespaceURI === DRAW_NAMESPACE && child.localName === "frame",
  );
  const hasImage = frameChildren.some((frame) => parseDrawFrame(frame, ctx.parts, ctx.graphicWrapStyles) !== null);
  if (!hasImage) {
    return { blocks: [parseTextBlock(el, catalog, isHeading)], media: [] };
  }

  const styleName = textAttr(el, "style-name") ?? undefined;
  const resolved = resolveParagraphStyle(catalog, styleName);
  let headingLevel: number | undefined;
  if (isHeading) {
    const attrValue = Number(textAttr(el, "outline-level") ?? "");
    headingLevel = (Number.isInteger(attrValue) && attrValue > 0 ? attrValue : resolved.headingLevel) ?? 1;
  }

  const blocks: Block[] = [];
  const media: MediaPart[] = [];
  let segment: Node[] = [];
  const flush = () => {
    if (segment.length === 0) {
      return;
    }
    const acc = collectInlineSegment(segment, resolved.runFormatting, catalog, el.ownerDocument);
    segment = [];
    const runs: Run[] = acc.pieces
      .filter((piece) => piece.text.length > 0)
      .map((piece) => ({ text: piece.text, formatting: { ...piece.formatting } }));
    if (runs.length === 0 && acc.opaqueInlines.length === 0 && acc.bookmarks.length === 0) {
      return;
    }
    const formatting: ParagraphFormatting = { ...resolved.paragraphFormatting };
    normalizeResolvedDirection(formatting);
    const paragraph: Paragraph = { type: "paragraph", runs, formatting };
    if (styleName) {
      paragraph.styleId = styleName;
    }
    if (acc.opaqueInlines.length > 0) {
      paragraph.opaqueInlines = acc.opaqueInlines;
    }
    if (acc.bookmarks.length > 0) {
      paragraph.bookmarks = acc.bookmarks;
    }
    if (headingLevel !== undefined) {
      paragraph.headingLevel = headingLevel;
    }
    blocks.push(paragraph);
  };
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const frameEl = node as Element;
      if (frameEl.namespaceURI === DRAW_NAMESPACE && frameEl.localName === "frame") {
        const parsed = parseDrawFrame(frameEl, ctx.parts, ctx.graphicWrapStyles);
        if (parsed) {
          flush();
          blocks.push(parsed.image);
          media.push(parsed.media);
          continue;
        }
      }
    }
    segment.push(node);
  }
  flush();
  return { blocks, media };
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

/** ODF's `fo:border`-family shorthand ("0.5pt solid #ff0000") or "none"/
 * "hidden" (SPEC-023). */
function parseBorderShorthand(value: string | null): BorderSide | undefined {
  if (!value) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "none" || trimmed === "hidden") {
    return { style: "none", widthPt: 0, color: "#000000" };
  }
  const match = /^(\S+)\s+(\S+)\s+(#[0-9a-fA-F]{6})$/.exec(trimmed);
  if (!match) {
    return undefined;
  }
  const widthPt = odfLengthToPt(match[1]) ?? 1;
  const style = mapOdfBorderStyle(match[2]);
  const color = odfColorToCss(match[3]) ?? "#000000";
  return { style, widthPt, color };
}

/** ODF's CSS2-derived border style keywords collapse to the model's small
 * enum (SPEC-023), the same simplification the .docx reader applies. */
function mapOdfBorderStyle(word: string): BorderLineStyle {
  const w = word.toLowerCase();
  if (w === "none" || w === "hidden") {
    return "none";
  }
  if (w === "double") {
    return "double";
  }
  if (w === "dashed") {
    return "dashed";
  }
  if (w === "dotted") {
    return "dotted";
  }
  return "single"; // "solid", "groove", "ridge", "inset", "outset", …
}

/** A `style:table-cell-properties` element's SPEC-023 fields: per-side
 * borders (falling back to the `fo:border` shorthand for any side without
 * its own `fo:border-*`), background shading, padding and vertical
 * alignment. */
interface ParsedCellStyle {
  borders?: CellBorders;
  shadingFill?: string;
  paddingCm?: number;
  verticalAlign?: CellVerticalAlign;
}

function parseCellStyleProperties(props: Element): ParsedCellStyle {
  const shorthand = foAttr(props, "border");
  const top = parseBorderShorthand(foAttr(props, "border-top") ?? shorthand);
  const bottom = parseBorderShorthand(foAttr(props, "border-bottom") ?? shorthand);
  const left = parseBorderShorthand(foAttr(props, "border-left") ?? shorthand);
  const right = parseBorderShorthand(foAttr(props, "border-right") ?? shorthand);
  const result: ParsedCellStyle = {};
  if (top || bottom || left || right) {
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
    result.borders = borders;
  }
  const shadingFill = odfColorToCss(foAttr(props, "background-color"));
  if (shadingFill !== undefined) {
    result.shadingFill = shadingFill;
  }
  const paddingCm = odfLengthToCm(foAttr(props, "padding-top") ?? foAttr(props, "padding"));
  if (paddingCm !== undefined) {
    result.paddingCm = paddingCm;
  }
  const vAlign = styleAttr(props, "vertical-align");
  if (vAlign === "top" || vAlign === "bottom") {
    result.verticalAlign = vAlign;
  } else if (vAlign === "middle") {
    result.verticalAlign = "center";
  }
  return result;
}

/** A `style:style` (family "table") element's SPEC-023 width fields:
 * `style:rel-width` is a percentage of the text area; its absence means a
 * fixed width (the `columnWidthsCm` sum) — ODF text tables have no distinct
 * "size to content" mode the way `w:tblLayout="autofit"` does, so `"auto"`
 * is reachable only by setting it in the editor, not by reading one. */
interface ParsedTableStyle {
  widthMode?: TableWidthMode;
  widthPercent?: number;
}

/** The three style families a table references, collected once per document
 * (mirroring collectTableColumnWidths' scan-both-parts pattern) and threaded
 * through parseTableElement/parseTableCellElement. */
interface TableStyleTables {
  columnWidths: Map<string, number>;
  cellStyles: Map<string, ParsedCellStyle>;
  rowMinHeights: Map<string, number>;
  tableStyles: Map<string, ParsedTableStyle>;
}

function collectTableCellStyles(contentDoc: Document, stylesDoc: Document | undefined): Map<string, ParsedCellStyle> {
  const styles = new Map<string, ParsedCellStyle>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-cell") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-cell-properties");
      if (name && props) {
        styles.set(name, parseCellStyleProperties(props));
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return styles;
}

function collectTableRowMinHeights(contentDoc: Document, stylesDoc: Document | undefined): Map<string, number> {
  const heights = new Map<string, number>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-row") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-row-properties");
      // `min-row-height` (a floor, matching the model) takes precedence over
      // `row-height` (a fixed height, approximated the same way as the
      // .docx reader treats `w:hRule="exact"`).
      const height = props
        ? (odfLengthToCm(styleAttr(props, "min-row-height")) ?? odfLengthToCm(styleAttr(props, "row-height")))
        : undefined;
      if (name && height !== undefined) {
        heights.set(name, height);
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return heights;
}

function collectTableStyles(contentDoc: Document, stylesDoc: Document | undefined): Map<string, ParsedTableStyle> {
  const styles = new Map<string, ParsedTableStyle>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-properties");
      if (!name || !props) {
        continue;
      }
      const relWidth = odfPercentage(styleAttr(props, "rel-width"));
      styles.set(
        name,
        relWidth !== undefined
          ? { widthMode: "percent", widthPercent: Math.min(100, Math.max(1, Math.round(relWidth))) }
          : { widthMode: "fixed" },
      );
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return styles;
}

/** `table:align` on the `table:table` element itself (SPEC-023 AC-22, not a
 * style property): "margins" is ODF's own "left, indented like a paragraph"
 * default and reads the same as "left" — table RTL is out of scope. */
function parseTableAlignmentAttr(tableEl: Element): TableAlignment | undefined {
  const val = tableAttr(tableEl, "align");
  if (val === "center" || val === "right") {
    return val;
  }
  if (val === "left" || val === "margins") {
    return "left";
  }
  return undefined;
}

function parseTableCellElement(
  cellEl: Element,
  catalog: OdfStyleCatalog,
  cellStyles: Map<string, ParsedCellStyle>,
): TableCell | null {
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
  // Merged cells (SPEC-023 AC-28): the origin carries its own span; the
  // grid slots it covers are separate `table:covered-table-cell` elements
  // (handled by the caller, not here).
  const colSpan = Number(tableAttr(cellEl, "number-columns-spanned") ?? "1");
  if (Number.isInteger(colSpan) && colSpan > 1) {
    cell.colSpan = colSpan;
  }
  const rowSpan = Number(tableAttr(cellEl, "number-rows-spanned") ?? "1");
  if (Number.isInteger(rowSpan) && rowSpan > 1) {
    cell.rowSpan = rowSpan;
  }
  const styleName = tableAttr(cellEl, "style-name");
  if (styleName !== null) {
    cell.odtStyleName = styleName;
    const parsedStyle = cellStyles.get(styleName);
    if (parsedStyle?.borders) {
      cell.borders = parsedStyle.borders;
    }
    if (parsedStyle?.shadingFill !== undefined) {
      cell.shadingFill = parsedStyle.shadingFill;
    }
    if (parsedStyle?.paddingCm !== undefined) {
      cell.paddingCm = parsedStyle.paddingCm;
    }
    if (parsedStyle?.verticalAlign !== undefined) {
      cell.verticalAlign = parsedStyle.verticalAlign;
    }
  }
  return cell;
}

/** Parses a `table:table` into a modelled Table (SPEC-022; merged cells and
 * formatting SPEC-023), or null when it uses features outside the model's
 * regular grid (header rows, nested tables, sub-tables, …) — the caller
 * then preserves the element verbatim as an opaque block (SPEC-008).
 *
 * A merged cell's covered neighbours are explicit `table:covered-table-cell`
 * elements in the source — unlike `.docx`'s inferred `w:gridSpan`/`w:vMerge`
 * chains, ODF already lists one element per grid slot, so no second pass is
 * needed: each becomes a `covered: true` placeholder directly. */
function parseTableElement(tableEl: Element, catalog: OdfStyleCatalog, styleTables: TableStyleTables): Table | null {
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
        if (cellEl.namespaceURI !== TABLE_NAMESPACE) {
          return null;
        }
        const cellRepeat = repeatCount(cellEl, "number-columns-repeated");
        if (cellRepeat === null) {
          return null;
        }
        let cell: TableCell | null;
        if (cellEl.localName === "table-cell") {
          cell = parseTableCellElement(cellEl, catalog, styleTables.cellStyles);
        } else if (cellEl.localName === "covered-table-cell") {
          cell = { blocks: [], covered: true };
        } else {
          return null;
        }
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
      const minHeightCm = rowStyleName ? styleTables.rowMinHeights.get(rowStyleName) : undefined;
      for (let i = 0; i < repeat; i++) {
        const row: TableRow = { cells: i === 0 ? cells : structuredClone(cells) };
        if (rowStyleName !== null) {
          row.odtStyleName = rowStyleName;
        }
        if (minHeightCm !== undefined) {
          row.minHeightCm = minHeightCm;
        }
        rows.push(row);
      }
    } else {
      // table-header-rows, sub-tables, title/desc, … — the model cannot
      // express these; keep the table verbatim.
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
  const declared = columnStyleNames.map((name) => (name ? styleTables.columnWidths.get(name) : undefined));
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
    const parsedStyle = styleTables.tableStyles.get(styleName);
    if (parsedStyle?.widthMode !== undefined) {
      table.widthMode = parsedStyle.widthMode;
    }
    if (parsedStyle?.widthPercent !== undefined) {
      table.widthPercent = parsedStyle.widthPercent;
    }
  }
  const alignment = parseTableAlignmentAttr(tableEl);
  if (alignment !== undefined) {
    table.alignment = alignment;
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
  const tableStyleTables: TableStyleTables = {
    columnWidths: tableColumnWidths,
    cellStyles: collectTableCellStyles(contentDoc, stylesDoc),
    rowMinHeights: collectTableRowMinHeights(contentDoc, stylesDoc),
    tableStyles: collectTableStyles(contentDoc, stylesDoc),
  };

  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const textRoot = body ? firstChildElement(body, OFFICE_NAMESPACE, "text") : undefined;
  if (!textRoot) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }

  const drawingCtx: OdfDrawingContext = { parts, graphicWrapStyles: collectGraphicWrapStyles(contentDoc, stylesDoc) };
  const mediaById = new Map<string, MediaPart>();
  const sections: Section[] = [];
  let blocks: Block[] = [];
  let opaqueBlocks: OpaqueBlock[] = [];
  // The geometry of the master page the CURRENT section's first paragraph
  // switched to (SPEC-028 AC-33); undefined for the first section, whose
  // master is the document default folded into defaultStyles below.
  let sectionGeometry: OdfMasterPageLayout | undefined;
  // The master page named by the document's very FIRST body paragraph. The
  // reader never opens an empty leading section for such a reference (see
  // the guard below) — but that master, not the first-listed one, is the one
  // the body actually starts on, so it must supply the document defaults and
  // the first section's geometry (SPEC-028 AC-17/18/19/33). Undefined when
  // the first paragraph names no master: the first-listed master stays the
  // default, exactly as before.
  let leadingMasterName: string | undefined;
  const finishSection = () => {
    const section: Section = { blocks };
    if (opaqueBlocks.length > 0) {
      section.opaqueBlocks = opaqueBlocks;
    }
    // Explicit geometry on every later section — the master page pins it,
    // so it must not float with a later defaultStyles edit. `start` stays
    // absent: ODF's master-page switch is always a page break (nextPage).
    if (sectionGeometry?.pageSize) {
      section.pageSize = { ...sectionGeometry.pageSize };
    }
    if (sectionGeometry?.margins) {
      section.margins = { ...sectionGeometry.margins };
    }
    sections.push(section);
  };
  for (const child of Array.from(textRoot.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h")) {
      const signals = paragraphAutoStyleSignals(catalog, textAttr(child, "style-name") ?? undefined);
      // A style:master-page-name on the paragraph's automatic style starts
      // a new section here (SPEC-028 AC-33) — except on the document's very
      // first block, where opening an empty leading section would be wrong.
      // But "first master page" is then the one this paragraph actually
      // names — NOT whichever is listed first in master-styles (a foreign
      // doc can list Standard first and have the body start on A5Master) —
      // so record it and let it provide the document defaults below.
      if (signals.masterPageName !== undefined) {
        if (sections.length === 0 && blocks.length === 0) {
          leadingMasterName = signals.masterPageName;
        } else {
          finishSection();
          blocks = [];
          opaqueBlocks = [];
          sectionGeometry = catalog.masterPageLayouts.get(signals.masterPageName);
        }
      }
      // fo:break-before="page" on the AUTOMATIC style materialises a manual
      // BreakBlock before the paragraph, and the flag is dropped from the
      // paragraph's own formatting (SPEC-028 AC-26/32). ODF has one channel
      // for both the manual break and the pageBreakBefore formatting flag,
      // so .odt import canonicalises the ambiguity to the break block —
      // what the writer's fo:break-before actually stood for; only a NAMED
      // style's break-before stays pure formatting (see
      // paragraphAutoStyleSignals).
      if (signals.pageBreakBefore) {
        blocks.push({ type: "break", kind: "page" });
      } else if (signals.columnBreakBefore) {
        // fo:break-before="column" (one attribute — mutually exclusive with
        // "page") likewise materialises a column BreakBlock (AC-10/11); no
        // formatting flag to strip afterwards, since the model has no
        // column counterpart of pageBreakBefore (SPEC-030).
        blocks.push({ type: "break", kind: "column" });
      }
      const { blocks: parsedBlocks, media } = parseBodyTextBlockOrSplit(
        child,
        catalog,
        child.localName === "h",
        drawingCtx,
      );
      if (signals.pageBreakBefore) {
        for (const parsed of parsedBlocks) {
          if (parsed.type === "paragraph" && parsed.formatting) {
            delete parsed.formatting.pageBreakBefore;
          }
        }
      }
      blocks.push(...parsedBlocks);
      for (const part of media) {
        mediaById.set(part.mediaId, part);
      }
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
      // Bulleted/numbered lists (SPEC-018 AC-21).
      parseListElement(child, catalog, 1, undefined, blocks, opaqueBlocks);
    } else if (child.namespaceURI === TABLE_NAMESPACE && child.localName === "table") {
      // Regular-grid tables are modelled (SPEC-022); anything the model
      // cannot express falls through to the opaque path below.
      const table = parseTableElement(child, catalog, tableStyleTables);
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
  finishSection();

  const fallback = createDefaultStyles();
  // The document default page geometry is the master the body actually
  // starts on: the master named by the first body paragraph when there is
  // one, else the first-listed master folded into catalog.pageSize/margins.
  // parsePageLayout folds the first-listed master for the no-lead case; when
  // a lead exists it wins (SPEC-028 AC-17/18/19/33 — a foreign doc must read
  // back the geometry it actually starts on, not whichever master happens to
  // be listed first).
  const leadingGeometry = leadingMasterName
    ? catalog.masterPageLayouts.get(leadingMasterName)
    : undefined;
  // The leading section's geometry is the master the very first paragraph
  // named, so it carries explicit pageSize/margins like every later section
  // that switches masters — reported by AC-17/18/19/33 and rendered/round-
  // tripped correctly rather than silently inheriting the first-LISTED
  // master. A document whose first paragraph names no master keeps the
  // first-listed master folded into defaults exactly as before.
  const leadingSection = sections[0];
  if (leadingSection && leadingGeometry?.pageSize) {
    leadingSection.pageSize = { ...leadingGeometry.pageSize };
  }
  if (leadingSection && leadingGeometry?.margins) {
    leadingSection.margins = { ...leadingGeometry.margins };
  }
  const defaultStyles: DefaultStyles = {
    pageSize: leadingGeometry?.pageSize ?? catalog.pageSize ?? fallback.pageSize,
    margins: leadingGeometry?.margins ?? catalog.margins ?? fallback.margins,
    fontFamily: catalog.defaultRunFormatting.fontFamily ?? fallback.fontFamily,
    fontSizePt: catalog.defaultRunFormatting.fontSizePt ?? fallback.fontSizePt,
  };

  // content.xml and meta.xml are rebuilt on save (body/title/timestamp);
  // everything else — mimetype, styles.xml, the manifest, settings,
  // embedded objects — is retained verbatim for the lifetime of the
  // document (SPEC-008). Media parts a modelled Image references (SPEC-024)
  // are excluded here — they travel in `media` instead (see docx/reader.ts's
  // parseDocx for why: the writer regenerates their manifest/relationship
  // entries from `media` on every save, same-format or not).
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (name !== "content.xml" && name !== "meta.xml" && !mediaById.has(name)) {
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

  return {
    sections,
    defaultStyles,
    styles: catalogToNamedStyles(catalog),
    opaque,
    sourceXml,
    media: Array.from(mediaById.values()),
  };
}
