// Pure block builders (SPEC-045 AC-1): one function per construct the API
// generates — heading, paragraph, list, table, image, page break, table of
// contents. Both the fluent `DocumentBuilder` (builder.ts) and the JSON
// document-description dispatcher (description.ts) call these same
// functions with already-validated primitive arguments, which is what
// makes the two surfaces a 1:1 mapping (AC-4) rather than two independent
// implementations. No serialisation happens here — these only ever produce
// `Block`s of the canonical model (documentModel.ts), which the existing
// writers (docx/writer.ts, odf/writer.ts) turn into bytes unchanged.
import {
  headingStyleId,
  resolveNamedStyle,
} from "../styles";
import { decodeImageDimensions } from "../imageDimensions";
import { fitImageSizeCm, textAreaWidthCm } from "../images";
import { bytesToBase64 } from "../opaque";
import { equalColumnWidthsCm } from "../tables";
import { tocEntriesFromHeadings } from "../toc";
import { isToc } from "../documentModel";
import type {
  Block,
  BreakBlock,
  DocumentModel,
  HeaderFooterContent,
  Image,
  ImageWrapMode,
  ListFormatting,
  ListType,
  Paragraph,
  ParagraphAlignment,
  ParagraphFormatting,
  Run,
  RunFormatting,
  Table,
  TableAlignment,
  TableCell,
  TableOfContents,
  TableRow,
  TableWidthMode,
  TocOptions,
} from "../documentModel";
import { DocumentGenerationError } from "./errors";
import { emuToCm, lengthToEmu, type LengthInput } from "./length";
import { buildRunsFromContent, contentFromTextOrRuns } from "./runs";
import {
  expectArray,
  expectBoolean,
  expectEnum,
  expectIntInRange,
  expectNonNegativeNumber,
  expectPlainObject,
  expectString,
  rejectUnknownKeys,
} from "./validate";

/** What every block builder needs from the document under construction: its
 * default page geometry (for image/table default sizing) and its style
 * catalogue (for heading formatting) — never its blocks, so building one
 * block never depends on another having been added first (TOC is the one
 * exception, resolved separately once the whole body exists — see
 * builder.ts's `finalizeTocEntries`). */
export interface BuildContext {
  doc: DocumentModel;
}

export const PARAGRAPH_ALIGNMENTS: readonly ParagraphAlignment[] = ["left", "center", "right", "justify"];

/** Bakes a run's effective formatting (AC-14/15/18's round-trip contract):
 * the document's default font family/size, then `styleFormatting` (a
 * heading's built-in style, empty for anything else), then the run's own
 * explicit formatting — highest precedence last. This mirrors what both
 * readers do when opening any real file: `docDefaultRunFormatting` is
 * always the base every run resolves against (docx/styles.ts's
 * `resolveStyle`), so a run this API leaves "unset" and a run a reader
 * resolves to the same default are the same content — `equalStructural`
 * only agrees if this API bakes it too, since it does not know a
 * document's own defaults are what produced an absent field. */
function bakeRun(ctx: BuildContext, run: Run, styleFormatting: RunFormatting = {}): Run {
  return {
    ...run,
    formatting: {
      fontFamily: ctx.doc.defaultStyles.fontFamily,
      fontSizePt: ctx.doc.defaultStyles.fontSizePt,
      ...styleFormatting,
      ...run.formatting,
    },
  };
}

function bakeRuns(ctx: BuildContext, runs: Run[], styleFormatting: RunFormatting = {}): Run[] {
  return runs.map((run) => bakeRun(ctx, run, styleFormatting));
}

// --- heading -----------------------------------------------------------

export const HEADING_OPTION_KEYS = ["alignment"] as const;

export interface HeadingOptions {
  alignment?: ParagraphAlignment;
}

/** A heading paragraph (AC-1): its formatting is baked directly from the
 * built-in Heading{level} style (lib/styles.ts) — the same visual result as
 * applying that style in the editor — without referencing the style by id,
 * so a generated document never needs a style catalogue of its own (see
 * builder.ts's doc comment for why). `headingLevel` alone is what both
 * writers and lib/toc.ts's heading scan key off; it round-trips via
 * `w:outlineLvl` (.docx) / `text:outline-level` (.odt) regardless of style
 * references. */
export function buildHeadingParagraph(
  ctx: BuildContext,
  content: unknown,
  level: unknown,
  options: HeadingOptions,
  path: string,
): Paragraph {
  const lvl = expectIntInRange(level, `${path}.level`, 1, 6);
  const runs = buildRunsFromContent(content, path);
  const resolved = resolveNamedStyle(ctx.doc, headingStyleId(lvl));
  const bakedRuns = bakeRuns(ctx, runs, resolved.runFormatting);
  const paragraph: Paragraph = { type: "paragraph", runs: bakedRuns, headingLevel: lvl };
  const formatting: ParagraphFormatting = { ...resolved.paragraphFormatting };
  if (options.alignment !== undefined) {
    formatting.alignment = options.alignment;
  }
  if (Object.keys(formatting).length > 0) {
    paragraph.formatting = formatting;
  }
  return paragraph;
}

// --- paragraph -----------------------------------------------------------

export const PARAGRAPH_OPTION_KEYS = ["alignment", "spaceBeforePt", "spaceAfterPt"] as const;

export interface ParagraphOptions {
  alignment?: ParagraphAlignment;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

export function buildParagraph(ctx: BuildContext, content: unknown, options: ParagraphOptions, path: string): Paragraph {
  const runs = bakeRuns(ctx, buildRunsFromContent(content, path));
  const paragraph: Paragraph = { type: "paragraph", runs };
  const formatting: ParagraphFormatting = {};
  if (options.alignment !== undefined) {
    formatting.alignment = options.alignment;
  }
  if (options.spaceBeforePt !== undefined) {
    formatting.spaceBeforePt = options.spaceBeforePt;
  }
  if (options.spaceAfterPt !== undefined) {
    formatting.spaceAfterPt = options.spaceAfterPt;
  }
  if (Object.keys(formatting).length > 0) {
    paragraph.formatting = formatting;
  }
  return paragraph;
}

// --- list ------------------------------------------------------------------

export const LIST_OPTION_KEYS = ["style", "startAt"] as const;

export interface ListOptions {
  style?: "bullet" | "number";
  startAt?: number;
}

/** One paragraph per item, single-level (nesting is out of scope — SPEC-045
 * lists this construct without mentioning levels). List identity is
 * adjacency (documentModel.ts's ListFormatting doc comment): consecutive
 * items form one list purely by sitting next to each other in the body, no
 * list id needed. */
export function buildListParagraphs(
  ctx: BuildContext,
  items: unknown,
  options: ListOptions,
  path: string,
): Paragraph[] {
  const arr = expectArray(items, path);
  if (arr.length === 0) {
    throw new DocumentGenerationError(`${path}: expected at least one item`);
  }
  const listType: ListType = options.style === "number" ? "number" : "bullet";
  return arr.map((item, i) => {
    const runs = bakeRuns(ctx, buildRunsFromContent(item, `${path}[${i}]`));
    const list: ListFormatting = { type: listType, level: 1 };
    if (options.startAt !== undefined) {
      // Carried on every item, not just the first: documentModel.ts's
      // ListFormatting.startAt doc comment says a real reader does the
      // same ("carried on every item of an imported list so a restart
      // after an interruption returns to the list's own start value") —
      // matching that is what keeps AC-14/15/18's round trip exact.
      list.startAt = options.startAt;
    }
    return { type: "paragraph", runs, list };
  });
}

// --- table -------------------------------------------------------------

export interface TableSpec {
  rows: unknown;
  columnWidths?: unknown;
  header?: unknown;
  alignment?: unknown;
  widthMode?: unknown;
  widthPercent?: unknown;
}

const TABLE_ALIGNMENTS: readonly TableAlignment[] = ["left", "center", "right"];
const TABLE_WIDTH_MODES: readonly TableWidthMode[] = ["auto", "fixed", "percent"];

export const TABLE_SPEC_KEYS = [
  "rows",
  "columnWidths",
  "header",
  "alignment",
  "widthMode",
  "widthPercent",
] as const;

function buildTableCell(ctx: BuildContext, spec: unknown, path: string): TableCell {
  const runs = bakeRuns(ctx, buildRunsFromContent(spec, path));
  return { blocks: [{ type: "paragraph", runs }] };
}

/** A table block (AC-1) from a raw, unvalidated spec — called both from the
 * JSON dispatcher and directly from the fluent builder, so validation
 * (including AC-9's unknown-property check) happens here once regardless
 * of entry path. */
export function buildTable(ctx: BuildContext, rawSpec: unknown, path: string): Table {
  const spec = expectPlainObject(rawSpec, path);
  rejectUnknownKeys(spec, TABLE_SPEC_KEYS, path);
  const rowsRaw = expectArray(spec.rows, `${path}.rows`);
  if (rowsRaw.length === 0) {
    throw new DocumentGenerationError(`${path}.rows: expected at least one row`);
  }
  const firstRow = expectArray(rowsRaw[0], `${path}.rows[0]`);
  const columns = firstRow.length;
  if (columns === 0) {
    throw new DocumentGenerationError(`${path}.rows[0]: expected at least one cell`);
  }
  const header = spec.header !== undefined ? expectBoolean(spec.header, `${path}.header`) : false;
  const rows: TableRow[] = rowsRaw.map((rowSpec, r) => {
    const cellsRaw = expectArray(rowSpec, `${path}.rows[${r}]`);
    if (cellsRaw.length !== columns) {
      throw new DocumentGenerationError(
        `${path}.rows[${r}]: expected ${columns} cells (every row must have the same number of cells as row 0)`,
      );
    }
    const cells = cellsRaw.map((cellSpec, c) => buildTableCell(ctx, cellSpec, `${path}.rows[${r}][${c}]`));
    const row: TableRow = { cells };
    if (r === 0 && header) {
      row.isHeader = true;
    }
    return row;
  });
  const columnWidthsCm =
    spec.columnWidths !== undefined
      ? expectArray(spec.columnWidths, `${path}.columnWidths`).map((w, i) =>
          emuToCm(lengthToEmu(w, `${path}.columnWidths[${i}]`)),
        )
      : equalColumnWidthsCm(ctx.doc, columns);
  if (columnWidthsCm.length !== columns) {
    throw new DocumentGenerationError(`${path}.columnWidths: expected ${columns} entries`);
  }
  // Both readers always resolve a table's width mode explicitly — "fixed"
  // when a source declares none (docx/reader.ts's parseTableWidth,
  // odf/reader.ts's table-width parsing) — so this is baked the same way
  // even when the caller never names it, for the same round-trip reason
  // bakeRun bakes run defaults (AC-14/15/18).
  const table: Table = { type: "table", rows, columnWidthsCm, widthMode: "fixed" };
  if (spec.alignment !== undefined) {
    table.alignment = expectEnum(spec.alignment, TABLE_ALIGNMENTS, `${path}.alignment`);
  }
  if (spec.widthMode !== undefined) {
    table.widthMode = expectEnum(spec.widthMode, TABLE_WIDTH_MODES, `${path}.widthMode`);
  }
  if (spec.widthPercent !== undefined) {
    table.widthPercent = expectNonNegativeNumber(spec.widthPercent, `${path}.widthPercent`);
  }
  return table;
}

// --- image -------------------------------------------------------------

const ACCEPTED_IMAGE_CONTENT_TYPES = ["image/png", "image/jpeg", "image/gif"] as const;
const IMAGE_WRAP_MODES: readonly ImageWrapMode[] = ["inline", "square", "behind", "infront"];

export const IMAGE_SPEC_KEYS = ["data", "contentType", "width", "height", "altText", "wrap"] as const;

export interface ImageSpec {
  data: unknown;
  contentType: unknown;
  width?: LengthInput;
  height?: LengthInput;
  altText?: unknown;
  wrap?: unknown;
}

function normalizeImageBytes(data: unknown, path: string): Uint8Array {
  if (typeof data === "string") {
    // Base64, the only shape JSON can carry (AC-4's description format);
    // the fluent builder additionally accepts raw bytes directly (below).
    let binary: string;
    try {
      binary = atob(data);
    } catch {
      throw new DocumentGenerationError(`${path}.data: not valid base64`);
    }
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }
  if (data instanceof Uint8Array) {
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  throw new DocumentGenerationError(
    `${path}.data: expected a base64 string, Uint8Array, ArrayBuffer or Buffer`,
  );
}

/** An image block (AC-1) from a raw, unvalidated spec — same "validate here
 * once" contract as buildTable. */
export function buildImageBlock(ctx: BuildContext, rawSpec: unknown, path: string): Image {
  const spec = expectPlainObject(rawSpec, path);
  rejectUnknownKeys(spec, IMAGE_SPEC_KEYS, path);
  const contentType = expectEnum(spec.contentType, ACCEPTED_IMAGE_CONTENT_TYPES, `${path}.contentType`);
  if (spec.data === undefined) {
    throw new DocumentGenerationError(`${path}.data: expected image bytes`);
  }
  const bytes = normalizeImageBytes(spec.data, path);
  const dims = decodeImageDimensions(bytes, contentType);
  if (!dims) {
    throw new DocumentGenerationError(`${path}.data: could not decode image dimensions for ${contentType}`);
  }
  const fitted = fitImageSizeCm(dims.widthPx, dims.heightPx, textAreaWidthCm(ctx.doc));
  const widthCm = spec.width !== undefined ? emuToCm(lengthToEmu(spec.width, `${path}.width`)) : fitted.widthCm;
  const heightCm =
    spec.height !== undefined ? emuToCm(lengthToEmu(spec.height, `${path}.height`)) : fitted.heightCm;
  const wrap = spec.wrap !== undefined ? expectEnum(spec.wrap, IMAGE_WRAP_MODES, `${path}.wrap`) : "inline";
  const mediaId = crypto.randomUUID();
  ctx.doc.media = [...(ctx.doc.media ?? []), { mediaId, contentType, data: bytesToBase64(bytes) }];
  const image: Image = {
    type: "image",
    mediaId,
    naturalWidthPx: dims.widthPx,
    naturalHeightPx: dims.heightPx,
    widthCm,
    heightCm,
    wrap,
  };
  if (spec.altText !== undefined) {
    image.altText = expectString(spec.altText, `${path}.altText`);
  }
  return image;
}

// --- page break --------------------------------------------------------

export function buildPageBreak(): BreakBlock {
  return { type: "break", kind: "page" };
}

// --- table of contents ---------------------------------------------------

export const TOC_SPEC_KEYS = ["levels", "showPageNumbers", "hyperlinksOnly"] as const;

export interface TocSpec {
  levels?: unknown;
  showPageNumbers?: unknown;
  hyperlinksOnly?: unknown;
}

/** A TOC block with empty `entries` (AC-1): its real entries are filled in
 * once the whole document exists (lib/toc.ts's `tocEntriesFromHeadings`
 * scans every heading in the body, including ones added after this call —
 * see builder.ts's `finalizeTocEntries`, run once at output time). Same
 * "validate here once" contract as buildTable/buildImageBlock. */
export function buildTocBlock(rawSpec: unknown, path: string): TableOfContents {
  const spec = expectPlainObject(rawSpec ?? {}, path);
  rejectUnknownKeys(spec, TOC_SPEC_KEYS, path);
  const levels =
    spec.levels !== undefined
      ? expectArray(spec.levels, `${path}.levels`).map((l, i) => expectIntInRange(l, `${path}.levels[${i}]`, 1, 6))
      : [1, 2, 3];
  if (levels.length === 0) {
    throw new DocumentGenerationError(`${path}.levels: expected at least one level`);
  }
  const showPageNumbers =
    spec.showPageNumbers !== undefined ? expectBoolean(spec.showPageNumbers, `${path}.showPageNumbers`) : true;
  const hyperlinksOnly =
    spec.hyperlinksOnly !== undefined ? expectBoolean(spec.hyperlinksOnly, `${path}.hyperlinksOnly`) : false;
  const options: TocOptions = { levels, showPageNumbers, hyperlinksOnly };
  return { type: "toc", options, entries: [] };
}

// --- block dispatch (shared by the JSON description and header/footer) ---

/** Strips the block-spec object's own `type` discriminator before handing
 * the rest to a builder (buildTable/buildImageBlock/buildTocBlock) that
 * validates its own key set from scratch — `type` is never one of theirs. */
function withoutType(obj: Record<string, unknown>): Record<string, unknown> {
  const rest: Record<string, unknown> = {};
  for (const key of Object.keys(obj)) {
    if (key !== "type") {
      rest[key] = obj[key];
    }
  }
  return rest;
}

/** Block types valid in the document body (AC-1). */
export const BODY_BLOCK_TYPES = [
  "heading",
  "paragraph",
  "list",
  "table",
  "image",
  "pageBreak",
  "toc",
] as const;

/** Block types valid inside a header/footer: page breaks and a table of
 * contents are document/section-level constructs the writers already
 * ignore when they appear in a header/footer part (docx/writer.ts's
 * headerFooterPartXml, odf/writer.ts's paragraphXml path), so this API
 * rejects them there up front with a clear error instead of silently
 * emitting nothing. */
export const HEADER_FOOTER_BLOCK_TYPES = ["heading", "paragraph", "list", "table", "image"] as const;

/** One JSON block-description object -> zero or more `Block`s (a list
 * expands to one paragraph per item). Shared by description.ts's body/
 * header/footer interpreters — this is the "1:1 mapping" AC-4 asks for. */
export function buildBlockFromSpec(
  ctx: BuildContext,
  spec: unknown,
  allowedTypes: readonly string[],
  path: string,
): Block[] {
  const obj = expectPlainObject(spec, path);
  const type = expectString(obj.type, `${path}.type`);
  if (!allowedTypes.includes(type)) {
    throw new DocumentGenerationError(
      `${path}.type: expected one of ${allowedTypes.map((t) => `"${t}"`).join(", ")}, got "${type}"`,
    );
  }
  switch (type) {
    case "heading": {
      rejectUnknownKeys(obj, ["type", "text", "runs", "level", "alignment"], path);
      const content = contentFromTextOrRuns(obj, path);
      const alignment =
        obj.alignment !== undefined ? expectEnum(obj.alignment, PARAGRAPH_ALIGNMENTS, `${path}.alignment`) : undefined;
      return [buildHeadingParagraph(ctx, content, obj.level, { alignment }, path)];
    }
    case "paragraph": {
      rejectUnknownKeys(obj, ["type", "text", "runs", "alignment", "spaceBeforePt", "spaceAfterPt"], path);
      const content = contentFromTextOrRuns(obj, path);
      const options: ParagraphOptions = {};
      if (obj.alignment !== undefined) {
        options.alignment = expectEnum(obj.alignment, PARAGRAPH_ALIGNMENTS, `${path}.alignment`);
      }
      if (obj.spaceBeforePt !== undefined) {
        options.spaceBeforePt = expectNonNegativeNumber(obj.spaceBeforePt, `${path}.spaceBeforePt`);
      }
      if (obj.spaceAfterPt !== undefined) {
        options.spaceAfterPt = expectNonNegativeNumber(obj.spaceAfterPt, `${path}.spaceAfterPt`);
      }
      return [buildParagraph(ctx, content, options, path)];
    }
    case "list": {
      rejectUnknownKeys(obj, ["type", "items", "style", "startAt"], path);
      const options: ListOptions = {};
      if (obj.style !== undefined) {
        options.style = expectEnum(obj.style, ["bullet", "number"] as const, `${path}.style`);
      }
      if (obj.startAt !== undefined) {
        options.startAt = expectIntInRange(obj.startAt, `${path}.startAt`, 1, 9999);
      }
      return buildListParagraphs(ctx, obj.items, options, path);
    }
    case "table": {
      return [buildTable(ctx, withoutType(obj), path)];
    }
    case "image": {
      return [buildImageBlock(ctx, withoutType(obj), path)];
    }
    case "pageBreak": {
      rejectUnknownKeys(obj, ["type"], path);
      return [buildPageBreak()];
    }
    case "toc": {
      return [buildTocBlock(withoutType(obj), path)];
    }
    default:
      // Unreachable: `allowedTypes.includes(type)` above already narrowed
      // to a known type — kept as a defensive exhaustiveness guard.
      throw new DocumentGenerationError(`${path}.type: unknown block type "${type}"`);
  }
}

export function buildBlocksFromSpecs(
  ctx: BuildContext,
  specs: unknown,
  allowedTypes: readonly string[],
  path: string,
): Block[] {
  const arr = expectArray(specs, path);
  return arr.flatMap((spec, i) => buildBlockFromSpec(ctx, spec, allowedTypes, `${path}[${i}]`));
}

/** A header/footer's content from a block-spec array, or from the
 * paragraph-content shorthand (a string or run array) for the common
 * single-line case. */
export function buildHeaderFooterContent(ctx: BuildContext, content: unknown, path: string): HeaderFooterContent {
  if (typeof content === "string") {
    return { blocks: [buildParagraph(ctx, content, {}, path)] };
  }
  if (Array.isArray(content) && content.every((item) => typeof item !== "object" || item === null || !("type" in (item as object)))) {
    return { blocks: [buildParagraph(ctx, content, {}, path)] };
  }
  return { blocks: buildBlocksFromSpecs(ctx, content, HEADER_FOOTER_BLOCK_TYPES, path) };
}

/** Fills in every TOC block's `entries` from the document's current
 * headings (SPEC-045 AC-1), in place. Run once, after the whole document —
 * every section's body and every header/footer — has been built: a TOC can
 * legitimately be inserted before headings that appear later in the body
 * (a report's front-page TOC), and lib/toc.ts's `tocEntriesFromHeadings`
 * always scans the full document regardless of where the TOC block sits. */
export function finalizeTocEntries(doc: DocumentModel): void {
  for (const section of doc.sections) {
    section.blocks = section.blocks.map((block) =>
      isToc(block) ? { ...block, entries: tocEntriesFromHeadings(doc, block.options) } : block,
    );
  }
}
