// Shared SPEC-008 plumbing for the two writers: base64 codecs for opaque
// part bytes (the model must stay JSON-serialisable, so parts are not held
// as raw Uint8Arrays), the widened writer result type, and the single place
// the cross-format "dropped content" report is derived from the model. The
// writer performing a conversion is the only component that knows both what
// an item is and that it cannot be represented in the target format, so the
// conversion report dialog and the tests both read the writer's `dropped`
// array rather than diffing models.
import {
  createDefaultStyles,
  isBreak,
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  mergeRunFormatting,
  tableLookIsEmpty,
  type Connector,
  type DocumentFormat,
  type DocumentModel,
  type Paragraph,
  type Run,
  type RunFormatting,
  type Shape,
  type Table,
} from "./documentModel";
import { parseXml, XmlParseError } from "./core/xml";
import type { BorderSide } from "./core/borders";
import { PATTERN_PRESET_TO_ODF_HATCH } from "./core/drawingMl";
// The chart-feature marker tables moved to core/chartFeatures.ts (SPEC-092)
// so the presentation rewrite-notice shares them; re-imported here unchanged
// so collectEditedChartDropRecords and every chartFeature expectation holds.
import { DOCX_CHART_FEATURE_MARKERS, ODT_CHART_FEATURE_MARKERS } from "./core/chartFeatures";
import { shapeDefinition } from "./core/shapeCatalogue";
import {
  diffParagraphFormatting,
  paragraphStyleBase,
  resolveCharacterStyle,
  splitRunFormatting,
} from "./styles";
import { paragraphHasTrackedChange } from "./renderFormatting";
import { equalColumnWidthsCm } from "./tables";
import { styleHasBanding, tableStyleFor } from "./tableStyles";

// --- base64 ----------------------------------------------------------------
//
// The codecs moved to core/packageParts.ts (they are package plumbing every
// product shares); re-exported here because the word-processor layer has
// always imported them from this module.
export { base64ToBytes, bytesToBase64, bytesToBase64Async } from "./core/packageParts";
import { base64ToBytes, decodedPartBytes } from "./core/packageParts";

/** The model's opaque parts as name -> decoded bytes, in model order. */
export function opaquePartBytes(model: DocumentModel): Map<string, Uint8Array> {
  return decodedPartBytes(model.opaque);
}

// --- writer result ---------------------------------------------------------

export type DropKind =
  | "chart"
  | "contentControl"
  | "customXml"
  | "embeddedObject"
  | "macro"
  | "opaqueBlock"
  | "opaqueInline"
  | "opaquePart"
  | "runShading"
  | "svgImage"
  // SPEC-054: a bar tab stop, which OpenDocument Text has no equivalent for —
  // a .docx → .odt save drops each one and reports it (AC-38).
  | "tabStop"
  | "trackedPropertyChange"
  // SPEC-051: a shape geometry outside the catalogue that a cross-format
  // save cannot carry into the target format (AC-61).
  | "shapeGeometry"
  // SPEC-073: a shape styling property one format cannot express — an
  // unmodelled preserved property crossing formats, or a modelled one the
  // target has no vocabulary for (AC-34).
  | "shapeStyle"
  // SPEC-052: a property of an *edited* chart the model has no field for —
  // the one case a same-format save can report a non-empty `dropped` for
  // (see "Rewriting a chart" — an untouched chart still reports nothing).
  | "chartFeature"
  // SPEC-050: content-level kinds a Markdown save can lose — Markdown has no
  // package, so unlike the kinds above these are detected by walking the
  // model's content rather than its unmodelled parts/nodes.
  | "pageSetup"
  | "headerFooter"
  | "pageNumberField"
  | "columns"
  | "note"
  | "comment"
  | "tableOfContents"
  | "tableOfFigures"
  | "bookmark"
  | "runFormatting"
  | "paragraphFormatting"
  | "tableStyling"
  | "trackedChange";

/** One item a writer could not represent in its target format. */
export interface DropRecord {
  kind: DropKind;
  /** Human-readable, e.g. "paragraph 3" or 'package part "customXml/item1.xml"'. */
  location: string;
  /** Why the target format cannot carry the item. */
  reason: string;
}

/** What both writers return (SPEC-008): the package bytes plus every item
 * that could not be represented in the target format. `dropped` is empty
 * for a same-format save — opaque content round-trips untouched. */
export interface WriteResult {
  bytes: Uint8Array;
  dropped: DropRecord[];
}

/** The initial (hidden) state of the conversion-report dialog. A typed
 * factory rather than an inline `as` cast because .vue scripts are parsed
 * by espree, which cannot parse TypeScript-only syntax. */
export function noConversionReport(): DropRecord[] | null {
  return null;
}

/** Display names for the conversion report dialog. */
export const DROP_KIND_LABELS: Record<DropKind, string> = {
  chart: "Chart",
  contentControl: "Content control",
  customXml: "Custom XML part",
  embeddedObject: "Embedded object",
  macro: "Macro",
  opaqueBlock: "Unsupported block content",
  opaqueInline: "Unsupported inline content",
  opaquePart: "Package part",
  runShading: "Run shading",
  svgImage: "SVG image",
  tabStop: "Bar tab stop",
  shapeGeometry: "Shape geometry",
  shapeStyle: "Shape styling",
  chartFeature: "Chart feature",
  trackedPropertyChange: "Tracked table property change",
  pageSetup: "Page setup",
  headerFooter: "Header/footer",
  pageNumberField: "Page number field",
  columns: "Multi-column layout",
  note: "Footnote/endnote",
  comment: "Comment",
  tableOfContents: "Table of contents field",
  tableOfFigures: "Table of figures field",
  bookmark: "Bookmark",
  runFormatting: "Character formatting",
  paragraphFormatting: "Paragraph formatting",
  tableStyling: "Table styling",
  trackedChange: "Tracked change",
};

/** DROP_KIND_LABELS behind a plain-string lookup, for the dialog template
 * (espree-parsed .vue scripts cannot use TS-only index narrowing). */
export function dropKindLabel(kind: string): string {
  return (DROP_KIND_LABELS as Record<string, string>)[kind] ?? "Unsupported content";
}

/** Typed pass-through for the dialog's runtime `Array` prop — .vue scripts
 * are parsed by espree, so the generic `PropType<DropRecord[]>` form is not
 * available there. */
export function asDropRecords(items: unknown): DropRecord[] {
  return Array.isArray(items) ? (items as DropRecord[]) : [];
}

// --- drop report -----------------------------------------------------------

const CHART_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** Package parts that are format plumbing rather than user content: the
 * target writer emits its own equivalents (packaging metadata, style and
 * settings tables), so listing them as "dropped" would only cry wolf. */
function isInfrastructurePart(source: DocumentFormat, name: string): boolean {
  if (source === "docx") {
    return (
      name === "[Content_Types].xml" ||
      name.includes("_rels/") ||
      name.startsWith("docProps/") ||
      name === "word/styles.xml" ||
      name === "word/fontTable.xml" ||
      name === "word/settings.xml" ||
      name === "word/webSettings.xml" ||
      name.startsWith("word/theme/")
    );
  }
  return (
    name === "mimetype" ||
    name.startsWith("META-INF/") ||
    name === "styles.xml" ||
    name === "settings.xml" ||
    name === "meta.xml" ||
    name === "manifest.rdf" ||
    name === "layout-cache" ||
    name.startsWith("Configurations2/") ||
    name.startsWith("Thumbnails/")
  );
}

function partKind(source: DocumentFormat, name: string): DropKind {
  if (source === "docx") {
    if (name.startsWith("customXml/")) {
      return "customXml";
    }
    if (name.startsWith("word/charts/")) {
      return "chart";
    }
    if (name.startsWith("word/embeddings/")) {
      return "embeddedObject";
    }
    if (name.startsWith("word/vbaProject")) {
      return "macro";
    }
  } else if (name.startsWith("ObjectReplacements/") || /^Object ?\d+\//.test(name)) {
    return "embeddedObject";
  }
  return "opaquePart";
}

/** Classifies an opaque node's XML by its content: block/inline `w:sdt` is a
 * content control, a `w:drawing` holding a DrawingML chart reference is a
 * chart. Anything unparseable or unrecognised falls back to the generic
 * kind. */
function nodeKind(xml: string, fallback: DropKind): DropKind {
  let doc: Document;
  try {
    doc = parseXml(xml);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return fallback;
    }
    throw error;
  }
  const root = doc.documentElement;
  if (root.localName === "sdt") {
    return "contentControl";
  }
  if (root.getElementsByTagNameNS(CHART_NAMESPACE, "*").length > 0) {
    return "chart";
  }
  return fallback;
}

const REASONS: Record<DropKind, string> = {
  chart: "charts cannot be converted between formats",
  contentControl: "content controls have no equivalent in the target format",
  customXml: "custom XML parts have no equivalent in the target format",
  embeddedObject: "embedded objects cannot be converted between formats",
  macro: "macros have no equivalent in the target format",
  opaqueBlock: "this content is not understood by the editor and cannot be converted",
  opaqueInline: "this content is not understood by the editor and cannot be converted",
  opaquePart: "this file has no equivalent in the target format",
  runShading:
    "the target format has a single text background, which the highlight colour occupies",
  svgImage: "Word (.docx) cannot store SVG images and this one could not be converted to PNG",
  tabStop: "OpenDocument Text has no bar tab stop — the vertical rule it draws is lost",
  shapeGeometry: "this shape's geometry is outside the catalogue and cannot be converted between formats",
  // Overridden per call by shapeStyleDropReason with the specific property;
  // this generic text is the fallback.
  shapeStyle: "this shape styling property cannot be expressed in the target format",
  // Overridden per call by collectEditedChartDropRecords with the specific
  // lost property; this generic text is the fallback record() would use if
  // anything ever routes a chartFeature drop through it instead.
  chartFeature: "this property of the original chart is not carried by the editor's chart model",
  trackedPropertyChange:
    "OpenDocument Text has no slot for a table/row/cell property change's previous state — the change is kept applied but no longer tracked",
  pageSetup: "Markdown carries no page geometry or section breaks",
  headerFooter: "Markdown has no header/footer concept",
  pageNumberField: "Markdown has no live page-number field",
  columns: "Markdown has no multi-column layout",
  note: "Markdown has no footnote/endnote mechanism",
  comment: "Markdown has no review-comment mechanism",
  tableOfContents: "Markdown cannot carry a live table of contents field — its rendered text is kept instead",
  tableOfFigures: "Markdown cannot carry a live table of figures field — its rendered text is kept instead",
  bookmark: "Markdown has no bookmark mechanism",
  runFormatting: "Markdown carries only bold, italic, strikethrough and code — other character formatting is lost",
  paragraphFormatting: "Markdown has no alignment, indent or spacing properties",
  tableStyling: "Markdown tables cannot merge cells or carry table styling",
  trackedChange: "Markdown cannot track pending changes — the change is kept applied but no longer tracked",
};

const FORMAT_LABELS: Record<DocumentFormat, string> = {
  docx: "Word (.docx)",
  odt: "OpenDocument Text (.odt)",
  md: "Markdown (.md)",
};

function record(kind: DropKind, location: string, target: DocumentFormat): DropRecord {
  return {
    kind,
    location,
    reason: `Not preserved when saving as ${FORMAT_LABELS[target]}: ${REASONS[kind]}`,
  };
}

/** The `.docx`/`.odt` package-level drops (opaque nodes/parts, SVG images,
 * run shading, tracked table-property changes): everything a cross-format
 * save between the two package formats will drop, in document order (opaque
 * nodes first, then unmodelled package parts). Empty when the model carries
 * no opaque content — a blank document converts losslessly between them.
 * The "approximate location" is expressed in paragraph numbers: the product
 * does not paginate, so page numbers do not exist to report. Unlike
 * `collectDropRecords`, this is only ever meaningful for a genuine format
 * change — a same-format save round-trips opaque content untouched. */
/** Every `sourcePart` a live, catalogue-type Chart block still references —
 * its part-level opaque entry converts through the block-level path (a
 * fresh part regenerated from the model in the target format) rather than
 * being genuinely dropped, so the generic "every remaining opaque part is a
 * drop" sweep below must not double-report it. An out-of-catalogue chart's
 * source part is deliberately excluded from this set: that one really is
 * dropped, and is already reported once via the isChart(blockAt) branch
 * above — excluding it here too would just make the sweep silently skip a
 * genuine drop instead of reporting it twice. */
function liveConvertibleChartSourceParts(model: DocumentModel): Set<string> {
  const parts = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isChart(block) && block.sourcePart && block.unsupportedChartXml === undefined) {
        parts.add(block.sourcePart);
      }
    }
  }
  return parts;
}

/** SPEC-073 AC-34: the styling properties of a shape/connector that a
 * cross-format save into `target` cannot express, each named so the
 * conversion report lists the specific loss rather than a silent flatten.
 * Preserved unmodelled styling (`unsupportedStyleXml`) is verbatim source
 * XML this module never translates, so it drops on any format change;
 * modelled properties drop only where the target lacks the vocabulary —
 * ODF has no compound line, no gradient past two stops, no per-stop
 * transparency and only the closed hatch table of pattern presets. */
function lostShapeStyleProperties(block: Shape | Connector, target: DocumentFormat): string[] {
  const lost: string[] = [];
  if (block.unsupportedStyleXml) {
    lost.push("unmodelled styling from the source file");
  }
  if (target === "odt") {
    if (block.outline?.compound && block.outline.compound !== "single") {
      lost.push("the outline's compound (multi-line) style");
    }
    if (isShape(block)) {
      const fill = block.fill;
      if (fill?.kind === "gradient") {
        if (fill.gradient.stops.length > 2) {
          lost.push("the gradient's intermediate stops");
        }
        if (fill.gradient.stops.some((stop) => stop.alpha !== undefined && stop.alpha < 1)) {
          lost.push("the gradient's per-stop transparency");
        }
      }
      if (fill?.kind === "pattern" && !PATTERN_PRESET_TO_ODF_HATCH[fill.preset]) {
        lost.push(`the "${fill.preset}" pattern preset`);
      }
    }
  }
  return lost;
}

function shapeStyleDrop(property: string, location: string, target: DocumentFormat): DropRecord {
  return {
    kind: "shapeStyle",
    location,
    reason: `Not preserved when saving as ${FORMAT_LABELS[target]}: the target format cannot express ${property}`,
  };
}

function collectPackageDropRecords(model: DocumentModel, source: DocumentFormat, target: DocumentFormat): DropRecord[] {
  const dropped: DropRecord[] = [];
  const convertibleChartParts = liveConvertibleChartSourceParts(model);
  let paragraphOffset = 0;
  for (const section of model.sections) {
    const opaqueBlocks = [...(section.opaqueBlocks ?? [])].sort((a, b) => a.blockIndex - b.blockIndex);
    for (let index = 0; index <= section.blocks.length; index++) {
      for (const block of opaqueBlocks) {
        if (Math.min(block.blockIndex, section.blocks.length) === index) {
          const location =
            index < section.blocks.length
              ? `before paragraph ${paragraphOffset + index + 1}`
              : "end of document";
          dropped.push(record(nodeKind(block.xml, "opaqueBlock"), location, target));
        }
      }
      if (index < section.blocks.length) {
        const paragraphNumber = paragraphOffset + index + 1;
        const blockAt = section.blocks[index];
        if (isImage(blockAt)) {
          // Images convert losslessly between formats (both carry media
          // parts) except SVG, which .docx cannot store at all (SPEC-024
          // AC-2 converts it to PNG at insert time — the only way one still
          // reaches a .docx save is a cross-format Save As of a .odt
          // document that has an SVG image, since raster conversion here
          // would need a browser canvas this pure module cannot assume).
          const media = (model.media ?? []).find((part) => part.mediaId === blockAt.mediaId);
          if (target === "docx" && media?.contentType === "image/svg+xml") {
            dropped.push(record("svgImage", `paragraph ${paragraphNumber}`, target));
          }
          continue;
        }
        if (isBreak(blockAt)) {
          // A break block carries no content and no unmodelled XML.
          continue;
        }
        if (isToc(blockAt) || isTof(blockAt)) {
          // A TOC/ToF's generated entries are plain data, not preserved XML —
          // both writers regenerate its field/index markup fresh, so there
          // is nothing here that could be dropped by a format conversion.
          continue;
        }
        if (isShape(blockAt)) {
          if (source !== target) {
            for (const property of lostShapeStyleProperties(blockAt, target)) {
              dropped.push(shapeStyleDrop(property, `paragraph ${paragraphNumber}`, target));
            }
          }
          // An unsupported geometry (SPEC-008 preservation) is verbatim XML
          // this module never translates between formats — any cross-format
          // save of it drops the geometry (SPEC-051 AC-61); a same-format
          // save re-emits it unchanged and reports nothing.
          if (blockAt.unsupportedGeometryXml && source !== target) {
            dropped.push(record("shapeGeometry", `paragraph ${paragraphNumber}`, target));
          } else if (blockAt.path && blockAt.pathSource && blockAt.pathSource.format !== target) {
            // A path reduced from arcs/quadratics on read (SPEC-071 AC-30):
            // a same-format save re-emits the preserved source (no loss), but
            // a cross-format save writes the converted cubics — a real, small
            // fidelity loss we report rather than hide.
            dropped.push(record("shapeGeometry", `paragraph ${paragraphNumber}`, target));
          } else if (source !== target) {
            // A catalogue geometry the target format has no clean equivalent
            // for: the writer emits the closest preset AND we report the drop
            // (SPEC-069 AC-9 / "Cross-format mapping").
            const def = shapeDefinition(blockAt.geometry);
            if (def?.approximateIn?.some((format) => format === target)) {
              dropped.push(record("shapeGeometry", `paragraph ${paragraphNumber}`, target));
            }
          }
          continue;
        }
        if (isChart(blockAt)) {
          // A chart whose type is outside the catalogue (SPEC-008
          // preservation) is verbatim XML this module never translates
          // between formats — a cross-format save of it drops the chart
          // (SPEC-052 AC "unsupported type... never rewritten"); a
          // catalogue chart type converts losslessly (its model IS the
          // source of truth for both formats), so nothing is reported for
          // it here. A same-format save reports nothing either way — see
          // collectEditedChartDropRecords for the "edited" exception.
          if (blockAt.unsupportedChartXml && source !== target) {
            dropped.push(record("chart", `paragraph ${paragraphNumber}`, target));
          }
          continue;
        }
        if (isConnector(blockAt)) {
          if (source !== target) {
            for (const property of lostShapeStyleProperties(blockAt, target)) {
              dropped.push(shapeStyleDrop(property, `paragraph ${paragraphNumber}`, target));
            }
          }
          // SPEC-070 AC-16/AC-32: a connector converts cleanly between formats
          // (its own preset + attachment refs). The one lossy case is an
          // endpoint attached to a *non-shape* drawing saved to `.odt`, whose
          // `draw:start-shape`/`draw:end-shape` needs a `draw:id` the ODF
          // writer only mints for shapes — that attachment drops to
          // free-floating and is reported. `.docx` (`wp:docPr/@id`) can name
          // any drawing target, so a cross-format `.odt → .docx` reports
          // nothing.
          if (source !== target && target === "odt") {
            for (const attachment of [blockAt.startAttachment, blockAt.endAttachment]) {
              if (!attachment) {
                continue;
              }
              const targetBlock = section.blocks[attachment.target.block];
              if (targetBlock && !isShape(targetBlock)) {
                dropped.push(record("shapeGeometry", `paragraph ${paragraphNumber}`, target));
              }
            }
          }
          continue;
        }
        if (target === "odt" && isTable(blockAt)) {
          // SPEC-039 AC-19: a table/row/cell property change (a merge/split,
          // a border/shading/padding/height/width/alignment edit) has no
          // .odt representation at all — unlike a paragraph mark or a table
          // row/cell insert/delete, which the .odt writer does track
          // (AC-16). The change is still applied (its current value is what
          // gets written either way), just no longer tracked as pending —
          // reported here rather than silently dropped.
          if (blockAt.changes?.propertyChange) {
            dropped.push(record("trackedPropertyChange", `table at paragraph ${paragraphNumber}`, target));
          }
          blockAt.rows.forEach((row, rowIndex) => {
            if (row.changes?.propertyChange) {
              dropped.push(
                record("trackedPropertyChange", `table at paragraph ${paragraphNumber}, row ${rowIndex + 1}`, target),
              );
            }
            row.cells.forEach((cell, colIndex) => {
              if (cell.changes?.propertyChange) {
                dropped.push(
                  record(
                    "trackedPropertyChange",
                    `table at paragraph ${paragraphNumber}, row ${rowIndex + 1}, column ${colIndex + 1}`,
                    target,
                  ),
                );
              }
            });
          });
        }
        // A table's cell paragraphs are inspected too (SPEC-022); their
        // records locate by the table's own block position.
        const paragraphs =
          blockAt.type === "table"
            ? blockAt.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
            : [blockAt];
        for (const paragraph of paragraphs) {
          for (const inline of paragraph.opaqueInlines ?? []) {
            dropped.push(record(nodeKind(inline.xml, "opaqueInline"), `paragraph ${paragraphNumber}`, target));
          }
          // SPEC-056 Format mapping: an equation carrying an `opaqueMath` node
          // — a construct outside the closed set, preserved verbatim in its
          // *source* format's markup (OMML or MathML) — has no expression in
          // the other format, so a cross-format save drops it and reports it
          // (AC-32). A fully-modelled equation converts losslessly and reports
          // nothing.
          if (paragraph.runs.some((run) => run.equation?.body.some((node) => node.kind === "opaqueMath"))) {
            dropped.push(record("opaqueInline", `equation at paragraph ${paragraphNumber}`, target));
          }
          // Run shading preserved beside a highlight (SPEC-013): ODF has one
          // background slot, which the highlight occupies, so a .docx → .odt
          // save loses the shading. One record per paragraph, not per run —
          // the report locates by paragraph anyway.
          if (
            target === "odt" &&
            paragraph.runs.some((run) => run.formatting.shadingFill !== undefined)
          ) {
            dropped.push(record("runShading", `paragraph ${paragraphNumber}`, target));
          }
        }
      }
    }
    paragraphOffset += section.blocks.length;
  }

  for (const part of model.opaque ?? []) {
    if (!isInfrastructurePart(source, part.name) && !convertibleChartParts.has(part.name)) {
      dropped.push(record(partKind(source, part.name), `package part "${part.name}"`, target));
    }
  }

  return dropped;
}

// --- SPEC-050: Markdown content-level drops ---------------------------------

/** The only run-formatting a Markdown save carries (the "Supported subset"
 * table): everything else — including anything a paragraph/character style
 * itself sets — must be flagged, but *only* the direct layer beyond a run's
 * style chain, so a heading's baked bold or a Quote paragraph's baked italic
 * (both style-driven, both representable via the construct itself) are not
 * mistaken for unsupported direct formatting. */
const ALLOWED_RUN_DIFF_KEYS = new Set(["bold", "italic", "strikethrough"]);

/** A run's formatting resolved from its paragraph's and its own style chain
 * alone (no direct formatting) — what `checkParagraph`/`withoutDropped` both
 * diff a run's actual formatting against. */
function runStyleBase(model: DocumentModel, paragraph: Paragraph, run: Run): RunFormatting {
  const paragraphBase = paragraphStyleBase(model, paragraph).runFormatting;
  return run.styleId ? mergeRunFormatting(paragraphBase, resolveCharacterStyle(model, run.styleId)) : paragraphBase;
}

/** Whether a table has merged cells or per-cell/row/table styling GFM's
 * plain pipe-table grid cannot carry (SPEC-050's "Supported subset": "Table
 * without merged cells"). */
function tableHasUnsupportedStyling(model: DocumentModel, table: Table): boolean {
  const hasMerge = table.rows.some((row) =>
    row.cells.some((cell) => (cell.colSpan ?? 1) > 1 || (cell.rowSpan ?? 1) > 1 || cell.covered),
  );
  const hasCellOrRowStyling = table.rows.some(
    (row, rowIndex) =>
      row.minHeightCm !== undefined ||
      // A header row is only representable as GFM's implicit first row
      // (the writer always treats row 0 as the header) — a header
      // anywhere else is what's actually unsupported.
      (row.isHeader && rowIndex !== 0) ||
      row.cells.some(
        (cell) =>
          cell.borders !== undefined ||
          cell.shadingFill !== undefined ||
          cell.paddingCm !== undefined ||
          cell.verticalAlign !== undefined,
      ),
  );
  const hasTableStyling =
    (table.alignment !== undefined && table.alignment !== "left") ||
    (table.widthMode !== undefined && table.widthMode !== "fixed") ||
    table.widthPercent !== undefined ||
    table.styleId !== undefined ||
    // SPEC-076 AC-74: a look flag, or a style carrying banding, is table
    // styling GFM has no way to carry.
    !tableLookIsEmpty(table.look) ||
    styleHasBanding(tableStyleFor(model.styles, table));
  // GFM tables carry no column-width info at all — a Markdown save always
  // reflows to an equal split of the text area, so any table whose widths
  // are not *already* that equal split loses them for real (a table
  // Markdown itself produced round-trips clean, since it starts equal).
  const equalWidths = equalColumnWidthsCm(model, table.columnWidthsCm.length);
  const hasCustomWidths = table.columnWidthsCm.some((width, index) => Math.abs(width - equalWidths[index]) >= 0.01);
  return hasMerge || hasCellOrRowStyling || hasTableStyling || hasCustomWidths;
}

/** Whether a table (or one of its rows/cells) carries pending tracked-change
 * state (SPEC-039) — Markdown has nowhere to keep it. */
function tableHasTrackedChanges(table: Table): boolean {
  if (table.changes?.propertyChange) {
    return true;
  }
  return table.rows.some(
    (row) =>
      row.changes?.insert ||
      row.changes?.delete ||
      row.changes?.propertyChange ||
      row.cells.some((cell) => cell.changes?.insert || cell.changes?.delete || cell.changes?.propertyChange),
  );
}

/** Everything a save to Markdown will drop that a `.docx`/`.odt` save never
 * would (SPEC-050's "What markdown cannot carry"): unlike
 * `collectPackageDropRecords`, this runs for *every* save to `.md`,
 * including one whose source is already `.md` — Markdown is not a
 * full-fidelity format the way `.docx`/`.odt` are, so content the editor
 * itself can create (a header, a tracked change) can be unrepresentable even
 * when re-saving a document that was already Markdown. */
function collectMarkdownDropRecords(model: DocumentModel): DropRecord[] {
  const target: DocumentFormat = "md";
  const dropped: DropRecord[] = [];
  const seenThreads = new Set<string>();
  const seenNotes = new Set<string>();

  function checkParagraph(paragraph: Paragraph, paragraphNumber: number): void {
    if (paragraphHasTrackedChange(paragraph)) {
      dropped.push(record("trackedChange", `paragraph ${paragraphNumber}`, target));
    }
    const paragraphBase = paragraphStyleBase(model, paragraph).paragraphFormatting;
    if (Object.keys(diffParagraphFormatting(paragraph.formatting ?? {}, paragraphBase)).length > 0) {
      dropped.push(record("paragraphFormatting", `paragraph ${paragraphNumber}`, target));
    }
    const hasUnsupportedRunFormatting = paragraph.runs.some((run) => {
      const base = runStyleBase(model, paragraph, run);
      // splitRunFormatting (not a raw diff against `base`) so a run's font
      // family/size baked in only because it happens to match the
      // document's own defaults — the "underlay" every reader bakes onto
      // every run regardless of any real override — is not mistaken for a
      // direct override Markdown would have to drop.
      const { diff } = splitRunFormatting(model, run.formatting, base);
      return (
        Object.keys(diff).some((key) => !ALLOWED_RUN_DIFF_KEYS.has(key)) ||
        run.formatting.lang !== undefined ||
        run.formatting.noProof === true
      );
    });
    if (hasUnsupportedRunFormatting) {
      dropped.push(record("runFormatting", `paragraph ${paragraphNumber}`, target));
    }
    if (paragraph.runs.some((run) => run.field)) {
      dropped.push(record("pageNumberField", `paragraph ${paragraphNumber}`, target));
    }
    for (const _bookmark of paragraph.bookmarks ?? []) {
      dropped.push(record("bookmark", `paragraph ${paragraphNumber}`, target));
    }
    for (const anchor of paragraph.commentAnchors ?? []) {
      if (!seenThreads.has(anchor.threadId)) {
        seenThreads.add(anchor.threadId);
        dropped.push(record("comment", `paragraph ${paragraphNumber}`, target));
      }
    }
    for (const run of paragraph.runs) {
      if (run.note && !seenNotes.has(run.note.noteId)) {
        seenNotes.add(run.note.noteId);
        dropped.push(record("note", `paragraph ${paragraphNumber}`, target));
      }
    }
  }

  // A single-section document's own page geometry lives on `defaultStyles`
  // directly, not `sections[0].pageSize`/`margins` (SPEC-028's "absent
  // falls back to defaultStyles" — the docx/odt readers never populate a
  // lone section's own override when the document has only one). A tenth
  // of a millimetre of slack absorbs the twip-grid rounding a value that
  // started life as a clean cm figure picks up on its way through .docx.
  const EPSILON_CM = 0.01;
  const nearlyEqual = (a: number, b: number): boolean => Math.abs(a - b) < EPSILON_CM;
  const standard = createDefaultStyles();
  const hasCustomDocumentPageSetup =
    !nearlyEqual(model.defaultStyles.pageSize.widthCm, standard.pageSize.widthCm) ||
    !nearlyEqual(model.defaultStyles.pageSize.heightCm, standard.pageSize.heightCm) ||
    !nearlyEqual(model.defaultStyles.margins.topCm, standard.margins.topCm) ||
    !nearlyEqual(model.defaultStyles.margins.rightCm, standard.margins.rightCm) ||
    !nearlyEqual(model.defaultStyles.margins.bottomCm, standard.margins.bottomCm) ||
    !nearlyEqual(model.defaultStyles.margins.leftCm, standard.margins.leftCm);
  if (hasCustomDocumentPageSetup) {
    dropped.push(record("pageSetup", "document page setup", target));
  }

  model.sections.forEach((section, sectionIndex) => {
    if (
      sectionIndex > 0 ||
      section.pageSize !== undefined ||
      section.margins !== undefined ||
      section.docxSectPr !== undefined ||
      section.pageNumberStart !== undefined ||
      section.titlePage ||
      section.evenAndOddHeaders
    ) {
      dropped.push(record("pageSetup", `section ${sectionIndex + 1}`, target));
    }
    if (section.columns && section.columns.count > 1) {
      dropped.push(record("columns", `section ${sectionIndex + 1}`, target));
    }
    const hasHeader = section.headers && Object.keys(section.headers).length > 0;
    const hasFooter = section.footers && Object.keys(section.footers).length > 0;
    if (hasHeader || hasFooter) {
      dropped.push(record("headerFooter", `section ${sectionIndex + 1}`, target));
    }
  });

  let paragraphOffset = 0;
  for (const section of model.sections) {
    section.blocks.forEach((block, index) => {
      const paragraphNumber = paragraphOffset + index + 1;
      if (isToc(block)) {
        dropped.push(record("tableOfContents", `paragraph ${paragraphNumber}`, target));
        return;
      }
      if (isTof(block)) {
        dropped.push(record("tableOfFigures", `paragraph ${paragraphNumber}`, target));
        return;
      }
      if (isTable(block)) {
        if (tableHasUnsupportedStyling(model, block)) {
          dropped.push(record("tableStyling", `table at paragraph ${paragraphNumber}`, target));
        }
        if (tableHasTrackedChanges(block)) {
          dropped.push(record("trackedChange", `table at paragraph ${paragraphNumber}`, target));
        }
        for (const row of block.rows) {
          for (const cell of row.cells) {
            for (const paragraph of cell.blocks) {
              checkParagraph(paragraph, paragraphNumber);
            }
          }
        }
        return;
      }
      if (isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block)) {
        return;
      }
      checkParagraph(block, paragraphNumber);
    });
    paragraphOffset += section.blocks.length;
  }

  return dropped;
}

/** Everything a cross-format save of `model` will drop. Two independent
 * passes, both able to fire on the same save: the package-level pass (any
 * genuine format change — its `.docx`/`.odt`-specific checks, e.g.
 * `svgImage`, stay inert when `target` is `"md"` since they gate on that
 * exact target format) and the Markdown content-level pass (every save to
 * `.md`, regardless of source — see `collectMarkdownDropRecords`, which
 * covers what the package-level pass cannot: content a `.md` *source* never
 * has opaque parts/nodes for in the first place). */
export function collectDropRecords(model: DocumentModel, target: DocumentFormat): DropRecord[] {
  const source = model.format;
  const dropped: DropRecord[] = [];
  if (source !== undefined && source !== target) {
    dropped.push(...collectPackageDropRecords(model, source, target));
  }
  if (target === "odt") {
    // Bar tab stops (SPEC-054 AC-38): OpenDocument Text has no bar tab, so
    // ANY save to .odt drops every one — regardless of source, unlike the
    // package-level opaque content above (a bar stop is a modelled feature
    // ODF cannot express, not preserved-but-untranslated XML). A brand-new
    // document has no source format, so this cannot ride on the cross-format
    // gate above.
    dropped.push(...collectBarTabDropRecords(model, target));
    // Border styles ODF cannot express (SPEC-074 AC-32) — same reasoning as
    // the bar tabs above: a modelled feature the target format lacks, so
    // every save to .odt reports it regardless of source.
    dropped.push(...collectBorderStyleDropRecords(model));
    // SPEC-076 AC-14, AC-70: header rows and band sizes ODF cannot carry.
    dropped.push(...collectOdfTableDropRecords(model));
  }
  if (target === "docx") {
    dropped.push(...collectPageBackgroundDropRecords(model));
  }
  if (target === "md") {
    dropped.push(...collectMarkdownDropRecords(model));
  }
  return dropped;
}

/** The border styles ODF has no keyword for (SPEC-074 AC-32): `dashDot` and
 * `dashDotDot` write as `dashed`, their nearest equivalent, and each
 * paragraph (or section page border) that uses one is reported. Exported so
 * a same-format `.odt` save reports them too — ODF's vocabulary is the
 * limit, not the conversion. */
export function collectBorderStyleDropRecords(model: DocumentModel): DropRecord[] {
  const dropped: DropRecord[] = [];
  const lost = (side: BorderSide | undefined): boolean =>
    side !== undefined && (side.style === "dashDot" || side.style === "dashDotDot");
  const reason =
    "Not preserved when saving as OpenDocument Text: ODF's border keywords have no dash-dot " +
    "pattern, so the border is written as its nearest equivalent, dashed.";
  let paragraphOffset = 0;
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, index) => {
      const paragraphNumber = paragraphOffset + index + 1;
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : [block];
      for (const paragraph of paragraphs) {
        if (paragraph.type !== "paragraph") {
          continue;
        }
        const borders = paragraph.formatting?.borders;
        if (
          borders &&
          (lost(borders.top) || lost(borders.bottom) || lost(borders.left) || lost(borders.right) || lost(borders.between))
        ) {
          dropped.push({ kind: "paragraphFormatting", location: `paragraph ${paragraphNumber}`, reason });
        }
      }
    });
    paragraphOffset += section.blocks.length;
    const page = section.pageBorders;
    if (page && (lost(page.top) || lost(page.bottom) || lost(page.left) || lost(page.right))) {
      dropped.push({ kind: "pageSetup", location: `section ${sectionIndex + 1} page border`, reason });
    }
  });
  return dropped;
}

/** What an ODF save cannot carry from SPEC-076's table model, reported under
 * the existing "tableStyling" kind:
 *
 * - an `isHeader` row OUTSIDE the leading run (AC-14). ODF's
 *   `table:table-header-rows` is a structural wrapper, so only a leading run
 *   is representable; a later header row writes as a plain row.
 * - a band size other than 1 (AC-70). ODF's table template has odd/even row
 *   and column entries and no band-size attribute at all, so the banding
 *   writes as size-1 and the declared size is lost.
 *
 * Exported so a same-format `.odt` save reports them too: ODF's vocabulary is
 * the limit, not the conversion. */
export function collectOdfTableDropRecords(model: DocumentModel): DropRecord[] {
  const dropped: DropRecord[] = [];
  let paragraphOffset = 0;
  for (const section of model.sections) {
    section.blocks.forEach((block, index) => {
      if (block.type !== "table") {
        return;
      }
      const location = `table at paragraph ${paragraphOffset + index + 1}`;
      let leading = 0;
      while (leading < block.rows.length && block.rows[leading].isHeader) {
        leading += 1;
      }
      block.rows.forEach((row, rowIndex) => {
        if (row.isHeader && rowIndex >= leading) {
          dropped.push({
            kind: "tableStyling",
            location: `${location}, row ${rowIndex + 1}`,
            reason:
              "Not preserved when saving as OpenDocument Text: ODF wraps only a leading run of " +
              "header rows, so this row is written as an ordinary row.",
          });
        }
      });
    });
    paragraphOffset += section.blocks.length;
  }
  for (const style of model.styles ?? []) {
    for (const axis of ["bandedRows", "bandedColumns"] as const) {
      const band = style.banding?.[axis];
      if (band !== undefined && band.size !== 1) {
        dropped.push({
          kind: "tableStyling",
          location: `table style "${style.name ?? style.id}" (${axis === "bandedRows" ? "row" : "column"} band size ${band.size})`,
          reason:
            "Not preserved when saving as OpenDocument Text: ODF's table template bands alternate " +
            "row by row and column by column, so the band size is written as 1.",
        });
      }
    }
  }
  return dropped;
}

/** The page background a `.docx` save cannot carry (SPEC-074 AC-32): OOXML
 * stores one `w:background` for the whole document while the model holds it
 * per section, so a save writes the first section's value and reports the
 * rest. Exported for the same-format `.docx` save, which loses it just as a
 * converted one does. */
export function collectPageBackgroundDropRecords(model: DocumentModel): DropRecord[] {
  if (model.sections.length < 2) {
    return [];
  }
  const first = model.sections[0]?.pageBackground;
  const disagreeing = model.sections.filter((section) => section.pageBackground !== first);
  if (disagreeing.length === 0) {
    return [];
  }
  return [
    {
      kind: "pageSetup",
      location: `${disagreeing.length} section${disagreeing.length === 1 ? "" : "s"} after the first`,
      reason:
        "Not preserved when saving as Word (.docx): OOXML holds one page background for the whole " +
        "document, so the first section's colour is written and the others are lost.",
    },
  ];
}

/** One `tabStop` DropRecord per bar tab stop in the model (SPEC-054 AC-38),
 * located by paragraph — table cell paragraphs included, addressed by the
 * table's own block position, mirroring collectPackageDropRecords. Exported so
 * a same-format .odt save (which reports edited-chart drops rather than the
 * cross-format collectDropRecords) can still surface every dropped bar rule. */
export function collectBarTabDropRecords(model: DocumentModel, target: DocumentFormat): DropRecord[] {
  const dropped: DropRecord[] = [];
  let paragraphOffset = 0;
  for (const section of model.sections) {
    section.blocks.forEach((block, index) => {
      const paragraphNumber = paragraphOffset + index + 1;
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : [block];
      for (const paragraph of paragraphs) {
        if (paragraph.type !== "paragraph") {
          continue;
        }
        for (const stop of paragraph.formatting?.tabStops ?? []) {
          if (stop.alignment === "bar") {
            dropped.push(record("tabStop", `paragraph ${paragraphNumber}`, target));
          }
        }
      }
    });
    paragraphOffset += section.blocks.length;
  }
  return dropped;
}

// --- SPEC-052: chartFeature drops for a same-format save of an edited chart -

/** Everything an *edited* chart's regenerated part loses relative to its
 * preserved original — SPEC-052's one exception to "a same-format save
 * reports nothing" (see "Rewriting a chart"). Both writers' chart-plan
 * modules call this to build the `chartFeature` entries in their same-format
 * save's `dropped`; every other chart (untouched, or with no `sourcePart` to
 * compare against, e.g. one created here) contributes nothing. */
export function collectEditedChartDropRecords(model: DocumentModel, target: "docx" | "odt"): DropRecord[] {
  const dropped: DropRecord[] = [];
  const opaqueByName = new Map((model.opaque ?? []).map((part) => [part.name, part.data]));
  const markers = target === "docx" ? DOCX_CHART_FEATURE_MARKERS : ODT_CHART_FEATURE_MARKERS;
  let chartNumber = 0;
  let paragraphOffset = 0;
  for (const section of model.sections) {
    section.blocks.forEach((block, index) => {
      if (!isChart(block)) {
        return;
      }
      chartNumber += 1;
      if (!block.edited || !block.sourcePart || block.unsupportedChartXml !== undefined) {
        return;
      }
      const data = opaqueByName.get(block.sourcePart);
      if (data === undefined) {
        return;
      }
      const xml = new TextDecoder("utf-8").decode(base64ToBytes(data));
      for (const marker of markers) {
        if (marker.test(xml)) {
          dropped.push({
            kind: "chartFeature",
            location: `Chart ${chartNumber} (paragraph ${paragraphOffset + index + 1})`,
            reason: marker.reason,
          });
        }
      }
    });
    paragraphOffset += section.blocks.length;
  }
  return dropped;
}

// --- SPEC-050: withoutDropped ------------------------------------------------

/** Masks `run`'s formatting down to what a real save-to-`.md`-and-reopen
 * would leave: its style chain's own contribution, plus whichever of the
 * allowed direct keys (bold/italic/strikethrough) it actually set, plus its
 * link target (a Markdown-representable hyperlink is never dropped). */
function maskRunFormatting(model: DocumentModel, paragraph: Paragraph, run: Run): RunFormatting {
  const base = runStyleBase(model, paragraph, run);
  const { underlay, diff } = splitRunFormatting(model, run.formatting, base);
  const allowedDirect: RunFormatting = {};
  for (const key of ALLOWED_RUN_DIFF_KEYS) {
    const value = diff[key as keyof RunFormatting];
    if (value !== undefined) {
      (allowedDirect as Record<string, unknown>)[key] = value;
    }
  }
  const masked = mergeRunFormatting(mergeRunFormatting(base, underlay), allowedDirect);
  if (run.formatting.linkHref !== undefined) {
    masked.linkHref = run.formatting.linkHref;
  }
  if (run.formatting.linkAnchor !== undefined) {
    masked.linkAnchor = run.formatting.linkAnchor;
  }
  return masked;
}

/** Masks one paragraph in place to what a real save-to-`.md`-and-reopen
 * would leave: its style's own paragraph formatting (direct alignment/
 * indent/spacing dropped), its runs' formatting similarly masked (field/note
 * runs removed — Markdown has nowhere to keep a computed value), its
 * tracked-change state cleared (current content kept, per the same
 * "kept applied but no longer tracked" convention `trackedPropertyChange`
 * already uses), and its bookmarks/comment anchors removed. */
function maskParagraph(model: DocumentModel, paragraph: Paragraph): void {
  const base = paragraphStyleBase(model, paragraph).paragraphFormatting;
  paragraph.formatting = Object.keys(base).length > 0 ? base : undefined;
  paragraph.runs = paragraph.runs.filter((run) => !run.field && !run.note);
  for (const run of paragraph.runs) {
    run.formatting = maskRunFormatting(model, paragraph, run);
    delete run.changes;
  }
  delete paragraph.changes;
  delete paragraph.bookmarks;
  delete paragraph.commentAnchors;
}

/** `source` with every item named in `dropped` masked out — the transform
 * `equalSemantic(withoutDropped(source, dropped), reopened)` compares
 * against (SPEC-008's "Writer API"), so the source's own now-unrepresentable
 * content never causes a false mismatch. Kind-driven, not location-driven:
 * `dropped`'s free-text `location` is for the conversion report a person
 * reads, not for this transform to parse back apart. Masking every instance
 * of a *kind* present anywhere in `dropped` is sound here because each mask
 * step recomputes its target from the model's own style chain — a paragraph
 * that never had unsupported formatting masks back to exactly what it
 * already had, so masking the whole document is a no-op wherever `dropped`
 * did not actually apply. A no-op `dropped` (nothing to mask) returns
 * `source` itself, unmodified. */
export function withoutDropped(source: DocumentModel, dropped: DropRecord[]): DocumentModel {
  if (dropped.length === 0) {
    return source;
  }
  const kinds = new Set(dropped.map((item) => item.kind));
  const model = structuredClone(source);

  if (kinds.has("pageSetup") || kinds.has("headerFooter") || kinds.has("columns")) {
    // Markdown has no section concept at all: a real save flattens every
    // section's blocks into one continuous flow, carrying no page setup,
    // headers/footers or columns forward.
    const blocks = model.sections.flatMap((section) => section.blocks);
    model.sections = [{ blocks }];
  }
  if (kinds.has("pageSetup")) {
    // AC-5: a Markdown document always opens with the *default* page setup
    // — a real reopen never carries the source's own page size/margins
    // forward, even the document-wide defaults (Markdown has no geometry of
    // any kind, only content).
    model.defaultStyles = createDefaultStyles();
  }

  const OPAQUE_FAMILY_KINDS: DropKind[] = [
    "chart",
    "contentControl",
    "customXml",
    "embeddedObject",
    "macro",
    "opaqueBlock",
    "opaqueInline",
    "opaquePart",
  ];
  if (OPAQUE_FAMILY_KINDS.some((kind) => kinds.has(kind))) {
    // Markdown carries no package and no unmodelled-XML side-lists at all —
    // a real save loses every preserved opaque part/node outright, the same
    // "preserved .docx/.odt opaque parts and nodes" bucket
    // collectMarkdownDropRecords's caller (collectPackageDropRecords) is
    // the one actually reporting, since a .md source never has any of this
    // to begin with.
    delete model.opaque;
    for (const section of model.sections) {
      delete section.opaqueBlocks;
    }
  }

  if (kinds.has("tableOfContents")) {
    for (const section of model.sections) {
      section.blocks = section.blocks.flatMap((block) =>
        isToc(block)
          ? block.entries.map((entry): Paragraph => ({ type: "paragraph", runs: [{ text: entry.text, formatting: {} }] }))
          : [block],
      );
    }
  }

  if (kinds.has("tableOfFigures")) {
    for (const section of model.sections) {
      section.blocks = section.blocks.flatMap((block) =>
        isTof(block)
          ? block.entries.map((entry): Paragraph => ({ type: "paragraph", runs: [{ text: entry.text, formatting: {} }] }))
          : [block],
      );
    }
  }

  if (kinds.has("comment")) {
    delete model.comments;
  }
  if (kinds.has("note")) {
    delete model.notes;
    delete model.footnoteSettings;
    delete model.endnoteSettings;
  }

  const paragraphs: Paragraph[] = [];
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isTable(block)) {
        if (kinds.has("tableStyling")) {
          block.rows.forEach((row, rowIndex) => {
            row.minHeightCm = undefined;
            // Row 0 as the header is exactly what a GFM table represents
            // (the writer always treats it that way) — only a header
            // anywhere else is actually unsupported (see
            // tableHasUnsupportedStyling).
            if (rowIndex !== 0) {
              row.isHeader = undefined;
            }
            row.docxTrPr = undefined;
            row.odtStyleName = undefined;
            for (const cell of row.cells) {
              cell.colSpan = undefined;
              cell.rowSpan = undefined;
              cell.covered = undefined;
              cell.borders = undefined;
              cell.shadingFill = undefined;
              cell.paddingCm = undefined;
              cell.verticalAlign = undefined;
              cell.docxTcPr = undefined;
              cell.odtStyleName = undefined;
              if (cell.blocks.length === 0) {
                cell.blocks = [{ type: "paragraph", runs: [] }];
              }
            }
          });
          block.alignment = undefined;
          block.widthMode = undefined;
          block.widthPercent = undefined;
          block.styleId = undefined;
          block.docxTblPr = undefined;
          block.odtStyleName = undefined;
          block.odtColumnStyleNames = undefined;
          // GFM carries no column-width info — a real save always reflows
          // to an equal split of the text area (see
          // tableHasUnsupportedStyling).
          block.columnWidthsCm = equalColumnWidthsCm(model, block.columnWidthsCm.length);
        }
        if (kinds.has("trackedChange")) {
          delete block.changes;
          for (const row of block.rows) {
            delete row.changes;
            for (const cell of row.cells) {
              delete cell.changes;
            }
          }
        }
        for (const row of block.rows) {
          for (const cell of row.cells) {
            paragraphs.push(...cell.blocks);
          }
        }
      } else if (!isImage(block) && !isShape(block) && !isChart(block) && !isConnector(block) && !isBreak(block) && !isToc(block) && !isTof(block)) {
        paragraphs.push(block);
      }
    }
  }

  for (const paragraph of paragraphs) {
    if (kinds.has("paragraphFormatting") || kinds.has("runFormatting") || kinds.has("pageNumberField") || kinds.has("trackedChange")) {
      maskParagraph(model, paragraph);
    }
    if (kinds.has("bookmark")) {
      delete paragraph.bookmarks;
    }
    if (kinds.has("comment")) {
      delete paragraph.commentAnchors;
    }
    if (OPAQUE_FAMILY_KINDS.some((kind) => kinds.has(kind))) {
      delete paragraph.opaqueInlines;
    }
  }

  return model;
}
