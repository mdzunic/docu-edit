// Shared SPEC-008 plumbing for the two writers: base64 codecs for opaque
// part bytes (the model must stay JSON-serialisable, so parts are not held
// as raw Uint8Arrays), the widened writer result type, and the single place
// the cross-format "dropped content" report is derived from the model. The
// writer performing a conversion is the only component that knows both what
// an item is and that it cannot be represented in the target format, so the
// conversion report dialog and the tests both read the writer's `dropped`
// array rather than diffing models.
import { isImage, type DocumentFormat, type DocumentModel } from "./documentModel";
import { parseXml, XmlParseError } from "./docx/xml";

// --- base64 ----------------------------------------------------------------

// Chunked so String.fromCharCode never exceeds the argument-count limit.
const BASE64_CHUNK = 0x8000;

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += BASE64_CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + BASE64_CHUNK));
  }
  return btoa(binary);
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** The model's opaque parts as name -> decoded bytes, in model order. */
export function opaquePartBytes(model: DocumentModel): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  for (const part of model.opaque ?? []) {
    parts.set(part.name, base64ToBytes(part.data));
  }
  return parts;
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
  | "svgImage";

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
};

const FORMAT_LABELS: Record<DocumentFormat, string> = {
  docx: "Word (.docx)",
  odt: "OpenDocument Text (.odt)",
};

function record(kind: DropKind, location: string, target: DocumentFormat): DropRecord {
  return {
    kind,
    location,
    reason: `Not preserved when saving as ${FORMAT_LABELS[target]}: ${REASONS[kind]}`,
  };
}

/** Everything a cross-format save of `model` will drop, in document order
 * (opaque nodes first, then unmodelled package parts). Empty when the model
 * carries no opaque content — a blank document converts losslessly. The
 * "approximate location" is expressed in paragraph numbers: the product
 * does not paginate, so page numbers do not exist to report. */
export function collectDropRecords(model: DocumentModel, target: DocumentFormat): DropRecord[] {
  const source = model.format;
  if (source === undefined || source === target) {
    return [];
  }

  const dropped: DropRecord[] = [];
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
    if (!isInfrastructurePart(source, part.name)) {
      dropped.push(record(partKind(source, part.name), `package part "${part.name}"`, target));
    }
  }

  return dropped;
}
