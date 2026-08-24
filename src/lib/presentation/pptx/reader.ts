// The `.pptx` -> PresentationModel pipeline (preparation milestone): unzip
// the OPC package, follow the root relationships to the presentation part
// (never assume "ppt/presentation.xml" is where it is — the same rule as
// the docx reader's AC-3), walk p:sldIdLst through the presentation part's
// relationships to each slide, and extract a read-only text outline per
// slide. Every package part is retained verbatim (SPEC-008 applied
// wholesale) — nothing is modelled deeply enough to rebuild yet, so the
// writer re-emits exactly these parts.
import { bytesToBase64Async, type OpaquePart } from "../../core/packageParts";
import { detectPackage } from "../../core/packageDetect";
import { parseCoreProperties, parseCustomProperties } from "../../core/opcProperties";
import type { DocumentProperties } from "../../core/documentProperties";
import { unzip, ZipError } from "../../core/zip";
import {
  childElements,
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  XmlParseError,
  type Relationship,
} from "../../core/xml";
import type { ParsedPresentation, SlidePreview } from "../model";

export class PptxOpenError extends Error {}

export const INVALID_PPTX_MESSAGE = "This file isn't a PowerPoint presentation we can open";

/** The main part's declared content type that makes an OPC package a
 * presentation (ECMA-376 PresentationML). */
export const PPTX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

/** The file's MIME type, for downloads. */
export const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CUSTOM_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** The text lines of one `a:p` paragraph: run text concatenated in order,
 * split into separate lines at explicit `a:br` breaks. Empty lines drop —
 * this is a preview outline, not a faithful text model. */
function paragraphLines(paragraph: Element): string[] {
  const lines: string[] = [];
  let current = "";
  for (const child of Array.from(paragraph.children)) {
    if (child.namespaceURI !== A_NAMESPACE) {
      continue;
    }
    if (child.localName === "br") {
      lines.push(current);
      current = "";
    } else if (child.localName === "r" || child.localName === "fld") {
      for (const t of childElements(child, A_NAMESPACE, "t")) {
        current += t.textContent ?? "";
      }
    }
  }
  lines.push(current);
  return lines.filter((line) => line.trim().length > 0);
}

/** Whether a `p:sp` is a title placeholder (`p:ph type="title"|"ctrTitle"`). */
function isTitleShape(sp: Element): boolean {
  const nvSpPr = firstChildElement(sp, P_NAMESPACE, "nvSpPr");
  const nvPr = nvSpPr && firstChildElement(nvSpPr, P_NAMESPACE, "nvPr");
  const ph = nvPr && firstChildElement(nvPr, P_NAMESPACE, "ph");
  const type = ph?.getAttribute("type");
  return type === "title" || type === "ctrTitle";
}

/** Collects `p:sp` shapes in document order, descending into `p:grpSp`
 * groups (grouped text is still slide text). */
function collectShapes(container: Element, into: Element[]): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI !== P_NAMESPACE) {
      continue;
    }
    if (child.localName === "sp") {
      into.push(child);
    } else if (child.localName === "grpSp") {
      collectShapes(child, into);
    }
  }
}

/** One slide part's preview: title from the first title placeholder, every
 * other shape's text as lines. */
function slidePreviewOf(slideDoc: Document): SlidePreview {
  const cSld = firstChildElement(slideDoc.documentElement, P_NAMESPACE, "cSld");
  const spTree = cSld && firstChildElement(cSld, P_NAMESPACE, "spTree");
  const preview: SlidePreview = { lines: [] };
  if (!spTree) {
    return preview;
  }
  const shapes: Element[] = [];
  collectShapes(spTree, shapes);
  for (const sp of shapes) {
    const txBody = firstChildElement(sp, P_NAMESPACE, "txBody");
    if (!txBody) {
      continue;
    }
    const lines = childElements(txBody, A_NAMESPACE, "p").flatMap(paragraphLines);
    if (lines.length === 0) {
      continue;
    }
    if (preview.title === undefined && isTitleShape(sp)) {
      preview.title = lines.join(" ");
    } else {
      preview.lines.push(...lines);
    }
  }
  return preview;
}

/** Parses a `.pptx` file's bytes into the slide previews + retained parts
 * the presentation model needs. Throws `PptxOpenError` for anything that
 * isn't a valid, openable `.pptx` package. */
export async function parsePptx(data: ArrayBuffer | Uint8Array): Promise<ParsedPresentation> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      throw new PptxOpenError(INVALID_PPTX_MESSAGE);
    }
    throw error;
  }
  const info = detectPackage(parts);
  if (info.container !== "opc" || info.mainContentType !== PPTX_MAIN_CONTENT_TYPE) {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }

  const readPart = (path: string): string | undefined => {
    const raw = parts.get(path);
    return raw ? decodeUtf8(raw) : undefined;
  };
  const mainText = readPart(info.mainPartPath);
  if (!mainText) {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }
  let mainDoc: Document;
  try {
    mainDoc = parseXml(mainText);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new PptxOpenError(INVALID_PPTX_MESSAGE);
    }
    throw error;
  }
  if (mainDoc.documentElement.namespaceURI !== P_NAMESPACE || mainDoc.documentElement.localName !== "presentation") {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }

  // Slide order comes from p:sldIdLst; each p:sldId's r:id resolves through
  // the presentation part's own relationships to the slide part.
  const mainRelsText = readPart(relsPathFor(info.mainPartPath));
  const mainRels: Map<string, Relationship> = mainRelsText ? parseRelationships(mainRelsText) : new Map();
  const slides: SlidePreview[] = [];
  const sldIdLst = firstChildElement(mainDoc.documentElement, P_NAMESPACE, "sldIdLst");
  for (const sldId of sldIdLst ? childElements(sldIdLst, P_NAMESPACE, "sldId") : []) {
    const relId = sldId.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
    const rel = relId ? mainRels.get(relId) : undefined;
    if (!rel || rel.type !== SLIDE_RELATIONSHIP) {
      continue;
    }
    const slideText = readPart(resolvePartPath(info.mainPartPath, rel.target));
    if (!slideText) {
      continue;
    }
    try {
      slides.push(slidePreviewOf(parseXml(slideText)));
    } catch (error) {
      if (!(error instanceof XmlParseError)) {
        throw error;
      }
      // An unreadable slide part still counts as a slide (the package keeps
      // it verbatim); its preview is just empty.
      slides.push({ lines: [] });
    }
  }

  // Document properties (SPEC-041), resolved through the root relationships
  // like the docx reader does. Tolerant: metadata never blocks opening.
  let properties: DocumentProperties = {};
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const coreRel = Array.from(rootRels.values()).find((rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP);
    const coreText = coreRel ? readPart(resolvePartPath("", coreRel.target)) : undefined;
    if (coreText) {
      properties = parseCoreProperties(coreText);
    }
    const customRel = Array.from(rootRels.values()).find((rel) => rel.type === CUSTOM_PROPERTIES_RELATIONSHIP);
    const customText = customRel ? readPart(resolvePartPath("", customRel.target)) : undefined;
    if (customText) {
      const { customProperties, unsupportedCustomProperties } = parseCustomProperties(customText);
      if (customProperties.length > 0) {
        properties.customProperties = customProperties;
      }
      if (unsupportedCustomProperties.length > 0) {
        properties.unsupportedCustomProperties = unsupportedCustomProperties;
      }
    }
  }

  // Retain the whole package, in package order (SPEC-008 wholesale).
  const opaque: OpaquePart[] = [];
  for (const [name, partBytes] of parts) {
    opaque.push({ name, data: await bytesToBase64Async(partBytes) });
  }

  return { slides, opaque, properties };
}
