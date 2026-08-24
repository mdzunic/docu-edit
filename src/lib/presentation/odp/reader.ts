// The `.odp` -> PresentationModel pipeline (preparation milestone): verify
// the package's `mimetype` entry declares an OpenDocument presentation,
// walk content.xml's office:presentation draw:pages, and extract a
// read-only text outline per slide. Every package part is retained
// verbatim (SPEC-008 applied wholesale), `mimetype` included — the writer
// re-emits them with `mimetype` first and STOREd.
import { bytesToBase64Async, type OpaquePart } from "../../core/packageParts";
import { readOdfMimetype } from "../../core/packageDetect";
import { parseOdfCustomProperties, parseOdfMeta } from "../../core/odfMeta";
import type { DocumentProperties } from "../../core/documentProperties";
import { unzip, ZipError } from "../../core/zip";
import { firstChildElement, parseXml, XmlParseError } from "../../core/xml";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, TEXT_NAMESPACE, presentationAttr } from "../../core/odfXml";
import type { ParsedPresentation, SlidePreview } from "../model";

export class OdpOpenError extends Error {}

export const INVALID_ODP_MESSAGE = "This file isn't an OpenDocument presentation we can open";

/** The package mimetype identifying an OpenDocument presentation. */
export const ODP_MIMETYPE = "application/vnd.oasis.opendocument.presentation";
/** An `.otp` template's mimetype — same package shape, opened the same way
 * (the same convention as the odt reader's OTT handling). */
export const OTP_MIMETYPE = "application/vnd.oasis.opendocument.presentation-template";

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** Whether an element sits inside a title-classed frame — walked upward so
 * text nested in a title frame's text box attributes correctly. */
function inTitleFrame(el: Element, pageEl: Element): boolean {
  for (let node = el.parentElement; node && node !== pageEl; node = node.parentElement) {
    if (
      node.namespaceURI === DRAW_NAMESPACE &&
      node.localName === "frame" &&
      presentationAttr(node, "class") === "title"
    ) {
      return true;
    }
  }
  return false;
}

/** One draw:page's preview: title from title-classed frames, every other
 * text:p's text as lines, in document order. */
function pagePreviewOf(pageEl: Element): SlidePreview {
  const preview: SlidePreview = { lines: [] };
  const titleParts: string[] = [];
  const paragraphs = pageEl.getElementsByTagNameNS(TEXT_NAMESPACE, "p");
  for (const paragraph of Array.from(paragraphs)) {
    const text = (paragraph.textContent ?? "").trim();
    if (text.length === 0) {
      continue;
    }
    if (inTitleFrame(paragraph, pageEl)) {
      titleParts.push(text);
    } else {
      preview.lines.push(text);
    }
  }
  if (titleParts.length > 0) {
    preview.title = titleParts.join(" ");
  }
  return preview;
}

/** Parses an `.odp` file's bytes into the slide previews + retained parts
 * the presentation model needs. Throws `OdpOpenError` for anything that
 * isn't a valid, openable `.odp` package. */
export async function parseOdp(data: ArrayBuffer | Uint8Array): Promise<ParsedPresentation> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      throw new OdpOpenError(INVALID_ODP_MESSAGE);
    }
    throw error;
  }
  const mimetype = readOdfMimetype(parts);
  if (mimetype !== ODP_MIMETYPE && mimetype !== OTP_MIMETYPE) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }

  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }
  let contentDoc: Document;
  try {
    contentDoc = parseXml(decodeUtf8(contentBytes));
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new OdpOpenError(INVALID_ODP_MESSAGE);
    }
    throw error;
  }
  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const presentation = body && firstChildElement(body, OFFICE_NAMESPACE, "presentation");
  if (!presentation) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }

  const slides: SlidePreview[] = [];
  for (const child of Array.from(presentation.children)) {
    if (child.namespaceURI === DRAW_NAMESPACE && child.localName === "page") {
      slides.push(pagePreviewOf(child));
    }
  }

  // Document properties from meta.xml (SPEC-041) — the same merge the odt
  // reader performs. Tolerant: metadata never blocks opening.
  const metaBytes = parts.get("meta.xml");
  const metaText = metaBytes ? decodeUtf8(metaBytes) : undefined;
  const properties: DocumentProperties = metaText ? parseOdfMeta(metaText) : {};
  if (metaText) {
    const { customProperties, unsupportedCustomProperties } = parseOdfCustomProperties(metaText);
    if (customProperties.length > 0) {
      properties.customProperties = customProperties;
    }
    if (unsupportedCustomProperties.length > 0) {
      properties.unsupportedCustomProperties = unsupportedCustomProperties;
    }
  }

  // Retain the whole package, in package order (SPEC-008 wholesale) —
  // `mimetype` included, so the writer can re-emit it byte-identical.
  const opaque: OpaquePart[] = [];
  for (const [name, partBytes] of parts) {
    opaque.push({ name, data: await bytesToBase64Async(partBytes) });
  }

  return { slides, opaque, properties };
}
