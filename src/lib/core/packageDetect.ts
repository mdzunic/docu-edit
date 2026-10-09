// Container-level format detection, shared by every product's open path.
// Both package families this repo reads are ZIPs that declare what they are
// inside the package — an ODF package in its `mimetype` entry, an OPC
// package in `[Content_Types].xml` plus the root relationships — so the
// open funnel can classify an unzipped package here and hand the original
// bytes to the right reader (.docx/.odt/.pptx/.odp), instead of each format
// re-implementing "is this file mine". Detection never trusts a file
// extension (SPEC-004/005's rule); everything here reads package contents.
import { parseRelationships, parseXml, resolvePartPath, XmlParseError, type Relationship } from "./xml";

/** The OPC root relationship that names a package's main document part —
 * the same relationship type for WordprocessingML and PresentationML. */
export const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** An ODF package: a ZIP whose `mimetype` entry declares the document type
 * (e.g. "application/vnd.oasis.opendocument.text" / "…presentation"). */
export interface OdfPackageInfo {
  container: "odf";
  mimetype: string;
}

/** An OPC package: a ZIP whose root relationships name a main document part.
 * `mainContentType` is that part's declared content type from
 * `[Content_Types].xml` (e.g. "…wordprocessingml.document.main+xml" /
 * "…presentationml.presentation.main+xml"), undefined when the package
 * declares none for it. */
export interface OpcPackageInfo {
  container: "opc";
  mainPartPath: string;
  mainContentType: string | undefined;
}

/** A ZIP that is neither an ODF nor an OPC package (no `mimetype` entry, no
 * resolvable officeDocument relationship). */
export interface PlainZipInfo {
  container: "zip";
}

export type PackageInfo = OdfPackageInfo | OpcPackageInfo | PlainZipInfo;

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** The trimmed `mimetype` entry of an ODF package, or undefined when the
 * ZIP has none (i.e. it is not an ODF package). */
export function readOdfMimetype(parts: Map<string, Uint8Array>): string | undefined {
  const entry = parts.get("mimetype");
  if (!entry) {
    return undefined;
  }
  const declared = decodeUtf8(entry).trim();
  return declared.length > 0 ? declared : undefined;
}

/** The declared content type of `partPath` per `[Content_Types].xml`:
 * an exact `Override` wins, else the `Default` for the part's extension.
 * Undefined when the part (or a readable content-types part) has neither. */
function contentTypeOf(parts: Map<string, Uint8Array>, partPath: string): string | undefined {
  const contentTypesEntry = parts.get("[Content_Types].xml");
  if (!contentTypesEntry) {
    return undefined;
  }
  let doc: Document;
  try {
    doc = parseXml(decodeUtf8(contentTypesEntry));
  } catch (error) {
    if (error instanceof XmlParseError) {
      return undefined;
    }
    throw error;
  }
  const extension = partPath.includes(".") ? partPath.slice(partPath.lastIndexOf(".") + 1).toLowerCase() : "";
  let byDefault: string | undefined;
  for (const child of Array.from(doc.documentElement.children)) {
    if (child.namespaceURI !== CONTENT_TYPES_NAMESPACE) {
      continue;
    }
    if (child.localName === "Override" && child.getAttribute("PartName") === `/${partPath}`) {
      return child.getAttribute("ContentType") ?? undefined;
    }
    if (child.localName === "Default" && (child.getAttribute("Extension") ?? "").toLowerCase() === extension) {
      byDefault = child.getAttribute("ContentType") ?? undefined;
    }
  }
  return byDefault;
}

/** The main document part an OPC package's root relationships point at,
 * or undefined when `_rels/.rels` is missing, unreadable, or has no
 * officeDocument relationship. */
function opcMainPartPath(parts: Map<string, Uint8Array>): string | undefined {
  const rootRelsEntry = parts.get("_rels/.rels");
  if (!rootRelsEntry) {
    return undefined;
  }
  let rels: Map<string, Relationship>;
  try {
    rels = parseRelationships(decodeUtf8(rootRelsEntry));
  } catch (error) {
    if (error instanceof XmlParseError) {
      return undefined;
    }
    throw error;
  }
  const officeDocRel = Array.from(rels.values()).find((rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP);
  return officeDocRel ? resolvePartPath("", officeDocRel.target) : undefined;
}

// --- media detection (SPEC-107) --------------------------------------------

/** The media formats `detectMedia` can recognise from bytes. Structurally
 * identical to media/model.ts's MediaFormat — core cannot import the product
 * layer, so the union is restated here and stays assignable. */
export type DetectedMediaFormat = "pdf" | "png" | "jpeg" | "gif" | "svg" | "webp";

/** How deep into the file the `%PDF-` header may sit (some real files carry
 * a preamble before it — the offset is honoured by the parser too). */
const PDF_HEADER_SCAN_BYTES = 1024;

/** How much of the file the SVG root-element scan reads. */
const SVG_SCAN_BYTES = 4096;

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function startsWithAscii(bytes: Uint8Array, text: string, offset = 0): boolean {
  if (bytes.length < offset + text.length) {
    return false;
  }
  for (let i = 0; i < text.length; i += 1) {
    if (bytes[offset + i] !== text.charCodeAt(i)) {
      return false;
    }
  }
  return true;
}

/** The byte offset of `%PDF-` within the leading scan window, or -1. */
export function pdfHeaderOffset(bytes: Uint8Array): number {
  const limit = Math.min(bytes.length, PDF_HEADER_SCAN_BYTES);
  for (let i = 0; i + 5 <= limit; i += 1) {
    if (
      bytes[i] === 0x25 && // %
      bytes[i + 1] === 0x50 && // P
      bytes[i + 2] === 0x44 && // D
      bytes[i + 3] === 0x46 && // F
      bytes[i + 4] === 0x2d // -
    ) {
      return i;
    }
  }
  return -1;
}

/** The index just past the `<!DOCTYPE …>` declaration starting at `pos`,
 * or -1 when it does not close inside `text`. Quoted literals and an
 * internal subset (`[ … ]` — where Illustrator declares its entities) may
 * themselves contain `>`. */
function skipDoctype(text: string, pos: number): number {
  let quote = "";
  let depth = 0;
  for (let i = pos + 2; i < text.length; i += 1) {
    const char = text[i];
    if (quote) {
      if (char === quote) {
        quote = "";
      }
    } else if (depth > 0 && text.startsWith("<!--", i)) {
      const end = text.indexOf("-->", i + 4);
      if (end < 0) {
        return -1;
      }
      i = end + 2;
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === "[") {
      depth += 1;
    } else if (char === "]") {
      depth = Math.max(0, depth - 1);
    } else if (char === ">" && depth === 0) {
      return i + 1;
    }
  }
  return -1;
}

/** Whether the leading bytes are an XML document whose ROOT element is
 * `<svg>` — the one media sniff that is structural rather than a magic
 * number (an SVG has none). Skips an optional BOM (UTF-8 or UTF-16), the
 * XML declaration, doctype and comments, then requires the first real
 * element to be `svg` (optionally namespace-prefixed). A comment or doctype
 * that merely mentions svg is not a match. */
function sniffSvg(bytes: Uint8Array): boolean {
  const limit = Math.min(bytes.length, SVG_SCAN_BYTES);
  const encoding =
    bytes[0] === 0xff && bytes[1] === 0xfe
      ? "utf-16le"
      : bytes[0] === 0xfe && bytes[1] === 0xff
        ? "utf-16be"
        : "utf-8";
  let text: string;
  try {
    // The decoder drops a leading BOM itself.
    text = new TextDecoder(encoding, { fatal: false }).decode(bytes.subarray(0, limit));
  } catch {
    return false;
  }
  let pos = 0;
  if (text.charCodeAt(0) === 0xfeff) {
    pos = 1; // a BOM the decoder kept
  }
  for (;;) {
    while (pos < text.length && /\s/.test(text[pos])) {
      pos += 1;
    }
    if (text[pos] !== "<") {
      return false;
    }
    if (text.startsWith("<?", pos)) {
      const end = text.indexOf("?>", pos);
      if (end < 0) {
        return false;
      }
      pos = end + 2;
      continue;
    }
    if (text.startsWith("<!--", pos)) {
      const end = text.indexOf("-->", pos);
      if (end < 0) {
        return false;
      }
      pos = end + 3;
      continue;
    }
    if (text.startsWith("<!", pos)) {
      const end = skipDoctype(text, pos);
      if (end < 0) {
        return false;
      }
      pos = end;
      continue;
    }
    // The first real element: `<svg` or `<prefix:svg`, followed by a name
    // boundary so `<svgfoo>` never matches.
    const match = /^<([A-Za-z_][\w.-]*:)?svg(?=[\s/>])/.exec(text.slice(pos));
    return match !== null;
  }
}

/** Sniffs a media file's format from its leading bytes (SPEC-107) — run by
 * the open funnel BEFORE any unzip attempt, since media files are not ZIPs.
 * Never consults the file extension. Returns null for anything that is not
 * one of the six media formats. */
export function detectMedia(bytes: Uint8Array): DetectedMediaFormat | null {
  if (PNG_SIGNATURE.every((byte, i) => bytes[i] === byte)) {
    return "png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "jpeg";
  }
  if (startsWithAscii(bytes, "GIF87a") || startsWithAscii(bytes, "GIF89a")) {
    return "gif";
  }
  if (startsWithAscii(bytes, "RIFF") && startsWithAscii(bytes, "WEBP", 8)) {
    return "webp";
  }
  if (pdfHeaderOffset(bytes) >= 0) {
    return "pdf";
  }
  // SVG last: the only structural (non-magic-number) check.
  if (sniffSvg(bytes)) {
    return "svg";
  }
  return null;
}

/** Classifies an unzipped package by its own declarations. ODF is checked
 * first (its `mimetype` entry is the format's mandated self-description);
 * a ZIP with neither declaration reports `container: "zip"` and the caller
 * decides the fallback (the historical one: try it as a `.docx`). */
export function detectPackage(parts: Map<string, Uint8Array>): PackageInfo {
  const mimetype = readOdfMimetype(parts);
  if (mimetype !== undefined) {
    return { container: "odf", mimetype };
  }
  const mainPartPath = opcMainPartPath(parts);
  if (mainPartPath !== undefined) {
    return { container: "opc", mainPartPath, mainContentType: contentTypeOf(parts, mainPartPath) };
  }
  return { container: "zip" };
}
