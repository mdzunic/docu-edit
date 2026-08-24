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
