// OPC document-properties parsing/serialisation (SPEC-041): docProps/core.xml
// (dc:title/dc:subject/dc:creator/dc:description/cp:keywords/cp:lastModifiedBy/
// dcterms:created/dcterms:modified) and docProps/custom.xml (typed custom
// properties). These parts are defined at the OPC package level, not by any
// one document format — a `.pptx` carries the identical parts — so the
// module lives in core; the .docx reader (parsing) and writer (building)
// share it so the field list is declared exactly once. `docProps/app.xml`
// is deliberately not modelled here: none of its schema elements (Company,
// Manager, Template, Application, word/character-count snapshot, ...)
// correspond to anything this feature edits, so it stays a plain preserved
// OpaquePart like any other unmodelled package part (see the docx reader's
// opaque-part loop) rather than being parsed and rebuilt for no field this
// editor actually changes.
import type { CustomProperty, CustomPropertyType, DocumentProperties, ExtraMetadataField, UnsupportedCustomProperty } from "./documentProperties";
import { childElements, escapeXml, parseXml, serializeElement, XmlParseError } from "./xml";

export const CORE_PROPERTIES_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
export const DC_NAMESPACE = "http://purl.org/dc/elements/1.1/";
export const DCTERMS_NAMESPACE = "http://purl.org/dc/terms/";
export const XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance";
export const CUSTOM_PROPERTIES_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
export const VT_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/** Standard fmtid every `.docx` custom property carries (ECMA-376 §22.3). */
const CUSTOM_PROPERTY_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/** A timestamp in the W3CDTF profile core.xml requires (whole seconds, UTC),
 * e.g. "2026-08-03T12:00:00Z". */
export function toW3CDTF(date: Date): string {
  return date.toISOString().replace(/\.\d+Z$/, "Z");
}

// --- core.xml ----------------------------------------------------------

const CORE_FIELD_ELEMENTS = new Set(["title", "subject", "creator", "description", "keywords", "lastModifiedBy", "created", "modified"]);

/** Display labels for known extra core.xml elements the dialog does not
 * expose for editing (AC-6) — an unrecognised element falls back to its
 * bare name. */
export const CORE_EXTRA_FIELD_LABELS: Record<string, string> = {
  category: "Category",
  contentStatus: "Content status",
  revision: "Revision",
  version: "Version",
  lastPrinted: "Last printed",
  language: "Language",
  identifier: "Identifier",
};

/** Parses `docProps/core.xml` into the Description-tab fields plus every
 * other element as a read-only `extraFields` entry (AC-6). Tolerant of a
 * malformed part (returns `{}`) since a source file's metadata should never
 * block opening the document itself. */
export function parseCoreProperties(xml: string): DocumentProperties {
  let doc: Document;
  try {
    doc = parseXml(xml);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return {};
    }
    throw error;
  }
  const root = doc.documentElement;
  const properties: DocumentProperties = {};
  const extraFields: ExtraMetadataField[] = [];

  for (const child of Array.from(root.children)) {
    const ns = child.namespaceURI;
    const name = child.localName;
    const text = child.textContent ?? "";
    if (ns === DC_NAMESPACE && name === "title") {
      properties.title = text;
    } else if (ns === DC_NAMESPACE && name === "subject") {
      properties.subject = text;
    } else if (ns === DC_NAMESPACE && name === "creator") {
      properties.author = text;
    } else if (ns === DC_NAMESPACE && name === "description") {
      properties.comments = text;
    } else if (ns === CORE_PROPERTIES_NAMESPACE && name === "keywords") {
      properties.keywords = text;
    } else if (ns === CORE_PROPERTIES_NAMESPACE && name === "lastModifiedBy") {
      properties.lastModifiedBy = text;
    } else if (ns === DCTERMS_NAMESPACE && name === "created") {
      properties.created = text;
    } else if (ns === DCTERMS_NAMESPACE && name === "modified") {
      properties.modified = text;
    } else if (!(ns === CORE_PROPERTIES_NAMESPACE && CORE_FIELD_ELEMENTS.has(name))) {
      extraFields.push({ name, xml: serializeElement(child), value: text });
    }
  }
  if (extraFields.length > 0) {
    properties.extraFields = extraFields;
  }
  return properties;
}

/** What `buildCorePropertiesXml` needs — already resolved by the caller
 * (the effective title falls back to the model's own title when the
 * Description-tab field is blank; `created` defaults to `modified` the
 * first time a document is saved). */
export interface CorePropertiesInput {
  title: string;
  subject?: string;
  author?: string;
  lastModifiedBy?: string;
  keywords?: string;
  comments?: string;
  created?: string;
  extraFields?: ExtraMetadataField[];
}

function fieldXml(tag: string, value: string | undefined): string {
  return value !== undefined && value.length > 0 ? `<${tag}>${escapeXml(value)}</${tag}>` : "";
}

/** `docProps/core.xml` (AC-2, AC-3, AC-5): the Description-tab fields plus
 * `dcterms:created`/`dcterms:modified` and any preserved extra fields
 * (AC-6). `dcterms:created` is optional per ECMA-376 and is omitted rather
 * than fabricated when `input.created` is unset — a document with no known
 * creation date (never saved before, or opened from a source with none)
 * must round-trip as "still unknown", not silently gain one on every save;
 * setting it once and keeping it stable across saves is the caller's job
 * (see DocumentView.vue's save handler). */
export function buildCorePropertiesXml(input: CorePropertiesInput, modified: Date): string {
  const modifiedStamp = toW3CDTF(modified);
  const createdXml = input.created
    ? `<dcterms:created xsi:type="dcterms:W3CDTF">${toW3CDTF(new Date(input.created))}</dcterms:created>`
    : "";
  const extra = (input.extraFields ?? []).map((field) => field.xml).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NAMESPACE}" xmlns:dc="${DC_NAMESPACE}"` +
    ` xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}">` +
    fieldXml("dc:title", input.title) +
    fieldXml("dc:subject", input.subject) +
    fieldXml("dc:creator", input.author) +
    fieldXml("cp:lastModifiedBy", input.lastModifiedBy) +
    fieldXml("cp:keywords", input.keywords) +
    fieldXml("dc:description", input.comments) +
    createdXml +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${modifiedStamp}</dcterms:modified>` +
    extra +
    "</cp:coreProperties>"
  );
}

// --- custom.xml ----------------------------------------------------------

/** Reads a `vt:*`-typed value element into a supported `CustomProperty`, or
 * `undefined` when its type isn't one this editor supports (AC-9/AC-10). */
function readSupportedValue(valueEl: Element): { type: CustomPropertyType; value: string } | undefined {
  if (valueEl.namespaceURI !== VT_NAMESPACE) {
    return undefined;
  }
  const text = valueEl.textContent ?? "";
  switch (valueEl.localName) {
    case "lpwstr":
    case "bstr":
    case "lpstr":
      return { type: "text", value: text };
    case "i4":
    case "i2":
    case "r8":
    case "r4":
      return Number.isFinite(Number(text)) ? { type: "number", value: text } : undefined;
    case "bool":
      return { type: "boolean", value: text === "1" || text.toLowerCase() === "true" ? "true" : "false" };
    case "filetime": {
      // vt:filetime is a full W3CDTF datetime; the editable "date" type is
      // date-only (AC-9), so the time-of-day is truncated — round-tripping
      // an edited date always re-emits midnight UTC (see buildCustomPropertiesXml).
      const date = text.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(date) ? { type: "date", value: date } : undefined;
    }
    default:
      return undefined;
  }
}

/** Parses `docProps/custom.xml` into typed custom properties (AC-7..10).
 * Tolerant of a malformed part (returns empty), for the same reason
 * `parseCoreProperties` is. */
export function parseCustomProperties(xml: string): {
  customProperties: CustomProperty[];
  unsupportedCustomProperties: UnsupportedCustomProperty[];
} {
  let doc: Document;
  try {
    doc = parseXml(xml);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return { customProperties: [], unsupportedCustomProperties: [] };
    }
    throw error;
  }
  const customProperties: CustomProperty[] = [];
  const unsupportedCustomProperties: UnsupportedCustomProperty[] = [];
  for (const propertyEl of childElements(doc.documentElement, CUSTOM_PROPERTIES_NAMESPACE, "property")) {
    const name = propertyEl.getAttribute("name");
    if (!name) {
      continue;
    }
    const valueEl = Array.from(propertyEl.children)[0];
    const supported = valueEl ? readSupportedValue(valueEl) : undefined;
    if (supported) {
      customProperties.push({ name, type: supported.type, value: supported.value });
    } else {
      unsupportedCustomProperties.push({
        name,
        rawType: valueEl?.localName ?? "unknown",
        rawXml: serializeElement(propertyEl),
      });
    }
  }
  return { customProperties, unsupportedCustomProperties };
}

/** The pids a set of preserved raw `<property>` elements already use, parsed
 * back out of their verbatim XML so fresh pids never collide with them. */
function reservedPids(unsupported: UnsupportedCustomProperty[]): Set<number> {
  const pids = new Set<number>();
  for (const property of unsupported) {
    const match = /\bpid="(\d+)"/.exec(property.rawXml);
    if (match) {
      pids.add(Number(match[1]));
    }
  }
  return pids;
}

function customValueXml(property: CustomProperty): string {
  switch (property.type) {
    case "text":
      return `<vt:lpwstr>${escapeXml(property.value)}</vt:lpwstr>`;
    case "number":
      return `<vt:r8>${escapeXml(property.value)}</vt:r8>`;
    case "boolean":
      return `<vt:bool>${property.value === "true" ? "true" : "false"}</vt:bool>`;
    case "date":
      return `<vt:filetime>${escapeXml(property.value)}T00:00:00Z</vt:filetime>`;
  }
}

/** `docProps/custom.xml` (AC-7..10): one `<property>` per supported custom
 * property (fresh pids, never colliding with a preserved unsupported
 * property's own pid) plus every unsupported property's verbatim source
 * element, unchanged. Returns `undefined` when there is nothing to write
 * (AC "round-trip... adding, editing or removing a custom property" implies
 * removing the last one drops the part). */
export function buildCustomPropertiesXml(
  customProperties: CustomProperty[],
  unsupportedCustomProperties: UnsupportedCustomProperty[],
): string | undefined {
  if (customProperties.length === 0 && unsupportedCustomProperties.length === 0) {
    return undefined;
  }
  const reserved = reservedPids(unsupportedCustomProperties);
  let nextPid = 2;
  const nextFreePid = (): number => {
    while (reserved.has(nextPid)) {
      nextPid++;
    }
    return nextPid++;
  };
  const supportedXml = customProperties
    .map(
      (property) =>
        `<property fmtid="${CUSTOM_PROPERTY_FMTID}" pid="${nextFreePid()}" name="${escapeXml(property.name)}">` +
        customValueXml(property) +
        `</property>`,
    )
    .join("");
  const unsupportedXml = unsupportedCustomProperties.map((property) => property.rawXml).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Properties xmlns="${CUSTOM_PROPERTIES_NAMESPACE}" xmlns:vt="${VT_NAMESPACE}">` +
    supportedXml +
    unsupportedXml +
    "</Properties>"
  );
}
