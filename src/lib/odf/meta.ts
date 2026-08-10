// .odt `meta.xml` parsing/serialisation (SPEC-041): `office:meta`'s
// dc:title/dc:subject/dc:description/meta:keyword/meta:initial-creator/
// dc:creator/meta:creation-date/dc:date and `meta:user-defined` custom
// properties. Shared by reader.ts (parsing) and writer.ts (building), the
// same split as docx/docProps.ts.
//
// Two ODF fields are counter-intuitively named relative to their OOXML
// counterparts: `meta:initial-creator` is the ORIGINAL author (OOXML's
// `dc:creator`), while ODF's own `dc:creator` names the person who LAST
// MODIFIED the document (OOXML's `cp:lastModifiedBy`) — see ODF 1.2 §3.1.1.
import type {
  CustomProperty,
  CustomPropertyType,
  DocumentProperties,
  ExtraMetadataField,
  UnsupportedCustomProperty,
} from "../documentModel";
import { childElements, parseXml, serializeElement, XmlParseError } from "../docx/xml";
import { DC_NAMESPACE, META_NAMESPACE, OFFICE_NAMESPACE } from "./xml";

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** A timestamp in the W3CDTF profile meta.xml requires (whole seconds, UTC),
 * e.g. "2026-08-03T12:00:00Z". */
export function toW3CDTF(date: Date): string {
  return date.toISOString().replace(/\.\d+Z$/, "Z");
}

/** Display labels for known extra `office:meta` elements the dialog does not
 * expose for editing (AC-6) — an unrecognised element falls back to its bare
 * name. */
export const META_EXTRA_FIELD_LABELS: Record<string, string> = {
  generator: "Generator",
  "editing-cycles": "Editing cycles",
  "editing-duration": "Editing duration",
  template: "Template",
  "print-date": "Last printed",
  "printed-by": "Printed by",
};

const KNOWN_META_ELEMENTS = new Set([
  "title",
  "subject",
  "description",
  "keyword",
  "initial-creator",
  "creation-date",
  "user-defined",
]);

/** Parses `meta.xml` into the Description-tab fields plus every other
 * `office:meta` child as a read-only `extraFields` entry (AC-6). Tolerant of
 * a malformed part, like `docProps.ts`'s `parseCoreProperties`. */
export function parseOdfMeta(xml: string): DocumentProperties {
  let doc: Document;
  try {
    doc = parseXml(xml);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return {};
    }
    throw error;
  }
  const meta = childElements(doc.documentElement, OFFICE_NAMESPACE, "meta")[0];
  if (!meta) {
    return {};
  }
  const properties: DocumentProperties = {};
  const extraFields: ExtraMetadataField[] = [];
  const keywords: string[] = [];

  for (const child of Array.from(meta.children)) {
    const ns = child.namespaceURI;
    const name = child.localName;
    const text = child.textContent ?? "";
    if (ns === DC_NAMESPACE && name === "title") {
      properties.title = text;
    } else if (ns === DC_NAMESPACE && name === "subject") {
      properties.subject = text;
    } else if (ns === DC_NAMESPACE && name === "description") {
      properties.comments = text;
    } else if (ns === DC_NAMESPACE && name === "creator") {
      properties.lastModifiedBy = text;
    } else if (ns === DC_NAMESPACE && name === "date") {
      properties.modified = text;
    } else if (ns === META_NAMESPACE && name === "keyword") {
      keywords.push(text);
    } else if (ns === META_NAMESPACE && name === "initial-creator") {
      properties.author = text;
    } else if (ns === META_NAMESPACE && name === "creation-date") {
      properties.created = text;
    } else if (ns === META_NAMESPACE && name === "user-defined") {
      // Handled by parseOdfCustomProperties below.
    } else if (!(ns === META_NAMESPACE && KNOWN_META_ELEMENTS.has(name))) {
      extraFields.push({ name, xml: serializeElement(child) });
    }
  }
  if (keywords.length > 0) {
    properties.keywords = keywords.join(", ");
  }
  if (extraFields.length > 0) {
    properties.extraFields = extraFields;
  }
  return properties;
}

function readSupportedValue(el: Element, valueType: string | null): { type: CustomPropertyType; value: string } | undefined {
  const text = el.textContent ?? "";
  switch (valueType) {
    case "string":
      return { type: "text", value: text };
    case "float":
      return Number.isFinite(Number(text)) ? { type: "number", value: text } : undefined;
    case "boolean":
      return text === "true" || text === "false" ? { type: "boolean", value: text } : undefined;
    case "date":
      return /^\d{4}-\d{2}-\d{2}/.test(text) ? { type: "date", value: text.slice(0, 10) } : undefined;
    default:
      return undefined;
  }
}

/** Parses `meta.xml`'s `meta:user-defined` elements into typed custom
 * properties (AC-7..10). */
export function parseOdfCustomProperties(xml: string): {
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
  const meta = childElements(doc.documentElement, OFFICE_NAMESPACE, "meta")[0];
  const customProperties: CustomProperty[] = [];
  const unsupportedCustomProperties: UnsupportedCustomProperty[] = [];
  if (!meta) {
    return { customProperties, unsupportedCustomProperties };
  }
  for (const el of childElements(meta, META_NAMESPACE, "user-defined")) {
    const name = el.getAttributeNS(META_NAMESPACE, "name");
    if (!name) {
      continue;
    }
    const valueType = el.getAttributeNS(META_NAMESPACE, "value-type");
    const supported = readSupportedValue(el, valueType);
    if (supported) {
      customProperties.push({ name, type: supported.type, value: supported.value });
    } else {
      unsupportedCustomProperties.push({
        name,
        rawType: valueType ?? "string",
        rawXml: serializeElement(el),
      });
    }
  }
  return { customProperties, unsupportedCustomProperties };
}

/** What `buildMetaXml` needs — already resolved by the caller, the same way
 * `docProps.ts`'s `CorePropertiesInput` is. */
export interface OdfMetaInput {
  title: string;
  subject?: string;
  author?: string;
  lastModifiedBy?: string;
  keywords?: string;
  comments?: string;
  created?: string;
  extraFields?: ExtraMetadataField[];
  customProperties?: CustomProperty[];
  unsupportedCustomProperties?: UnsupportedCustomProperty[];
}

function fieldXml(tag: string, value: string | undefined): string {
  return value !== undefined && value.length > 0 ? `<${tag}>${escapeXml(value)}</${tag}>` : "";
}

function keywordsXml(keywords: string | undefined): string {
  if (!keywords) {
    return "";
  }
  return keywords
    .split(",")
    .map((keyword) => keyword.trim())
    .filter((keyword) => keyword.length > 0)
    .map((keyword) => `<meta:keyword>${escapeXml(keyword)}</meta:keyword>`)
    .join("");
}

function customValueXml(property: CustomProperty): string {
  switch (property.type) {
    case "text":
      return `<meta:user-defined meta:name="${escapeXml(property.name)}" meta:value-type="string">${escapeXml(property.value)}</meta:user-defined>`;
    case "number":
      return `<meta:user-defined meta:name="${escapeXml(property.name)}" meta:value-type="float">${escapeXml(property.value)}</meta:user-defined>`;
    case "boolean":
      return `<meta:user-defined meta:name="${escapeXml(property.name)}" meta:value-type="boolean">${property.value === "true" ? "true" : "false"}</meta:user-defined>`;
    case "date":
      return `<meta:user-defined meta:name="${escapeXml(property.name)}" meta:value-type="date">${escapeXml(property.value)}</meta:user-defined>`;
  }
}

/** `meta.xml` (AC-2, AC-3, AC-5, AC-7): the Description-tab fields,
 * `meta:creation-date`/`dc:date`, custom properties and any preserved extra
 * fields (AC-6). `meta:creation-date` is omitted rather than fabricated when
 * `input.created` is unset — same contract as `docProps.ts`'s
 * `buildCorePropertiesXml`, and for the same round-trip-fidelity reason. */
export function buildMetaXml(input: OdfMetaInput, modified: Date): string {
  const createdXml = input.created
    ? `<meta:creation-date>${toW3CDTF(new Date(input.created))}</meta:creation-date>`
    : "";
  const customXml = [
    ...(input.customProperties ?? []).map(customValueXml),
    ...(input.unsupportedCustomProperties ?? []).map((property) => property.rawXml),
  ].join("");
  const extraXml = (input.extraFields ?? []).map((field) => field.xml).join("");
  return (
    // Matches odf/writer.ts's own XML_DECLARATION (no `standalone`, unlike
    // .docx's parts) so a document with no metadata edits keeps byte-stable
    // output.
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta xmlns:office="${OFFICE_NAMESPACE}" xmlns:meta="${META_NAMESPACE}"` +
    ` xmlns:dc="${DC_NAMESPACE}" office:version="1.2">` +
    "<office:meta>" +
    fieldXml("dc:title", input.title) +
    fieldXml("dc:description", input.comments) +
    fieldXml("dc:subject", input.subject) +
    keywordsXml(input.keywords) +
    fieldXml("meta:initial-creator", input.author) +
    fieldXml("dc:creator", input.lastModifiedBy) +
    createdXml +
    `<dc:date>${toW3CDTF(modified)}</dc:date>` +
    customXml +
    extraXml +
    "</office:meta>" +
    "</office:document-meta>"
  );
}
