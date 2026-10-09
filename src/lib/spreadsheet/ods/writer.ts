// The .ods writer (SPEC-121 "From scratch", "The .ods substitution path").
// Three paths, as xlsx/writer.ts:
//
//  1. No retained package: mimetype (first, STOREd), META-INF/manifest.xml
//     (version 1.2), meta.xml (core's buildMetaXml: the title and `modified`),
//     styles.xml (a table-cell default style and the `Default` cell style in
//     Calibri 11 pt, page layout `pm1`, master page `Default`), content.xml
//     (`office:version="1.2"`; `table:calculation-settings` only for the 1904
//     system or a custom `odsNullDate`; the generated `de-` automatic styles;
//     per sheet one table with one column and one empty row) and a
//     settings.xml holding only the view's `ActiveTable`.
//  2. Retained package, nothing changed: every part verbatim (mimetype first
//     and STOREd).
//  3. The substitution path: content.xml = the envelope prefix (the
//     generated `de-ta` table styles appended to its automatic styles; the
//     sheet-reference pass over its content validations) + each sheet's
//     table text (a clean original's or copy's sourceXml with the pass and
//     two start-tag splices — `table:name` when it differs from the model's,
//     `table:style-name` → a `de-ta` style when the model's visibility or
//     tab colour differs from what the source table style declares; a new
//     sheet's generated table) + the suffix (the pass, and the named ranges
//     and database ranges on deleted sheets removed). settings.xml's
//     ActiveTable names the active sheet and its per-table entries follow
//     renames, deletions and duplicates (a recoloured table's TabColor item
//     is dropped — the style attribute is authoritative). Embedded objects
//     and pictures only a deleted table showed are dropped with their
//     manifest entries; an object's chart ranges get the pass. Every other
//     part is verbatim.
//
// Interim until SPEC-122: a dirty sheet with a source is written as a clean
// one (source + pass + splices). LibreOffice does not recalculate an .ods on
// load by default, so a cached value the pipeline recalculated stays stale
// in the file — the documented cost of the interim. New sheets are empty.
import { now as clockNow } from "../../core/clock";
import { buildMetaXml } from "../../core/odfMeta";
import { decodedPartBytes } from "../../core/packageParts";
import { resolveColor, DEFAULT_THEME } from "../../core/theme";
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import type { SheetRemap } from "../formula/sheetRefs";
import type { Sheet, SpreadsheetModel } from "../model";
import { copyRemap, sourceSheetRemap } from "../sourceNames";
import { applySplices, attributeValue, decodeXmlText, escapeAttributeValue, escapeElementText, startTagAt, type Splice } from "../xmlSplice";
import { decodeUtf8, encodeUtf8 } from "../xlsx/package";
import {
  isObjectPart,
  objectsReferenced,
  picturesReferenced,
  readManifest,
  setTableAttribute,
  sourceTables,
  spliceManifest,
  tableNameOf,
  tableStyleNameOf,
} from "./package";
import { passOdsText, removeRangesOnDeletedSheets, rootNamespaces } from "./sheetNamePass";
import { namespaceDeclarations, OdsStyleBook, OdsStyleRegistry, readOdsContentStyles, readOdsStyleSheet, emptyOdsStyleSheet } from "./styles";

export interface WriteOdsOptions {
  /** meta.xml's `dc:date` from scratch (defaults to core `now()`). */
  modified?: Date;
}

export const ODS_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";
const ODF_DECLARATION = `<?xml version="1.0" encoding="UTF-8"?>`;
const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const OF_NS = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
const TABLEOOO_NS = "http://openoffice.org/2009/table";
const CONFIG_NS = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const OOO_NS = "http://openoffice.org/2004/office";
const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

const encode = (text: string): Uint8Array => encodeUtf8(text);

/** Serialises a workbook to .ods bytes (see the module header). */
export async function writeOds(model: SpreadsheetModel, options: WriteOdsOptions = {}): Promise<Uint8Array> {
  if (model.opaque.length === 0) {
    return zip(buildFromScratch(model, options));
  }
  const parts = decodedPartBytes(model.opaque);
  if (!model.structureDirty && !model.sheets.some((sheet) => sheet.dirty)) {
    return zip(mimetypeFirst([...parts].map(([name, data]) => ({ name, data }))));
  }
  return zip(mimetypeFirst(buildSubstitution(model, parts)));
}

/** ODF's rule: `mimetype` first and STOREd. */
function mimetypeFirst(entries: ZipEntryInput[]): ZipEntryInput[] {
  const mimetype = entries.find((entry) => entry.name === "mimetype");
  return mimetype ? [{ ...mimetype, store: true }, ...entries.filter((entry) => entry !== mimetype)] : entries;
}

// --- from scratch ---------------------------------------------------------------------------

/** A tab colour as the hex `tableooo:tab-color` writes. */
function tabColorHex(sheet: Sheet, model: SpreadsheetModel): string | undefined {
  return sheet.tabColor === undefined ? undefined : resolveColor(sheet.tabColor, model.theme ?? DEFAULT_THEME);
}

/** A generated table (SPEC-121 "From scratch"): one column, one empty row,
 * its styles from the registry. */
function generatedTable(sheet: Sheet, model: SpreadsheetModel, registry: OdsStyleRegistry, masterPage: string): string {
  const table = registry.tableStyle({ masterPage, display: sheet.visibility === "visible", tabColor: tabColorHex(sheet, model) });
  const column = registry.columnStyle(sheet.defaultColWidthPx);
  const row = registry.rowStyle(sheet.defaultRowHeightPt, true);
  return (
    `<table:table table:name="${escapeAttributeValue(sheet.name)}" table:style-name="${table}">` +
    `<table:table-column table:style-name="${column}" table:default-cell-style-name="Default"/>` +
    `<table:table-row table:style-name="${row}"><table:table-cell/></table:table-row></table:table>`
  );
}

/** `table:calculation-settings` for a date origin other than LibreOffice's
 * default 1899-12-30. */
function calculationSettings(model: SpreadsheetModel): string {
  const date = model.odsNullDate ?? (model.dateSystem === "1904" ? "1904-01-01" : undefined);
  return date === undefined ? "" : `<table:calculation-settings><table:null-date table:date-value="${date}"/></table:calculation-settings>`;
}

export const FROM_SCRATCH_ODS_STYLES_XML =
  `${ODF_DECLARATION}<office:document-styles xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:text="${TEXT_NS}" ` +
  `xmlns:table="${TABLE_NS}" xmlns:fo="${FO_NS}" xmlns:svg="${SVG_NS}" office:version="1.2">` +
  `<office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri"/></office:font-face-decls>` +
  `<office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style>` +
  `<style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:style></office:styles>` +
  `<office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles>` +
  `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`;

export const FROM_SCRATCH_ODS_MANIFEST_XML =
  `${ODF_DECLARATION}<manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="1.2">` +
  `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODS_MIMETYPE}"/>` +
  `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/></manifest:manifest>`;

/** settings.xml holding only the view's ActiveTable. */
export function fromScratchSettingsXml(activeTable: string): string {
  return (
    `${ODF_DECLARATION}<office:document-settings xmlns:office="${OFFICE_NS}" xmlns:config="${CONFIG_NS}" xmlns:ooo="${OOO_NS}" office:version="1.2">` +
    `<office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views">` +
    `<config:config-item-map-entry><config:config-item config:name="ActiveTable" config:type="string">${escapeElementText(activeTable)}</config:config-item>` +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
  );
}

function buildFromScratch(model: SpreadsheetModel, options: WriteOdsOptions): ZipEntryInput[] {
  const modified = options.modified ?? clockNow();
  const registry = new OdsStyleRegistry();
  const tables = model.sheets.map((sheet) => generatedTable(sheet, model, registry, "Default")).join("");
  const content =
    `${ODF_DECLARATION}<office:document-content xmlns:office="${OFFICE_NS}" xmlns:style="${STYLE_NS}" xmlns:text="${TEXT_NS}" ` +
    `xmlns:table="${TABLE_NS}" xmlns:fo="${FO_NS}" xmlns:of="${OF_NS}" xmlns:tableooo="${TABLEOOO_NS}" office:version="1.2">` +
    `<office:automatic-styles>${registry.xml()}</office:automatic-styles>` +
    `<office:body><office:spreadsheet>${calculationSettings(model)}${tables}</office:spreadsheet></office:body></office:document-content>`;
  const active = model.sheets[model.activeSheetIndex] ?? model.sheets[0];
  return [
    { name: "mimetype", data: encode(ODS_MIMETYPE), store: true },
    { name: "META-INF/manifest.xml", data: encode(FROM_SCRATCH_ODS_MANIFEST_XML) },
    { name: "meta.xml", data: encode(buildMetaXml({ title: model.title }, modified)) },
    { name: "styles.xml", data: encode(FROM_SCRATCH_ODS_STYLES_XML) },
    { name: "content.xml", data: encode(content) },
    { name: "settings.xml", data: encode(fromScratchSettingsXml(active?.name ?? "")) },
  ];
}

// --- substitution ---------------------------------------------------------------------------

/** The source table styles as the reader reads them. */
function sourceStyleBook(prefix: string, stylesXml: string | undefined): OdsStyleBook {
  const root = /<[\w.-]+:document-content\b[^>]*>/.exec(prefix)?.[0] ?? "";
  const declarations = namespaceDeclarations(
    [...root.matchAll(/\s(xmlns(?::[\w.-]+)?)\s*=\s*"([^"]*)"/g)].map((match) => ({ name: match[1]!, value: decodeXmlText(match[2]!) })),
  );
  const slices = [...prefix.matchAll(/<office:(font-face-decls|automatic-styles)\b[\s\S]*?<\/office:\1>/g)].map((match) => match[0]);
  const sheet = stylesXml === undefined ? emptyOdsStyleSheet() : readOdsStyleSheet(stylesXml);
  return new OdsStyleBook(sheet, readOdsContentStyles(slices, declarations));
}

/** A table style's master page, read from the envelope's automatic styles. */
function masterPageOf(prefix: string, styleName: string | undefined): string {
  if (styleName === undefined) {
    return "Default";
  }
  for (const match of prefix.matchAll(/<style:style\b[^>]*>/g)) {
    if (attributeValue(match[0], "style:name") === styleName && attributeValue(match[0], "style:family") === "table") {
      return attributeValue(match[0], "style:master-page-name") ?? "Default";
    }
  }
  return "Default";
}

/** The envelope prefix with generated styles added to its automatic styles
 * (the element created before `<office:body` when absent). */
function withGeneratedStyles(prefix: string, styles: string): string {
  if (styles === "") {
    return prefix;
  }
  const close = prefix.lastIndexOf("</office:automatic-styles>");
  if (close >= 0) {
    return prefix.slice(0, close) + styles + prefix.slice(close);
  }
  const selfClosing = /<office:automatic-styles\s*\/>/.exec(prefix);
  if (selfClosing) {
    return `${prefix.slice(0, selfClosing.index)}<office:automatic-styles>${styles}</office:automatic-styles>${prefix.slice(selfClosing.index + selfClosing[0].length)}`;
  }
  const body = prefix.indexOf("<office:body");
  const at = body >= 0 ? body : prefix.length;
  return `${prefix.slice(0, at)}<office:automatic-styles>${styles}</office:automatic-styles>${prefix.slice(at)}`;
}

function buildSubstitution(model: SpreadsheetModel, parts: Map<string, Uint8Array>): ZipEntryInput[] {
  const envelope = model.odsContentEnvelope;
  const contentBytes = parts.get("content.xml");
  if (!envelope || !contentBytes) {
    throw new Error("The workbook's retained package has no content envelope");
  }
  const sourceContent = decodeUtf8(contentBytes);
  const stylesXml = parts.has("styles.xml") ? decodeUtf8(parts.get("styles.xml")!) : undefined;
  const remap = sourceSheetRemap(model, parts);
  const namespaces = rootNamespaces(envelope.prefix);
  const registry = new OdsStyleRegistry([envelope.prefix, envelope.suffix, stylesXml ?? ""]);
  const book = sourceStyleBook(envelope.prefix, stylesXml);
  const recoloured = new Set<string>();

  const tables = model.sheets.map((sheet) => {
    if (sheet.sourceXml === undefined) {
      return generatedTable(sheet, model, registry, "Default");
    }
    let text = passOdsText(sheet.sourceXml, copyRemap(remap, sheet), namespaces);
    if (tableNameOf(text) !== sheet.name) {
      text = setTableAttribute(text, "table:name", sheet.name);
    }
    const styleName = tableStyleNameOf(text);
    const declared = book.tableStyle(styleName);
    const hidden = sheet.visibility !== "visible";
    const color = tabColorHex(sheet, model);
    if (declared.hidden !== hidden || declared.tabColor !== color) {
      const name = registry.tableStyle({ masterPage: masterPageOf(envelope.prefix, styleName), display: !hidden, tabColor: color });
      text = setTableAttribute(text, "table:style-name", name);
      if (declared.tabColor !== color) {
        recoloured.add(sheet.name);
      }
    }
    return text;
  });
  const prefix = withGeneratedStyles(passOdsText(envelope.prefix, remap, namespaces), registry.xml());
  const suffix = passOdsText(removeRangesOnDeletedSheets(envelope.suffix, remap), remap, namespaces);
  const content = prefix + tables.join("") + suffix;

  // Objects and pictures only a deleted source table showed.
  const deletedNames = new Set(remap.deleted.map((name) => name.toUpperCase()));
  const deletedTexts = sourceTables(sourceContent, envelope.prefix.length, envelope.suffix.length)
    .filter((table) => deletedNames.has(table.name.toUpperCase()))
    .map((table) => sourceContent.slice(table.start, table.end));
  const shownObjects = new Set(objectsReferenced(content));
  const shownPictures = new Set([...picturesReferenced(content), ...picturesReferenced(stylesXml ?? "")]);
  const droppedObjects = deletedTexts.flatMap(objectsReferenced).filter((object) => !shownObjects.has(object));
  const droppedPictures = new Set(deletedTexts.flatMap(picturesReferenced).filter((picture) => !shownPictures.has(picture)));
  const dropped = (name: string): boolean => droppedPictures.has(name) || droppedObjects.some((object) => isObjectPart(name, object));

  // Each embedded object's chart ranges: a copy's own object takes its remap.
  const objectRemap = new Map<string, SheetRemap>();
  for (const sheet of model.sheets) {
    if (sheet.copiedFrom !== undefined && sheet.sourceXml !== undefined) {
      for (const object of objectsReferenced(sheet.sourceXml)) {
        const shared = model.sheets.some((other) => other !== sheet && objectsReferenced(other.sourceXml ?? "").includes(object));
        if (!shared) {
          objectRemap.set(object, copyRemap(remap, sheet));
        }
      }
    }
  }

  const active = model.sheets[model.activeSheetIndex];
  const out: ZipEntryInput[] = [];
  for (const [name, bytes] of parts) {
    if (dropped(name)) {
      continue;
    }
    let data = bytes;
    if (name === "content.xml") {
      data = encode(content);
    } else if (name === "settings.xml") {
      data = encode(settingsXmlFor(decodeUtf8(bytes), model, remap, recoloured, active?.name));
    } else if (name === "META-INF/manifest.xml" && (droppedObjects.length > 0 || droppedPictures.size > 0)) {
      const xml = decodeUtf8(bytes);
      const remove = new Set(
        readManifest(xml)
          .map((entry) => entry.fullPath)
          .filter((path) => droppedPictures.has(path) || droppedObjects.some((object) => path === `${object}/` || isObjectPart(path, object))),
      );
      data = encode(spliceManifest(xml, remove, []));
    } else if (/^Object [^/]+\/content\.xml$/.test(name)) {
      const object = name.slice(0, name.indexOf("/"));
      const xml = decodeUtf8(bytes);
      const next = passOdsText(xml, objectRemap.get(object) ?? remap, rootNamespaces(xml));
      data = next === xml ? bytes : encode(next);
    }
    out.push({ name, data });
  }
  return out;
}

// --- settings.xml ----------------------------------------------------------------------------

/** settings.xml following the model (see the module header); every other
 * byte verbatim. */
function settingsXmlFor(xml: string, model: SpreadsheetModel, remap: SheetRemap, recoloured: ReadonlySet<string>, activeName: string | undefined): string {
  const splices: Splice[] = [];
  const renamed = new Map(Object.entries(remap.renamed).map(([from, to]) => [from.toUpperCase(), to] as const));
  const deleted = new Set(remap.deleted.map((name) => name.toUpperCase()));
  const copiesOf = new Map<string, Sheet[]>();
  for (const sheet of model.sheets) {
    if (sheet.copiedFrom !== undefined) {
      const key = sheet.copiedFrom.toUpperCase();
      copiesOf.set(key, [...(copiesOf.get(key) ?? []), sheet]);
    }
  }
  const stack: { local: string; name?: string; start: number; tagEnd: number }[] = [];
  let tablesDepth = -1;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      const name = scanAttr(event.attrs, "name", CONFIG_NS);
      const frame = { local: event.local, name, start: event.start, tagEnd: event.end };
      if (event.local === "config-item-map-named" && name === "Tables") {
        tablesDepth = stack.length + 1;
      }
      if (event.selfClosing) {
        return undefined;
      }
      stack.push(frame);
      return undefined;
    }
    if (event.type !== "close") {
      return undefined;
    }
    const frame = stack.pop();
    if (!frame) {
      return undefined;
    }
    if (frame.local === "config-item" && frame.name === "ActiveTable" && activeName !== undefined) {
      const current = decodeXmlText(xml.slice(frame.tagEnd, event.start));
      if (current !== activeName) {
        splices.push({ start: frame.tagEnd, end: event.start, text: escapeElementText(activeName) });
      }
    } else if (frame.local === "config-item-map-entry" && stack.length === tablesDepth && frame.name !== undefined) {
      const key = frame.name.toUpperCase();
      const element = xml.slice(frame.start, event.end);
      if (deleted.has(key) && !renamed.has(key)) {
        splices.push({ start: frame.start, end: event.end, text: copiesFor(key, element) });
        return undefined;
      }
      const newName = renamed.get(key) ?? frame.name;
      let own = renameEntry(element, newName);
      if (recoloured.has(newName)) {
        own = withoutTabColor(own);
      }
      const replaced = own + copiesFor(key, element);
      if (replaced !== element) {
        splices.push({ start: frame.start, end: event.end, text: replaced });
      }
    } else if (frame.local === "config-item-map-named" && frame.name === "Tables") {
      tablesDepth = -1;
    }
    return undefined;
  });
  return applySplices(xml, splices);

  function copiesFor(key: string, element: string): string {
    return (copiesOf.get(key) ?? [])
      .map((copy) => {
        const entry = renameEntry(element, copy.name);
        return recoloured.has(copy.name) ? withoutTabColor(entry) : entry;
      })
      .join("");
  }
}

function renameEntry(element: string, name: string): string {
  const tag = startTagAt(element, 0);
  const next = tag.replace(/(\sconfig:name\s*=\s*")[^"]*"/, `$1${escapeAttributeValue(name)}"`);
  return next + element.slice(tag.length);
}

function withoutTabColor(element: string): string {
  return element.replace(/<config:config-item\b[^>]*\bconfig:name="TabColor"[^>]*>[^<]*<\/config:config-item>/, "");
}
