// Package integrity for SPEC-121's writer specs ("Tests": "a
// packageIntegrity(bytes) helper ... runs over every fixture × every
// command"). Framework-free: it reads the ZIP and scans parts with
// core/xmlScanner.ts, never a DOM. Returns the problems found ([] = sound).
//
// OPC (.xlsx): every internal relationship target exists and every `.rels`
// belongs to a part; every `r:` reference in an XML part names a relationship
// of that part; every part (but [Content_Types].xml and `.rels` parts) has a
// content type and no Override names a missing part; worksheet parts carry
// an Override (not just the xml Default); table ids and table
// names/displayNames are unique across the tables (case-insensitively) and
// clash with no definedName; xl/workbook.xml's sheetIds and sheet names are
// unique; at most one worksheet is `tabSelected`; every
// `definedName@localSheetId` is a sheet index; every VML part has its own
// `o:idmap` block; every XML part is well-formed.
// ODF (.ods): `mimetype` is the first entry and STOREd (read off the local
// header) and matches the manifest root's media type; every manifest
// file-entry but "/" and directories exists; every part but `mimetype` and the
// manifest is listed; content.xml's sheet names, `xml:id`s and `draw:id`s are
// unique; every embedded object or image link resolves; content.xml,
// settings.xml and the manifest are well-formed.
import { resolvePartPath } from "../../core/xml";
import { scanAttr, scanXml, type XmlScanEvent } from "../../core/xmlScanner";
import { unzip } from "../../core/zip";

const REL_NS = new Set(["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"]);
const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK_NS = "http://www.w3.org/1999/xlink";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

const decoder = new TextDecoder("utf-8");

/** The problems of a workbook package ([] = sound); see the module header. */
export async function packageIntegrity(bytes: Uint8Array): Promise<string[]> {
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    return [`not a ZIP: ${String(error)}`];
  }
  if (parts.has("[Content_Types].xml")) {
    return opcProblems(parts);
  }
  if (parts.has("mimetype") || parts.has("META-INF/manifest.xml")) {
    return odfProblems(bytes, parts);
  }
  return ["neither an OPC nor an ODF package"];
}

/** Scans a part, reporting it as malformed instead of throwing. */
function scan(problems: string[], name: string, text: string, onEvent: (event: XmlScanEvent) => void): boolean {
  try {
    scanXml(text, onEvent);
    return true;
  } catch (error) {
    problems.push(`${name} is not well-formed: ${String(error)}`);
    return false;
  }
}

function duplicates(values: readonly string[]): string[] {
  const seen = new Set<string>();
  const repeated = new Set<string>();
  for (const value of values) {
    (seen.has(value) ? repeated : seen).add(value);
  }
  return [...repeated];
}

// --- OPC ---------------------------------------------------------------------------

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function relsOwner(name: string): string | undefined {
  const match = /^(.*?)_rels\/([^/]*)\.rels$/.exec(name);
  return match ? `${match[1]}${match[2]}` : undefined;
}

function opcProblems(parts: Map<string, Uint8Array>): string[] {
  const problems: string[] = [];
  const text = (name: string): string => decoder.decode(parts.get(name));
  const lowerNames = new Set([...parts.keys()].map((name) => name.toLowerCase()));

  // Relationships.
  const relsOf = new Map<string, Rel[]>();
  for (const name of parts.keys()) {
    const owner = relsOwner(name);
    if (owner === undefined) {
      continue;
    }
    if (owner !== "" && !parts.has(owner)) {
      problems.push(`${name} belongs to the missing part ${owner}`);
    }
    const rels: Rel[] = [];
    scan(problems, name, text(name), (event) => {
      if (event.type === "open" && event.local === "Relationship") {
        rels.push({
          id: scanAttr(event.attrs, "Id") ?? "",
          type: scanAttr(event.attrs, "Type") ?? "",
          target: scanAttr(event.attrs, "Target") ?? "",
          external: scanAttr(event.attrs, "TargetMode") === "External",
        });
      }
    });
    for (const id of duplicates(rels.map((rel) => rel.id))) {
      problems.push(`${name} repeats the relationship id ${id}`);
    }
    for (const rel of rels) {
      const target = resolvePartPath(owner, rel.target);
      if (!rel.external && !lowerNames.has(target.toLowerCase())) {
        problems.push(`${name} ${rel.id} targets the missing part ${target}`);
      }
    }
    relsOf.set(owner, rels);
  }
  const targetsOfKind = (kind: string): string[] =>
    [...relsOf].flatMap(([owner, rels]) => rels.filter((rel) => !rel.external && rel.type.endsWith(`/${kind}`)).map((rel) => resolvePartPath(owner, rel.target)));

  // Content types.
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  scan(problems, "[Content_Types].xml", text("[Content_Types].xml"), (event) => {
    if (event.type !== "open") {
      return;
    }
    if (event.local === "Override") {
      const partName = (scanAttr(event.attrs, "PartName") ?? "").replace(/^\//, "").toLowerCase();
      if (overrides.has(partName)) {
        problems.push(`two Overrides name /${partName}`);
      }
      overrides.set(partName, scanAttr(event.attrs, "ContentType") ?? "");
    } else if (event.local === "Default") {
      defaults.set((scanAttr(event.attrs, "Extension") ?? "").toLowerCase(), scanAttr(event.attrs, "ContentType") ?? "");
    }
  });
  for (const partName of overrides.keys()) {
    if (!lowerNames.has(partName)) {
      problems.push(`an Override names the missing part /${partName}`);
    }
  }
  for (const name of parts.keys()) {
    if (name === "[Content_Types].xml" || name.endsWith(".rels")) {
      continue;
    }
    const extension = name.includes(".") ? name.slice(name.lastIndexOf(".") + 1).toLowerCase() : "";
    if (!overrides.has(name.toLowerCase()) && !defaults.has(extension)) {
      problems.push(`${name} has no content type`);
    }
  }
  for (const sheet of new Set(targetsOfKind("worksheet"))) {
    const type = overrides.get(sheet.toLowerCase());
    if (type === undefined) {
      problems.push(`the worksheet ${sheet} has no Override`);
    } else if (!type.endsWith("worksheet+xml")) {
      problems.push(`the worksheet ${sheet} is declared ${type}`);
    }
  }

  // Every r: reference names a relationship of its part; every XML part parses.
  const worksheetParts = new Set(targetsOfKind("worksheet"));
  let selectedSheets: string[] = [];
  for (const name of parts.keys()) {
    if (!/\.xml$/i.test(name) || name === "[Content_Types].xml") {
      continue;
    }
    const ids = new Set((relsOf.get(name) ?? []).map((rel) => rel.id));
    let selected = false;
    scan(problems, name, text(name), (event) => {
      if (event.type !== "open") {
        return;
      }
      for (const attr of event.attrs) {
        if (attr.ns !== undefined && REL_NS.has(attr.ns) && ["id", "embed", "link", "pict", "dm", "lo", "qs", "cs"].includes(attr.local) && attr.value !== "" && !ids.has(attr.value)) {
          problems.push(`${name} <${event.name}> refers to the missing relationship ${attr.value}`);
        }
      }
      if (worksheetParts.has(name) && event.local === "sheetView" && ["1", "true"].includes(scanAttr(event.attrs, "tabSelected") ?? "")) {
        selected = true;
      }
    });
    if (selected) {
      selectedSheets = [...selectedSheets, name];
    }
  }
  if (selectedSheets.length > 1) {
    problems.push(`more than one worksheet is tabSelected: ${selectedSheets.join(", ")}`);
  }

  // The workbook: sheets and defined names.
  const officeDocument = (relsOf.get("") ?? []).find((rel) => !rel.external && rel.type.endsWith("/officeDocument"));
  const workbookPath = officeDocument === undefined ? undefined : resolvePartPath("", officeDocument.target);
  const definedNames: string[] = [];
  if (workbookPath === undefined) {
    problems.push("no workbook part");
  } else {
    const sheetIds: string[] = [];
    const sheetNames: string[] = [];
    const localIds: string[] = [];
    const workbookRels = new Set((relsOf.get(workbookPath) ?? []).map((rel) => rel.id));
    scan(problems, workbookPath, text(workbookPath), (event) => {
      if (event.type !== "open") {
        return;
      }
      if (event.local === "sheet") {
        sheetIds.push(scanAttr(event.attrs, "sheetId") ?? "");
        sheetNames.push((scanAttr(event.attrs, "name") ?? "").toUpperCase());
        const rId = event.attrs.find((attr) => attr.local === "id" && attr.ns !== undefined && REL_NS.has(attr.ns))?.value;
        if (rId === undefined || !workbookRels.has(rId)) {
          problems.push(`the sheet "${scanAttr(event.attrs, "name")}" has no relationship (${rId ?? "no r:id"})`);
        }
      } else if (event.local === "definedName") {
        definedNames.push((scanAttr(event.attrs, "name") ?? "").toUpperCase());
        const local = scanAttr(event.attrs, "localSheetId");
        if (local !== undefined) {
          localIds.push(local);
        }
      }
    });
    for (const id of duplicates(sheetIds)) {
      problems.push(`duplicate sheetId ${id}`);
    }
    for (const name of duplicates(sheetNames)) {
      problems.push(`duplicate sheet name ${name}`);
    }
    for (const local of localIds) {
      const index = Number(local);
      if (!Number.isInteger(index) || index < 0 || index >= sheetIds.length) {
        problems.push(`definedName@localSheetId ${local} is not a sheet index (0..${sheetIds.length - 1})`);
      }
    }
  }

  // Tables.
  const tableIds: string[] = [];
  const tableNames: string[] = [];
  for (const name of [...parts.keys()].filter((part) => /^xl\/tables\/[^/]+\.xml$/.test(part))) {
    let root = true;
    scan(problems, name, text(name), (event) => {
      if (event.type === "open" && root) {
        root = false;
        tableIds.push(scanAttr(event.attrs, "id") ?? "");
        const own = new Set(["name", "displayName"].map((attr) => (scanAttr(event.attrs, attr) ?? "").toUpperCase()).filter((value) => value !== ""));
        tableNames.push(...own);
      }
    });
  }
  for (const id of duplicates(tableIds)) {
    problems.push(`duplicate table id ${id}`);
  }
  for (const name of duplicates(tableNames)) {
    problems.push(`duplicate table name ${name}`);
  }
  for (const name of new Set(tableNames)) {
    if (definedNames.includes(name)) {
      problems.push(`the table name ${name} is also a definedName`);
    }
  }

  // VML shape-id blocks.
  const blocks: string[] = [];
  for (const name of [...parts.keys()].filter((part) => part.toLowerCase().endsWith(".vml"))) {
    for (const match of text(name).matchAll(/<o:idmap\b[^>]*\bdata="([^"]*)"/g)) {
      blocks.push(...match[1]!.split(",").map((value) => value.trim()).filter((value) => value !== ""));
    }
  }
  for (const block of duplicates(blocks)) {
    problems.push(`two VML parts share the o:idmap block ${block}`);
  }
  return problems;
}

// --- ODF ------------------------------------------------------------------------------

/** The name and compression method of a ZIP's first local entry. */
function firstEntry(bytes: Uint8Array): { name: string; method: number } | undefined {
  if (bytes.length < 30 || bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
    return undefined;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const nameLength = view.getUint16(26, true);
  return { name: decoder.decode(bytes.subarray(30, 30 + nameLength)), method: view.getUint16(8, true) };
}

function odfProblems(bytes: Uint8Array, parts: Map<string, Uint8Array>): string[] {
  const problems: string[] = [];
  const text = (name: string): string => decoder.decode(parts.get(name));
  const first = firstEntry(bytes);
  if (first?.name !== "mimetype") {
    problems.push(`the first entry is ${first?.name ?? "unreadable"}, not mimetype`);
  } else if (first.method !== 0) {
    problems.push(`mimetype is compressed (method ${first.method}), not STOREd`);
  }

  const listed: string[] = [];
  let rootType: string | undefined;
  if (!parts.has("META-INF/manifest.xml")) {
    problems.push("no META-INF/manifest.xml");
  } else {
    scan(problems, "META-INF/manifest.xml", text("META-INF/manifest.xml"), (event) => {
      if (event.type === "open" && event.local === "file-entry") {
        const path = scanAttr(event.attrs, "full-path", MANIFEST_NS) ?? "";
        listed.push(path);
        if (path === "/") {
          rootType = scanAttr(event.attrs, "media-type", MANIFEST_NS);
        }
      }
    });
  }
  if (parts.has("mimetype") && rootType !== undefined && text("mimetype") !== rootType) {
    problems.push(`mimetype "${text("mimetype")}" is not the manifest's "${rootType}"`);
  }
  for (const path of duplicates(listed)) {
    problems.push(`the manifest lists ${path} twice`);
  }
  for (const path of listed) {
    if (path !== "/" && !path.endsWith("/") && !parts.has(path)) {
      problems.push(`the manifest lists the missing ${path}`);
    }
  }
  for (const name of parts.keys()) {
    if (name !== "mimetype" && name !== "META-INF/manifest.xml" && !name.endsWith("/") && !listed.includes(name)) {
      problems.push(`${name} is not in the manifest`);
    }
  }

  if (!parts.has("content.xml")) {
    problems.push("no content.xml");
  } else {
    const sheetNames: string[] = [];
    const xmlIds: string[] = [];
    const drawIds: string[] = [];
    const stack: string[] = [];
    scan(problems, "content.xml", text("content.xml"), (event) => {
      if (event.type === "open") {
        if (event.ns === TABLE_NS && event.local === "table" && stack.at(-1) === `${OFFICE_NS} spreadsheet`) {
          sheetNames.push((scanAttr(event.attrs, "name", TABLE_NS) ?? "").toUpperCase());
        }
        const xmlId = scanAttr(event.attrs, "id", XML_NS);
        if (xmlId !== undefined) {
          xmlIds.push(xmlId);
        }
        const drawId = scanAttr(event.attrs, "id", DRAW_NS);
        if (drawId !== undefined) {
          drawIds.push(drawId);
        }
        if (event.ns === DRAW_NS && (event.local === "object" || event.local === "image")) {
          const href = (scanAttr(event.attrs, "href", XLINK_NS) ?? "").replace(/^\.\//, "");
          if (href !== "" && !href.startsWith("#") && !/^[a-z][\w+.-]*:/i.test(href) && !listed.includes(href) && !listed.includes(`${href}/`) && !parts.has(href)) {
            problems.push(`draw:${event.local} links to the missing ${href}`);
          }
        }
        if (!event.selfClosing) {
          stack.push(`${event.ns} ${event.local}`);
        }
      } else if (event.type === "close") {
        stack.pop();
      }
    });
    for (const name of duplicates(sheetNames)) {
      problems.push(`duplicate table:name ${name}`);
    }
    for (const id of duplicates(xmlIds)) {
      problems.push(`duplicate xml:id ${id}`);
    }
    for (const id of duplicates(drawIds)) {
      problems.push(`duplicate draw:id ${id}`);
    }
  }
  if (parts.has("settings.xml")) {
    scan(problems, "settings.xml", text("settings.xml"), () => undefined);
  }
  return problems;
}
