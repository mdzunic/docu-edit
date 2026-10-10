// The .xlsx reader (SPEC-116, extended by SPEC-117): an OPC package whose
// main part declares SpreadsheetML (plain, template, macro-enabled or
// macro-enabled template) → the workbook model. Reads the sheet list, the
// date system, the active tab, the core properties, the theme and the style
// table from the small parts with DOMParser (core/xml.ts); reads shared
// strings and every worksheet with the forward-only scanner
// (core/xmlScanner.ts, in xlsx/sheetXml.ts) — a million-cell sheet as a DOM
// costs an order of magnitude more memory than its text.
//
// What is read (everything else — named cell styles, dxfs,
// drawings, comments, tables, calcChain, VBA, custom XML — stays in `opaque`,
// untouched; nothing here writes, so a download stays byte-identical):
// - sheets in `<sheets>` order: name, `state` → visibility, and the
//   relationship type → kind (worksheet → "grid"; chartsheet, dialogsheet,
//   macrosheet → "foreign", shown as a placeholder and carried verbatim);
// - the theme part, found through the workbook's `…/theme` relationship
//   (core/themeDrawingMl.ts readThemePart) → `theme`;
// - the styles part, through the `…/styles` relationship (xlsx/styles.ts):
//   cellXfs → `styles` one-to-one, `<indexedColors>` → `indexedColors`, and
//   the Normal style's font, whose maximum digit width (geometry.ts) turns
//   every column width into px;
//   SPEC-118: each xf's number format (`numFmts`, the built-ins), and a
//   derived style per implied format a `t="d"` value with a General style
//   needs (impliedFormats.ts), appended to `styles`;
// - shared strings: `<t>` or a run of `<r>`s with their `<rPr>` fonts,
//   phonetic `rPh`/`phoneticPr` skipped, `_xHHHH_` escapes decoded;
// - each worksheet (xlsx/sheetXml.ts): cells (rows and cells without `r`
//   placed sequentially; `t` absent/n, s, str, inlineStr, b, e, d; a
//   formula's cached `<v>` is the value — `<f>` is SPEC-119's; `s` →
//   `styleId`, a style-only cell stored), row heights/visibility/styles,
//   column spans, default sizes, merges, the tab colour and the first sheet
//   view (frozen pane, gridlines, zeros, zoom, top-left cell, selection);
// - `workbookPr@date1904`, `bookViews/workbookView@activeTab`,
//   `docProps/core.xml`;
// - SPEC-120: `calcPr@fullCalcOnLoad`, `@calcMode="manual"` and `@iterate`
//   → `calc` (read only: the element stays in the retained part, verbatim).
// - SPEC-122: `workbookProtection@lockStructure` "1"/"true" →
//   `structureProtected`; each worksheet's `sheetProtection@sheet` →
//   `protection`, its cells' `sstIndex`, `preservedAttrs` and `preservedXml`
//   (xlsx/sheetXml.ts). Read only: the protection XML stays verbatim.
// A foreign sheet keeps the defaults: no columns, no merges, 64 px columns
// and 15 pt rows. The cell ceiling stops the scan the moment the stored
// count would pass MAX_WORKBOOK_CELLS; a cell or row past the grid is a parse
// error naming its address.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import { parseCoreProperties } from "../../core/opcProperties";
import type { DocumentTheme } from "../../core/theme";
import { readThemePart } from "../../core/themeDrawingMl";
import { parseRelationships, parseXml, relsPathFor, resolvePartPath, XmlParseError, type Relationship } from "../../core/xml";
import { XmlScanError } from "../../core/xmlScanner";
import { unzip, ZipError } from "../../core/zip";
import type { DocumentProperties } from "../../core/documentProperties";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT, maxDigitWidthPx } from "../geometry";
import { ImpliedFormatStyles } from "../impliedFormats";
import type { CalcSettings, Sheet } from "../model";
import { CellBudget, encodeOpaque, parseError, partText, resolveActiveSheet, type ParsedSpreadsheet } from "../parsed";
import { readSharedStrings, readWorksheet } from "./sheetXml";
import { isSpreadsheetMl, readXlsxStyles } from "./styles";

export { decodeXmlEscapes, readSharedStrings, type SharedString } from "./sheetXml";

/** The `r:` relationship-id attribute namespace, transitional and strict. */
const RELATIONSHIP_ID_NAMESPACES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

const CORE_PROPERTIES_RELATIONSHIP = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

/** One `<sheet>` of the workbook part, resolved through its relationship. */
interface WorkbookSheet {
  name: string;
  visibility: Sheet["visibility"];
  kind: Sheet["kind"];
  partName: string;
}

function childElements(parent: Element, local: string): Element[] {
  return Array.from(parent.children).filter((child) => child.localName === local && isSpreadsheetMl(child.namespaceURI));
}

function relationshipId(element: Element): string | null {
  for (const ns of RELATIONSHIP_ID_NAMESPACES) {
    const id = element.getAttributeNS(ns, "id");
    if (id) {
      return id;
    }
  }
  return null;
}

/** The kind a sheet's relationship type makes it: only a worksheet is a
 * grid this spec can show. */
function sheetKindOf(relationshipType: string): Sheet["kind"] {
  return relationshipType.endsWith("/worksheet") ? "grid" : "foreign";
}

function readRelationships(parts: Map<string, Uint8Array>, partPath: string): Map<string, Relationship> {
  const bytes = parts.get(relsPathFor(partPath));
  if (!bytes) {
    return new Map();
  }
  try {
    return parseRelationships(partText(bytes));
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw parseError(relsPathFor(partPath), "its relationships are not well-formed XML");
    }
    throw error;
  }
}

/** What the workbook part says: its sheets, date system and active tab, and
 * the shared strings, theme and styles parts its relationships point at. */
interface WorkbookInfo {
  sheets: WorkbookSheet[];
  date1904: boolean;
  activeTab: number;
  /** SPEC-120: calcPr's fullCalcOnLoad, calcMode and iterate. */
  calc: CalcSettings;
  /** SPEC-122: workbookProtection@lockStructure. */
  structureProtected: boolean;
  sharedStringsPath: string | undefined;
  themePath: string | undefined;
  stylesPath: string | undefined;
}

/** An xsd:boolean attribute's truth ("1" or "true"). */
function isTrue(value: string | null): boolean {
  return value === "1" || value === "true";
}

/** SPEC-120 "Calculation settings": `calcPr@fullCalcOnLoad`, `@calcMode`
 * ("manual"; "auto" and "autoNoTable" are automatic) and `@iterate`. */
function readCalcPr(calcPr: Element | undefined): CalcSettings {
  const calc: CalcSettings = {};
  if (!calcPr) {
    return calc;
  }
  if (isTrue(calcPr.getAttribute("fullCalcOnLoad"))) {
    calc.fullCalcOnLoad = true;
  }
  if (calcPr.getAttribute("calcMode") === "manual") {
    calc.manual = true;
  }
  if (isTrue(calcPr.getAttribute("iterate"))) {
    calc.iterate = true;
  }
  return calc;
}

/** The workbook part (see WorkbookInfo). */
function readWorkbook(parts: Map<string, Uint8Array>, workbookPath: string): WorkbookInfo {
  const bytes = parts.get(workbookPath);
  if (!bytes) {
    throw parseError(workbookPath, "the workbook part is missing");
  }
  let root: Element;
  try {
    root = parseXml(partText(bytes)).documentElement;
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw parseError(workbookPath, "the workbook part is not well-formed XML");
    }
    throw error;
  }
  const relationships = readRelationships(parts, workbookPath);
  const date1904Attr = childElements(root, "workbookPr")[0]?.getAttribute("date1904");
  const bookViews = childElements(root, "bookViews")[0];
  const firstView = bookViews ? childElements(bookViews, "workbookView")[0] : undefined;
  const activeTab = Number(firstView?.getAttribute("activeTab") ?? 0);

  const sheets: WorkbookSheet[] = [];
  const sheetsElement = childElements(root, "sheets")[0];
  for (const element of sheetsElement ? childElements(sheetsElement, "sheet") : []) {
    const name = element.getAttribute("name") ?? `Sheet${sheets.length + 1}`;
    const rId = relationshipId(element);
    const relationship = rId ? relationships.get(rId) : undefined;
    if (!relationship) {
      throw parseError(workbookPath, `sheet "${name}" has no part in the package`);
    }
    const state = element.getAttribute("state");
    sheets.push({
      name,
      visibility: state === "hidden" ? "hidden" : state === "veryHidden" ? "veryHidden" : "visible",
      kind: sheetKindOf(relationship.type),
      partName: resolvePartPath(workbookPath, relationship.target),
    });
  }
  if (sheets.length === 0) {
    throw parseError(workbookPath, "the workbook has no sheets");
  }
  // The part of the first relationship whose type ends with `suffix`
  // (transitional and strict types share their last segment).
  const target = (suffix: string): string | undefined => {
    const relationship = [...relationships.values()].find((rel) => rel.type.endsWith(suffix));
    return relationship ? resolvePartPath(workbookPath, relationship.target) : undefined;
  };
  return {
    sheets,
    date1904: date1904Attr === "1" || date1904Attr === "true",
    activeTab: Number.isFinite(activeTab) ? activeTab : 0,
    calc: readCalcPr(childElements(root, "calcPr")[0]),
    structureProtected: isTrue(childElements(root, "workbookProtection")[0]?.getAttribute("lockStructure") ?? null),
    sharedStringsPath: target("/sharedStrings"),
    themePath: target("/theme"),
    stylesPath: target("/styles"),
  };
}

/** The workbook's theme part, read (undefined: none, or not a usable theme). */
function readTheme(parts: Map<string, Uint8Array>, themePath: string | undefined): DocumentTheme | undefined {
  const bytes = themePath ? parts.get(themePath) : undefined;
  return bytes ? readThemePart(partText(bytes)) : undefined;
}

function coreProperties(parts: Map<string, Uint8Array>): DocumentProperties {
  const rootRelationships = readRelationships(parts, "");
  const core = [...rootRelationships.values()].find((rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP);
  const bytes = core ? parts.get(resolvePartPath("", core.target)) : undefined;
  return bytes ? parseCoreProperties(partText(bytes)) : {};
}

/** Wraps a scanner failure in a part as the workbook's parse error. */
function scanned<T>(partName: string, read: () => T): T {
  try {
    return read();
  } catch (error) {
    if (error instanceof XmlScanError) {
      throw parseError(partName, `the part is not well-formed XML — ${error.message}`);
    }
    throw error;
  }
}

/** Parses an `.xlsx` (or `.xlsm`/`.xltx`/`.xltm`) package. `preunzipped`,
 * when the open funnel already inflated the package, is used instead of
 * `data` (an upgrade passes the retained parts and no bytes at all). */
export async function parseXlsx(
  data: Uint8Array | ArrayBuffer,
  preunzipped?: Map<string, Uint8Array>,
): Promise<ParsedSpreadsheet> {
  let parts = preunzipped;
  if (!parts) {
    try {
      parts = await unzip(data);
    } catch (error) {
      if (error instanceof ZipError) {
        throw new SpreadsheetOpenError("This file isn't an Excel workbook we can open");
      }
      throw error;
    }
  }
  const info = detectPackage(parts);
  const detected = spreadsheetFormatOf(info);
  if (info.container !== "opc" || detected === null || detected === "xlsb" || detected.format !== "xlsx") {
    throw new SpreadsheetOpenError("This file isn't an Excel workbook we can open");
  }
  const workbook = readWorkbook(parts, info.mainPartPath);
  const dateSystem = workbook.date1904 ? "1904" : "1900";
  const theme = readTheme(parts, workbook.themePath);
  const stylesBytes = workbook.stylesPath ? parts.get(workbook.stylesPath) : undefined;
  const styleTable = readXlsxStyles(stylesBytes ? partText(stylesBytes) : undefined, theme);
  const indexedColors = styleTable.indexedColors;
  // Every column width converts through the Normal font's maximum digit
  // width — a pinned table, never a measured font (geometry.ts).
  const mdw = maxDigitWidthPx(styleTable.defaultFont.family, styleTable.defaultFont.sizePt);
  const sharedStringsBytes = workbook.sharedStringsPath ? parts.get(workbook.sharedStringsPath) : undefined;
  const sharedStrings = sharedStringsBytes
    ? scanned(workbook.sharedStringsPath!, () => readSharedStrings(partText(sharedStringsBytes), indexedColors))
    : [];

  const budget = new CellBudget();
  const implied = new ImpliedFormatStyles(styleTable.styles);
  const sheets: Sheet[] = [];
  for (const entry of workbook.sheets) {
    const sheet: Sheet = {
      id: crypto.randomUUID(),
      name: entry.name,
      visibility: entry.visibility,
      kind: entry.kind,
      rows: [],
      partName: entry.partName,
      columns: [],
      defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
      defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
      merges: [],
    };
    const bytes = parts.get(entry.partName);
    if (entry.kind === "grid" && bytes) {
      const content = scanned(entry.partName, () =>
        readWorksheet(partText(bytes), entry.name, { sharedStrings, dateSystem, budget, mdw, indexedColors, implied }),
      );
      if (content === null) {
        throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
      Object.assign(sheet, content);
    }
    sheets.push(sheet);
  }

  const parsed: ParsedSpreadsheet = {
    sheets,
    styles: styleTable.styles,
    activeSheetIndex: resolveActiveSheet(workbook.activeTab, sheets),
    dateSystem,
    opaque: await encodeOpaque(parts),
    properties: coreProperties(parts),
  };
  if (theme !== undefined) {
    parsed.theme = theme;
  }
  if (indexedColors !== undefined) {
    parsed.indexedColors = indexedColors;
  }
  if (detected.packageVariant) {
    parsed.packageVariant = detected.packageVariant;
  }
  if (Object.keys(workbook.calc).length > 0) {
    parsed.calc = workbook.calc;
  }
  if (workbook.structureProtected) {
    parsed.structureProtected = true;
  }
  return parsed;
}
