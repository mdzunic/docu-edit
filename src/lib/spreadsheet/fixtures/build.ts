// DOM-free builders for the SPEC-116 workbook fixtures — the same pattern as
// presentation/fixtures/build.ts and media/fixtures/build.ts: the checked-in
// e2e fixtures are generated from here by
// `npx vite-node scripts/build-spreadsheet-fixtures.ts`, and the unit specs
// build the same bytes, so the two can never drift apart.
//
// They emulate real producers, because a fixture shaped like our own writer's
// output proves nothing:
// - mini-book.xlsx — Excel-shaped (Microsoft 365): part order and namespaces
//   as Excel writes them, `mc:Ignorable="x14ac xr xr2 xr3"` with `xr:uid`s,
//   `spans` and `x14ac:dyDescent` on rows, shared strings with a rich-text
//   item, an `_x000D_` escape and a phonetic run, a shared formula, a
//   `t="str"` formula result, a `#DIV/0!` error, booleans and a date serial,
//   `xl/calcChain.xml`, `docProps/app.xml` with `TitlesOfParts`, a theme and
//   a styles part. Three sheets: "Summary" (B2 = "Revenue"), "Data" (500 rows
//   written WITHOUT `r` attributes, as several non-Excel producers do) and
//   "Hidden" (`state="hidden"`).
// - mini-book.ods — LibreOffice-shaped (Calc 24.8): unindented content.xml
//   with LibreOffice's namespace set, three tables — "Archive" (hidden by its
//   automatic table style, and first, so the active table is the first TAB),
//   "Inventory" (active in settings.xml) and "Totals" — the trailing
//   `number-rows-repeated="1048553"` / `number-columns-repeated="16380"`
//   runs, a repeated non-empty row, covered cells under a span, every
//   `office:value-type`, `calcext:value-type`, `of:=` formulas with cached
//   values, an error cell, header rows, and no `table:null-date`.
// - spec116-1904.xlsx (`date1904="1"`), spec116-macros.xlsm (a dummy
//   `xl/vbaProject.bin` with its relationship and content type),
//   spec116-template.xltx (the template main content type), spec116-wide.ods
//   (one cell at column 16,000 of row 900,000, reached by empty runs).
// spec116-numbers.xlsx is NOT built here — Apple Numbers exported it (see
// e2e/fixtures/spec116-numbers.md); the specs read the committed file.
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import { dateToSerial } from "../serial";

const encoder = new TextEncoder();

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;
const ODF_XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8"?>\n`;

function part(name: string, text: string): ZipEntryInput {
  return { name, data: encoder.encode(text) };
}

// ============================================================================
// Excel
// ============================================================================

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/** The worksheet root Excel 365 writes, namespaces and `xr:uid` included. */
function worksheetOpen(uid: string, codeName?: string): string {
  return (
    `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
    `mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ` +
    `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ` +
    `xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" ` +
    `xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{${uid}}">` +
    (codeName ? `<sheetPr codeName="${codeName}"/>` : "")
  );
}

function worksheetXml(options: {
  uid: string;
  dimension: string;
  selected?: boolean;
  activeCell?: string;
  cols?: string;
  sheetData: string;
  codeName?: string;
}): string {
  const active = options.activeCell ?? "A1";
  return (
    XML_DECLARATION +
    worksheetOpen(options.uid, options.codeName) +
    `<dimension ref="${options.dimension}"/>` +
    `<sheetViews><sheetView ${options.selected ? 'tabSelected="1" ' : ""}workbookViewId="0">` +
    `<selection activeCell="${active}" sqref="${active}"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>` +
    (options.cols ?? "") +
    `<sheetData>${options.sheetData}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  );
}

/** One Excel-shaped `<row>`: `r`, `spans` and `x14ac:dyDescent`. */
function excelRow(r: number, spans: string, cells: string): string {
  return `<row r="${r}" spans="${spans}" x14ac:dyDescent="0.35">${cells}</row>`;
}

const sharedCell = (ref: string, index: number, style?: number): string =>
  `<c r="${ref}"${style ? ` s="${style}"` : ""} t="s"><v>${index}</v></c>`;
const numberCell = (ref: string, value: number | string, style?: number): string =>
  `<c r="${ref}"${style ? ` s="${style}"` : ""}><v>${value}</v></c>`;

/** A modest Office 2023 theme (DrawingML) — untouched by this spec, read by
 * SPEC-117. */
const EXCEL_THEME =
  XML_DECLARATION +
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>` +
  `<a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
  `<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>` +
  `<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

/** Excel's styles part: the default font, a bold header font, a date xf
 * (built-in numFmt 14) and a bold xf. Untouched by this spec (SPEC-117). */
const EXCEL_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

function coreXml(title: string): string {
  return (
    XML_DECLARATION +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
    `xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${title}</dc:title><dc:creator>${FIXTURE_WORKBOOK_AUTHOR}</dc:creator><cp:lastModifiedBy>${FIXTURE_WORKBOOK_AUTHOR}</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-01T09:00:00Z</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-08T17:30:00Z</dcterms:modified></cp:coreProperties>`
  );
}

function appXml(sheetNames: readonly string[]): string {
  const titles = sheetNames.map((name) => `<vt:lpstr>${name}</vt:lpstr>`).join("");
  return (
    XML_DECLARATION +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>` +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>` +
    `<vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts>` +
    `<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>` +
    `<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`
  );
}

interface ExcelSheet {
  name: string;
  state?: "hidden" | "veryHidden";
  xml: string;
}

/** Assembles an Excel-shaped package in Excel's own part order. */
async function excelPackage(options: {
  mainContentType: string;
  title: string;
  sheets: ExcelSheet[];
  sharedStrings?: string;
  calcChain?: string;
  workbookPr: string;
  activeTab?: number;
  /** Extra parts (name → bytes), their workbook relationship and content
   * type declaration — the VBA project. */
  extra?: { name: string; data: Uint8Array; relationship: string; contentType: string };
}): Promise<Uint8Array> {
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="${options.mainContentType}"/>`,
    ...options.sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`),
    `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
    ...(options.sharedStrings
      ? [`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`]
      : []),
    ...(options.calcChain
      ? [`<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>`]
      : []),
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
  ];
  const contentTypes =
    XML_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    (options.extra ? options.extra.contentType : "") +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>${overrides.join("")}</Types>`;
  const rootRels =
    XML_DECLARATION +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId3" Type="${REL_TYPE}/extended-properties" Target="docProps/app.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId1" Type="${REL_TYPE}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const n = options.sheets.length;
  const workbookRels = [
    ...options.sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${REL_TYPE}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).toReversed(),
    `<Relationship Id="rId${n + 1}" Type="${REL_TYPE}/theme" Target="theme/theme1.xml"/>`,
    `<Relationship Id="rId${n + 2}" Type="${REL_TYPE}/styles" Target="styles.xml"/>`,
    ...(options.sharedStrings ? [`<Relationship Id="rId${n + 3}" Type="${REL_TYPE}/sharedStrings" Target="sharedStrings.xml"/>`] : []),
    ...(options.calcChain ? [`<Relationship Id="rId${n + 4}" Type="${REL_TYPE}/calcChain" Target="calcChain.xml"/>`] : []),
    ...(options.extra ? [options.extra.relationship] : []),
  ];
  const sheetsXml = options.sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${index + 1}"${sheet.state ? ` state="${sheet.state}"` : ""} r:id="rId${index + 1}"/>`)
    .join("");
  const workbook =
    XML_DECLARATION +
    `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
    `mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" ` +
    `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ` +
    `xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" ` +
    `xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" ` +
    `xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">` +
    `<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="28127"/>${options.workbookPr}` +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15">` +
    `<x15ac:absPath url="C:\\Users\\fixture\\Documents\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>` +
    `</mc:Choice></mc:AlternateContent>` +
    `<xr:revisionPtr revIDLastSave="0" documentId="8_{6A3D0C51-2F8B-4F0E-9C1D-4B7E2A9F1C33}" xr6:coauthVersionLast="47" ` +
    `xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>` +
    `<bookViews><workbookView xWindow="-110" yWindow="-110" windowWidth="19420" windowHeight="10300"` +
    `${options.activeTab ? ` activeTab="${options.activeTab}"` : ""} xr2:uid="{1E0B4C7A-5D2F-4C9B-8E3A-7F6D5C4B3A21}"/></bookViews>` +
    `<sheets>${sheetsXml}</sheets><calcPr calcId="191029"/></workbook>`;

  const entries: ZipEntryInput[] = [
    part("[Content_Types].xml", contentTypes),
    part("_rels/.rels", rootRels),
    part("xl/workbook.xml", workbook),
    part("xl/_rels/workbook.xml.rels", XML_DECLARATION + `<Relationships xmlns="${PKG_REL_NS}">${workbookRels.join("")}</Relationships>`),
    ...options.sheets.map((sheet, index) => part(`xl/worksheets/sheet${index + 1}.xml`, sheet.xml)),
    part("xl/theme/theme1.xml", EXCEL_THEME),
    part("xl/styles.xml", EXCEL_STYLES),
  ];
  if (options.sharedStrings) {
    entries.push(part("xl/sharedStrings.xml", options.sharedStrings));
  }
  if (options.calcChain) {
    entries.push(part("xl/calcChain.xml", options.calcChain));
  }
  if (options.extra) {
    entries.push({ name: options.extra.name, data: options.extra.data });
  }
  entries.push(part("docProps/core.xml", coreXml(options.title)), part("docProps/app.xml", appXml(options.sheets.map((sheet) => sheet.name))));
  return zip(entries);
}

export const FIXTURE_WORKBOOK_AUTHOR = "Fixture Author";

/** An Excel shared-strings part over plain strings (Excel writes no inline
 * strings, even for a one-cell workbook). */
function sharedStringsXml(strings: readonly string[], count = strings.length): string {
  const items = strings.map((text) => `<si><t>${text}</t></si>`).join("");
  return XML_DECLARATION + `<sst xmlns="${MAIN_NS}" count="${count}" uniqueCount="${strings.length}">${items}</sst>`;
}

/** mini-book.xlsx's shared strings, in index order. Index 0 is rich text,
 * 7 carries an `_x000D_` escape, 22 a phonetic run. */
const MINI_BOOK_STRINGS = [
  `<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">Q3 </t></r>` +
    `<r><rPr><sz val="11"/><color theme="1"/><rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr><t>report</t></r></si>`,
  "<si><t>Region</t></si>",
  "<si><t>Revenue</t></si>",
  "<si><t>Cost</t></si>",
  "<si><t>Margin</t></si>",
  "<si><t>Notes</t></si>",
  "<si><t>North</t></si>",
  `<si><t xml:space="preserve">Line one_x000D_\nLine two</t></si>`,
  "<si><t>South</t></si>",
  "<si><t>Total</t></si>",
  "<si><t>Per unit</t></si>",
  "<si><t>Updated</t></si>",
  "<si><t>Audited</t></si>",
  "<si><t>Label</t></si>",
  "<si><t>Big</t></si>",
  "<si><t>Tiny</t></si>",
  "<si><t>ID</t></si>",
  "<si><t>Amount</t></si>",
  "<si><t>Parity</t></si>",
  "<si><t>even</t></si>",
  "<si><t>odd</t></si>",
  "<si><t>Secret</t></si>",
  `<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>`,
];

/** What mini-book.xlsx holds, for the specs. */
export const MINI_BOOK_XLSX = {
  sheetNames: ["Summary", "Data", "Hidden"],
  dataRows: 500,
  /** Summary's cells (A1 notation) → the value the reader must produce. */
  summary: {
    A1: "Q3 report",
    B2: "Revenue",
    B3: 1200,
    C3: 800,
    D3: 400,
    E3: "Line one\r\nLine two",
    B4: 950.5,
    D4: 250.25,
    B5: 2150.5,
    C5: 1500.25,
    D5: 650.25,
    B6: { error: "#DIV/0!" },
    B7: 45573,
    B8: true,
    C8: false,
    B9: "North region",
    B10: 123456789012,
    B11: 0.0000000001234,
  },
} as const;

export async function buildMiniBookXlsxBytes(): Promise<Uint8Array> {
  const summaryRows = [
    excelRow(1, "1:5", sharedCell("A1", 0, 2)),
    excelRow(2, "1:5", [sharedCell("A2", 1, 2), sharedCell("B2", 2, 2), sharedCell("C2", 3, 2), sharedCell("D2", 4, 2), sharedCell("E2", 5, 2)].join("")),
    excelRow(
      3,
      "1:5",
      sharedCell("A3", 6) + numberCell("B3", 1200) + numberCell("C3", 800) +
        `<c r="D3"><f t="shared" ref="D3:D4" si="0">B3-C3</f><v>400</v></c>` + sharedCell("E3", 7),
    ),
    excelRow(4, "1:5", sharedCell("A4", 8) + numberCell("B4", 950.5) + numberCell("C4", 700.25) + `<c r="D4"><f t="shared" si="0"/><v>250.25</v></c>`),
    excelRow(
      5,
      "1:5",
      sharedCell("A5", 9) + `<c r="B5"><f>SUM(B3:B4)</f><v>2150.5</v></c><c r="C5"><f>SUM(C3:C4)</f><v>1500.25</v></c>` +
        `<c r="D5"><f>B5-C5</f><v>650.25</v></c>`,
    ),
    excelRow(6, "1:5", sharedCell("A6", 10) + `<c r="B6" t="e"><f>B5/0</f><v>#DIV/0!</v></c>`),
    excelRow(7, "1:5", sharedCell("A7", 11) + numberCell("B7", 45573, 1)),
    excelRow(8, "1:5", sharedCell("A8", 12) + `<c r="B8" t="b"><v>1</v></c><c r="C8" t="b"><v>0</v></c>`),
    excelRow(9, "1:5", sharedCell("A9", 13) + `<c r="B9" t="str"><f>A3&amp;" region"</f><v>North region</v></c>`),
    excelRow(10, "1:5", sharedCell("A10", 14) + numberCell("B10", 123456789012)),
    excelRow(11, "1:5", sharedCell("A11", 15) + numberCell("B11", "1.234E-10")),
  ].join("");
  const summary = worksheetXml({
    uid: "3F1A9C2E-7B4D-4E8A-9C1F-2D6B5A4E3C11",
    dimension: "A1:E11",
    selected: true,
    activeCell: "B2",
    cols: `<cols><col min="1" max="1" width="14.6328125" customWidth="1"/></cols>`,
    sheetData: summaryRows,
  });

  // "Data": every row and cell without `r` — placed sequentially.
  const dataRows = [`<row><c t="s"><v>16</v></c><c t="s"><v>17</v></c><c t="s"><v>18</v></c></row>`];
  for (let i = 1; i < MINI_BOOK_XLSX.dataRows; i += 1) {
    dataRows.push(`<row><c><v>${i}</v></c><c><v>${i * 1.5}</v></c><c t="s"><v>${i % 2 === 0 ? 19 : 20}</v></c></row>`);
  }
  const data = worksheetXml({
    uid: "8C2D4E6F-1A3B-4C5D-8E7F-9A0B1C2D3E44",
    dimension: `A1:C${MINI_BOOK_XLSX.dataRows}`,
    sheetData: dataRows.join(""),
  });

  const hidden = worksheetXml({
    uid: "5B7C9D1E-3F5A-4B7C-9D1E-3F5A7B9C1D55",
    dimension: "A1:B2",
    sheetData: excelRow(1, "1:2", sharedCell("A1", 21) + numberCell("B1", 42)) + excelRow(2, "1:2", sharedCell("A2", 22)),
  });

  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${MINI_BOOK_STRINGS.length + MINI_BOOK_XLSX.dataRows + 1}" uniqueCount="${MINI_BOOK_STRINGS.length}">` +
    `${MINI_BOOK_STRINGS.join("")}</sst>`;
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}"><c r="D3" i="1" l="1"/><c r="D4"/><c r="B5"/><c r="C5"/><c r="D5"/><c r="B6"/><c r="B9"/></calcChain>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Mini book",
    sheets: [
      { name: "Summary", xml: summary },
      { name: "Data", xml: data },
      { name: "Hidden", state: "hidden", xml: hidden },
    ],
    sharedStrings,
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
  });
}

/** spec116-1904.xlsx: a Mac-era workbook on the 1904 date system. */
export const SPEC116_1904 = {
  /** 2026-10-08 as a 1904-system serial. */
  dateSerial: dateToSerial(2026, 10, 8, "1904"),
};

export async function buildSpec1904XlsxBytes(): Promise<Uint8Array> {
  const sheet = worksheetXml({
    uid: "0D2E4F6A-8B0C-4D2E-8F4A-6B8C0D2E4F66",
    dimension: "A1:B2",
    selected: true,
    sheetData:
      excelRow(1, "1:2", sharedCell("A1", 0) + numberCell("B1", SPEC116_1904.dateSerial, 1)) +
      excelRow(2, "1:2", sharedCell("A2", 1) + numberCell("B2", 0.5)),
  });
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "1904 dates",
    sheets: [{ name: "Dates", xml: sheet }],
    sharedStrings: sharedStringsXml(["Shipped", "Noon"]),
    workbookPr: `<workbookPr date1904="1" defaultThemeVersion="202300"/>`,
  });
}

/** The dummy VBA project spec116-macros.xlsm carries — bytes this product
 * keeps, never code it reads (an OLE header and a marker). */
export const SPEC116_VBA_PROJECT = new Uint8Array([
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00, 0x00, 0x00, 0x53, 0x50, 0x45, 0x43, 0x2d, 0x31, 0x31, 0x36,
  0x20, 0x76, 0x62, 0x61, 0x50, 0x72, 0x6f, 0x6a, 0x65, 0x63, 0x74, 0xff, 0x00, 0x01, 0x02, 0x03,
]);

export async function buildSpecMacrosXlsmBytes(): Promise<Uint8Array> {
  const sheet = worksheetXml({
    uid: "6E8F0A2B-4C6D-4E8F-8A2B-4C6D8E0F2A77",
    dimension: "A1",
    selected: true,
    codeName: "Sheet1",
    sheetData: excelRow(1, "1:1", sharedCell("A1", 0)),
  });
  return excelPackage({
    mainContentType: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    title: "Macros",
    sheets: [{ name: "Macros", xml: sheet }],
    sharedStrings: sharedStringsXml(["Run the macro"]),
    workbookPr: `<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>`,
    extra: {
      name: "xl/vbaProject.bin",
      data: SPEC116_VBA_PROJECT,
      relationship: `<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`,
      contentType: `<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>`,
    },
  });
}

export async function buildSpecTemplateXltxBytes(): Promise<Uint8Array> {
  const sheet = worksheetXml({
    uid: "7F9A1B3C-5D7E-4F9A-9B3C-5D7E9F1A3B88",
    dimension: "A1:B2",
    selected: true,
    sheetData:
      excelRow(1, "1:2", sharedCell("A1", 0)) +
      excelRow(2, "1:2", sharedCell("A2", 1) + numberCell("B2", 0)),
  });
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    title: "Invoice template",
    sheets: [{ name: "Template", xml: sheet }],
    sharedStrings: sharedStringsXml(["Invoice", "Amount"]),
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
  });
}

// ============================================================================
// LibreOffice
// ============================================================================

/** The namespace set LibreOffice Calc declares on content.xml. */
const LO_CONTENT_NAMESPACES =
  `xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" ` +
  `xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:xhtml="http://www.w3.org/1999/xhtml" ` +
  `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" ` +
  `xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:dom="http://www.w3.org/2001/xml-events" ` +
  `xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" ` +
  `xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:ooo="http://openoffice.org/2004/office" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:ooow="http://openoffice.org/2004/writer" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:oooc="http://openoffice.org/2004/calc" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" ` +
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" ` +
  `xmlns:rpt="http://openoffice.org/2005/report" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" ` +
  `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
  `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" office:version="1.3"`;

const LO_FONT_FACES =
  `<office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" ` +
  `style:font-family-generic="swiss" style:font-pitch="variable"/><style:font-face style:name="DejaVu Sans" ` +
  `svg:font-family="&apos;DejaVu Sans&apos;" style:font-family-generic="system" style:font-pitch="variable"/></office:font-face-decls>`;

/** The automatic styles LibreOffice writes for these tables: columns and
 * rows, a visible (ta1) and a HIDDEN (ta2) table style, and the data styles
 * the value cells reference. */
const LO_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.3417in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>` +
  `<number:date-style style:name="N37" number:automatic-order="true"><number:month number:style="long"/><number:text>/</number:text>` +
  `<number:day number:style="long"/><number:text>/</number:text><number:year/></number:date-style>` +
  `<number:percentage-style style:name="N11"><number:number number:decimal-places="0" number:min-decimal-places="0" number:min-integer-digits="1"/>` +
  `<number:text>%</number:text></number:percentage-style>` +
  `<number:currency-style style:name="N122P0" style:volatile="true"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>` +
  `<number:currency-style style:name="N122"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N122P0"/></number:currency-style>` +
  `<number:time-style style:name="N41" number:truncate-on-overflow="false"><number:hours number:style="long"/><number:text>:</number:text>` +
  `<number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N11"/>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N122"/>` +
  `<style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N41"/>` +
  `<style:style style:name="ce5" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
  `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `</office:automatic-styles>`;

const loString = (text: string, style?: string): string =>
  `<table:table-cell${style ? ` table:style-name="${style}"` : ""} office:value-type="string" calcext:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
const loFloat = (value: number, formula?: string): string =>
  `<table:table-cell${formula ? ` table:formula="${formula}"` : ""} office:value-type="float" office:value="${value}" calcext:value-type="float"><text:p>${value}</text:p></table:table-cell>`;
const loCurrency = (value: number, text: string, formula?: string): string =>
  `<table:table-cell table:style-name="ce3"${formula ? ` table:formula="${formula}"` : ""} office:value-type="currency" office:currency="USD" ` +
  `office:value="${value}" calcext:value-type="currency"><text:p>${text}</text:p></table:table-cell>`;
const loEmpty = (repeat: number): string => `<table:table-cell table:number-columns-repeated="${repeat}"/>`;
const loRow = (cells: string, style = "ro1", repeat?: number): string =>
  `<table:table-row table:style-name="${style}"${repeat ? ` table:number-rows-repeated="${repeat}"` : ""}>${cells}</table:table-row>`;

function loContent(tables: string): string {
  return (
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${LO_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false" ` +
    `table:use-regular-expressions="false" table:use-wildcards="true"/>${tables}<table:named-expressions/>` +
    `</office:spreadsheet></office:body></office:document-content>`
  );
}

const LO_STYLES =
  ODF_XML_DECLARATION +
  `<office:document-styles ${LO_CONTENT_NAMESPACES}>${LO_FONT_FACES}<office:styles>` +
  `<style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/>` +
  `<style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="en" fo:country="US"/></style:default-style>` +
  `<number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
  `<style:style style:name="Default" style:family="table-cell"/>` +
  `<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-size="24pt" fo:font-weight="bold"/></style:style>` +
  `</office:styles><office:automatic-styles><style:page-layout style:name="Mpm1"><style:page-layout-properties style:writing-mode="lr-tb"/>` +
  `<style:header-style><style:header-footer-properties fo:min-height="0.2953in" fo:margin-bottom="0.0984in"/></style:header-style>` +
  `<style:footer-style><style:header-footer-properties fo:min-height="0.2953in" fo:margin-top="0.0984in"/></style:footer-style></style:page-layout>` +
  `</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="Mpm1">` +
  `<style:header><text:p><text:sheet-name>???</text:sheet-name></text:p></style:header>` +
  `<style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer></style:master-page></office:master-styles></office:document-styles>`;

function loMeta(title: string, tableCount: number, cellCount: number): string {
  return (
    ODF_XML_DECLARATION +
    `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
    `xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.3">` +
    `<office:meta><meta:creation-date>2026-10-01T09:00:00.123456789</meta:creation-date><dc:date>2026-10-08T17:30:00.987654321</dc:date>` +
    `<meta:editing-duration>PT12M34S</meta:editing-duration><meta:editing-cycles>3</meta:editing-cycles>` +
    `<meta:generator>LibreOffice/24.8.2.1$Linux_X86_64 LibreOffice_project/0f794b6e29741098670a3b95d60478a65d05ef13</meta:generator>` +
    `<dc:title>${title}</dc:title><meta:initial-creator>${FIXTURE_WORKBOOK_AUTHOR}</meta:initial-creator><dc:creator>${FIXTURE_WORKBOOK_AUTHOR}</dc:creator>` +
    `<meta:document-statistic meta:table-count="${tableCount}" meta:cell-count="${cellCount}" meta:object-count="0"/></office:meta></office:document-meta>`
  );
}

function loSettings(tables: readonly string[], active: string): string {
  const entries = tables
    .map(
      (name) =>
        `<config:config-item-map-entry config:name="${name}"><config:config-item config:name="CursorPositionX" config:type="int">0</config:config-item>` +
        `<config:config-item config:name="CursorPositionY" config:type="int">0</config:config-item>` +
        `<config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>` +
        `<config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item></config:config-item-map-entry>`,
    )
    .join("");
  return (
    ODF_XML_DECLARATION +
    `<office:document-settings xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" ` +
    `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3">` +
    `<office:settings><config:config-item-set config:name="ooo:view-settings">` +
    `<config:config-item config:name="VisibleAreaTop" config:type="int">0</config:config-item>` +
    `<config:config-item config:name="VisibleAreaLeft" config:type="int">0</config:config-item>` +
    `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    `<config:config-item config:name="ViewId" config:type="string">view1</config:config-item>` +
    `<config:config-item-map-named config:name="Tables">${entries}</config:config-item-map-named>` +
    `<config:config-item config:name="ActiveTable" config:type="string">${active}</config:config-item>` +
    `<config:config-item config:name="HorizontalScrollbarWidth" config:type="int">1309</config:config-item>` +
    `<config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>` +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>` +
    `<config:config-item-set config:name="ooo:configuration-settings">` +
    `<config:config-item config:name="IsDocumentShared" config:type="boolean">false</config:config-item>` +
    `<config:config-item config:name="SaveVersionOnClose" config:type="boolean">false</config:config-item>` +
    `</config:config-item-set></office:settings></office:document-settings>`
  );
}

const LO_MANIFEST =
  ODF_XML_DECLARATION +
  `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3" ` +
  `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">` +
  `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
  `<manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/>` +
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/></manifest:manifest>`;

const LO_MANIFEST_RDF =
  ODF_XML_DECLARATION +
  `<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about="content.xml">` +
  `<rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/></rdf:Description>` +
  `<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="content.xml"/>` +
  `</rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>` +
  `</rdf:Description></rdf:RDF>\n`;

/** A 2×2 PNG — the thumbnail LibreOffice stores in every package. */
const THUMBNAIL_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0, 0, 0, 253, 212, 154,
  115, 0, 0, 0, 18, 73, 68, 65, 84, 120, 156, 99, 48, 238, 56, 99, 220, 113, 134, 1, 66, 1, 0, 40, 110, 6, 29, 254, 136,
  5, 56, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

/** Assembles a LibreOffice-shaped package in Calc's own entry order, with
 * `mimetype` first and STOREd. */
function loPackage(content: string, meta: string, settings: string): Promise<Uint8Array> {
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    part("meta.xml", meta),
    part("settings.xml", settings),
    part("content.xml", content),
    { name: "Thumbnails/thumbnail.png", data: THUMBNAIL_PNG },
    part("styles.xml", LO_STYLES),
    part("manifest.rdf", LO_MANIFEST_RDF),
    part("META-INF/manifest.xml", LO_MANIFEST),
  ]);
}

/** What mini-book.ods holds, for the specs. */
export const MINI_BOOK_ODS = {
  tableNames: ["Archive", "Inventory", "Totals"],
  activeTable: "Inventory",
  /** Inventory's used rows: 23 + the trailing 1048553-row run = 1048576. */
  inventoryRows: 23,
  /** Inventory's cells (A1 notation) → the value the reader must produce. */
  inventory: {
    A1: "Item",
    D1: "Total",
    A2: "Widget",
    B2: 4,
    C2: 2.5,
    D2: 10,
    D3: 12.75,
    A4: "Spare",
    A5: "Spare",
    A6: "Spare",
    D6: 0.5,
    B7: 0.15,
    B8: dateToSerial(2026, 10, 8, "1900"),
    C8: (13 * 60 + 45) / (24 * 60),
    B9: true,
    C9: false,
    B10: { error: "#DIV/0!" },
    A11: "Merged note",
    C11: 7,
    A12: "First  line\ttabbed\nSecond\nthird",
    B18: "Widget x4",
    B19: 123456789012,
    B20: -42.5,
    B21: (36 * 60 + 5) / (24 * 60),
    B22: 61,
    A23: "Done",
  },
} as const;

export async function buildMiniBookOdsBytes(): Promise<Uint8Array> {
  const archive =
    `<table:table table:name="Archive" table:style-name="ta2">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString("Old data") + loFloat(2019) + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1048575) +
    `</table:table>`;

  const inventory =
    `<table:table table:name="Inventory" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="4" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16380" table:default-cell-style-name="Default"/>` +
    `<table:table-header-rows>` +
    loRow(loString("Item", "ce5") + loString("Qty", "ce5") + loString("Price", "ce5") + loString("Total", "ce5") + loEmpty(16380)) +
    `</table:table-header-rows>` +
    loRow(loString("Widget") + loFloat(4) + loCurrency(2.5, "$2.50") + loCurrency(10, "$10.00", "of:=[.B2]*[.C2]") + loEmpty(16380)) +
    loRow(loString("Gadget") + loFloat(3) + loCurrency(4.25, "$4.25") + loCurrency(12.75, "$12.75", "of:=[.B3]*[.C3]") + loEmpty(16380)) +
    // A repeated NON-empty row: materialised three times.
    loRow(loString("Spare") + loFloat(1) + loCurrency(0.5, "$0.50") + loCurrency(0.5, "$0.50") + loEmpty(16380), "ro1", 3) +
    loRow(
      loString("Discount") +
        `<table:table-cell table:style-name="ce2" office:value-type="percentage" office:value="0.15" calcext:value-type="percentage"><text:p>15%</text:p></table:table-cell>` +
        loEmpty(16382),
    ) +
    loRow(
      loString("Delivered") +
        `<table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="2026-10-08" calcext:value-type="date"><text:p>10/08/2026</text:p></table:table-cell>` +
        `<table:table-cell table:style-name="ce4" office:value-type="time" office:time-value="PT13H45M00S" calcext:value-type="time"><text:p>13:45:00</text:p></table:table-cell>` +
        loEmpty(16381),
    ) +
    loRow(
      loString("In stock") +
        `<table:table-cell office:value-type="boolean" office:boolean-value="true" calcext:value-type="boolean"><text:p>TRUE</text:p></table:table-cell>` +
        `<table:table-cell office:value-type="boolean" office:boolean-value="false" calcext:value-type="boolean"><text:p>FALSE</text:p></table:table-cell>` +
        loEmpty(16381),
    ) +
    loRow(
      loString("Ratio") +
        `<table:table-cell table:formula="of:=1/0" office:value-type="float" office:value="0" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>` +
        loEmpty(16382),
    ) +
    // Covered cells under a two-column span.
    loRow(
      `<table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="1" office:value-type="string" calcext:value-type="string">` +
        `<text:p>Merged note</text:p></table:table-cell><table:covered-table-cell/>` +
        loFloat(7) +
        loEmpty(16381),
      "ro2",
    ) +
    loRow(
      `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>First<text:s text:c="2"/>line<text:tab/>tabbed</text:p>` +
        `<text:p>Second<text:line-break/>third</text:p></table:table-cell>` +
        loEmpty(16383),
      "ro2",
    ) +
    // Five empty rows in the middle: advanced past, never materialised.
    loRow(loEmpty(16384), "ro1", 5) +
    loRow(
      loString("Note") +
        `<table:table-cell table:formula="of:=CONCATENATE([.A2];&quot; x&quot;;[.B2])" office:value-type="string" ` +
        `office:string-value="Widget x4" calcext:value-type="string"><text:p>Widget x4</text:p></table:table-cell>` +
        loEmpty(16382),
    ) +
    loRow(loString("Big") + loFloat(123456789012) + loEmpty(16382)) +
    loRow(loString("Negative") + loFloat(-42.5) + loEmpty(16382)) +
    loRow(
      loString("Long time") +
        `<table:table-cell table:style-name="ce4" office:value-type="time" office:time-value="PT36H05M00S" calcext:value-type="time"><text:p>36:05:00</text:p></table:table-cell>` +
        loEmpty(16382),
    ) +
    loRow(
      loString("Leap") +
        `<table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="1900-03-01" calcext:value-type="date"><text:p>03/01/1900</text:p></table:table-cell>` +
        loEmpty(16382),
    ) +
    loRow(loString("Done") + loEmpty(16383)) +
    // LibreOffice's trailing runs to the grid's edge.
    loRow(loEmpty(16384), "ro1", 1048553) +
    `</table:table>`;

  const totals =
    `<table:table table:name="Totals" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString("Total value") + loCurrency(24.25, "$24.25", "of:=SUM([Inventory.D2:.D6])") + loEmpty(16382)) +
    loRow(loString("Items") + loFloat(5, "of:=COUNTA([Inventory.A2:.A6])") + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1048574) +
    `</table:table>`;

  return loPackage(
    loContent(archive + inventory + totals),
    loMeta("Mini book", 3, 64),
    loSettings(MINI_BOOK_ODS.tableNames, MINI_BOOK_ODS.activeTable),
  );
}

/** spec116-wide.ods: one cell far out — column 16,000 of row 900,000. */
export const SPEC116_WIDE = {
  tableName: "Wide",
  /** 0-based coordinates of the far cell. */
  row: 899_999,
  col: 15_999,
  value: 900000.16,
} as const;

export async function buildSpecWideOdsBytes(): Promise<Uint8Array> {
  const table =
    `<table:table table:name="${SPEC116_WIDE.tableName}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString("Corner") + loEmpty(16383)) +
    loRow(loEmpty(16384), "ro1", SPEC116_WIDE.row - 1) +
    loRow(loEmpty(SPEC116_WIDE.col) + loFloat(SPEC116_WIDE.value) + loEmpty(16384 - SPEC116_WIDE.col - 1)) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - SPEC116_WIDE.row - 1) +
    `</table:table>`;
  return loPackage(loContent(table), loMeta("Wide", 1, 2), loSettings([SPEC116_WIDE.tableName], SPEC116_WIDE.tableName));
}
