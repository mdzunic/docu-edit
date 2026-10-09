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
// - spec117-styles.xlsx (SPEC-117) — Excel 365 on the classic Office theme
//   (Calibri): a styles part with every font, fill (pattern, tinted, indexed
//   with an `<indexedColors>` override, gradient), border (all thirteen
//   styles and a diagonal), alignment, protection and quotePrefix construct,
//   named styles, a dxf and Excel's extLst; sheets "Styled" (custom, best-fit,
//   hidden and styled columns, custom-height, hidden and styled rows,
//   style-only cells, a merge, a frozen pane, 85% zoom, a two-range
//   selection, a theme tab colour, a rich string), "Plain" (no gridlines) and
//   "Split" (a split, not frozen, pane). SPEC117_STYLES_XLSX records the facts.
// - spec118-formats.xlsx (SPEC-118) — Excel 365: `numFmts` with custom ids
//   ≥ 164 ([$-F800] long date, [$€-407], red parentheses, # ??/??, a
//   conditional pair, `;;;`, d-mmm-yyyy) and the explicit 43/44 Excel always
//   writes; cells on built-ins 4, 9, 10, 11, 14, 22, 46 and 49; A2:A4 the
//   same date-time under 14, 22 and the F800 code; serials 0, 60, 61; a
//   narrow F (60 px — General drops to "1.23E+11") and G (30 px — a date
//   shows only #). SPEC118_FORMATS_XLSX records the facts.
// - spec118-generated.xlsx — openpyxl-shaped: inline strings, `t="d"` ISO
//   dates under a General style, no theme, no numFmts (implied formats).
// - spec118-formats.ods — LibreOffice-shaped data styles: N2/N2P0 (a red
//   negative through style:map), a de-DE currency, a German long textual
//   date, an elapsed time style, fraction, scientific, percentage, boolean,
//   text, quarter + year, week-of-year and era styles, and unstyled date,
//   time and percentage cells — 1899-12-25 among them (an ODS-only date).
// - spec119-formulas.xlsx (SPEC-119) — Excel 365-shaped: sheet "Calc" with
//   a normal formula, a shared block in Excel's order (master first), a
//   shared block as a non-Excel producer writes it (a follower before its
//   master, the last follower's shift running off the grid → #REF!), a CSE
//   array, `_xlfn.` and `_xlfn._xlws.` functions, a dynamic array with
//   `cm="1"` and xl/metadata.xml, LET with `_xlpm.` parameters, SINGLE, a
//   data table, a 3-D sum over Jan:Mar, 'O''Brien'!A1, A:A and 1:1, an
//   external reference with xl/externalLinks/externalLink1.xml, Sheet2!#REF!,
//   a multi-line formula with indentation and the immediately-invoked LAMBDA
//   Excel writes but this grammar does not parse; "Jan" (the second sheet)
//   has `showFormulas="1"`. SPEC119_FORMULAS_XLSX records the facts.
// - spec119-formulas.ods (SPEC-119) — LibreOffice-shaped, `xmlns:of`,
//   `xmlns:oooc` and `xmlns:msoxl` on the root: bracketed references (another
//   sheet's quoted name included), `~` and `!`, an inline array,
//   COM.MICROSOFT.CONCAT, a matrix formula with both span attributes, a
//   repeated formula cell, an `oooc:=` and an `msoxl:=` formula and a formula
//   cell without a cached value. SPEC119_FORMULAS_ODS records the facts.
// - spec120-calc.xlsx (SPEC-120) — Excel-shaped: "Calc" with a SUM cached
//   999 (wrong on purpose: the cache is trusted until F9), TODAY() cached
//   with an old date, an `_xll.` add-in call and the self-contained
//   conformance rows with Excel's values; "Loop", a two-cell cycle; no
//   fullCalcOnLoad. spec120-fullcalc.xlsx — XlsxWriter-shaped, every formula
//   cached 0 and `fullCalcOnLoad="1"`. spec120-manual.xlsx — Excel-shaped,
//   `calcMode="manual"`. spec120-calc.ods — LibreOffice-shaped, explicit
//   `table:calculation-settings` (case-sensitive, wildcards), AutoCalculate in
//   settings.xml and two formula cells without cached values.
// spec116-numbers.xlsx and spec120-numbers.xlsx are NOT built here — Apple
// Numbers exported them (see e2e/fixtures/spec116-numbers.md and
// spec120-numbers.md); the specs read the committed files.
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
  /** SPEC-120: the workbook's `<calcPr>` (default Excel's `calcId` alone). */
  calcPr?: string;
  /** Extra parts (name → bytes), their workbook relationship and content
   * type declaration — the VBA project. */
  extra?: { name: string; data: Uint8Array; relationship: string; contentType: string };
  /** The theme and styles parts (default: EXCEL_THEME and EXCEL_STYLES —
   * every SPEC-116 fixture's bytes). */
  theme?: string;
  styles?: string;
  /** SPEC-119: further parts (after the shared strings), their workbook
   * relationships (with their own ids), content type overrides, and what the
   * workbook part holds between `</sheets>` and `<calcPr>`. */
  additions?: { parts: ZipEntryInput[]; relationships: string[]; overrides: string[]; workbookAfterSheets?: string };
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
    ...(options.additions?.overrides ?? []),
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
    ...(options.additions?.relationships ?? []),
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
    `<sheets>${sheetsXml}</sheets>${options.additions?.workbookAfterSheets ?? ""}${options.calcPr ?? `<calcPr calcId="191029"/>`}</workbook>`;

  const entries: ZipEntryInput[] = [
    part("[Content_Types].xml", contentTypes),
    part("_rels/.rels", rootRels),
    part("xl/workbook.xml", workbook),
    part("xl/_rels/workbook.xml.rels", XML_DECLARATION + `<Relationships xmlns="${PKG_REL_NS}">${workbookRels.join("")}</Relationships>`),
    ...options.sheets.map((sheet, index) => part(`xl/worksheets/sheet${index + 1}.xml`, sheet.xml)),
    part("xl/theme/theme1.xml", options.theme ?? EXCEL_THEME),
    part("xl/styles.xml", options.styles ?? EXCEL_STYLES),
  ];
  if (options.sharedStrings) {
    entries.push(part("xl/sharedStrings.xml", options.sharedStrings));
  }
  entries.push(...(options.additions?.parts ?? []));
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

// --- spec117-styles.xlsx (SPEC-117 "Fixtures") --------------------------------
// Excel 365 on the classic Office theme (Calibri 11, defaultThemeVersion
// 166925): every construct SPEC-117 reads, in the shape Excel writes it.

/** The classic Office theme (2013–2022): Calibri / Calibri Light and exactly
 * the scheme core/theme.ts's DEFAULT_THEME carries. */
const SPEC117_THEME =
  XML_DECLARATION +
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
  `<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>` +
  `<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` +
  `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` +
  `<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"><a:shade val="98000"/><a:satMod val="102000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>` +
  `<a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">` +
  `<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" ` +
  `id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`;

/** The thirteen SpreadsheetML border styles, in ST_BorderStyle order — one
 * border (1–13), one xf (SPEC117_XF.borders + k) and one row-16 cell each. */
export const SPEC117_BORDER_STYLES = [
  "thin",
  "medium",
  "dashed",
  "dotted",
  "thick",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
] as const;

/** Each border style's colour element: Excel's automatic `indexed="64"`,
 * with a theme colour, an rgb and an `auto="1"` among them. */
const SPEC117_BORDER_COLORS: Partial<Record<(typeof SPEC117_BORDER_STYLES)[number], string>> = {
  medium: `<color theme="4"/>`,
  thick: `<color rgb="FFFF0000"/>`,
  hair: `<color auto="1"/>`,
};

/** spec117-styles.xlsx's cellXfs indices — a cell's `s`. */
export const SPEC117_XF = {
  normal: 0,
  header: 1, // bold red (FFFF0000) on solid yellow (FFFFFF00)
  title: 2, // Heading 1, centred
  trellis: 3, // darkTrellis fill
  tinted: 4, // solid accent1, lighter 40%
  gradient: 5, // linear gradient, 90°
  indexed: 6, // indexed font colour + an overridden indexed fill
  wrap: 7,
  rotate90: 8,
  indent: 9,
  diagonal: 10,
  quotePrefix: 11,
  italicTint: 12, // italic accent1, darker 25%
  doubleUnderline: 13,
  good: 14, // the "Good" named style
  column: 15, // columns J:L — 0.0% on the tint fill
  row: 16, // row 18 (customFormat)
  borders: 17, // 17 + k: SPEC117_BORDER_STYLES[k] on all four edges
  unlocked: 30,
  shrink: 31,
  stacked: 32, // textRotation 255
  strike: 33, // strike + superscript
} as const;

/** The 64-entry legacy palette as Excel writes `<indexedColors>`, entry 44
 * (#99ccff in the legacy table) overridden to light pink. */
const SPEC117_INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "FFB6C1", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

const SPEC117_FONT_TAIL = `<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`;

const SPEC117_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>` +
  `<fonts count="8" x14ac:knownFonts="1">` +
  `<font><sz val="11"/><color theme="1"/>${SPEC117_FONT_TAIL}` + // 0 Normal
  `<font><b/><sz val="11"/><color rgb="FFFF0000"/>${SPEC117_FONT_TAIL}` + // 1 bold red
  `<font><i/><sz val="11"/><color theme="4" tint="-0.249977111117893"/>${SPEC117_FONT_TAIL}` + // 2 italic accent1 darker 25%
  `<font><u val="double"/><sz val="11"/><color theme="1"/>${SPEC117_FONT_TAIL}` + // 3 double underline
  `<font><b/><sz val="15"/><color theme="3"/>${SPEC117_FONT_TAIL}` + // 4 Heading 1
  `<font><sz val="11"/><color indexed="10"/>${SPEC117_FONT_TAIL}` + // 5 indexed red
  `<font><sz val="11"/><color rgb="FF006100"/>${SPEC117_FONT_TAIL}` + // 6 Good
  `<font><strike/><vertAlign val="superscript"/><sz val="11"/><color theme="1"/>${SPEC117_FONT_TAIL}` + // 7 strike superscript
  `</fonts>` +
  `<fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.39997558519241921"/><bgColor indexed="64"/></patternFill></fill>` +
  `<fill><patternFill patternType="darkTrellis"><fgColor theme="5"/><bgColor rgb="FFFFFFCC"/></patternFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor indexed="44"/><bgColor indexed="64"/></patternFill></fill>` +
  `<fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill></fills>` +
  `<borders count="16"><border><left/><right/><top/><bottom/><diagonal/></border>` +
  SPEC117_BORDER_STYLES.map((style) => {
    const color = SPEC117_BORDER_COLORS[style] ?? `<color indexed="64"/>`;
    const edge = (name: string) => `<${name} style="${style}">${color}</${name}>`;
    return `<border>${edge("left")}${edge("right")}${edge("top")}${edge("bottom")}<diagonal/></border>`;
  }).join("") +
  `<border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border>` + // 14 Heading 1
  `<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border>` + // 15
  `</borders>` +
  `<cellStyleXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="0" fontId="6" fillId="7" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="0" fontId="4" fillId="0" borderId="14" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
  `<cellXfs count="34"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="4" fillId="0" borderId="14" xfId="2" applyAlignment="1"><alignment horizontal="center"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="5" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="2"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="15" xfId="0" applyBorder="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<xf numFmtId="0" fontId="6" fillId="7" borderId="0" xfId="1"/>` +
  `<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>` +
  SPEC117_BORDER_STYLES.map((_, k) => `<xf numFmtId="0" fontId="0" fillId="0" borderId="${k + 1}" xfId="0" applyBorder="1"/>`).join("") +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" shrinkToFit="1"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" textRotation="255"/></xf>` +
  `<xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="3"><cellStyle name="Good" xfId="1" builtinId="26"/><cellStyle name="Heading 1" xfId="2" builtinId="16"/>` +
  `<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<colors><indexedColors>${SPEC117_INDEXED_COLORS.map((rgb) => `<rgbColor rgb="00${rgb}"/>`).join("")}</indexedColors>` +
  `<mruColors><color rgb="FFFFB6C1"/><color rgb="FF66FF33"/></mruColors></colors>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
  `<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>` +
  `<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
  `<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;

/** spec117-styles.xlsx's plain shared strings, in index order after the rich
 * item 0 (A1). */
const SPEC117_STRINGS = [
  "Header",
  "Merged title",
  "Tinted",
  "Gradient",
  "Indexed",
  "Good",
  "Wrapped text that runs over several lines",
  "Rotated",
  "Indented",
  "Italic accent",
  "Diagonal",
  "Double underline",
  "Shrink this text to fit",
  "Stacked",
  "Unlocked",
  "Strike",
  "hidden row",
  ...SPEC117_BORDER_STYLES,
  "007",
  "Styled row",
  "Name",
  "Score",
  "Ada",
  "Grace",
  "Split panes",
];

/** spec117-styles.xlsx's rich shared string (item 0): two runs, two fonts. */
const SPEC117_RICH_STRING =
  `<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">Styled </t></r>` +
  `<r><rPr><i/><sz val="14"/><color rgb="FFC00000"/><rFont val="Cambria"/><family val="1"/></rPr><t>workbook</t></r></si>`;

/** A shared-string cell for one of SPEC117_STRINGS. */
function spec117Text(ref: string, text: string, style?: number): string {
  const index = SPEC117_STRINGS.indexOf(text);
  if (index < 0) {
    throw new Error(`"${text}" is not one of spec117-styles.xlsx's strings`);
  }
  return sharedCell(ref, index + 1, style);
}

/** A style-only cell, as Excel writes a formatted empty cell. */
const styleOnlyCell = (ref: string, style: number): string => `<c r="${ref}" s="${style}"/>`;

/** One Calibri-era row in Excel's attribute order (r, spans, s, customFormat,
 * ht, hidden, customHeight, …, x14ac:dyDescent). `spans` covers the row's
 * 16-row block: A:Z for rows 1–16, A:A below. */
function spec117Row(r: number, attrs: string, cells: string): string {
  return `<row r="${r}" spans="${r <= 16 ? "1:26" : "1:1"}"${attrs} x14ac:dyDescent="0.25">${cells}</row>`;
}

/** The columns row 16's border cells sit in (every other one, skipping the
 * hidden G, so each cell's four edges show). */
const SPEC117_BORDER_COLUMNS = ["A", "C", "E", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z"] as const;

/** How many shared-string cells (`t="s"`) a worksheet part holds — the sst's `count`. */
function sharedStringReferences(worksheetPart: string): number {
  return worksheetPart.split(` t="s"`).length - 1;
}

const SPEC117_PAGE_MARGINS = `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`;
const SPEC117_FORMAT_PR = `<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>`;

/** What spec117-styles.xlsx holds, for the unit and e2e specs. */
export const SPEC117_STYLES_XLSX = {
  sheetNames: ["Styled", "Plain", "Split"],
  activeSheetIndex: 0,
  themeName: "Office Theme",
  /** The Normal font — Calibri 11, so the maximum digit width is 7 px. */
  defaultFont: { family: "Calibri", sizePt: 11, mdw: 7 },
  /** B2: the shared string "Header" — bold red text on a solid yellow fill. */
  header: { ref: "B2", value: "Header", styleId: SPEC117_XF.header, color: "#ff0000", fill: "#ffff00" },
  /** B3:B12 = 10, 20, …, 100; nothing in column B below row 12. */
  columnB: { first: "B3", last: "B12", values: [10, 20, 30, 40, 50, 60, 70, 80, 90, 100] },
  /** B3:B5 = 10, 20, 30: the status-bar aggregates of that selection. */
  b3b5: { range: "B3:B5", average: 20, count: 3, sum: 60 },
  /** C5:E5 merged; C5 "Merged title" (Heading 1, centred), D5 and E5
   * style-only members. */
  merge: { ref: "C5:E5", range: { top: 4, left: 2, bottom: 4, right: 4 }, value: "Merged title", styleId: SPEC117_XF.title },
  /** D3: a style-only cell with the darkTrellis fill. */
  styleOnly: { ref: "D3", styleId: SPEC117_XF.trellis },
  /** Column G (index 6) is hidden — G2 holds a value; F and H are visible. */
  hiddenColumn: { index: 6, letter: "G", markerTitle: "Hidden column G", valueRef: "G2", value: 7 },
  /** Row 14 (index 13) is hidden and holds A14 = "hidden row". */
  hiddenRow: { index: 13, ref: "A14", value: "hidden row" },
  /** Rows with `ht` + `customHeight`: row 2 (20.25 pt) and row 5 (24 pt). */
  rowHeightsPt: { 2: 20.25, 5: 24 },
  /** Column widths in px for Calibri 11 (MDW 7); the default is 64 px. */
  columnWidthsPx: { A: 131, B: 89, F: 97 },
  /** Columns J:L (indices 9–11) carry style 15. */
  styledColumns: { min: 9, max: 11, styleId: SPEC117_XF.column },
  /** Row 18 (index 17): s="16" customFormat="1". */
  styledRow: { index: 17, styleId: SPEC117_XF.row },
  /** A17: "007" entered with a leading apostrophe. */
  quotePrefix: { ref: "A17", value: "007", styleId: SPEC117_XF.quotePrefix },
  /** Row 16: one cell per border style, in SPEC117_BORDER_STYLES order. */
  borderRow: { index: 15, columns: SPEC117_BORDER_COLUMNS },
  frozen: { rows: 2, cols: 1 },
  zoomPercent: 85,
  topLeftCell: "B3",
  /** The bottomRight pane's selection (the active pane). */
  selection: {
    active: "B3",
    sqref: "B3:B5 D7:E8",
    ranges: [
      { top: 2, left: 1, bottom: 4, right: 1 },
      { top: 6, left: 3, bottom: 7, right: 4 },
    ],
    activeRangeIndex: 0,
  },
  /** sheetPr/tabColor theme="5" (accent2) tint="0.39997558519241921". */
  tabColorTint: "0.39997558519241921",
  /** The rich shared string at A1. */
  rich: { ref: "A1", value: "Styled workbook", runs: ["Styled ", "workbook"] },
  /** `<indexedColors>` entry 44 differs from the legacy palette. */
  indexedOverride: { index: 44, color: "#ffb6c1" },
  /** "Plain": showGridLines="0". */
  plain: { showGridLines: false, selection: "B2" },
  /** "Split": a split (not frozen) pane, its active pane's selection D6. */
  split: { topLeftCell: "C5", selection: "D6" },
} as const;

export async function buildSpec117StylesXlsxBytes(): Promise<Uint8Array> {
  const X = SPEC117_XF;
  const columnB = SPEC117_STYLES_XLSX.columnB.values;
  const rows = [
    spec117Row(1, "", sharedCell("A1", 0)),
    spec117Row(2, ` ht="20.25" customHeight="1"`, spec117Text("B2", "Header", X.header) + numberCell("G2", 7)),
    spec117Row(3, "", numberCell("B3", columnB[0]) + styleOnlyCell("D3", X.trellis)),
    spec117Row(4, "", numberCell("B4", columnB[1])),
    spec117Row(
      5,
      ` ht="24" customHeight="1"`,
      numberCell("B5", columnB[2]) + spec117Text("C5", "Merged title", X.title) + styleOnlyCell("D5", X.title) + styleOnlyCell("E5", X.title),
    ),
    spec117Row(6, "", numberCell("B6", columnB[3])),
    spec117Row(
      7,
      "",
      numberCell("B7", columnB[4]) + spec117Text("C7", "Tinted", X.tinted) + spec117Text("D7", "Gradient", X.gradient) +
        spec117Text("E7", "Indexed", X.indexed) + spec117Text("F7", "Good", X.good),
    ),
    spec117Row(
      8,
      "",
      numberCell("B8", columnB[5]) + spec117Text("C8", "Wrapped text that runs over several lines", X.wrap) +
        spec117Text("D8", "Rotated", X.rotate90) + spec117Text("E8", "Indented", X.indent) + spec117Text("F8", "Italic accent", X.italicTint),
    ),
    spec117Row(
      9,
      "",
      numberCell("B9", columnB[6]) + spec117Text("C9", "Diagonal", X.diagonal) + spec117Text("D9", "Double underline", X.doubleUnderline) +
        spec117Text("E9", "Shrink this text to fit", X.shrink) + spec117Text("F9", "Stacked", X.stacked),
    ),
    spec117Row(10, "", numberCell("B10", columnB[7]) + spec117Text("C10", "Unlocked", X.unlocked) + spec117Text("D10", "Strike", X.strike)),
    spec117Row(11, "", numberCell("B11", columnB[8]) + numberCell("C11", 0.256, X.column)),
    spec117Row(12, "", numberCell("B12", columnB[9])),
    spec117Row(14, ` hidden="1"`, spec117Text("A14", "hidden row")),
    spec117Row(
      16,
      "",
      SPEC117_BORDER_STYLES.map((style, k) => spec117Text(`${SPEC117_BORDER_COLUMNS[k]}16`, style, X.borders + k)).join(""),
    ),
    spec117Row(17, "", spec117Text("A17", "007", X.quotePrefix)),
    spec117Row(18, ` s="${X.row}" customFormat="1"`, spec117Text("A18", "Styled row", X.row)),
  ].join("");
  const styled =
    XML_DECLARATION +
    worksheetOpen("A3C1E5F7-9B2D-4F6A-8C1E-3D5F7A9B1C17") +
    `<sheetPr><tabColor theme="5" tint="${SPEC117_STYLES_XLSX.tabColorTint}"/></sheetPr><dimension ref="A1:Z18"/>` +
    `<sheetViews><sheetView tabSelected="1" zoomScale="85" zoomScaleNormal="85" workbookViewId="0">` +
    `<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>` +
    `<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/>` +
    `<selection pane="bottomRight" activeCell="B3" sqref="B3:B5 D7:E8"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<cols><col min="1" max="1" width="18.7109375" customWidth="1"/><col min="2" max="2" width="12.7109375" customWidth="1"/>` +
    `<col min="6" max="6" width="13.8515625" bestFit="1" customWidth="1"/><col min="7" max="7" width="9.140625" hidden="1"/>` +
    `<col min="10" max="12" width="9.140625" style="${X.column}"/></cols>` +
    `<sheetData>${rows}</sheetData><mergeCells count="1"><mergeCell ref="C5:E5"/></mergeCells>${SPEC117_PAGE_MARGINS}</worksheet>`;

  const plain =
    XML_DECLARATION +
    worksheetOpen("B4D2F6A8-0C3E-4A7B-9D2F-4E6A8B0C2D27") +
    `<dimension ref="A1:B3"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<sheetData>` +
    `<row r="1" spans="1:2" x14ac:dyDescent="0.25">${spec117Text("A1", "Name")}${spec117Text("B1", "Score")}</row>` +
    `<row r="2" spans="1:2" x14ac:dyDescent="0.25">${spec117Text("A2", "Ada")}${numberCell("B2", 95)}</row>` +
    `<row r="3" spans="1:2" x14ac:dyDescent="0.25">${spec117Text("A3", "Grace")}${numberCell("B3", 88)}</row>` +
    `</sheetData>${SPEC117_PAGE_MARGINS}</worksheet>`;

  const split =
    XML_DECLARATION +
    worksheetOpen("C5E3A7B9-1D4F-4B8C-8E3A-5F7B9C1D3E37") +
    `<dimension ref="A1:D6"/><sheetViews><sheetView workbookViewId="0">` +
    `<pane xSplit="2010" ySplit="1500" topLeftCell="C5" activePane="bottomRight" state="split"/>` +
    `<selection pane="topRight" activeCell="C1" sqref="C1"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/>` +
    `<selection pane="bottomRight" activeCell="D6" sqref="D6"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<sheetData>` +
    `<row r="1" spans="1:4" x14ac:dyDescent="0.25">${spec117Text("A1", "Split panes")}${numberCell("C1", 1)}</row>` +
    `<row r="5" spans="1:4" x14ac:dyDescent="0.25">${numberCell("A5", 2)}</row>` +
    `<row r="6" spans="1:4" x14ac:dyDescent="0.25">${numberCell("D6", 42)}</row>` +
    `</sheetData>${SPEC117_PAGE_MARGINS}</worksheet>`;

  const count = sharedStringReferences(styled) + sharedStringReferences(plain) + sharedStringReferences(split);
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${count}" uniqueCount="${SPEC117_STRINGS.length + 1}">` +
    SPEC117_RICH_STRING +
    SPEC117_STRINGS.map((text) => `<si><t>${text}</t></si>`).join("") +
    `</sst>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Styles",
    sheets: [
      { name: "Styled", xml: styled },
      { name: "Plain", xml: plain },
      { name: "Split", xml: split },
    ],
    sharedStrings,
    workbookPr: `<workbookPr defaultThemeVersion="166925"/>`,
    theme: SPEC117_THEME,
    styles: SPEC117_STYLES,
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

/** The `table:calculation-settings` Calc 24.8 writes for its defaults. */
const LO_CALCULATION_SETTINGS = `<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>`;

function loContent(tables: string, namespaces = LO_CONTENT_NAMESPACES, calculationSettings = LO_CALCULATION_SETTINGS): string {
  return (
    ODF_XML_DECLARATION +
    `<office:document-content ${namespaces}><office:scripts/>${LO_FONT_FACES}${LO_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${calculationSettings}${tables}<table:named-expressions/>` +
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

/** settings.xml as Calc writes it; `configuration` adds items to
 * `ooo:configuration-settings` (SPEC-120's AutoCalculate). */
function loSettings(tables: readonly string[], active: string, configuration = ""): string {
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
    configuration +
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

// ============================================================================
// SPEC-117: spec117-styles.ods (LibreOffice 24-shaped)
// ============================================================================
// Every construct SPEC-117 reads from an .ods, in the shape Calc writes it:
// `co…`/`ro…`/`ta…`/`ce…`/`T…` automatic styles in Calc's order (`ta1` with
// `tableooo:tab-color`, every `ce…` parented to "Default" but the merged
// title's, parented to the common "Heading"); styles.xml's table-cell
// default-style (Liberation Sans 10pt) and "Default"/"Heading" common styles
// (LO_STYLES — Calc writes no table-column default-style); column runs with
// `table:default-cell-style-name` "Default" and a `ce…` name, a wider column
// and column G `table:visibility="collapse"`; rows with custom heights
// (`use-optimal-row-height="false"`) and a hidden one; a whole-row format
// written as a styled `number-columns-repeated="16380"` run to XFD; a 3×1
// span over styled covered cells; a `text:span` run; borders, wrap,
// rotation, repeat-content, indent, vertical alignment, shrink-to-fit and a
// diagonal; the trailing `number-rows-repeated` run to row 1,048,576; and
// settings.xml with both split modes 2 on "Styled", ShowGrid false on
// "Plain", a split (mode 1, not a freeze) on "Split", zoom 85 and a cursor.

/** The automatic styles Calc writes for spec117-styles.ods. */
const SPEC117_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="1.2in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.3in" fo:break-before="auto" style:use-optimal-row-height="false"/></style:style>` +
  `<style:style style:name="ro3" style:family="table-row"><style:table-row-properties style:row-height="0.45in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ro4" style:family="table-row"><style:table-row-properties style:row-height="0.6in" fo:break-before="auto" style:use-optimal-row-height="false"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb" tableooo:tab-color="#729fcf"/></style:style>` +
  `<style:style style:name="ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ffff00"/>` +
  `<style:text-properties fo:color="#ff0000" fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border="0.74pt solid #000000"/></style:style>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border="1.75pt solid #000000"/></style:style>` +
  `<style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border="2.5pt solid #000000"/></style:style>` +
  `<style:style style:name="ce5" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border="0.74pt dashed #000000"/></style:style>` +
  `<style:style style:name="ce6" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:border="0.74pt fine-dashed #000000"/></style:style>` +
  `<style:style style:name="ce7" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:border-line-width="0.0102in 0.0071in 0.0102in" ` +
  `fo:border="1.1pt double #000000"/></style:style>` +
  `<style:style style:name="ce8" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:wrap-option="wrap"/></style:style>` +
  `<style:style style:name="ce9" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic" ` +
  `style:font-style-asian="italic" style:font-style-complex="italic"/></style:style>` +
  `<style:style style:name="ce10" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#dddddd"/>` +
  `<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `<style:style style:name="ce11" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:rotation-angle="90" style:rotation-align="none"/></style:style>` +
  `<style:style style:name="ce12" style:family="table-cell" style:parent-style-name="Heading"><style:table-cell-properties style:text-align-source="fix" style:repeat-content="false"/>` +
  `<style:paragraph-properties fo:text-align="center" fo:margin-left="0in"/></style:style>` +
  `<style:style style:name="ce13" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:text-align-source="fix" style:repeat-content="true"/>` +
  `<style:paragraph-properties fo:margin-left="0in"/></style:style>` +
  `<style:style style:name="ce14" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:text-align-source="fix" style:repeat-content="false"/>` +
  `<style:paragraph-properties fo:text-align="start" fo:margin-left="0.1874in"/></style:style>` +
  `<style:style style:name="ce15" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:vertical-align="middle" style:shrink-to-fit="true"/></style:style>` +
  `<style:style style:name="ce16" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:diagonal-bl-tr="0.74pt solid #ff0000" ` +
  `style:diagonal-tl-br="none" fo:border-bottom="1.75pt dashed #0000ff" fo:border-left="none" fo:border-right="none" fo:border-top="none"/></style:style>` +
  `<style:style style:name="T1" style:family="text"><style:text-properties fo:color="#c9211e" fo:font-weight="bold" style:font-weight-asian="bold" ` +
  `style:font-weight-complex="bold"/></style:style>` +
  `</office:automatic-styles>`;

function spec117OdsContent(tables: string): string {
  return (
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC117_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false" ` +
    `table:use-regular-expressions="false" table:use-wildcards="true"/>${tables}<table:named-expressions/>` +
    `</office:spreadsheet></office:body></office:document-content>`
  );
}

const loConfigItem = (name: string, type: string, value: string | number | boolean): string =>
  `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`;

/** One table's entry in settings.xml's Tables map, as Calc writes it. */
function spec117TableSettings(
  name: string,
  view: {
    cursor: [number, number];
    split: [number, number];
    splitPosition: [number, number];
    activeSplitRange: number;
    position: { left: number; right: number; top: number; bottom: number };
    zoom: number;
    showGrid: boolean;
  },
): string {
  return (
    `<config:config-item-map-entry config:name="${name}">` +
    loConfigItem("CursorPositionX", "int", view.cursor[0]) +
    loConfigItem("CursorPositionY", "int", view.cursor[1]) +
    loConfigItem("HorizontalSplitMode", "short", view.split[0]) +
    loConfigItem("VerticalSplitMode", "short", view.split[1]) +
    loConfigItem("HorizontalSplitPosition", "int", view.splitPosition[0]) +
    loConfigItem("VerticalSplitPosition", "int", view.splitPosition[1]) +
    loConfigItem("ActiveSplitRange", "short", view.activeSplitRange) +
    loConfigItem("PositionLeft", "int", view.position.left) +
    loConfigItem("PositionRight", "int", view.position.right) +
    loConfigItem("PositionTop", "int", view.position.top) +
    loConfigItem("PositionBottom", "int", view.position.bottom) +
    loConfigItem("ZoomType", "short", 0) +
    loConfigItem("ZoomValue", "int", view.zoom) +
    loConfigItem("PageViewZoomValue", "int", 60) +
    loConfigItem("ShowGrid", "boolean", view.showGrid) +
    loConfigItem("AnchoredTextOverflowLegacy", "boolean", false) +
    loConfigItem("LegacySingleLineFontwork", "boolean", false) +
    loConfigItem("ConnectorUseSnapRect", "boolean", false) +
    loConfigItem("IgnoreBreakAfterMultilineField", "boolean", false) +
    `</config:config-item-map-entry>`
  );
}

function spec117OdsSettings(): string {
  const tables =
    spec117TableSettings("Styled", {
      cursor: [1, 2],
      split: [2, 2],
      splitPosition: [1, 2],
      activeSplitRange: 3,
      position: { left: 0, right: 1, top: 0, bottom: 2 },
      zoom: 85,
      showGrid: true,
    }) +
    spec117TableSettings("Plain", {
      cursor: [0, 0],
      split: [0, 0],
      splitPosition: [0, 0],
      activeSplitRange: 2,
      position: { left: 0, right: 0, top: 0, bottom: 0 },
      zoom: 100,
      showGrid: false,
    }) +
    spec117TableSettings("Split", {
      cursor: [0, 0],
      split: [1, 0],
      splitPosition: [3216, 0],
      activeSplitRange: 2,
      position: { left: 0, right: 4, top: 0, bottom: 0 },
      zoom: 100,
      showGrid: true,
    });
  return (
    ODF_XML_DECLARATION +
    `<office:document-settings xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" ` +
    `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3">` +
    `<office:settings><config:config-item-set config:name="ooo:view-settings">` +
    loConfigItem("VisibleAreaTop", "int", 0) +
    loConfigItem("VisibleAreaLeft", "int", 0) +
    loConfigItem("VisibleAreaWidth", "int", 13544) +
    loConfigItem("VisibleAreaHeight", "int", 4516) +
    `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    loConfigItem("ViewId", "string", "view1") +
    `<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>` +
    loConfigItem("ActiveTable", "string", SPEC117_STYLES_ODS.activeTable) +
    loConfigItem("HorizontalScrollbarWidth", "int", 1309) +
    loConfigItem("ZoomType", "short", 0) +
    loConfigItem("ZoomValue", "int", 85) +
    loConfigItem("PageViewZoomValue", "int", 60) +
    loConfigItem("ShowPageBreakPreview", "boolean", false) +
    loConfigItem("ShowZeroValues", "boolean", true) +
    loConfigItem("ShowNotes", "boolean", true) +
    loConfigItem("ShowGrid", "boolean", true) +
    loConfigItem("GridColor", "int", 12632256) +
    loConfigItem("ShowPageBreaks", "boolean", true) +
    loConfigItem("HasColumnRowHeaders", "boolean", true) +
    loConfigItem("HasSheetTabs", "boolean", true) +
    loConfigItem("IsOutlineSymbolsSet", "boolean", true) +
    loConfigItem("IsValueHighlightingEnabled", "boolean", false) +
    loConfigItem("IsSnapToRaster", "boolean", false) +
    loConfigItem("RasterIsVisible", "boolean", false) +
    loConfigItem("RasterResolutionX", "int", 1000) +
    loConfigItem("RasterResolutionY", "int", 1000) +
    loConfigItem("RasterSubdivisionX", "int", 1) +
    loConfigItem("RasterSubdivisionY", "int", 1) +
    loConfigItem("IsRasterAxisSynchronized", "boolean", true) +
    loConfigItem("AnchoredTextOverflowLegacy", "boolean", false) +
    loConfigItem("LegacySingleLineFontwork", "boolean", false) +
    loConfigItem("ConnectorUseSnapRect", "boolean", false) +
    loConfigItem("IgnoreBreakAfterMultilineField", "boolean", false) +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>` +
    `<config:config-item-set config:name="ooo:configuration-settings">` +
    loConfigItem("IsDocumentShared", "boolean", false) +
    loConfigItem("SaveVersionOnClose", "boolean", false) +
    loConfigItem("ShowGrid", "boolean", true) +
    loConfigItem("ShowZeroValues", "boolean", true) +
    `</config:config-item-set></office:settings></office:document-settings>`
  );
}

/** What spec117-styles.ods holds, for the unit and e2e specs. Row and column
 * numbers here are 1-based / A1 unless named `index` (0-based). */
export const SPEC117_STYLES_ODS = {
  tableNames: ["Styled", "Plain", "Split"],
  activeTable: "Styled",
  /** "Styled": the active table — every construct the spec reads. */
  styled: {
    /** B2: bold red text on yellow (ce1, parented to "Default"). */
    header: { a1: "B2", value: "Header", styleName: "ce1", bold: true, color: "#ff0000", fill: "#ffff00" },
    /** B3:B12 hold 10, 20, …, 100; B13 and everything below is empty. */
    numbers: { from: "B3", to: "B12", values: [10, 20, 30, 40, 50, 60, 70, 80, 90, 100] },
    /** C5 spans C5:E5 (3×1); the covered D5 and E5 carry its style. */
    merge: { range: "C5:E5", anchor: "C5", value: "Merged title", styleName: "ce12", covered: ["D5", "E5"] },
    /** ce12 = the common "Heading" (24pt bold) + centred. */
    titleFont: { family: "Liberation Sans", sizePt: 24, bold: true },
    /** Column G (`table:visibility="collapse"`); F and H are visible. */
    hiddenColumn: { index: 6, letter: "G" },
    /** Column B is wider (co2, 1.2in). */
    wideColumn: { index: 1, widthPx: 115.2, styleName: "co2" },
    /** Column A's default cell style is ce9 (italic). */
    styledColumn: { index: 0, styleName: "ce9" },
    /** co1 (0.889in) is the last column run's style; ro1 (0.178in) the last row run's. */
    defaultColWidthPx: 85.344,
    defaultRowHeightPt: 12.816,
    /** Row 1: four headings in ce10 and a ce10 run of 16380 cells to XFD —
     * the row's style, storing no cell of its own. */
    bannerRow: { index: 0, styleName: "ce10", values: ["Region", "Amount", "Title", "Notes"], repeatedRun: 16380 },
    /** Rows with heights of their own (index → pt; custom = use-optimal false). */
    sizedRows: [
      { index: 0, heightPt: 21.6, customHeight: true },
      { index: 4, heightPt: 32.4 },
      { index: 6, heightPt: 43.2, customHeight: true },
      { index: 15, heightPt: 21.6, customHeight: true },
    ],
    /** Row 14 is hidden (it holds D14); row 16 is stored for its height alone. */
    hiddenRow: { index: 13, a1: "D14", value: "Hidden row" },
    emptySizedRow: { index: 15 },
    /** C3: "Mixed bold run", the middle word a text:span in T1. */
    richText: {
      a1: "C3",
      value: "Mixed bold run",
      runs: [{ text: "Mixed " }, { text: "bold", font: { bold: true, color: "#c9211e" } }, { text: " run" }],
    },
    borders: { D7: "thin", D8: "medium", D9: "thick", D10: "dashed", D11: "hair", D12: "double" },
    /** D6: a red thin diagonal (bottom-left → top-right) over a blue 1.75pt dashed bottom edge. */
    diagonal: { a1: "D6", diagonal: { style: "thin", color: "#ff0000", up: true }, bottom: { style: "mediumDashed", color: "#0000ff" } },
    alignment: {
      F7: { wrap: true },
      F8: { textRotation: 90 },
      F9: { horizontal: "fill" },
      F10: { horizontal: "left", indent: 2 },
      F11: { vertical: "center", shrinkToFit: true },
    },
    tabColor: "#729fcf",
    /** settings.xml: both split modes 2 → rows 2 and column A frozen; the
     * scrolled pane starts at B3; zoom 85; the cursor on B3. */
    frozen: { rows: 2, cols: 1 },
    topLeftCell: "B3",
    zoomPercent: 85,
    cursor: "B3",
    /** 16 rows written, then the trailing run to the grid's edge. */
    trailingRowRun: 1_048_576 - 16,
  },
  /** "Plain": ShowGrid false, zoom 100 (→ no zoomPercent). */
  plain: { showGridLines: false },
  /** "Split": a split (mode 1) — not a freeze; the normal view. */
  split: { a1: "A1", value: "Split view" },
} as const;

/** A styled empty run, as Calc writes a format reaching past the data. */
const loStyledRun = (styleName: string, repeat: number): string =>
  `<table:table-cell table:style-name="${styleName}" table:number-columns-repeated="${repeat}"/>`;

export async function buildSpec117StylesOdsBytes(): Promise<Uint8Array> {
  /** A row of `B` = value, `D` and `F` optional styled strings. */
  const numberRow = (value: number, d?: [string, string], f?: [string, string], style = "ro1"): string =>
    loRow(
      loEmpty(1) +
        loFloat(value) +
        (d || f ? loEmpty(1) + (d ? loString(d[0], d[1]) : loEmpty(1)) : "") +
        (f ? loEmpty(1) + loString(f[0], f[1]) + loEmpty(16378) : d ? loEmpty(16380) : loEmpty(16382)),
      style,
    );
  const styled =
    `<table:table table:name="Styled" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:default-cell-style-name="ce9"/>` +
    `<table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="4" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:visibility="collapse" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16377" table:default-cell-style-name="Default"/>` +
    // Row 1: a whole-row format — four headings, then the style run to XFD.
    loRow(SPEC117_STYLES_ODS.styled.bannerRow.values.map((text) => loString(text, "ce10")).join("") + loStyledRun("ce10", 16380), "ro2") +
    loRow(loEmpty(1) + loString("Header", "ce1") + loEmpty(16382)) +
    loRow(
      loEmpty(1) +
        loFloat(10) +
        `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>Mixed <text:span text:style-name="T1">bold</text:span> run</text:p>` +
        `</table:table-cell>` +
        loEmpty(16381),
    ) +
    numberRow(20) +
    loRow(
      loEmpty(1) +
        loFloat(30) +
        `<table:table-cell table:style-name="ce12" table:number-columns-spanned="3" table:number-rows-spanned="1" office:value-type="string" ` +
        `calcext:value-type="string"><text:p>Merged title</text:p></table:table-cell>` +
        `<table:covered-table-cell table:style-name="ce12" table:number-columns-repeated="2"/>` +
        loEmpty(16379),
      "ro3",
    ) +
    numberRow(40, ["Diagonal", "ce16"]) +
    numberRow(50, ["Thin", "ce2"], ["Wrapped text in a narrow column", "ce8"], "ro4") +
    numberRow(60, ["Medium", "ce3"], ["Up", "ce11"]) +
    numberRow(70, ["Thick", "ce4"], ["=", "ce13"]) +
    numberRow(80, ["Dashed", "ce5"], ["Indented", "ce14"]) +
    numberRow(90, ["Hair", "ce6"], ["Middle", "ce15"]) +
    numberRow(100, ["Double", "ce7"]) +
    loRow(loEmpty(16384)) +
    `<table:table-row table:style-name="ro1" table:visibility="collapse">${loEmpty(3)}${loString("Hidden row")}${loEmpty(16380)}</table:table-row>` +
    loRow(loEmpty(16384)) +
    loRow(loEmpty(16384), "ro2") +
    loRow(loEmpty(16384), "ro1", SPEC117_STYLES_ODS.styled.trailingRowRun) +
    `</table:table>`;

  const plain =
    `<table:table table:name="Plain" table:style-name="ta2">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString("Plain sheet") + loFloat(1) + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1_048_575) +
    `</table:table>`;

  const split =
    `<table:table table:name="Split" table:style-name="ta2">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString(SPEC117_STYLES_ODS.split.value) + loEmpty(16383)) +
    loRow(loEmpty(16384), "ro1", 1_048_575) +
    `</table:table>`;

  return loPackage(spec117OdsContent(styled + plain + split), loMeta("Spec 117 styles", 3, 36), spec117OdsSettings());
}

// ============================================================================
// SPEC-118: number formats
// ============================================================================

/** The serial every SPEC-118 date fixture holds: 2026-10-08 13:45:00. */
export const SPEC118_SERIAL = 46303.5729166667;

/** spec118-formats.xlsx's custom number formats (ids ≥ 164), as Excel 365
 * writes them, and the explicit entries it writes for 43 and 44. */
export const SPEC118_NUMFMTS: Readonly<Record<number, string>> = {
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  164: "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
  165: "[$€-407]#,##0.00",
  166: "0.00;[Red]\\(0.00\\)",
  167: "# ??/??",
  168: "[>100][Red]0;[Blue]0",
  169: ";;;",
  170: "#,##0.00_);[Red]\\(#,##0.00\\)",
  171: "d-mmm-yyyy",
};

/** spec118-formats.xlsx's cellXfs, index → numFmtId (xf 1 is the bold header). */
const SPEC118_XF_NUMFMTS = [0, 0, 14, 22, 164, 170, 4, 10, 46, 171, 9, 11, 49, 165, 166, 167, 168, 169, 43, 44] as const;

/** The xf index of each format the "Formats" sheet uses. */
export const SPEC118_XF = {
  header: 1,
  shortDate: 2, // 14
  dateTime: 3, // 22
  systemLongDate: 4, // 164 [$-F800]…
  accountingRed: 5, // 170
  number2: 6, // 4
  percent2: 7, // 10
  elapsed: 8, // 46
  dMmmYyyy: 9, // 171
  percent0: 10, // 9
  scientific: 11, // 11
  text: 12, // 49
  euro: 13, // 165
  redParens: 14, // 166
  fraction: 15, // 167
  conditional: 16, // 168
  hidden: 17, // 169
  comma: 18, // 43
  currency: 19, // 44
} as const;

/** An attribute value as Excel escapes it (`&`, `<`, `"`). */
function excelAttr(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");
}

const SPEC118_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<numFmts count="${Object.keys(SPEC118_NUMFMTS).length}">` +
  Object.entries(SPEC118_NUMFMTS)
    .map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${excelAttr(code)}"/>`)
    .join("") +
  `</numFmts>` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/>${SPEC117_FONT_TAIL}` +
  `<font><b/><sz val="11"/><color theme="1"/>${SPEC117_FONT_TAIL}</fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="43" fontId="0" fillId="0" borderId="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="44" fontId="0" fillId="0" borderId="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
  `<cellXfs count="${SPEC118_XF_NUMFMTS.length}">` +
  SPEC118_XF_NUMFMTS.map((numFmtId, index) => {
    if (index === SPEC118_XF.header) {
      return `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>`;
    }
    const xfId = numFmtId === 43 ? 1 : numFmtId === 44 ? 2 : 0;
    return `<xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0" xfId="${xfId}"${numFmtId === 0 ? "" : ` applyNumberFormat="1"`}/>`;
  }).join("") +
  `</cellXfs>` +
  `<cellStyles count="3"><cellStyle name="Comma" xfId="1" builtinId="3"/><cellStyle name="Currency" xfId="2" builtinId="4"/>` +
  `<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
  `<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`;

const SPEC118_STRINGS = ["Date", "Amount", "Ratio", "Hours", "Other", "Big", "Day", "Text value"] as const;

/** What spec118-formats.xlsx holds (sheet "Formats", Calibri 11 — MDW 7), for
 * the unit and e2e specs: each cell's value, xf and what en-US shows. */
export const SPEC118_FORMATS_XLSX = {
  sheetName: "Formats",
  /** Column widths in px: A and B wide enough for the long texts, F narrow
   * enough that General drops to "1.23E+11", G too narrow for any date. */
  columnWidthsPx: { A: 200, B: 100, E: 100, F: 60, G: 30 },
  cells: {
    A2: { value: SPEC118_SERIAL, xf: SPEC118_XF.shortDate, enUS: "10/8/2026", deDE: "08.10.2026" },
    A3: { value: SPEC118_SERIAL, xf: SPEC118_XF.dateTime, enUS: "10/8/2026 13:45", deDE: "08.10.2026 13:45" },
    A4: { value: SPEC118_SERIAL, xf: SPEC118_XF.systemLongDate, enUS: "Thursday, October 8, 2026", deDE: "Donnerstag, 8. Oktober 2026" },
    A6: { value: 0, xf: SPEC118_XF.dMmmYyyy, enUS: "0-Jan-1900" },
    A7: { value: 60, xf: SPEC118_XF.dMmmYyyy, enUS: "29-Feb-1900" },
    A8: { value: 61, xf: SPEC118_XF.dMmmYyyy, enUS: "1-Mar-1900" },
    B2: { value: -1234.5, xf: SPEC118_XF.accountingRed, enUS: "(1,234.50)", deDE: "(1.234,50)", color: "#ff0000" },
    B3: { value: 1234.5, xf: SPEC118_XF.euro, enUS: "€1,234.50" },
    B4: { value: -3, xf: SPEC118_XF.redParens, enUS: "(3.00)", color: "#ff0000" },
    B5: { value: 1234.5, xf: SPEC118_XF.number2, enUS: "1,234.50", deDE: "1.234,50" },
    B6: { value: 150, xf: SPEC118_XF.conditional, enUS: "150", color: "#ff0000" },
    B7: { value: -5, xf: SPEC118_XF.conditional, enUS: "-5", color: "#0000ff" },
    C2: { value: 0.125, xf: SPEC118_XF.percent2, enUS: "12.50%", deDE: "12,50%" },
    D2: { value: 1.5, xf: SPEC118_XF.elapsed, enUS: "36:00:00" },
    E2: { value: 0.25, xf: SPEC118_XF.percent0, enUS: "25%" },
    E3: { value: 12345, xf: SPEC118_XF.scientific, enUS: "1.23E+04" },
    E4: { value: "Text value", xf: SPEC118_XF.text, enUS: "Text value" },
    E5: { value: Math.PI, xf: SPEC118_XF.fraction, enUS: "3 14/99" },
    E6: { value: 42, xf: SPEC118_XF.hidden, enUS: "" },
    E7: { value: 1234.5, xf: SPEC118_XF.comma },
    E8: { value: -1234.5, xf: SPEC118_XF.currency },
    F2: { value: 123456789012, xf: 0, enUS: "1.23E+11" },
    G2: { value: 46303, xf: SPEC118_XF.shortDate, enUS: "###" },
  },
  /** The formula bar for A2:A4 (the same serial). */
  editText: { enUS: "10/8/2026 1:45:00 PM", deDE: "08.10.2026 13:45:00" },
} as const;

/** One of spec118-formats.xlsx's Excel-shaped rows (A:G). */
const spec118Row = (r: number, cells: string): string => `<row r="${r}" spans="1:7" x14ac:dyDescent="0.25">${cells}</row>`;

export async function buildSpec118FormatsXlsxBytes(): Promise<Uint8Array> {
  const X = SPEC118_XF;
  const C = SPEC118_FORMATS_XLSX.cells;
  const text = (ref: string, value: (typeof SPEC118_STRINGS)[number], style?: number): string => sharedCell(ref, SPEC118_STRINGS.indexOf(value), style);
  const row = spec118Row;
  const rows = [
    row(
      1,
      text("A1", "Date", X.header) + text("B1", "Amount", X.header) + text("C1", "Ratio", X.header) + text("D1", "Hours", X.header) +
        text("E1", "Other", X.header) + text("F1", "Big", X.header) + text("G1", "Day", X.header),
    ),
    row(
      2,
      numberCell("A2", C.A2.value, C.A2.xf) + numberCell("B2", C.B2.value, C.B2.xf) + numberCell("C2", C.C2.value, C.C2.xf) +
        numberCell("D2", C.D2.value, C.D2.xf) + numberCell("E2", C.E2.value, C.E2.xf) + numberCell("F2", C.F2.value) + numberCell("G2", C.G2.value, C.G2.xf),
    ),
    row(3, numberCell("A3", C.A3.value, C.A3.xf) + numberCell("B3", C.B3.value, C.B3.xf) + numberCell("E3", C.E3.value, C.E3.xf)),
    row(4, numberCell("A4", C.A4.value, C.A4.xf) + numberCell("B4", C.B4.value, C.B4.xf) + text("E4", "Text value", C.E4.xf)),
    row(5, numberCell("B5", C.B5.value, C.B5.xf) + numberCell("E5", C.E5.value, C.E5.xf)),
    row(6, numberCell("A6", C.A6.value, C.A6.xf) + numberCell("B6", C.B6.value, C.B6.xf) + numberCell("E6", C.E6.value, C.E6.xf)),
    row(7, numberCell("A7", C.A7.value, C.A7.xf) + numberCell("B7", C.B7.value, C.B7.xf) + numberCell("E7", C.E7.value, C.E7.xf)),
    row(8, numberCell("A8", C.A8.value, C.A8.xf) + numberCell("E8", C.E8.value, C.E8.xf)),
  ].join("");
  // Widths through Excel's own px → characters inverse at MDW 7.
  const width = (px: number): number => pxToXlsxWidthMdw7(px);
  const W = SPEC118_FORMATS_XLSX.columnWidthsPx;
  const sheet =
    XML_DECLARATION +
    worksheetOpen("D6F4B8C0-2E5A-4C9D-9F4B-6A8C0E2F4A18") +
    `<dimension ref="A1:G8"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<cols><col min="1" max="1" width="${width(W.A)}" customWidth="1"/><col min="2" max="2" width="${width(W.B)}" customWidth="1"/>` +
    `<col min="5" max="5" width="${width(W.E)}" customWidth="1"/><col min="6" max="6" width="${width(W.F)}" customWidth="1"/>` +
    `<col min="7" max="7" width="${width(W.G)}" customWidth="1"/></cols>` +
    `<sheetData>${rows}</sheetData>${SPEC117_PAGE_MARGINS}</worksheet>`;
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${sharedStringReferences(sheet)}" uniqueCount="${SPEC118_STRINGS.length}">` +
    SPEC118_STRINGS.map((value) => `<si><t>${value}</t></si>`).join("") +
    `</sst>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Formats",
    sheets: [{ name: SPEC118_FORMATS_XLSX.sheetName, xml: sheet }],
    sharedStrings,
    workbookPr: `<workbookPr defaultThemeVersion="166925"/>`,
    theme: SPEC117_THEME,
    styles: SPEC118_STYLES,
  });
}

/** Excel's px → `<col width>` inverse for a 7 px maximum digit width
 * (geometry.ts pxToXlsxWidth, restated so the builders stay self-contained). */
function pxToXlsxWidthMdw7(px: number): number {
  const mdw = 7;
  const characters = Math.trunc(((px - 5) / mdw) * 100 + 0.5) / 100;
  return Math.trunc(((characters * mdw + 5) / mdw) * 256) / 256;
}

// --- openpyxl -----------------------------------------------------------------

const OPENPYXL_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

/** What spec118-generated.xlsx holds: an openpyxl-shaped workbook whose
 * dates are `t="d"` ISO values under a General style (no theme, no numFmts)
 * — the reader gives them the implied `yyyy-mm-dd` (+ time) format. */
export const SPEC118_GENERATED_XLSX = {
  sheetName: "Generated",
  cells: {
    A1: "Created",
    B1: { iso: "2026-10-08T13:45:00", shows: "2026-10-08 13:45:00" },
    A2: "Due",
    B2: { iso: "2026-10-08T00:00:00", shows: "2026-10-08", styleId: 0 },
    A3: "Centred",
    B3: { iso: "2026-10-09T00:00:00", shows: "2026-10-09", styleId: 1 },
  },
} as const;

/** openpyxl's inline string cell. */
const openpyxlInline = (ref: string, value: string): string => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;

export async function buildSpec118GeneratedXlsxBytes(): Promise<Uint8Array> {
  const G = SPEC118_GENERATED_XLSX.cells;
  const inline = openpyxlInline;
  const sheet =
    OPENPYXL_DECLARATION +
    `<worksheet xmlns="${MAIN_NS}"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr/></sheetPr><dimension ref="A1:B3"/>` +
    `<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    `<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/><cols><col min="2" max="2" width="22" customWidth="1"/></cols><sheetData>` +
    `<row r="1">${inline("A1", G.A1)}<c r="B1" t="d"><v>${G.B1.iso}</v></c></row>` +
    `<row r="2">${inline("A2", G.A2)}<c r="B2" s="0" t="d"><v>${G.B2.iso}</v></c></row>` +
    `<row r="3">${inline("A3", G.A3)}<c r="B3" s="1" t="d"><v>${G.B3.iso}</v></c></row>` +
    `</sheetData><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/></worksheet>`;
  const styles =
    OPENPYXL_DECLARATION +
    `<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font></fonts>` +
    `<fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0"/>` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0"><alignment horizontal="center"/></xf></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" hidden="0"/></cellStyles>` +
    `<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  const workbook =
    OPENPYXL_DECLARATION +
    `<workbook xmlns="${MAIN_NS}"><workbookPr/><workbookProtection/><bookViews><workbookView visibility="visible" minimized="0" ` +
    `showHorizontalScroll="1" showVerticalScroll="1" showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="0" autoFilterDateGrouping="1"/>` +
    `</bookViews><sheets><sheet xmlns:r="${REL_NS}" name="${SPEC118_GENERATED_XLSX.sheetName}" sheetId="1" state="visible" r:id="rId1"/></sheets>` +
    `<definedNames/><calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>`;
  const contentTypes =
    OPENPYXL_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  return zip([
    part(
      "docProps/app.xml",
      OPENPYXL_DECLARATION +
        `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application>` +
        `<AppVersion>3.1</AppVersion></Properties>`,
    ),
    part(
      "docProps/core.xml",
      OPENPYXL_DECLARATION +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
        `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>openpyxl</dc:creator>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-08T12:00:00Z</dcterms:created>` +
        `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-08T12:00:00Z</dcterms:modified></cp:coreProperties>`,
    ),
    part("xl/worksheets/sheet1.xml", sheet),
    part("xl/styles.xml", styles),
    part("xl/workbook.xml", workbook),
    part(
      "xl/_rels/workbook.xml.rels",
      OPENPYXL_DECLARATION +
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Type="${REL_TYPE}/worksheet" Target="/xl/worksheets/sheet1.xml" Id="rId1"/>` +
        `<Relationship Type="${REL_TYPE}/styles" Target="styles.xml" Id="rId2"/></Relationships>`,
    ),
    part(
      "_rels/.rels",
      OPENPYXL_DECLARATION +
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Type="${REL_TYPE}/officeDocument" Target="xl/workbook.xml" Id="rId1"/>` +
        `<Relationship Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" Id="rId2"/>` +
        `<Relationship Type="${REL_TYPE}/extended-properties" Target="docProps/app.xml" Id="rId3"/></Relationships>`,
    ),
    part("[Content_Types].xml", contentTypes),
  ]);
}

// --- LibreOffice data styles ----------------------------------------------------

/** spec118-formats.ods's data styles (Calc 24.8 shapes, content.xml's
 * automatic styles), each with the table-cell style ceN that names it. */
const SPEC118_ODS_DATA_STYLES =
  // N2: 0.00;[Red]-0.00 — LibreOffice's red negative, the positive half a P0 sub-style.
  `<number:number-style style:name="N2P0" style:volatile="true"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/></number:number-style>` +
  `<number:number-style style:name="N2"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N2P0"/></number:number-style>` +
  // N3: euro, German — #,##0.00 [$€-407] with a red negative.
  `<number:currency-style style:name="N3P0" style:volatile="true" number:language="de" number:country="DE">` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>` +
  `<number:text> </number:text><number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol></number:currency-style>` +
  `<number:currency-style style:name="N3" number:language="de" number:country="DE"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>` +
  `<number:text> </number:text><number:currency-symbol number:language="de" number:country="DE">€</number:currency-symbol>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N3P0"/></number:currency-style>` +
  // N4: a long textual date in German, whatever the reader's locale.
  `<number:date-style style:name="N4" number:language="de" number:country="DE"><number:day-of-week number:style="long"/><number:text>, </number:text>` +
  `<number:day/><number:text>. </number:text><number:month number:style="long" number:textual="true"/><number:text> </number:text>` +
  `<number:year number:style="long"/></number:date-style>` +
  // N5: [HH]:MM:SS — elapsed hours.
  `<number:time-style style:name="N5" number:truncate-on-overflow="false"><number:hours number:style="long"/><number:text>:</number:text>` +
  `<number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>` +
  // N6: # ??/??
  `<number:number-style style:name="N6"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="2" ` +
  `number:min-denominator-digits="2" loext:max-numerator-digits="2" number:max-denominator-value="99"/></number:number-style>` +
  // N7: 0.00E+00
  `<number:number-style style:name="N7"><number:scientific-number number:decimal-places="2" number:min-decimal-places="2" ` +
  `number:min-integer-digits="1" number:min-exponent-digits="2" number:exponent-interval="1" number:forced-exponent-sign="true"/></number:number-style>` +
  // N8: 0.00%
  `<number:percentage-style style:name="N8"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
  `<number:text>%</number:text></number:percentage-style>` +
  // N9: BOOLEAN
  `<number:boolean-style style:name="N9"><number:boolean/></number:boolean-style>` +
  // N10: "Note: "@
  `<number:text-style style:name="N10"><number:text>Note: </number:text><number:text-content/></number:text-style>` +
  // N11: Q YYYY
  `<number:date-style style:name="N11"><number:quarter/><number:text> </number:text><number:year number:style="long"/></number:date-style>` +
  // N12: "Week "WW
  `<number:date-style style:name="N12"><number:text>Week </number:text><number:week-of-year/></number:date-style>` +
  // N13: an era — out of scope, read as General.
  `<number:date-style style:name="N13" number:language="ja" number:country="JP"><number:era number:style="long"/><number:year/></number:date-style>`;

/** What spec118-formats.ods holds (table "Formats": a label in A, the value
 * in the wide B), for the unit and e2e specs. */
export const SPEC118_FORMATS_ODS = {
  tableName: "Formats",
  rows: {
    2: { label: "Red negative", style: "ce1", code: "0.00;[Red]-0.00", enUS: "-1.00", deDE: "-1,00", color: "#ff0000" },
    3: { label: "Currency", style: "ce2", code: "#,##0.00 [$€-407];[Red]-#,##0.00 [$€-407]", enUS: "1,234.50 €", deDE: "1.234,50 €" },
    4: { label: "Long date", style: "ce3", code: "[$-407]dddd, d. mmmm yyyy", enUS: "Donnerstag, 8. Oktober 2026" },
    5: { label: "Elapsed", style: "ce4", code: "[hh]:mm:ss", enUS: "36:00:00" },
    6: { label: "Fraction", style: "ce5", code: "# ??/??", enUS: "3 14/99" },
    7: { label: "Scientific", style: "ce6", code: "0.00E+00", enUS: "1.23E+04" },
    8: { label: "Percent", style: "ce7", code: "0.00%", enUS: "12.50%" },
    9: { label: "Boolean", style: "ce8", code: '"TRUE";"TRUE";"FALSE"', enUS: "TRUE" },
    10: { label: "Text", style: "ce9", code: '"Note: "@', enUS: "Note: ok" },
    11: { label: "Quarter", style: "ce10", code: "Q yyyy", enUS: "Q4 2026" },
    12: { label: "Week", style: "ce11", code: '"Week "WW', enUS: "Week 41" },
    13: { label: "Era", style: "ce12", code: undefined, enUS: "46303" },
    14: { label: "Unstyled date", style: undefined, code: "yyyy-mm-dd", enUS: "1899-12-25" },
    15: { label: "Unstyled date and time", style: undefined, code: "yyyy-mm-dd hh:mm:ss", enUS: "2026-10-08 13:45:00" },
    16: { label: "Unstyled time", style: undefined, code: "[hh]:mm:ss", enUS: "01:30:00" },
    17: { label: "Unstyled percentage", style: undefined, code: "0.00%", enUS: "50.00%" },
  },
} as const;

const SPEC118_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="1.5in"/></style:style>` +
  `<style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="2.5in"/></style:style>` +
  `<style:style style:name="co3" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  SPEC118_ODS_DATA_STYLES +
  Array.from({ length: 12 }, (_, k) => `<style:style style:name="ce${k + 1}" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N${k + 2}"/>`).join("") +
  `</office:automatic-styles>`;

/** A LibreOffice value cell: its style (if any), value attributes and text. */
const loValueCell = (style: string | undefined, attrs: string, text: string): string =>
  `<table:table-cell${style ? ` table:style-name="${style}"` : ""} ${attrs}><text:p>${text}</text:p></table:table-cell>`;

export async function buildSpec118FormatsOdsBytes(): Promise<Uint8Array> {
  const R = SPEC118_FORMATS_ODS.rows;
  const cell = loValueCell;
  const valueRow = (label: string, value: string): string => loRow(loString(label) + value + loEmpty(16382));
  const date = (style: string | undefined, iso: string, text: string): string =>
    cell(style, `office:value-type="date" office:date-value="${iso}" calcext:value-type="date"`, text);
  const float = (style: string | undefined, value: number, text: string): string =>
    cell(style, `office:value-type="float" office:value="${value}" calcext:value-type="float"`, text);
  const table =
    `<table:table table:name="${SPEC118_FORMATS_ODS.tableName}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co3" table:number-columns-repeated="16382" table:default-cell-style-name="Default"/>` +
    loRow(loString("Format") + loString("Value") + loEmpty(16382)) +
    valueRow(R[2].label, float(R[2].style, -1, "-1.00")) +
    valueRow(
      R[3].label,
      cell(R[3].style, `office:value-type="currency" office:currency="EUR" office:value="1234.5" calcext:value-type="currency"`, "1.234,50 €"),
    ) +
    valueRow(R[4].label, date(R[4].style, "2026-10-08T13:45:00", "Donnerstag, 8. Oktober 2026")) +
    valueRow(R[5].label, cell(R[5].style, `office:value-type="time" office:time-value="PT36H00M00S" calcext:value-type="time"`, "36:00:00")) +
    valueRow(R[6].label, float(R[6].style, Math.PI, "3 14/99")) +
    valueRow(R[7].label, float(R[7].style, 12345, "1.23E+04")) +
    valueRow(R[8].label, cell(R[8].style, `office:value-type="percentage" office:value="0.125" calcext:value-type="percentage"`, "12.50%")) +
    valueRow(R[9].label, cell(R[9].style, `office:value-type="boolean" office:boolean-value="true" calcext:value-type="boolean"`, "TRUE")) +
    valueRow(R[10].label, cell(R[10].style, `office:value-type="string" calcext:value-type="string"`, "ok")) +
    valueRow(R[11].label, date(R[11].style, "2026-10-08", "Q4 2026")) +
    valueRow(R[12].label, date(R[12].style, "2026-10-08", "Week 41")) +
    valueRow(R[13].label, date(R[13].style, "2026-10-08", "Reiwa 8")) +
    valueRow(R[14].label, date(undefined, "1899-12-25", "12/25/99")) +
    valueRow(R[15].label, date(undefined, "2026-10-08T13:45:00", "10/08/26 13:45")) +
    valueRow(R[16].label, cell(undefined, `office:value-type="time" office:time-value="PT01H30M00S" calcext:value-type="time"`, "01:30:00")) +
    valueRow(R[17].label, cell(undefined, `office:value-type="percentage" office:value="0.5" calcext:value-type="percentage"`, "50.00%")) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 17) +
    `</table:table>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC118_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false" ` +
    `table:use-regular-expressions="false" table:use-wildcards="true"/>${table}<table:named-expressions/>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return loPackage(content, loMeta("Spec 118 formats", 1, 34), loSettings([SPEC118_FORMATS_ODS.tableName], SPEC118_FORMATS_ODS.tableName));
}

// ============================================================================
// SPEC-119 — formulas
// ============================================================================

/** An XML attribute value / element text as Excel and LibreOffice escape it. */
function xmlEscape(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

/** Excel writes `"` literally in element text, and `>` as `&gt;`. */
function xmlText(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

const SPEC119_STRINGS = ["Value", "Amount"] as const;

/** spec119-formulas.xlsx's "Calc" `<f>` texts, as Excel wrote them. */
const SPEC119_WIRE = {
  B5: "SUM(B2:B4)",
  C2: "B2*$F$1",
  D3: "A1048575*2",
  E2: "A2:A4*2",
  G2: '_xlfn.IFS(A2>1,"big",TRUE,"small")',
  H2: "_xlfn._xlws.SORT(A2:A6)",
  I2: "_xlfn.LET(_xlpm.x,5,_xlpm.x*2)",
  J2: "_xlfn.SINGLE(A2:A4)",
  L2: "F2*10",
  M2: "SUM(Jan:Mar!B2)",
  N2: "'O''Brien'!A1*2",
  O2: "SUM(A:A)+SUM(1:1)",
  P2: "[1]Sheet1!$A$1+1",
  Q2: "SUM(Sheet2!#REF!)",
  R2: 'IF(B2>15,\n    "high",\n    "low")',
  S2: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)",
} as const;

/** What spec119-formulas.xlsx holds, for the unit and e2e specs. */
export const SPEC119_FORMULAS_XLSX = {
  sheetNames: ["Calc", "Jan", "Feb", "Mar", "O'Brien"],
  /** "Calc"'s `<f>` texts as written. */
  wire: SPEC119_WIRE,
  /** "Calc"'s formula cells → the canonical formula the reader must give. */
  formulas: {
    B5: "SUM(B2:B4)",
    C2: "B2*$F$1",
    C3: "B3*$F$1",
    C4: "B4*$F$1",
    C5: "B5*$F$1",
    C6: "B6*$F$1",
    C7: "B7*$F$1",
    C8: "B8*$F$1",
    C9: "B9*$F$1",
    D2: "A1048574*2",
    D3: "A1048575*2",
    D4: "A1048576*2",
    D5: "#REF!*2",
    E2: "A2:A4*2",
    G2: 'IFS(A2>1,"big",TRUE,"small")',
    H2: "SORT(A2:A6)",
    I2: "LET(x,5,x*2)",
    J2: "SINGLE(A2:A4)",
    L2: "F2*10",
    M2: "SUM(Jan:Mar!B2)",
    N2: "'O''Brien'!A1*2",
    O2: "SUM(A:A)+SUM(1:1)",
    P2: "[1]Sheet1!$A$1+1",
    Q2: "SUM(Sheet2!#REF!)",
    R2: 'IF(B2>15,\n    "high",\n    "low")',
  },
  /** The formula this grammar does not parse: kept as its `<f>` text. */
  unparsed: { S2: SPEC119_WIRE.S2 },
  /** The cached values of the formula cells the specs look at. */
  values: { B5: 60, C3: 40, C5: 120, D5: { error: "#REF!" }, E2: 6, E3: 2, E4: 10, G2: "big", H2: 1, H6: 5, I2: 10, L3: 10, S2: 3 },
  arrays: { E2: { top: 1, left: 4, bottom: 3, right: 4 }, H2: { top: 1, left: 7, bottom: 5, right: 7 } },
  /** The data table's `<f>`, verbatim, on L3 (L3:L5, column input F2). */
  dataTable: { L3: `<f t="dataTable" ref="L3:L5" dt2D="0" dtr="0" r1="F2"/>` },
  /** The formula bar's text for the e2e journey's cells. */
  formulaBar: { B5: "=SUM(B2:B4)", C3: "=B3*$F$1", E3: "{=A2:A4*2}", G2: '=IFS(A2>1,"big",TRUE,"small")', H4: "=SORT(A2:A6)", L4: "{=TABLE(,F2)}" },
} as const;

/** Excel's dynamic-array cell metadata part (the XLDAPR future metadata). */
const SPEC119_METADATA =
  XML_DECLARATION +
  `<metadata xmlns="${MAIN_NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
  `<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" ` +
  `splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>` +
  `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">` +
  `<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
  `<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

/** The external workbook "[1]" names: Rates.xlsx, its Sheet1!A1 cached. */
const SPEC119_EXTERNAL_LINK =
  XML_DECLARATION +
  `<externalLink xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14" ` +
  `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><externalBook xmlns:r="${REL_NS}" r:id="rId1">` +
  `<sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>100</v></cell></row>` +
  `</sheetData></sheetDataSet></externalBook></externalLink>`;

const SPEC119_EXTERNAL_LINK_RELS =
  XML_DECLARATION +
  `<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_TYPE}/externalLinkPath" Target="Rates.xlsx" TargetMode="External"/></Relationships>`;

/** One of spec119-formulas.xlsx's Excel-shaped rows. */
const spec119Row = (r: number, spans: string, cells: string): string => `<row r="${r}" spans="${spans}" x14ac:dyDescent="0.35">${cells}</row>`;

/** A formula cell as Excel writes it: `t` for a string or error result. */
function formulaCell(ref: string, f: string, v: string | number, options: { t?: "str" | "e"; cm?: number } = {}): string {
  return `<c r="${ref}"${options.cm === undefined ? "" : ` cm="${options.cm}"`}${options.t ? ` t="${options.t}"` : ""}>${f}<v>${v}</v></c>`;
}

/** A shared formula's follower, as Excel writes one. */
const follower = (ref: string, si: number, v: number | string, t?: "e"): string =>
  `<c r="${ref}"${t ? ` t="${t}"` : ""}><f t="shared" si="${si}"/><v>${v}</v></c>`;

/** A normal formula's `<f>`. */
const normalF = (text: string): string => `<f>${xmlText(text)}</f>`;

export async function buildSpec119FormulasXlsxBytes(): Promise<Uint8Array> {
  const W = SPEC119_WIRE;
  const text = (ref: string, value: (typeof SPEC119_STRINGS)[number]): string => sharedCell(ref, SPEC119_STRINGS.indexOf(value));
  const calcRows = [
    spec119Row(1, "1:19", text("A1", "Value") + text("B1", "Amount") + numberCell("F1", 2)),
    spec119Row(
      2,
      "1:19",
      numberCell("A2", 3) + numberCell("B2", 10) +
        `<c r="C2"><f t="shared" ref="C2:C9" si="0">${xmlText(W.C2)}</f><v>20</v></c>` +
        // A non-Excel producer's order: this follower comes before its master (D3).
        follower("D2", 1, 0) +
        `<c r="E2"><f t="array" ref="E2:E4">${xmlText(W.E2)}</f><v>6</v></c>` +
        numberCell("F2", 1) +
        formulaCell("G2", normalF(W.G2), "big", { t: "str" }) +
        formulaCell("H2", `<f t="array" ref="H2:H6">${xmlText(W.H2)}</f>`, 1, { cm: 1 }) +
        formulaCell("I2", normalF(W.I2), 10) +
        formulaCell("J2", normalF(W.J2), 3) +
        formulaCell("L2", normalF(W.L2), 10) +
        formulaCell("M2", normalF(W.M2), 6) +
        formulaCell("N2", normalF(W.N2), 42) +
        formulaCell("O2", normalF(W.O2), 17) +
        formulaCell("P2", normalF(W.P2), 101) +
        formulaCell("Q2", normalF(W.Q2), "#REF!", { t: "e" }) +
        formulaCell("R2", normalF(W.R2), "low", { t: "str" }) +
        formulaCell("S2", normalF(W.S2), 3),
    ),
    spec119Row(
      3,
      "1:12",
      numberCell("A3", 1) + numberCell("B3", 20) + follower("C3", 0, 40) +
        `<c r="D3"><f t="shared" ref="D2:D5" si="1">${xmlText(W.D3)}</f><v>0</v></c>` +
        numberCell("E3", 2) + numberCell("H3", 2) + numberCell("K3", 1) +
        `<c r="L3">${SPEC119_FORMULAS_XLSX.dataTable.L3}<v>10</v></c>`,
    ),
    spec119Row(4, "1:12", numberCell("A4", 5) + numberCell("B4", 30) + follower("C4", 0, 60) + follower("D4", 1, 0) + numberCell("E4", 10) + numberCell("H4", 3) + numberCell("K4", 2) + numberCell("L4", 20)),
    spec119Row(5, "1:12", numberCell("A5", 2) + formulaCell("B5", normalF(W.B5), 60) + follower("C5", 0, 120) + follower("D5", 1, "#REF!", "e") + numberCell("H5", 4) + numberCell("K5", 3) + numberCell("L5", 30)),
    spec119Row(6, "1:8", numberCell("A6", 4) + follower("C6", 0, 0) + numberCell("H6", 5)),
    spec119Row(7, "3:3", follower("C7", 0, 0)),
    spec119Row(8, "3:3", follower("C8", 0, 0)),
    spec119Row(9, "3:3", follower("C9", 0, 0)),
  ].join("");
  const sheet = (uid: string, dimension: string, view: string, rows: string): string =>
    XML_DECLARATION +
    worksheetOpen(uid) +
    `<dimension ref="${dimension}"/><sheetViews>${view}</sheetViews><sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>` +
    `<sheetData>${rows}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  const calc = sheet(
    "4A6C8E0B-2D4F-4A6C-8E0B-2D4F6A8C0E19",
    "A1:S9",
    `<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B5" sqref="B5"/></sheetView>`,
    calcRows,
  );
  const jan = sheet(
    "5B7D9F1C-3E5A-4B7D-9F1C-3E5A7B9D1F20",
    "B2:C2",
    `<sheetView showFormulas="1" workbookViewId="0"><selection activeCell="C2" sqref="C2"/></sheetView>`,
    spec119Row(2, "2:3", numberCell("B2", 1) + formulaCell("C2", normalF("B2*2"), 2)),
  );
  const plain = (uid: string, ref: string, value: number): string =>
    sheet(uid, ref, `<sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView>`, spec119Row(Number(ref.slice(1)), `${ref.charCodeAt(0) - 64}:${ref.charCodeAt(0) - 64}`, numberCell(ref, value)));
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${SPEC119_STRINGS.length}" uniqueCount="${SPEC119_STRINGS.length}">` +
    SPEC119_STRINGS.map((value) => `<si><t>${value}</t></si>`).join("") +
    `</sst>`;
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}"><c r="C2" i="1" l="1"/><c r="C3"/><c r="C4"/><c r="B5"/><c r="C5"/><c r="C6"/><c r="C7"/><c r="C8"/><c r="C9"/>` +
    `<c r="D3"/><c r="D2"/><c r="D4"/><c r="D5"/><c r="E2" a="1"/><c r="G2"/><c r="H2" a="1"/><c r="I2"/><c r="J2"/><c r="L2"/><c r="L3"/>` +
    `<c r="M2"/><c r="N2"/><c r="O2"/><c r="P2"/><c r="Q2"/><c r="R2"/><c r="S2"/><c r="C2" i="2"/></calcChain>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Formulas",
    sheets: [
      { name: "Calc", xml: calc },
      { name: "Jan", xml: jan },
      { name: "Feb", xml: plain("6C8E0A2D-4F6B-4C8E-8A2D-4F6B8C0E2A21", "B2", 2) },
      { name: "Mar", xml: plain("7D9F1B3E-5A7C-4D9F-9B3E-5A7C9D1F3B22", "B2", 3) },
      { name: "O'Brien", xml: plain("8E0A2C4F-6B8D-4E0A-8C4F-6B8D0E2A4C23", "A1", 21) },
    ],
    sharedStrings,
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    additions: {
      parts: [
        part("xl/externalLinks/externalLink1.xml", SPEC119_EXTERNAL_LINK),
        part("xl/externalLinks/_rels/externalLink1.xml.rels", SPEC119_EXTERNAL_LINK_RELS),
        part("xl/metadata.xml", SPEC119_METADATA),
      ],
      relationships: [
        `<Relationship Id="rId20" Type="${REL_TYPE}/externalLink" Target="externalLinks/externalLink1.xml"/>`,
        `<Relationship Id="rId21" Type="${REL_TYPE}/sheetMetadata" Target="metadata.xml"/>`,
      ],
      overrides: [
        `<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`,
        `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>`,
      ],
      workbookAfterSheets: `<externalReferences><externalReference r:id="rId20"/></externalReferences>`,
    },
  });
}

/** spec119-formulas.ods's "Calc" `table:formula` values, as Calc wrote them. */
const SPEC119_ODS_WIRE = {
  B5: "of:=SUM([.B2:.B4])",
  B6: "of:=SUM([.B2:.B3]~[.B4])",
  C6: "of:=SUM([.A2:.B4]![.B2:.C3])",
  B7: "of:=[$'My Sheet'.A1]+[$Sheet2.B1]",
  B8: "of:=SUM({1;2|3;4})",
  C8: 'of:=COM.MICROSOFT.CONCAT([.A2];"-";[.B2])',
  B9: "of:=[.B2:.B4]*2",
  B12: "of:=[.B2]*3",
  B13: "oooc:=SUM([.B2:.B3])",
  C13: "msoxl:=SUM(B2:B4)",
  B14: "of:=[.B2]+1",
} as const;

/** What spec119-formulas.ods holds, for the unit and e2e specs. */
export const SPEC119_FORMULAS_ODS = {
  tableNames: ["Calc", "My Sheet", "Sheet2"],
  wire: SPEC119_ODS_WIRE,
  /** "Calc"'s formula cells → the canonical formula the reader must give
   * (B12:C12 is one cell repeated: the text is copied, never shifted). */
  formulas: {
    B5: "SUM(B2:B4)",
    B6: "SUM((B2:B3,B4))",
    C6: "SUM(A2:B4 B2:C3)",
    B7: "'My Sheet'!A1+Sheet2!B1",
    B8: "SUM({1,2;3,4})",
    C8: 'CONCAT(A2,"-",B2)',
    B9: "B2:B4*2",
    B12: "B2*3",
    C12: "B2*3",
    B13: "SUM(B2:B3)",
    C13: "SUM(B2:B4)",
    B14: "B2+1",
  },
  values: { B5: 60, B6: 60, C6: 30, B7: 12, B8: 10, C8: "3-10", B9: 20, B10: 40, B11: 60, B12: 30, C12: 30, B13: 30, C13: 60, B14: null },
  /** The matrix formula B9:B11. */
  matrix: { B9: { top: 8, left: 1, bottom: 10, right: 1 } },
  formulaBar: { B5: "=SUM(B2:B4)", B7: "='My Sheet'!A1+Sheet2!B1", B10: "{=B2:B4*2}" },
} as const;

/** LibreOffice's root declarations plus `xmlns:msoxl`, which Calc declares
 * for formulas in Excel's syntax. */
const SPEC119_CONTENT_NAMESPACES = LO_CONTENT_NAMESPACES.replace(
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"`,
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:msoxl="http://schemas.microsoft.com/office/excel/formula"`,
);

/** A LibreOffice formula cell with a float result. */
const loFormula = (formula: string, value: number, extra = ""): string =>
  `<table:table-cell${extra} table:formula="${xmlEscape(formula)}" office:value-type="float" office:value="${value}" calcext:value-type="float"><text:p>${value}</text:p></table:table-cell>`;

export async function buildSpec119FormulasOdsBytes(): Promise<Uint8Array> {
  const W = SPEC119_ODS_WIRE;
  const skip = loEmpty(1);
  const calc =
    `<table:table table:name="Calc" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loString("Value") + loString("Amount") + loEmpty(16382)) +
    loRow(loFloat(3) + loFloat(10) + loEmpty(16382)) +
    loRow(loFloat(1) + loFloat(20) + loEmpty(16382)) +
    loRow(loFloat(5) + loFloat(30) + loEmpty(16382)) +
    loRow(skip + loFormula(W.B5, 60) + loEmpty(16382)) +
    loRow(skip + loFormula(W.B6, 60) + loFormula(W.C6, 30) + loEmpty(16381)) +
    loRow(skip + loFormula(W.B7, 12) + loEmpty(16382)) +
    loRow(
      skip +
        loFormula(W.B8, 10) +
        `<table:table-cell table:formula="${xmlEscape(W.C8)}" office:value-type="string" office:string-value="3-10" calcext:value-type="string"><text:p>3-10</text:p></table:table-cell>` +
        loEmpty(16381),
    ) +
    loRow(skip + loFormula(W.B9, 20, ` table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="3"`) + loEmpty(16382)) +
    loRow(skip + loFloat(40) + loEmpty(16382)) +
    loRow(skip + loFloat(60) + loEmpty(16382)) +
    loRow(skip + loFormula(W.B12, 30, ` table:number-columns-repeated="2"`) + loEmpty(16381)) +
    loRow(skip + loFormula(W.B13, 30) + loFormula(W.C13, 60) + loEmpty(16381)) +
    loRow(skip + `<table:table-cell table:formula="${xmlEscape(W.B14)}"/>` + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1048562) +
    `</table:table>`;
  const mySheet =
    `<table:table table:name="My Sheet" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loFloat(5) + loEmpty(16383)) +
    loRow(loEmpty(16384), "ro1", 1048575) +
    `</table:table>`;
  const sheet2 =
    `<table:table table:name="Sheet2" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(skip + loFloat(7) + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1048575) +
    `</table:table>`;
  return loPackage(
    loContent(calc + mySheet + sheet2, SPEC119_CONTENT_NAMESPACES),
    loMeta("Formulas", 3, 34),
    loSettings(SPEC119_FORMULAS_ODS.tableNames, "Calc"),
  );
}

// --- SPEC-120 -----------------------------------------------------------------------

/** spec120-calc.xlsx's conformance rows (SPEC-120 "Conformance"): the
 * self-contained ones — no cell inputs — one per row of "Calc"'s column G,
 * each cached with the value Excel gives it. */
export const SPEC120_CONFORMANCE: readonly { ref: string; formula: string; value: number | string | boolean | { error: "#VALUE!" | "#NUM!" | "#N/A" } }[] = [
  { ref: "G1", formula: '"1"+"2"', value: 3 },
  { ref: "G2", formula: '"1"+"$4.00"', value: 5 },
  { ref: "G3", formula: '"6/1/2001"-"5/1/2001"', value: 31 },
  { ref: "G4", formula: 'SQRT("8+1")', value: { error: "#VALUE!" } },
  { ref: "G5", formula: '"A"&TRUE', value: "ATRUE" },
  { ref: "G6", formula: '1+""', value: { error: "#VALUE!" } },
  { ref: "G7", formula: "TRUE+TRUE", value: 2 },
  { ref: "G8", formula: '1="1"', value: false },
  { ref: "G9", formula: '1<"1"', value: true },
  { ref: "G10", formula: '"z"<TRUE', value: true },
  { ref: "G11", formula: '"A"="a"', value: true },
  { ref: "G12", formula: "0.1+0.2=0.3", value: true },
  { ref: "G13", formula: "0.1+0.2-0.3", value: 0 },
  { ref: "G14", formula: "-2^2", value: 4 },
  { ref: "G15", formula: "2^3^2", value: 64 },
  { ref: "G16", formula: "0^0", value: { error: "#NUM!" } },
  { ref: "G17", formula: '1/3&""', value: "0.333333333333333" },
  { ref: "G18", formula: "#N/A+1/0", value: { error: "#N/A" } },
  { ref: "G19", formula: "IF(TRUE,1,1/0)", value: 1 },
  { ref: "G20", formula: 'SUM("3",TRUE)', value: 4 },
  { ref: "G21", formula: 'MAX("a")', value: { error: "#VALUE!" } },
  { ref: "G22", formula: "ROUND(2.5,0)", value: 3 },
  { ref: "G23", formula: "ROUND(1.005,2)", value: 1.01 },
  { ref: "G24", formula: "MOD(-3,2)", value: 1 },
  { ref: "G25", formula: "INT(-8.9)", value: -9 },
  { ref: "G26", formula: "CEILING(-2.5,2)", value: -2 },
  { ref: "G27", formula: "FLOOR(0.3,0.1)", value: 0.3 },
  { ref: "G28", formula: "DATE(2008,14,2)", value: 39846 },
  { ref: "G29", formula: "DATE(1900,2,29)", value: 60 },
  { ref: "G30", formula: "WEEKDAY(46303)", value: 5 },
  { ref: "G31", formula: 'VLOOKUP("a",{"c",1;"a",2},2,)', value: 2 },
  { ref: "G32", formula: 'MATCH("B*",{"apple","banana"},0)', value: 2 },
  { ref: "G33", formula: 'LEN("😀")', value: 2 },
  { ref: "G34", formula: "PROPER(\"it's\")", value: "It'S" },
  { ref: "G35", formula: 'TRIM("  a   b  ")', value: "a b" },
  { ref: "G36", formula: 'TEXT(0.5,"0%")', value: "50%" },
];

/** What spec120-calc.xlsx holds (Excel-shaped): "Calc" — A1:A3 1, 2, 3; B2
 * `SUM(A1:A3)` cached 999, deliberately wrong (the cache is trusted until
 * F9); D2 `TODAY()` (built-in format 14) cached 45000; E2 `_xll.CUSTOM(B2)`
 * cached 42, an add-in function no registry will have; the conformance rows
 * in column G — and "Loop": A1 `B1+1`, B1 `A1+1`, cached 0. Its calcPr has
 * no fullCalcOnLoad. */
export const SPEC120_CALC_XLSX = {
  sheetNames: ["Calc", "Loop"],
  sum: { ref: "B2", formula: "SUM(A1:A3)", cached: 999, computed: 6 },
  today: { ref: "D2", formula: "TODAY()", cached: 45000 },
  addIn: { ref: "E2", wire: "_xll.CUSTOM(B2)", cached: 42 },
  loop: { A1: "B1+1", B1: "A1+1" },
  conformance: SPEC120_CONFORMANCE,
} as const;

/** A conformance row's `<c>` as Excel writes it: `t` for text, boolean and
 * error results. */
function conformanceCell(row: (typeof SPEC120_CONFORMANCE)[number]): string {
  const { value } = row;
  if (typeof value === "string") {
    return formulaCell(row.ref, normalF(row.formula), xmlText(value), { t: "str" });
  }
  if (typeof value === "boolean") {
    return `<c r="${row.ref}" t="b">${normalF(row.formula)}<v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === "object") {
    return formulaCell(row.ref, normalF(row.formula), value.error, { t: "e" });
  }
  return formulaCell(row.ref, normalF(row.formula), value);
}

export async function buildSpec120CalcXlsxBytes(): Promise<Uint8Array> {
  const C = SPEC120_CALC_XLSX;
  const g = (r: number): string => conformanceCell(SPEC120_CONFORMANCE[r - 1]!);
  const rows = [
    spec119Row(1, "1:7", numberCell("A1", 1) + g(1)),
    spec119Row(
      2,
      "1:7",
      numberCell("A2", 2) +
        formulaCell(C.sum.ref, normalF(C.sum.formula), C.sum.cached) +
        `<c r="${C.today.ref}" s="1"><f>${C.today.formula}</f><v>${C.today.cached}</v></c>` +
        formulaCell(C.addIn.ref, normalF(C.addIn.wire), C.addIn.cached) +
        g(2),
    ),
    spec119Row(3, "1:7", numberCell("A3", 3) + g(3)),
    ...SPEC120_CONFORMANCE.slice(3).map((_, i) => spec119Row(i + 4, "7:7", g(i + 4))),
  ].join("");
  const sheet = (uid: string, dimension: string, selected: boolean, body: string): string =>
    XML_DECLARATION +
    worksheetOpen(uid) +
    `<dimension ref="${dimension}"/><sheetViews><sheetView ${selected ? 'tabSelected="1" ' : ""}workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/><cols><col min="4" max="4" width="10.6328125" customWidth="1"/></cols>` +
    `<sheetData>${body}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  const loop = spec119Row(1, "1:2", formulaCell("A1", normalF(C.loop.A1), 0) + formulaCell("B1", normalF(C.loop.B1), 0));
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}"><c r="B2" i="1" l="1"/><c r="D2"/><c r="E2"/>` +
    SPEC120_CONFORMANCE.map((row) => `<c r="${row.ref}"/>`).join("") +
    `<c r="A1" i="2" l="1"/><c r="B1"/></calcChain>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Calculation",
    sheets: [
      { name: "Calc", xml: sheet("9F1B3D5E-7A9C-4E1B-8D3F-5A7C9E1B3D24", `A1:G${SPEC120_CONFORMANCE.length}`, true, rows) },
      { name: "Loop", xml: sheet("0A2C4E6F-8B0D-4F2A-9C4E-6B8D0F2A4C25", "A1:B1", false, loop) },
    ],
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
  });
}

/** What spec120-fullcalc.xlsx holds: an XlsxWriter-shaped workbook — as a
 * library that never calculates writes one — with `fullCalcOnLoad="1"` and
 * every formula cached as 0. "Totals": A1:A3 labels, B1:B3 10/20/30,
 * C1:C3 `B*2`, row 4 the totals. */
export const SPEC120_FULLCALC_XLSX = {
  sheetName: "Totals",
  labels: ["North", "South", "West", "Total"],
  numbers: { B1: 10, B2: 20, B3: 30 },
  formulas: { C1: "B1*2", C2: "B2*2", C3: "B3*2", B4: "SUM(B1:B3)", C4: "SUM(C1:C3)", D4: "AVERAGE(B1:B3)" },
  /** What each formula computes to (every one is cached 0). */
  computed: { C1: 20, C2: 40, C3: 60, B4: 60, C4: 120, D4: 20 },
} as const;

/** XlsxWriter's declaration (no standalone attribute line break quirks). */
const XLSXWRITER_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

export async function buildSpec120FullCalcXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC120_FULLCALC_XLSX;
  const label = (ref: string, index: number): string => `<c r="${ref}" t="s"><v>${index}</v></c>`;
  const zero = (ref: keyof typeof F.formulas): string => `<c r="${ref}"><f>${F.formulas[ref]}</f><v>0</v></c>`;
  const sheet =
    XLSXWRITER_DECLARATION +
    `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="A1:D4"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/><sheetData>` +
    `<row r="1" spans="1:3">${label("A1", 0)}<c r="B1"><v>${F.numbers.B1}</v></c>${zero("C1")}</row>` +
    `<row r="2" spans="1:3">${label("A2", 1)}<c r="B2"><v>${F.numbers.B2}</v></c>${zero("C2")}</row>` +
    `<row r="3" spans="1:3">${label("A3", 2)}<c r="B3"><v>${F.numbers.B3}</v></c>${zero("C3")}</row>` +
    `<row r="4" spans="1:4">${label("A4", 3)}${zero("B4")}${zero("C4")}${zero("D4")}</row>` +
    `</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  const strings =
    XLSXWRITER_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${F.labels.length}" uniqueCount="${F.labels.length}">` +
    F.labels.map((text) => `<si><t>${text}</t></si>`).join("") +
    `</sst>`;
  const styles =
    XLSXWRITER_DECLARATION +
    `<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
    `<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  const workbook =
    XLSXWRITER_DECLARATION +
    `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>` +
    `<workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>` +
    `<sheets><sheet name="${F.sheetName}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>`;
  const contentTypes =
    XLSXWRITER_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>` +
    `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`;
  return zip([
    part("[Content_Types].xml", contentTypes),
    part(
      "_rels/.rels",
      XLSXWRITER_DECLARATION +
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_TYPE}/officeDocument" Target="xl/workbook.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_TYPE}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
    ),
    part(
      "docProps/app.xml",
      XLSXWRITER_DECLARATION +
        `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
        `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application>` +
        `<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><AppVersion>12.0000</AppVersion></Properties>`,
    ),
    part(
      "docProps/core.xml",
      XLSXWRITER_DECLARATION +
        `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
        `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
        `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${FIXTURE_WORKBOOK_AUTHOR}</dc:creator>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-08T12:00:00Z</dcterms:created>` +
        `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-08T12:00:00Z</dcterms:modified></cp:coreProperties>`,
    ),
    part("xl/workbook.xml", workbook),
    part(
      "xl/_rels/workbook.xml.rels",
      XLSXWRITER_DECLARATION +
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_TYPE}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL_TYPE}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_TYPE}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    ),
    part("xl/styles.xml", styles),
    part("xl/sharedStrings.xml", strings),
    part("xl/worksheets/sheet1.xml", sheet),
  ]);
}

/** What spec120-manual.xlsx holds (Excel-shaped, `calcMode="manual"`):
 * "Manual" — A1 5, A2 7, A3 `SUM(A1:A2)` cached 12. */
export const SPEC120_MANUAL_XLSX = {
  sheetName: "Manual",
  calcPr: `<calcPr calcId="191029" calcMode="manual"/>`,
} as const;

export async function buildSpec120ManualXlsxBytes(): Promise<Uint8Array> {
  const rows =
    spec119Row(1, "1:1", numberCell("A1", 5)) + spec119Row(2, "1:1", numberCell("A2", 7)) + spec119Row(3, "1:1", formulaCell("A3", normalF("SUM(A1:A2)"), 12));
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: "Manual",
    sheets: [
      {
        name: SPEC120_MANUAL_XLSX.sheetName,
        xml: worksheetXml({ uid: "1B3D5F7A-9C1E-4A3B-8D5F-7A9C1E3B5D26", dimension: "A1:A3", selected: true, sheetData: rows }),
      },
    ],
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    calcPr: SPEC120_MANUAL_XLSX.calcPr,
  });
}

/** What spec120-calc.ods holds (LibreOffice-shaped): explicit
 * `table:calculation-settings` — case-sensitive and wildcards — settings.xml's
 * AutoCalculate, and "Calc": A1:A3 1/2/3, B1:B3 "Apple"/"apple"/"APPLE", and
 * two formula cells without cached values, A4 `SUM(A1:A3)` and B4
 * `COUNTIF(B1:B3,"apple")` (1 when case-sensitive). */
export const SPEC120_CALC_ODS = {
  tableName: "Calc",
  calculationSettings:
    `<table:calculation-settings table:case-sensitive="true" table:automatic-find-labels="false" ` +
    `table:use-regular-expressions="false" table:use-wildcards="true"/>`,
  wire: { A4: "of:=SUM([.A1:.A3])", B4: 'of:=COUNTIF([.B1:.B3];"apple")' },
  formulas: { A4: "SUM(A1:A3)", B4: 'COUNTIF(B1:B3,"apple")' },
  computed: { A4: 6, B4: 1 },
} as const;

export async function buildSpec120CalcOdsBytes(): Promise<Uint8Array> {
  const O = SPEC120_CALC_ODS;
  const uncached = (wire: string): string => `<table:table-cell table:formula="${xmlEscape(wire)}"/>`;
  const calc =
    `<table:table table:name="${O.tableName}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loFloat(1) + loString("Apple") + loEmpty(16382)) +
    loRow(loFloat(2) + loString("apple") + loEmpty(16382)) +
    loRow(loFloat(3) + loString("APPLE") + loEmpty(16382)) +
    loRow(uncached(O.wire.A4) + uncached(O.wire.B4) + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1048572) +
    `</table:table>`;
  return loPackage(
    loContent(calc, LO_CONTENT_NAMESPACES, O.calculationSettings),
    loMeta("Calculation", 1, 8),
    loSettings([O.tableName], O.tableName, `<config:config-item config:name="AutoCalculate" config:type="boolean">true</config:config-item>`),
  );
}
