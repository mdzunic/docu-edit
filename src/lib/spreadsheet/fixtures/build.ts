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
// - spec121-links.xlsx (SPEC-121) — Excel 365-shaped: the reference web the
//   sheet commands rewrite — cross-sheet, 3-D and VLOOKUP formulas, defined
//   names (workbook- and sheet-scoped, _xlnm.Print_Area), a chart, legacy +
//   threaded comments with VML and persons, a table, conditional formatting,
//   a validation and a hyperlink naming another sheet, a theme tab colour, a
//   hidden and a very hidden sheet, calcChain, TitlesOfParts.
//   spec121-links.ods — its LibreOffice-shaped twin (named range/expression,
//   database range, validation, conditional format, annotation, link, an
//   embedded chart object, settings.xml TabColor). SPEC121_LINKS_XLSX /
//   SPEC121_LINKS_ODS record the facts.
// - spec122-shapes.xlsx (SPEC-122) — Excel 365-shaped: what a rebuilt sheet
//   must keep. "Edit" (the active tab; B2 a plain General number C2 reads):
//   shared-formula blocks, a legacy array, a dynamic array (`cm="1"`), a data
//   table, rich, phonetic (`ph="1"`) and `_x000D_` shared strings, a picture
//   in a cell (`vm="1"`, xl/richData), rows with `spans`, `x14ac:dyDescent`
//   and `outlineLevel`, grouped custom-width columns, a merge, a frozen row,
//   formats 14, 22, 164 and 165, a quote-prefixed number with its
//   `ignoredErrors`, a classic and an x14 data-bar conditional format,
//   external and internal hyperlinks, a page setup with printer settings;
//   "Locked" (`sheetProtection`, Excel's SHA-512 hash, one unlocked cell);
//   "Inline" (openpyxl-shaped inline strings). spec122-shapes.ods — the
//   LibreOffice 24.8-shaped twin: the 1048553-row / 16380-column tails,
//   repeated non-empty runs, a 2×2 span, an annotation, a link, header rows
//   and a row group, table-local names and conditional formats, a validated
//   cell, a red-negative `style:map` pair, a `de-ce1` style, and a protected
//   table with one `style:cell-protect="none"` cell. SPEC122_SHAPES_XLSX /
//   SPEC122_SHAPES_ODS record the facts.
// - spec123-styles.xlsx (SPEC-123) — Excel 365-shaped (Aptos, the 2023
//   theme): named cell styles as Excel writes them — cellStyleXfs Normal,
//   Good, Heading 1, the workbook's own "Invoice total" and a hidden
//   Followed Hyperlink with Excel's `applyX="0"` exclusions, cellStyles in
//   Excel's alphabetical order with `builtinId`s (the custom one an
//   `xr:uid` instead), cellXfs naming them by `xfId`; tinted theme colours,
//   custom numFmt 164, a dxf and Excel's x14/x15 extLst. "Budget" (active):
//   Heading 1 headers (F1 empty), 0.00 and 0.00% columns, Good and
//   Good-centred texts, no row 5, row 6 a `<row s customFormat="1">` total
//   with SUMs, column H a `<col style>`, and H6 — where the styled row
//   meets the styled column — the style-only cell Excel stores there.
//   "Notes": two texts no formatting flow touches.
// - spec123-styles.ods — the LibreOffice 24.8-shaped twin: styles.xml's
//   common Default, Heading, Heading_20_1 ("Heading 1"), Good and
//   Invoice_20_total ("Invoice total"); content.xml's ce styles each
//   parented to one of them, N104/N111 and the red-minus N105/N105P0 pair
//   through `style:map`; column H's `table:default-cell-style-name`; row 6
//   formatted to XFD as a styled `number-columns-repeated` run.
//   SPEC123_STYLES_XLSX / SPEC123_STYLES_ODS record the facts.
// - spec124-structure.xlsx / .ods (SPEC-124) and spec125-fill.xlsx / .ods
//   (SPEC-125) — see their own sections below.
// - spec128-filter.xlsx (SPEC-128) — Excel 365-shaped on the classic theme:
//   "Sales" (active) a 12-record list under an `autoFilter` (`xr:uid`, a
//   `filters blank="1"` values column, a `customFilters and="1"` pair, an
//   `autoFilter/sortState`), `sheetPr filterMode="1"`, four filter-hidden
//   rows and one user-hidden row every criterion passes, half its rows
//   written without `r`; "Colours" an autoFilter with a `colorFilter` and
//   its dxf; "Log" a worksheet `sortState` sorting by cell colour; a hidden
//   `_xlnm._FilterDatabase` per autofilter. spec128-filter.ods — the
//   LibreOffice 24.8-shaped twin of "Sales": `__Anonymous_Sheet_DB__0` with
//   `table:filter-set-item`s, `table:visibility="filter"` and `"collapse"`
//   rows, the trailing repeated rows. SPEC128_FILTER_XLSX /
//   SPEC128_FILTER_ODS record the facts.
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
  /** SPEC-121: content-type Defaults Excel writes before `rels` (the `vml`
   * extension), and the docProps/app.xml text (default: appXml of the sheet
   * names). */
  defaults?: string[];
  app?: string;
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
    (options.defaults ?? []).join("") +
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
  entries.push(
    part("docProps/core.xml", coreXml(options.title)),
    part("docProps/app.xml", options.app ?? appXml(options.sheets.map((sheet) => sheet.name))),
  );
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

function loMeta(title: string, tableCount: number, cellCount: number, objectCount = 0): string {
  return (
    ODF_XML_DECLARATION +
    `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
    `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
    `xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.3">` +
    `<office:meta><meta:creation-date>2026-10-01T09:00:00.123456789</meta:creation-date><dc:date>2026-10-08T17:30:00.987654321</dc:date>` +
    `<meta:editing-duration>PT12M34S</meta:editing-duration><meta:editing-cycles>3</meta:editing-cycles>` +
    `<meta:generator>LibreOffice/24.8.2.1$Linux_X86_64 LibreOffice_project/0f794b6e29741098670a3b95d60478a65d05ef13</meta:generator>` +
    `<dc:title>${title}</dc:title><meta:initial-creator>${FIXTURE_WORKBOOK_AUTHOR}</meta:initial-creator><dc:creator>${FIXTURE_WORKBOOK_AUTHOR}</dc:creator>` +
    `<meta:document-statistic meta:table-count="${tableCount}" meta:cell-count="${cellCount}" meta:object-count="${objectCount}"/></office:meta></office:document-meta>`
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
 * `mimetype` first and STOREd; `styles` is styles.xml (default LO_STYLES —
 * every pre-SPEC-123 fixture's bytes). */
function loPackage(content: string, meta: string, settings: string, styles = LO_STYLES): Promise<Uint8Array> {
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    part("meta.xml", meta),
    part("settings.xml", settings),
    part("content.xml", content),
    { name: "Thumbnails/thumbnail.png", data: THUMBNAIL_PNG },
    part("styles.xml", styles),
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

/** XlsxWriter's shared-string cell. */
const xlsxWriterLabel = (ref: string, index: number): string => `<c r="${ref}" t="s"><v>${index}</v></c>`;

export async function buildSpec120FullCalcXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC120_FULLCALC_XLSX;
  const label = xlsxWriterLabel;
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

// ============================================================================
// SPEC-121 — spec121-links.xlsx / spec121-links.ods: the reference web
// ============================================================================
// The workbooks the sheet commands are proven on: every place a sheet name
// lives outside a cell, in the shape its producer writes it.
// - spec121-links.xlsx — Excel 365-shaped. "Summary" (the active tab:
//   cross-sheet, 3-D and VLOOKUP formulas, every cached value what the
//   formula computes); "Data" (a theme + tint tab colour; a drawing whose
//   chart's `c:f` texts name Data, with Excel's chart style and colour parts;
//   a threaded comment, its legacy placeholder comment and VML shape, and the
//   workbook's persons part; Table1 over A1:B5; conditional formatting and a
//   list validation naming Lookup; an internal hyperlink); "Lookup";
//   "Archive" (hidden); "Vault" (veryHidden). workbook.xml carries
//   workbook-scoped names on Data and Lookup, a Data-scoped name and
//   `_xlnm.Print_Area`, sorted as Excel sorts them; calcChain.xml; app.xml
//   with the Worksheets and Named Ranges heading pairs.
// - spec121-links.ods — the LibreOffice 24.8-shaped twin: `of:=` formulas
//   with `[$Data.B2]` references; a named range, a named expression and a
//   database range (and a Data-local named range); a content validation in
//   the prefix; on Data a conditional format naming Lookup, an annotation, a
//   `text:a` internal link and an embedded chart (Object 1/,
//   ObjectReplacements/Object 1) framed in `table:shapes` with
//   `xml:id`/`draw:id`; Archive hidden by its table style; settings.xml's
//   ActiveTable and per-table entries, Data's carrying TabColor beside the
//   style's `tableooo:tab-color`.
// SPEC121_LINKS_XLSX and SPEC121_LINKS_ODS record the facts the specs use.

/** spec121-links.xlsx's shared strings, in index order. */
const SPEC121_STRINGS = ["Code", "Target", "Double", "3-D sum", "Direct", "Amount", "Lookup!A1", "Archived", "Secret"] as const;

/** The threaded comment's id — Excel gives the legacy placeholder comment
 * the same GUID as its `xr:uid` and its author `tc=…`. */
const SPEC121_THREAD_ID = "{B0D2F4B5-9B1D-4FD0-8CF4-B59B1DD05C3B}";
const SPEC121_PERSON_ID = "{C1E3A5C6-AC2E-4A1E-9DA5-C6AC2E1E6D3C}";
const SPEC121_COMMENT_TEXT = "Checked against the ledger.";

const SPEC121_THREADED_COMMENT_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
const SPEC121_PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person";
const SPEC121_CHART_STYLE_REL = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const SPEC121_CHART_COLORS_REL = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

/** What spec121-links.xlsx holds, for the unit and e2e specs. Cell
 * addresses are A1; `sheetId`s and `localSheetId`s as written. */
export const SPEC121_LINKS_XLSX = {
  title: "Links",
  sheets: [
    { name: "Summary", sheetId: 1, rId: "rId1", partName: "xl/worksheets/sheet1.xml", visibility: "visible", uid: "{1C3E5A7B-9D1F-4B3C-8E5A-7B9D1F3C5E31}" },
    { name: "Data", sheetId: 2, rId: "rId2", partName: "xl/worksheets/sheet2.xml", visibility: "visible", uid: "{2D4F6B8C-0E2A-4C4D-9F6B-8C0E2A4D6F32}" },
    { name: "Lookup", sheetId: 3, rId: "rId3", partName: "xl/worksheets/sheet3.xml", visibility: "visible", uid: "{3E5A7C9D-1F3B-4D5E-8A7C-9D1F3B5E7A33}" },
    { name: "Archive", sheetId: 4, rId: "rId4", partName: "xl/worksheets/sheet4.xml", visibility: "hidden", uid: "{4F6B8DAE-2A4C-4E6F-9B8D-AE2A4C6F8B34}" },
    { name: "Vault", sheetId: 5, rId: "rId5", partName: "xl/worksheets/sheet5.xml", visibility: "veryHidden", uid: "{5A7C9EBF-3B5D-4F7A-8C9E-BF3B5D7A9C35}" },
  ],
  sheetNames: ["Summary", "Data", "Lookup", "Archive", "Vault"],
  /** Summary — the only sheetView with tabSelected="1"; no workbookView@activeTab. */
  activeSheetIndex: 0,
  /** Every stored cell, per sheet (a formula cell's value is its cache). */
  values: {
    Summary: { A1: "Code", B1: "Target", C1: "Double", D1: "3-D sum", E1: "Direct", A2: 103, B2: 360, C2: 200, D2: 101, E2: 240 },
    Data: { A1: "Code", B1: "Amount", D1: "Lookup!A1", A2: 101, B2: 100, A3: 102, B3: 200, A4: 103, B4: 300, A5: 104, B5: 400 },
    Lookup: { A1: 101, B1: 120, A2: 102, B2: 240, A3: 103, B3: 360, A4: 104, B4: 480, A5: 105, B5: 600 },
    Archive: { A1: "Archived", B1: 2019 },
    Vault: { A1: "Secret", B1: 42 },
  },
  /** Summary's formula cells: the `<f>` text, the canonical formula, the
   * cached value (= what it computes; A2 holds the key 103). */
  formulas: {
    B2: { wire: "VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)", formula: "VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)", value: 360 },
    C2: { wire: "Data!B2*2", formula: "Data!B2*2", value: 200 },
    D2: { wire: "SUM(Data:Lookup!A1)", formula: "SUM(Data:Lookup!A1)", value: 101 },
    E2: { wire: "Lookup!B2", formula: "Lookup!B2", value: 240 },
  },
  /** Data's `<sheetPr><tabColor theme="5" tint="…"/>` (theme 5 = accent2 after Excel's swap). */
  tabColor: { sheet: "Data", theme: 5, slot: "accent2", tint: "0.39997558519241921" },
  /** workbook.xml's definedNames, in file order (Excel's sort). */
  definedNames: [
    { name: "_xlnm.Print_Area", localSheetId: 1, text: "Data!$A$1:$B$5" },
    { name: "Amounts", localSheetId: 1, text: "Data!$B$2:$B$5" },
    { name: "LookupRange", text: "Lookup!$A$1:$B$5" },
    { name: "Rate", text: "Data!$B$2" },
  ],
  calcPr: `<calcPr calcId="191029"/>`,
  /** docProps/app.xml: HeadingPairs and TitlesOfParts. */
  headingPairs: [
    ["Worksheets", 5],
    ["Named Ranges", 4],
  ],
  titlesOfParts: ["Summary", "Data", "Lookup", "Archive", "Vault", "Data!Print_Area", "Data!Amounts", "LookupRange", "Rate"],
  /** xl/calcChain.xml's cells (`i` = Summary's sheetId, written on the first). */
  calcChain: [
    { r: "B2", i: 1 },
    { r: "C2", i: 1 },
    { r: "D2", i: 1 },
    { r: "E2", i: 1 },
  ],
  /** xl/worksheets/_rels/sheet2.xml.rels, by id. */
  dataRelationships: {
    rId1: { type: `${REL_TYPE}/drawing`, target: "../drawings/drawing1.xml" },
    rId2: { type: `${REL_TYPE}/vmlDrawing`, target: "../drawings/vmlDrawing1.vml" },
    rId3: { type: `${REL_TYPE}/table`, target: "../tables/table1.xml" },
    rId4: { type: `${REL_TYPE}/comments`, target: "../comments1.xml" },
    rId5: { type: SPEC121_THREADED_COMMENT_REL, target: "../threadedComments/threadedComment1.xml" },
  },
  drawing: { partName: "xl/drawings/drawing1.xml", from: "D3", to: "F12", chartRId: "rId1" },
  chart: {
    partName: "xl/charts/chart1.xml",
    style: "xl/charts/style1.xml",
    colors: "xl/charts/colors1.xml",
    /** Every `c:f` text, in document order: series name, categories, values. */
    formulas: ["Data!$B$1", "Data!$A$2:$A$5", "Data!$B$2:$B$5"],
  },
  table: {
    partName: "xl/tables/table1.xml",
    id: 1,
    name: "Table1",
    displayName: "Table1",
    ref: "A1:B5",
    columns: ["Code", "Amount"],
    uid: "{8DAFC1E2-6E8A-4CAD-9FC1-E26E8AAD2F38}",
  },
  comment: { partName: "xl/comments1.xml", ref: "B2", uid: SPEC121_THREAD_ID, author: `tc=${SPEC121_THREAD_ID}`, text: SPEC121_COMMENT_TEXT },
  /** The legacy comment's VML: `o:idmap data` and the shape's id (also its `o:spid`). */
  vml: { partName: "xl/drawings/vmlDrawing1.vml", idmap: 1, shapeIds: ["_x0000_s1025"] },
  threadedComment: {
    partName: "xl/threadedComments/threadedComment1.xml",
    ref: "B2",
    id: SPEC121_THREAD_ID,
    personId: SPEC121_PERSON_ID,
    text: SPEC121_COMMENT_TEXT,
  },
  /** Related from the WORKBOOK, not from a sheet. */
  person: { partName: "xl/persons/person.xml", id: SPEC121_PERSON_ID, rId: "rId10", type: SPEC121_PERSON_REL },
  conditionalFormatting: { sheet: "Data", sqref: "B2:B5", formula: "B2>Lookup!$B$1", dxfId: 0 },
  dataValidation: { sheet: "Data", sqref: "A2:A5", formula1: "Lookup!$A$1:$A$5", uid: "{6B8DAFC0-4C6E-4A8B-9DAF-C04C6E8B0D36}" },
  hyperlink: { sheet: "Data", ref: "D1", location: "Lookup!A1", display: "Lookup!A1", uid: "{7C9EB0D1-5D7F-4B9C-8EB0-D15D7F9C1E37}" },
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
    "xl/worksheets/sheet4.xml",
    "xl/worksheets/sheet5.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/worksheets/_rels/sheet2.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/charts/chart1.xml",
    "xl/charts/style1.xml",
    "xl/charts/colors1.xml",
    "xl/charts/_rels/chart1.xml.rels",
    "xl/tables/table1.xml",
    "xl/drawings/vmlDrawing1.vml",
    "xl/comments1.xml",
    "xl/threadedComments/threadedComment1.xml",
    "xl/persons/person.xml",
    "xl/calcChain.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

/** Excel's styles part for the links workbook: Normal, a bold xf (1), the
 * Hyperlink cell style's xf (2) and the conditional format's dxf 0 (light
 * red fill, dark red text — Excel's preset). */
const SPEC121_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="3" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><u/><sz val="11"/><color theme="10"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="1"/></cellXfs>` +
  `<cellStyles count="2"><cellStyle name="Hyperlink" xfId="1" builtinId="8"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

/** A chart text block as Excel writes one (`c:txPr`), in the minor font. */
function spec121ChartText(size: number, lumMod: number, lumOff: number): string {
  return (
    `<c:txPr><a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>` +
    `<a:p><a:pPr><a:defRPr sz="${size}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1">` +
    `<a:lumMod val="${lumMod}"/><a:lumOff val="${lumOff}"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>` +
    `<a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`
  );
}

/** Excel's thin grey chart line (`c:spPr`). */
const SPEC121_CHART_LINE =
  `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/>` +
  `<a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`;

/** A `c:numCache` over numbers. */
function spec121NumCache(values: readonly number[]): string {
  return (
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`).join("") +
    `</c:numCache>`
  );
}

/** xl/charts/chart1.xml: a clustered column chart over Data!B2:B5 by
 * Data!A2:A5, its series named by Data!B1 — Excel 365's chartSpace. */
function spec121ChartXml(): string {
  const [name, categories, values] = SPEC121_LINKS_XLSX.chart.formulas;
  const data = SPEC121_LINKS_XLSX.values.Data;
  return (
    XML_DECLARATION +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:r="${REL_NS}" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice>` +
    `<mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>` +
    `<c:chart><c:title><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>${spec121ChartText(1400, 65000, 35000)}</c:title>` +
    `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>${name}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${data.B1}</c:v></c:pt>` +
    `</c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>` +
    `<c:invertIfNegative val="0"/>` +
    `<c:cat><c:numRef><c:f>${categories}</c:f>${spec121NumCache([data.A2, data.A3, data.A4, data.A5])}</c:numRef></c:cat>` +
    `<c:val><c:numRef><c:f>${values}</c:f>${spec121NumCache([data.B2, data.B3, data.B4, data.B5])}</c:numRef></c:val>` +
    `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">` +
    `<c16:uniqueId val="{00000000-5E7A-4B9C-8D1E-2F3A4B5C6D7E}"/></c:ext></c:extLst></c:ser>` +
    `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>` +
    `<c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1146207512"/><c:axId val="1146212312"/></c:barChart>` +
    `<c:catAx><c:axId val="1146207512"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    `${SPEC121_CHART_LINE}${spec121ChartText(900, 65000, 35000)}<c:crossAx val="1146212312"/><c:crosses val="autoZero"/><c:auto val="1"/>` +
    `<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="1146212312"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
    `<c:majorGridlines>${SPEC121_CHART_LINE}</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>` +
    `${spec121ChartText(900, 65000, 35000)}<c:crossAx val="1146207512"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">` +
    `<a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>` +
    `</c:chartSpace>`
  );
}

const spec121StyleFont = (lumMod: number, lumOff: number): string =>
  `<cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="${lumMod}"/><a:lumOff val="${lumOff}"/></a:schemeClr></cs:fontRef>`;
const spec121StyleRefs = (fill = `<cs:fillRef idx="0"/>`): string => `<cs:lnRef idx="0"/>${fill}<cs:effectRef idx="0"/>`;

/** xl/charts/style1.xml — Excel's chart style 201: one entry per chart
 * element, in CT_ChartStyle order. */
function spec121ChartStyleXml(): string {
  const font = spec121StyleFont;
  const refs = spec121StyleRefs;
  const text = (size: number): string => `${refs()}${font(65000, 35000)}<cs:defRPr sz="${size}" kern="1200"/>`;
  const line = `<cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>`;
  const series = `${refs(`<cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>`)}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>`;
  const entries: [string, string][] = [
    ["axisTitle", text(1000)],
    ["categoryAxis", `${text(900)}${line}`],
    ["chartArea", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></cs:spPr><cs:defRPr sz="1000"/>`],
    ["dataLabel", text(900)],
    ["dataLabelCallout", text(900)],
    ["dataPoint", `${series}<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>`],
    ["dataPoint3D", `${series}<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>`],
    ["dataPointLine", `${series}<cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr>`],
    ["dataPointMarker", `${series}<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>`],
    ["dataPointWireframe", `${series}<cs:spPr><a:ln w="9525" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr>`],
    ["dataTable", `${text(900)}${line}`],
    ["downBar", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></cs:spPr>`],
    ["dropLine", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["errorBar", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["floor", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>`],
    ["gridlineMajor", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["gridlineMinor", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["hiLoLine", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["leaderLine", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["legend", text(900)],
    ["plotArea", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>`],
    ["plotArea3D", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>`],
    ["seriesAxis", `${text(900)}${line}`],
    ["seriesLine", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${line}`],
    ["title", `${refs()}${font(65000, 35000)}<cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0"/>`],
    ["trendline", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></cs:spPr>`],
    ["trendlineLabel", text(900)],
    ["upBar", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></cs:spPr>`],
    ["valueAxis", text(900)],
    ["wall", `${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>`],
  ];
  return (
    XML_DECLARATION +
    `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201">` +
    entries
      .map(([element, body]) => (element === "dataPointWireframe" ? `<cs:dataPointMarkerLayout symbol="circle" size="5"/>` : "") + `<cs:${element}>${body}</cs:${element}>`)
      .join("") +
    `</cs:chartStyle>`
  );
}

/** xl/charts/colors1.xml — Excel's "colorful" palette 10. */
const SPEC121_CHART_COLORS =
  XML_DECLARATION +
  `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">` +
  `<a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/>` +
  `<a:schemeClr val="accent6"/><cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/>` +
  `<a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/>` +
  `<a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/>` +
  `<a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/>` +
  `<a:lumOff val="50000"/></cs:variation></cs:colorStyle>`;

/** xl/drawings/drawing1.xml: the chart's graphicFrame, anchored D3 → F12. */
const SPEC121_DRAWING =
  XML_DECLARATION +
  `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>5</xdr:col><xdr:colOff>609600</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>184150</xdr:rowOff></xdr:to>` +
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">` +
  `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{D2F4B6D7-BD3F-4B2F-8EB6-D7BD3F2F7E3D}"/></a:ext></a:extLst>` +
  `</xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL_NS}" r:id="${SPEC121_LINKS_XLSX.drawing.chartRId}"/>` +
  `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;

/** xl/drawings/vmlDrawing1.vml — the legacy comment's note shape, as Excel
 * writes VML: no declaration, CRLF line breaks. */
const SPEC121_VML = [
  `<xml xmlns:v="urn:schemas-microsoft-com:vml"`,
  ` xmlns:o="urn:schemas-microsoft-com:office:office"`,
  ` xmlns:x="urn:schemas-microsoft-com:office:excel">`,
  ` <o:shapelayout v:ext="edit">`,
  `  <o:idmap v:ext="edit" data="${SPEC121_LINKS_XLSX.vml.idmap}"/>`,
  ` </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"`,
  `  path="m,l,21600r21600,l21600,xe">`,
  `  <v:stroke joinstyle="miter"/>`,
  `  <v:path gradientshapeok="t" o:connecttype="rect"/>`,
  ` </v:shapetype><v:shape id="_x0000_s1025" o:spid="_x0000_s1025" type="#_x0000_t202" style='position:absolute;`,
  `  margin-left:107.25pt;margin-top:16pt;width:108pt;height:59.25pt;z-index:1;`,
  `  visibility:hidden' fillcolor="infoBackground [80]" o:insetmode="auto">`,
  `  <v:fill color2="infoBackground [80]"/>`,
  `  <v:shadow color="none [81]" obscured="t"/>`,
  `  <v:path o:connecttype="none"/>`,
  `  <v:textbox style='mso-direction-alt:auto'>`,
  `   <div style='text-align:left'></div>`,
  `  </v:textbox>`,
  `  <x:ClientData ObjectType="Note">`,
  `   <x:MoveWithCells/>`,
  `   <x:SizeWithCells/>`,
  `   <x:Anchor>`,
  `    2, 15, 1, 2, 4, 15, 5, 4</x:Anchor>`,
  `   <x:AutoFill>False</x:AutoFill>`,
  `   <x:Row>1</x:Row>`,
  `   <x:Column>1</x:Column>`,
  `  </x:ClientData>`,
  ` </v:shape></xml>`,
].join("\r\n");

/** xl/comments1.xml — the placeholder Excel writes beside a threaded
 * comment for older readers. */
const SPEC121_COMMENTS =
  XML_DECLARATION +
  `<comments xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors><author>${SPEC121_LINKS_XLSX.comment.author}</author></authors>` +
  `<commentList><comment ref="${SPEC121_LINKS_XLSX.comment.ref}" authorId="0" shapeId="0" xr:uid="${SPEC121_LINKS_XLSX.comment.uid}"><text><t>` +
  `[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file ` +
  `is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ${SPEC121_COMMENT_TEXT}` +
  `</t></text></comment></commentList></comments>`;

const SPEC121_THREADED_COMMENTS =
  XML_DECLARATION +
  `<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="${MAIN_NS}">` +
  `<threadedComment ref="${SPEC121_LINKS_XLSX.threadedComment.ref}" dT="2026-10-08T17:20:00.00" personId="${SPEC121_PERSON_ID}" ` +
  `id="${SPEC121_THREAD_ID}"><text>${SPEC121_COMMENT_TEXT}</text></threadedComment></ThreadedComments>`;

const SPEC121_PERSONS =
  XML_DECLARATION +
  `<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="${MAIN_NS}">` +
  `<person displayName="${FIXTURE_WORKBOOK_AUTHOR}" id="${SPEC121_PERSON_ID}" userId="${FIXTURE_WORKBOOK_AUTHOR}" providerId="None"/></personList>`;

/** xl/tables/table1.xml — Table1 over Data!A1:B5, header row on, no filter buttons. */
function spec121TableXml(): string {
  const T = SPEC121_LINKS_XLSX.table;
  return (
    XML_DECLARATION +
    `<table xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" ` +
    `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" ` +
    `id="${T.id}" xr:uid="${T.uid}" name="${T.name}" displayName="${T.displayName}" ref="${T.ref}" totalsRowShown="0">` +
    `<tableColumns count="${T.columns.length}"><tableColumn id="1" xr3:uid="{9EB0D2F3-7F9B-4DBE-8AD2-F37F9BBE3A39}" name="${T.columns[0]}"/>` +
    `<tableColumn id="2" xr3:uid="{AFC1E3A4-8A0C-4ECF-9BE3-A48A0CCF4B3A}" name="${T.columns[1]}"/></tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`
  );
}

/** docProps/app.xml with both heading pairs Excel writes when a workbook
 * has defined names: the sheets, then the named ranges. */
function spec121AppXml(): string {
  const F = SPEC121_LINKS_XLSX;
  const pairs = F.headingPairs.map(([heading, count]) => `<vt:variant><vt:lpstr>${heading}</vt:lpstr></vt:variant><vt:variant><vt:i4>${count}</vt:i4></vt:variant>`);
  return (
    XML_DECLARATION +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>` +
    `<HeadingPairs><vt:vector size="${pairs.length * 2}" baseType="variant">${pairs.join("")}</vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${F.titlesOfParts.length}" baseType="lpstr">${F.titlesOfParts.map((title) => `<vt:lpstr>${title}</vt:lpstr>`).join("")}` +
    `</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>` +
    `<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`
  );
}

/** A links-workbook worksheet in Excel 365's element order (CT_Worksheet:
 * sheetPr, dimension, sheetViews, sheetFormatPr, sheetData, …,
 * conditionalFormatting, dataValidations, hyperlinks, pageMargins, …,
 * drawing, legacyDrawing, tableParts). */
function spec121Worksheet(
  index: number,
  body: { sheetPr?: string; dimension: string; activeCell: string; rows: string; afterData?: string; tail?: string },
): string {
  const sheet = SPEC121_LINKS_XLSX.sheets[index]!;
  return (
    XML_DECLARATION +
    worksheetOpen(sheet.uid.slice(1, -1)) +
    (body.sheetPr ?? "") +
    `<dimension ref="${body.dimension}"/><sheetViews><sheetView ${index === SPEC121_LINKS_XLSX.activeSheetIndex ? 'tabSelected="1" ' : ""}workbookViewId="0">` +
    `<selection activeCell="${body.activeCell}" sqref="${body.activeCell}"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/><sheetData>${body.rows}</sheetData>${body.afterData ?? ""}` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${body.tail ?? ""}</worksheet>`
  );
}

// The facts' tab colour as attribute text (destructured: a fixture is not a
// colour reader — core/themeGuard.spec.ts).
const { theme: TAB_COLOR_THEME, tint: TAB_COLOR_TINT } = SPEC121_LINKS_XLSX.tabColor;

export async function buildSpec121LinksXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC121_LINKS_XLSX;
  const V = F.values;
  /** A literal cell: a number, or one of SPEC121_STRINGS by shared index. */
  const cell = (ref: string, value: number | string, style?: number): string => {
    if (typeof value === "number") {
      return numberCell(ref, value, style);
    }
    const index = SPEC121_STRINGS.indexOf(value as (typeof SPEC121_STRINGS)[number]);
    if (index < 0) {
      throw new Error(`"${value}" is not one of spec121-links.xlsx's strings`);
    }
    return sharedCell(ref, index, style);
  };
  const formula = (ref: keyof typeof F.formulas): string => formulaCell(ref, normalF(F.formulas[ref].wire), F.formulas[ref].value);
  const BOLD = 1;
  const HYPERLINK = 2;

  const summary = spec121Worksheet(0, {
    dimension: "A1:E2",
    activeCell: "C2",
    rows:
      excelRow(1, "1:5", (["A1", "B1", "C1", "D1", "E1"] as const).map((ref) => cell(ref, V.Summary[ref], BOLD)).join("")) +
      excelRow(2, "1:5", cell("A2", V.Summary.A2) + formula("B2") + formula("C2") + formula("D2") + formula("E2")),
  });

  const dataRows = [2, 3, 4, 5] as const;
  const data = spec121Worksheet(1, {
    sheetPr: `<sheetPr><tabColor theme="${TAB_COLOR_THEME}" tint="${TAB_COLOR_TINT}"/></sheetPr>`,
    dimension: "A1:D5",
    activeCell: "B2",
    rows:
      excelRow(1, "1:4", cell("A1", V.Data.A1) + cell("B1", V.Data.B1) + cell("D1", V.Data.D1, HYPERLINK)) +
      dataRows.map((r) => excelRow(r, "1:4", cell(`A${r}`, V.Data[`A${r}`]) + cell(`B${r}`, V.Data[`B${r}`]))).join(""),
    afterData:
      `<conditionalFormatting sqref="${F.conditionalFormatting.sqref}"><cfRule type="expression" dxfId="${F.conditionalFormatting.dxfId}" priority="1">` +
      `<formula>${xmlText(F.conditionalFormatting.formula)}</formula></cfRule></conditionalFormatting>` +
      `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${F.dataValidation.sqref}" ` +
      `xr:uid="${F.dataValidation.uid}"><formula1>${F.dataValidation.formula1}</formula1></dataValidation></dataValidations>` +
      `<hyperlinks><hyperlink ref="${F.hyperlink.ref}" location="${F.hyperlink.location}" display="${F.hyperlink.display}" xr:uid="${F.hyperlink.uid}"/></hyperlinks>`,
    tail: `<drawing r:id="rId1"/><legacyDrawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId3"/></tableParts>`,
  });

  const lookupRows = [1, 2, 3, 4, 5] as const;
  const lookup = spec121Worksheet(2, {
    dimension: "A1:B5",
    activeCell: "A1",
    rows: lookupRows.map((r) => excelRow(r, "1:2", cell(`A${r}`, V.Lookup[`A${r}`]) + cell(`B${r}`, V.Lookup[`B${r}`]))).join(""),
  });
  const archive = spec121Worksheet(3, { dimension: "A1:B1", activeCell: "A1", rows: excelRow(1, "1:2", cell("A1", V.Archive.A1) + cell("B1", V.Archive.B1)) });
  const vault = spec121Worksheet(4, { dimension: "A1:B1", activeCell: "A1", rows: excelRow(1, "1:2", cell("A1", V.Vault.A1) + cell("B1", V.Vault.B1)) });
  const worksheets = [summary, data, lookup, archive, vault];

  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${SPEC121_STRINGS.length}">` +
    SPEC121_STRINGS.map((text) => `<si><t>${text}</t></si>`).join("") +
    `</sst>`;
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}">` +
    F.calcChain.map((entry, index) => (index === 0 ? `<c r="${entry.r}" i="${entry.i}" l="1"/>` : `<c r="${entry.r}"/>`)).join("") +
    `</calcChain>`;
  const definedNames =
    `<definedNames>` +
    F.definedNames
      .map((name) => `<definedName name="${name.name}"${"localSheetId" in name ? ` localSheetId="${name.localSheetId}"` : ""}>${name.text}</definedName>`)
      .join("") +
    `</definedNames>`;
  const rels = (relationships: string): string => XML_DECLARATION + `<Relationships xmlns="${PKG_REL_NS}">${relationships}</Relationships>`;
  const R = F.dataRelationships;
  const relationship = (id: keyof typeof R): string => `<Relationship Id="${id}" Type="${R[id].type}" Target="${R[id].target}"/>`;

  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheets.map((sheet, index) => ({
      name: sheet.name,
      ...(sheet.visibility === "visible" ? {} : { state: sheet.visibility }),
      xml: worksheets[index]!,
    })),
    sharedStrings,
    calcChain,
    calcPr: F.calcPr,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    styles: SPEC121_STYLES,
    defaults: [`<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`],
    app: spec121AppXml(),
    additions: {
      parts: [
        part("xl/worksheets/_rels/sheet2.xml.rels", rels(relationship("rId3") + relationship("rId2") + relationship("rId1") + relationship("rId5") + relationship("rId4"))),
        part(F.drawing.partName, SPEC121_DRAWING),
        part("xl/drawings/_rels/drawing1.xml.rels", rels(`<Relationship Id="${F.drawing.chartRId}" Type="${REL_TYPE}/chart" Target="../charts/chart1.xml"/>`)),
        part(F.chart.partName, spec121ChartXml()),
        part(F.chart.style, spec121ChartStyleXml()),
        part(F.chart.colors, SPEC121_CHART_COLORS),
        part(
          "xl/charts/_rels/chart1.xml.rels",
          rels(
            `<Relationship Id="rId2" Type="${SPEC121_CHART_COLORS_REL}" Target="colors1.xml"/>` +
              `<Relationship Id="rId1" Type="${SPEC121_CHART_STYLE_REL}" Target="style1.xml"/>`,
          ),
        ),
        part(F.table.partName, spec121TableXml()),
        part(F.vml.partName, SPEC121_VML),
        part(F.comment.partName, SPEC121_COMMENTS),
        part(F.threadedComment.partName, SPEC121_THREADED_COMMENTS),
        part(F.person.partName, SPEC121_PERSONS),
      ],
      relationships: [`<Relationship Id="${F.person.rId}" Type="${F.person.type}" Target="persons/person.xml"/>`],
      overrides: [
        `<Override PartName="/${F.drawing.partName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
        `<Override PartName="/${F.chart.partName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
        `<Override PartName="/${F.chart.style}" ContentType="application/vnd.ms-office.chartstyle+xml"/>`,
        `<Override PartName="/${F.chart.colors}" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`,
        `<Override PartName="/${F.table.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
        `<Override PartName="/${F.comment.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
        `<Override PartName="/${F.threadedComment.partName}" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>`,
        `<Override PartName="/${F.person.partName}" ContentType="application/vnd.ms-excel.person+xml"/>`,
      ],
      workbookAfterSheets: definedNames,
    },
  });
}

// --- spec121-links.ods ---------------------------------------------------------

/** What spec121-links.ods holds, for the unit and e2e specs. */
export const SPEC121_LINKS_ODS = {
  title: "Links",
  tableNames: ["Summary", "Data", "Lookup", "Archive"],
  visibility: ["visible", "visible", "visible", "hidden"],
  /** Each table's automatic table style: ta2 carries the tab colour, ta3 `table:display="false"`. */
  tableStyles: ["ta1", "ta2", "ta1", "ta3"],
  activeTable: "Summary",
  activeSheetIndex: 0,
  /** Data's `tableooo:tab-color`, and the same colour as settings.xml's TabColor int. */
  tabColor: { table: "Data", styleName: "ta2", hex: "#f1a983", settingsValue: 0xf1a983 },
  values: {
    Summary: { A1: "Code", B1: "Target", C1: "Double", D1: "3-D sum", E1: "Direct", A2: 103, B2: 360, C2: 200, D2: 101, E2: 240 },
    Data: { A1: "Code", B1: "Amount", D1: "Lookup!A1", A2: 101, B2: 100, A3: 102, B3: 200, A4: 103, B4: 300, A5: 104, B5: 400 },
    Lookup: { A1: 101, B1: 120, A2: 102, B2: 240, A3: 103, B3: 360, A4: 104, B4: 480, A5: 105, B5: 600 },
    Archive: { A1: "Archived", B1: 2019 },
  },
  /** Summary's formula cells: `table:formula`, canonical formula, cached value. */
  formulas: {
    B2: { wire: "of:=VLOOKUP([.A2];[$Lookup.$A$1:.$B$5];2;0)", formula: "VLOOKUP(A2,Lookup!$A$1:$B$5,2,0)", value: 360 },
    C2: { wire: "of:=[$Data.B2]*2", formula: "Data!B2*2", value: 200 },
    D2: { wire: "of:=SUM([$Data.A1:$Lookup.A1])", formula: "SUM(Data:Lookup!A1)", value: 101 },
    E2: { wire: "of:=[$Lookup.B2]", formula: "Lookup!B2", value: 240 },
  },
  /** content.xml's suffix: `table:named-expressions`, then `table:database-ranges`. */
  namedRange: { name: "LookupRange", baseCellAddress: "$Lookup.$A$1", cellRangeAddress: "$Lookup.$A$1:.$B$5" },
  namedExpression: { name: "Rate", baseCellAddress: "$Data.$B$2", expression: "of:=[$Data.$B$2]" },
  databaseRange: { name: "Table1", targetRangeAddress: "Data.A1:Data.B5" },
  /** Data's own (sheet-local) named range, inside its table text. */
  localNamedRange: { table: "Data", name: "Amounts", baseCellAddress: "$Data.$B$2", cellRangeAddress: "$Data.$B$2:.$B$5" },
  /** content.xml's prefix: the list validation Data!A2:A5 use. */
  contentValidation: {
    name: "val1",
    condition: "of:cell-content-is-in-list([$Lookup.$A$1:.$A$5])",
    baseCellAddress: "Data.A2",
    cells: ["A2", "A3", "A4", "A5"],
  },
  conditionalFormat: {
    table: "Data",
    targetRangeAddress: "Data.B2:Data.B5",
    value: "formula-is([.B2]>[$Lookup.$B$1])",
    baseCellAddress: "Data.B2",
    applyStyleName: "Bad",
  },
  annotation: { table: "Data", a1: "B2", creator: FIXTURE_WORKBOOK_AUTHOR, date: "2026-10-08T17:20:00", text: SPEC121_COMMENT_TEXT },
  link: { table: "Data", a1: "D1", href: "#Lookup.A1", text: "Lookup!A1" },
  /** The embedded chart: its frame in Data's `table:shapes`, the object's
   * parts and the ranges its content.xml names. */
  chart: {
    frame: { xmlId: "id1", drawId: "id1", name: "Object 1", endCellAddress: "Data.F12" },
    notifyOnUpdateOfRanges: "Data.A1:Data.B5",
    href: "./Object 1",
    replacementHref: "./ObjectReplacements/Object 1",
    objectParts: ["Object 1/content.xml", "Object 1/styles.xml", "Object 1/meta.xml"],
    replacement: "ObjectReplacements/Object 1",
    plotAreaRange: "Data.A1:Data.B5",
    categoriesRange: "Data.A2:Data.A5",
    valuesRange: "Data.B2:Data.B5",
    labelCell: "Data.B1",
  },
  /** settings.xml's Tables map entries, in order (Data's carries TabColor). */
  settingsTables: ["Summary", "Data", "Lookup", "Archive"],
  /** Every entry, in package order (mimetype first, STOREd). */
  parts: [
    "mimetype",
    "meta.xml",
    "settings.xml",
    "content.xml",
    "Thumbnails/thumbnail.png",
    "Object 1/content.xml",
    "Object 1/styles.xml",
    "Object 1/meta.xml",
    "ObjectReplacements/Object 1",
    "styles.xml",
    "manifest.rdf",
    "META-INF/manifest.xml",
  ],
} as const;

/** content.xml's automatic styles: columns, rows, the three table styles,
 * the bold header, the chart frame's and the annotation's graphic and
 * paragraph styles — in Calc's order. */
const SPEC121_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb" ` +
  `tableooo:tab-color="${SPEC121_LINKS_ODS.tabColor.hex}"/></style:style>` +
  `<style:style style:name="ta3" style:family="table" style:master-page-name="Default"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
  `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center" ` +
  `draw:textarea-vertical-align="middle" draw:color-mode="standard" draw:luminance="0%" draw:contrast="0%" draw:gamma="100%" draw:red="0%" draw:green="0%" ` +
  `draw:blue="0%" fo:clip="rect(0in, 0in, 0in, 0in)" draw:image-opacity="100%" style:mirror="none"/></style:style>` +
  `<style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-color="#000000" draw:fill="solid" ` +
  `draw:fill-color="#ffffc0" draw:auto-grow-height="true" draw:auto-grow-width="false" fo:padding-top="0.0398in" fo:padding-bottom="0.0398in" ` +
  `fo:padding-left="0.0398in" fo:padding-right="0.0398in" draw:shadow="visible" draw:shadow-offset-x="0.0236in" draw:shadow-offset-y="0.0236in" ` +
  `draw:shadow-color="#000000"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>` +
  `<style:style style:name="P2" style:family="paragraph"><style:paragraph-properties fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0in" ` +
  `fo:margin-bottom="0in" fo:text-indent="0in" fo:border="none"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="9pt"/></style:style>` +
  `</office:automatic-styles>`;

/** styles.xml: Calc's common styles plus "Status" and its child "Bad", the
 * style the conditional format applies. */
const SPEC121_ODS_STYLES = LO_STYLES.replace(
  `</office:styles>`,
  `<style:style style:name="Status" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
    `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
    `<style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Status"><style:table-cell-properties fo:background-color="#ffcccc"/>` +
    `<style:text-properties fo:color="#cc0000"/></style:style></office:styles>`,
);

/** The chart object's content.xml: a bar chart over Data, its plot area,
 * categories, series values and label naming the sheet, and the local
 * table of cached values Calc writes beside it. */
function spec121OdsChartContent(): string {
  const C = SPEC121_LINKS_ODS.chart;
  const D = SPEC121_LINKS_ODS.values.Data;
  const localRows = ([2, 3, 4, 5] as const)
    .map(
      (r) =>
        `<table:table-row><table:table-cell office:value-type="string"><text:p>${D[`A${r}`]}</text:p></table:table-cell>` +
        `<table:table-cell office:value-type="float" office:value="${D[`B${r}`]}"><text:p>${D[`B${r}`]}</text:p></table:table-cell></table:table-row>`,
    )
    .join("");
  return (
    ODF_XML_DECLARATION +
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
    `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
    `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
    `xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
    `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
    `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:chartooo="http://openoffice.org/2010/chart" office:version="1.3">` +
    `<office:automatic-styles><number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
    `<style:style style:name="ch1" style:family="chart"><style:graphic-properties draw:stroke="none" draw:fill-color="#ffffff"/></style:style>` +
    `<style:style style:name="ch2" style:family="chart"><style:chart-properties chart:include-hidden-cells="false" chart:auto-position="true" ` +
    `chart:auto-size="true" chart:treat-empty-cells="leave-gap" chart:right-angled-axes="true"/></style:style>` +
    `<style:style style:name="ch3" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:display-label="true" ` +
    `chart:logarithmic="false" chart:reverse-direction="false" text:line-break="false" chart:link-data-style-to-source="true" chart:axis-position="0" ` +
    `chart:axis-label-position="outside-start" chart:tick-mark-position="at-labels"/><style:graphic-properties svg:stroke-color="#b3b3b3"/>` +
    `<style:text-properties fo:font-size="10pt"/></style:style>` +
    `<style:style style:name="ch4" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:display-label="true" ` +
    `chart:logarithmic="false" chart:reverse-direction="false" chart:link-data-style-to-source="true" chart:axis-position="start"/>` +
    `<style:graphic-properties svg:stroke-color="#b3b3b3"/><style:text-properties fo:font-size="10pt"/></style:style>` +
    `<style:style style:name="ch5" style:family="chart"><style:graphic-properties svg:stroke-color="#b3b3b3"/></style:style>` +
    `<style:style style:name="ch6" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:link-data-style-to-source="true"/>` +
    `<style:graphic-properties draw:fill-color="#004586"/></style:style>` +
    `<style:style style:name="ch7" style:family="chart"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>` +
    `<style:style style:name="ch8" style:family="chart"><style:graphic-properties svg:stroke-color="#b3b3b3" draw:fill="none"/></style:style>` +
    `</office:automatic-styles><office:body><office:chart>` +
    `<chart:chart svg:width="2.5in" svg:height="1.8in" xlink:href=".." xlink:type="simple" chart:class="chart:bar" chart:style-name="ch1">` +
    `<chart:plot-area chart:style-name="ch2" table:cell-range-address="${C.plotAreaRange}" chart:data-source-has-labels="both" ` +
    `svg:x="0.1in" svg:y="0.1in" svg:width="2.3in" svg:height="1.6in">` +
    `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch3"><chart:categories table:cell-range-address="${C.categoriesRange}"/></chart:axis>` +
    `<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch4"><chart:grid chart:style-name="ch5" chart:class="major"/></chart:axis>` +
    `<chart:series chart:style-name="ch6" chart:values-cell-range-address="${C.valuesRange}" chart:label-cell-address="${C.labelCell}" ` +
    `chart:class="chart:bar"><chart:data-point chart:repeated="4"/></chart:series>` +
    `<chart:wall chart:style-name="ch7"/><chart:floor chart:style-name="ch8"/></chart:plot-area>` +
    `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
    `<table:table-columns><table:table-column/></table:table-columns><table:table-header-rows><table:table-row>` +
    `<table:table-cell><text:p/></table:table-cell><table:table-cell office:value-type="string"><text:p>${D.B1}</text:p></table:table-cell>` +
    `</table:table-row></table:table-header-rows><table:table-rows>${localRows}</table:table-rows></table:table>` +
    `</chart:chart></office:chart></office:body></office:document-content>`
  );
}

const SPEC121_ODS_CHART_STYLES =
  ODF_XML_DECLARATION +
  `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:styles>` +
  `<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf" fo:wrap-option="no-wrap"/>` +
  `<style:paragraph-properties style:text-autospace="ideograph-alpha" style:punctuation-wrap="simple" style:line-break="strict" ` +
  `style:writing-mode="page" style:font-independent-line-spacing="false"/><style:text-properties style:use-window-font-color="true" ` +
  `fo:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" fo:font-size="10pt" fo:language="en" fo:country="US"/>` +
  `</style:default-style></office:styles></office:document-styles>`;

const SPEC121_ODS_CHART_META =
  ODF_XML_DECLARATION +
  `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
  `xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.3"><office:meta>` +
  `<meta:generator>LibreOffice/24.8.2.1$Linux_X86_64 LibreOffice_project/0f794b6e29741098670a3b95d60478a65d05ef13</meta:generator>` +
  `</office:meta></office:document-meta>`;

/** The chart's replacement image: a stand-in for the GDI metafile Calc
 * stores (its "VCLMTF" signature, then a few bytes). */
const SPEC121_ODS_REPLACEMENT = new Uint8Array([0x56, 0x43, 0x4c, 0x4d, 0x54, 0x46, 0x01, 0x00, 0x31, 0x00, 0x00, 0x00, 0x53, 0x50, 0x45, 0x43, 0x2d, 0x31, 0x32, 0x31]);

/** META-INF/manifest.xml: Calc's entries plus the chart object's, in
 * Calc's order (the object's files, its directory, its replacement). */
const SPEC121_ODS_MANIFEST = LO_MANIFEST.replace(
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>`,
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/meta.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
    `<manifest:file-entry manifest:full-path="ObjectReplacements/Object 1" ` +
    `manifest:media-type="application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;"/>`,
);

/** settings.xml: ActiveTable and one Tables entry per table — Data's
 * carrying TabColor (an RGB int) after ShowGrid. */
function spec121OdsSettings(): string {
  const O = SPEC121_LINKS_ODS;
  const view = {
    cursor: [0, 0] as [number, number],
    split: [0, 0] as [number, number],
    splitPosition: [0, 0] as [number, number],
    activeSplitRange: 2,
    position: { left: 0, right: 0, top: 0, bottom: 0 },
    zoom: 100,
    showGrid: true,
  };
  const showGrid = loConfigItem("ShowGrid", "boolean", true);
  const tables = O.settingsTables
    .map((name) => {
      const entry = spec117TableSettings(name, view);
      return name === O.tabColor.table ? entry.replace(showGrid, showGrid + loConfigItem("TabColor", "int", O.tabColor.settingsValue)) : entry;
    })
    .join("");
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
    loConfigItem("ActiveTable", "string", O.activeTable) +
    loConfigItem("HorizontalScrollbarWidth", "int", 1309) +
    loConfigItem("ZoomType", "short", 0) +
    loConfigItem("ZoomValue", "int", 100) +
    loConfigItem("PageViewZoomValue", "int", 60) +
    loConfigItem("ShowPageBreakPreview", "boolean", false) +
    loConfigItem("ShowZeroValues", "boolean", true) +
    loConfigItem("ShowNotes", "boolean", true) +
    loConfigItem("ShowGrid", "boolean", true) +
    loConfigItem("GridColor", "int", 12632256) +
    loConfigItem("ShowPageBreaks", "boolean", true) +
    loConfigItem("HasColumnRowHeaders", "boolean", true) +
    loConfigItem("HasSheetTabs", "boolean", true) +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>` +
    `<config:config-item-set config:name="ooo:configuration-settings">` +
    loConfigItem("IsDocumentShared", "boolean", false) +
    loConfigItem("SaveVersionOnClose", "boolean", false) +
    `</config:config-item-set></office:settings></office:document-settings>`
  );
}

export async function buildSpec121LinksOdsBytes(): Promise<Uint8Array> {
  const O = SPEC121_LINKS_ODS;
  const V = O.values;
  const formula = (ref: keyof typeof O.formulas): string => loFormula(O.formulas[ref].wire, O.formulas[ref].value);
  const cell = (value: number | string, style?: string): string => (typeof value === "number" ? loFloat(value) : loString(value, style));
  const columns = `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>`;
  const tail = (rowsWritten: number): string => loRow(loEmpty(16384), "ro1", 1_048_576 - rowsWritten);
  const table = (index: number, inner: string): string =>
    `<table:table table:name="${O.tableNames[index]}" table:style-name="${O.tableStyles[index]}">${inner}</table:table>`;

  const summary = table(
    0,
    columns +
      loRow((["A1", "B1", "C1", "D1", "E1"] as const).map((ref) => cell(V.Summary[ref], "ce1")).join("") + loEmpty(16379)) +
      loRow(cell(V.Summary.A2) + formula("B2") + formula("C2") + formula("D2") + formula("E2") + loEmpty(16379)) +
      tail(2),
  );

  const C = O.chart;
  const frame =
    `<draw:frame table:end-cell-address="${C.frame.endCellAddress}" table:end-x="0.5in" table:end-y="0.1in" draw:z-index="0" ` +
    `xml:id="${C.frame.xmlId}" draw:id="${C.frame.drawId}" draw:name="${C.frame.name}" draw:style-name="gr1" draw:text-style-name="P1" ` +
    `svg:width="2.5in" svg:height="1.8in" svg:x="2.67in" svg:y="0.36in">` +
    `<draw:object draw:notify-on-update-of-ranges="${C.notifyOnUpdateOfRanges}" xlink:href="${C.href}" xlink:type="simple" xlink:show="embed" ` +
    `xlink:actuate="onLoad"><loext:p/></draw:object>` +
    `<draw:image xlink:href="${C.replacementHref}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
  const validated = (value: number): string =>
    `<table:table-cell table:content-validation-name="${O.contentValidation.name}" office:value-type="float" office:value="${value}" ` +
    `calcext:value-type="float"><text:p>${value}</text:p></table:table-cell>`;
  const A = O.annotation;
  const annotated =
    `<table:table-cell office:value-type="float" office:value="${V.Data.B2}" calcext:value-type="float">` +
    `<office:annotation draw:style-name="gr2" draw:text-style-name="P2" svg:width="1.6in" svg:height="0.5in" svg:x="2.1in" svg:y="0.05in" ` +
    `draw:caption-point-x="-0.1in" draw:caption-point-y="0.12in"><dc:creator>${A.creator}</dc:creator><dc:date>${A.date}</dc:date>` +
    `<text:p text:style-name="P2">${A.text}</text:p></office:annotation><text:p>${V.Data.B2}</text:p></table:table-cell>`;
  const link =
    `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>` +
    `<text:a xlink:href="${O.link.href}" xlink:type="simple">${O.link.text}</text:a></text:p></table:table-cell>`;
  const L = O.localNamedRange;
  const F = O.conditionalFormat;
  const data = table(
    1,
    `<office:forms form:automatic-focus="false" form:apply-design-mode="false"/><table:shapes>${frame}</table:shapes>` +
      columns +
      loRow(cell(V.Data.A1) + cell(V.Data.B1) + loEmpty(1) + link + loEmpty(16380)) +
      loRow(validated(V.Data.A2) + annotated + loEmpty(16382)) +
      ([3, 4, 5] as const).map((r) => loRow(validated(V.Data[`A${r}`]) + cell(V.Data[`B${r}`]) + loEmpty(16382))).join("") +
      tail(5) +
      `<table:named-expressions><table:named-range table:name="${L.name}" table:base-cell-address="${L.baseCellAddress}" ` +
      `table:cell-range-address="${L.cellRangeAddress}"/></table:named-expressions>` +
      `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${F.targetRangeAddress}">` +
      `<calcext:condition calcext:apply-style-name="${F.applyStyleName}" calcext:value="${xmlEscape(F.value)}" calcext:base-cell-address="${F.baseCellAddress}"/>` +
      `</calcext:conditional-format></calcext:conditional-formats>`,
  );

  const lookup = table(
    2,
    columns + ([1, 2, 3, 4, 5] as const).map((r) => loRow(cell(V.Lookup[`A${r}`]) + cell(V.Lookup[`B${r}`]) + loEmpty(16382))).join("") + tail(5),
  );
  const archive = table(3, columns + loRow(cell(V.Archive.A1) + cell(V.Archive.B1) + loEmpty(16382)) + tail(1));

  const V1 = O.contentValidation;
  const validations =
    `<table:content-validations><table:content-validation table:name="${V1.name}" table:condition="${xmlEscape(V1.condition)}" ` +
    `table:allow-empty-cell="true" table:display-list="unsorted" table:base-cell-address="${V1.baseCellAddress}">` +
    `<table:error-message table:message-type="stop" table:display="true"/></table:content-validation></table:content-validations>`;
  const N = O.namedRange;
  const E = O.namedExpression;
  const suffix =
    `<table:named-expressions><table:named-range table:name="${N.name}" table:base-cell-address="${N.baseCellAddress}" ` +
    `table:cell-range-address="${N.cellRangeAddress}"/><table:named-expression table:name="${E.name}" table:base-cell-address="${E.baseCellAddress}" ` +
    `table:expression="${xmlEscape(E.expression)}"/></table:named-expressions>` +
    `<table:database-ranges><table:database-range table:name="${O.databaseRange.name}" ` +
    `table:target-range-address="${O.databaseRange.targetRangeAddress}"/></table:database-ranges>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC121_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${validations}${summary}${data}${lookup}${archive}${suffix}` +
    `</office:spreadsheet></office:body></office:document-content>`;

  const cellCount = Object.values(V).reduce((sum, cells) => sum + Object.keys(cells).length, 0);
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    part("meta.xml", loMeta(O.title, O.tableNames.length, cellCount, 1)),
    part("settings.xml", spec121OdsSettings()),
    part("content.xml", content),
    { name: "Thumbnails/thumbnail.png", data: THUMBNAIL_PNG },
    part("Object 1/content.xml", spec121OdsChartContent()),
    part("Object 1/styles.xml", SPEC121_ODS_CHART_STYLES),
    part("Object 1/meta.xml", SPEC121_ODS_CHART_META),
    { name: C.replacement, data: SPEC121_ODS_REPLACEMENT },
    part("styles.xml", SPEC121_ODS_STYLES),
    part("manifest.rdf", LO_MANIFEST_RDF),
    part("META-INF/manifest.xml", SPEC121_ODS_MANIFEST),
  ]);
}

// ============================================================================
// SPEC-122 — spec122-shapes.xlsx / spec122-shapes.ods: what a rebuilt sheet keeps
// ============================================================================
// The workbooks the cell editor and the sheet writers are proven on: every
// construct SPEC-122's "Fixtures" lists, in the shape its producer writes it.
// - spec122-shapes.xlsx — Excel 365-shaped (Aptos, the 2023 theme). "Edit",
//   the active tab: a budget whose B2 is a plain General number the e2e
//   retypes (C2 reads it); C2:C5 and D2:D5 shared-formula blocks (master
//   first), E2:E5 a legacy array, G2:G5 a dynamic array (`cm="1"` and
//   xl/metadata.xml's XLDAPR), J2:J4 a one-variable data table over F2, rich,
//   phonetic (`ph="1"` on its cell) and `_x000D_` shared strings, B16 a
//   picture placed in its cell (`t="e"` `#VALUE!` with `vm="1"`, the rich
//   value in xl/richData), rows with the block's `spans`, `x14ac:dyDescent`
//   and `outlineLevel`, grouped custom-width columns, A8:E8 merged, row 1
//   frozen, dates on 14 and 22, custom formats 164 and 165, a quote-prefixed
//   "0042" whose number-stored-as-text error is ignored, a classic and an x14
//   data-bar conditional format, external and internal hyperlinks, and a page
//   setup with its printer settings. Every cached value is what the formula
//   computes, written as Excel writes numbers (17 digits when 15 do not
//   round-trip); no formula is volatile and none outside "Edit" reads it.
//   "Locked": `sheetProtection` with Excel 2013's SHA-512 hash of "fixture",
//   A1 locked, B2 the one unlocked cell. "Inline": inline strings as openpyxl
//   writes them, in the same package. calcChain.xml, TitlesOfParts.
// - spec122-shapes.ods — LibreOffice 24.8-shaped. "Data" (active): 23 rows
//   over A:D, so its tails are the 1048553-row and 16380-column runs; header
//   rows and a row group, a repeated non-empty row and a repeated non-empty
//   cell, a 2×2 span over covered cells, an annotation, an external
//   `text:a` link, a validated cell (its `table:content-validations` before
//   the tables), a cell in a `de-ce1` style, a red-negative N2/N2P0 pair and a
//   USD currency pair through `style:map`, every value type and `of:=`
//   formulas the engine reproduces, then the table-local named range and
//   conditional format after the rows. "Locked": `table:protected` with
//   Calc's SHA-256 key and `loext:table-protection`, B2 in a
//   `style:cell-protect="none"` style.

/** A `<v>` number as Excel writes it: 15 significant digits when they
 * round-trip, else 17 (its "%.17g" — 2000/3 is 666.66666666666663). */
function excelNumberText(value: number): string {
  const short = Number(value.toPrecision(15));
  return short === value ? String(short) : value.toPrecision(17).replace(/(\.\d*?)0+$/, "$1");
}

/** spec122-shapes.xlsx's cellXfs indices — a cell's `s`. */
export const SPEC122_XF = {
  normal: 0,
  header: 1, // bold on accent1 lighter 80%, a thin accent1 bottom edge
  date: 2, // built-in 14
  dateTime: 3, // built-in 22
  redNegative: 4, // custom 164
  percent: 5, // custom 165
  total: 6, // bold between a thin and a double accent1 edge
  hyperlink: 7, // the Hyperlink cell style
  quotePrefix: 8,
  unlocked: 9, // applyProtection="1" + <protection locked="0"/>
  note: 10, // italic Text 1 lighter 35%, wrapped
} as const;

/** spec122-shapes.xlsx's custom number formats. */
export const SPEC122_NUMFMTS = { 164: "#,##0.00;[Red]\\-#,##0.00", 165: "0.0%" } as const;

/** The shared strings' special items: a rich string whose first run has no
 * `rPr` (it keeps the cell's font) and whose last run is tinted, a phonetic
 * reading, and Excel's CR escape. */
const SPEC122_RICH_ITEM =
  `<si><r><t xml:space="preserve">Figures are </t></r>` +
  `<r><rPr><b/><sz val="11"/><color rgb="FFC00000"/><rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr><t>estimates</t></r>` +
  `<r><rPr><sz val="11"/><color theme="1" tint="0.34998626667073579"/><rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr>` +
  `<t xml:space="preserve"> until the audit</t></r></si>`;
const SPEC122_PHONETIC_ITEM = `<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>`;
const SPEC122_CR_ITEM = `<si><t xml:space="preserve">Line one_x000D_\nLine two</t></si>`;

/** spec122-shapes.xlsx's shared strings in index order (first use, row-major,
 * sheet order): each item's value as the reader gives it, and its XML when
 * it is more than a lone `<t>`. */
const SPEC122_SST: readonly { value: string; xml?: string }[] = [
  ...["Item", "Amount", "Double", "Running", "Share", "Rate", "Week", "Growth", "Rent", "Food", "Travel", "Books", "Total", "Per month"].map(
    (value) => ({ value }),
  ),
  { value: "Figures are estimates until the audit", xml: SPEC122_RICH_ITEM },
  { value: "Line one\r\nLine two", xml: SPEC122_CR_ITEM },
  { value: "Updated" },
  { value: "Logged" },
  { value: "Variance" },
  { value: "東京", xml: SPEC122_PHONETIC_ITEM },
  ...["Tokyo office", "Code", "0042", "Statements", "Locked!A1", "Receipt", "Approved total", "Adjustment", "Final"].map((value) => ({ value })),
];

/** "Edit"'s data table, its `<f>` verbatim (column input cell F2). */
const SPEC122_DATA_TABLE = `<f t="dataTable" ref="J2:J4" dt2D="0" dtr="0" r1="F2"/>`;
/** The x14 data bar's id, shared by its classic rule and its extension. */
const SPEC122_DATA_BAR_ID = "{1F5A3C7E-9B2D-4E6F-8A1C-3D5E7F9A1B2C}";
const SPEC122_SHEET_PROTECTION =
  `<sheetProtection algorithmName="SHA-512" hashValue="v29aHT5sWJ4HIjjGD0KTaqmRB9g25knlAqPdHUn15k+DY135r8qGmV09ra/OyqkT8v+qDWbrdwrKh+3d9+qBHg==" ` +
  `saltValue="XBqefTsvSmCBx+nQ8rSmyA==" spinCount="100000" sheet="1" objects="1" scenarios="1"/>`;
const SPEC122_PRINTER_SETTINGS_REL = `${REL_TYPE}/printerSettings`;
const SPEC122_RICH_VALUE_REL = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
const SPEC122_RD_REL = "http://schemas.microsoft.com/office/2017/06/relationships";

/** What spec122-shapes.xlsx holds, for the unit, round-trip and e2e specs.
 * Cells are A1 notation; a formula cell's value is its cached result. */
export const SPEC122_SHAPES_XLSX = {
  title: "Shapes",
  sheets: [
    { name: "Edit", sheetId: 1, partName: "xl/worksheets/sheet1.xml", uid: "{5D2B8E41-7C3A-4F19-9E6D-2A4C6E8F0B11}" },
    { name: "Locked", sheetId: 2, partName: "xl/worksheets/sheet2.xml", uid: "{6E3C9F52-8D4B-4A2A-8F7E-3B5D7F9A1C22}" },
    { name: "Inline", sheetId: 3, partName: "xl/worksheets/sheet3.xml" },
  ],
  sheetNames: ["Edit", "Locked", "Inline"],
  /** Edit — the only tabSelected sheetView; workbookView has no activeTab (0, Excel's default). */
  activeSheetIndex: 0,
  xf: SPEC122_XF,
  numFmts: SPEC122_NUMFMTS,
  edit: {
    /** The e2e's cell: a plain number in General — not merged, in no array
     * or shared block — whose dependents (C2, D2:D5, B6, …) are all on Edit. */
    target: { ref: "B2", value: 1200, dependent: { ref: "C2", formula: "B2*2" } },
    activeCell: "B2",
    dimension: "A1:J16",
    /** Every stored cell → the reader's value. */
    values: {
      A1: "Item", B1: "Amount", C1: "Double", D1: "Running", E1: "Share", F1: "Rate", G1: "Week", I1: "Growth", J1: 2000,
      A2: "Rent", B2: 1200, C2: 2400, D2: 1200, E2: 1200 / 2000, F2: 0, G2: 1, I2: 0, J2: 2000,
      A3: "Food", B3: 450.5, C3: 901, D3: 1650.5, E3: 450.5 / 2000, G3: 2, I3: 0.5, J3: 3000,
      A4: "Travel", B4: 300, C4: 600, D4: 1950.5, E4: 300 / 2000, G4: 3, I4: 1, J4: 4000,
      A5: "Books", B5: 49.5, C5: 99, D5: 2000, E5: 49.5 / 2000, G5: 4,
      A6: "Total", B6: 2000, C6: 4000,
      A7: "Per month", B7: 2000 / 3,
      A8: "Figures are estimates until the audit",
      A9: "Line one\r\nLine two",
      A10: "Updated", B10: dateToSerial(2026, 10, 8, "1900"),
      A11: "Logged", B11: dateToSerial(2026, 10, 8, "1900") + (13 * 60 + 45) / (24 * 60),
      A12: "Variance", B12: -125.5,
      A13: "東京", B13: "Tokyo office",
      A14: "Code", B14: "0042",
      A15: "Statements", B15: "Locked!A1",
      A16: "Receipt", B16: { error: "#VALUE!" },
    },
    /** Every formula cell → its canonical formula (followers expanded). */
    formulas: {
      J1: "B6*(1+F2)",
      C2: "B2*2", C3: "B3*2", C4: "B4*2", C5: "B5*2",
      D2: "SUM($B$2:B2)", D3: "SUM($B$2:B3)", D4: "SUM($B$2:B4)", D5: "SUM($B$2:B5)",
      E2: "B2:B5/B6",
      G2: "SEQUENCE(4)",
      B6: "SUM(B2:B5)", C6: "SUM(C2:C5)", B7: "B6/3",
    },
    /** The shared blocks' masters, as written (followers carry `si` only). */
    sharedBlocks: [
      { master: "C2", ref: "C2:C5", si: 0, wire: "B2*2" },
      { master: "D2", ref: "D2:D5", si: 1, wire: "SUM($B$2:B2)" },
    ],
    array: { anchor: "E2", ref: "E2:E5" },
    dynamicArray: { anchor: "G2", ref: "G2:G5", wire: "_xlfn.SEQUENCE(4)", cm: 1 },
    dataTable: { anchor: "J2", ref: "J2:J4", xml: SPEC122_DATA_TABLE, inputCell: "F2", inputs: { I2: 0, I3: 0.5, I4: 1 } },
    merge: "A8:E8",
    frozen: { rows: 1, cols: 0 },
    topLeftCell: "A2",
    /** A8's runs; A13's phonetic item (its reading is skipped) with ph="1";
     * A9 Excel's `_x000D_` escape of CR. */
    rich: { ref: "A8", runs: ["Figures are ", "estimates", " until the audit"], sstIndex: 14 },
    phonetic: { ref: "A13", value: "東京", sstIndex: 19, preservedAttrs: ` ph="1"` },
    carriageReturn: { ref: "A9", sstIndex: 15 },
    /** B16: a picture placed in the cell — `#VALUE!` for older readers. */
    valueMetadata: { ref: "B16", preservedAttrs: ` vm="1"` },
    /** Rows 2–5 (indices 1–4) are outline level 1; row 9 is 29 pt (two wrapped lines). */
    spans: "1:10",
    outlineRows: [1, 2, 3, 4],
    tallRow: { index: 8, heightPt: 29 },
    /** Columns A (custom width) and C:D (custom width, outline level 1). */
    columns: [
      { min: 0, max: 0, sourceWidthChars: 16.6328125 },
      { min: 2, max: 3, sourceWidthChars: 10.6328125, sourceAttrs: `outlineLevel="1"` },
    ],
    numberFormats: {
      B10: { builtinFormatId: 14 },
      B11: { builtinFormatId: 22 },
      B12: { numFmtId: 164, code: SPEC122_NUMFMTS[164] },
      E2: { numFmtId: 165, code: SPEC122_NUMFMTS[165] },
    },
    quotePrefix: { ref: "B14", value: "0042", xf: SPEC122_XF.quotePrefix },
    ignoredErrors: `<ignoredErrors><ignoredError sqref="B14" numberStoredAsText="1"/></ignoredErrors>`,
    hyperlinks: {
      external: { ref: "A15", rId: "rId1", target: "https://example.com/statements", uid: "{8A4C6E2F-0B1D-4C3E-9F5A-7B9D1F3A5C44}" },
      internal: { ref: "B15", location: "Locked!A1", uid: "{9B5D7F3A-1C2E-4D4F-8A6B-8C0E2A4B6D55}" },
    },
    printerSettings: { rId: "rId2", partName: "xl/printerSettings/printerSettings1.bin" },
    conditionalFormats: {
      dataBar: { sqref: "B2:B5", id: SPEC122_DATA_BAR_ID },
      cellIs: { sqref: "C2:C5", dxfId: 0, formula: "1000" },
    },
    /** The worksheet's children, CT_Worksheet order. */
    children: [
      "dimension",
      "sheetViews",
      "sheetFormatPr",
      "cols",
      "sheetData",
      "mergeCells",
      "conditionalFormatting",
      "conditionalFormatting",
      "hyperlinks",
      "pageMargins",
      "pageSetup",
      "ignoredErrors",
      "extLst",
    ],
  },
  locked: {
    values: { A1: "Approved total", B1: 2000, A2: "Adjustment", B2: 150, A3: "Final", B3: 2150 },
    formulas: { B3: "B1+B2" },
    /** A1: a locked cell (xf 0) — and the sheet's active cell, so typing there is refused. */
    protectedCell: { ref: "A1", value: "Approved total" },
    activeCell: "A1",
    /** B2: the sheet's one unlocked cell. */
    unlockedCell: { ref: "B2", value: 150, xf: SPEC122_XF.unlocked },
    password: "fixture",
    sheetProtection: SPEC122_SHEET_PROTECTION,
    children: ["dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "pageMargins"],
  },
  inline: {
    values: { A1: "Source", B1: "Rows", A2: "ledger.csv", B2: 128, A3: "Checked by script", B3: true, A4: "  padded  " },
    rich: { ref: "A3", runs: [{ text: "Checked", font: { bold: true } }, { text: " by script" }] },
  },
  /** xl/calcChain.xml's cells (`i` = the sheet's sheetId, written on each sheet's first). */
  calcChain: [
    ...["C2", "C3", "C4", "C5", "D2", "D3", "D4", "D5", "B6", "C6", "B7", "E2", "J1", "G2", "J2"].map((r) => ({ r, i: 1 })),
    { r: "B3", i: 2 },
  ],
  /** xl/metadata.xml as Excel relates and types it. */
  metadata: {
    partName: "xl/metadata.xml",
    rId: "rId8",
    type: `${REL_TYPE}/sheetMetadata`,
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
  },
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/printerSettings/printerSettings1.bin",
    "xl/metadata.xml",
    "xl/richData/rdrichvalue.xml",
    "xl/richData/rdrichvaluestructure.xml",
    "xl/richData/rdRichValueTypes.xml",
    "xl/richData/richValueRel.xml",
    "xl/richData/_rels/richValueRel.xml.rels",
    "xl/media/image1.png",
    "xl/calcChain.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

const SPEC122_FONT_TAIL = `<name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>`;

/** spec122-shapes.xlsx's styles part (see SPEC122_XF): font 1 is the 6 pt
 * phonetic font the phonetic item names, fill 2 and font 4 theme colours
 * with Excel's tints, dxf 0 the classic conditional format's. */
const SPEC122_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<numFmts count="2"><numFmt numFmtId="164" formatCode="${excelAttr(SPEC122_NUMFMTS[164])}"/><numFmt numFmtId="165" formatCode="${SPEC122_NUMFMTS[165]}"/></numFmts>` +
  `<fonts count="5" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/>${SPEC122_FONT_TAIL}` +
  `<font><sz val="6"/><name val="Aptos Narrow"/><family val="3"/><charset val="128"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/>${SPEC122_FONT_TAIL}` +
  `<font><u/><sz val="11"/><color theme="10"/>${SPEC122_FONT_TAIL}` +
  `<font><i/><sz val="11"/><color theme="1" tint="0.34998626667073579"/>${SPEC122_FONT_TAIL}</fonts>` +
  `<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>` +
  `<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>` +
  `<border><left/><right/><top/><bottom style="thin"><color theme="4"/></bottom><diagonal/></border>` +
  `<border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/></bottom><diagonal/></border></borders>` +
  `<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
  `<cellXfs count="11"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>` +
  `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyProtection="1"><protection locked="0"/></xf>` +
  `<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>` +
  `<cellStyles count="2"><cellStyle name="Hyperlink" xfId="1" builtinId="8"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
  `<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>` +
  `<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
  `<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;

/** xl/metadata.xml: the dynamic array's XLDAPR (cm="1") and the placed
 * picture's XLRICHVALUE (vm="1" → rich value 0). */
const SPEC122_METADATA =
  XML_DECLARATION +
  `<metadata xmlns="${MAIN_NS}" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" ` +
  `xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="2">` +
  `<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" ` +
  `clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>` +
  `<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" ` +
  `clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>` +
  `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">` +
  `<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
  `<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>` +
  `<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata><valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata></metadata>`;

const SPEC122_RICH_DATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

/** The placed picture's rich value: a `_localImage` naming relationship 0
 * of richValueRel.xml (CalcOrigin 5: placed in the cell). */
const SPEC122_RICH_VALUE = XML_DECLARATION + `<rvData xmlns="${SPEC122_RICH_DATA_NS}" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`;
const SPEC122_RICH_VALUE_STRUCTURE =
  XML_DECLARATION +
  `<rvStructures xmlns="${SPEC122_RICH_DATA_NS}" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>`;
const SPEC122_RICH_VALUE_TYPES =
  XML_DECLARATION +
  `<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" ` +
  `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${MAIN_NS}"><global><keyFlags>` +
  `<key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
  ["_DisplayString", "_Flags", "_Format", "_SubLabel", "_Attribution", "_Icon", "_Display", "_CanonicalPropertyNames", "_ClassificationId"]
    .map((key) => `<key name="${key}"><flag name="ExcludeFromCalcComparison" value="1"/></key>`)
    .join("") +
  `</keyFlags></global></rvTypesInfo>`;
const SPEC122_RICH_VALUE_RELS =
  XML_DECLARATION +
  `<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="${REL_NS}"><rel r:id="rId1"/></richValueRels>`;

/** xl/printerSettings/printerSettings1.bin: a stand-in for the DEVMODE
 * Windows Excel caches with a page setup — the device name in UTF-16, then
 * a few fields (bytes this product keeps, never reads). */
function spec122PrinterSettings(): Uint8Array {
  const device = "Microsoft Print to PDF";
  const bytes = new Uint8Array(72);
  for (let i = 0; i < device.length; i += 1) {
    bytes[i * 2] = device.charCodeAt(i);
  }
  bytes.set([0x01, 0x04, 0x00, 0x05, 0xdc, 0x00, 0x24, 0x03], 64);
  return bytes;
}

/** A shared-string cell for the SPEC122_SST item whose value is `value`;
 * `extra` adds attributes after `t` (Excel's order: r, s, t, cm, vm, ph). */
function spec122Text(ref: string, value: string, style?: number, extra = ""): string {
  const index = SPEC122_SST.findIndex((item) => item.value === value);
  if (index < 0) {
    throw new Error(`"${value}" is not one of spec122-shapes.xlsx's strings`);
  }
  return `<c r="${ref}"${style ? ` s="${style}"` : ""} t="s"${extra}><v>${index}</v></c>`;
}

const spec122Number = (ref: string, value: number, style?: number): string =>
  `<c r="${ref}"${style ? ` s="${style}"` : ""}><v>${excelNumberText(value)}</v></c>`;

/** A formula cell with a numeric result: `<f>` (any shape) and the cached `<v>`. */
const spec122Formula = (ref: string, f: string, value: number, style?: number, cm?: number): string =>
  `<c r="${ref}"${style ? ` s="${style}"` : ""}${cm === undefined ? "" : ` cm="${cm}"`}>${f}<v>${excelNumberText(value)}</v></c>`;

/** One of the Excel-shaped rows: `r`, the 16-row block's `spans` (every
 * column A:J holds a cell in rows 1–16), the row's own attributes in
 * Excel's order, `x14ac:dyDescent`. */
const spec122Row = (r: number, spans: string, attrs: string, cells: string): string =>
  `<row r="${r}" spans="${spans}"${attrs} x14ac:dyDescent="0.35">${cells}</row>`;

/** "Edit" — see the section header and SPEC122_SHAPES_XLSX.edit. */
function spec122EditSheet(): string {
  const E = SPEC122_SHAPES_XLSX.edit;
  const V = E.values;
  const X = SPEC122_XF;
  const [C, D] = E.sharedBlocks;
  const row = (r: number, cells: string, attrs = ""): string => spec122Row(r, E.spans, attrs, cells);
  const outlined = ` outlineLevel="1"`;
  const shared = (ref: string, si: number, value: number): string => spec122Formula(ref, `<f t="shared" si="${si}"/>`, value);
  const rows = [
    row(
      1,
      (["A1", "B1", "C1", "D1", "E1", "F1", "G1", "I1"] as const).map((ref) => spec122Text(ref, V[ref], X.header)).join("") +
        spec122Formula("J1", normalF(E.formulas.J1), V.J1),
    ),
    row(
      2,
      spec122Text("A2", V.A2) +
        spec122Number("B2", V.B2) +
        spec122Formula("C2", `<f t="shared" ref="${C.ref}" si="${C.si}">${xmlText(C.wire)}</f>`, V.C2) +
        spec122Formula("D2", `<f t="shared" ref="${D.ref}" si="${D.si}">${xmlText(D.wire)}</f>`, V.D2) +
        spec122Formula("E2", `<f t="array" ref="${E.array.ref}">${xmlText(E.formulas.E2)}</f>`, V.E2, X.percent) +
        spec122Number("F2", V.F2) +
        spec122Formula("G2", `<f t="array" ref="${E.dynamicArray.ref}">${E.dynamicArray.wire}</f>`, V.G2, undefined, E.dynamicArray.cm) +
        spec122Number("I2", V.I2, X.percent) +
        spec122Formula("J2", SPEC122_DATA_TABLE, V.J2),
      outlined,
    ),
    row(
      3,
      spec122Text("A3", V.A3) + spec122Number("B3", V.B3) + shared("C3", C.si, V.C3) + shared("D3", D.si, V.D3) + spec122Number("E3", V.E3, X.percent) +
        spec122Number("G3", V.G3) + spec122Number("I3", V.I3, X.percent) + spec122Number("J3", V.J3),
      outlined,
    ),
    row(
      4,
      spec122Text("A4", V.A4) + spec122Number("B4", V.B4) + shared("C4", C.si, V.C4) + shared("D4", D.si, V.D4) + spec122Number("E4", V.E4, X.percent) +
        spec122Number("G4", V.G4) + spec122Number("I4", V.I4, X.percent) + spec122Number("J4", V.J4),
      outlined,
    ),
    row(
      5,
      spec122Text("A5", V.A5) + spec122Number("B5", V.B5) + shared("C5", C.si, V.C5) + shared("D5", D.si, V.D5) + spec122Number("E5", V.E5, X.percent) +
        spec122Number("G5", V.G5),
      outlined,
    ),
    row(
      6,
      spec122Text("A6", V.A6, X.total) + spec122Formula("B6", normalF(E.formulas.B6), V.B6, X.total) + spec122Formula("C6", normalF(E.formulas.C6), V.C6, X.total),
    ),
    row(7, spec122Text("A7", V.A7) + spec122Formula("B7", normalF(E.formulas.B7), V.B7)),
    row(8, spec122Text("A8", V.A8)),
    row(9, spec122Text("A9", V.A9, X.note), ` ht="${E.tallRow.heightPt}"`),
    row(10, spec122Text("A10", V.A10) + spec122Number("B10", V.B10, X.date)),
    row(11, spec122Text("A11", V.A11) + spec122Number("B11", V.B11, X.dateTime)),
    row(12, spec122Text("A12", V.A12) + spec122Number("B12", V.B12, X.redNegative)),
    row(13, spec122Text("A13", V.A13, undefined, E.phonetic.preservedAttrs) + spec122Text("B13", V.B13)),
    row(14, spec122Text("A14", V.A14) + spec122Text("B14", V.B14, X.quotePrefix)),
    row(15, spec122Text("A15", V.A15, X.hyperlink) + spec122Text("B15", V.B15, X.hyperlink)),
    row(16, spec122Text("A16", V.A16) + `<c r="B16" t="e"${E.valueMetadata.preservedAttrs}><v>${V.B16.error}</v></c>`),
  ].join("");
  const H = E.hyperlinks;
  const CF = E.conditionalFormats;
  const x14 = `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"`;
  const [firstColumn, groupedColumns] = E.columns;
  return (
    XML_DECLARATION +
    worksheetOpen(SPEC122_SHAPES_XLSX.sheets[0].uid.slice(1, -1)) +
    `<dimension ref="${E.dimension}"/><sheetViews><sheetView tabSelected="1" workbookViewId="0">` +
    `<pane ySplit="${E.frozen.rows}" topLeftCell="${E.topLeftCell}" activePane="bottomLeft" state="frozen"/>` +
    `<selection pane="bottomLeft" activeCell="${E.activeCell}" sqref="${E.activeCell}"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" outlineLevelRow="1" outlineLevelCol="1" x14ac:dyDescent="0.35"/>` +
    `<cols><col min="${firstColumn.min + 1}" max="${firstColumn.max + 1}" width="${firstColumn.sourceWidthChars}" customWidth="1"/>` +
    `<col min="${groupedColumns.min + 1}" max="${groupedColumns.max + 1}" width="${groupedColumns.sourceWidthChars}" customWidth="1" ${groupedColumns.sourceAttrs}/></cols>` +
    `<sheetData>${rows}</sheetData><mergeCells count="1"><mergeCell ref="${E.merge}"/></mergeCells>` +
    `<conditionalFormatting sqref="${CF.dataBar.sqref}"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/>` +
    `<color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${x14}><x14:id>${CF.dataBar.id}</x14:id></ext></extLst>` +
    `</cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="${CF.cellIs.sqref}"><cfRule type="cellIs" dxfId="${CF.cellIs.dxfId}" priority="1" operator="greaterThan">` +
    `<formula>${CF.cellIs.formula}</formula></cfRule></conditionalFormatting>` +
    `<hyperlinks><hyperlink ref="${H.external.ref}" r:id="${H.external.rId}" xr:uid="${H.external.uid}"/>` +
    `<hyperlink ref="${H.internal.ref}" location="${H.internal.location}" display="${H.internal.location}" xr:uid="${H.internal.uid}"/></hyperlinks>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `<pageSetup paperSize="9" orientation="landscape" r:id="${E.printerSettings.rId}"/>${E.ignoredErrors}` +
    `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${x14}><x14:conditionalFormattings>` +
    `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${CF.dataBar.id}">` +
    `<x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>` +
    `<x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/>` +
    `</x14:dataBar></x14:cfRule><xm:sqref>${CF.dataBar.sqref}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>`
  );
}

/** "Locked": the protected sheet (sheetProtection after sheetData, in
 * CT_Worksheet order), its active cell the locked A1. */
function spec122LockedSheet(): string {
  const L = SPEC122_SHAPES_XLSX.locked;
  const V = L.values;
  const row = (r: number, cells: string): string => spec122Row(r, "1:2", "", cells);
  return (
    XML_DECLARATION +
    worksheetOpen(SPEC122_SHAPES_XLSX.sheets[1].uid.slice(1, -1)) +
    `<dimension ref="A1:B3"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="${L.activeCell}" sqref="${L.activeCell}"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/><cols><col min="1" max="1" width="16.6328125" customWidth="1"/></cols><sheetData>` +
    row(1, spec122Text("A1", V.A1) + spec122Number("B1", V.B1)) +
    row(2, spec122Text("A2", V.A2) + spec122Number("B2", V.B2, L.unlockedCell.xf)) +
    row(3, spec122Text("A3", V.A3) + spec122Formula("B3", normalF(L.formulas.B3), V.B3, SPEC122_XF.total)) +
    `</sheetData>${L.sheetProtection}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  );
}

/** "Inline": openpyxl's shape — inline strings (one of them rich), `t="n"`
 * numbers, rows without spans, no `xr:uid`, its own sizes and margins. */
function spec122InlineSheet(): string {
  const V = SPEC122_SHAPES_XLSX.inline.values;
  const inline = openpyxlInline;
  return (
    OPENPYXL_DECLARATION +
    `<worksheet xmlns="${MAIN_NS}"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr/></sheetPr><dimension ref="A1:B4"/>` +
    `<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    `<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/><sheetData>` +
    `<row r="1">${inline("A1", V.A1)}${inline("B1", V.B1)}</row>` +
    `<row r="2">${inline("A2", V.A2)}<c r="B2" t="n"><v>${V.B2}</v></c></row>` +
    `<row r="3"><c r="A3" t="inlineStr"><is><r><rPr><b val="1"/></rPr><t>Checked</t></r><r><t xml:space="preserve"> by script</t></r></is></c>` +
    `<c r="B3" t="b"><v>1</v></c></row>` +
    `<row r="4"><c r="A4" t="inlineStr"><is><t xml:space="preserve">${V.A4}</t></is></c></row>` +
    `</sheetData><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/></worksheet>`
  );
}

export async function buildSpec122ShapesXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC122_SHAPES_XLSX;
  const worksheets = [spec122EditSheet(), spec122LockedSheet(), spec122InlineSheet()];
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${SPEC122_SST.length}">` +
    SPEC122_SST.map((item) => item.xml ?? `<si><t>${item.value}</t></si>`).join("") +
    `</sst>`;
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}">` +
    F.calcChain
      .map((entry, index) => {
        const first = index === 0 || F.calcChain[index - 1]!.i !== entry.i;
        const array = entry.r === F.edit.array.anchor || entry.r === F.edit.dynamicArray.anchor;
        return `<c r="${entry.r}"${first ? ` i="${entry.i}"` : ""}${index === 0 ? ` l="1"` : ""}${array ? ` a="1"` : ""}/>`;
      })
      .join("") +
    `</calcChain>`;
  const rels = (relationships: string): string => XML_DECLARATION + `<Relationships xmlns="${PKG_REL_NS}">${relationships}</Relationships>`;
  const H = F.edit.hyperlinks.external;
  const P = F.edit.printerSettings;
  const M = F.metadata;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheets.map((sheet, index) => ({ name: sheet.name, xml: worksheets[index]! })),
    sharedStrings,
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    styles: SPEC122_STYLES,
    defaults: [
      `<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>`,
      `<Default Extension="png" ContentType="image/png"/>`,
    ],
    additions: {
      parts: [
        part(
          "xl/worksheets/_rels/sheet1.xml.rels",
          rels(
            `<Relationship Id="${P.rId}" Type="${SPEC122_PRINTER_SETTINGS_REL}" Target="../printerSettings/printerSettings1.bin"/>` +
              `<Relationship Id="${H.rId}" Type="${REL_TYPE}/hyperlink" Target="${H.target}" TargetMode="External"/>`,
          ),
        ),
        { name: P.partName, data: spec122PrinterSettings() },
        part(M.partName, SPEC122_METADATA),
        part("xl/richData/rdrichvalue.xml", SPEC122_RICH_VALUE),
        part("xl/richData/rdrichvaluestructure.xml", SPEC122_RICH_VALUE_STRUCTURE),
        part("xl/richData/rdRichValueTypes.xml", SPEC122_RICH_VALUE_TYPES),
        part("xl/richData/richValueRel.xml", SPEC122_RICH_VALUE_RELS),
        part("xl/richData/_rels/richValueRel.xml.rels", rels(`<Relationship Id="rId1" Type="${REL_TYPE}/image" Target="../media/image1.png"/>`)),
        { name: "xl/media/image1.png", data: THUMBNAIL_PNG },
      ],
      relationships: [
        `<Relationship Id="${M.rId}" Type="${M.type}" Target="metadata.xml"/>`,
        `<Relationship Id="rId9" Type="${SPEC122_RICH_VALUE_REL}" Target="richData/richValueRel.xml"/>`,
        `<Relationship Id="rId10" Type="${SPEC122_RD_REL}/rdRichValue" Target="richData/rdrichvalue.xml"/>`,
        `<Relationship Id="rId11" Type="${SPEC122_RD_REL}/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>`,
        `<Relationship Id="rId12" Type="${SPEC122_RD_REL}/rdRichValueTypes" Target="richData/rdRichValueTypes.xml"/>`,
      ],
      overrides: [
        `<Override PartName="/${M.partName}" ContentType="${M.contentType}"/>`,
        `<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>`,
        `<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>`,
        `<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>`,
        `<Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>`,
      ],
    },
  });
}

// --- spec122-shapes.ods ---------------------------------------------------------

/** "Data"'s annotation (A14), verbatim — what the reader keeps as the
 * cell's preservedXml. */
const SPEC122_ODS_ANNOTATION =
  `<office:annotation draw:style-name="gr1" draw:text-style-name="P1" svg:width="1.6in" svg:height="0.5in" svg:x="1.6in" svg:y="2.3in" ` +
  `draw:caption-point-x="-0.1in" draw:caption-point-y="0.12in"><dc:creator>${FIXTURE_WORKBOOK_AUTHOR}</dc:creator><dc:date>2026-10-08T17:20:00</dc:date>` +
  `<text:p text:style-name="P1">Matches the delivery note.</text:p></office:annotation>`;

/** "Data"'s link paragraph (A15), verbatim — the cell's preservedText. */
const SPEC122_ODS_LINK = `<text:p><text:a xlink:href="https://example.com/catalogue" xlink:type="simple">Catalogue</text:a></text:p>`;

/** What spec122-shapes.ods holds, for the unit and round-trip specs. Cells
 * are A1 notation; a formula cell's value is its cached result. */
export const SPEC122_SHAPES_ODS = {
  title: "Shapes",
  tableNames: ["Data", "Locked"],
  activeTable: "Data",
  activeSheetIndex: 0,
  data: {
    /** Every stored cell → the reader's value (rows 4–6 are one repeated row,
     * B7:D7 one repeated cell; the covered B12, A13 and B13 store nothing). */
    values: {
      A1: "Item", B1: "Qty", C1: "Price", D1: "Total",
      A2: "Widget", B2: 4, C2: 2.5, D2: 10,
      A3: "Gadget", B3: 3, C3: 4.25, D3: 12.75,
      A4: "Spare", B4: 1, C4: 0.5, D4: 0.5,
      A5: "Spare", B5: 1, C5: 0.5, D5: 0.5,
      A6: "Spare", B6: 1, C6: 0.5, D6: 0.5,
      A7: "Zero", B7: 0, C7: 0, D7: 0,
      A8: "Discount", B8: 0.15,
      A9: "Delivered", B9: dateToSerial(2026, 10, 8, "1900"), C9: (13 * 60 + 45) / (24 * 60),
      A10: "In stock", B10: true, C10: false,
      A11: "Balance", B11: -42.5,
      A12: "Merged note", C12: 7,
      A14: "Reviewed",
      A15: "Catalogue",
      A16: "Quantity", B16: 5,
      A17: "Highlighted",
      A18: "Total qty", B18: 10,
      A19: "Total value", B19: 24.25,
      A20: "Big orders", B20: 2,
      A21: "Label", B21: "Widget x4",
      A22: "Long time", B22: (36 * 60 + 5) / (24 * 60),
      A23: "Done",
    },
    /** Formula cells: `table:formula`, the canonical formula, the cached value. */
    formulas: {
      D2: { wire: "of:=[.B2]*[.C2]", formula: "B2*C2", value: 10 },
      D3: { wire: "of:=[.B3]*[.C3]", formula: "B3*C3", value: 12.75 },
      B18: { wire: "of:=SUM([.B2:.B7])", formula: "SUM(B2:B7)", value: 10 },
      B19: { wire: "of:=SUM([.D2:.D7])", formula: "SUM(D2:D7)", value: 24.25 },
      B20: { wire: 'of:=COUNTIF([.B2:.B7];">1")', formula: 'COUNTIF(B2:B7,">1")', value: 2 },
      B21: { wire: 'of:=[.A2]&" x"&[.B2]', formula: 'A2&" x"&B2', value: "Widget x4" },
    },
    /** 23 rows written over A:D; LibreOffice's tails reach the grid's edge. */
    usedRows: 23,
    usedColumns: 4,
    trailingRowRun: 1_048_553,
    trailingColumnRun: 16_380,
    /** Row 1 in `table:table-header-rows`; rows 2–7 in a `table:table-row-group`. */
    headerRows: { first: 0, last: 0 },
    rowGroup: { first: 1, last: 6 },
    repeatedRow: { first: "A4", count: 3 },
    repeatedCell: { first: "B7", count: 3, value: 0 },
    /** A12 spans A12:B13 (2×2) over three covered cells. */
    merge: { anchor: "A12", range: "A12:B13", covered: ["B12", "A13", "B13"] },
    annotation: { a1: "A14", xml: SPEC122_ODS_ANNOTATION, creator: FIXTURE_WORKBOOK_AUTHOR, text: "Matches the delivery note." },
    link: { a1: "A15", href: "https://example.com/catalogue", text: "Catalogue", paragraph: SPEC122_ODS_LINK },
    /** B16's validation, declared in content.xml's prefix. */
    validation: {
      a1: "B16",
      name: "val1",
      condition: "of:cell-content-is-whole-number() and cell-content-is-between(1;10)",
      baseCellAddress: "Data.B16",
      preservedAttrs: ` table:content-validation-name="val1"`,
    },
    /** A17's automatic style is literally named de-ce1 (a yellow fill): the
     * table writer's generated names must step past it. */
    collision: { a1: "A17", styleName: "de-ce1", fill: "#ffff00" },
    /** B11 in ce6 → N2, the red negative (N2P0 through style:map). */
    redNegative: { a1: "B11", styleName: "ce6", dataStyle: "N2", code: "0.00;[Red]-0.00" },
    /** C2:D6 and B19 in ce3 → N122, USD with a red negative (N122P0 through style:map). */
    currency: { styleName: "ce3", dataStyle: "N122" },
    localNamedRange: { name: "Quantities", baseCellAddress: "$Data.$B$2", cellRangeAddress: "$Data.$B$2:.$B$7" },
    conditionalFormat: { targetRangeAddress: "Data.D2:Data.D7", value: ">10", baseCellAddress: "Data.D2", applyStyleName: "Bad" },
  },
  locked: {
    values: { A1: "Approved total", B1: 24.25, A2: "Adjustment", B2: 1.5, A3: "Final", B3: 25.75 },
    formulas: { B3: { wire: "of:=[.B1]+[.B2]", formula: "B1+B2", value: 25.75 } },
    protectedCell: { a1: "A1" },
    /** B2: the table's one cell whose style (ce7) says style:cell-protect="none". */
    unprotectedCell: { a1: "B2", styleName: "ce7", value: 1.5 },
    password: "fixture",
    /** Calc's key: base64 SHA-256 of the UTF-8 password. */
    protectionKey: "8W0F7GspJI0sYa2x6SY/eOT3us4blVAUotF4cs/kBk0=",
    digestAlgorithm: "http://www.w3.org/2000/09/xmldsig#sha256",
    tableProtection: `<loext:table-protection loext:select-protected-cells="true" loext:select-unprotected-cells="true"/>`,
  },
  /** Every entry, in package order (mimetype first, STOREd). */
  parts: ["mimetype", "meta.xml", "settings.xml", "content.xml", "Thumbnails/thumbnail.png", "styles.xml", "manifest.rdf", "META-INF/manifest.xml"],
} as const;

/** content.xml's automatic styles, in Calc's order: columns, rows, the table,
 * the data styles (N2/N2P0 and N122/N122P0 pairs through style:map), the
 * cell styles (ce7 unprotected), the de-ce1 a previous save named, and the
 * annotation's graphic and paragraph styles. */
const SPEC122_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="1.5in"/></style:style>` +
  `<style:style style:name="co3" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="1in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<number:number-style style:name="N2P0" style:volatile="true"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/></number:number-style>` +
  `<number:number-style style:name="N2"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N2P0"/></number:number-style>` +
  `<number:percentage-style style:name="N11"><number:number number:decimal-places="0" number:min-decimal-places="0" number:min-integer-digits="1"/>` +
  `<number:text>%</number:text></number:percentage-style>` +
  `<number:date-style style:name="N37" number:automatic-order="true"><number:month number:style="long"/><number:text>/</number:text>` +
  `<number:day number:style="long"/><number:text>/</number:text><number:year/></number:date-style>` +
  `<number:time-style style:name="N41" number:truncate-on-overflow="false"><number:hours number:style="long"/><number:text>:</number:text>` +
  `<number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:time-style>` +
  `<number:currency-style style:name="N122P0" style:volatile="true"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>` +
  `<number:currency-style style:name="N122"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>` +
  `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N122P0"/></number:currency-style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
  `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N11"/>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N122"/>` +
  `<style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>` +
  `<style:style style:name="ce5" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N41"/>` +
  `<style:style style:name="ce6" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2"/>` +
  `<style:style style:name="ce7" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:cell-protect="none" ` +
  `style:print-content="true"/></style:style>` +
  `<style:style style:name="de-ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ffff00"/></style:style>` +
  `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-color="#000000" draw:fill="solid" ` +
  `draw:fill-color="#ffffc0" draw:auto-grow-height="true" draw:auto-grow-width="false" fo:padding-top="0.0398in" fo:padding-bottom="0.0398in" ` +
  `fo:padding-left="0.0398in" fo:padding-right="0.0398in" draw:shadow="visible" draw:shadow-offset-x="0.0236in" draw:shadow-offset-y="0.0236in" ` +
  `draw:shadow-color="#000000"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0in" ` +
  `fo:margin-bottom="0in" fo:text-indent="0in" fo:border="none"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="9pt"/></style:style>` +
  `</office:automatic-styles>`;

/** A LibreOffice typed cell: its style (if any), value attributes and text. */
const spec122OdsCell = (style: string | undefined, attrs: string, text: string): string =>
  `<table:table-cell${style ? ` table:style-name="${style}"` : ""} ${attrs}><text:p>${text}</text:p></table:table-cell>`;

/** "Data" — see the section header and SPEC122_SHAPES_ODS.data. */
function spec122OdsDataTable(): string {
  const D = SPEC122_SHAPES_ODS.data;
  const V = D.values;
  const F = D.formulas;
  const cell = spec122OdsCell;
  /** A row of label A, one value cell B, and the empty run to XFD. */
  const labelled = (label: string, value: string): string => loRow(loString(label) + value + loEmpty(16382));
  const header = loRow((["A1", "B1", "C1", "D1"] as const).map((ref) => loString(V[ref], "ce1")).join("") + loEmpty(16380));
  const group =
    loRow(loString(V.A2) + loFloat(V.B2) + loCurrency(V.C2, "$2.50") + loCurrency(V.D2, "$10.00", F.D2.wire) + loEmpty(16380)) +
    loRow(loString(V.A3) + loFloat(V.B3) + loCurrency(V.C3, "$4.25") + loCurrency(V.D3, "$12.75", F.D3.wire) + loEmpty(16380)) +
    // A repeated NON-empty row: rows 4–6.
    loRow(loString(V.A4) + loFloat(V.B4) + loCurrency(V.C4, "$0.50") + loCurrency(V.D4, "$0.50") + loEmpty(16380), "ro1", D.repeatedRow.count) +
    // A repeated NON-empty cell: B7:D7.
    loRow(
      loString(V.A7) +
        `<table:table-cell table:number-columns-repeated="${D.repeatedCell.count}" office:value-type="float" office:value="0" calcext:value-type="float">` +
        `<text:p>0</text:p></table:table-cell>` +
        loEmpty(16380),
    );
  const A = D.annotation;
  const rows =
    labelled(V.A8, cell("ce2", `office:value-type="percentage" office:value="${V.B8}" calcext:value-type="percentage"`, "15%")) +
    loRow(
      loString(V.A9) +
        cell("ce4", `office:value-type="date" office:date-value="2026-10-08" calcext:value-type="date"`, "10/08/2026") +
        cell("ce5", `office:value-type="time" office:time-value="PT13H45M00S" calcext:value-type="time"`, "13:45:00") +
        loEmpty(16381),
    ) +
    loRow(
      loString(V.A10) +
        cell(undefined, `office:value-type="boolean" office:boolean-value="true" calcext:value-type="boolean"`, "TRUE") +
        cell(undefined, `office:value-type="boolean" office:boolean-value="false" calcext:value-type="boolean"`, "FALSE") +
        loEmpty(16381),
    ) +
    labelled(V.A11, cell(D.redNegative.styleName, `office:value-type="float" office:value="${V.B11}" calcext:value-type="float"`, "-42.50")) +
    // A 2×2 span: B12 covered here, A13:B13 on the next row.
    loRow(
      `<table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2" office:value-type="string" calcext:value-type="string">` +
        `<text:p>${V.A12}</text:p></table:table-cell><table:covered-table-cell/>` +
        loFloat(V.C12) +
        loEmpty(16381),
    ) +
    loRow(`<table:covered-table-cell table:number-columns-repeated="2"/>` + loEmpty(16382)) +
    loRow(`<table:table-cell office:value-type="string" calcext:value-type="string">${A.xml}<text:p>${V.A14}</text:p></table:table-cell>` + loEmpty(16383)) +
    loRow(`<table:table-cell office:value-type="string" calcext:value-type="string">${D.link.paragraph}</table:table-cell>` + loEmpty(16383)) +
    labelled(
      V.A16,
      `<table:table-cell${D.validation.preservedAttrs} office:value-type="float" office:value="${V.B16}" calcext:value-type="float"><text:p>${V.B16}</text:p></table:table-cell>`,
    ) +
    loRow(loString(V.A17, D.collision.styleName) + loEmpty(16383)) +
    labelled(V.A18, loFloat(V.B18, F.B18.wire)) +
    labelled(V.A19, loCurrency(V.B19, "$24.25", F.B19.wire)) +
    labelled(V.A20, loFormula(F.B20.wire, V.B20)) +
    labelled(
      V.A21,
      `<table:table-cell table:formula="${xmlEscape(F.B21.wire)}" office:value-type="string" office:string-value="${V.B21}" calcext:value-type="string">` +
        `<text:p>${V.B21}</text:p></table:table-cell>`,
    ) +
    labelled(V.A22, cell("ce5", `office:value-type="time" office:time-value="PT36H05M00S" calcext:value-type="time"`, "36:05:00")) +
    loRow(loString(V.A23) + loEmpty(16383)) +
    // LibreOffice's trailing run to the grid's edge.
    loRow(loEmpty(16384), "ro1", D.trailingRowRun);
  const N = D.localNamedRange;
  const C = D.conditionalFormat;
  return (
    `<table:table table:name="${SPEC122_SHAPES_ODS.tableNames[0]}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co3" table:number-columns-repeated="3" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="${D.trailingColumnRun}" table:default-cell-style-name="Default"/>` +
    `<table:table-header-rows>${header}</table:table-header-rows><table:table-row-group>${group}</table:table-row-group>${rows}` +
    `<table:named-expressions><table:named-range table:name="${N.name}" table:base-cell-address="${N.baseCellAddress}" ` +
    `table:cell-range-address="${N.cellRangeAddress}"/></table:named-expressions>` +
    `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${C.targetRangeAddress}">` +
    `<calcext:condition calcext:apply-style-name="${C.applyStyleName}" calcext:value="${xmlEscape(C.value)}" calcext:base-cell-address="${C.baseCellAddress}"/>` +
    `</calcext:conditional-format></calcext:conditional-formats></table:table>`
  );
}

/** "Locked": protected with a password, its one unprotected cell B2. */
function spec122OdsLockedTable(): string {
  const L = SPEC122_SHAPES_ODS.locked;
  const V = L.values;
  return (
    `<table:table table:name="${SPEC122_SHAPES_ODS.tableNames[1]}" table:style-name="ta1" table:protected="true" ` +
    `table:protection-key="${L.protectionKey}" table:protection-key-digest-algorithm="${L.digestAlgorithm}">${L.tableProtection}` +
    `<table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16383" table:default-cell-style-name="Default"/>` +
    loRow(loString(V.A1) + loCurrency(V.B1, "$24.25") + loEmpty(16382)) +
    loRow(
      loString(V.A2) +
        spec122OdsCell(L.unprotectedCell.styleName, `office:value-type="float" office:value="${V.B2}" calcext:value-type="float"`, String(V.B2)) +
        loEmpty(16382),
    ) +
    loRow(loString(V.A3) + loFormula(L.formulas.B3.wire, V.B3) + loEmpty(16382)) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 3) +
    `</table:table>`
  );
}

export async function buildSpec122ShapesOdsBytes(): Promise<Uint8Array> {
  const O = SPEC122_SHAPES_ODS;
  const V1 = O.data.validation;
  const validations =
    `<table:content-validations><table:content-validation table:name="${V1.name}" table:condition="${xmlEscape(V1.condition)}" ` +
    `table:allow-empty-cell="true" table:base-cell-address="${V1.baseCellAddress}">` +
    `<table:error-message table:message-type="stop" table:display="true"/></table:content-validation></table:content-validations>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC122_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${validations}${spec122OdsDataTable()}${spec122OdsLockedTable()}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const cellCount = Object.keys(O.data.values).length + Object.keys(O.locked.values).length;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    part("meta.xml", loMeta(O.title, O.tableNames.length, cellCount)),
    part("settings.xml", loSettings(O.tableNames, O.activeTable)),
    part("content.xml", content),
    { name: "Thumbnails/thumbnail.png", data: THUMBNAIL_PNG },
    // Calc's common styles plus "Status" and "Bad", the conditional format's style.
    part("styles.xml", SPEC121_ODS_STYLES),
    part("manifest.rdf", LO_MANIFEST_RDF),
    part("META-INF/manifest.xml", LO_MANIFEST),
  ]);
}

// ============================================================================
// SPEC-123 — spec123-styles.xlsx / spec123-styles.ods: named cell styles
// ============================================================================
// The workbooks the formatting commands, the cell styles gallery and the
// named-style readers and writers are proven on (SPEC-123 "Fixtures"), in
// the shapes their producers write them.
// - spec123-styles.xlsx — Excel 365 (Aptos Narrow, the 2023 theme).
//   styles.xml in CT_Stylesheet order: numFmt 164 (`#,##0.0;[Red]\-#,##0.0`,
//   Excel's escape); fonts Normal, Good's #006100, Heading 1's 15 pt bold
//   dk2, the bold "Invoice total", Followed Hyperlink's underlined folHlink
//   and an italic accent1 darker 25% (a negative tint); fills none, gray125,
//   Good's #C6EFCE (no bgColor — Excel's own shape for the style's fill) and
//   accent1 lighter 80% (a positive tint); Heading 1's thick accent1 bottom
//   and Invoice total's thin top + double bottom; cellStyleXfs Normal, Good,
//   Heading 1, Invoice total and Followed Hyperlink, each excluding what it
//   lacks with `applyX="0"`; cellXfs naming them by `xfId` (`applyX="1"`
//   only where a cell differs from its style); cellStyles alphabetically,
//   `builtinId` on the built-ins, `hidden="1"` on Followed Hyperlink and an
//   `xr:uid` on the custom style; one dxf, the table-style defaults and
//   Excel's x14/x15 extLst — none of which a formatting command may disturb.
//   "Budget" (active): row 1 Heading 1 headers A1:E1 (F1 empty), B2:C4 on
//   built-in 2 (0.00), D2:D4 on built-in 10 (0.00%), E2:E4 plain, F3 Good
//   and F4 Good centred, no row 5, row 6 `s customFormat="1"` on Invoice
//   total with "Total" and two SUMs, column H `<col style>` (the italic) —
//   and H6, where the styled row meets the styled column, the style-only
//   cell Excel stores there (the row's style, applied after the column's),
//   which is why the dimension is A1:H6 and every row spans 1:8. "Notes":
//   two texts, a part no formatting flow touches. No calcChain.
// - spec123-styles.ods — LibreOffice 24.8: styles.xml's common table-cell
//   styles Default, Heading, Heading_20_1 ("Heading 1"), Good and
//   Invoice_20_total ("Invoice total") — LO_STYLES with its "Heading"
//   restated as Calc writes it (colour, style and weight with their asian
//   and complex twins) and the three others after it; content.xml's
//   automatic styles in Calc's order — co/ro/ta, the data styles N104
//   (0.00), N105P0 + N105 (#,##0.0, a red minus through `style:map`) and
//   N111 (0.00%), then ce1–ce7, each `style:parent-style-name`d to a common
//   style; column H (co2, wider) with `table:default-cell-style-name="ce6"`;
//   the same values, an empty row 5, row 6 formatted to XFD as Calc writes a
//   whole-row format (its cells, then a ce7 `number-columns-repeated="16381"`
//   run — the reader's Row.styleId) and the trailing run to row 1,048,576.

/** spec123-styles.xlsx's cellXfs indices — a cell's `s`. */
export const SPEC123_XF = {
  normal: 0,
  numbers: 1, // built-in 2 (0.00)
  heading: 2, // the Heading 1 style, as it is
  good: 3, // the Good style, as it is
  total: 4, // the workbook's own "Invoice total" (numFmt 164)
  columnH: 5, // column H's style: italic accent1 darker 25%
  percent: 6, // built-in 10 (0.00%)
  goodCentred: 7, // Good + centred
} as const;

/** spec123-styles.xlsx's one custom number format (Excel escapes the minus). */
export const SPEC123_NUMFMTS = { 164: "#,##0.0;[Red]\\-#,##0.0" } as const;

/** "Budget"'s values, row-major (A1 notation) — both fixtures hold them;
 * a formula cell's value is its cached result. */
const SPEC123_BUDGET_VALUES = {
  A1: "Item", B1: "Q1", C1: "Q2", D1: "Rate", E1: "Units",
  A2: "Rent", B2: 1234.5, C2: 2500, D2: 0.125, E2: 10, F2: "paid",
  A3: "Food", B3: 980.25, C3: 75, D3: 0.2, E3: 20, F3: "OK",
  A4: "Travel", B4: 410, C4: 615.5, D4: 0.05, E4: 30, F4: "late",
  A6: "Total", B6: 2624.75, C6: 3190.5,
} as const;

/** "Notes": two texts no formatting flow touches. */
const SPEC123_NOTES_VALUES = { A1: "Prepared by the finance team", A2: "Figures in US dollars" } as const;

/** The 1-based row number of an A1 reference. */
const a1RowNumber = (ref: string): number => Number(/\d+$/.exec(ref)![0]);

/** What spec123-styles.xlsx holds, for the unit, round-trip and e2e specs.
 * Cells are A1 notation; `index` fields are 0-based. */
export const SPEC123_STYLES_XLSX = {
  title: "Budget",
  sheets: [
    { name: "Budget", sheetId: 1, partName: "xl/worksheets/sheet1.xml", uid: "{4B7D1E93-2C5A-4E8F-9A3B-6C1D8E2F4A31}" },
    { name: "Notes", sheetId: 2, partName: "xl/worksheets/sheet2.xml", uid: "{5C8E2FA4-3D6B-4F9A-8B4C-7D2E9F3A5B42}" },
  ],
  sheetNames: ["Budget", "Notes"],
  /** Budget — the only tabSelected sheetView; workbookView has no activeTab. */
  activeSheetIndex: 0,
  xf: SPEC123_XF,
  numFmts: SPEC123_NUMFMTS,
  /** cellXfs index → its `xfId` (the cellStyleXfs entry, i.e. the named style). */
  xfIds: [0, 0, 2, 1, 3, 0, 0, 1],
  /** The cellStyles, in file order (Excel's alphabetical one): `xfId` the
   * cellStyleXfs index, `includes` the aspects the style xf does not
   * exclude with `applyX="0"` (in STYLE_ASPECTS order). */
  namedStyles: [
    { name: "Followed Hyperlink", xfId: 4, builtinId: 9, hidden: true, includes: ["font"] },
    { name: "Good", xfId: 1, builtinId: 26, includes: ["font", "fill"] },
    { name: "Heading 1", xfId: 2, builtinId: 16, includes: ["font", "borders"] },
    { name: "Invoice total", xfId: 3, uid: "{2E6A9C41-7B3D-4C5E-8F1A-9B4D6E2C8A53}", includes: ["numberFormat", "font", "fill", "borders"] },
    { name: "Normal", xfId: 0, builtinId: 0, includes: ["numberFormat", "font", "fill", "borders", "alignment", "protection"] },
  ],
  /** The theme tints, as written: font 5 (column H) and fill 3 (Invoice total). */
  tints: { columnHFont: "-0.249977111117893", totalFill: "0.79998168889431442" },
  budget: {
    dimension: "A1:H6",
    spans: "1:8",
    activeCell: "A1",
    /** Every stored cell holding a value → the reader's value. */
    values: SPEC123_BUDGET_VALUES,
    /** Every stored cell's `s` (absent: xf 0). */
    styles: {
      A1: SPEC123_XF.heading, B1: SPEC123_XF.heading, C1: SPEC123_XF.heading, D1: SPEC123_XF.heading, E1: SPEC123_XF.heading,
      B2: SPEC123_XF.numbers, C2: SPEC123_XF.numbers, D2: SPEC123_XF.percent,
      B3: SPEC123_XF.numbers, C3: SPEC123_XF.numbers, D3: SPEC123_XF.percent, F3: SPEC123_XF.good,
      B4: SPEC123_XF.numbers, C4: SPEC123_XF.numbers, D4: SPEC123_XF.percent, F4: SPEC123_XF.goodCentred,
      A6: SPEC123_XF.total, B6: SPEC123_XF.total, C6: SPEC123_XF.total, H6: SPEC123_XF.total,
    },
    formulas: { B6: "SUM(B2:B4)", C6: "SUM(C2:C4)" },
    /** Cells stored for their style alone (`<c r s/>`, value null): H6, the
     * styled row × styled column intersection. */
    styleOnly: ["H6"],
    /** F1 holds nothing — the header row stops at E1. */
    emptyCell: "F1",
    /** Rows 1–6 are used; row 5 (index 4) is not written at all. */
    usedRows: 6,
    emptyRow: { index: 4 },
    /** Column H: `<col min="8" max="8" width="12.7109375" style="5" customWidth="1"/>`. */
    styledColumn: { index: 7, letter: "H", styleId: SPEC123_XF.columnH, sourceWidthChars: 12.7109375 },
    /** Row 6: `s="4" customFormat="1"` — Invoice total from A6 to XFD6. */
    styledRow: { index: 5, styleId: SPEC123_XF.total },
    /** What a few positions show, en-US. */
    looks: {
      A1: { bold: true, sizePt: 15, bottomEdge: "thick" },
      B2: { numberFormat: "0.00", text: "1234.50" },
      D2: { numberFormat: "0.00%", text: "12.50%" },
      F3: { color: "#006100", fill: "#c6efce" },
      F4: { color: "#006100", fill: "#c6efce", horizontal: "center" },
      H1: { italic: true },
      B6: { numberFormat: SPEC123_NUMFMTS[164], text: "2,624.8", bold: true, topEdge: "thin", bottomEdge: "double" },
    },
  },
  notes: { values: SPEC123_NOTES_VALUES },
  partNames: {
    styles: "xl/styles.xml",
    theme: "xl/theme/theme1.xml",
    sharedStrings: "xl/sharedStrings.xml",
    budget: "xl/worksheets/sheet1.xml",
    notes: "xl/worksheets/sheet2.xml",
  },
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

/** spec123-styles.xlsx's shared strings, in index order (first use,
 * row-major, sheet order). */
const SPEC123_STRINGS: readonly string[] = [
  ...Object.values(SPEC123_BUDGET_VALUES).filter((value) => typeof value === "string"),
  ...Object.values(SPEC123_NOTES_VALUES),
];

/** A shared-string cell for one of SPEC123_STRINGS. */
function spec123Text(ref: string, text: string, style?: number): string {
  const index = SPEC123_STRINGS.indexOf(text);
  if (index < 0) {
    throw new Error(`"${text}" is not one of spec123-styles.xlsx's strings`);
  }
  return sharedCell(ref, index, style);
}

/** spec123-styles.xlsx's styles part (see SPEC123_XF and the section header). */
const SPEC123_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<numFmts count="1"><numFmt numFmtId="164" formatCode="${excelAttr(SPEC123_NUMFMTS[164])}"/></numFmts>` +
  `<fonts count="6" x14ac:knownFonts="1">` +
  `<font><sz val="11"/><color theme="1"/>${SPEC122_FONT_TAIL}` + // 0 Normal
  `<font><sz val="11"/><color rgb="FF006100"/>${SPEC122_FONT_TAIL}` + // 1 Good
  `<font><b/><sz val="15"/><color theme="3"/>${SPEC122_FONT_TAIL}` + // 2 Heading 1
  `<font><b/><sz val="11"/><color theme="1"/>${SPEC122_FONT_TAIL}` + // 3 Invoice total
  `<font><u/><sz val="11"/><color theme="11"/>${SPEC122_FONT_TAIL}` + // 4 Followed Hyperlink
  `<font><i/><sz val="11"/><color theme="4" tint="${SPEC123_STYLES_XLSX.tints.columnHFont}"/>${SPEC122_FONT_TAIL}` + // 5 column H
  `</fonts>` +
  `<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>` + // 2 Good
  `<fill><patternFill patternType="solid"><fgColor theme="4" tint="${SPEC123_STYLES_XLSX.tints.totalFill}"/><bgColor indexed="64"/></patternFill></fill>` + // 3 Invoice total
  `</fills>` +
  `<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>` +
  `<border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border>` + // 1 Heading 1
  `<border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/></bottom><diagonal/></border>` + // 2 Invoice total
  `</borders>` +
  `<cellStyleXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="164" fontId="3" fillId="3" borderId="2" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="0" fontId="4" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
  `</cellStyleXfs>` +
  `<cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="2"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>` +
  `<xf numFmtId="164" fontId="3" fillId="3" borderId="2" xfId="3"/>` +
  `<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>` +
  `<cellStyles count="5">` +
  SPEC123_STYLES_XLSX.namedStyles
    .map((style) => {
      const builtin = "builtinId" in style ? ` builtinId="${style.builtinId}"` : "";
      const hidden = "hidden" in style ? ` hidden="1"` : "";
      const uid = "uid" in style ? ` xr:uid="${style.uid}"` : "";
      return `<cellStyle name="${style.name}" xfId="${style.xfId}"${builtin}${hidden}${uid}/>`;
    })
    .join("") +
  `</cellStyles>` +
  `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
  `<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>` +
  `<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">` +
  `<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;

/** "Budget" — see the section header and SPEC123_STYLES_XLSX.budget. Its
 * rows are written from the facts: each row's cells in their order, every
 * row with the block's `spans`, row 6 with its style. */
function spec123BudgetSheet(): string {
  const F = SPEC123_STYLES_XLSX;
  const B = F.budget;
  const values: Readonly<Record<string, string | number>> = B.values;
  const styles: Readonly<Record<string, number>> = B.styles;
  const formulas: Readonly<Record<string, string>> = B.formulas;
  const styleOnly: readonly string[] = B.styleOnly;
  const cell = (ref: string): string => {
    const style = styles[ref];
    if (styleOnly.includes(ref)) {
      return styleOnlyCell(ref, style!);
    }
    const value = values[ref]!;
    const formula = formulas[ref];
    if (formula !== undefined) {
      return spec122Formula(ref, normalF(formula), value as number, style);
    }
    return typeof value === "number" ? spec122Number(ref, value, style) : spec123Text(ref, value, style);
  };
  const refs = [...Object.keys(values), ...styleOnly];
  const rows = [...new Set(refs.map(a1RowNumber))].map((r) => {
    const attrs = r === B.styledRow.index + 1 ? ` s="${B.styledRow.styleId}" customFormat="1"` : "";
    return spec122Row(r, B.spans, attrs, refs.filter((ref) => a1RowNumber(ref) === r).map(cell).join(""));
  });
  const H = B.styledColumn;
  return worksheetXml({
    uid: F.sheets[0].uid.slice(1, -1),
    dimension: B.dimension,
    selected: true,
    activeCell: B.activeCell,
    cols: `<cols><col min="${H.index + 1}" max="${H.index + 1}" width="${H.sourceWidthChars}" style="${H.styleId}" customWidth="1"/></cols>`,
    sheetData: rows.join(""),
  });
}

export async function buildSpec123StylesXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC123_STYLES_XLSX;
  const budget = spec123BudgetSheet();
  const notes = worksheetXml({
    uid: F.sheets[1].uid.slice(1, -1),
    dimension: "A1:A2",
    sheetData: Object.entries(F.notes.values)
      .map(([ref, text]) => excelRow(a1RowNumber(ref), "1:1", spec123Text(ref, text)))
      .join(""),
  });
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: [
      { name: F.sheetNames[0], xml: budget },
      { name: F.sheetNames[1], xml: notes },
    ],
    sharedStrings: sharedStringsXml(SPEC123_STRINGS, sharedStringReferences(budget) + sharedStringReferences(notes)),
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    styles: SPEC123_STYLES,
  });
}

// --- spec123-styles.ods ---------------------------------------------------------

/** What spec123-styles.ods holds, for the unit, round-trip and e2e specs.
 * Cells are A1 notation; `index` fields are 0-based. */
export const SPEC123_STYLES_ODS = {
  title: "Budget",
  tableNames: ["Budget", "Notes"],
  activeTable: "Budget",
  activeSheetIndex: 0,
  budget: {
    /** Every stored cell → the reader's value (the same as the xlsx's,
     * whose H6 Calc does not store: row 6's style reaches it). */
    values: SPEC123_BUDGET_VALUES,
    /** Each value cell's automatic style (absent: none). */
    styles: {
      A1: "ce1", B1: "ce1", C1: "ce1", D1: "ce1", E1: "ce1",
      B2: "ce2", C2: "ce2", D2: "ce3",
      B3: "ce2", C3: "ce2", D3: "ce3", F3: "ce4",
      B4: "ce2", C4: "ce2", D4: "ce3", F4: "ce5",
      A6: "ce7", B6: "ce7", C6: "ce7",
    },
    /** Formula cells: `table:formula`, the canonical formula, the cached
     * value and the text Calc wrote for it. */
    formulas: {
      B6: { wire: "of:=SUM([.B2:.B4])", formula: "SUM(B2:B4)", value: 2624.75, text: "2,624.8" },
      C6: { wire: "of:=SUM([.C2:.C4])", formula: "SUM(C2:C4)", value: 3190.5, text: "3,190.5" },
    },
    emptyCell: "F1",
    /** Rows 1–6 are written; row 5 (index 4) is an empty row; the trailing
     * run reaches the grid's edge. */
    usedRows: 6,
    emptyRow: { index: 4 },
    trailingRowRun: 1_048_576 - 6,
    /** Column H: `table:style-name="co2"` (wider) and
     * `table:default-cell-style-name="ce6"` (italic). */
    columnDefault: { index: 7, letter: "H", styleName: "ce6", columnStyle: "co2" },
    /** Row 6: A6:C6 in ce7, then a ce7 run of 16381 empty cells to XFD — the
     * row's style, storing no cell of its own. */
    styledRow: { index: 5, styleName: "ce7", repeatedRun: 16_381 },
    /** What a few positions show, en-US. */
    looks: {
      A1: { bold: true, sizePt: 18 },
      B2: { numberFormat: "0.00", text: "1234.50" },
      D2: { numberFormat: "0.00%", text: "12.50%" },
      F3: { color: "#006600", fill: "#ccffcc" },
      F4: { color: "#006600", fill: "#ccffcc", horizontal: "center" },
      H1: { italic: true },
      B6: { numberFormat: "#,##0.0;[Red]-#,##0.0", text: "2,624.8", bold: true, topEdge: "thin", bottomEdge: "double", borderColor: "#156082" },
    },
  },
  notes: { values: SPEC123_NOTES_VALUES },
  /** styles.xml's common table-cell styles, in file order: `includes` the
   * aspects each sets, itself or through its parents (STYLE_ASPECTS order). */
  commonStyles: [
    { name: "Default" },
    { name: "Heading", parent: "Default", includes: ["font"] },
    { name: "Heading_20_1", displayName: "Heading 1", parent: "Heading", includes: ["font"] },
    { name: "Good", parent: "Default", includes: ["font", "fill"] },
    { name: "Invoice_20_total", displayName: "Invoice total", parent: "Default", includes: ["font", "borders"] },
  ],
  /** content.xml's automatic cell styles: the common style each is
   * parented to, its data style and what it adds of its own. */
  cellStyles: {
    ce1: { parent: "Heading_20_1" },
    ce2: { parent: "Default", dataStyle: "N104" },
    ce3: { parent: "Default", dataStyle: "N111" },
    ce4: { parent: "Good" },
    ce5: { parent: "Good", horizontal: "center" },
    ce6: { parent: "Default", italic: true },
    ce7: { parent: "Invoice_20_total", dataStyle: "N105" },
  },
  /** The data styles' codes as SPEC-118 reads them (N105 is the N105P0 pair). */
  dataStyles: { N104: "0.00", N105: "#,##0.0;[Red]-#,##0.0", N111: "0.00%" },
  /** Every entry, in package order (mimetype first, STOREd). */
  parts: ["mimetype", "meta.xml", "settings.xml", "content.xml", "Thumbnails/thumbnail.png", "styles.xml", "manifest.rdf", "META-INF/manifest.xml"],
} as const;

/** styles.xml: LO_STYLES with Calc's own "Heading" (its asian and complex
 * twins) followed by "Heading 1", "Good" and the workbook's "Invoice total". */
const SPEC123_ODS_STYLES = LO_STYLES.replace(
  `<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-size="24pt" fo:font-weight="bold"/></style:style>`,
  `<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:color="#000000" fo:font-size="24pt" ` +
    `fo:font-style="normal" fo:font-weight="bold" style:font-size-asian="24pt" style:font-style-asian="normal" style:font-weight-asian="bold" ` +
    `style:font-size-complex="24pt" style:font-style-complex="normal" style:font-weight-complex="bold"/></style:style>` +
    `<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="table-cell" style:parent-style-name="Heading">` +
    `<style:text-properties fo:font-size="18pt" style:font-size-asian="18pt" style:font-size-complex="18pt"/></style:style>` +
    `<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ccffcc"/>` +
    `<style:text-properties fo:color="#006600"/></style:style>` +
    `<style:style style:name="Invoice_20_total" style:display-name="Invoice total" style:family="table-cell" style:parent-style-name="Default">` +
    `<style:table-cell-properties style:border-line-width-bottom="0.0102in 0.0157in 0.0102in" fo:border-bottom="2.6pt double #156082" ` +
    `fo:border-left="none" fo:border-right="none" fo:border-top="0.74pt solid #156082"/>` +
    `<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>`,
);

/** content.xml's automatic styles, in Calc's order: columns (co2 column H's
 * wider width), the row, the table, the data styles, the cell styles. */
const SPEC123_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.9272in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<number:number-style style:name="N104"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/></number:number-style>` +
  `<number:number-style style:name="N105P0" style:volatile="true"><number:number number:decimal-places="1" number:min-decimal-places="1" ` +
  `number:min-integer-digits="1" number:grouping="true"/></number:number-style>` +
  `<number:number-style style:name="N105"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
  `<number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1" number:grouping="true"/>` +
  `<style:map style:condition="value()&gt;=0" style:apply-style-name="N105P0"/></number:number-style>` +
  `<number:percentage-style style:name="N111"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
  `<number:text>%</number:text></number:percentage-style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading_20_1"/>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N104"/>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N111"/>` +
  `<style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Good"/>` +
  `<style:style style:name="ce5" style:family="table-cell" style:parent-style-name="Good"><style:table-cell-properties style:text-align-source="fix" ` +
  `style:repeat-content="false"/><style:paragraph-properties fo:text-align="center" fo:margin-left="0in"/></style:style>` +
  `<style:style style:name="ce6" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic" ` +
  `style:font-style-asian="italic" style:font-style-complex="italic"/></style:style>` +
  `<style:style style:name="ce7" style:family="table-cell" style:parent-style-name="Invoice_20_total" style:data-style-name="N105"/>` +
  `</office:automatic-styles>`;

/** "Budget" — see the section header and SPEC123_STYLES_ODS.budget: each
 * written row's cells from the facts, then its run to XFD (row 6's styled). */
function spec123OdsBudgetTable(): string {
  const B = SPEC123_STYLES_ODS.budget;
  const values: Readonly<Record<string, string | number>> = B.values;
  const styles: Readonly<Record<string, string>> = B.styles;
  const formulas: Readonly<Record<string, { wire: string; text: string }>> = B.formulas;
  /** The text Calc shows for a number in its cell's data style. */
  const shown = (ref: string, value: number): string => {
    const formula = formulas[ref];
    if (formula !== undefined) {
      return formula.text;
    }
    const style = styles[ref];
    return style === "ce2" ? value.toFixed(2) : style === "ce3" ? `${(value * 100).toFixed(2)}%` : String(value);
  };
  const cell = (ref: string): string => {
    const value = values[ref]!;
    const style = styles[ref];
    if (typeof value === "string") {
      return loString(value, style);
    }
    const type = style === "ce3" ? "percentage" : "float";
    const formula = formulas[ref] === undefined ? "" : `table:formula="${formulas[ref].wire}" `;
    return loValueCell(style, `${formula}office:value-type="${type}" office:value="${value}" calcext:value-type="${type}"`, shown(ref, value));
  };
  const H = B.columnDefault;
  const R = B.styledRow;
  const rows: string[] = [];
  for (let r = 1; r <= B.usedRows; r += 1) {
    const refs = Object.keys(values).filter((ref) => a1RowNumber(ref) === r);
    const rest = 16_384 - refs.length;
    rows.push(loRow(refs.map(cell).join("") + (r === R.index + 1 ? loStyledRun(R.styleName, rest) : loEmpty(rest))));
  }
  return (
    `<table:table table:name="${SPEC123_STYLES_ODS.tableNames[0]}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="${H.index}" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="${H.columnStyle}" table:default-cell-style-name="${H.styleName}"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="${16_384 - H.index - 1}" table:default-cell-style-name="Default"/>` +
    rows.join("") +
    loRow(loEmpty(16_384), "ro1", B.trailingRowRun) +
    `</table:table>`
  );
}

export async function buildSpec123StylesOdsBytes(): Promise<Uint8Array> {
  const O = SPEC123_STYLES_ODS;
  const notesValues = Object.values(O.notes.values);
  const notes =
    `<table:table table:name="${O.tableNames[1]}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    notesValues.map((text) => loRow(loString(text) + loEmpty(16_383))).join("") +
    loRow(loEmpty(16_384), "ro1", 1_048_576 - notesValues.length) +
    `</table:table>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC123_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${spec123OdsBudgetTable()}${notes}<table:named-expressions/>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const cellCount = Object.keys(O.budget.values).length + notesValues.length;
  return loPackage(content, loMeta(O.title, O.tableNames.length, cellCount), loSettings(O.tableNames, O.activeTable), SPEC123_ODS_STYLES);
}

// ============================================================================
// SPEC-124 — spec124-structure.xlsx / spec124-structure.ods: structure
// ============================================================================
// The workbooks the structural round trip (tables, charts, drawings,
// comments, validations, conditional formatting, merges, frozen panes and
// named ranges) is proven on — every construct must survive an untouched
// download byte-for-byte, the same ground SPEC-121's fixtures stand on.
// - spec124-structure.xlsx — Excel 365-shaped. "Sheet1" (active, its first
//   column frozen): 30 data rows (B qty, C price, D a cycling A/B/C code
//   validated against a local P1:P3 list, E a description — the longest at
//   E12 — F a shared-formula total B*C), a plain formula and a legacy array
//   on H, a table ("Rates" over J1:K6, its own autoFilter), two merges
//   (M1:N1, M2:N3), a workbook autoFilter, a conditional format over C, a
//   hyperlink to Sheet2, a drawing with a picture and a chart, a legacy
//   comment with its VML note, and an x14 list validation naming Sheet2.
//   "Sheet2": cross-sheet, 3-D and literal formulas. calcChain.xml,
//   docProps/app.xml's Worksheets/Named Ranges heading pairs.
// - spec124-structure.ods — its LibreOffice-shaped twin: a content
//   validation, a formula column, a chart frame embedded directly in a
//   cell (not `table:shapes`), a merge, a named range, a named expression,
//   a database range and a table-local conditional format; settings.xml's
//   per-table frozen-pane entries.
// SPEC124_STRUCTURE_XLSX and SPEC124_STRUCTURE_ODS record the facts.

const SPEC124_ITEM_COUNT = 30;
const SPEC124_LONG_DESCRIPTION = "An unusually long description that needs a much wider column";
const SPEC124_CODES = ["A", "B", "C"] as const;

/** Sheet1's 30 data rows (r = 2..31, i = r-1): item, qty (i), price (10+i),
 * a cycling code, a description (row 12 the longest) and the shared
 * formula's total (qty*price). */
const SPEC124_ROWS = Array.from({ length: SPEC124_ITEM_COUNT }, (_, index) => {
  const i = index + 1;
  const row = i + 1;
  return {
    row,
    item: `Item ${i}`,
    qty: i,
    price: 10 + i,
    code: SPEC124_CODES[(i - 1) % 3]!,
    description: row === 12 ? SPEC124_LONG_DESCRIPTION : `Note ${i}`,
    total: i * (10 + i),
  };
});

/** What spec124-structure.xlsx holds, for the unit and e2e specs. */
export const SPEC124_STRUCTURE_XLSX = {
  title: "Structure",
  sheetNames: ["Sheet1", "Sheet2"],
  activeSheetIndex: 0,
  rows: SPEC124_ROWS,
  headers: { A1: "Item", B1: "Qty", C1: "Price", D1: "Code", E1: "Description", F1: "Total" } as Record<string, string>,
  /** First-column freeze: `<pane xSplit="1" .../>`. */
  frozen: { rows: 0, cols: 1 },
  topLeftCell: "B1",
  /** Column C's custom width — every other column default. */
  cols: [{ min: 2, max: 2, width: "12.7109375" }],
  sharedFormula: { master: "F2", ref: "F2:F31", si: 0, wire: "B2*C2" },
  simpleFormula: { ref: "H5", formula: "B5*2", value: 8 },
  arrayFormula: { anchor: "H20", ref: "H20:H21", formula: "B20:B21*2", values: { H20: 38, H21: 40 } },
  table: {
    partName: "xl/tables/table1.xml",
    id: 1,
    name: "Rates",
    displayName: "Rates",
    ref: "J1:K6",
    columns: ["Region", "Rate"],
    regions: ["North", "South", "East", "West", "Central"],
    rates: [0.1, 0.2, 0.3, 0.4, 0.5],
    uid: "{0A1B2C3D-1001-4AA1-8B1C-2D3E4F5A6B70}",
    columnUids: ["{0A1B2C3D-1002-4AA1-8B1C-2D3E4F5A6B70}", "{0A1B2C3D-1003-4AA1-8B1C-2D3E4F5A6B70}"],
  },
  merges: [
    { ref: "M1:N1", value: "Notes" },
    { ref: "M2:N3", value: "Reviewed quarterly" },
  ],
  validationList: { ref: "P1:P3", values: ["A", "B", "C"] },
  autoFilter: { ref: "A1:F31" },
  conditionalFormatting: { sqref: "C2:C31", formula: "C2>AVERAGE($C$2:$C$31)", dxfId: 0 },
  dataValidation: { sqref: "D2:D31", formula1: "$P$1:$P$3", uid: "{0A1B2C3D-1004-4AA1-8B1C-2D3E4F5A6B70}" },
  hyperlink: { ref: "A31", location: "Sheet2!A1", display: "Item 30", uid: "{0A1B2C3D-1005-4AA1-8B1C-2D3E4F5A6B70}" },
  drawing: {
    partName: "xl/drawings/drawing1.xml",
    rels: "xl/drawings/_rels/drawing1.xml.rels",
    anchors: {
      picture: { from: { col: 7, row: 22 }, to: { col: 9, row: 28 }, embedRId: "rId2" },
      chart: { from: { col: 11, row: 7 }, to: { col: 17, row: 21 }, rId: "rId1" },
    },
  },
  image: { partName: "xl/media/image1.png" },
  chart: {
    partName: "xl/charts/chart1.xml",
    /** Every `c:f` text, in document order: series name, categories, values. */
    formulas: ["Sheet1!$C$1", "Sheet1!$A$2:$A$31", "Sheet1!$C$2:$C$31"],
  },
  comment: { partName: "xl/comments1.xml", ref: "B3", author: FIXTURE_WORKBOOK_AUTHOR, text: "Check this quantity." },
  /** The legacy comment's VML note: no threaded comment pairs with it. */
  vml: { partName: "xl/drawings/vmlDrawing1.vml", anchor: "2, 15, 1, 10, 4, 15, 5, 4", row: 2, column: 1, idmap: 1, shapeId: "_x0000_s1025" },
  x14Validation: { sqref: "A2:A31", formula1: "Sheet2!$D$1:$D$3", uid: "{0A1B2C3D-1006-4AA1-8B1C-2D3E4F5A6B70}" },
  /** xl/worksheets/_rels/sheet1.xml.rels, by id. */
  sheetRelationships: {
    rId1: { type: `${REL_TYPE}/drawing`, target: "../drawings/drawing1.xml" },
    rId2: { type: `${REL_TYPE}/vmlDrawing`, target: "../drawings/vmlDrawing1.vml" },
    rId3: { type: `${REL_TYPE}/table`, target: "../tables/table1.xml" },
    rId4: { type: `${REL_TYPE}/comments`, target: "../comments1.xml" },
  },
  sheet2: {
    values: { B1: 0.2, B5: 10, D1: "North", D2: "South", D3: "West" } as Record<string, number | string>,
    formulas: {
      A1: { wire: "Sheet1!A5", formula: "Sheet1!A5", value: "Item 4", t: "str" as const },
      A2: { wire: "SUM('Sheet1'!B2:B20)", formula: "SUM('Sheet1'!B2:B20)", value: 190 },
      A3: { wire: "SUM(Sheet1:Sheet2!A1)", formula: "SUM(Sheet1:Sheet2!A1)", value: 0 },
      A4: { wire: "Sheet1!B5*2", formula: "Sheet1!B5*2", value: 8 },
      A5: { wire: "B5+1", formula: "B5+1", value: 11 },
    },
  },
  /** workbook.xml's definedNames, in file order. */
  definedNames: [
    { name: "_xlnm._FilterDatabase", localSheetId: 0, hidden: true, text: "Sheet1!$A$1:$F$31" },
    { name: "_xlnm.Print_Area", localSheetId: 0, text: "Sheet1!$A$1:$F$31" },
    { name: "Prices", text: "Sheet1!$C$2:$C$31" },
    { name: "TaxRate", text: "Sheet2!$B$1" },
  ],
  /** docProps/app.xml: HeadingPairs and TitlesOfParts. */
  headingPairs: [
    ["Worksheets", 2],
    ["Named Ranges", 3],
  ] as [string, number][],
  titlesOfParts: ["Sheet1", "Sheet2", "Sheet1!Print_Area", "Prices", "TaxRate"],
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/charts/chart1.xml",
    "xl/tables/table1.xml",
    "xl/drawings/vmlDrawing1.vml",
    "xl/comments1.xml",
    "xl/media/image1.png",
    "xl/calcChain.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

const SPEC124_SHEET1_UID = "4A5B6C7D-0010-4AA1-8B1C-2D3E4F5A6B70";
const SPEC124_SHEET2_UID = "5B6C7D8E-0011-4AA1-8B1C-2D3E4F5A6B70";

/** Excel's styles part: Normal, a bold header xf and the conditional
 * format's dxf 0 (light red fill, dark red text — Excel's preset). */
const SPEC124_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

const SPEC124_HEADER_XF = 1;

/** A `c:strCache` over strings. */
function spec124StrCache(values: readonly string[]): string {
  return (
    `<c:strCache><c:ptCount val="${values.length}"/>` +
    values.map((value, idx) => `<c:pt idx="${idx}"><c:v>${xmlText(value)}</c:v></c:pt>`).join("") +
    `</c:strCache>`
  );
}

/** xl/charts/chart1.xml: a clustered bar chart over Sheet1!$C$2:$C$31 by
 * Sheet1!$A$2:$A$31, its series named by Sheet1!$C$1 — no separate style or
 * colours part (unlike spec121ChartXml, which Excel also writes those
 * alongside). */
function spec124ChartXml(): string {
  const F = SPEC124_STRUCTURE_XLSX;
  const [seriesName, categories, values] = F.chart.formulas;
  const items = F.rows.map((row) => row.item);
  const prices = F.rows.map((row) => row.price);
  return (
    XML_DECLARATION +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart><c:title><c:overlay val="0"/>${spec121ChartText(1400, 65000, 35000)}</c:title><c:autoTitleDeleted val="0"/>` +
    `<c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/>` +
    `<c:tx><c:strRef><c:f>${seriesName}</c:f>${spec124StrCache(["Price"])}</c:strRef></c:tx>` +
    `<c:invertIfNegative val="0"/>` +
    `<c:cat><c:strRef><c:f>${categories}</c:f>${spec124StrCache(items)}</c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>${values}</c:f>${spec121NumCache(prices)}</c:numRef></c:val>` +
    `</c:ser><c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
    `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
    `<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${SPEC121_CHART_LINE}${spec121ChartText(900, 65000, 35000)}` +
    `<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
    `<c:majorGridlines>${SPEC121_CHART_LINE}</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${spec121ChartText(900, 65000, 35000)}<c:crossAx val="1"/><c:crosses val="autoZero"/>` +
    `<c:crossBetween val="between"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `</c:chartSpace>`
  );
}

/** xl/drawings/drawing1.xml: a picture (oneCell) and the chart's
 * graphicFrame, each its own twoCellAnchor. */
/** A `<xdr:from>`/`<xdr:to>` pair at zero offset. */
function spec124DrawingAnchor(from: { col: number; row: number }, to: { col: number; row: number }): string {
  return (
    `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`
  );
}

function spec124DrawingXml(): string {
  const D = SPEC124_STRUCTURE_XLSX.drawing.anchors;
  const anchor = spec124DrawingAnchor;
  return (
    XML_DECLARATION +
    `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">` +
    `<xdr:twoCellAnchor editAs="oneCell">${anchor(D.picture.from, D.picture.to)}` +
    `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${D.picture.embedRId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
    `<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${anchor(D.chart.from, D.chart.to)}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${D.chart.rId}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>` +
    `</xdr:wsDr>`
  );
}

/** xl/tables/table1.xml — "Rates" over J1:K6, with its own autoFilter
 * (Excel shape, like spec121TableXml). */
function spec124TableXml(): string {
  const T = SPEC124_STRUCTURE_XLSX.table;
  const columns = T.columns.map((name, idx) => `<tableColumn id="${idx + 1}" xr3:uid="${T.columnUids[idx]}" name="${name}"/>`).join("");
  return (
    XML_DECLARATION +
    `<table xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" ` +
    `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" ` +
    `id="${T.id}" xr:uid="${T.uid}" name="${T.name}" displayName="${T.displayName}" ref="${T.ref}" totalsRowShown="0">` +
    `<autoFilter ref="${T.ref}"/>` +
    `<tableColumns count="${T.columns.length}">${columns}</tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`
  );
}

/** xl/drawings/vmlDrawing1.vml — the legacy comment's note shape, as Excel
 * writes VML (see SPEC121_VML): no declaration, CRLF line breaks. */
const SPEC124_VML = [
  `<xml xmlns:v="urn:schemas-microsoft-com:vml"`,
  ` xmlns:o="urn:schemas-microsoft-com:office:office"`,
  ` xmlns:x="urn:schemas-microsoft-com:office:excel">`,
  ` <o:shapelayout v:ext="edit">`,
  `  <o:idmap v:ext="edit" data="${SPEC124_STRUCTURE_XLSX.vml.idmap}"/>`,
  ` </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"`,
  `  path="m,l,21600r21600,l21600,xe">`,
  `  <v:stroke joinstyle="miter"/>`,
  `  <v:path gradientshapeok="t" o:connecttype="rect"/>`,
  ` </v:shapetype><v:shape id="${SPEC124_STRUCTURE_XLSX.vml.shapeId}" o:spid="${SPEC124_STRUCTURE_XLSX.vml.shapeId}" type="#_x0000_t202" style='position:absolute;`,
  `  margin-left:107.25pt;margin-top:16pt;width:108pt;height:59.25pt;z-index:1;`,
  `  visibility:hidden' fillcolor="infoBackground [80]" o:insetmode="auto">`,
  `  <v:fill color2="infoBackground [80]"/>`,
  `  <v:shadow color="none [81]" obscured="t"/>`,
  `  <v:path o:connecttype="none"/>`,
  `  <v:textbox style='mso-direction-alt:auto'>`,
  `   <div style='text-align:left'></div>`,
  `  </v:textbox>`,
  `  <x:ClientData ObjectType="Note">`,
  `   <x:MoveWithCells/>`,
  `   <x:SizeWithCells/>`,
  `   <x:Anchor>`,
  `    ${SPEC124_STRUCTURE_XLSX.vml.anchor}</x:Anchor>`,
  `   <x:AutoFill>False</x:AutoFill>`,
  `   <x:Row>${SPEC124_STRUCTURE_XLSX.vml.row}</x:Row>`,
  `   <x:Column>${SPEC124_STRUCTURE_XLSX.vml.column}</x:Column>`,
  `  </x:ClientData>`,
  ` </v:shape></xml>`,
].join("\r\n");

/** xl/comments1.xml — a plain legacy comment (no threaded comment pairs
 * with it this time). */
function spec124CommentsXml(): string {
  const C = SPEC124_STRUCTURE_XLSX.comment;
  return (
    XML_DECLARATION +
    `<comments xmlns="${MAIN_NS}"><authors><author>${C.author}</author></authors>` +
    `<commentList><comment ref="${C.ref}" authorId="0" shapeId="0"><text><t>${xmlText(C.text)}</t></text></comment></commentList></comments>`
  );
}

/** The `<extLst>` x14 data validation naming Sheet2's D1:D3 list. */
function spec124X14ValidationsXml(): string {
  const X = SPEC124_STRUCTURE_XLSX.x14Validation;
  return (
    `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
    `<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
    `<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" xr:uid="${X.uid}">` +
    `<x14:formula1><xm:f>${X.formula1}</xm:f></x14:formula1><xm:sqref>${X.sqref}</xm:sqref>` +
    `</x14:dataValidation></x14:dataValidations></ext></extLst>`
  );
}

/** docProps/app.xml with the Worksheets/Named Ranges heading pairs (see
 * spec121AppXml). */
function spec124AppXml(): string {
  const F = SPEC124_STRUCTURE_XLSX;
  const pairs = F.headingPairs.map(([heading, count]) => `<vt:variant><vt:lpstr>${heading}</vt:lpstr></vt:variant><vt:variant><vt:i4>${count}</vt:i4></vt:variant>`);
  return (
    XML_DECLARATION +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>` +
    `<HeadingPairs><vt:vector size="${pairs.length * 2}" baseType="variant">${pairs.join("")}</vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${F.titlesOfParts.length}" baseType="lpstr">${F.titlesOfParts.map((title) => `<vt:lpstr>${title}</vt:lpstr>`).join("")}` +
    `</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>` +
    `<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`
  );
}

/** Sheet1's `<sheetData>`: the 30-row block (A..F), H5 and the H20:H21
 * array, the "Rates" table (J:K), the merges (M) and the local validation
 * list (P), each row's cells assembled in ascending column order. */
function spec124Sheet1Rows(str: (value: string) => number): string {
  const F = SPEC124_STRUCTURE_XLSX;

  const mainCells = new Map<number, string>();
  mainCells.set(
    1,
    (["A1", "B1", "C1", "D1", "E1", "F1"] as const).map((ref) => sharedCell(ref, str(F.headers[ref]!), SPEC124_HEADER_XF)).join(""),
  );
  for (const data of F.rows) {
    const r = data.row;
    const totalCell =
      r === 2
        ? formulaCell(`F${r}`, `<f t="shared" ref="${F.sharedFormula.ref}" si="${F.sharedFormula.si}">${xmlText(F.sharedFormula.wire)}</f>`, data.total)
        : follower(`F${r}`, F.sharedFormula.si, data.total);
    mainCells.set(
      r,
      sharedCell(`A${r}`, str(data.item)) + numberCell(`B${r}`, data.qty) + numberCell(`C${r}`, data.price) + sharedCell(`D${r}`, str(data.code)) +
        sharedCell(`E${r}`, str(data.description)) + totalCell,
    );
  }

  const extra = new Map<number, string>([
    [5, formulaCell("H5", normalF(F.simpleFormula.formula), F.simpleFormula.value)],
    [20, formulaCell("H20", `<f t="array" ref="${F.arrayFormula.ref}">${xmlText(F.arrayFormula.formula)}</f>`, F.arrayFormula.values.H20)],
    [21, numberCell("H21", F.arrayFormula.values.H21)],
  ]);

  const tableCells = new Map<number, string>();
  tableCells.set(1, sharedCell("J1", str(F.table.columns[0]!)) + sharedCell("K1", str(F.table.columns[1]!)));
  F.table.regions.forEach((region, index) => {
    const r = index + 2;
    tableCells.set(r, sharedCell(`J${r}`, str(region)) + numberCell(`K${r}`, F.table.rates[index]!));
  });

  const mergeCells = new Map<number, string>([
    [1, sharedCell("M1", str(F.merges[0]!.value))],
    [2, sharedCell("M2", str(F.merges[1]!.value))],
  ]);

  const listCells = new Map<number, string>();
  F.validationList.values.forEach((value, index) => listCells.set(index + 1, sharedCell(`P${index + 1}`, str(value))));

  const rows: string[] = [];
  for (let r = 1; r <= 31; r += 1) {
    const cells = (mainCells.get(r) ?? "") + (extra.get(r) ?? "") + (tableCells.get(r) ?? "") + (mergeCells.get(r) ?? "") + (listCells.get(r) ?? "");
    if (cells !== "") {
      rows.push(excelRow(r, r <= 16 ? "1:16" : "1:8", cells));
    }
  }
  return rows.join("");
}

/** Sheet1 in Excel 365's CT_Worksheet order: dimension, sheetViews,
 * sheetFormatPr, cols, sheetData, autoFilter, mergeCells,
 * conditionalFormatting, dataValidations, hyperlinks, pageMargins, drawing,
 * legacyDrawing, tableParts, extLst. */
function spec124Sheet1Xml(str: (value: string) => number): string {
  const F = SPEC124_STRUCTURE_XLSX;
  const col = F.cols[0]!;
  return (
    XML_DECLARATION +
    worksheetOpen(SPEC124_SHEET1_UID) +
    `<dimension ref="A1:P31"/>` +
    `<sheetViews><sheetView tabSelected="1" workbookViewId="0">` +
    `<pane xSplit="${F.frozen.cols}" topLeftCell="${F.topLeftCell}" activePane="topRight" state="frozen"/>` +
    `<selection pane="topRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>` +
    `<cols><col min="${col.min + 1}" max="${col.max + 1}" width="${col.width}" customWidth="1"/></cols>` +
    `<sheetData>${spec124Sheet1Rows(str)}</sheetData>` +
    `<autoFilter ref="${F.autoFilter.ref}"/>` +
    `<mergeCells count="${F.merges.length}">${F.merges.map((merge) => `<mergeCell ref="${merge.ref}"/>`).join("")}</mergeCells>` +
    `<conditionalFormatting sqref="${F.conditionalFormatting.sqref}"><cfRule type="expression" dxfId="${F.conditionalFormatting.dxfId}" priority="1">` +
    `<formula>${xmlText(F.conditionalFormatting.formula)}</formula></cfRule></conditionalFormatting>` +
    `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${F.dataValidation.sqref}" ` +
    `xr:uid="${F.dataValidation.uid}"><formula1>${F.dataValidation.formula1}</formula1></dataValidation></dataValidations>` +
    `<hyperlinks><hyperlink ref="${F.hyperlink.ref}" location="${F.hyperlink.location}" display="${F.hyperlink.display}" xr:uid="${F.hyperlink.uid}"/></hyperlinks>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `<drawing r:id="rId1"/><legacyDrawing r:id="rId2"/>` +
    `<tableParts count="1"><tablePart r:id="rId3"/></tableParts>` +
    spec124X14ValidationsXml() +
    `</worksheet>`
  );
}

/** Sheet2's `<sheetData>`: cross-sheet, 3-D and literal formulas/values. */
function spec124Sheet2Rows(str: (value: string) => number): string {
  const S = SPEC124_STRUCTURE_XLSX.sheet2;
  return (
    excelRow(
      1,
      "1:4",
      formulaCell("A1", normalF(S.formulas.A1.wire), S.formulas.A1.value, { t: S.formulas.A1.t }) +
        numberCell("B1", S.values.B1 as number) +
        sharedCell("D1", str(S.values.D1 as string)),
    ) +
    excelRow(2, "1:4", formulaCell("A2", normalF(S.formulas.A2.wire), S.formulas.A2.value) + sharedCell("D2", str(S.values.D2 as string))) +
    excelRow(3, "1:4", formulaCell("A3", normalF(S.formulas.A3.wire), S.formulas.A3.value) + sharedCell("D3", str(S.values.D3 as string))) +
    excelRow(4, "1:4", formulaCell("A4", normalF(S.formulas.A4.wire), S.formulas.A4.value)) +
    excelRow(5, "1:4", formulaCell("A5", normalF(S.formulas.A5.wire), S.formulas.A5.value) + numberCell("B5", S.values.B5 as number))
  );
}

export async function buildSpec124StructureXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC124_STRUCTURE_XLSX;
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  const str = (value: string): number => {
    let index = sstIndex.get(value);
    if (index === undefined) {
      index = sst.length;
      sst.push(value);
      sstIndex.set(value, index);
    }
    return index;
  };

  const sheet1 = spec124Sheet1Xml(str);
  const sheet2 = worksheetXml({ uid: SPEC124_SHEET2_UID, dimension: "A1:D5", sheetData: spec124Sheet2Rows(str) });
  const worksheets = [sheet1, sheet2];

  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${sst.length}">` +
    sst.map((text) => `<si><t>${xmlText(text)}</t></si>`).join("") +
    `</sst>`;

  const calcEntries: { r: string; i: number; array?: true }[] = [
    ...F.rows.map((row) => ({ r: `F${row.row}`, i: 1 })),
    { r: "H5", i: 1 },
    { r: "H20", i: 1, array: true as const },
    ...(["A1", "A2", "A3", "A4", "A5"] as const).map((r) => ({ r, i: 2 })),
  ];
  const calcChain =
    XML_DECLARATION +
    `<calcChain xmlns="${MAIN_NS}">` +
    calcEntries
      .map((entry, index) => {
        const first = index === 0 || calcEntries[index - 1]!.i !== entry.i;
        return `<c r="${entry.r}"${first ? ` i="${entry.i}"` : ""}${index === 0 ? ` l="1"` : ""}${entry.array ? ` a="1"` : ""}/>`;
      })
      .join("") +
    `</calcChain>`;

  const definedNames =
    `<definedNames>` +
    F.definedNames
      .map(
        (name) =>
          `<definedName name="${name.name}"${"localSheetId" in name ? ` localSheetId="${name.localSheetId}"` : ""}${"hidden" in name && name.hidden ? ` hidden="1"` : ""}>${name.text}</definedName>`,
      )
      .join("") +
    `</definedNames>`;

  const rels = (relationships: string): string => XML_DECLARATION + `<Relationships xmlns="${PKG_REL_NS}">${relationships}</Relationships>`;
  const sheetRels = Object.entries(F.sheetRelationships)
    .map(([id, rel]) => `<Relationship Id="${id}" Type="${rel.type}" Target="${rel.target}"/>`)
    .join("");
  const D = F.drawing.anchors;

  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: [
      { name: F.sheetNames[0], xml: sheet1 },
      { name: F.sheetNames[1], xml: sheet2 },
    ],
    sharedStrings,
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    styles: SPEC124_STYLES,
    defaults: [
      `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`,
      `<Default Extension="png" ContentType="image/png"/>`,
    ],
    app: spec124AppXml(),
    additions: {
      parts: [
        part("xl/worksheets/_rels/sheet1.xml.rels", rels(sheetRels)),
        part(F.drawing.partName, spec124DrawingXml()),
        part(
          F.drawing.rels,
          rels(
            `<Relationship Id="${D.chart.rId}" Type="${REL_TYPE}/chart" Target="../charts/chart1.xml"/>` +
              `<Relationship Id="${D.picture.embedRId}" Type="${REL_TYPE}/image" Target="../media/image1.png"/>`,
          ),
        ),
        part(F.chart.partName, spec124ChartXml()),
        part(F.table.partName, spec124TableXml()),
        part(F.vml.partName, SPEC124_VML),
        part(F.comment.partName, spec124CommentsXml()),
        { name: F.image.partName, data: THUMBNAIL_PNG },
      ],
      relationships: [],
      overrides: [
        `<Override PartName="/${F.drawing.partName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
        `<Override PartName="/${F.chart.partName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
        `<Override PartName="/${F.table.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
        `<Override PartName="/${F.comment.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
      ],
      workbookAfterSheets: definedNames,
    },
  });
}

// --- spec124-structure.ods ---------------------------------------------------------

/** Sheet1's 10 data rows (r = 2..11, i = r-1): item, amount (i*100), a
 * cycling code and the doubling formula's value. */
const SPEC124_ODS_ROWS = Array.from({ length: 10 }, (_, index) => {
  const i = index + 1;
  const row = i + 1;
  return { row, item: `Item ${i}`, amount: i * 100, code: SPEC124_CODES[(i - 1) % 3]!, formula: `of:=[.B${row}]*2`, value: i * 100 * 2 };
});

/** What spec124-structure.ods holds, for the unit and e2e specs. */
export const SPEC124_STRUCTURE_ODS = {
  title: "Structure",
  tableNames: ["Sheet1", "Sheet2"],
  activeTable: "Sheet1",
  activeSheetIndex: 0,
  frozen: { rows: 0, cols: 1 },
  printRanges: "Sheet1.A1:Sheet1.C11",
  rows: SPEC124_ODS_ROWS,
  headers: { A1: "Item", B1: "Amount", C1: "Code" },
  contentValidation: {
    name: "val1",
    condition: "of:cell-content-is-in-list([.$F$1:.$F$3])",
    baseCellAddress: "Sheet1.C2",
    cells: SPEC124_ODS_ROWS.map((row) => `C${row.row}`),
  },
  validationList: { ref: "F1:F3", values: ["A", "B", "C"] },
  frame: {
    ref: "E2",
    endCellAddress: "Sheet1.G8",
    notifyRanges: "Sheet1.A1:Sheet1.B11",
    href: "./Object 1",
    replacementHref: "./ObjectReplacements/Object 1",
  },
  merge: { ref: "A13:C13", value: "Merged note", covered: ["B13", "C13"] },
  conditionalFormat: { targetRangeAddress: "Sheet1.B2:Sheet1.B11", value: ">500", baseCellAddress: "Sheet1.B2", applyStyleName: "Accent" },
  sheet2: {
    formulas: {
      A1: { wire: "of:=[$Sheet1.B5]", formula: "Sheet1!B5", value: 400 },
      A2: { wire: "of:=SUM([$Sheet1.B2:.B11])", formula: "SUM(Sheet1!B2:B11)", value: 5500 },
    },
  },
  namedRange: { name: "Amounts", baseCellAddress: "$Sheet1.$B$2", cellRangeAddress: "$Sheet1.$B$2:.$B$11" },
  namedExpression: { name: "DoubleFirst", baseCellAddress: "$Sheet1.$A$1", expression: "of:=[$Sheet1.$B$2]*2" },
  databaseRange: { name: "__Anonymous_Sheet_DB__0", targetRangeAddress: "Sheet1.A1:Sheet1.C11", displayFilterButtons: true },
  chart: {
    plotAreaRange: "Sheet1.A1:Sheet1.B11",
    categoriesRange: "Sheet1.A2:Sheet1.A11",
    valuesRange: "Sheet1.B2:Sheet1.B11",
    labelCell: "Sheet1.B1",
    replacement: "ObjectReplacements/Object 1",
    objectParts: ["Object 1/content.xml", "Object 1/styles.xml", "Object 1/meta.xml"],
  },
  settingsTables: ["Sheet1", "Sheet2"],
  /** Every entry, in package order (mimetype first, STOREd). */
  parts: [
    "mimetype",
    "meta.xml",
    "settings.xml",
    "content.xml",
    "Thumbnails/thumbnail.png",
    "Object 1/content.xml",
    "Object 1/styles.xml",
    "Object 1/meta.xml",
    "ObjectReplacements/Object 1",
    "styles.xml",
    "manifest.rdf",
    "META-INF/manifest.xml",
  ],
} as const;

/** content.xml's automatic styles: columns, rows, the table, the bold
 * header, and the embedded chart frame's graphic and paragraph styles (gr1,
 * P1 — see SPEC121_ODS_AUTOMATIC_STYLES). */
const SPEC124_ODS_AUTOMATIC_STYLES =
  `<office:automatic-styles>` +
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
  `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
  `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center" ` +
  `draw:textarea-vertical-align="middle" draw:color-mode="standard" draw:luminance="0%" draw:contrast="0%" draw:gamma="100%" draw:red="0%" draw:green="0%" ` +
  `draw:blue="0%" fo:clip="rect(0in, 0in, 0in, 0in)" draw:image-opacity="100%" style:mirror="none"/></style:style>` +
  `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>` +
  `</office:automatic-styles>`;

/** styles.xml: Calc's common styles plus "Accent", the conditional
 * format's style (referenced only by name). */
const SPEC124_ODS_STYLES = LO_STYLES.replace(
  `</office:styles>`,
  `<style:style style:name="Accent" style:family="table-cell" style:parent-style-name="Default"/></office:styles>`,
);

/** content.xml's prefix: the list validation Sheet1!C2:C11 use. */
function spec124OdsContentValidations(): string {
  const V = SPEC124_STRUCTURE_ODS.contentValidation;
  return (
    `<table:content-validations><table:content-validation table:name="${V.name}" table:condition="${xmlEscape(V.condition)}" ` +
    `table:allow-empty-cell="true" table:base-cell-address="${V.baseCellAddress}" table:display-list="unsorted">` +
    `<table:error-message table:display="true" table:message-type="stop"/></table:content-validation></table:content-validations>`
  );
}

/** E2's chart frame, embedded directly in the cell (not `table:shapes`) —
 * see buildSpec121LinksOdsBytes's frame for the general shape. */
function spec124OdsFrame(): string {
  const F = SPEC124_STRUCTURE_ODS.frame;
  return (
    `<draw:frame draw:z-index="0" draw:name="Chart 1" draw:style-name="gr1" draw:text-style-name="P1" svg:width="8cm" svg:height="5cm" ` +
    `svg:x="0cm" svg:y="0cm" table:end-cell-address="${F.endCellAddress}" table:end-x="1cm" table:end-y="0.2cm">` +
    `<draw:object draw:notify-on-update-of-ranges="${F.notifyRanges}" xlink:href="${F.href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">` +
    `<loext:p/></draw:object>` +
    `<draw:image xlink:href="${F.replacementHref}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`
  );
}

/** Sheet1: the header, the 10 data rows (E2 the chart frame, C validated
 * from F1:F3, D the doubling formula), the empty row 12, the merged row 13,
 * the LibreOffice tail, then the table-local conditional format. */
function spec124OdsSheet1(): string {
  const F = SPEC124_STRUCTURE_ODS;
  const validatedCode = (code: string): string =>
    `<table:table-cell table:content-validation-name="${F.contentValidation.name}" office:value-type="string" calcext:value-type="string">` +
    `<text:p>${code}</text:p></table:table-cell>`;
  const listCell = (physicalRow: number): string => {
    const index = physicalRow - 1;
    return index >= 0 && index < F.validationList.values.length ? loString(F.validationList.values[index]!) : loEmpty(1);
  };

  const header = loRow(loString(F.headers.A1, "ce1") + loString(F.headers.B1, "ce1") + loString(F.headers.C1, "ce1") + loEmpty(2) + listCell(1) + loEmpty(16378));
  const dataRows = F.rows
    .map((row) => {
      const eCell = row.row === 2 ? `<table:table-cell>${spec124OdsFrame()}</table:table-cell>` : loEmpty(1);
      return loRow(
        loString(row.item) + loFloat(row.amount) + validatedCode(row.code) + loFormula(row.formula, row.value) + eCell + listCell(row.row) + loEmpty(16378),
      );
    })
    .join("");
  const row12 = loRow(loEmpty(16384));
  const M = F.merge;
  const row13 = loRow(
    `<table:table-cell table:number-columns-spanned="3" table:number-rows-spanned="1" office:value-type="string" calcext:value-type="string">` +
      `<text:p>${M.value}</text:p></table:table-cell><table:covered-table-cell/><table:covered-table-cell/>` +
      loEmpty(16381),
  );
  const tail = loRow(loEmpty(16384), "ro1", 1_048_576 - 13);
  const C = F.conditionalFormat;
  return (
    `<table:table table:name="Sheet1" table:style-name="ta1" table:print-ranges="${F.printRanges}">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    header +
    dataRows +
    row12 +
    row13 +
    tail +
    `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${C.targetRangeAddress}">` +
    `<calcext:condition calcext:apply-style-name="${C.applyStyleName}" calcext:value="${xmlEscape(C.value)}" calcext:base-cell-address="${C.baseCellAddress}"/>` +
    `</calcext:conditional-format></calcext:conditional-formats>` +
    `</table:table>`
  );
}

/** Sheet2: a single-cell cross-sheet reference and a SUM over Sheet1. */
function spec124OdsSheet2(): string {
  const F = SPEC124_STRUCTURE_ODS.sheet2;
  return (
    `<table:table table:name="Sheet2" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(loFormula(F.formulas.A1.wire, F.formulas.A1.value) + loEmpty(16383)) +
    loRow(loFormula(F.formulas.A2.wire, F.formulas.A2.value) + loEmpty(16383)) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 2) +
    `</table:table>`
  );
}

/** The chart object's content.xml: a bar chart over Sheet1, its plot area,
 * categories, series values and label naming the sheet (see
 * spec121OdsChartContent). */
function spec124OdsChartContent(): string {
  const C = SPEC124_STRUCTURE_ODS.chart;
  const rows = SPEC124_STRUCTURE_ODS.rows;
  const localRows = rows
    .map(
      (row) =>
        `<table:table-row><table:table-cell office:value-type="string"><text:p>${row.item}</text:p></table:table-cell>` +
        `<table:table-cell office:value-type="float" office:value="${row.amount}"><text:p>${row.amount}</text:p></table:table-cell></table:table-row>`,
    )
    .join("");
  return (
    ODF_XML_DECLARATION +
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
    `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
    `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
    `xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
    `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
    `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:chartooo="http://openoffice.org/2010/chart" office:version="1.3">` +
    `<office:automatic-styles><number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
    `<style:style style:name="ch1" style:family="chart"><style:graphic-properties draw:stroke="none" draw:fill-color="#ffffff"/></style:style>` +
    `<style:style style:name="ch2" style:family="chart"><style:chart-properties chart:include-hidden-cells="false" chart:auto-position="true" ` +
    `chart:auto-size="true" chart:treat-empty-cells="leave-gap" chart:right-angled-axes="true"/></style:style>` +
    `<style:style style:name="ch3" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:display-label="true" ` +
    `chart:logarithmic="false" chart:reverse-direction="false" text:line-break="false" chart:link-data-style-to-source="true" chart:axis-position="0" ` +
    `chart:axis-label-position="outside-start" chart:tick-mark-position="at-labels"/><style:graphic-properties svg:stroke-color="#b3b3b3"/>` +
    `<style:text-properties fo:font-size="10pt"/></style:style>` +
    `<style:style style:name="ch4" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:display-label="true" ` +
    `chart:logarithmic="false" chart:reverse-direction="false" chart:link-data-style-to-source="true" chart:axis-position="start"/>` +
    `<style:graphic-properties svg:stroke-color="#b3b3b3"/><style:text-properties fo:font-size="10pt"/></style:style>` +
    `<style:style style:name="ch5" style:family="chart"><style:graphic-properties svg:stroke-color="#b3b3b3"/></style:style>` +
    `<style:style style:name="ch6" style:family="chart" style:data-style-name="N0"><style:chart-properties chart:link-data-style-to-source="true"/>` +
    `<style:graphic-properties draw:fill-color="#004586"/></style:style>` +
    `<style:style style:name="ch7" style:family="chart"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>` +
    `<style:style style:name="ch8" style:family="chart"><style:graphic-properties svg:stroke-color="#b3b3b3" draw:fill="none"/></style:style>` +
    `</office:automatic-styles><office:body><office:chart>` +
    `<chart:chart svg:width="8cm" svg:height="5cm" xlink:href=".." xlink:type="simple" chart:class="chart:bar" chart:style-name="ch1">` +
    `<chart:plot-area chart:style-name="ch2" table:cell-range-address="${C.plotAreaRange}" chart:data-source-has-labels="both" ` +
    `svg:x="0.1in" svg:y="0.1in" svg:width="2.3in" svg:height="1.6in">` +
    `<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch3"><chart:categories table:cell-range-address="${C.categoriesRange}"/></chart:axis>` +
    `<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch4"><chart:grid chart:style-name="ch5" chart:class="major"/></chart:axis>` +
    `<chart:series chart:style-name="ch6" chart:values-cell-range-address="${C.valuesRange}" chart:label-cell-address="${C.labelCell}" ` +
    `chart:class="chart:bar"><chart:data-point chart:repeated="${rows.length}"/></chart:series>` +
    `<chart:wall chart:style-name="ch7"/><chart:floor chart:style-name="ch8"/></chart:plot-area>` +
    `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
    `<table:table-columns><table:table-column/></table:table-columns><table:table-header-rows><table:table-row>` +
    `<table:table-cell><text:p/></table:table-cell><table:table-cell office:value-type="string"><text:p>${SPEC124_STRUCTURE_ODS.headers.B1}</text:p></table:table-cell>` +
    `</table:table-row></table:table-header-rows><table:table-rows>${localRows}</table:table-rows></table:table>` +
    `</chart:chart></office:chart></office:body></office:document-content>`
  );
}

/** The chart's replacement image: a stand-in for the GDI metafile Calc
 * stores (its "VCLMTF" signature, then a few bytes — see
 * SPEC121_ODS_REPLACEMENT). */
const SPEC124_ODS_REPLACEMENT = new Uint8Array([0x56, 0x43, 0x4c, 0x4d, 0x54, 0x46, 0x01, 0x00, 0x31, 0x00, 0x00, 0x00, 0x53, 0x50, 0x45, 0x43, 0x2d, 0x31, 0x32, 0x34]);

/** META-INF/manifest.xml: Calc's entries plus the chart object's, in
 * Calc's order (see SPEC121_ODS_MANIFEST). */
const SPEC124_ODS_MANIFEST = LO_MANIFEST.replace(
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>`,
  `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/meta.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="Object 1/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
    `<manifest:file-entry manifest:full-path="ObjectReplacements/Object 1" ` +
    `manifest:media-type="application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;"/>`,
);

/** settings.xml: ActiveTable and one Tables entry per table — Sheet1's
 * carrying the frozen first column (see spec121OdsSettings). */
function spec124OdsSettings(): string {
  const O = SPEC124_STRUCTURE_ODS;
  const tables =
    spec117TableSettings(O.tableNames[0], {
      cursor: [0, 0],
      split: [2, 0],
      splitPosition: [1, 0],
      activeSplitRange: 3,
      position: { left: 0, right: 1, top: 0, bottom: 0 },
      zoom: 100,
      showGrid: true,
    }) +
    spec117TableSettings(O.tableNames[1], {
      cursor: [0, 0],
      split: [0, 0],
      splitPosition: [0, 0],
      activeSplitRange: 2,
      position: { left: 0, right: 0, top: 0, bottom: 0 },
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
    `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    loConfigItem("ViewId", "string", "view1") +
    `<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>` +
    loConfigItem("ActiveTable", "string", O.activeTable) +
    loConfigItem("HorizontalScrollbarWidth", "int", 1309) +
    loConfigItem("ZoomValue", "int", 100) +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>` +
    `<config:config-item-set config:name="ooo:configuration-settings">` +
    loConfigItem("IsDocumentShared", "boolean", false) +
    loConfigItem("SaveVersionOnClose", "boolean", false) +
    `</config:config-item-set></office:settings></office:document-settings>`
  );
}

export async function buildSpec124StructureOdsBytes(): Promise<Uint8Array> {
  const O = SPEC124_STRUCTURE_ODS;
  const N = O.namedRange;
  const E = O.namedExpression;
  const suffix =
    `<table:named-expressions><table:named-range table:name="${N.name}" table:base-cell-address="${N.baseCellAddress}" ` +
    `table:cell-range-address="${N.cellRangeAddress}"/><table:named-expression table:name="${E.name}" table:base-cell-address="${E.baseCellAddress}" ` +
    `table:expression="${xmlEscape(E.expression)}"/></table:named-expressions>` +
    `<table:database-ranges><table:database-range table:name="${O.databaseRange.name}" table:target-range-address="${O.databaseRange.targetRangeAddress}" ` +
    `table:display-filter-buttons="${O.databaseRange.displayFilterButtons}"/></table:database-ranges>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC124_ODS_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${spec124OdsContentValidations()}${spec124OdsSheet1()}${spec124OdsSheet2()}${suffix}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const cellCount = 3 + O.rows.length * 4 + 1 + 2;

  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    part("meta.xml", loMeta(O.title, O.tableNames.length, cellCount, 1)),
    part("settings.xml", spec124OdsSettings()),
    part("content.xml", content),
    { name: "Thumbnails/thumbnail.png", data: THUMBNAIL_PNG },
    part("Object 1/content.xml", spec124OdsChartContent()),
    part("Object 1/styles.xml", SPEC121_ODS_CHART_STYLES),
    part("Object 1/meta.xml", SPEC121_ODS_CHART_META),
    { name: O.chart.replacement, data: SPEC124_ODS_REPLACEMENT },
    part("styles.xml", SPEC124_ODS_STYLES),
    part("manifest.rdf", LO_MANIFEST_RDF),
    part("META-INF/manifest.xml", SPEC124_ODS_MANIFEST),
  ]);
}

// ============================================================================
// SPEC-125 — spec125-fill.xlsx / spec125-fill.ods: copy, paste, fill and move
// ============================================================================
// The workbook SPEC-125's e2e flow runs on, in both producers' shapes
// (Excel 365 and LibreOffice 24.8):
// - "Sheet1" (active): B2:C3 = 10, 20 / =B2*2, =B2+C2 — the block the flow
//   copies; E1:F1 merged, holding "Merged"; A10:A12 = 5, 6, 7 — the cells it
//   cuts; J2 "Mon" — the fill handle's seed; L2:L21 = 1…20 beside M2 =L2*2 —
//   the double-click fill; N2 "Total" above an empty N3:N6 — Ctrl/Cmd+D.
// - "Sheet2": A1 =SUM(Sheet1!A10:A12) and A2 =Sheet1!B2*3, the cross-sheet
//   formulas a cut re-points.
// - "Series": one seed column per series kind — its label in row 1, its
//   seeds below: numbers (one, a step, a trend, decimals), dates (by day,
//   month-ends, a month step, a year step, a week step), a time, weekday and
//   month names, quarters with and without a year, counters (plain, zero
//   padded, stepped), an ordinal, a formula, a boolean and plain text.
// Every cached value is what its formula computes. SPEC125_FILL records the
// facts the specs use.

/** A seed of spec125-fill's "Series" sheet. */
type Spec125Seed = number | string | boolean | { date: [number, number, number] } | { time: number } | { formula: string; value: number; odsWire?: string };

export const SPEC125_FILL = {
  title: "Fill",
  sheetNames: ["Sheet1", "Sheet2", "Series"],
  activeSheetIndex: 0,
  block: { B2: 10, C2: 20, B3: { formula: "B2*2", value: 20 }, C3: { formula: "B2+C2", value: 30 } },
  merge: { ref: "E1:F1", value: "Merged" },
  cut: { A10: 5, A11: 6, A12: 7 },
  monday: { ref: "J2", value: "Mon" },
  guide: { column: "L", first: 2, last: 21 },
  doubleClick: { ref: "M2", formula: "L2*2", value: 2 },
  fillDown: { ref: "N2", value: "Total" },
  sheet2: { A1: { formula: "SUM(Sheet1!A10:A12)", value: 18 }, A2: { formula: "Sheet1!B2*3", value: 30 } },
  series: [
    { label: "Number", seeds: [5] },
    { label: "Step", seeds: [1, 3] },
    { label: "Trend", seeds: [1, 2, 4] },
    { label: "Decimal", seeds: [0.1, 0.2] },
    { label: "Day", seeds: [{ date: [2026, 1, 31] }] },
    { label: "Month end", seeds: [{ date: [2026, 1, 31] }, { date: [2026, 2, 28] }] },
    { label: "Month", seeds: [{ date: [2026, 1, 15] }, { date: [2026, 2, 15] }] },
    { label: "Year", seeds: [{ date: [2024, 3, 1] }, { date: [2025, 3, 1] }] },
    { label: "Week", seeds: [{ date: [2026, 1, 1] }, { date: [2026, 1, 8] }] },
    { label: "Hour", seeds: [{ time: 9 }] },
    { label: "Weekday", seeds: ["Mon"] },
    { label: "Months", seeds: ["Jan", "Apr"] },
    { label: "Quarter", seeds: ["Q3 2026"] },
    { label: "Qtr", seeds: ["Qtr 4"] },
    { label: "Item", seeds: ["Item 9"] },
    { label: "Padded", seeds: ["Item 009"] },
    { label: "Odd items", seeds: ["Item 1", "Item 3"] },
    { label: "Ordinal", seeds: ["1st"] },
    { label: "Formula", seeds: [{ formula: "A2*2", value: 10 }] },
    { label: "Logical", seeds: [true] },
    { label: "Text", seeds: ["abc"] },
  ] as { label: string; seeds: Spec125Seed[] }[],
} as const;

const SPEC125_SHEET1_UID = "5A1C3E5F-7B9D-4F1A-8C2E-4A6C8E0A2C51";
const SPEC125_SHEET2_UID = "6B2D4F60-8CAE-4A2B-9D3F-5B7D9F1B3D52";
const SPEC125_SERIES_UID = "7C3E5071-9DBF-4B3C-AE40-6C8EA02C4E53";
const SPEC125_DATE_XF = 1;
const SPEC125_TIME_XF = 2;
const SPEC125_HEADER_XF = 3;

/** Excel's styles part with a date (14), a time (20, h:mm) and a bold xf. */
const SPEC125_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

/** The "Series" sheet's seeds as A1 → seed, its labels in row 1. */
function spec125SeriesGrid(): Map<string, Spec125Seed> {
  const grid = new Map<string, Spec125Seed>();
  SPEC125_FILL.series.forEach((column, c) => {
    const letters = String.fromCharCode(65 + c);
    grid.set(`${letters}1`, column.label);
    column.seeds.forEach((seed, r) => grid.set(`${letters}${r + 2}`, seed));
  });
  return grid;
}

/** Sheet1's cells as A1 → value or formula, row by row. */
function spec125Sheet1Cells(): Map<string, Spec125Seed> {
  const F = SPEC125_FILL;
  const cells = new Map<string, Spec125Seed>([
    ["A1", "Amounts"],
    ["B1", "Price"],
    ["C1", "Qty"],
    ["E1", F.merge.value],
    ["J1", "Days"],
    ["L1", "Count"],
    ["M1", "Double"],
    ["N1", "Label"],
    ["B2", F.block.B2],
    ["C2", F.block.C2],
    ["B3", F.block.B3],
    ["C3", F.block.C3],
    [F.monday.ref, F.monday.value],
    [F.doubleClick.ref, { formula: F.doubleClick.formula, value: F.doubleClick.value }],
    [F.fillDown.ref, F.fillDown.value],
    ["A10", F.cut.A10],
    ["A11", F.cut.A11],
    ["A12", F.cut.A12],
  ]);
  for (let r = F.guide.first; r <= F.guide.last; r += 1) {
    cells.set(`${F.guide.column}${r}`, r - F.guide.first + 1);
  }
  return cells;
}

/** Cells grouped by row, each row's in column order. */
function spec125Rows(cells: Map<string, Spec125Seed>): Map<number, [string, Spec125Seed][]> {
  const rows = new Map<number, [string, Spec125Seed][]>();
  for (const [ref, seed] of cells) {
    const r = Number(/\d+/.exec(ref)![0]);
    rows.set(r, [...(rows.get(r) ?? []), [ref, seed]]);
  }
  for (const entries of rows.values()) {
    entries.sort((a, b) => a[0].length - b[0].length || a[0].localeCompare(b[0]));
  }
  return new Map([...rows.entries()].toSorted((a, b) => a[0] - b[0]));
}

function spec125XlsxCell(ref: string, seed: Spec125Seed, str: (value: string) => number, header: boolean): string {
  if (typeof seed === "string") {
    return sharedCell(ref, str(seed), header ? SPEC125_HEADER_XF : undefined);
  }
  if (typeof seed === "number") {
    return numberCell(ref, seed);
  }
  if (typeof seed === "boolean") {
    return `<c r="${ref}" t="b"><v>${seed ? 1 : 0}</v></c>`;
  }
  if ("date" in seed) {
    return numberCell(ref, dateToSerial(seed.date[0], seed.date[1], seed.date[2], "1900"), SPEC125_DATE_XF);
  }
  if ("time" in seed) {
    return numberCell(ref, seed.time / 24, SPEC125_TIME_XF);
  }
  return formulaCell(ref, normalF(seed.formula), seed.value);
}

function spec125XlsxRows(cells: Map<string, Spec125Seed>, str: (value: string) => number, spans: string): string {
  return [...spec125Rows(cells)]
    .map(([r, entries]) => excelRow(r, spans, entries.map(([ref, seed]) => spec125XlsxCell(ref, seed, str, r === 1)).join("")))
    .join("");
}

export async function buildSpec125FillXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC125_FILL;
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  const str = (value: string): number => {
    let index = sstIndex.get(value);
    if (index === undefined) {
      index = sst.length;
      sst.push(value);
      sstIndex.set(value, index);
    }
    return index;
  };
  const sheet1 =
    XML_DECLARATION +
    worksheetOpen(SPEC125_SHEET1_UID) +
    `<dimension ref="A1:N21"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>` +
    `<sheetData>${spec125XlsxRows(spec125Sheet1Cells(), str, "1:14")}</sheetData>` +
    `<mergeCells count="1"><mergeCell ref="${F.merge.ref}"/></mergeCells>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  const sheet2 = worksheetXml({
    uid: SPEC125_SHEET2_UID,
    dimension: "A1:A2",
    sheetData:
      excelRow(1, "1:1", formulaCell("A1", normalF(F.sheet2.A1.formula), F.sheet2.A1.value)) +
      excelRow(2, "1:1", formulaCell("A2", normalF(F.sheet2.A2.formula), F.sheet2.A2.value)),
  });
  const series = worksheetXml({ uid: SPEC125_SERIES_UID, dimension: "A1:U4", sheetData: spec125XlsxRows(spec125SeriesGrid(), str, "1:21") });
  const worksheets = [sheet1, sheet2, series];
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${sst.length}">` +
    sst.map((text) => `<si><t>${xmlText(text)}</t></si>`).join("") +
    `</sst>`;
  const calcChain = XML_DECLARATION + `<calcChain xmlns="${MAIN_NS}"><c r="B3" i="1" l="1"/><c r="C3"/><c r="M2"/><c r="A1" i="2"/><c r="A2"/><c r="S2" i="3"/></calcChain>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheetNames.map((name, index) => ({ name, xml: worksheets[index]! })),
    sharedStrings,
    calcChain,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    styles: SPEC125_STYLES,
  });
}

// --- spec125-fill.ods ---------------------------------------------------------------

/** One seed as Calc writes its cell (dates with ce1's MM/DD/YYYY, times with
 * ce4's HH:MM:SS, headers bold with ce5). */
function spec125OdsCell(seed: Spec125Seed, header: boolean): string {
  if (typeof seed === "string") {
    return loString(xmlEscape(seed), header ? "ce5" : undefined);
  }
  if (typeof seed === "number") {
    return loFloat(seed);
  }
  if (typeof seed === "boolean") {
    return `<table:table-cell office:value-type="boolean" office:boolean-value="${seed}" calcext:value-type="boolean"><text:p>${seed ? "TRUE" : "FALSE"}</text:p></table:table-cell>`;
  }
  if ("date" in seed) {
    const [y, m, d] = seed.date;
    const iso = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    return `<table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="${iso}" calcext:value-type="date"><text:p>${String(m).padStart(2, "0")}/${String(d).padStart(2, "0")}/${y}</text:p></table:table-cell>`;
  }
  if ("time" in seed) {
    const hours = String(seed.time).padStart(2, "0");
    return `<table:table-cell table:style-name="ce4" office:value-type="time" office:time-value="PT${hours}H00M00S" calcext:value-type="time"><text:p>${hours}:00:00</text:p></table:table-cell>`;
  }
  // A same-sheet formula's references in OpenFormula's brackets.
  return loFormula(seed.odsWire ?? `of:=${seed.formula.replace(/\$?[A-Z]{1,3}\$?\d+/g, (ref) => `[.${ref}]`)}`, seed.value);
}

/** A table's rows: every row up to its last cell, empty runs between,
 * Calc's trailing run of empty rows. */
function spec125OdsTable(name: string, cells: Map<string, Spec125Seed>, extra = ""): string {
  const rows = spec125Rows(cells);
  const out: string[] = [];
  let next = 1;
  for (const [r, entries] of rows) {
    if (r > next) {
      out.push(loRow(loEmpty(16384), "ro1", r - next));
    }
    let col = 0;
    let row = "";
    for (const [ref, seed] of entries) {
      const at = ref.replace(/\d+/, "").split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      if (at > col) {
        row += loEmpty(at - col);
      }
      const merged = name === "Sheet1" && ref === "E1";
      row += merged
        ? `<table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="1" office:value-type="string" calcext:value-type="string"><text:p>${seed as string}</text:p></table:table-cell><table:covered-table-cell/>`
        : spec125OdsCell(seed, r === 1);
      col = at + (merged ? 2 : 1);
    }
    out.push(loRow(row + loEmpty(16384 - col)));
    next = r + 1;
  }
  out.push(loRow(loEmpty(16384), "ro1", 1_048_576 - (next - 1)));
  return (
    `<table:table table:name="${name}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    out.join("") +
    extra +
    `</table:table>`
  );
}

export async function buildSpec125FillOdsBytes(): Promise<Uint8Array> {
  const F = SPEC125_FILL;
  const sheet2 = new Map<string, Spec125Seed>([
    ["A1", { ...F.sheet2.A1, odsWire: "of:=SUM([$Sheet1.A10:.A12])" }],
    ["A2", { ...F.sheet2.A2, odsWire: "of:=[$Sheet1.B2]*3" }],
  ]);
  const tables =
    spec125OdsTable("Sheet1", spec125Sheet1Cells()) +
    spec125OdsTable("Sheet2", sheet2) +
    spec125OdsTable("Series", spec125SeriesGrid());
  const cellCount = spec125Sheet1Cells().size + sheet2.size + spec125SeriesGrid().size;
  return loPackage(loContent(tables), loMeta(F.title, F.sheetNames.length, cellCount), loSettings(F.sheetNames, "Sheet1"));
}

// --- SPEC-126: CSV and TSV --------------------------------------------------------------------
// Six delimited-text files, each emulating a producer. They are LITERAL
// bytes — spelled out here character by character and encoded by the small
// tables below, never by csv/write.ts, which would only prove our writer
// agrees with itself:
// - spec126-excel-utf8.csv — Excel "CSV UTF-8": a BOM, CRLF after every
//   record, minimal quoting (a field with a comma, one with an in-cell LF,
//   one with `""`), `10/8/2026`, `12%`, `1.50`, `TRUE`, `007`, a 16-digit
//   number, `1-2`, `SEPT2`, `2026-10-08`, `123456789012`, empty fields;
//   rectangular (five fields a record).
// - spec126-semicolon-1252.csv — a German Excel export: Windows-1252, `;`,
//   CRLF, header `Name;Zip;Amount`, row 2 `Müller;01234;1234,5`, a `€` and
//   `1.234,56` elsewhere and a `31.12.2025` date.
// - spec126-quoted.csv — a database dump: UTF-8 without a BOM, LF, every
//   non-empty field quoted, ragged records, no line ending after the last.
// - spec126-sep.csv — Excel's `sep=;` first line, then three `;` records.
// - spec126-unicode.tsv — Excel "Unicode Text": a UTF-16 LE BOM, tab, CRLF,
//   a quoted field holding a tab and a `""`, Czech, Greek and an
//   astral-plane emoji.
// - spec126-mac-pipe.csv — classic Mac: lone CR records, `|`, Windows-1252.
// SPEC126_CSV records the facts.

/** The few non-ASCII Windows-1252 bytes the fixtures spell. */
const WINDOWS_1252_BYTES: Readonly<Record<string, number>> = {
  "€": 0x80,
  "Ü": 0xdc,
  "ß": 0xdf,
  "è": 0xe8,
  "é": 0xe9,
  "ö": 0xf6,
  "û": 0xfb,
  "ü": 0xfc,
};

/** `text` as Windows-1252 bytes: ASCII as itself, the table above for the
 * rest (a character outside both is a fixture bug). */
function windows1252Literal(text: string): Uint8Array {
  return Uint8Array.from(Array.from(text), (char) => {
    const code = char.charCodeAt(0);
    const byte = code < 0x80 ? code : WINDOWS_1252_BYTES[char];
    if (byte === undefined) {
      throw new Error(`no Windows-1252 byte spelled for ${char}`);
    }
    return byte;
  });
}

/** `text` as UTF-16 LE code units after the FF FE BOM. */
function utf16LeLiteral(text: string): Uint8Array {
  const out = new Uint8Array(2 + text.length * 2);
  out[0] = 0xff;
  out[1] = 0xfe;
  for (let at = 0; at < text.length; at += 1) {
    const unit = text.charCodeAt(at);
    out[2 + at * 2] = unit & 0xff;
    out[3 + at * 2] = unit >> 8;
  }
  return out;
}

function withUtf8Bom(text: string): Uint8Array {
  const body = encoder.encode(text);
  const out = new Uint8Array(3 + body.length);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}

export const SPEC126_CSV = {
  excelUtf8: {
    fileName: "spec126-excel-utf8.csv",
    text:
      "Name,Note,Date,Share,Price\r\n" +
      '"Smith, Jane","Line one\nLine two",10/8/2026,12%,1.50\r\n' +
      'Bob,"Says ""hi""",,TRUE,007\r\n' +
      "Ann,1234567890123456,1-2,SEPT2,\r\n" +
      "Zoë,Plain text,2026-10-08,,123456789012\r\n",
    /** The one-field edit the round trip makes: A3's text, replaced by
     * exactly the typed text. */
    edit: { ref: "A3", row: 2, col: 0, from: "Bob", to: "Robert" },
  },
  semicolon1252: {
    fileName: "spec126-semicolon-1252.csv",
    text:
      "Name;Zip;Amount\r\n" +
      "Müller;01234;1234,5\r\n" +
      "Schröder;80331;1.234,56\r\n" +
      "Café Ü (€);10115;-7,25\r\n" +
      "Stichtag;;31.12.2025\r\n",
    edit: { ref: "C2", row: 1, col: 2, from: "1234,5", to: "99" },
  },
  quoted: {
    fileName: "spec126-quoted.csv",
    text: '"id","name","city"\n"1","Ada","London"\n"2","Grace"\n"3","Linus","Helsinki","FI"\n"4","","Oslo"',
  },
  sep: {
    fileName: "spec126-sep.csv",
    text: "sep=;\r\nItem;Qty;Price\r\nApples;3;0.5\r\nPears;12;1.25\r\n",
  },
  unicode: {
    fileName: "spec126-unicode.tsv",
    text:
      "Jazyk\tPozdrav\tPoznámka\r\n" +
      "Čeština\tDobrý den\tŽluťoučký kůň\r\n" +
      'Ελληνικά\tΓειά σου\t"tab\there ""quoted"""\r\n' +
      "Emoji\t😀\t42\r\n",
    edit: { ref: "B2", row: 1, col: 1, from: "Dobrý den", to: "Ahoj" },
  },
  macPipe: {
    fileName: "spec126-mac-pipe.csv",
    text: "Produkt|Preis|Menge\rCrème brûlée|4.50|2\rCafé|2.20|10\r",
  },
} as const;

export function buildSpec126ExcelUtf8CsvBytes(): Uint8Array {
  return withUtf8Bom(SPEC126_CSV.excelUtf8.text);
}

export function buildSpec126Semicolon1252CsvBytes(): Uint8Array {
  return windows1252Literal(SPEC126_CSV.semicolon1252.text);
}

export function buildSpec126QuotedCsvBytes(): Uint8Array {
  return encoder.encode(SPEC126_CSV.quoted.text);
}

export function buildSpec126SepCsvBytes(): Uint8Array {
  return encoder.encode(SPEC126_CSV.sep.text);
}

export function buildSpec126UnicodeTsvBytes(): Uint8Array {
  return utf16LeLiteral(SPEC126_CSV.unicode.text);
}

export function buildSpec126MacPipeCsvBytes(): Uint8Array {
  return windows1252Literal(SPEC126_CSV.macPipe.text);
}

// ============================================================================
// SPEC-128 — spec128-filter.xlsx / spec128-filter.ods: sort and autofilter
// ============================================================================
// The workbooks SPEC-128's readers, writers and e2e flow run on, each in its
// producer's own shapes:
// - spec128-filter.xlsx — Excel 365 on the classic Office theme (Calibri 11,
//   rows `x14ac:dyDescent="0.25"`). "Sales" (active): Region, Rep, Amount,
//   Date (bold headers, dates on built-in numFmt 14) over 12 records in Date
//   order; `<sheetPr filterMode="1"/>` before `<dimension>`; an `autoFilter`
//   over A1:D13 with its `xr:uid` — Region a `filters blank="1"` list (East,
//   North, West: both South rows fail), Amount a `customFilters and="1"`
//   pair (>= 100, <= 900: the 50 and the 950 fail) — and Excel's
//   `autoFilter/sortState` (ascending by Date, `xmlns:xlrd2` on it, as Excel
//   writes it); those four rows `hidden="1"`, plus row 12, which every
//   criterion passes — the row the user hid. Rows 2, 4, 6, 8, 10 and 12 are
//   written WITHOUT `r` (their cells too), as a non-Excel producer writes
//   them; each directly follows its predecessor, so sequential placement
//   lands it on its own row. "Colours": an autoFilter whose `colorFilter
//   dxfId="0"` (styles.xml's one dxf, a solid yellow fill) hides the two
//   unfilled rows. "Log": a worksheet-level `sortState` whose only condition
//   is `sortBy="cellColor"` (no autoFilter, nothing hidden). xl/workbook.xml
//   carries a hidden `_xlnm._FilterDatabase` per autofilter.
// - spec128-filter.ods — LibreOffice 24.8's twin of "Sales": rows 4, 5, 9
//   and 10 `table:visibility="filter"`, row 12 `"collapse"`, dates as
//   `office:value-type="date"` under a date data style, the trailing
//   1048563-row run, and after `table:named-expressions` the anonymous
//   sheet database range `__Anonymous_Sheet_DB__0` with
//   `table:display-filter-buttons="true"` and its `table:filter` —
//   `table:filter-set-item`s for Region (the empty item is "(Blanks)") and a
//   `>=`/`<=` pair with `table:data-type="number"` for Amount.
// SPEC128_FILTER_XLSX and SPEC128_FILTER_ODS record the facts: rows are
// 1-based sheet rows, `index`/`indices` 0-based, `range`s 0-based GridRanges.

/** What the file does with a "Sales" record: shows it, hides it because a
 * criterion rejects it, or hides it although every criterion passes (the
 * user hid it). */
type Spec128RowState = "shown" | "filter" | "user";

interface Spec128SalesRow {
  /** 1-based sheet row. */
  row: number;
  /** null: no A cell at all — a blank, which `filters blank="1"` passes. */
  region: string | null;
  rep: string;
  amount: number;
  date: readonly [number, number, number];
  /** The date as a 1900-system serial. */
  serial: number;
  state: Spec128RowState;
}

const spec128Row = (
  row: number,
  region: string | null,
  rep: string,
  amount: number,
  date: readonly [number, number, number],
  state: Spec128RowState,
): Spec128SalesRow => ({ row, region, rep, amount, date, serial: dateToSerial(date[0], date[1], date[2], "1900"), state });

/** "Sales"'s 12 records, in Date order (the sort record's order). */
const SPEC128_SALES_ROWS: readonly Spec128SalesRow[] = [
  spec128Row(2, "East", "Ana", 420, [2026, 1, 5], "shown"),
  spec128Row(3, "West", "Ben", 310, [2026, 1, 12], "shown"),
  spec128Row(4, "South", "Cy", 275, [2026, 1, 19], "filter"),
  spec128Row(5, "North", "Dee", 50, [2026, 2, 2], "filter"),
  spec128Row(6, "East", "Eve", 610, [2026, 2, 9], "shown"),
  spec128Row(7, null, "Fay", 180, [2026, 2, 16], "shown"),
  spec128Row(8, "West", "Gus", 730, [2026, 3, 2], "shown"),
  spec128Row(9, "South", "Hal", 505, [2026, 3, 9], "filter"),
  spec128Row(10, "East", "Ivy", 950, [2026, 3, 16], "filter"),
  spec128Row(11, "North", "Jo", 390, [2026, 3, 23], "shown"),
  spec128Row(12, "East", "Kim", 220, [2026, 4, 6], "user"),
  spec128Row(13, "North", "Lee", 845, [2026, 4, 13], "shown"),
];

/** "Colours"'s items: `yellow` cells carry the fill the colour filter keeps. */
const SPEC128_COLOUR_ROWS = [
  { row: 2, item: "Apples", qty: 12, yellow: true },
  { row: 3, item: "Pears", qty: 7, yellow: false },
  { row: 4, item: "Plums", qty: 30, yellow: true },
  { row: 5, item: "Figs", qty: 4, yellow: false },
  { row: 6, item: "Limes", qty: 18, yellow: true },
] as const;

/** "Log"'s entries, already sorted yellow-first as its cell-colour sort left them. */
const SPEC128_LOG_ROWS = [
  { row: 2, entry: "Audit", hours: 3, yellow: true },
  { row: 3, entry: "Backup", hours: 1, yellow: true },
  { row: 4, entry: "Cleanup", hours: 2, yellow: false },
  { row: 5, entry: "Deploy", hours: 5, yellow: false },
  { row: 6, entry: "Review", hours: 4, yellow: false },
] as const;

const SPEC128_XLRD2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
const SPEC128_SALES_AUTOFILTER_UID = "{4D5E6F70-0128-4182-BDCE-3F4A5B6C7D84}";
const SPEC128_COLOURS_AUTOFILTER_UID = "{5E6F7081-0128-4293-8EDF-4A5B6C7D8E95}";

/** "Sales"'s `<autoFilter>`, verbatim — what the reader keeps as `sourceXml`. */
const SPEC128_SALES_AUTOFILTER_XML =
  `<autoFilter ref="A1:D13" xr:uid="${SPEC128_SALES_AUTOFILTER_UID}">` +
  `<filterColumn colId="0"><filters blank="1"><filter val="East"/><filter val="North"/><filter val="West"/></filters></filterColumn>` +
  `<filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="100"/>` +
  `<customFilter operator="lessThanOrEqual" val="900"/></customFilters></filterColumn>` +
  `<sortState xmlns:xlrd2="${SPEC128_XLRD2_NS}" ref="A2:D13"><sortCondition ref="D2:D13"/></sortState></autoFilter>`;

/** "Colours"'s `<autoFilter>`, verbatim. */
const SPEC128_COLOURS_AUTOFILTER_XML =
  `<autoFilter ref="A1:B6" xr:uid="${SPEC128_COLOURS_AUTOFILTER_UID}"><filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn></autoFilter>`;

/** "Log"'s worksheet-level `<sortState>`, verbatim. */
const SPEC128_LOG_SORTSTATE_XML =
  `<sortState xmlns:xlrd2="${SPEC128_XLRD2_NS}" ref="A2:B6"><sortCondition sortBy="cellColor" ref="A2:A6" dxfId="0"/></sortState>`;

/** What spec128-filter.xlsx holds, for the unit, round-trip and e2e specs. */
export const SPEC128_FILTER_XLSX = {
  title: "Filter",
  sheetNames: ["Sales", "Colours", "Log"],
  activeSheetIndex: 0,
  /** cellXfs indices: 1 numFmt 14 (m/d/yyyy), 2 bold, 3 solid yellow fill. */
  xf: { date: 1, header: 2, yellow: 3 },
  /** styles.xml's one dxf: the solid yellow fill both the colour filter and
   * the cell-colour sort name. */
  dxf: { id: 0, fill: "#ffff00" },
  sales: {
    partName: "xl/worksheets/sheet1.xml",
    dimension: "A1:D13",
    headers: { A1: "Region", B1: "Rep", C1: "Amount", D1: "Date" },
    rows: SPEC128_SALES_ROWS,
    /** `<row>`s written without `r` (their cells without `r` too). */
    rowsWithoutR: [2, 4, 6, 8, 10, 12],
    /** `<sheetPr filterMode="1"/>`. */
    filterMode: true,
    autoFilter: {
      ref: "A1:D13",
      range: { top: 0, left: 0, bottom: 12, right: 3 },
      uid: SPEC128_SALES_AUTOFILTER_UID,
      xml: SPEC128_SALES_AUTOFILTER_XML,
      /** Region: `filters blank="1"` with three values. */
      region: { colId: 0, values: ["East", "North", "West"], blanks: true },
      /** Amount: `customFilters and="1"`, 100 <= amount <= 900. */
      amount: {
        colId: 2,
        join: "and",
        conditions: [
          { operator: "greaterThanOrEqual", value: "100" },
          { operator: "lessThanOrEqual", value: "900" },
        ],
      },
      /** `autoFilter/sortState`: the data rows, ascending by Date. */
      sort: { ref: "A2:D13", range: { top: 1, left: 0, bottom: 12, right: 3 }, conditionRef: "D2:D13", keys: [{ col: 3, descending: false }] },
    },
    /** Rows a criterion rejects (South ×2, 50, 950) — `filterHidden` on read. */
    filterHiddenRows: [4, 5, 9, 10],
    filterHiddenIndices: [3, 4, 8, 9],
    /** Row 12 (East, 220) passes every criterion yet is hidden — `hidden` on read. */
    userHiddenRow: 12,
    userHiddenIndex: 11,
    /** Data rows of the range (M) and those not filter-hidden (N). */
    records: 12,
    status: "8 of 12 records found",
    /** After unchecking "West" in Region's popover (rows 3 and 8 join the filtered). */
    statusWithoutWest: "6 of 12 records found",
    /** The Amount column's visible cells, top to bottom, as opened. */
    visibleAmounts: [420, 310, 610, 180, 730, 390, 845],
    /** …after Clear filters (only the user-hidden row 12 stays hidden). */
    visibleAmountsCleared: [420, 310, 275, 50, 610, 180, 730, 505, 950, 390, 845],
    /** …after Sort Z to A on Amount (row 12's 220 neither moves nor is moved into). */
    visibleAmountsDescending: [950, 845, 730, 610, 505, 420, 390, 310, 275, 180, 50],
  },
  colours: {
    partName: "xl/worksheets/sheet2.xml",
    dimension: "A1:B6",
    headers: { A1: "Item", B1: "Qty" },
    rows: SPEC128_COLOUR_ROWS,
    filterMode: true,
    autoFilter: {
      ref: "A1:B6",
      range: { top: 0, left: 0, bottom: 5, right: 1 },
      uid: SPEC128_COLOURS_AUTOFILTER_UID,
      xml: SPEC128_COLOURS_AUTOFILTER_XML,
      /** A `colorFilter` — a `preserved` criterion labelled "colour". */
      colorFilter: { colId: 0, dxfId: 0, xml: `<colorFilter dxfId="0"/>` },
    },
    /** The unfilled rows the colour filter hides — `filterHidden` on read
     * (a preserved criterion keeps the file's state). */
    filterHiddenRows: [3, 5],
    filterHiddenIndices: [2, 4],
  },
  log: {
    partName: "xl/worksheets/sheet3.xml",
    dimension: "A1:B6",
    headers: { A1: "Entry", B1: "Hours" },
    rows: SPEC128_LOG_ROWS,
    /** The worksheet `<sortState>`: its one condition sorts by cell colour,
     * so it is not a key — the element is kept as `sourceXml`. */
    sortState: {
      ref: "A2:B6",
      range: { top: 1, left: 0, bottom: 5, right: 1 },
      xml: SPEC128_LOG_SORTSTATE_XML,
      condition: { sortBy: "cellColor", ref: "A2:A6", dxfId: 0 },
    },
  },
  /** xl/workbook.xml's definedNames, in file order. */
  definedNames: [
    { name: "_xlnm._FilterDatabase", localSheetId: 0, hidden: true, text: "Sales!$A$1:$D$13" },
    { name: "_xlnm._FilterDatabase", localSheetId: 1, hidden: true, text: "Colours!$A$1:$B$6" },
  ],
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

const SPEC128_SALES_UID = "1A2B3C4D-0128-4E5F-8A9B-0C1D2E3F4A51";
const SPEC128_COLOURS_UID = "2B3C4D5E-0128-4F60-9BAC-1D2E3F4A5B62";
const SPEC128_LOG_UID = "3C4D5E6F-0128-4071-ACBD-2E3F4A5B6C73";

/** Excel's styles part on the classic theme: Normal, a date xf (numFmt 14),
 * a bold header xf and a solid-yellow xf; the one dxf the colour filter and
 * the cell-colour sort name (Excel's shape for a fill dxf it records from a
 * filter or sort: a solid pattern, the colour in `bgColor`). */
const SPEC128_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor indexed="64"/><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

/** A Calibri-era `<row>`: `r` unless omitted, `spans`, `hidden`, `x14ac:dyDescent`. */
function spec128XlsxRow(r: number | undefined, spans: string, hidden: boolean, cells: string): string {
  return `<row${r === undefined ? "" : ` r="${r}"`} spans="${spans}"${hidden ? ` hidden="1"` : ""} x14ac:dyDescent="0.25">${cells}</row>`;
}

/** A cell: `r` unless its row omits it, then `s`, `t` and the value. */
function spec128XlsxCell(ref: string | undefined, value: number, options: { s?: number; t?: "s" } = {}): string {
  return `<c${ref === undefined ? "" : ` r="${ref}"`}${options.s ? ` s="${options.s}"` : ""}${options.t ? ` t="${options.t}"` : ""}><v>${value}</v></c>`;
}

/** A worksheet in Excel 365's CT_Worksheet order: sheetPr (filterMode),
 * dimension, sheetViews, sheetFormatPr, sheetData, autoFilter or sortState,
 * pageMargins. */
function spec128WorksheetXml(options: { uid: string; filterMode: boolean; dimension: string; selected: boolean; sheetData: string; after: string }): string {
  return (
    XML_DECLARATION +
    worksheetOpen(options.uid) +
    (options.filterMode ? `<sheetPr filterMode="1"/>` : "") +
    `<dimension ref="${options.dimension}"/>` +
    `<sheetViews><sheetView ${options.selected ? 'tabSelected="1" ' : ""}workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<sheetData>${options.sheetData}</sheetData>` +
    options.after +
    SPEC117_PAGE_MARGINS +
    `</worksheet>`
  );
}

export async function buildSpec128FilterXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC128_FILTER_XLSX;
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  const str = (value: string): number => {
    let index = sstIndex.get(value);
    if (index === undefined) {
      index = sst.length;
      sst.push(value);
      sstIndex.set(value, index);
    }
    return index;
  };
  const XF = F.xf;
  const headerRow = (headers: Record<string, string>, spans: string): string =>
    spec128XlsxRow(1, spans, false, Object.entries(headers).map(([ref, text]) => spec128XlsxCell(ref, str(text), { s: XF.header, t: "s" })).join(""));

  // Interned in Excel's first-use order: row by row, sheet by sheet.
  const salesHeader = headerRow(F.sales.headers, "1:4");
  const withoutR: readonly number[] = F.sales.rowsWithoutR;
  const salesRows = F.sales.rows.map((row) => {
    const bare = withoutR.includes(row.row);
    const ref = (col: string): string | undefined => (bare ? undefined : `${col}${row.row}`);
    const cells =
      (row.region === null ? "" : spec128XlsxCell(ref("A"), str(row.region), { t: "s" })) +
      spec128XlsxCell(ref("B"), str(row.rep), { t: "s" }) +
      spec128XlsxCell(ref("C"), row.amount) +
      spec128XlsxCell(ref("D"), row.serial, { s: XF.date });
    return spec128XlsxRow(bare ? undefined : row.row, "1:4", row.state !== "shown", cells);
  });
  const sales = spec128WorksheetXml({
    uid: SPEC128_SALES_UID,
    filterMode: F.sales.filterMode,
    dimension: F.sales.dimension,
    selected: true,
    sheetData: salesHeader + salesRows.join(""),
    after: F.sales.autoFilter.xml,
  });

  const colours = spec128WorksheetXml({
    uid: SPEC128_COLOURS_UID,
    filterMode: F.colours.filterMode,
    dimension: F.colours.dimension,
    selected: false,
    sheetData:
      headerRow(F.colours.headers, "1:2") +
      F.colours.rows
        .map((row) =>
          spec128XlsxRow(
            row.row,
            "1:2",
            !row.yellow,
            spec128XlsxCell(`A${row.row}`, str(row.item), { s: row.yellow ? XF.yellow : undefined, t: "s" }) + spec128XlsxCell(`B${row.row}`, row.qty),
          ),
        )
        .join(""),
    after: F.colours.autoFilter.xml,
  });

  const log = spec128WorksheetXml({
    uid: SPEC128_LOG_UID,
    filterMode: false,
    dimension: F.log.dimension,
    selected: false,
    sheetData:
      headerRow(F.log.headers, "1:2") +
      F.log.rows
        .map((row) =>
          spec128XlsxRow(
            row.row,
            "1:2",
            false,
            spec128XlsxCell(`A${row.row}`, str(row.entry), { s: row.yellow ? XF.yellow : undefined, t: "s" }) + spec128XlsxCell(`B${row.row}`, row.hours),
          ),
        )
        .join(""),
    after: F.log.sortState.xml,
  });

  const worksheets = [sales, colours, log];
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${sst.length}">` +
    sst.map((text) => `<si><t>${xmlText(text)}</t></si>`).join("") +
    `</sst>`;
  const definedNames =
    `<definedNames>` +
    F.definedNames
      .map((name) => `<definedName name="${name.name}" localSheetId="${name.localSheetId}"${name.hidden ? ` hidden="1"` : ""}>${name.text}</definedName>`)
      .join("") +
    `</definedNames>`;

  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheetNames.map((name, index) => ({ name, xml: worksheets[index]! })),
    sharedStrings,
    workbookPr: `<workbookPr defaultThemeVersion="166925"/>`,
    theme: SPEC117_THEME,
    styles: SPEC128_STYLES,
    additions: { parts: [], relationships: [], overrides: [], workbookAfterSheets: definedNames },
  });
}

// --- spec128-filter.ods ---------------------------------------------------------------

/** The anonymous sheet database range carrying "Sales"'s autofilter, verbatim
 * as Calc writes it — what the reader keeps as `sourceXml`. */
const SPEC128_ODS_DATABASE_RANGE_XML =
  `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sales.A1:Sales.D13" table:display-filter-buttons="true">` +
  `<table:filter><table:filter-and>` +
  `<table:filter-condition table:field-number="0" table:value="East" table:operator="=">` +
  `<table:filter-set-item table:value=""/><table:filter-set-item table:value="East"/>` +
  `<table:filter-set-item table:value="North"/><table:filter-set-item table:value="West"/></table:filter-condition>` +
  `<table:filter-condition table:field-number="2" table:value="100" table:operator="&gt;=" table:data-type="number"/>` +
  `<table:filter-condition table:field-number="2" table:value="900" table:operator="&lt;=" table:data-type="number"/>` +
  `</table:filter-and></table:filter></table:database-range>`;

/** What spec128-filter.ods holds, for the unit, round-trip and e2e specs. */
export const SPEC128_FILTER_ODS = {
  title: "Filter",
  tableNames: ["Sales"],
  activeTable: "Sales",
  activeSheetIndex: 0,
  headers: { A1: "Region", B1: "Rep", C1: "Amount", D1: "Date" },
  /** The same records as SPEC128_FILTER_XLSX.sales.rows. */
  rows: SPEC128_SALES_ROWS,
  /** Header cells bold (ce5); dates in ce1 — N37, MM/DD/YY. */
  styles: { header: "ce5", date: "ce1", dateDataStyle: "N37" },
  databaseRange: {
    name: "__Anonymous_Sheet_DB__0",
    targetRangeAddress: "Sales.A1:Sales.D13",
    displayFilterButtons: true,
    range: { top: 0, left: 0, bottom: 12, right: 3 },
    xml: SPEC128_ODS_DATABASE_RANGE_XML,
    /** Region: `=` with set items; the empty item is "(Blanks)". */
    region: { fieldNumber: 0, setItems: ["", "East", "North", "West"], values: ["East", "North", "West"], blanks: true },
    /** Amount: a numeric `>=`/`<=` pair (custom, and). */
    amount: {
      fieldNumber: 2,
      conditions: [
        { operator: ">=", value: "100" },
        { operator: "<=", value: "900" },
      ],
    },
  },
  /** `table:visibility="filter"` rows — `filterHidden` on read. */
  filterRows: [4, 5, 9, 10],
  filterIndices: [3, 4, 8, 9],
  /** The `table:visibility="collapse"` row — `hidden` on read. */
  collapseRow: 12,
  collapseIndex: 11,
  /** 13 rows written, then Calc's trailing run to the grid's edge. */
  trailingRowRun: 1_048_576 - 13,
  records: 12,
  status: "8 of 12 records found",
} as const;

/** Two digits, as Calc's ISO dates and its MM/DD/YY text spell them. */
const spec128Pad2 = (n: number): string => String(n).padStart(2, "0");

/** A Calc row with its visibility: `filter` (a filter hid it) or `collapse`
 * (the user did), in Calc's attribute order. */
function spec128OdsRow(cells: string, visibility?: "filter" | "collapse"): string {
  return `<table:table-row table:style-name="ro1"${visibility ? ` table:visibility="${visibility}"` : ""}>${cells}</table:table-row>`;
}

export async function buildSpec128FilterOdsBytes(): Promise<Uint8Array> {
  const O = SPEC128_FILTER_ODS;
  const header = spec128OdsRow(Object.values(O.headers).map((text) => loString(text, O.styles.header)).join("") + loEmpty(16380));
  const rows = O.rows.map((row) => {
    const [y, m, d] = row.date;
    const cells =
      (row.region === null ? `<table:table-cell/>` : loString(row.region)) +
      loString(row.rep) +
      loFloat(row.amount) +
      `<table:table-cell table:style-name="${O.styles.date}" office:value-type="date" office:date-value="${y}-${spec128Pad2(m)}-${spec128Pad2(d)}" calcext:value-type="date">` +
      `<text:p>${spec128Pad2(m)}/${spec128Pad2(d)}/${spec128Pad2(y % 100)}</text:p></table:table-cell>` +
      loEmpty(16380);
    return spec128OdsRow(cells, row.state === "filter" ? "filter" : row.state === "user" ? "collapse" : undefined);
  });
  const table =
    `<table:table table:name="${O.tableNames[0]}" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    header +
    rows.join("") +
    loRow(loEmpty(16384), "ro1", O.trailingRowRun) +
    `</table:table>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${LO_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${table}<table:named-expressions/>` +
    `<table:database-ranges>${O.databaseRange.xml}</table:database-ranges>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const cellCount = Object.keys(O.headers).length + O.rows.reduce((sum, row) => sum + (row.region === null ? 3 : 4), 0);
  return loPackage(content, loMeta(O.title, O.tableNames.length, cellCount), loSettings(O.tableNames, O.activeTable));
}

// ============================================================================
// SPEC-129 (xlsx) — spec129-cf.xlsx: conditional formatting
// ============================================================================
// Excel 365 on the classic Office theme (Calibri 11), its worksheets rooted
// with `mc:Ignorable="x14ac xr xr2 xr3"`, every `<conditionalFormatting>`
// carrying an `xr:uid`, and styles.xml's three `<dxfs>` written as Excel
// writes a dxf fill — a `patternFill` with no patternType and only `bgColor`:
// dxf 0 Excel's "Light red fill with dark red text", dxf 1 its "Yellow fill
// with dark yellow text", dxf 2 bold theme text on a tinted theme fill.
// "Scores" (active), 12 records under header row 1, rules by priority:
//   1. B2:B13 cellIs greaterThan 100 (dxf 0) — B3 (140) is the only score
//      above 100; B4 is 90.
//   2. C2:C13 a solid data bar with its x14 twin (one GUID, gradient="0",
//      axis automatic, autoMin/autoMax): distinct values, C2 the clearly
//      negative one, C5 slightly negative.
//   3. D2:D13 3Arrows, thresholds percent 0/33/67 with gte="0"; D2 the
//      column's largest value (so it shows the green up arrow, index 2).
//   4. E2:E13 a 3-colour scale (min, 50th percentile, max) of tinted theme
//      colours.
//   5. F2:F13 an x14-only 3Stars icon set (no base twin).
//   6. G2:G13 a custom x14 icon set (`custom="1"` + `x14:cfIcon`s).
//   7. A2:H13 expression `$H2="late"` with stopIfTrue (dxf 1) — rows 3 and 4
//      are "on time"; ranked last, so it hides no bar, icon or scale.
// Its `<extLst>` holds the conditional-formatting ext and, after it, a
// sparkline ext that must stay verbatim.
// "More": top10 percent (rank 20), aboveAverage stdDev="1", one block with
// duplicateValues and containsText "an" (Excel's formula), containsBlanks
// and timePeriod last7Days, each with Excel's own `<formula>`.
// SPEC129_CF_XLSX records the facts (rows 1-based, ranges 0-based).

const SPEC129_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SPEC129_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
const SPEC129_DATABAR_GUID = "{8A1E2F3C-0129-4D5E-9F60-7A8B9C0D1E2F}";
const SPEC129_STARS_GUID = "{9B2F3041-0129-4E6F-8071-8B9CAD1E2F30}";
const SPEC129_CUSTOM_GUID = "{AC304152-0129-4F70-9182-9CADBE2F3041}";

/** The x14 data bar twin of "Scores"' C column, verbatim. */
const SPEC129_DATABAR_X14 =
  `<x14:cfRule type="dataBar" id="${SPEC129_DATABAR_GUID}"><x14:dataBar minLength="0" maxLength="100" gradient="0">` +
  `<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>`;

/** The x14-only 3Stars rule, verbatim. */
const SPEC129_STARS_X14 =
  `<x14:cfRule type="iconSet" priority="5" id="${SPEC129_STARS_GUID}"><x14:iconSet iconSet="3Stars">` +
  `<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>` +
  `</x14:iconSet></x14:cfRule>`;

/** The custom icon set — preserved verbatim by the reader. */
const SPEC129_CUSTOM_X14 =
  `<x14:cfRule type="iconSet" priority="6" id="${SPEC129_CUSTOM_GUID}"><x14:iconSet iconSet="3Arrows" custom="1">` +
  `<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>` +
  `<x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="3Symbols" iconId="1"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule>`;

/** The sparkline ext after the conditional formattings, verbatim. */
const SPEC129_SPARKLINE_EXT =
  `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${SPEC129_X14_NS}"><x14:sparklineGroups xmlns:xm="${SPEC129_XM_NS}">` +
  `<x14:sparklineGroup displayEmptyCellsAs="gap" xr2:uid="{BD415263-0129-4081-A293-ADBECF304152}"><x14:colorSeries theme="4" tint="-0.499984740745262"/>` +
  `<x14:colorNegative theme="5"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers theme="4" tint="-0.499984740745262"/><x14:colorFirst theme="4" tint="0.39997558519241921"/>` +
  `<x14:colorLast theme="4" tint="0.39997558519241921"/><x14:colorHigh theme="4"/><x14:colorLow theme="4"/>` +
  `<x14:sparklines><x14:sparkline><xm:f>Scores!B2:G2</xm:f><xm:sqref>I2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`;

/** "Scores"' base blocks, verbatim, in document order. */
const SPEC129_SCORES_BLOCKS = [
  `<conditionalFormatting sqref="B2:B13" xr:uid="{C1526374-0129-4192-B3A4-BECFD0415263}"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="C2:C13" xr:uid="{D2637485-0129-42A3-84B5-CFD0E1526374}"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>` +
    `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${SPEC129_X14_NS}"><x14:id>${SPEC129_DATABAR_GUID}</x14:id></ext></extLst></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="D2:D13" xr:uid="{E3748596-0129-43B4-95C6-D0E1F2637485}"><cfRule type="iconSet" priority="3"><iconSet iconSet="3Arrows">` +
    `<cfvo type="percent" val="0"/><cfvo type="percent" val="33" gte="0"/><cfvo type="percent" val="67" gte="0"/></iconSet></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="E2:E13" xr:uid="{F48596A7-0129-44C5-A6D7-E1F203748596}"><cfRule type="colorScale" priority="4"><colorScale>` +
    `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color theme="5" tint="0.39997558519241921"/><color theme="0" tint="-0.249977111117893"/>` +
    `<color theme="9"/></colorScale></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="A2:H13" xr:uid="{0596A7B8-0129-45D6-B7E8-F20314859607}"><cfRule type="expression" dxfId="1" priority="7" stopIfTrue="1"><formula>$H2="late"</formula></cfRule></conditionalFormatting>`,
] as const;

/** "Scores"' conditional-formatting ext (the twin, the stars, the custom set). */
const SPEC129_CF_EXT =
  `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${SPEC129_X14_NS}"><x14:conditionalFormattings>` +
  `<x14:conditionalFormatting xmlns:xm="${SPEC129_XM_NS}">${SPEC129_DATABAR_X14}<xm:sqref>C2:C13</xm:sqref></x14:conditionalFormatting>` +
  `<x14:conditionalFormatting xmlns:xm="${SPEC129_XM_NS}">${SPEC129_STARS_X14}<xm:sqref>F2:F13</xm:sqref></x14:conditionalFormatting>` +
  `<x14:conditionalFormatting xmlns:xm="${SPEC129_XM_NS}">${SPEC129_CUSTOM_X14}<xm:sqref>G2:G13</xm:sqref></x14:conditionalFormatting>` +
  `</x14:conditionalFormattings></ext>`;

/** "More"'s base blocks, verbatim. */
const SPEC129_MORE_BLOCKS = [
  `<conditionalFormatting sqref="A2:A11" xr:uid="{16A7B8C9-0129-46E7-88F9-0314259607A8}"><cfRule type="top10" dxfId="2" priority="1" percent="1" rank="20"/></conditionalFormatting>`,
  `<conditionalFormatting sqref="B2:B11" xr:uid="{27B8C9DA-0129-47F8-990A-14253607A8B9}"><cfRule type="aboveAverage" dxfId="1" priority="2" stdDev="1"/></conditionalFormatting>`,
  `<conditionalFormatting sqref="C2:C11" xr:uid="{38C9DAEB-0129-4809-AA1B-2536478B9CA0}"><cfRule type="duplicateValues" dxfId="0" priority="3"/>` +
    `<cfRule type="containsText" dxfId="2" priority="4" operator="containsText" text="an"><formula>NOT(ISERROR(SEARCH("an",C2)))</formula></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="D2:D11" xr:uid="{49DAEBFC-0129-491A-BB2C-36478A9CB0D1}"><cfRule type="containsBlanks" dxfId="1" priority="5"><formula>LEN(TRIM(D2))=0</formula></cfRule></conditionalFormatting>`,
  `<conditionalFormatting sqref="E2:E11" xr:uid="{5AEBFC0D-0129-4A2B-8C3D-4789BADC1E20}"><cfRule type="timePeriod" dxfId="0" priority="6" timePeriod="last7Days">` +
    `<formula>AND(TODAY()-FLOOR(E2,1)&lt;=6,FLOOR(E2,1)&lt;=TODAY())</formula></cfRule></conditionalFormatting>`,
] as const;

/** What spec129-cf.xlsx holds, for the unit, round-trip and e2e specs. */
export const SPEC129_CF_XLSX = {
  title: "Conditional formats",
  sheetNames: ["Scores", "More"],
  activeSheetIndex: 0,
  /** cellXfs: 1 a date (numFmt 14), 2 a bold header. */
  xf: { date: 1, header: 2 },
  /** styles.xml's dxfs as the model reads them. */
  dxfs: [
    { font: { color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" } },
    { font: { color: "#9c5700" }, fill: { kind: "solid", color: "#ffeb9c" } },
    { font: { bold: true, color: { theme: "dk1" } }, fill: { kind: "solid", color: { theme: "accent1", lumMod: 1 - 0.79998168889431442, lumOff: 0.79998168889431442, xlsxTint: "0.79998168889431442" } } },
  ],
  dxfXml: [
    `<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>`,
    `<dxf><font><color rgb="FF9C5700"/></font><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>`,
    `<dxf><font><b/><color theme="1"/></font><fill><patternFill><bgColor theme="4" tint="0.79998168889431442"/></patternFill></fill></dxf>`,
  ],
  scores: {
    partName: "xl/worksheets/sheet1.xml",
    dimension: "A1:H13",
    headers: ["Name", "Score", "Change", "Trend", "Heat", "Rating", "Flags", "Status"],
    names: ["Ana", "Ben", "Cy", "Dee", "Eve", "Fay", "Gus", "Hal", "Ivy", "Jo", "Kim", "Lee"],
    /** B2..B13: only B3 is above 100. */
    scores: [75, 140, 90, 60, 88, 95, 70, 82, 100, 55, 99, 64],
    /** C2..C13: distinct; C2 clearly negative, C5 slightly. */
    changes: [-30, 45, 10, -8, 80, 25, 60, 15, 35, 70, 50, 5],
    negativeCells: ["C2", "C5"],
    /** D2..D13: D2 the largest. */
    trends: [98, 12, 45, 67, 23, 88, 34, 51, 5, 76, 40, 60],
    heat: [55, 12, 78, 34, 90, 23, 67, 45, 81, 5, 62, 39],
    ratings: [3, 5, 1, 4, 2, 5, 3, 1, 4, 2, 3, 5],
    flags: [7, 2, 9, 4, 6, 1, 8, 3, 5, 10, 2, 6],
    /** H2..H13: rows 3 and 4 are on time. */
    statuses: ["late", "on time", "on time", "late", "on time", "on time", "late", "on time", "on time", "late", "on time", "on time"],
    blocks: SPEC129_SCORES_BLOCKS,
    cfExt: SPEC129_CF_EXT,
    sparklineExt: SPEC129_SPARKLINE_EXT,
    cellIs: { sqref: "B2:B13", priority: 1, dxfId: 0, operator: "greaterThan", formula: "100", hot: "B3", hotValue: 140, coolCell: "B4", coolValue: 90 },
    dataBar: { sqref: "C2:C13", priority: 2, guid: SPEC129_DATABAR_GUID, color: "#638ec6", x14Xml: SPEC129_DATABAR_X14 },
    arrows: { sqref: "D2:D13", priority: 3, set: "3Arrows", topCell: "D2", topIcon: "3Arrows:2" },
    colorScale: { sqref: "E2:E13", priority: 4 },
    stars: { sqref: "F2:F13", priority: 5, guid: SPEC129_STARS_GUID },
    custom: { sqref: "G2:G13", priority: 6, guid: SPEC129_CUSTOM_GUID, xml: SPEC129_CUSTOM_X14 },
    late: { sqref: "A2:H13", priority: 7, dxfId: 1, formula: `$H2="late"` },
  },
  more: {
    partName: "xl/worksheets/sheet2.xml",
    dimension: "A1:E11",
    headers: ["Sales", "Hours", "Fruit", "Notes", "Date"],
    sales: [120, 340, 560, 90, 780, 230, 410, 650, 300, 870],
    hours: [8, 7.5, 9, 6, 30, 8, 7, 8.5, 9, 7],
    fruit: ["Apple", "Banana", "Cherry", "Apple", "Mango", "Kiwi", "Banana", "Grape", "Orange", "Plum"],
    /** null: a blank cell. */
    notes: ["ok", null, "check", null, "ok", "ok", null, "late", "ok", null],
    /** E2..E11 as [y, m, d]. */
    dates: [
      [2026, 10, 1],
      [2026, 10, 3],
      [2026, 10, 5],
      [2026, 10, 6],
      [2026, 10, 7],
      [2026, 10, 8],
      [2026, 10, 9],
      [2026, 10, 10],
      [2026, 10, 11],
      [2026, 10, 12],
    ],
    blocks: SPEC129_MORE_BLOCKS,
  },
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

/** Excel's styles part for spec129-cf.xlsx (see the section header). */
const SPEC129_STYLES =
  XML_DECLARATION +
  `<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" ` +
  `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="3">${SPEC129_CF_XLSX.dxfXml.join("")}</dxfs>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

const SPEC129_SCORES_UID = "6BFC0D1E-0129-4B3C-9D4E-589ACBED2F31";
const SPEC129_MORE_UID = "7C0D1E2F-0129-4C4D-AE5F-69ABDCFE3042";

/** A worksheet in Excel 365's CT_Worksheet order with its conditional
 * formats (and extLst) after sheetData. */
function spec129WorksheetXml(options: { uid: string; dimension: string; selected: boolean; sheetData: string; after: string }): string {
  return (
    XML_DECLARATION +
    worksheetOpen(options.uid) +
    `<dimension ref="${options.dimension}"/>` +
    `<sheetViews><sheetView ${options.selected ? 'tabSelected="1" ' : ""}workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    SPEC117_FORMAT_PR +
    `<sheetData>${options.sheetData}</sheetData>` +
    options.after +
    `</worksheet>`
  );
}

/** A row of spec129-cf.xlsx as Excel 365 writes one. */
function spec129Row(r: number, spans: string, cells: string): string {
  return `<row r="${r}" spans="${spans}" x14ac:dyDescent="0.25">${cells}</row>`;
}

/** A number cell of spec129-cf.xlsx. */
function spec129Number(ref: string, value: number, style?: number): string {
  return `<c r="${ref}"${style ? ` s="${style}"` : ""}><v>${value}</v></c>`;
}

export async function buildSpec129CfXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC129_CF_XLSX;
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  const str = (value: string): number => {
    let index = sstIndex.get(value);
    if (index === undefined) {
      index = sst.length;
      sst.push(value);
      sstIndex.set(value, index);
    }
    return index;
  };
  const columns = "ABCDEFGH";
  const text = (ref: string, value: string, style?: number): string => `<c r="${ref}"${style ? ` s="${style}"` : ""} t="s"><v>${str(value)}</v></c>`;

  const S = F.scores;
  const scoresRows = [spec129Row(1, "1:8", S.headers.map((header, col) => text(`${columns[col]}1`, header, F.xf.header)).join(""))];
  S.names.forEach((name, i) => {
    const r = i + 2;
    scoresRows.push(
      spec129Row(
        r,
        "1:8",
        text(`A${r}`, name) +
          spec129Number(`B${r}`, S.scores[i]!) +
          spec129Number(`C${r}`, S.changes[i]!) +
          spec129Number(`D${r}`, S.trends[i]!) +
          spec129Number(`E${r}`, S.heat[i]!) +
          spec129Number(`F${r}`, S.ratings[i]!) +
          spec129Number(`G${r}`, S.flags[i]!) +
          text(`H${r}`, S.statuses[i]!),
      ),
    );
  });
  const scores = spec129WorksheetXml({
    uid: SPEC129_SCORES_UID,
    dimension: S.dimension,
    selected: true,
    sheetData: scoresRows.join(""),
    after: S.blocks.join("") + SPEC117_PAGE_MARGINS + `<extLst>${S.cfExt}${S.sparklineExt}</extLst>`,
  });

  const M = F.more;
  const moreRows = [spec129Row(1, "1:5", M.headers.map((header, col) => text(`${columns[col]}1`, header, F.xf.header)).join(""))];
  M.sales.forEach((sales, i) => {
    const r = i + 2;
    const note = M.notes[i];
    const [y, m, d] = M.dates[i]!;
    moreRows.push(
      spec129Row(
        r,
        "1:5",
        spec129Number(`A${r}`, sales) +
          spec129Number(`B${r}`, M.hours[i]!) +
          text(`C${r}`, M.fruit[i]!) +
          (note === null || note === undefined ? "" : text(`D${r}`, note)) +
          spec129Number(`E${r}`, dateToSerial(y, m, d, "1900"), F.xf.date),
      ),
    );
  });
  const more = spec129WorksheetXml({
    uid: SPEC129_MORE_UID,
    dimension: M.dimension,
    selected: false,
    sheetData: moreRows.join(""),
    after: M.blocks.join("") + SPEC117_PAGE_MARGINS,
  });

  const worksheets = [scores, more];
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${sst.length}">` +
    sst.map((value) => `<si><t>${xmlText(value)}</t></si>`).join("") +
    `</sst>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheetNames.map((name, index) => ({ name, xml: worksheets[index]! })),
    sharedStrings,
    workbookPr: `<workbookPr defaultThemeVersion="166925"/>`,
    theme: SPEC117_THEME,
    styles: SPEC129_STYLES,
  });
}
// ============================================================================
// end of SPEC-129 (xlsx)
// ============================================================================

// ============================================================================
// SPEC-129 (ODS) — spec129-cf.ods / spec129-legacy.ods: conditional formatting
// ============================================================================
// - spec129-cf.ods — LibreOffice 24.8's twin of spec129-cf.xlsx: "Scores"
//   (Name, Score, Change, Trend, Heat, Stars, Custom, Status over 12
//   records; B3 = 140, B4 = 90, D2 the Trend column's largest, negative
//   Changes) and "More", each table closing with Calc's
//   `calcext:conditional-formats` after its rows. The conditions apply the
//   common styles of styles.xml: "Bad" (fill #ffc7ce, text #9c0006) and
//   "Good" (#c6efce, #006100), both children of the empty "Status", and
//   "Accent 3" written `Accent_20_3` (fill #dddddd) under "Accent" (bold) —
//   so its differential style is its own fill plus Accent's weight, short of
//   "Default". Scores: B2:B13 `>100` → Bad; A2:H13 `formula-is([.$H2]="late")`
//   → Accent 3 (rows 3 and 4 are not late); a data bar on C2:C13 (negative
//   and axis colours, no gradient, auto minimum and maximum); a 3Arrows icon
//   set on D2:D13 whose 33 and 67 thresholds are `calcext:gte="false"`; a
//   3-colour scale on E2:E13 (minimum, 50th percentile, maximum); 3Stars on
//   F2:F13; a custom icon set on G2:G13 (`calcext:custom="true"`, carried
//   verbatim). More: one format on A2:A7 holding top-percent(10), above-
//   average and duplicate; contains-text("late") and the blanks formula on
//   B2:B7; date-is last-7-days on C2:C7.
// - spec129-legacy.ods — an older producer's shape: no calcext rules, only
//   `style:map` on content.xml's automatic cell styles. ce6 (`cell-content()
//   >100` → Bad, then `cell-content-is-between(50,100)` → Good, written from
//   Legacy.B2) styles B2:D6 cell by cell and then a 994-row repeated run of
//   empty cells, B7:D1000; ce7 (`is-true-formula([.E2]="late")` → Bad,
//   written from Legacy.E2) is column E's default cell style, which every
//   unstyled cell of the column takes — the header, the flags and Calc's
//   trailing run — so its region is E1:E1048576, one rectangle.
// SPEC129_CF_ODS and SPEC129_LEGACY_ODS record the facts; ranges are
// 0-based GridRanges, formulas canonical and relative to each format's
// origin.

/** styles.xml for both: Calc's common styles plus the conditional ones. */
const SPEC129_ODS_STYLES = LO_STYLES.replace(
  `</office:styles>`,
  `<style:style style:name="Accent" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" ` +
    `style:font-weight-asian="bold" style:font-weight-complex="bold"/></style:style>` +
    `<style:style style:name="Accent_20_3" style:display-name="Accent 3" style:family="table-cell" style:parent-style-name="Accent">` +
    `<style:table-cell-properties fo:background-color="#dddddd"/></style:style>` +
    `<style:style style:name="Status" style:family="table-cell" style:parent-style-name="Default"/>` +
    `<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Status"><style:table-cell-properties fo:background-color="#c6efce"/>` +
    `<style:text-properties fo:color="#006100"/></style:style>` +
    `<style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Status"><style:table-cell-properties fo:background-color="#ffc7ce"/>` +
    `<style:text-properties fo:color="#9c0006"/></style:style></office:styles>`,
);

/** One "Scores" record: Name, Score, Change, Trend, Heat, Stars, Custom, Status. */
type Spec129Record = readonly [string, number, number, number, number, number, number, string];

const SPEC129_SCORES: readonly Spec129Record[] = [
  ["Ana", 95, 12, 98, 10, 1, 3, "late"],
  ["Ben", 140, -8, 40, 20, 2, 1, "on time"],
  ["Cy", 90, 5, 55, 30, 3, 2, "on time"],
  ["Dee", 105, -15, 60, 40, 4, 3, "late"],
  ["Eve", 70, 20, 25, 50, 5, 1, "on time"],
  ["Fay", 130, 0, 80, 60, 1, 2, "on time"],
  ["Gus", 60, 9, 30, 70, 2, 3, "late"],
  ["Hal", 115, -3, 45, 80, 3, 1, "on time"],
  ["Ivy", 85, 14, 70, 90, 4, 2, "on time"],
  ["Jo", 100, 7, 15, 100, 5, 3, "on time"],
  ["Kai", 125, -11, 65, 110, 1, 1, "late"],
  ["Lu", 75, 3, 20, 120, 2, 2, "on time"],
];

/** Each Scores and More format verbatim, as Calc writes it (and as the
 * reader keeps it as `sourceXml`). */
const SPEC129_ODS_SCORES_FORMATS = [
  `<calcext:conditional-format calcext:target-range-address="Scores.B2:Scores.B13">` +
    `<calcext:condition calcext:apply-style-name="Bad" calcext:value="&gt;100" calcext:base-cell-address="Scores.B2"/></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.A2:Scores.H13">` +
    `<calcext:condition calcext:apply-style-name="Accent_20_3" calcext:value="formula-is([.$H2]=&quot;late&quot;)" calcext:base-cell-address="Scores.A2"/></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.C2:Scores.C13">` +
    `<calcext:data-bar calcext:gradient="false" calcext:negative-color="#ff0000" calcext:positive-color="#638ec6" calcext:axis-color="#000000">` +
    `<calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/><calcext:formatting-entry calcext:value="0" calcext:type="auto-maximum"/>` +
    `</calcext:data-bar></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.D2:Scores.D13">` +
    `<calcext:icon-set calcext:icon-set-type="3Arrows"><calcext:formatting-entry calcext:value="0" calcext:type="percent"/>` +
    `<calcext:formatting-entry calcext:value="33" calcext:type="percent" calcext:gte="false"/>` +
    `<calcext:formatting-entry calcext:value="67" calcext:type="percent" calcext:gte="false"/></calcext:icon-set></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.E2:Scores.E13">` +
    `<calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#f8696b"/>` +
    `<calcext:color-scale-entry calcext:value="50" calcext:type="percentile" calcext:color="#ffeb84"/>` +
    `<calcext:color-scale-entry calcext:value="0" calcext:type="maximum" calcext:color="#63be7b"/></calcext:color-scale></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.F2:Scores.F13">` +
    `<calcext:icon-set calcext:icon-set-type="3Stars"><calcext:formatting-entry calcext:value="0" calcext:type="percent"/>` +
    `<calcext:formatting-entry calcext:value="33" calcext:type="percent"/><calcext:formatting-entry calcext:value="67" calcext:type="percent"/>` +
    `</calcext:icon-set></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="Scores.G2:Scores.G13">` +
    `<calcext:icon-set calcext:icon-set-type="3Arrows" calcext:custom="true">` +
    `<calcext:custom-iconset calcext:custom-iconset-name="3Stars" calcext:custom-iconset-index="0"/>` +
    `<calcext:custom-iconset calcext:custom-iconset-name="3Flags" calcext:custom-iconset-index="1"/>` +
    `<calcext:custom-iconset calcext:custom-iconset-name="3Arrows" calcext:custom-iconset-index="2"/>` +
    `<calcext:formatting-entry calcext:value="0" calcext:type="percent"/><calcext:formatting-entry calcext:value="33" calcext:type="percent"/>` +
    `<calcext:formatting-entry calcext:value="67" calcext:type="percent"/></calcext:icon-set></calcext:conditional-format>`,
] as const;

const SPEC129_ODS_MORE_FORMATS = [
  `<calcext:conditional-format calcext:target-range-address="More.A2:More.A7">` +
    `<calcext:condition calcext:apply-style-name="Good" calcext:value="top-percent(10)" calcext:base-cell-address="More.A2"/>` +
    `<calcext:condition calcext:apply-style-name="Accent_20_3" calcext:value="above-average" calcext:base-cell-address="More.A2"/>` +
    `<calcext:condition calcext:apply-style-name="Bad" calcext:value="duplicate" calcext:base-cell-address="More.A2"/></calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="More.B2:More.B7">` +
    `<calcext:condition calcext:apply-style-name="Bad" calcext:value="contains-text(&quot;late&quot;)" calcext:base-cell-address="More.B2"/>` +
    `<calcext:condition calcext:apply-style-name="Accent_20_3" calcext:value="formula-is(LEN(TRIM([.B2]))=0)" calcext:base-cell-address="More.B2"/>` +
    `</calcext:conditional-format>`,
  `<calcext:conditional-format calcext:target-range-address="More.C2:More.C7">` +
    `<calcext:date-is calcext:style="Good" calcext:date="last-7-days"/></calcext:conditional-format>`,
] as const;

/** What spec129-cf.ods holds, for the unit, round-trip and e2e specs. */
export const SPEC129_CF_ODS = {
  title: "Conditional formats",
  tableNames: ["Scores", "More"],
  activeTable: "Scores",
  headers: ["Name", "Score", "Change", "Trend", "Heat", "Stars", "Custom", "Status"],
  records: SPEC129_SCORES,
  /** The common styles the rules apply. */
  styles: {
    bad: { name: "Bad", fill: "#ffc7ce", text: "#9c0006" },
    good: { name: "Good", fill: "#c6efce", text: "#006100" },
    accent3: { name: "Accent_20_3", displayName: "Accent 3", fill: "#dddddd", bold: true },
  },
  /** The dxfs they read as, in first-reference order. */
  dxfs: [
    { font: { color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" }, sourceStyleName: "Bad" },
    { font: { bold: true }, fill: { kind: "solid", color: "#dddddd" }, sourceStyleName: "Accent_20_3" },
    { font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" }, sourceStyleName: "Good" },
  ],
  scoresFormats: SPEC129_ODS_SCORES_FORMATS,
  moreFormats: SPEC129_ODS_MORE_FORMATS,
  /** More's values: numbers, notes (B3 and B6 blank) and due dates. */
  more: {
    values: [10, 50, 50, 90, 30, 70],
    notes: ["late fee", null, "ok", "late", null, "fine"],
    dates: ["2026-10-05", "2026-10-06", "2026-10-07", "2026-10-08", "2026-10-09", "2026-10-10"],
  },
} as const;

/** A Calc date cell (LO_AUTOMATIC_STYLES's ce1, the N37 date style). */
const spec129Date = (iso: string): string =>
  `<table:table-cell table:style-name="ce1" office:value-type="date" office:date-value="${iso}" calcext:value-type="date">` +
  `<text:p>${iso.slice(5, 7)}/${iso.slice(8, 10)}/${iso.slice(2, 4)}</text:p></table:table-cell>`;

export async function buildSpec129CfOdsBytes(): Promise<Uint8Array> {
  const O = SPEC129_CF_ODS;
  const header = loRow(O.headers.map((text) => loString(text, "ce5")).join("") + loEmpty(16384 - O.headers.length));
  const records = O.records.map(([name, score, change, trend, heat, stars, custom, status]) =>
    loRow(loString(name) + [score, change, trend, heat, stars, custom].map((value) => loFloat(value)).join("") + loString(status) + loEmpty(16376)),
  );
  const scores =
    `<table:table table:name="Scores" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    header +
    records.join("") +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 13) +
    `<calcext:conditional-formats>${O.scoresFormats.join("")}</calcext:conditional-formats></table:table>`;
  const moreRows = O.more.values.map((value, i) => {
    const note = O.more.notes[i];
    return loRow(loFloat(value) + (note === null || note === undefined ? `<table:table-cell/>` : loString(note)) + spec129Date(O.more.dates[i]!) + loEmpty(16381));
  });
  const more =
    `<table:table table:name="More" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(["Value", "Note", "Due"].map((text) => loString(text, "ce5")).join("") + loEmpty(16381)) +
    moreRows.join("") +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 7) +
    `<calcext:conditional-formats>${O.moreFormats.join("")}</calcext:conditional-formats></table:table>`;
  const cellCount = O.headers.length * 13 + 3 + 16;
  return loPackage(loContent(scores + more), loMeta(O.title, O.tableNames.length, cellCount), loSettings(O.tableNames, O.activeTable), SPEC129_ODS_STYLES);
}

/** content.xml's automatic styles for spec129-legacy.ods: Calc's, plus the
 * two mapped cell styles. */
const SPEC129_LEGACY_AUTOMATIC_STYLES = LO_AUTOMATIC_STYLES.replace(
  `</office:automatic-styles>`,
  `<style:style style:name="ce6" style:family="table-cell" style:parent-style-name="Default">` +
    `<style:map style:condition="cell-content()&gt;100" style:apply-style-name="Bad" style:base-cell-address="Legacy.B2"/>` +
    `<style:map style:condition="cell-content-is-between(50,100)" style:apply-style-name="Good" style:base-cell-address="Legacy.B2"/></style:style>` +
    `<style:style style:name="ce7" style:family="table-cell" style:parent-style-name="Default">` +
    `<style:map style:condition="is-true-formula([.E2]=&quot;late&quot;)" style:apply-style-name="Bad" style:base-cell-address="Legacy.E2"/></style:style>` +
    `</office:automatic-styles>`,
);

/** What spec129-legacy.ods holds. */
export const SPEC129_LEGACY_ODS = {
  title: "Legacy conditions",
  tableNames: ["Legacy"],
  activeTable: "Legacy",
  /** Rows 2–6: an item, three quarters (in ce6) and a flag (column E's ce7). */
  rows: [
    ["North", 120, 80, 40, "late"],
    ["South", 95, 101, 50, "ok"],
    ["East", 60, 30, 100, "late"],
    ["West", 150, 75, 20, "ok"],
    ["Centre", 99, 100, 101, "ok"],
  ],
  /** ce6's repeated empty run after the data: rows 7–1000. */
  emptyRun: 994,
  styles: { quarters: "ce6", flags: "ce7" },
  /** The formats the reader makes, in first-use order (E1 is met first). */
  formats: [
    { ranges: [{ top: 0, left: 4, bottom: 1_048_575, right: 4 }], rules: [{ kind: "expression", formula: `E1="late"`, style: "Bad" }] },
    {
      ranges: [{ top: 1, left: 1, bottom: 999, right: 3 }],
      rules: [
        { kind: "cellIs", operator: "greaterThan", formulas: ["100"], style: "Bad" },
        { kind: "cellIs", operator: "between", formulas: ["50", "100"], style: "Good" },
      ],
    },
  ],
} as const;

export async function buildSpec129LegacyOdsBytes(): Promise<Uint8Array> {
  const O = SPEC129_LEGACY_ODS;
  const styled = (cell: string): string => cell.replace("<table:table-cell", `<table:table-cell table:style-name="${O.styles.quarters}"`);
  const header = loRow(["Item", "Q1", "Q2", "Q3"].map((text) => loString(text, "ce5")).join("") + loString("Flag") + loEmpty(16379));
  const rows = O.rows.map(([item, q1, q2, q3, flag]) => loRow(loString(item) + [q1, q2, q3].map((value) => styled(loFloat(value))).join("") + loString(flag) + loEmpty(16379)));
  const table =
    `<table:table table:name="Legacy" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="4" table:default-cell-style-name="Default"/>` +
    `<table:table-column table:style-name="co1" table:default-cell-style-name="${O.styles.flags}"/>` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16379" table:default-cell-style-name="Default"/>` +
    header +
    rows.join("") +
    loRow(`<table:table-cell/><table:table-cell table:style-name="${O.styles.quarters}" table:number-columns-repeated="3"/>` + loEmpty(16380), "ro1", O.emptyRun) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - 6 - O.emptyRun) +
    `</table:table>`;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${SPEC129_LEGACY_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}${table}<table:named-expressions/>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return loPackage(content, loMeta(O.title, 1, 5 * 6), loSettings(O.tableNames, O.activeTable), SPEC129_ODS_STYLES);
}

// ============================================================================
// SPEC-130 (xlsx) — spec130-validation.xlsx: data validation and in-cell
// dropdowns, as Excel 365 writes them
// ============================================================================
//
// Excel's shapes: the worksheet root with `mc:Ignorable="x14ac xr xr2 xr3"`,
// every `<dataValidation>` carrying an `xr:uid` (attributes in the schema's
// order, as Excel writes them), one with `imeMode="noControl"`, a list
// hiding its arrow with `showDropDown="1"`, and the cross-sheet list Excel
// 2010+ moves into the worksheet extLst's `{CCE6A557-…}` ext as an
// `x14:dataValidation` (`xm:f`, `xm:sqref`, the ext declaring `xmlns:x14`,
// the list `xmlns:xm`). A workbook-scoped `definedName` "Colours" names
// Lists!$B$1:$B$3 (docProps/app.xml lists it under Named Ranges).
// "Form" (active): labels in A2:A10, B3 holds 5 and the rest of B2:B10 is
// empty, E2:E6 the five options B8 lists, D2 empty and under no rule:
//   B2  list "Low,Medium,High", input message Priority / Pick one, stop alert
//   B3  whole between 1 and 10, warning alert
//   B4  decimal greater than 0, information alert, imeMode="noControl"
//   B5  date between 2026-01-01 and 2026-12-31 (1900 serials)
//   B6  text length <= 5
//   B7  custom ISNUMBER(SEARCH("@",B7))
//   B8  list $E$2:$E$6 with showDropDown="1" (no arrow)
//   B9  list Lists!$A$1:$A$4 — the x14 ext
//   B10 list Colours (the defined name)
// "Lists": A1:A4 the regions B9 lists, B1:B3 the colours "Colours" names.
// SPEC130_VALIDATION_XLSX records the facts.

const SPEC130_X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SPEC130_XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
const SPEC130_FORM_UID = "1A2B3C4D-0130-4E5F-8A6B-7C8D9E0F1A2B";
const SPEC130_LISTS_UID = "2B3C4D5E-0130-4F60-9B7C-8D9E0F1A2B3C";
const SPEC130_DATE_FROM = dateToSerial(2026, 1, 1, "1900");
const SPEC130_DATE_TO = dateToSerial(2026, 12, 31, "1900");

/** "Form"'s base rules, verbatim, in document order. */
const SPEC130_BASE_RULES = {
  priority:
    `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Invalid priority" error="Choose Low, Medium or High." ` +
    `promptTitle="Priority" prompt="Pick one" sqref="B2" xr:uid="{3C4D5E6F-0130-4071-AC8D-9E0F1A2B3C4D}"><formula1>"Low,Medium,High"</formula1></dataValidation>`,
  score:
    `<dataValidation type="whole" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Out of range" ` +
    `error="Scores run from 1 to 10." sqref="B3" xr:uid="{4D5E6F70-0130-4182-BD9E-0F1A2B3C4D5E}"><formula1>1</formula1><formula2>10</formula2></dataValidation>`,
  amount:
    `<dataValidation type="decimal" errorStyle="information" imeMode="noControl" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" ` +
    `sqref="B4" xr:uid="{5E6F7081-0130-4293-8EAF-1A2B3C4D5E6F}"><formula1>0</formula1></dataValidation>`,
  due:
    `<dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5" xr:uid="{6F708192-0130-43A4-9FB0-2B3C4D5E6F70}">` +
    `<formula1>${SPEC130_DATE_FROM}</formula1><formula2>${SPEC130_DATE_TO}</formula2></dataValidation>`,
  code:
    `<dataValidation type="textLength" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B6" ` +
    `xr:uid="{708192A3-0130-44B5-A0C1-3C4D5E6F7081}"><formula1>5</formula1></dataValidation>`,
  email:
    `<dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B7" xr:uid="{8192A3B4-0130-45C6-B1D2-4D5E6F708192}">` +
    `<formula1>ISNUMBER(SEARCH("@",B7))</formula1></dataValidation>`,
  option:
    `<dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="B8" xr:uid="{92A3B4C5-0130-46D7-82E3-5E6F708192A3}">` +
    `<formula1>$E$2:$E$6</formula1></dataValidation>`,
  colour:
    `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B10" xr:uid="{A3B4C5D6-0130-47E8-93F4-6F708192A3B4}">` +
    `<formula1>Colours</formula1></dataValidation>`,
} as const;

/** The cross-sheet list, verbatim, as Excel 2010+ writes it. */
const SPEC130_X14_RULE =
  `<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" xr:uid="{B4C5D6E7-0130-48F9-A405-708192A3B4C5}">` +
  `<x14:formula1><xm:f>Lists!$A$1:$A$4</xm:f></x14:formula1><xm:sqref>B9</xm:sqref></x14:dataValidation>`;

/** "Form"'s `<dataValidations>` block, verbatim. */
const SPEC130_BLOCK = `<dataValidations count="${Object.keys(SPEC130_BASE_RULES).length}">${Object.values(SPEC130_BASE_RULES).join("")}</dataValidations>`;

/** "Form"'s data-validation ext, verbatim. */
const SPEC130_EXT =
  `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${SPEC130_X14_NS}">` +
  `<x14:dataValidations count="1" xmlns:xm="${SPEC130_XM_NS}">${SPEC130_X14_RULE}</x14:dataValidations></ext>`;

/** What spec130-validation.xlsx holds, for the unit, round-trip and e2e specs. */
export const SPEC130_VALIDATION_XLSX = {
  title: "Data validation",
  sheetNames: ["Form", "Lists"],
  activeSheetIndex: 0,
  /** cellXfs: 2 a bold header (EXCEL_STYLES). */
  xf: { header: 2 },
  definedName: { name: "Colours", text: "Lists!$B$1:$B$3" },
  form: {
    partName: "xl/worksheets/sheet1.xml",
    dimension: "A1:E10",
    headers: { A1: "Field", B1: "Entry", E1: "Options" },
    /** A2..A10. */
    labels: ["Priority", "Score", "Amount", "Due date", "Code", "Email", "Option", "Item", "Colour"],
    /** B3's number — the only value in B2:B10. */
    score: { ref: "B3", value: 5 },
    /** E2..E6, the list B8 offers. */
    options: ["Alpha", "Bravo", "Charlie", "Delta", "Echo"],
    /** A cell under no rule. */
    free: "D2",
    rules: SPEC130_BASE_RULES,
    x14Rule: SPEC130_X14_RULE,
    block: SPEC130_BLOCK,
    ext: SPEC130_EXT,
    dates: { from: SPEC130_DATE_FROM, to: SPEC130_DATE_TO },
  },
  lists: {
    partName: "xl/worksheets/sheet2.xml",
    dimension: "A1:B4",
    /** A1..A4, the list B9 offers. */
    regions: ["North", "South", "East", "West"],
    /** B1..B3, what "Colours" names. */
    colours: ["Red", "Green", "Blue"],
  },
  /** Every part, in package order. */
  parts: [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/theme/theme1.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "docProps/core.xml",
    "docProps/app.xml",
  ],
} as const;

/** A worksheet in Excel 365's CT_Worksheet order: `after` follows sheetData. */
function spec130WorksheetXml(options: { uid: string; dimension: string; selected: boolean; sheetData: string; after: string }): string {
  return (
    XML_DECLARATION +
    worksheetOpen(options.uid) +
    `<dimension ref="${options.dimension}"/>` +
    `<sheetViews><sheetView ${options.selected ? 'tabSelected="1" ' : ""}workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>` +
    `<sheetData>${options.sheetData}</sheetData>` +
    options.after +
    `</worksheet>`
  );
}

/** docProps/app.xml with the Worksheets and Named Ranges heading pairs. */
function spec130AppXml(): string {
  const F = SPEC130_VALIDATION_XLSX;
  const titles = [...F.sheetNames, F.definedName.name];
  return (
    XML_DECLARATION +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>` +
    `<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${F.sheetNames.length}</vt:i4></vt:variant>` +
    `<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${titles.length}" baseType="lpstr">${titles.map((title) => `<vt:lpstr>${title}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts>` +
    `<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>` +
    `<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`
  );
}

export async function buildSpec130ValidationXlsxBytes(): Promise<Uint8Array> {
  const F = SPEC130_VALIDATION_XLSX;
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  const text = (ref: string, value: string, style?: number): string => {
    let index = sstIndex.get(value);
    if (index === undefined) {
      index = sst.length;
      sst.push(value);
      sstIndex.set(value, index);
    }
    return sharedCell(ref, index, style);
  };

  const formRows: string[] = [excelRow(1, "1:5", text("A1", F.form.headers.A1, F.xf.header) + text("B1", F.form.headers.B1, F.xf.header) + text("E1", F.form.headers.E1, F.xf.header))];
  F.form.labels.forEach((label, i) => {
    const r = i + 2;
    const score = `B${r}` === F.form.score.ref ? numberCell(F.form.score.ref, F.form.score.value) : "";
    const option = F.form.options[i];
    formRows.push(excelRow(r, "1:5", text(`A${r}`, label) + score + (option === undefined ? "" : text(`E${r}`, option))));
  });
  const form = spec130WorksheetXml({
    uid: SPEC130_FORM_UID,
    dimension: F.form.dimension,
    selected: true,
    sheetData: formRows.join(""),
    after: F.form.block + SPEC117_PAGE_MARGINS + `<extLst>${F.form.ext}</extLst>`,
  });

  const listRows = F.lists.regions.map((region, i) => {
    const colour = F.lists.colours[i];
    return excelRow(i + 1, "1:2", text(`A${i + 1}`, region) + (colour === undefined ? "" : text(`B${i + 1}`, colour)));
  });
  const lists = spec130WorksheetXml({ uid: SPEC130_LISTS_UID, dimension: F.lists.dimension, selected: false, sheetData: listRows.join(""), after: SPEC117_PAGE_MARGINS });

  const worksheets = [form, lists];
  const sharedStrings =
    XML_DECLARATION +
    `<sst xmlns="${MAIN_NS}" count="${worksheets.reduce((sum, xml) => sum + sharedStringReferences(xml), 0)}" uniqueCount="${sst.length}">` +
    sst.map((value) => `<si><t>${xmlText(value)}</t></si>`).join("") +
    `</sst>`;
  return excelPackage({
    mainContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    title: F.title,
    sheets: F.sheetNames.map((name, index) => ({ name, xml: worksheets[index]! })),
    sharedStrings,
    workbookPr: `<workbookPr defaultThemeVersion="202300"/>`,
    additions: {
      parts: [],
      relationships: [],
      overrides: [],
      workbookAfterSheets: `<definedNames><definedName name="${F.definedName.name}">${F.definedName.text}</definedName></definedNames>`,
    },
    app: spec130AppXml(),
  });
}
// ============================================================================
// end of SPEC-130 (xlsx)
// ============================================================================
// SPEC-130 (ODS): spec130-validation.ods — LibreOffice-shaped content
// validations: the prefix's `table:content-validations` (Calc 24.8's
// attribute order, `of:` conditions, between operands joined with `,`) and
// the cells' `table:content-validation-name`; a rule over C2:C1000 carried
// by a `table:number-rows-repeated` run of attribute-only empty cells, a
// sorted list, a two-paragraph help message, an error macro and an orphan.
// ============================================================================

/** One `table:content-validation` as Calc writes it. */
function spec130OdsValidation(
  name: string,
  condition: string,
  base: string,
  options: {
    displayList?: "unsorted" | "none" | "sort-ascending";
    help?: { title: string; paragraphs: readonly string[] };
    error?: { type: "stop" | "warning" | "information"; title?: string; paragraphs?: readonly string[]; display: boolean };
    macro?: string;
  },
): string {
  const paragraphs = (lines: readonly string[] | undefined): string => (lines ?? []).map((line) => `<text:p>${xmlText(line)}</text:p>`).join("");
  let children = "";
  if (options.help) {
    children += `<table:help-message table:title="${xmlEscape(options.help.title)}" table:display="true">${paragraphs(options.help.paragraphs)}</table:help-message>`;
  }
  if (options.error) {
    const E = options.error;
    children +=
      `<table:error-message table:message-type="${E.type}"${E.title === undefined ? "" : ` table:title="${xmlEscape(E.title)}"`} table:display="${E.display}"` +
      (E.paragraphs ? `>${paragraphs(E.paragraphs)}</table:error-message>` : `/>`);
  }
  if (options.macro !== undefined) {
    children +=
      `<table:error-macro table:execute="true"/><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:error" ` +
      `xlink:href="${xmlEscape(options.macro)}" xlink:type="simple"/></office:event-listeners>`;
  }
  return (
    `<table:content-validation table:name="${name}" table:condition="${xmlEscape(condition)}" table:allow-empty-cell="true"` +
    `${options.displayList === undefined ? "" : ` table:display-list="${options.displayList}"`} table:base-cell-address="${base}">${children}</table:content-validation>`
  );
}

const SPEC130_ODS_STOP = { type: "stop", display: true } as const;

/** Form's rules, val1–val11 in Calc's order (each B cell's own, then the
 * notes column's), verbatim. */
const SPEC130_ODS_DEFINITIONS = [
  spec130OdsValidation("val1", `of:cell-content-is-in-list("Low";"Medium";"High")`, "Form.B2", {
    displayList: "unsorted",
    help: { title: "Priority", paragraphs: ["Pick one"] },
    error: { type: "stop", title: "Invalid priority", paragraphs: ["Choose Low, Medium or High"], display: true },
  }),
  spec130OdsValidation("val2", "of:cell-content-is-whole-number() and cell-content-is-between(1,10)", "Form.B3", {
    error: { type: "warning", title: "Out of range", paragraphs: ["Enter a whole number from 1 to 10"], display: true },
  }),
  spec130OdsValidation("val3", "of:cell-content-is-decimal-number() and cell-content()>0", "Form.B4", {
    error: { type: "information", paragraphs: ["Prices are positive"], display: true },
  }),
  spec130OdsValidation("val4", "of:cell-content-is-date() and cell-content-is-between(45658,46022)", "Form.B5", { error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val5", "of:cell-content-text-length()<=5", "Form.B6", { error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val6", `of:is-true-formula(ISNUMBER(SEARCH("@";[.B7])))`, "Form.B7", { error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val7", "of:cell-content-is-in-list([.$E$2:.$E$6])", "Form.B8", { displayList: "none", error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val8", "of:cell-content-is-in-list([$Lists.$A$1:.$A$4])", "Form.B9", { displayList: "sort-ascending", error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val9", "of:cell-content-is-in-list(Colours)", "Form.B10", { displayList: "unsorted", error: SPEC130_ODS_STOP }),
  spec130OdsValidation("val10", "of:cell-content-is-whole-number() and cell-content()>=0", "Form.B11", {
    macro: "vnd.sun.star.script:Standard.Checks.RejectCode?language=Basic&location=application",
  }),
  spec130OdsValidation("val11", "of:cell-content-text-length()<=200", "Form.C2", {
    help: { title: "Notes", paragraphs: ["Free text", "Keep it short"] },
    error: { type: "stop", display: false },
  }),
] as const;

/** A definition no cell references. */
const SPEC130_ODS_ORPHAN = spec130OdsValidation("val12", "of:cell-content-is-decimal-number() and cell-content()>=0", "Form.F2", { error: SPEC130_ODS_STOP });

const SPEC130_ODS_NO_INPUT = { show: false, title: "", text: "" } as const;
const SPEC130_ODS_STOP_ERROR = { show: true, style: "stop", title: "", text: "" } as const;

/** What spec130-validation.ods holds, for the unit and e2e specs. Ranges in
 * A1 notation; formulas canonical, relative to each rule's origin. */
export const SPEC130_VALIDATION_ODS = {
  title: "Data validation",
  tableNames: ["Form", "Lists"],
  activeTable: "Form",
  /** A1:A11. */
  labels: ["Field", "Priority", "Quantity", "Price", "Due", "Code", "Email", "Region", "Fruit", "Colour", "Check"],
  headers: { B1: "Value", C1: "Notes", E1: "Regions" },
  /** B3's value; the rest of B2:B11 and all of C2:C1000 are empty. */
  B3: 5,
  /** E2:E6, B8's list source. */
  regions: ["North", "South", "East", "West", "Centre"],
  /** Lists!A1:A4 (B9's sorted list) and Lists!B1:B3 (the name Colours). */
  fruits: ["Pear", "Apple", "Mango", "Banana"],
  colours: ["Red", "Green", "Blue"],
  namedRange: { name: "Colours", baseCellAddress: "$Lists.$B$1", cellRangeAddress: "$Lists.$B$1:.$B$3" },
  definitions: SPEC130_ODS_DEFINITIONS,
  orphan: SPEC130_ODS_ORPHAN,
  /** The notes rule's rows 12–1000: one repeated row of attribute-only cells. */
  notesRun: { firstRow: 12, rows: 989 },
  /** Form's rules as the reader makes them (model order = definition order). */
  rules: [
    {
      ranges: ["B2"],
      type: "list",
      operator: "between",
      formula1: `"Low,Medium,High"`,
      allowBlank: true,
      showDropdown: true,
      input: { show: true, title: "Priority", text: "Pick one" },
      error: { show: true, style: "stop", title: "Invalid priority", text: "Choose Low, Medium or High" },
      odsName: "val1",
    },
    {
      ranges: ["B3"],
      type: "whole",
      operator: "between",
      formula1: "1",
      formula2: "10",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: { show: true, style: "warning", title: "Out of range", text: "Enter a whole number from 1 to 10" },
      odsName: "val2",
    },
    {
      ranges: ["B4"],
      type: "decimal",
      operator: "greaterThan",
      formula1: "0",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: { show: true, style: "information", title: "", text: "Prices are positive" },
      odsName: "val3",
    },
    {
      ranges: ["B5"],
      type: "date",
      operator: "between",
      formula1: "45658",
      formula2: "46022",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val4",
    },
    {
      ranges: ["B6"],
      type: "textLength",
      operator: "lessThanOrEqual",
      formula1: "5",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val5",
    },
    {
      ranges: ["B7"],
      type: "custom",
      operator: "between",
      formula1: `ISNUMBER(SEARCH("@",B7))`,
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val6",
    },
    {
      ranges: ["B8"],
      type: "list",
      operator: "between",
      formula1: "$E$2:$E$6",
      allowBlank: true,
      showDropdown: false,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val7",
    },
    {
      ranges: ["B9"],
      type: "list",
      operator: "between",
      formula1: "Lists!$A$1:$A$4",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val8",
      odsListOrder: "sort-ascending",
    },
    {
      ranges: ["B10"],
      type: "list",
      operator: "between",
      formula1: "Colours",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      error: SPEC130_ODS_STOP_ERROR,
      odsName: "val9",
    },
    {
      ranges: ["B11"],
      type: "whole",
      operator: "greaterThanOrEqual",
      formula1: "0",
      allowBlank: true,
      showDropdown: true,
      input: SPEC130_ODS_NO_INPUT,
      // The error macro is kept, never run: the rule enforces nothing.
      error: { show: false, style: "stop", title: "", text: "" },
      odsName: "val10",
    },
    {
      ranges: ["C2:C1000"],
      type: "textLength",
      operator: "lessThanOrEqual",
      formula1: "200",
      allowBlank: true,
      showDropdown: true,
      input: { show: true, title: "Notes", text: "Free text\nKeep it short" },
      error: { show: false, style: "stop", title: "", text: "" },
      odsName: "val11",
    },
  ],
} as const;

/** An empty cell carrying only its validation's name. */
const spec130Validated = (name: string): string => `<table:table-cell table:content-validation-name="${name}"/>`;

export async function buildSpec130ValidationOdsBytes(): Promise<Uint8Array> {
  const O = SPEC130_VALIDATION_ODS;
  const header =
    loString(O.labels[0], "ce5") + loString(O.headers.B1, "ce5") + loString(O.headers.C1, "ce5") + loEmpty(1) + loString(O.headers.E1, "ce5") + loEmpty(16379);
  const rows = O.labels.slice(1).map((label, index) => {
    const name = `val${index + 1}`;
    const b =
      index === 1
        ? `<table:table-cell table:content-validation-name="${name}" office:value-type="float" office:value="${O.B3}" calcext:value-type="float"><text:p>${O.B3}</text:p></table:table-cell>`
        : spec130Validated(name);
    const region = O.regions[index];
    const rest = region === undefined ? loEmpty(16381) : loEmpty(1) + loString(region) + loEmpty(16379);
    return loRow(loString(label) + b + spec130Validated("val11") + rest);
  });
  const form =
    `<table:table table:name="Form" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    loRow(header) +
    rows.join("") +
    loRow(loEmpty(2) + spec130Validated("val11") + loEmpty(16381), "ro1", O.notesRun.rows) +
    loRow(loEmpty(16384), "ro1", 1_048_576 - (O.notesRun.firstRow - 1) - O.notesRun.rows) +
    `</table:table>`;
  const listRows = O.fruits.map((fruit, index) => {
    const colour = O.colours[index];
    return loRow(loString(fruit) + (colour === undefined ? loEmpty(16383) : loString(colour) + loEmpty(16382)));
  });
  const lists =
    `<table:table table:name="Lists" table:style-name="ta1">` +
    `<table:table-column table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"/>` +
    listRows.join("") +
    loRow(loEmpty(16384), "ro1", 1_048_576 - O.fruits.length) +
    `</table:table>`;
  const N = O.namedRange;
  const content =
    ODF_XML_DECLARATION +
    `<office:document-content ${LO_CONTENT_NAMESPACES}><office:scripts/>${LO_FONT_FACES}${LO_AUTOMATIC_STYLES}` +
    `<office:body><office:spreadsheet>${LO_CALCULATION_SETTINGS}` +
    `<table:content-validations>${O.definitions.join("")}${O.orphan}</table:content-validations>` +
    form +
    lists +
    `<table:named-expressions><table:named-range table:name="${N.name}" table:base-cell-address="${N.baseCellAddress}" table:cell-range-address="${N.cellRangeAddress}"/></table:named-expressions>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const cellCount = 4 + O.labels.length - 1 + 1 + O.regions.length + O.fruits.length + O.colours.length;
  return loPackage(content, loMeta(O.title, O.tableNames.length, cellCount), loSettings(O.tableNames, O.activeTable));
}
