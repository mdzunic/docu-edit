// The .xlsx reader's phases (SPEC-137 "The read phases"; the types and the
// rules are readPhases.ts's). Together they read a package exactly as the
// single pass does (reader.ts parseXlsx + open.ts assembleSpreadsheet + the
// chart refresh readSpreadsheetBytes ran) — open.spec.ts pins the equality —
// with the O(cells) parts read where DOMParser does not exist:
//
// 1. xlsxSourcePartNames (main thread) names the shared strings, the styles
//    part and every grid sheet's worksheet part; prepareXlsxSources
//    (DOMParser-free) tells them apart by their root element (an SML `sst`,
//    `styleSheet`, anything else a worksheet it leaves for phase 3), reads
//    the styles part's `<indexedColors>` palette with the scanner
//    (styles.ts readXlsxIndexedColors) and the shared strings against it.
// 2. readXlsxShell (main thread): everything parseXlsx does but scan the
//    worksheets — detection with the same refusals, workbook.xml, theme,
//    styles, persons, chart sheets, every sheet's id (in sheet order),
//    defined names, `opaque`, properties … — the shell assembled exactly as
//    assembleSpreadsheet assembles a model; one job per grid sheet whose part
//    exists; the scan context as plain JSON (copies of the style and dxf
//    tables the scans append to, the implied-format record, the palette, the
//    MDW, the sheet names). Phase 1 decided the shared strings' roles and
//    palette without the workbook's relationships, so the shell checks its
//    choice against them (the same part, the same palette); where they differ
//    — a part only one parser accepts, an SST whose root is not `sst` — the
//    shell reads the table itself and the context carries it.
// 3. startXlsxSheetScan (DOMParser-free, stepwise): sheetXml.ts's worksheet
//    read, every sheet counted against one ScanBudget; then the read rule
//    for hidden rows (filter.ts classifyHiddenRows) as reader.ts applies it;
//    the rows emitted as RowChunks, the rest returned as `fields`, the style
//    and dxf entries this sheet appended as `appended`, and the part's text
//    with its cells cut out as `deferred`. Conditional formats and
//    validations are named with placeholders: the scan never mints an id.
// 4. The shell's completion (main thread, after assembleWorkbook joined
//    rows, fields and appended entries): placeholder ids re-minted; each
//    sheet's drawing layer, tables, notes, threads and hyperlinks read from
//    the package through the deferred text (the worksheet part is never
//    decoded again); the media; the card preview; the display refresh of
//    uncached charts.
//
// Errors: every refusal is the single pass's. A workbook with one fault
// fails with its message; with several, the first reported may differ,
// since the phases visit the parts in another order.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import type { MediaPart } from "../../core/media";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import type { Dxf } from "../conditionalFormats";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { classifyHiddenRows } from "../filter";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT, maxDigitWidthPx } from "../geometry";
import { ImpliedFormatStyles } from "../impliedFormats";
import type { CellStyle, Sheet, SpreadsheetModel } from "../model";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { workbookFormatContext } from "../numberFormat/render";
import { encodeOpaque, partText, resolveActiveSheet, type ParsedSpreadsheet } from "../parsed";
import { buildPreview } from "../preview";
import { setShellCompletion, type RowChunk, type ScanBudget, type SheetScanJob, type SheetScanResult, type SheetScanSteps, type WorkbookSources } from "../readPhases";
import { chunkRows } from "../rowChunks";
import { readXlsxPersons, readXlsxSheetAnnotations } from "./annotations";
import { readXlsxDefinedNames } from "./definedNames";
import { readXlsxChartSheet, readXlsxSheetDrawings } from "./drawingXml";
import { partTextOf, relationshipKind, relationshipsOf, workbookPathOf } from "./package";
import { readXlsxPageSetup } from "./pageSetup";
import { coreProperties, detectXlsx, readTheme, readWorkbook, scanned, type WorkbookInfo } from "./reader";
import type { SharedStringTable } from "./sharedStrings";
import { readSharedStrings, startWorksheetRead } from "./sheetXml";
import { isSpreadsheetMl, readXlsxIndexedColors, readXlsxStyles } from "./styles";
import { readXlsxSheetTables } from "./tableXml";

/** What a worksheet scan needs from the shell phase — plain JSON (it is
 * cloned, or posted to the parse worker, once per read, and every job of
 * the read scans with the same object, in sheet order). */
export interface XlsxScanContext {
  format: "xlsx";
  dateSystem: "1900" | "1904";
  /** The Normal font's maximum digit width (column widths → px). */
  mdw: number;
  /** The workbook's `<indexedColors>` palette (tab colours, inline runs). */
  indexedColors?: string[];
  /** A copy of the style table (cellXfs), which implied formats grow. */
  styles: CellStyle[];
  /** ImpliedFormatStyles' derived-style record, shared by every job. */
  implied: Record<string, number>;
  /** A copy of the differential styles, which x14 rules' dxfs grow. */
  dxfs: Dxf[];
  /** Every sheet's name by index (messages, conditional formats). */
  sheetNames: string[];
  /** The shared strings, when the shell read them itself (see the header). */
  sharedStrings?: SharedStringTable;
}

/** What a worksheet scan leaves for the main thread to finish — plain JSON. */
export interface XlsxScanDeferred {
  format: "xlsx";
  /** The worksheet part's text with its cells cut out (sheetXml.ts
   * WorksheetRead.withoutCells). */
  worksheetXml: string;
}

/** The locale the read rule and the preview format in (the reader knows no
 * preference — open.ts's READ_PREVIEW_LOCALE). */
const READ_LOCALE = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];

/** The root element of a part's text: its local name and namespace, or
 * undefined when the scanner cannot reach one. Reads only up to it. */
function rootOf(text: string): { local: string; ns: string | undefined } | undefined {
  let root: { local: string; ns: string | undefined } | undefined;
  try {
    scanXml(text, (event) => {
      if (event.type === "open") {
        root = { local: event.local, ns: event.ns };
        return "stop";
      }
      return undefined;
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  return root;
}

/** Whether a part's root is SpreadsheetML's `local` — decoding only its
 * first 64 KiB while that holds the root's start tag. */
function partRootIs(bytes: Uint8Array | undefined, local: string): boolean {
  if (bytes === undefined) {
    return false;
  }
  const head = bytes.length > 65_536 ? rootOf(partText(bytes.subarray(0, 65_536))) : undefined;
  const root = head ?? rootOf(partText(bytes));
  return root !== undefined && root.local === local && isSpreadsheetMl(root.ns);
}

/** The package's main part and workbook, or undefined when the shell will
 * refuse it (it reports why). */
function workbookOf(parts: Map<string, Uint8Array>): WorkbookInfo | undefined {
  try {
    const info = detectPackage(parts);
    const detected = spreadsheetFormatOf(info);
    if (info.container !== "opc" || detected === null || detected === "xlsb" || detected.format !== "xlsx") {
      return undefined;
    }
    return readWorkbook(parts, info.mainPartPath);
  } catch {
    return undefined;
  }
}

/** The parts phase 1 and phase 3 read (main thread, before phase 1 — the
 * parse client posts copies of exactly these to the worker): the shared
 * strings, the styles part, then every grid sheet's worksheet part — none
 * when the shell will refuse the package. */
export function xlsxSourcePartNames(parts: Map<string, Uint8Array>): string[] {
  const workbook = workbookOf(parts);
  if (!workbook) {
    return [];
  }
  const names = new Set<string>();
  for (const name of [workbook.sharedStringsPath, workbook.stylesPath]) {
    if (name !== undefined && parts.has(name)) {
      names.add(name);
    }
  }
  for (const entry of workbook.sheets) {
    if (entry.kind === "grid" && parts.has(entry.partName)) {
      names.add(entry.partName);
    }
  }
  return [...names];
}

/** Phase 1 (DOMParser-free): the shared strings, their run colours resolved
 * against the styles part's palette (see the module header). A malformed
 * SST is refused as the single pass refuses it. */
export function prepareXlsxSources(texts: Map<string, string>): WorkbookSources {
  let sst: { name: string; text: string } | undefined;
  let styles: string | undefined;
  for (const [name, text] of texts) {
    const root = rootOf(text);
    if (root === undefined || !isSpreadsheetMl(root.ns)) {
      continue;
    }
    if (root.local === "sst") {
      sst ??= { name, text };
    } else if (root.local === "styleSheet") {
      styles ??= text;
    }
  }
  const palette = readXlsxIndexedColors(styles);
  const sharedStrings = sst === undefined ? [] : scanned(sst.name, () => readSharedStrings(sst.text, palette));
  return { format: "xlsx", sharedStrings };
}

function samePalette(a: readonly string[] | undefined, b: readonly string[] | undefined): boolean {
  return a === b || (a !== undefined && b !== undefined && a.length === b.length && a.every((color, index) => color === b[index]));
}

/** A fresh grid/foreign sheet as the single pass starts each one. */
function sheetShell(entry: WorkbookInfo["sheets"][number]): Sheet {
  return {
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
    conditionalFormats: [],
    validations: [],
  };
}

/** Whether a sheet part relates a drawing (the only way a chart is read). */
function relatesDrawing(parts: Map<string, Uint8Array>, partName: string): boolean {
  try {
    return relationshipsOf(parts, partName).some((rel) => relationshipKind(rel.type) === "drawing");
  } catch {
    return true; // its reader reports the broken relationships
  }
}

/** The scan's placeholder ids (conditional formats, rules, validations)
 * replaced by fresh ones — one per distinct placeholder. */
function remintScanIds(sheet: Sheet): void {
  const ids = new Map<string, string>();
  const fresh = (id: string): string => {
    let next = ids.get(id);
    if (next === undefined) {
      next = crypto.randomUUID();
      ids.set(id, next);
    }
    return next;
  };
  if (sheet.conditionalFormats.length > 0) {
    sheet.conditionalFormats = sheet.conditionalFormats.map((format) => ({ ...format, id: fresh(format.id), rules: format.rules.map((rule) => ({ ...rule, id: fresh(rule.id) })) }));
  }
  if (sheet.validations.length > 0) {
    sheet.validations = sheet.validations.map((rule) => ({ ...rule, id: fresh(rule.id) }));
  }
}

/** Phase 2 (main thread): the shell, the scan jobs and their context, and
 * the completion assembleWorkbook runs (see the module header). */
export async function readXlsxShell(
  parts: Map<string, Uint8Array>,
  sources: WorkbookSources,
  title: string,
): Promise<{ shell: SpreadsheetModel; jobs: SheetScanJob[]; context: XlsxScanContext }> {
  const { info, detected } = detectXlsx(parts);
  const workbook = readWorkbook(parts, info.mainPartPath);
  const dateSystem = workbook.date1904 ? "1904" : "1900";
  const theme = readTheme(parts, workbook.themePath);
  const stylesBytes = workbook.stylesPath ? parts.get(workbook.stylesPath) : undefined;
  const stylesText = stylesBytes ? partText(stylesBytes) : undefined;
  const styleTable = readXlsxStyles(stylesText, theme);
  const indexedColors = styleTable.indexedColors;
  const mdw = maxDigitWidthPx(styleTable.defaultFont.family, styleTable.defaultFont.sizePt);
  // Phase 1 read the part rooted `sst` (the SST comes first among the
  // sources) against the palette of the part rooted `styleSheet`: the right
  // table exactly when that is the workbook's SST and the palette is the
  // style table's. Else the shell reads it (and refuses it) as the single
  // pass does, here, in the single pass's order.
  const sstBytes = workbook.sharedStringsPath ? parts.get(workbook.sharedStringsPath) : undefined;
  const phase1Palette = partRootIs(stylesBytes, "styleSheet") ? readXlsxIndexedColors(stylesText) : undefined;
  let sharedStrings: SharedStringTable | undefined;
  if ((sstBytes !== undefined && !partRootIs(sstBytes, "sst")) || !samePalette(phase1Palette, indexedColors) || sources.format !== "xlsx") {
    sharedStrings = sstBytes ? scanned(workbook.sharedStringsPath!, () => readSharedStrings(partText(sstBytes), indexedColors)) : [];
  }
  const persons = readXlsxPersons(parts);
  const sheets: Sheet[] = [];
  const jobs: SheetScanJob[] = [];
  let drawings = false;
  for (const [sheetIndex, entry] of workbook.sheets.entries()) {
    const sheet = sheetShell(entry);
    const bytes = parts.get(entry.partName);
    if (entry.kind === "grid" && bytes) {
      jobs.push({ sheetIndex, partName: entry.partName });
      drawings ||= relatesDrawing(parts, entry.partName);
    } else if (entry.chartsheet && bytes) {
      const chart = readXlsxChartSheet(parts, entry.partName, sheet);
      if (chart) {
        sheet.kind = "chart";
        sheet.drawings = chart.drawings;
        sheet.drawingPart = chart.drawingPart;
        drawings = true;
      }
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
    dxfs: styleTable.dxfs,
  };
  const workbookPath = workbookPathOf(parts);
  const workbookXml = workbookPath === undefined ? undefined : partTextOf(parts, workbookPath);
  if (workbookXml !== undefined) {
    const names = readXlsxDefinedNames(workbookXml, sheets.map((sheet) => sheet.id));
    if (names.length > 0) {
      parsed.names = names;
    }
  }
  if (theme !== undefined) {
    parsed.theme = theme;
  }
  if (indexedColors !== undefined) {
    parsed.indexedColors = indexedColors;
  }
  if (styleTable.namedStyles !== undefined) {
    parsed.namedStyles = styleTable.namedStyles;
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
  if (persons !== undefined && persons.length > 0) {
    parsed.persons = persons;
  }
  // The one assembly every open path shares (open.ts imports this module,
  // so it is loaded by now); its preview is rebuilt once the rows are in.
  const { assembleSpreadsheet } = await import("../open");
  // The completion is synchronous: the chart codec loads here, when some
  // sheet relates a drawing (a chart is only ever read through one).
  const charts = drawings ? await import("../chartsOpen") : undefined;
  const shell = assembleSpreadsheet(title, "xlsx", parsed);
  const context: XlsxScanContext = {
    format: "xlsx",
    dateSystem,
    mdw,
    styles: styleTable.styles.slice(),
    implied: {},
    dxfs: styleTable.dxfs.slice(),
    sheetNames: sheets.map((sheet) => sheet.name),
  };
  if (indexedColors !== undefined) {
    context.indexedColors = indexedColors;
  }
  if (sharedStrings !== undefined) {
    context.sharedStrings = sharedStrings;
  }
  setShellCompletion(shell, (model, results) => {
    const media = new Map<string, MediaPart>();
    for (const result of results) {
      const sheet = model.sheets[result.sheetIndex];
      const partName = sheet?.partName;
      if (!sheet || partName === undefined) {
        continue;
      }
      remintScanIds(sheet);
      const xml = result.deferred?.format === "xlsx" ? result.deferred.worksheetXml : (partTextOf(parts, partName) ?? "");
      const layer = readXlsxSheetDrawings(parts, partName, sheet, media, xml);
      if (layer.drawingPart !== undefined) {
        sheet.drawingPart = layer.drawingPart;
      }
      if (layer.drawings.length > 0) {
        sheet.drawings = layer.drawings;
      }
      // SPEC-133: the worksheet's <tableParts> → its xl/tables/tableN.xml parts.
      const tables = readXlsxSheetTables(parts, partName, xml, () => crypto.randomUUID());
      if (tables.length > 0) {
        sheet.tables = tables;
      }
      // SPEC-135: notes, threads and hyperlinks.
      Object.assign(sheet, readXlsxSheetAnnotations(parts, partName, xml, { ...(persons ? { persons } : {}), ...(indexedColors ? { indexedColors } : {}) }));
      // SPEC-136: the page elements (the deferred text keeps everything after sheetData).
      const pageSetup = readXlsxPageSetup(xml);
      if (pageSetup !== undefined) {
        sheet.pageSetup = pageSetup;
      }
    }
    if (media.size > 0) {
      model.media = [...media.values()];
    }
    model.preview = buildPreview(model, READ_LOCALE);
    if (charts && model.sheets.some((sheet) => (sheet.drawings ?? []).some((drawing) => drawing.content.kind === "chart" && drawing.content.refs !== undefined))) {
      charts.refreshUncachedCharts(model);
    }
  });
  return { shell, jobs, context };
}

/** Phase 3 (DOMParser-free), stepwise: one worksheet part (see the module
 * header). Every job of one read shares `budget` and `context`. */
export function startXlsxSheetScan(
  job: SheetScanJob,
  text: string,
  sources: WorkbookSources,
  context: XlsxScanContext,
  emit: (chunk: RowChunk) => void,
  budget: ScanBudget,
): SheetScanSteps {
  const sheetName = context.sheetNames[job.sheetIndex] ?? "";
  const stylesBefore = context.styles.length;
  const dxfsBefore = context.dxfs.length;
  const holder: Pick<SpreadsheetModel, "dxfs"> = { dxfs: context.dxfs };
  let cells = 0;
  let placeholders = 0;
  const read = startWorksheetRead(text, sheetName, {
    sharedStrings: context.sharedStrings ?? sources.sharedStrings ?? [],
    dateSystem: context.dateSystem,
    budget: {
      take(count: number): boolean {
        if (budget.cellsLeft < count) {
          return false;
        }
        budget.cellsLeft -= count;
        cells += count;
        return true;
      },
    },
    mdw: context.mdw,
    ...(context.indexedColors === undefined ? {} : { indexedColors: context.indexedColors }),
    implied: new ImpliedFormatStyles(context.styles, context.implied),
    dxfs: holder,
    mint: () => `scan-${job.sheetIndex}-${(placeholders += 1)}`,
  });
  let result: SheetScanResult | undefined;
  return {
    step(rows: number): boolean {
      if (result) {
        return true;
      }
      if (!scanned(job.partName, () => read.step(rows))) {
        return false;
      }
      const content = read.content();
      if (content === null) {
        throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
      context.dxfs = holder.dxfs ?? context.dxfs;
      const { rows: read_, ...fields } = content;
      let sheetRows = read_;
      const filter = content.autoFilter?.range;
      if (filter !== undefined && sheetRows.some((row) => row.hidden === true && row.index > filter.top && row.index <= filter.bottom)) {
        // SPEC-128's read rule, as reader.ts applies it to the read sheet.
        const sheet: Sheet = { ...sheetShellFor(job, sheetName), ...content };
        const formatContext = workbookFormatContext({ dateSystem: context.dateSystem, format: "xlsx", indexedColors: context.indexedColors }, READ_LOCALE);
        sheetRows = classifyHiddenRows({ styles: context.styles }, sheet, formatContext);
      }
      for (const [chunk, rowsOfChunk] of chunkRows(sheetRows)) {
        emit({ sheetIndex: job.sheetIndex, chunk, rows: rowsOfChunk });
      }
      result = { sheetIndex: job.sheetIndex, fields, cells, deferred: { format: "xlsx", worksheetXml: read.withoutCells() } };
      const styles = context.styles.slice(stylesBefore);
      const dxfs = context.dxfs.slice(dxfsBefore);
      if (styles.length > 0 || dxfs.length > 0) {
        result.appended = { ...(styles.length > 0 ? { styles } : {}), ...(dxfs.length > 0 ? { dxfs } : {}) };
      }
      return true;
    },
    result(): SheetScanResult {
      if (!result) {
        throw new Error("startXlsxSheetScan: the sheet is not scanned yet");
      }
      return result;
    },
  };
}

/** The sheet the read rule judges (the single pass's sheet at that point). */
function sheetShellFor(job: SheetScanJob, name: string): Sheet {
  return {
    id: "",
    name,
    visibility: "visible",
    kind: "grid",
    rows: [],
    partName: job.partName,
    columns: [],
    defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
    merges: [],
    conditionalFormats: [],
    validations: [],
  };
}
