// The .ods reader's phases (SPEC-137 "The read phases"; the types and the
// rules are readPhases.ts's). Together they read a package exactly as the
// single pass does (reader.ts parseOds + open.ts assembleSpreadsheet + the
// chart refresh readSpreadsheetBytes ran) — open.spec.ts pins the equality —
// with content.xml, the one O(cells) part, read where DOMParser does not
// exist. They share the single pass's code: reader.ts's OdsContentOutline
// (content.xml outside its tables) and OdsTableScan (one table's events).
//
// 1. odsSourcePartNames (main thread) names content.xml — nothing when the
//    shell will refuse the package (no OpenDocument spreadsheet, or no
//    content part). prepareOdsSources (DOMParser-free) scans it once with
//    OdsContentOutline, skipping the tables' insides: the envelope (prefix
//    and suffix), each top-level table's sourceXml offsets — sliced exactly
//    as the single pass slices them, so prefix + every sourceXml + suffix is
//    content.xml — and the outline (WorkbookSources.odsOutline): each table's
//    start tag and the namespaces in scope there, the root's declarations,
//    the style sections' offsets, the null date (at each table and at the
//    end), the calculation settings and the structure protection. A
//    malformed part, or one with no table, is refused as the single pass
//    refuses it.
// 2. readOdsShell (main thread): detection with the same refusals,
//    styles.xml, the style book over the prefix's style sections (refusing
//    them as the single pass does); the shell — one sheet per table, ids
//    minted here in sheet order, `opaque`, the envelope, the dates, the
//    structure protection, the named styles — assembled as
//    assembleSpreadsheet assembles a model; one job per table (its sourceXml
//    range in content.xml); and the scan context: the style book's plain-data
//    projection (OdsStyleBook.scanProjection — the book itself holds DOM
//    elements) with the growing tables it starts (styles, implied formats,
//    names met, dxfs), the dates, the namespaces.
// 3. startOdsSheetScan (DOMParser-free, stepwise): one table cut out of
//    content.xml by offset (core/xmlScanner.ts's range scan, from its start
//    tag, under the namespaces in scope there — no copy of the text, and every
//    offset the whole text's), its events through OdsTableScan against
//    ods/scanBook.ts's book, every sheet counted against one ScanBudget,
//    pausing between rows; then its conditional formats from its sourceXml
//    (ods/conditionalFormats.ts, named with placeholders: the scan never
//    mints an id) against the shared dxf table. The rows go out as RowChunks,
//    the rest as `fields`, the style and dxf entries this sheet added as
//    `appended`, and as `deferred` what needs DOMParser or the package: the
//    drawing elements, the content-validation names' rectangles, the table's
//    own `table:named-expressions` element.
// 4. The shell's completion (main thread, after assembleWorkbook joined rows,
//    fields and appended entries), in the single pass's order: placeholder
//    ids re-minted; each sheet's drawings under its geometry, sheet by sheet
//    (so the media come in the single pass's order); the suffix's
//    autofilters; the prefix's content validations and the orphans; the
//    names (the suffix's, then each table's); settings.xml's view settings,
//    active table and AutoCalculate; meta.xml; the media; the display refresh
//    of uncached charts; the card preview; the generic chart refresh the old
//    readSpreadsheetBytes ran.
//
// Errors: every refusal is the single pass's. A workbook with one fault
// fails with its message; with several, the first reported may differ, since
// the phases visit the parts in another order (phase 1 checks all of
// content.xml before styles.xml is read). Content between two tables (which
// ODF's schema rules out) is the later table's sourceXml and is read by the
// outline, as in the single pass; the one shape read differently is an
// `office:spreadsheet` element nested in a table outside its cells and
// drawings, which the single pass's outline follows and phase 1 skips.
import { parseOdfMeta } from "../../core/odfMeta";
import { scanXml, startXmlScan, type XmlScanAttribute } from "../../core/xmlScanner";
import type { GridRange } from "../address";
import type { Dxf } from "../conditionalFormats";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { buildGridMetrics, ODS_DEFAULT_COL_WIDTH_PX, ODS_DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import type { CalcSettings, Sheet, SheetChartContent, SpreadsheetModel } from "../model";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { encodeOpaque, parseError, partText, resolveActiveSheet, type ParsedSpreadsheet } from "../parsed";
import { buildPreview } from "../preview";
import { setShellCompletion, type RowChunk, type ScanBudget, type SheetScanJob, type SheetScanResult, type SheetScanSteps, type WorkbookSources } from "../readPhases";
import { chunkRows } from "../rowChunks";
import { readOdsTableConditionalFormats } from "./conditionalFormats";
import { readOdsContentValidations } from "./contentValidations";
import { readOdsAutoFilters } from "./databaseRanges";
import { OdsMediaReader, readOdsDrawings, type OdsDrawingSource } from "./drawingXml";
import { odsTableNamesSpan, readOdsTableNames, readOdsWorkbookNames } from "./namedExpressions";
import { readOdsPageSetups } from "./pageSetup";
import { applyViewSettings, detectOds, OdsContentOutline, OdsTableScan, refreshUncachedOdsCharts, scanned, type OdsDates, type OdsTableScanEnv } from "./reader";
import { OdsScanBook, type OdsScanBookState } from "./scanBook";
import { readOdsSettings } from "./settings";
import { emptyOdsStyleSheet, OdsStyleBook, readOdsContentStyles, readOdsStyleSheet, type OdsStyleProjection } from "./styles";

/** What phase 1's scan of content.xml found outside the tables (the
 * single pass's OdsContentOutline at the end of its scan, as plain data). */
export interface OdsSourcesOutline {
  /** Each table's start tag (its sourceXml starts at or before it). */
  tags: number[];
  /** The namespaces in scope at the tables' start tags (prefix → URI, ""
   * the default): what a table's range scan starts with. */
  scope: [string, string][];
  /** The root's namespace declarations, as source text. */
  declarations: string;
  /** The root's prefix → URI declarations (formula dialect prefixes). */
  namespaces: [string, string][];
  /** The style sections the single pass's style book reads — those met
   * before the first table: [start, end) in content.xml (in the prefix). */
  styleSlices: { start: number; end: number }[];
  /** The dates in effect at each table, and at the end (the workbook's). */
  tableDates: OdsDates[];
  dates: OdsDates;
  /** SPEC-120: `table:calculation-settings` (settings.xml's AutoCalculate
   * joins in the completion). */
  calc: CalcSettings;
  /** SPEC-122: `office:spreadsheet@table:structure-protected="true"`. */
  structureProtected: boolean;
}

/** What a table scan needs from the shell phase — plain JSON (cloned, or
 * posted to the parse worker, once per read; every job of the read scans
 * with the same object, in sheet order, and grows its tables). */
export interface OdsScanContext extends OdsScanBookState {
  format: "ods";
  /** The dates each table's values count from (by sheet index). */
  tableDates: OdsDates[];
  dates: OdsDates;
  /** The root's namespace declarations (notes parse under them). */
  declarations: string;
  /** The root's prefix → URI declarations. */
  namespaces: [string, string][];
  /** The namespaces in scope at a table's start tag. */
  scope: [string, string][];
  /** The differential styles the conditional formats apply, which grow. */
  dxfs: Dxf[];
  /** The style book, projected (see ods/scanBook.ts). */
  book: OdsStyleProjection;
}

/** What a table scan leaves for the main thread to finish — plain JSON. */
export interface OdsScanDeferred {
  format: "ods";
  /** SPEC-131: the drawing elements met, in document order. */
  drawings: OdsDrawingSource[];
  /** SPEC-130: each content validation name's rectangles, first met first. */
  validations: [string, GridRange[]][];
  /** SPEC-133: the table's own `table:named-expressions` element, verbatim. */
  namesXml?: string;
}

/** The locale the preview formats in (open.ts's READ_PREVIEW_LOCALE). */
const READ_LOCALE = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];

/** The part phases 1 and 3 read (main thread, before phase 1 — the parse
 * client posts a copy of it to the worker): content.xml, or none when the
 * shell will refuse the package. */
export function odsSourcePartNames(parts: Map<string, Uint8Array>): string[] {
  try {
    detectOds(parts);
  } catch {
    return [];
  }
  return parts.has("content.xml") ? ["content.xml"] : [];
}

/** `scope` with an element's own namespace declarations (the same map when
 * it declares none), as the scanner scopes them. */
function withDeclarations(scope: Map<string, string>, attrs: readonly XmlScanAttribute[]): Map<string, string> {
  let declared: Map<string, string> | undefined;
  for (const attr of attrs) {
    if (attr.name === "xmlns" || attr.name.startsWith("xmlns:")) {
      declared ??= new Map(scope);
      declared.set(attr.name === "xmlns" ? "" : attr.name.slice(6), attr.value);
    }
  }
  return declared ?? scope;
}

/** Phase 1 (DOMParser-free): content.xml's envelope, tables and outline
 * (see the module header). No content.xml: nothing — the shell refuses. */
export function prepareOdsSources(texts: Map<string, string>): WorkbookSources {
  const xml = texts.get("content.xml");
  if (xml === undefined) {
    return { format: "ods" };
  }
  return scanned("content.xml", () => outlineOf(xml));
}

function outlineOf(xml: string): WorkbookSources {
  const outline = new OdsContentOutline(xml);
  const tables: { start: number; end: number }[] = [];
  const tags: number[] = [];
  const tableDates: OdsDates[] = [];
  let styleSlices: { start: number; end: number }[] | undefined;
  let tableScope: [string, string][] | undefined;
  /** The open elements' scopes, outside the tables. */
  const scopes: Map<string, string>[] = [new Map()];
  /** The table being skipped, and the depth inside it. */
  let skipping: { start: number } | null = null;
  let depth = 0;
  scanXml(xml, (event) => {
    if (skipping !== null) {
      if (event.type === "open") {
        depth += event.selfClosing ? 0 : 1;
      } else if (event.type === "close") {
        if (depth === 0) {
          tables.push({ start: skipping.start, end: event.end });
          outline.tableEnded(event.end, false);
          skipping = null;
        } else {
          depth -= 1;
        }
      }
      return undefined;
    }
    const scope = scopes.at(-1)!;
    const opened = outline.event(event);
    if (opened !== undefined && event.type === "open") {
      tags.push(event.start);
      tableDates.push(outline.dates);
      // The single pass builds its style book at the first table's start tag.
      styleSlices ??= outline.styleSlices.slice();
      tableScope ??= [...scope];
      if (event.selfClosing) {
        tables.push({ start: opened.start, end: event.end });
        outline.tableEnded(event.end, true);
      } else {
        skipping = { start: opened.start };
        depth = 0;
      }
      return undefined;
    }
    if (event.type === "open" && !event.selfClosing) {
      scopes.push(withDeclarations(scope, event.attrs));
    } else if (event.type === "close") {
      scopes.pop();
    }
    return undefined;
  });
  if (tables.length === 0 || outline.firstTableStart < 0) {
    throw parseError("content.xml", "the spreadsheet has no tables");
  }
  return {
    format: "ods",
    odsEnvelope: { prefix: xml.slice(0, outline.firstTableStart), suffix: xml.slice(outline.previousTableEnd) },
    odsTables: tables,
    odsOutline: {
      tags,
      scope: tableScope ?? [],
      declarations: outline.declarations,
      namespaces: [...outline.namespaces],
      styleSlices: styleSlices ?? [],
      tableDates,
      dates: outline.dates,
      calc: outline.calc(),
      structureProtected: outline.structureProtected,
    },
  };
}

/** A grid sheet as the shell holds it until its table's scan fills it. */
function sheetShell(index: number): Sheet {
  return {
    id: crypto.randomUUID(),
    name: `Sheet${index + 1}`,
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: ODS_DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: ODS_DEFAULT_ROW_HEIGHT_PT,
    merges: [],
    conditionalFormats: [],
    validations: [],
  };
}

/** The scan's placeholder ids (conditional formats and rules) replaced by
 * fresh ones — one per distinct placeholder. */
function remintScanIds(sheet: Sheet): void {
  if (sheet.conditionalFormats.length === 0) {
    return;
  }
  const ids = new Map<string, string>();
  const fresh = (id: string): string => {
    let next = ids.get(id);
    if (next === undefined) {
      next = crypto.randomUUID();
      ids.set(id, next);
    }
    return next;
  };
  sheet.conditionalFormats = sheet.conditionalFormats.map((format) => ({ ...format, id: fresh(format.id), rules: format.rules.map((rule) => ({ ...rule, id: fresh(rule.id) })) }));
}

/** Whether a chart (one read through its object, `Object N/content.xml`)
 * carries references — what the generic display refresh looks at. */
function hasChartRefs(model: SpreadsheetModel): boolean {
  return model.sheets.some((sheet) => (sheet.drawings ?? []).some((drawing) => drawing.content.kind === "chart" && drawing.content.refs !== undefined));
}

/** Phase 2 (main thread): the shell, the scan jobs and their context, and
 * the completion assembleWorkbook runs (see the module header). */
export async function readOdsShell(
  parts: Map<string, Uint8Array>,
  sources: WorkbookSources,
  title: string,
): Promise<{ shell: SpreadsheetModel; jobs: SheetScanJob[]; context: OdsScanContext }> {
  const detected = detectOds(parts);
  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw parseError("content.xml", "the spreadsheet's content part is missing");
  }
  const stylesBytes = parts.get("styles.xml");
  const stylesXml = stylesBytes ? partText(stylesBytes) : undefined;
  const sheetStyles = stylesXml !== undefined ? scanned("styles.xml", () => readOdsStyleSheet(stylesXml)) : emptyOdsStyleSheet();
  const { odsEnvelope: envelope, odsTables: tables, odsOutline: outline } = sources;
  if (sources.format !== "ods" || !envelope || !tables || !outline) {
    throw new Error("the .ods shell phase reads phase 1's sources of content.xml");
  }
  const book = scanned("content.xml", () => new OdsStyleBook(sheetStyles, readOdsContentStyles(outline.styleSlices.map(({ start, end }) => envelope.prefix.slice(start, end)), outline.declarations)));
  const parsed: ParsedSpreadsheet = {
    sheets: tables.map((_, index) => sheetShell(index)),
    styles: book.styles,
    activeSheetIndex: 0,
    dateSystem: outline.dates.dateSystem,
    odsContentEnvelope: { prefix: envelope.prefix, suffix: envelope.suffix },
    opaque: await encodeOpaque(parts),
    properties: {},
    dxfs: [],
  };
  if (outline.dates.odsNullDate !== undefined) {
    parsed.odsNullDate = outline.dates.odsNullDate;
  }
  if (detected.packageVariant) {
    parsed.packageVariant = detected.packageVariant;
  }
  if (outline.structureProtected) {
    parsed.structureProtected = true;
  }
  const namedStyles = book.namedStyles();
  if (namedStyles.length > 0) {
    parsed.namedStyles = namedStyles;
  }
  // The one assembly every open path shares (open.ts imports this module,
  // so it is loaded by now); its preview is rebuilt once the rows are in.
  const { assembleSpreadsheet } = await import("../open");
  // The completion is synchronous: the chart codec loads here when the
  // package embeds an object (a chart is only ever read from one).
  const charts = [...parts.keys()].some((name) => name !== "content.xml" && name.endsWith("/content.xml")) ? await import("../chartsOpen") : undefined;
  const shell = assembleSpreadsheet(title, "ods", parsed);
  const jobs: SheetScanJob[] = tables.map((range, sheetIndex) => ({ sheetIndex, partName: "content.xml", range: { start: range.start, end: range.end } }));
  const context: OdsScanContext = {
    format: "ods",
    tableDates: outline.tableDates,
    dates: outline.dates,
    declarations: outline.declarations,
    namespaces: outline.namespaces,
    scope: outline.scope,
    styles: book.styles.slice(),
    implied: {},
    cellIds: [],
    formatted: [],
    dxfs: [],
    book: book.scanProjection(),
  };
  const namespaces = new Map(outline.namespaces);
  setShellCompletion(shell, (model, results) => {
    const deferredOf = new Map(results.flatMap((result) => (result.deferred?.format === "ods" ? [[result.sheetIndex, result.deferred] as const] : [])));
    const media = new OdsMediaReader(parts);
    const uncachedCharts = new Set<SheetChartContent>();
    const regions: Map<string, GridRange[]>[] = [];
    model.sheets.forEach((sheet, index) => {
      remintScanIds(sheet);
      const deferred = deferredOf.get(index);
      if (deferred && deferred.drawings.length > 0) {
        // SPEC-131: under the table's geometry, sheet by sheet (the media's order).
        sheet.drawings = scanned("content.xml", () => readOdsDrawings(deferred.drawings, buildGridMetrics(sheet), { declarations: outline.declarations, styles: book, media, uncachedCharts }));
      }
      regions.push(new Map(deferred?.validations ?? []));
    });
    const { prefix, suffix } = envelope;
    scanned("content.xml", () => {
      for (const [index, autoFilter] of readOdsAutoFilters(suffix, model.sheets.map((sheet) => sheet.name))) {
        model.sheets[index]!.autoFilter = autoFilter;
      }
      const validations = readOdsContentValidations(prefix, regions);
      model.sheets.forEach((sheet, index) => {
        sheet.validations = validations.sheets[index]!;
      });
      if (validations.orphans.length > 0) {
        model.odsOrphanValidations = validations.orphans;
      }
      model.names = [
        ...readOdsWorkbookNames(suffix, namespaces),
        ...model.sheets.flatMap((sheet, index) => {
          const names = deferredOf.get(index)?.namesXml;
          return names === undefined ? [] : readOdsTableNames(names, sheet.id, namespaces);
        }),
        // SPEC-136: each table's page setup; its print ranges and header rows/columns as reserved names.
        ...readOdsPageSetups(model.sheets, prefix, stylesXml),
      ];
    });
    const settingsBytes = parts.get("settings.xml");
    const settings = settingsBytes ? scanned("settings.xml", () => readOdsSettings(partText(settingsBytes))) : undefined;
    if (settings) {
      applyViewSettings(model.sheets, settings);
    }
    const activeName = settings?.activeTable;
    const activeIndex = activeName === undefined ? 0 : model.sheets.findIndex((sheet) => sheet.name === activeName);
    model.activeSheetIndex = resolveActiveSheet(Math.max(activeIndex, 0), model.sheets);
    const meta = parts.get("meta.xml");
    const properties = meta ? parseOdfMeta(partText(meta)) : {};
    if (Object.keys(properties).length > 0) {
      model.properties = properties;
    }
    const calc: CalcSettings = { ...outline.calc };
    if (settings?.autoCalculate === false) {
      calc.manual = true;
    }
    if (Object.keys(calc).length > 0) {
      model.calc = calc;
    }
    const pictures = media.media();
    if (pictures.length > 0) {
      model.media = pictures;
    }
    // SPEC-132: the single pass refreshed its uncached charts on the parsed
    // workbook, before assembly — the same fields, seen through one.
    const view: ParsedSpreadsheet = {
      sheets: model.sheets,
      styles: model.styles,
      activeSheetIndex: model.activeSheetIndex,
      dateSystem: model.dateSystem,
      opaque: model.opaque,
      properties: model.properties ?? {},
      dxfs: model.dxfs ?? [],
      names: model.names,
    };
    if (model.odsNullDate !== undefined) {
      view.odsNullDate = model.odsNullDate;
    }
    refreshUncachedOdsCharts(view, uncachedCharts);
    model.sheets = view.sheets;
    model.preview = buildPreview(model, READ_LOCALE);
    if (charts && hasChartRefs(model)) {
      charts.refreshUncachedCharts(model);
    }
  });
  return { shell, jobs, context };
}

/** Phase 3 (DOMParser-free), stepwise: one table (see the module header).
 * Every job of one read shares `budget` and `context`. */
export function startOdsSheetScan(
  job: SheetScanJob,
  text: string,
  sources: WorkbookSources,
  context: OdsScanContext,
  emit: (chunk: RowChunk) => void,
  budget: ScanBudget,
): SheetScanSteps {
  const book = new OdsScanBook(context.book, context);
  const stylesBefore = context.styles.length;
  const dxfsBefore = context.dxfs.length;
  const namespaces = new Map(context.namespaces);
  const range = job.range ?? { start: 0, end: text.length };
  let cells = 0;
  const env: OdsTableScanEnv = {
    xml: text,
    styles: book,
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
    namespaces,
    declarations: () => context.declarations,
    dates: context.tableDates[job.sheetIndex] ?? context.dates,
  };
  const read: { table: OdsTableScan | null; end: number; tooLarge: boolean; rowsAt: number; rows: number } = { table: null, end: -1, tooLarge: false, rowsAt: 0, rows: 1 };
  const scan = startXmlScan(
    text,
    (event) => {
      const table = read.table;
      if (table === null) {
        // The range's root: the table's start tag.
        if (event.type === "open") {
          read.table = new OdsTableScan(env, event.attrs, 1, range.start, job.sheetIndex);
          if (event.selfClosing) {
            read.end = event.end;
            return "stop";
          }
        }
        return undefined;
      }
      const verdict = table.event(event);
      if (verdict === "end") {
        read.end = event.end;
        return "stop";
      }
      if (verdict === "stop") {
        read.tooLarge = true;
        return "stop";
      }
      return table.rowsRead - read.rowsAt >= read.rows ? "pause" : undefined;
    },
    { from: sources.odsOutline?.tags[job.sheetIndex] ?? range.start, to: range.end, scope: new Map(context.scope) },
  );
  let result: SheetScanResult | undefined;
  return {
    step(rows: number): boolean {
      if (result) {
        return true;
      }
      read.rowsAt = read.table?.rowsRead ?? 0;
      read.rows = Math.max(1, rows);
      if (!scanned("content.xml", () => scan.resume())) {
        return false;
      }
      const table = read.table;
      const layout = read.tooLarge || table === null ? null : table.finish();
      if (table === null || layout === null) {
        throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
      const sourceXml = text.slice(range.start, read.end);
      const fields: Partial<Sheet> = {
        name: table.name,
        visibility: table.hidden ? "hidden" : "visible",
        sourceXml,
        columns: layout.columns,
        defaultColWidthPx: layout.defaultColWidthPx,
        defaultRowHeightPt: layout.defaultRowHeightPt,
        merges: layout.merges,
      };
      if (table.tabColor !== undefined) {
        fields.tabColor = table.tabColor;
      }
      if (table.protected) {
        fields.protection = { sheet: true };
      }
      // SPEC-135: the cells' notes and links.
      Object.assign(fields, table.annotations.result());
      // SPEC-129: the table's conditional formats, as the single pass reads
      // them after its scan, the dxfs they apply joining the shared table.
      let placeholders = 0;
      const holder = { dxfs: context.dxfs };
      fields.conditionalFormats = scanned("content.xml", () =>
        readOdsTableConditionalFormats(sourceXml, { styles: book, namespaces, mint: () => `scan-${job.sheetIndex}-${(placeholders += 1)}` }, holder),
      );
      context.dxfs = holder.dxfs;
      for (const [chunk, rowsOfChunk] of chunkRows(layout.rows)) {
        emit({ sheetIndex: job.sheetIndex, chunk, rows: rowsOfChunk });
      }
      const deferred: OdsScanDeferred = { format: "ods", drawings: table.drawings, validations: [...table.validations.result()] };
      const names = odsTableNamesSpan(sourceXml);
      if (names) {
        deferred.namesXml = sourceXml.slice(names.start, names.end);
      }
      result = { sheetIndex: job.sheetIndex, fields, cells, deferred };
      const styles = context.styles.slice(stylesBefore);
      const dxfs = context.dxfs.slice(dxfsBefore);
      if (styles.length > 0 || dxfs.length > 0) {
        result.appended = { ...(styles.length > 0 ? { styles } : {}), ...(dxfs.length > 0 ? { dxfs } : {}) };
      }
      return true;
    },
    result(): SheetScanResult {
      if (!result) {
        throw new Error("startOdsSheetScan: the table is not scanned yet");
      }
      return result;
    },
  };
}
