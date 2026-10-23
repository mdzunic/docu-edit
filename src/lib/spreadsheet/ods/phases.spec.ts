// SPEC-137: the .ods read phases (ods/phases.ts) on their own — phase 1's
// outline of content.xml against the single pass's slicing, phase 3's
// stepwise table scan (it pauses between rows, and small steps read what one
// step reads), and phase 3's style book (ods/scanBook.ts) answering every
// question exactly as the style book it was projected from. The whole read
// against the single pass is open.spec.ts's ("the phase split equals the old
// single pass"); without DOMParser, open.workerSafety.spec.ts's.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { unzip } from "../../core/zip";
import { MAX_WORKBOOK_CELLS } from "../model";
import { CellBudget, partText } from "../parsed";
import type { RowChunk, SheetScanResult } from "../readPhases";
import { readSpreadsheetBytes } from "../open";
import { odsSourcePartNames, prepareOdsSources, readOdsShell, startOdsSheetScan } from "./phases";
import { parseOds, readOdsContent } from "./reader";
import { OdsScanBook } from "./scanBook";
import { decodeOdsStyleName, odsStyleBookFromSource } from "./styles";

const FIXTURES = join(__dirname, "../../../../e2e/fixtures");

/** A fixture unzipped, its source part decoded and phase 1 run over it. */
async function phase1(name: string) {
  const parts = await unzip(new Uint8Array(readFileSync(join(FIXTURES, name))));
  const texts = new Map(odsSourcePartNames(parts).map((part) => [part, partText(parts.get(part)!)] as const));
  return { parts, texts, sources: prepareOdsSources(texts) };
}

const CONTENT_ROOT =
  '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">';

describe("phase 1: content.xml outlined", () => {
  it.each(["mini-book.ods", "spec116-wide.ods", "spec133-names.ods", "spec131-drawings.ods", "spec130-validation.ods"])(
    "cuts %s into the envelope and each table's sourceXml exactly as the single pass slices it",
    async (name) => {
      const { parts, texts, sources } = await phase1(name);
      const xml = texts.get("content.xml")!;
      const { prefix, suffix } = sources.odsEnvelope!;
      const slices = sources.odsTables!.map(({ start, end }) => xml.slice(start, end));
      expect(prefix + slices.join("") + suffix).toBe(xml);
      const single = await parseOds(new Uint8Array(0), parts);
      expect(single.odsContentEnvelope).toEqual({ prefix, suffix });
      expect(single.sheets.map((sheet) => sheet.sourceXml)).toEqual(slices);
      // Each table's range scan starts at its own start tag, inside its sourceXml.
      sources.odsOutline!.tags.forEach((tag, index) => {
        const range = sources.odsTables![index]!;
        expect(tag).toBeGreaterThanOrEqual(range.start);
        expect(tag).toBeLessThan(range.end);
        expect(xml.slice(range.start, tag).trim()).toBe("");
      });
    },
  );

  it("reads what the single pass reads outside the tables, a table met after a null date counting from it", () => {
    const xml = `${CONTENT_ROOT}<office:body><office:spreadsheet table:structure-protected="true"><table:calculation-settings table:case-sensitive="false"><table:null-date table:date-value="1904-01-01"/><table:iteration table:status="enable"/></table:calculation-settings>\n  <table:table table:name="A"/><!-- between --><table:table table:name="B"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const sources = prepareOdsSources(new Map([["content.xml", xml]]));
    const single = readOdsContent(xml, new CellBudget())!;
    expect(sources.odsEnvelope).toEqual({ prefix: single.prefix, suffix: single.suffix });
    expect(sources.odsTables!.map(({ start, end }) => xml.slice(start, end))).toEqual(single.sheets.map((sheet) => sheet.sourceXml));
    expect(sources.odsOutline).toMatchObject({ dates: single.dates, calc: single.calc, structureProtected: true, tableDates: [single.dates, single.dates] });
    expect(sources.odsOutline!.tags.map((tag) => xml.slice(tag, tag + 13))).toEqual(["<table:table ", "<table:table "]);
  });

  it("reads a self-closing table, and whatever stands between two tables, as the single pass does — a later null date counting for the later table", async () => {
    const parts = (await phase1("mini-book.ods")).parts;
    const content =
      `${CONTENT_ROOT}<office:body><office:spreadsheet>\n  <table:table table:name="A"/>\n  <!-- between -->\n  <table:null-date table:date-value="1904-01-01"/><text:p>stray</text:p> stray text ` +
      '<table:table table:name="B"><table:table-row><table:table-cell office:value-type="date" office:date-value="1904-01-02"><text:p>x</text:p></table:table-cell></table:table-row></table:table>\n</office:spreadsheet></office:body></office:document-content>';
    parts.set("content.xml", new TextEncoder().encode(content));
    let minted = 0;
    const ids = vi.spyOn(crypto, "randomUUID").mockImplementation(() => `00000000-0000-4000-8000-${String((minted += 1)).padStart(12, "0")}`);
    try {
      const { assembleSpreadsheet } = await import("../open");
      const single = assembleSpreadsheet("Book", "ods", await parseOds(new Uint8Array(0), parts));
      minted = 0;
      const phased = await readSpreadsheetBytes(new Uint8Array(0), "ods", "Book", parts);
      expect(phased).toEqual(single);
      expect(phased.dateSystem).toBe("1904");
      expect(phased.sheets.map((sheet) => [sheet.name, sheet.rows[0]?.cells[0]?.value])).toEqual([
        ["A", undefined],
        ["B", 1],
      ]);
      // B's sourceXml starts where A ends: what stands between them is B's.
      expect(phased.sheets[1]!.sourceXml).toMatch(/^\n {2}<!-- between -->\n {2}<table:null-date [^>]*\/><text:p>stray<\/text:p> stray text <table:table table:name="B">/);
    } finally {
      ids.mockRestore();
    }
  });

  it("refuses a content.xml with no table, and a malformed one, as the single pass does", () => {
    const empty = `${CONTENT_ROOT}<office:body><office:spreadsheet/></office:body></office:document-content>`;
    const malformed = `${CONTENT_ROOT}<office:body><office:spreadsheet><table:table><table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    for (const xml of [empty, malformed]) {
      const single = (() => {
        try {
          return readOdsContent(xml, new CellBudget());
        } catch (error) {
          return error;
        }
      })();
      expect(single).toBeInstanceOf(Error);
      expect(() => prepareOdsSources(new Map([["content.xml", xml]]))).toThrow((single as Error).message);
    }
  });

  it("names no part, and reads none, for a package the shell refuses", async () => {
    const xlsx = await unzip(new Uint8Array(readFileSync(join(FIXTURES, "mini-book.xlsx"))));
    expect(odsSourcePartNames(xlsx)).toEqual([]);
    const { parts } = await phase1("mini-book.ods");
    parts.delete("content.xml");
    expect(odsSourcePartNames(parts)).toEqual([]);
    expect(prepareOdsSources(new Map())).toEqual({ format: "ods" });
    // The shell then refuses it as the single pass does.
    await expect(readSpreadsheetBytes(new Uint8Array(0), "ods", "Book", parts)).rejects.toThrow("the spreadsheet's content part is missing");
    await expect(readSpreadsheetBytes(new Uint8Array(0), "ods", "Book", xlsx)).rejects.toThrow("This file isn't an OpenDocument spreadsheet we can open");
    await expect(readSpreadsheetBytes(new Uint8Array([1, 2, 3]), "ods", "Book")).rejects.toThrow("This file isn't an OpenDocument spreadsheet we can open");
  });
});

describe("phase 3: a table scanned stepwise", () => {
  /** Every job of a fixture scanned in steps of `rows` rows. */
  async function scanInSteps(name: string, rows: number) {
    const { parts, texts, sources } = await phase1(name);
    const { jobs, context } = await readOdsShell(parts, sources, name);
    const scanContext = structuredClone(context);
    const budget = { cellsLeft: MAX_WORKBOOK_CELLS };
    const chunks: RowChunk[] = [];
    const results: SheetScanResult[] = [];
    const pauses: number[] = [];
    for (const job of jobs) {
      const scan = startOdsSheetScan(job, texts.get("content.xml")!, sources, scanContext, (chunk) => chunks.push(chunk), budget);
      let paused = 0;
      while (!scan.step(rows)) {
        paused += 1;
      }
      pauses.push(paused);
      results.push(scan.result());
    }
    const rowElements = sources.odsTables!.map(({ start, end }) => texts.get("content.xml")!.slice(start, end).match(/<table:table-row[\s/>]/g)?.length ?? 0);
    return { chunks, results, pauses, rowElements, styles: scanContext.styles, dxfs: scanContext.dxfs, cellsLeft: budget.cellsLeft };
  }

  it.each(["mini-book.ods", "spec116-wide.ods", "spec117-styles.ods", "spec135-annotations.ods"])("pauses after every `rows` row elements, and small steps read what one step reads (%s)", async (name) => {
    const whole = await scanInSteps(name, Number.POSITIVE_INFINITY);
    expect(whole.pauses.every((paused) => paused === 0)).toBe(true);
    for (const rows of [1, 7, 64]) {
      const stepped = await scanInSteps(name, rows);
      // A step ends right after its `rows`-th row: floor(row elements / rows) pauses per table.
      expect(stepped.pauses).toEqual(stepped.rowElements.map((count) => Math.floor(count / rows)));
      expect(stepped.chunks).toEqual(whole.chunks);
      expect(stepped.results).toEqual(whole.results);
      expect(stepped.styles).toEqual(whole.styles);
      expect(stepped.dxfs).toEqual(whole.dxfs);
      expect(stepped.cellsLeft).toBe(whole.cellsLeft);
    }
    expect(Math.max(...(await scanInSteps(name, 1)).pauses)).toBeGreaterThan(0);
  });

  it("emits a long table's rows as SHEET_ROWS_CHUNK bands only once the table is read", async () => {
    const { parts, texts, sources } = await phase1("mini-book.ods");
    const { jobs, context } = await readOdsShell(parts, sources, "mini-book");
    const rowsOf = (job: (typeof jobs)[number]): number => texts.get("content.xml")!.slice(job.range!.start, job.range!.end).match(/<table:table-row[\s/>]/g)?.length ?? 0;
    const data = jobs.toSorted((a, b) => rowsOf(b) - rowsOf(a))[0]!;
    expect(rowsOf(data)).toBeGreaterThan(5);
    const chunks: RowChunk[] = [];
    const scan = startOdsSheetScan(data, texts.get("content.xml")!, sources, structuredClone(context), (chunk) => chunks.push(chunk), { cellsLeft: MAX_WORKBOOK_CELLS });
    expect(scan.step(5)).toBe(false);
    expect(chunks).toEqual([]);
    while (!scan.step(5)) {
      // the rest of the table
    }
    expect(chunks.length).toBeGreaterThan(0);
    expect(chunks.every((chunk) => chunk.sheetIndex === data.sheetIndex)).toBe(true);
    expect(scan.result().cells).toBeGreaterThan(0);
  });
});

describe("phase 3's style book answers as the style book it was projected from", () => {
  it.each(["spec117-styles.ods", "spec118-formats.ods", "spec123-styles.ods", "spec129-cf.ods", "spec129-legacy.ods", "spec119-formulas.ods"])("%s", async (name) => {
    const { parts, sources } = await phase1(name);
    const stylesBytes = parts.get("styles.xml");
    const real = odsStyleBookFromSource(sources.odsEnvelope!.prefix, stylesBytes ? partText(stylesBytes) : undefined);
    // Projected and cloned, as the context crosses to the worker.
    const projection = structuredClone(real.scanProjection());
    expect(real.styles).toHaveLength(1); // the projection adds nothing
    const scan = new OdsScanBook(projection, { styles: structuredClone(real.styles), implied: {}, cellIds: [], formatted: [] });

    // The cell style table, grown in the same first-reference order — every
    // defined name and two no part defines, each met twice — then every
    // implied format over every entry.
    const names = [...projection.cellStyles.map(([each]) => each), "Default", "no-such-style", "Another missing one"];
    expect(projection.cellStyles.length).toBeGreaterThan(0);
    for (const each of [...names, ...names.toReversed()]) {
      expect(scan.cellStyleId(each)).toBe(real.cellStyleId(each));
    }
    const grown = real.styles.length;
    for (let id = 0; id < grown; id += 1) {
      for (const code of ["yyyy-mm-dd", "[hh]:mm:ss", "0.00%"]) {
        expect(scan.impliedStyleId(id, code)).toBe(real.impliedStyleId(id, code));
      }
    }
    expect(scan.styles).toEqual(real.styles);

    // Every other lookup, for every name the projection holds and one it does not.
    const missing = "no-such-style";
    for (const each of [...projection.textStyles.map(([text]) => text), missing]) {
      expect(scan.textFont(each)).toEqual(real.textFont(each));
      expect(scan.textStyleModelled(each)).toBe(real.textStyleModelled(each));
    }
    for (const each of [...projection.tableStyles.map(([table]) => table), missing, undefined]) {
      expect(scan.tableStyle(each)).toEqual(real.tableStyle(each));
    }
    for (const each of [...projection.columnWidths.map(([column]) => column), missing, undefined]) {
      expect(scan.columnWidthPx(each)).toBe(real.columnWidthPx(each));
    }
    for (const each of [...projection.rowHeights.map(([row]) => row), missing, undefined]) {
      expect(scan.rowHeight(each)).toEqual(real.rowHeight(each));
    }
    expect(scan.mappedCellStyles()).toEqual(real.mappedCellStyles());
    for (const each of [...names, ...real.mappedCellStyles()]) {
      expect(scan.cellStyleMaps(each)).toEqual(real.cellStyleMaps(each));
    }
    const common = projection.conditional.map(([each]) => each);
    for (const each of [...common, ...common.map(decodeOdsStyleName), ...projection.displayNames.map(([display]) => display), missing]) {
      expect(scan.conditionalStyle(each)).toEqual(real.conditionalStyle(each));
    }
  });

  it("grows the read's style and dxf tables to the single pass's (the phased model's tables)", async () => {
    for (const name of ["spec117-styles.ods", "spec118-formats.ods", "spec129-cf.ods", "spec129-legacy.ods"]) {
      const bytes = new Uint8Array(readFileSync(join(FIXTURES, name)));
      const single = await parseOds(bytes);
      const phased = await readSpreadsheetBytes(bytes, "ods", name);
      expect(phased.styles.length).toBeGreaterThan(1);
      expect(phased.styles).toEqual(single.styles);
      expect(phased.dxfs).toEqual(single.dxfs);
    }
  });
});
