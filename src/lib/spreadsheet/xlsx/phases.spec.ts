// SPEC-137's .xlsx phases in isolation: the stepwise worksheet scan (pauses
// once per step of rows, and equals the one-step scan), the lean cells both
// paths allocate (AC "Readers allocate lean cells"), and the scanner-read
// indexed palette phase 1 resolves shared-string colours against, which must
// equal readXlsxStyles' for every fixture.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { MAX_WORKBOOK_CELLS, type Row } from "../model";
import { readSpreadsheetBytes } from "../open";
import { CellBudget, partText } from "../parsed";
import type { RowChunk } from "../readPhases";
import { startXlsxSheetScan, type XlsxScanContext } from "./phases";
import { readWorksheet, startWorksheetRead, type SharedString } from "./sheetXml";
import { readXlsxIndexedColors, readXlsxStyles } from "./styles";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const FIXTURES = join(__dirname, "../../../../e2e/fixtures");
const fixtures = readdirSync(FIXTURES)
  .filter((name) => /\.(xlsx|xlsm|xltx|xltm)$/i.test(name) && !name.startsWith("spec137-"))
  .toSorted();

/** A worksheet of `count` rows: a number, a shared string, a date. */
function bigSheet(count: number): string {
  let rows = "";
  for (let r = 1; r <= count; r += 1) {
    rows += `<row r="${r}" spans="1:3"><c r="A${r}"><v>${r}</v></c><c r="B${r}" t="s"><v>${r % 2}</v></c><c r="C${r}"><v>${r}.25</v></c></row>`;
  }
  return `<worksheet xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C${count}"/><sheetData>${rows}</sheetData><drawing r:id="rId1"/></worksheet>`;
}

const STRINGS: SharedString[] = [{ text: "even" }, { text: "odd", runs: [{ text: "odd", font: { bold: true } }] }];

function context(): XlsxScanContext {
  return { format: "xlsx", dateSystem: "1900", mdw: 7, styles: [{}], implied: {}, dxfs: [], sheetNames: ["Data"] };
}

describe("the stepwise worksheet scan", () => {
  const xml = bigSheet(5_000);

  it("startWorksheetRead pauses once per 2,000 rows and reads what readWorksheet reads", () => {
    const options = () => ({ sharedStrings: STRINGS, dateSystem: "1900" as const, budget: new CellBudget(), mdw: 7 });
    const read = startWorksheetRead(xml, "Data", options());
    const steps: boolean[] = [];
    for (let done = false; !done; ) {
      done = read.step(2_000);
      steps.push(done);
    }
    expect(steps).toEqual([false, false, true]);
    expect(read.content()).toEqual(readWorksheet(xml, "Data", options()));
    expect(read.withoutCells()).toBe(`<worksheet xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C5000"/><sheetData></sheetData><drawing r:id="rId1"/></worksheet>`);
  });

  it("startXlsxSheetScan with step(2000) pauses twice, emits 512-row chunks and equals the one-step scan", () => {
    const scan = (rows: number) => {
      const chunks: RowChunk[] = [];
      const budget = { cellsLeft: MAX_WORKBOOK_CELLS };
      const steps = startXlsxSheetScan({ sheetIndex: 0, partName: "xl/worksheets/sheet1.xml" }, xml, { format: "xlsx", sharedStrings: STRINGS }, context(), (chunk) => chunks.push(chunk), budget);
      let pauses = 0;
      while (!steps.step(rows)) {
        pauses += 1;
      }
      return { pauses, chunks, result: steps.result(), left: budget.cellsLeft };
    };
    const stepped = scan(2_000);
    const whole = scan(Number.POSITIVE_INFINITY);
    expect(stepped.pauses).toBe(2);
    expect(whole.pauses).toBe(0);
    expect(stepped.chunks).toEqual(whole.chunks);
    expect(stepped.result).toEqual(whole.result);
    expect(stepped.chunks.map((chunk) => chunk.chunk)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(stepped.chunks.every((chunk) => chunk.rows.every((row) => Math.floor(row.index / 512) === chunk.chunk))).toBe(true);
    expect(stepped.result.cells).toBe(15_000);
    expect(stepped.left).toBe(MAX_WORKBOOK_CELLS - 15_000);
    expect(stepped.result.deferred).toEqual({ format: "xlsx", worksheetXml: expect.stringContaining("<sheetData></sheetData><drawing") });
  });

  it("refuses past the ceiling counted across every job of one read", () => {
    const budget = { cellsLeft: 14_999 };
    const steps = startXlsxSheetScan({ sheetIndex: 0, partName: "xl/worksheets/sheet1.xml" }, xml, { format: "xlsx", sharedStrings: STRINGS }, context(), () => undefined, budget);
    expect(() => steps.step(Number.POSITIVE_INFINITY)).toThrow("more than 2,000,000 filled cells");
  });

  it("a scanner failure is the reader's parse error, naming the part", () => {
    const steps = startXlsxSheetScan({ sheetIndex: 0, partName: "xl/worksheets/sheet9.xml" }, `<worksheet xmlns="${MAIN}"><sheetData><row>`, { format: "xlsx", sharedStrings: [] }, context(), () => undefined, { cellsLeft: 10 });
    expect(() => steps.step(Number.POSITIVE_INFINITY)).toThrow(/^This workbook can't be opened: the part is not well-formed XML — .*\(xl\/worksheets\/sheet9\.xml\)$/);
  });
});

describe("lean cells (SPEC-137)", () => {
  it("a plain number is exactly { col, value }, integral stays integral, optional fields only when present, rows index-first", () => {
    const xml = `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>1</v></c><c r="D1"><v>1.5E+3</v></c><c r="E1" s="1"><v>2.5</v></c></row></sheetData></worksheet>`;
    const content = readWorksheet(xml, "S", { sharedStrings: STRINGS, dateSystem: "1900", budget: new CellBudget(), mdw: 7 })!;
    const row: Row = content.rows[0]!;
    expect(Object.keys(row).slice(0, 2)).toEqual(["index", "cells"]);
    const [a, b, c, d, e] = row.cells;
    expect(Object.keys(a!)).toEqual(["col", "value"]);
    expect(a!.value).toBe(42);
    expect(Number.isInteger(d!.value)).toBe(true);
    expect(Object.keys(d!)).toEqual(["col", "value"]);
    expect(Object.keys(e!)).toEqual(["col", "value", "styleId"]);
    expect(Object.keys(b!).slice(0, 2)).toEqual(["col", "value"]);
    // Every cell reading one shared string holds the table's own value and runs.
    expect(b!.value).toBe(STRINGS[1]!.text);
    expect(b!.runs).toBe(STRINGS[1]!.runs);
    expect(c!.runs).toBe(b!.runs);
  });

  it("no cell of any fixture holds an undefined-valued key, and every plain number cell is { col, value }", async () => {
    for (const name of fixtures) {
      const model = await readSpreadsheetBytes(new Uint8Array(readFileSync(join(FIXTURES, name))), "xlsx", name);
      for (const sheet of model.sheets) {
        for (const row of sheet.rows) {
          expect(Object.keys(row).slice(0, 2), name).toEqual(["index", "cells"]);
          for (const cell of row.cells) {
            expect(Object.keys(cell).slice(0, 2), name).toEqual(["col", "value"]);
            expect(Object.values(cell).includes(undefined), name).toBe(false);
          }
        }
      }
    }
  });
});

describe("readXlsxIndexedColors — phase 1's palette", () => {
  it("equals readXlsxStyles' palette for every fixture", async () => {
    for (const name of fixtures) {
      const parts = await unzip(new Uint8Array(readFileSync(join(FIXTURES, name))));
      for (const [part, bytes] of parts) {
        if (part.endsWith("styles.xml")) {
          const text = partText(bytes);
          expect(readXlsxIndexedColors(text), `${name} ${part}`).toEqual(readXlsxStyles(text, undefined).indexedColors);
        }
      }
    }
  });

  it("reads the first colors' first indexedColors as readXlsxStyles does: up to 64, an unreadable entry the legacy colour", () => {
    const entries = Array.from({ length: 70 }, (_, index) => (index === 1 ? `<rgbColor rgb="nonsense"/>` : `<rgbColor rgb="FF${String(index).padStart(6, "0")}"/>`)).join("");
    const cases = [
      `<styleSheet xmlns="${MAIN}"><colors><indexedColors>${entries}</indexedColors></colors></styleSheet>`,
      `<styleSheet xmlns="${MAIN}"><colors><mruColors/></colors><colors><indexedColors>${entries}</indexedColors></colors></styleSheet>`,
      `<x:styleSheet xmlns:x="${MAIN}"><x:colors><x:indexedColors><x:rgbColor rgb="FF00FF00"/></x:indexedColors></x:colors></x:styleSheet>`,
      `<styleSheet xmlns="${MAIN}"><colors><indexedColors/></colors></styleSheet>`,
      `<styleSheet xmlns="urn:other"><colors><indexedColors><rgbColor rgb="FF00FF00"/></indexedColors></colors></styleSheet>`,
      `<styleSheet xmlns="${MAIN}"><colors><indexedColors><rgbColor rgb="FF00FF00"/></indexedColors></colors>`,
    ];
    for (const xml of cases) {
      expect(readXlsxIndexedColors(xml), xml.slice(0, 80)).toEqual(readXlsxStyles(xml, undefined).indexedColors);
    }
    expect(readXlsxIndexedColors(cases[0])).toHaveLength(64);
    expect(readXlsxIndexedColors(cases[0])![1]).toBe("#ffffff");
  });
});
