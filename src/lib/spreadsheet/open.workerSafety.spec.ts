// SPEC-137 "The read phases": per-sheet reading never touches DOMParser. Every
// workbook fixture in e2e/fixtures/ — the .xlsx family and .ods/.ots (the
// generated spec137-* benchmark files aside) — is read phase by phase with
// DOMParser, XMLSerializer and document stubbed to undefined around exactly
// the two DOMParser-free phases — prepareWorkbookSources (1) and scanSheet
// (3) — and the assembled model must equal readSpreadsheetBytes'. The shell
// (2) and assembly (4) run on the main thread with the DOM, as the parse
// worker's client runs them. Ids are minted by a counter reset before each
// read, so the two reads' ids match.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { unzip } from "../core/zip";
import { MAX_WORKBOOK_CELLS, type SpreadsheetFormat } from "./model";
import { assembleWorkbook, prepareWorkbookSources, readSpreadsheetBytes, readWorkbookShell, scanSheet, sourcePartNames, type RowChunk, type SheetScanResult } from "./open";
import { partText } from "./parsed";

const FIXTURES = join(__dirname, "../../../e2e/fixtures");
const XLSX_FAMILY = /\.(xlsx|xlsm|xltx|xltm)$/i;
const ODS_FAMILY = /\.(ods|ots)$/i;
const fixtures = readdirSync(FIXTURES)
  .filter((name) => (XLSX_FAMILY.test(name) || ODS_FAMILY.test(name)) && !name.startsWith("spec137-"))
  .toSorted();

/** A fixture's format, by its family. */
const formatOf = (name: string): SpreadsheetFormat => (ODS_FAMILY.test(name) ? "ods" : "xlsx");

let minted = 0;

/** crypto.randomUUID as a counter from zero. */
function countIds(): void {
  minted = 0;
  vi.spyOn(globalThis.crypto, "randomUUID").mockImplementation(() => `00000000-0000-4000-8000-${String((minted += 1)).padStart(12, "0")}` as `${string}-${string}-${string}-${string}-${string}`);
}

/** Runs `read` with no DOMParser, XMLSerializer or document. */
function withoutDom<T>(read: () => T): T {
  vi.stubGlobal("DOMParser", undefined);
  vi.stubGlobal("XMLSerializer", undefined);
  vi.stubGlobal("document", undefined);
  try {
    return read();
  } finally {
    vi.unstubAllGlobals();
  }
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SPEC-137: phases 1 and 3 read every workbook fixture without DOMParser", () => {
  it("finds the fixtures of both families", () => {
    expect(fixtures.filter((name) => XLSX_FAMILY.test(name)).length).toBeGreaterThan(20);
    expect(fixtures.filter((name) => ODS_FAMILY.test(name)).length).toBeGreaterThanOrEqual(19);
  });

  for (const name of fixtures) {
    it(name, async () => {
      const format = formatOf(name);
      const bytes = new Uint8Array(readFileSync(join(FIXTURES, name)));
      countIds();
      const expected = await readSpreadsheetBytes(bytes, format, name);
      countIds();
      const parts = await unzip(bytes);
      const texts = new Map<string, string>();
      for (const part of sourcePartNames(parts, format)) {
        texts.set(part, partText(parts.get(part)!));
      }
      expect(texts.size).toBeGreaterThan(0);
      const sources = withoutDom(() => prepareWorkbookSources(texts, format));
      const { shell, jobs, context } = await readWorkbookShell(parts, sources, name);
      expect(jobs.length).toBeGreaterThan(0);
      const scanContext = structuredClone(context);
      const budget = { cellsLeft: MAX_WORKBOOK_CELLS };
      const chunks: RowChunk[] = [];
      const results: SheetScanResult[] = withoutDom(() => jobs.map((job) => scanSheet(job, texts.get(job.partName)!, sources, scanContext, (chunk) => chunks.push(chunk), budget)));
      // What crosses the worker boundary is plain data.
      expect(structuredClone({ sources, results, chunks })).toEqual({ sources, results, chunks });
      expect(assembleWorkbook(shell, results, chunks)).toEqual(expected);
    });
  }
});
