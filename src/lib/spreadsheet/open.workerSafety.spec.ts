// SPEC-137 "The read phases": per-sheet reading never touches DOMParser. Every
// workbook fixture of the .xlsx family in e2e/fixtures/ (the generated
// spec137-* benchmark files aside) is read phase by phase with DOMParser,
// XMLSerializer and document stubbed to undefined around exactly the two
// DOMParser-free phases — prepareWorkbookSources (1) and scanSheet (3) — and
// the assembled model must equal readSpreadsheetBytes'. The shell (2) and
// assembly (4) run on the main thread with the DOM, as the parse worker's
// client runs them. Ids are minted by a counter reset before each read, so
// the two reads' ids match. (.ods keeps the single pass — its phases are
// not built — so its fixtures are not listed here.)
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { unzip } from "../core/zip";
import { MAX_WORKBOOK_CELLS } from "./model";
import { assembleWorkbook, prepareWorkbookSources, readSpreadsheetBytes, readWorkbookShell, scanSheet, sourcePartNames, type RowChunk, type SheetScanResult } from "./open";
import { partText } from "./parsed";

const FIXTURES = join(__dirname, "../../../e2e/fixtures");
const XLSX_FAMILY = /\.(xlsx|xlsm|xltx|xltm)$/i;
const fixtures = readdirSync(FIXTURES)
  .filter((name) => XLSX_FAMILY.test(name) && !name.startsWith("spec137-"))
  .toSorted();

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

describe("SPEC-137: phases 1 and 3 read every .xlsx-family fixture without DOMParser", () => {
  it("finds the fixtures", () => {
    expect(fixtures.length).toBeGreaterThan(20);
  });

  for (const name of fixtures) {
    it(name, async () => {
      const bytes = new Uint8Array(readFileSync(join(FIXTURES, name)));
      countIds();
      const expected = await readSpreadsheetBytes(bytes, "xlsx", name);
      countIds();
      const parts = await unzip(bytes);
      const texts = new Map<string, string>();
      for (const part of sourcePartNames(parts, "xlsx")) {
        texts.set(part, partText(parts.get(part)!));
      }
      const sources = withoutDom(() => prepareWorkbookSources(texts, "xlsx"));
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
