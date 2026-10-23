// SPEC-137: the parse client (AC-14..AC-20) over a fake phase layer. The
// format's phases (xlsx/phases.ts) are mocked with a toy package — text
// parts, one row per line — and everything above them is real: open.ts's
// dispatch, readPhased and assembleWorkbook, parseSession.ts, workerRpc and
// the client. So the protocol, the inline fallback and its yields, the
// progress, the retry and the error marshalling are pinned whatever the real
// readers do; parseSession.spec.ts runs the real readers over every fixture.
//
// With a Worker, the global is an in-process stand-in: each message is
// structured-cloned (with its transfer list) and answered on a later task
// by a real parseSession — or, for the failure cases, not answered.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { XmlScanError } from "../core/xmlScanner";
import { unzip } from "../core/zip";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../spreadsheet/errors";
import type { Row, SpreadsheetModel } from "../spreadsheet/model";
import { readPhased, readSpreadsheetBytes } from "../spreadsheet/open";
import { createParseSessions, type SheetParseRequest, type SheetParseResponse } from "../spreadsheet/parseSession";
import { setShellCompletion, type RowChunk, type ScanBudget, type SheetScanJob, type SheetScanResult, type SheetScanSteps, type WorkbookSources } from "../spreadsheet/readPhases";
import { SPREADSHEET_READER_VERSION } from "../spreadsheet/readerVersion";
import { chunkOf } from "../spreadsheet/rowChunks";
import { readSpreadsheetOffThread, SCAN_YIELD_ROWS } from "./spreadsheetParseClient";
import { isWorkerSupported } from "./workerRpc";

/** What happened, in order: posts, yields, progress. */
const trace = vi.hoisted(() => ({ events: [] as string[] }));

vi.mock("../core/idleScheduler", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../core/idleScheduler")>()),
  yieldToMainThread: () => {
    trace.events.push("yield");
    return Promise.resolve();
  },
}));

vi.mock("../spreadsheet/xlsx/phases", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../spreadsheet/xlsx/phases")>()),
  xlsxSourcePartNames: (parts: Map<string, Uint8Array>) => toyPhases.sourcePartNames(parts),
  prepareXlsxSources: (texts: Map<string, string>) => toyPhases.prepare(texts),
  readXlsxShell: (parts: Map<string, Uint8Array>, sources: WorkbookSources, title: string) => toyPhases.shell(parts, sources, title),
  startXlsxSheetScan: (job: SheetScanJob, text: string, sources: WorkbookSources, context: ToyContext, emit: (chunk: RowChunk) => void, budget: ScanBudget) =>
    toyPhases.startScan(job, text, sources, context, emit, budget),
}));

// --- the toy package ---------------------------------------------------------------------------
// "strings.txt": the shared strings, one per line. "sheets/<name>.txt": one
// sheet, one line per row — `<index>:<value>,<value>…`, `#n` naming shared
// string n — or a command: `!take n` counts n cells it does not store,
// `!xml n` is malformed markup at offset n, `!bug` a reader bug. "broken.txt"
// makes the shell phase refuse the package. Any other part is not a source.

const SHEET_PART = /^sheets\/(.+)\.txt$/;

interface ToyContext {
  format: "xlsx";
  sheetNames: string[];
}

const toyPhases = {
  /** Shell phases run (phase 2, main thread only). */
  shells: 0,
  /** Every step() a scan was asked for, in rows. */
  steps: [] as number[],

  sourcePartNames(parts: Map<string, Uint8Array>): string[] {
    return [...parts.keys()].filter((name) => name === "strings.txt" || SHEET_PART.test(name));
  },

  prepare(texts: Map<string, string>): WorkbookSources {
    const strings = texts.get("strings.txt") ?? "";
    return { format: "xlsx", sharedStrings: strings === "" ? [] : strings.split("\n").map((text) => ({ text })) };
  },

  async shell(parts: Map<string, Uint8Array>, _sources: WorkbookSources, title: string): Promise<{ shell: SpreadsheetModel; jobs: SheetScanJob[]; context: ToyContext }> {
    toyPhases.shells += 1;
    if (parts.has("broken.txt")) {
      throw new SpreadsheetOpenError("This file isn't an Excel workbook we can open");
    }
    const partNames = [...parts.keys()].filter((name) => SHEET_PART.test(name));
    const sheetNames = partNames.map((partName) => SHEET_PART.exec(partName)![1]!);
    const shell: SpreadsheetModel = {
      kind: "spreadsheet",
      id: crypto.randomUUID(),
      title,
      format: "xlsx",
      readerVersion: SPREADSHEET_READER_VERSION,
      sheets: partNames.map((partName, index) => ({
        id: crypto.randomUUID(),
        name: sheetNames[index]!,
        visibility: "visible",
        kind: "grid",
        rows: [],
        partName,
        columns: [],
        defaultColWidthPx: 64,
        defaultRowHeightPt: 15,
        merges: [],
        conditionalFormats: [],
        validations: [],
      })),
      activeSheetIndex: 0,
      dateSystem: "1900",
      opaque: [],
      styles: [{}],
      defaultStyleId: 0,
      names: [],
      dxfs: [],
    };
    // The main thread's completion: here, the last sheet that holds rows becomes active.
    setShellCompletion(shell, (model, results) => {
      model.activeSheetIndex = results.findLast((result) => result.cells > 0)?.sheetIndex ?? 0;
    });
    return { shell, jobs: partNames.map((partName, sheetIndex) => ({ sheetIndex, partName })), context: { format: "xlsx", sheetNames } };
  },

  startScan(job: SheetScanJob, text: string, sources: WorkbookSources, context: ToyContext, emit: (chunk: RowChunk) => void, budget: ScanBudget): SheetScanSteps {
    const lines = text === "" ? [] : text.split("\n");
    let at = 0;
    let cells = 0;
    let band: Row[] = [];
    let bandIndex = -1;
    const take = (count: number): void => {
      cells += count;
      budget.cellsLeft -= count;
      if (budget.cellsLeft < 0) {
        throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
    };
    const flush = (): void => {
      if (band.length > 0) {
        emit({ sheetIndex: job.sheetIndex, chunk: bandIndex, rows: band });
        band = [];
      }
    };
    const command = (line: string): void => {
      const [name, argument] = line.slice(1).split(" ");
      if (name === "take") {
        take(Number(argument));
      } else if (name === "xml") {
        throw new XmlScanError("Closing tag </row> has no open element", Number(argument));
      } else {
        throw new TypeError(`${context.sheetNames[job.sheetIndex]}'s rows are not iterable`);
      }
    };
    return {
      step(rows) {
        toyPhases.steps.push(rows);
        for (const end = Math.min(lines.length, at + rows); at < end; at += 1) {
          const line = lines[at]!;
          if (line.startsWith("!")) {
            command(line);
            continue;
          }
          const [index, values] = line.split(":");
          const row: Row = {
            index: Number(index),
            cells: values!.split(",").map((value, col) => ({ col, value: value.startsWith("#") ? sources.sharedStrings![Number(value.slice(1))]!.text : Number(value) })),
          };
          take(row.cells.length);
          if (chunkOf(row.index) !== bandIndex) {
            flush();
            bandIndex = chunkOf(row.index);
          }
          band.push(row);
        }
        if (at < lines.length) {
          return false;
        }
        flush();
        return true;
      },
      result(): SheetScanResult {
        return { sheetIndex: job.sheetIndex, fields: { defaultRowHeightPt: 20 }, cells };
      },
    };
  },
};

const encoder = new TextEncoder();

/** A toy package: the given sheets and shared strings, and a part that is
 * no source. */
function toyWorkbook(sheets: Record<string, string[]>, strings: string[] = []): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>([["docProps/app.txt", encoder.encode("not a source part")]]);
  parts.set("strings.txt", encoder.encode(strings.join("\n")));
  for (const [name, lines] of Object.entries(sheets)) {
    parts.set(`sheets/${name}.txt`, encoder.encode(lines.join("\n")));
  }
  return parts;
}

function rowLines(count: number, line: (index: number) => string, from = 0): string[] {
  return Array.from({ length: count }, (_, at) => `${from + at}:${line(from + at)}`);
}

/** Data: 1,200 rows (bands 0–2); Sparse: rows 0 and 5,000 (bands 0 and 9);
 * Empty: none — five chunks. */
function book(): Map<string, Uint8Array> {
  return toyWorkbook(
    {
      Data: rowLines(1_200, (index) => `#${index % 3},${index},${index * 2}`),
      Sparse: ["0:#0", "5000:1.5"],
      Empty: [],
    },
    ["alpha", "beta", "gamma"],
  );
}

// --- the in-process Worker ----------------------------------------------------------------------

type Posted = SheetParseRequest & { id: number };
type Listener = (event: { data?: unknown; message?: string; error?: unknown }) => void;

/** How a worker misbehaves for one request: answered → false (answer it as
 * usual), true (the misbehaviour replaced the answer). */
type Misbehaviour = (worker: InProcessWorker, request: Posted) => boolean;

class InProcessWorker {
  static instances: InProcessWorker[] = [];
  static misbehave: Misbehaviour | undefined;

  readonly url: string;
  readonly options: WorkerOptions | undefined;
  readonly received: Posted[] = [];
  readonly replies: SheetParseResponse[] = [];
  readonly transferred: ArrayBuffer[] = [];
  terminated = false;
  private readonly sessions = createParseSessions();
  private readonly listeners = new Map<string, Listener[]>();

  constructor(url: URL | string, options?: WorkerOptions) {
    this.url = String(url);
    this.options = options;
    InProcessWorker.instances.push(this);
  }

  addEventListener(type: string, listener: Listener): void {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  postMessage(message: Posted, transfer: Transferable[] = []): void {
    trace.events.push(message.op === "chunk" ? `post chunk ${message.index}` : `post ${message.op}`);
    this.transferred.push(...(transfer as ArrayBuffer[]));
    const request = structuredClone(message, { transfer });
    this.received.push(request);
    setTimeout(() => {
      if (this.terminated || InProcessWorker.misbehave?.(this, request)) {
        return;
      }
      const reply = this.sessions.handle(request);
      this.replies.push(reply);
      this.emit("message", { data: structuredClone({ id: request.id, ...reply }) });
    }, 0);
  }

  terminate(): void {
    this.terminated = true;
  }

  emit(type: string, event: { data?: unknown; message?: string; error?: unknown }): void {
    for (const listener of this.listeners.get(type) ?? []) {
      listener(event);
    }
  }
}

// --- helpers -----------------------------------------------------------------------------------

let minted = 0;

beforeEach(() => {
  trace.events = [];
  toyPhases.shells = 0;
  toyPhases.steps = [];
  minted = 0;
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    minted += 1;
    return `00000000-0000-4000-8000-${String(minted).padStart(12, "0")}`;
  });
  InProcessWorker.instances = [];
  InProcessWorker.misbehave = undefined;
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** The inline phases' model of `parts` (ids minted from 1, the counters left
 * as they were). */
async function inlineModel(parts: Map<string, Uint8Array>): Promise<SpreadsheetModel> {
  const counted = { minted, shells: toyPhases.shells, steps: toyPhases.steps, events: trace.events };
  minted = 0;
  const model = await readPhased(parts, "xlsx", "Book");
  ({ minted, shells: toyPhases.shells, steps: toyPhases.steps, events: trace.events } = counted);
  return model;
}

/** A read and the progress it reported. */
async function readOffThread(parts: Map<string, Uint8Array>): Promise<{ model: SpreadsheetModel; progress: number[] }> {
  const progress: number[] = [];
  const model = await readSpreadsheetOffThread(parts, "xlsx", "Book", {
    onProgress: (fraction) => {
      progress.push(fraction);
      trace.events.push(`progress ${fraction}`);
    },
  });
  return { model, progress };
}

/** The model with its minted ids blanked: a retried read mints more of them. */
function withoutIds(model: SpreadsheetModel): unknown {
  return { ...model, id: "", sheets: model.sheets.map((sheet) => ({ ...sheet, id: "" })) };
}

// --- the tests -----------------------------------------------------------------------------------

describe("readSpreadsheetOffThread without a Worker (Vitest/jsdom)", () => {
  it("runs the same phases inline, handing the main thread back every SCAN_YIELD_ROWS rows", async () => {
    expect(isWorkerSupported()).toBe(false);
    expect(SCAN_YIELD_ROWS).toBe(2_000);
    const parts = toyWorkbook({ Data: rowLines(5_000, (index) => `${index},${index + 1}`), Small: rowLines(10, (index) => `${index}`) });
    const { model, progress } = await readOffThread(parts);
    // Data: 2,000 + 2,000 + 1,000 rows, a yield after each step that did
    // not finish the sheet; Small: one step.
    expect(toyPhases.steps).toEqual([2_000, 2_000, 2_000, 2_000]);
    expect(trace.events.filter((event) => event === "yield")).toHaveLength(2);
    expect(trace.events.slice(0, 4)).toEqual(["yield", "yield", "progress 0.5", "progress 1"]);
    expect(progress).toEqual([0.5, 1]);
    expect(model.sheets.map((sheet) => sheet.rows.length)).toEqual([5_000, 10]);
    expect(model).toEqual(await inlineModel(parts));
  });

  it("leaves a delimited-text file to readSpreadsheetBytes, which refuses it as no package", async () => {
    await expect(readSpreadsheetOffThread(new Map(), "csv", "Book")).rejects.toThrow("a .csv workbook is read by readCsvBytes, not as a package");
  });
});

describe("readSpreadsheetOffThread with a Worker: the parse protocol", () => {
  beforeEach(() => {
    vi.stubGlobal("Worker", InProcessWorker);
  });

  it("starts one module worker for the open and terminates it when the open settles", async () => {
    expect(isWorkerSupported()).toBe(true);
    await readOffThread(book());
    expect(InProcessWorker.instances).toHaveLength(1);
    const [worker] = InProcessWorker.instances;
    expect(worker!.url).toMatch(/\/workers\/spreadsheetParse\.worker\.ts/);
    expect(worker!.options).toEqual({ type: "module" });
    expect(worker!.terminated).toBe(true);
  });

  it("posts copies of the source parts only, so the parts keep their own bytes", async () => {
    const parts = book();
    const before = new Map([...parts].map(([name, bytes]) => [name, bytes.slice()]));
    await readOffThread(parts);
    const [worker] = InProcessWorker.instances;
    const prepare = worker!.received[0] as Extract<Posted, { op: "prepare" }>;
    expect(prepare.op).toBe("prepare");
    expect(prepare.format).toBe("xlsx");
    expect(prepare.sources.map((source) => source.name)).toEqual(["strings.txt", "sheets/Data.txt", "sheets/Sparse.txt", "sheets/Empty.txt"]);
    for (const source of prepare.sources) {
      expect(source.bytes).toEqual(before.get(source.name));
    }
    // The copies were transferred, never the parts' own buffers…
    expect(worker!.transferred).toHaveLength(prepare.sources.length);
    for (const bytes of parts.values()) {
      expect(worker!.transferred).not.toContain(bytes.buffer);
    }
    // …which are whole.
    expect(parts).toEqual(before);
  });

  it("keeps the shared strings in the worker: the prepared reply carries the rest of the sources", async () => {
    await readOffThread(book());
    const [worker] = InProcessWorker.instances;
    expect(worker!.replies[0]).toEqual({ op: "prepared", session: 1, sources: { format: "xlsx" } });
  });

  it("runs the shell phase here, then pulls every chunk one per request with a yield between pulls, then disposes", async () => {
    await readOffThread(book());
    expect(toyPhases.shells).toBe(1);
    expect(trace.events).toEqual([
      "post prepare",
      "post scan",
      "post chunk 0",
      "progress 0.2",
      "yield",
      "post chunk 1",
      "progress 0.4",
      "yield",
      "post chunk 2",
      "progress 0.6",
      "yield",
      "post chunk 3",
      "progress 0.8",
      "yield",
      "post chunk 4",
      "progress 1",
      "post dispose",
    ]);
    const [worker] = InProcessWorker.instances;
    const scanned = worker!.replies[1] as Extract<SheetParseResponse, { op: "scanned" }>;
    expect(scanned.chunkCount).toBe(5);
    // One SHEET_ROWS_CHUNK band of one sheet per reply, in order.
    const chunks = worker!.replies.slice(2, 7).map((reply) => (reply as Extract<SheetParseResponse, { op: "chunk" }>).chunk);
    expect(chunks.map((chunk) => [chunk.sheetIndex, chunk.chunk, chunk.rows.length])).toEqual([
      [0, 0, 512],
      [0, 1, 512],
      [0, 2, 176],
      [1, 0, 1],
      [1, 9, 1],
    ]);
    expect(worker!.replies[7]).toEqual({ op: "disposed" });
  });

  it("reads the model the inline phases read, completed on the main thread", async () => {
    const parts = book();
    const { model } = await readOffThread(parts);
    expect(model).toEqual(await inlineModel(parts));
    expect(model.sheets.map((sheet) => [sheet.name, sheet.rows.length, sheet.defaultRowHeightPt])).toEqual([
      ["Data", 1_200, 20],
      ["Sparse", 2, 20],
      ["Empty", 0, 20],
    ]);
    expect(model.sheets[0]!.rows[4]!.cells[0]!.value).toBe("beta");
    expect(model.activeSheetIndex).toBe(1);
  });

  it("reports progress 1 once for a workbook with no rows", async () => {
    const { model, progress } = await readOffThread(toyWorkbook({ Empty: [] }));
    expect(progress).toEqual([1]);
    expect(trace.events).toEqual(["post prepare", "post scan", "progress 1", "post dispose"]);
    expect(model.sheets.map((sheet) => sheet.rows)).toEqual([[]]);
  });

  describe("an error the read raised in the worker comes back as a reply, rethrown as the same class and message and never retried", () => {
    it.each<{ what: string; sheets: Record<string, string[]>; thrown: (error: unknown) => void }>([
      {
        what: "the cell-ceiling refusal, counted across every sheet",
        sheets: { First: ["!take 1500000"], Second: ["!take 600000"] },
        thrown: (error: unknown) => {
          expect(error).toBeInstanceOf(SpreadsheetOpenError);
          expect((error as Error).message).toBe(WORKBOOK_TOO_LARGE_MESSAGE);
        },
      },
      {
        what: "an XmlScanError, offset and all",
        sheets: { First: ["0:1", "!xml 17"] },
        thrown: (error: unknown) => {
          expect(error).toBeInstanceOf(XmlScanError);
          expect((error as XmlScanError).message).toBe("Closing tag </row> has no open element (at offset 17)");
          expect((error as XmlScanError).offset).toBe(17);
        },
      },
      {
        what: "any other error, by name",
        sheets: { First: ["!bug"] },
        thrown: (error: unknown) => {
          expect(error).toBeInstanceOf(Error);
          expect((error as Error).name).toBe("TypeError");
          expect((error as Error).message).toBe("First's rows are not iterable");
        },
      },
    ])("$what", async ({ sheets, thrown }) => {
      const failure = await readOffThread(toyWorkbook(sheets)).then(
        () => undefined,
        (error: unknown) => error,
      );
      thrown(failure);
      expect(toyPhases.shells).toBe(1);
      expect(InProcessWorker.instances).toHaveLength(1);
      expect(InProcessWorker.instances[0]!.terminated).toBe(true);
    });

    it("refuses inline alike, the same budget across every sheet", async () => {
      await expect(inlineModel(toyWorkbook({ First: ["!take 1500000"], Second: ["!take 600000"] }))).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
    });
  });

  it("lets an error of the main thread's own shell phase through as it is, never retried", async () => {
    const parts = book();
    parts.set("broken.txt", encoder.encode(""));
    await expect(readOffThread(parts)).rejects.toThrow(new SpreadsheetOpenError("This file isn't an Excel workbook we can open"));
    expect(toyPhases.shells).toBe(1);
    expect(InProcessWorker.instances[0]!.terminated).toBe(true);
  });

  describe("a worker that fails mid-open is torn down and the open retried once, inline", () => {
    it.each<{ what: string; misbehave?: Misbehaviour; construct?: "throws"; shells: number }>([
      {
        what: "an error event while it scans",
        misbehave: (worker, request) => {
          if (request.op !== "scan") {
            return false;
          }
          worker.emit("error", { message: "the worker crashed" });
          return true;
        },
        shells: 2,
      },
      {
        what: "a reply that cannot be deserialised",
        misbehave: (worker, request) => {
          if (request.op !== "chunk" || request.index !== 1) {
            return false;
          }
          worker.emit("messageerror", {});
          return true;
        },
        shells: 2,
      },
      {
        what: "a worker that dies while no request is in flight (the shell phase is running)",
        misbehave: (worker, request) => {
          if (request.op !== "prepare") {
            return false;
          }
          setTimeout(() => worker.emit("error", { message: "out of memory" }), 0);
          return false;
        },
        shells: 2,
      },
      {
        what: "a worker that cannot be constructed",
        construct: "throws",
        shells: 1,
      },
    ])("$what", async ({ misbehave, construct, shells }) => {
      InProcessWorker.misbehave = misbehave;
      if (construct === "throws") {
        vi.stubGlobal("Worker", function refusedWorker(): never {
          throw new DOMException("Refused to create a worker", "SecurityError");
        });
      }
      const parts = book();
      const { model, progress } = await readOffThread(parts);
      expect(withoutIds(model)).toEqual(withoutIds(await inlineModel(parts)));
      expect(toyPhases.shells).toBe(shells);
      // Inline: progress per sheet, ending at 1.
      expect(progress.at(-1)).toBe(1);
      expect(InProcessWorker.instances.length).toBeLessThanOrEqual(1);
      expect(InProcessWorker.instances.every((worker) => worker.terminated)).toBe(true);
    });

    it("once: when the inline read fails too, its error is the open's", async () => {
      InProcessWorker.misbehave = (worker) => {
        worker.emit("error", { message: "the worker crashed" });
        return true;
      };
      await expect(readOffThread(toyWorkbook({ First: ["!take 1500000"], Second: ["!take 600000"] }))).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
      expect(InProcessWorker.instances).toHaveLength(1);
      expect(toyPhases.shells).toBe(1);
    });
  });
});

// --- .ods/.ots: the real phases, read like the xlsx family ------------------------------------------
// The toy layer stands in for xlsx/phases.ts only; an .ods runs
// ods/phases.ts itself, over a fixture, through the same client.

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");

async function odsParts(name = "mini-book.ods"): Promise<Map<string, Uint8Array>> {
  return unzip(new Uint8Array(readFileSync(join(E2E_FIXTURES, name))));
}

/** mini-book.ods with one table of `rows` single-number rows for content.xml. */
async function longOds(rows: number): Promise<Map<string, Uint8Array>> {
  const parts = await odsParts();
  const row = '<table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>';
  const content =
    '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">' +
    `<office:body><office:spreadsheet><table:table table:name="Long">${row.repeat(rows)}</table:table></office:spreadsheet></office:body></office:document-content>`;
  parts.set("content.xml", encoder.encode(content));
  return parts;
}

/** readSpreadsheetBytes' model of an .ods (ids minted from 1, the counters left as they were). */
async function inlineOds(parts: Map<string, Uint8Array>): Promise<SpreadsheetModel> {
  const counted = { minted, events: trace.events };
  minted = 0;
  trace.events = [];
  const model = await readSpreadsheetBytes(new Uint8Array(0), "ods", "Book", parts);
  ({ minted, events: trace.events } = counted);
  return model;
}

describe("readSpreadsheetOffThread reads an .ods like the xlsx family", () => {
  it("without a Worker: the same phases inline, handing the main thread back every SCAN_YIELD_ROWS rows", async () => {
    expect(isWorkerSupported()).toBe(false);
    const parts = await longOds(4_500);
    // An earlier test's settled open may still have a task queued: let it run first.
    await new Promise((resolve) => setTimeout(resolve, 20));
    trace.events = [];
    const progress: number[] = [];
    const model = await readSpreadsheetOffThread(parts, "ods", "Book", { onProgress: (fraction) => progress.push(fraction) });
    // One table of 4,500 rows: 2,000 + 2,000 + 500, a yield after each step that did not finish it.
    expect(trace.events.filter((event) => event === "yield")).toHaveLength(2);
    expect(progress).toEqual([1]);
    expect(model.sheets.map((sheet) => [sheet.name, sheet.rows.length])).toEqual([["Long", 4_500]]);
    expect(model).toEqual(await inlineOds(await longOds(4_500)));
  });

  describe("with a Worker", () => {
    beforeEach(() => {
      vi.stubGlobal("Worker", InProcessWorker);
    });

    it("posts a copy of content.xml only, runs the shell here, pulls every chunk with a yield between pulls, then disposes — the model readSpreadsheetBytes reads", async () => {
      const parts = await odsParts();
      const before = new Map([...parts].map(([name, bytes]) => [name, bytes.slice()]));
      minted = 0;
      const progress: number[] = [];
      const model = await readSpreadsheetOffThread(parts, "ods", "Book", { onProgress: (fraction) => progress.push(fraction) });
      expect(InProcessWorker.instances).toHaveLength(1);
      const [worker] = InProcessWorker.instances;
      expect(worker!.terminated).toBe(true);
      const prepare = worker!.received[0] as Extract<Posted, { op: "prepare" }>;
      expect(prepare.format).toBe("ods");
      expect(prepare.sources.map((source) => source.name)).toEqual(["content.xml"]);
      // (Compared as text: an inflated entry's array may come from another realm.)
      const text = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);
      expect(text(prepare.sources[0]!.bytes)).toBe(text(before.get("content.xml")));
      expect(worker!.transferred).not.toContain(parts.get("content.xml")!.buffer);
      for (const [name, bytes] of parts) {
        expect(text(bytes)).toBe(text(before.get(name)));
      }
      // The envelope and the tables' offsets come back for the shell phase.
      const prepared = worker!.replies[0] as Extract<SheetParseResponse, { op: "prepared" }>;
      expect(prepared.sources.format).toBe("ods");
      expect(prepared.sources.odsTables!.length).toBe(model.sheets.length);
      const { chunkCount } = worker!.replies[1] as Extract<SheetParseResponse, { op: "scanned" }>;
      expect(chunkCount).toBeGreaterThan(0);
      expect(worker!.received.map((request) => request.op)).toEqual(["prepare", "scan", ...Array.from({ length: chunkCount }, () => "chunk"), "dispose"]);
      expect(trace.events.filter((event) => event === "yield")).toHaveLength(chunkCount - 1);
      expect(progress).toEqual(Array.from({ length: chunkCount }, (_, index) => (index + 1) / chunkCount));
      expect(model).toEqual(await inlineOds(await odsParts()));
    });

    it("rethrows a malformed content.xml as the inline read's error — the same class and message — and never retries", async () => {
      const corrupt = async (): Promise<Map<string, Uint8Array>> => {
        const parts = await odsParts();
        parts.set("content.xml", encoder.encode(new TextDecoder().decode(parts.get("content.xml")).replace("</table:table-row>", "</table:table-rows>")));
        return parts;
      };
      const inline = await readSpreadsheetBytes(new Uint8Array(0), "ods", "Book", await corrupt()).then(
        () => undefined,
        (error: unknown) => error,
      );
      expect(inline).toBeInstanceOf(SpreadsheetOpenError);
      const thrown = await readSpreadsheetOffThread(await corrupt(), "ods", "Book").then(
        () => undefined,
        (error: unknown) => error,
      );
      expect(thrown).toBeInstanceOf(SpreadsheetOpenError);
      expect((thrown as Error).message).toBe((inline as Error).message);
      expect(InProcessWorker.instances).toHaveLength(1);
      expect(InProcessWorker.instances[0]!.received.map((request) => request.op)).toEqual(["prepare"]);
    });
  });
});

describe("spreadsheetParse.worker.ts", () => {
  it("answers each message with the session's reply under the request's id", async () => {
    const posted: unknown[] = [];
    let listener: ((event: { data: Posted }) => void) | undefined;
    vi.stubGlobal("self", {
      addEventListener: (_type: "message", callback: (event: { data: Posted }) => void) => {
        listener = callback;
      },
      postMessage: (message: unknown) => posted.push(message),
    });
    await import("../../workers/spreadsheetParse.worker");
    expect(listener).toBeDefined();
    listener!({ data: { id: 7, op: "prepare", format: "xlsx", sources: [{ name: "strings.txt", bytes: encoder.encode("alpha") }] } });
    listener!({ data: { id: 8, op: "chunk", session: 1, index: 0 } });
    listener!({ data: { id: 9, op: "dispose", session: 1 } });
    expect(posted).toEqual([
      { id: 7, op: "prepared", session: 1, sources: { format: "xlsx" } },
      { id: 8, op: "error", name: "Error", message: "workbook parse session 1 has no chunk 0 to send" },
      { id: 9, op: "disposed" },
    ]);
  });
});
