// SPEC-137 AC-21: parseSession.ts's message handler, driven directly the way
// the parse worker drives it — every request and every reply through
// structuredClone, as postMessage carries them — over every workbook fixture
// in e2e/fixtures, the xlsx family and .ods/.ots: prepare → the shell phase
// on the "main thread" → scan → every chunk pulled → dispose →
// assembleWorkbook, whose model equals readSpreadsheetBytes'. Then the error
// replies (a malformed worksheet or table; the cell-ceiling refusal, counted
// across every sheet of the session), the chunks (SHEET_ROWS_CHUNK-row bands
// of row index, in order, each sent once) and the session's lifetime
// (dispose frees it). The protocol's logic apart from the real phases — the
// client, its fallback and retry — is workers/spreadsheetParseClient.spec.ts's.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { detectPackage, spreadsheetFormatOf } from "../core/packageDetect";
import { XmlScanError } from "../core/xmlScanner";
import { unzip } from "../core/zip";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "./errors";
import type { SpreadsheetFormat, SpreadsheetModel } from "./model";
import { assembleWorkbook, readSpreadsheetBytes, readWorkbookShell, sourcePartNames, type RowChunk, type SheetScanResult } from "./open";
import { createParseSessions, errorFromReply, errorReply, type ParseSessions, type SheetParseRequest, type SheetParseResponse } from "./parseSession";
import { chunkOf, chunkRows } from "./rowChunks";

// The cell ceiling, lowered by the refusal tests (read live: a session
// takes MAX_WORKBOOK_CELLS when it scans).
const ceiling = vi.hoisted(() => ({ cells: undefined as number | undefined }));
vi.mock("./model", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./model")>();
  return {
    ...actual,
    get MAX_WORKBOOK_CELLS() {
      return ceiling.cells ?? actual.MAX_WORKBOOK_CELLS;
    },
  };
});

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");

const ODS_FAMILY = /\.(ods|ots)$/i;

/** Every committed workbook fixture, xlsx family and .ods/.ots (the
 * generated spec137-* benchmark files are not committed, and too large for a
 * unit test). */
const FIXTURES = readdirSync(E2E_FIXTURES)
  .filter((name) => (/\.(xlsx|xlsm|xltx|xltm)$/i.test(name) || ODS_FAMILY.test(name)) && !name.startsWith("spec137-"))
  .toSorted();

let minted = 0;

/** Ids minted from a counter, reset before each read, so two reads of one
 * file mint the same ids in the same order. */
function resetIds(): void {
  minted = 0;
}

beforeEach(() => {
  resetIds();
  vi.spyOn(crypto, "randomUUID").mockImplementation(() => {
    minted += 1;
    return `00000000-0000-4000-8000-${String(minted).padStart(12, "0")}`;
  });
});

afterEach(() => {
  ceiling.cells = undefined;
  vi.restoreAllMocks();
});

interface Fixture {
  bytes: Uint8Array;
  format: SpreadsheetFormat;
}

function fixture(name: string): Fixture {
  return { bytes: new Uint8Array(readFileSync(join(E2E_FIXTURES, name))), format: ODS_FAMILY.test(name) ? "ods" : "xlsx" };
}

/** The package as the open funnel hands it on — unzipped afresh per read,
 * so no read shares a part with another — and the format it detects. */
async function inflated(bytes: Uint8Array, expected: SpreadsheetFormat): Promise<Map<string, Uint8Array>> {
  const parts = await unzip(bytes);
  const detected = spreadsheetFormatOf(detectPackage(parts));
  expect(detected !== null && detected !== "xlsb" ? detected.format : detected).toBe(expected);
  return parts;
}

/** One request as the worker receives it, and its reply as the main thread
 * does: each through structuredClone, as postMessage carries them. */
function post(sessions: ParseSessions, request: SheetParseRequest): SheetParseResponse {
  return structuredClone(sessions.handle(structuredClone(request)));
}

function expectReply<Op extends SheetParseResponse["op"]>(reply: SheetParseResponse, op: Op): Extract<SheetParseResponse, { op: Op }> {
  expect(reply).toMatchObject({ op });
  return reply as Extract<SheetParseResponse, { op: Op }>;
}

/** The source parts' copies, as the client posts them. */
function sourceCopies(parts: Map<string, Uint8Array>, format: SpreadsheetFormat): { name: string; bytes: Uint8Array }[] {
  return sourcePartNames(parts, format).flatMap((name) => {
    const bytes = parts.get(name);
    return bytes ? [{ name, bytes: bytes.slice() }] : [];
  });
}

interface Prepared {
  session: number;
  shell: SpreadsheetModel;
  scan: Extract<SheetParseRequest, { op: "scan" }>;
}

/** prepare, then the shell phase on the "main thread": what the scan needs. */
async function prepare(sessions: ParseSessions, parts: Map<string, Uint8Array>, format: SpreadsheetFormat, title: string): Promise<Prepared> {
  const prepared = expectReply(post(sessions, { op: "prepare", format, sources: sourceCopies(parts, format) }), "prepared");
  expect(prepared.sources).not.toHaveProperty("sharedStrings");
  expect(prepared.sources.format).toBe(format);
  const { shell, jobs, context } = await readWorkbookShell(parts, prepared.sources, title);
  return { session: prepared.session, shell, scan: { op: "scan", session: prepared.session, jobs, context } };
}

interface SessionRead {
  model: SpreadsheetModel;
  results: SheetScanResult[];
  chunks: RowChunk[];
  session: number;
}

/** The whole protocol through one session, as the parse client runs it. */
async function readThroughSession(sessions: ParseSessions, parts: Map<string, Uint8Array>, format: SpreadsheetFormat, title: string): Promise<SessionRead> {
  const { session, shell, scan } = await prepare(sessions, parts, format, title);
  const scanned = expectReply(post(sessions, scan), "scanned");
  const chunks: RowChunk[] = [];
  for (let index = 0; index < scanned.chunkCount; index += 1) {
    chunks.push(expectReply(post(sessions, { op: "chunk", session, index }), "chunk").chunk);
  }
  expectReply(post(sessions, { op: "dispose", session }), "disposed");
  return { model: assembleWorkbook(shell, scanned.results, chunks), results: scanned.results, chunks, session };
}

describe("a workbook read through the parse session equals readSpreadsheetBytes' (every fixture, xlsx family and .ods)", () => {
  it("has fixtures to read, templates and macro-enabled ones among them", () => {
    expect(FIXTURES.length).toBeGreaterThan(20);
    expect(FIXTURES.some((name) => name.endsWith(".xltx"))).toBe(true);
    expect(FIXTURES.some((name) => name.endsWith(".xlsm"))).toBe(true);
  });

  it("has every .ods fixture among them", () => {
    expect(FIXTURES.filter((name) => ODS_FAMILY.test(name)).length).toBeGreaterThanOrEqual(19);
  });

  it.each(FIXTURES)("%s", async (name) => {
    const { bytes, format } = fixture(name);
    resetIds();
    const inline = await readSpreadsheetBytes(bytes, format, name);
    resetIds();
    const { model } = await readThroughSession(createParseSessions(), await inflated(bytes, format), format, name);
    expect(model).toEqual(inline);
  });
});

describe("the session's chunks", () => {
  // mini-book.xlsx: three sheets of one band each (Data: rows 0–499);
  // spec132-charts.xlsx: a chart sheet among them, which has no rows to send.
  // (Bands past the first, and empty bands between, are the parse client
  // spec's toy workbook's.)
  it.each(["mini-book.xlsx", "spec124-structure.xlsx", "spec132-charts.xlsx", "mini-book.ods", "spec116-wide.ods"])("are SHEET_ROWS_CHUNK-row bands of row index, sheet by sheet in order (%s)", async (name) => {
    const { bytes, format } = fixture(name);
    const inline = await readSpreadsheetBytes(bytes, format, name);
    const { chunks } = await readThroughSession(createParseSessions(), await inflated(bytes, format), format, name);
    for (const chunk of chunks) {
      expect(chunk.rows.length).toBeGreaterThan(0);
      for (const row of chunk.rows) {
        expect(chunkOf(row.index)).toBe(chunk.chunk);
      }
    }
    const order = chunks.map((chunk) => [chunk.sheetIndex, chunk.chunk] as const);
    expect(order).toEqual(order.toSorted((a, b) => a[0] - b[0] || a[1] - b[1]));
    // One chunk per band that holds rows: what the inline read's rows cut into bands give.
    expect(order).toEqual(inline.sheets.flatMap((sheet, sheetIndex) => chunkRows(sheet.rows).map(([chunk]) => [sheetIndex, chunk] as const)));
  });

  it("are pulled by index, each once", async () => {
    const { bytes, format } = fixture("mini-book.xlsx");
    const sessions = createParseSessions();
    const { session, scan } = await prepare(sessions, await inflated(bytes, format), format, "mini-book");
    const { chunkCount } = expectReply(post(sessions, scan), "scanned");
    expect(chunkCount).toBe(3);
    const second = expectReply(post(sessions, { op: "chunk", session, index: 1 }), "chunk").chunk;
    const first = expectReply(post(sessions, { op: "chunk", session, index: 0 }), "chunk").chunk;
    expect([first.sheetIndex, second.sheetIndex]).toEqual([0, 1]);
    // Sent, so released: the worker no longer holds it.
    expect(post(sessions, { op: "chunk", session, index: 0 })).toEqual({ op: "error", name: "Error", message: `workbook parse session ${session} has no chunk 0 to send` });
    expect(post(sessions, { op: "chunk", session, index: 3 })).toEqual({ op: "error", name: "Error", message: `workbook parse session ${session} has no chunk 3 to send` });
  });
});

describe("the session's lifetime", () => {
  it("numbers each prepare as a session of its own", async () => {
    const { bytes, format } = fixture("mini-book.xlsx");
    const sessions = createParseSessions();
    const parts = await inflated(bytes, format);
    const first = expectReply(post(sessions, { op: "prepare", format, sources: sourceCopies(parts, format) }), "prepared");
    const second = expectReply(post(sessions, { op: "prepare", format, sources: sourceCopies(parts, format) }), "prepared");
    expect(second.session).not.toBe(first.session);
  });

  it("scans once: the source texts are let go of when read", async () => {
    const { bytes, format } = fixture("mini-book.xlsx");
    const sessions = createParseSessions();
    const { session, scan } = await prepare(sessions, await inflated(bytes, format), format, "mini-book");
    expectReply(post(sessions, scan), "scanned");
    expect(post(sessions, scan)).toEqual({ op: "error", name: "Error", message: `workbook parse session ${session} was already scanned` });
  });

  it("is freed by dispose: nothing of it can be pulled after, and disposing again is harmless", async () => {
    const { bytes, format } = fixture("mini-book.xlsx");
    const sessions = createParseSessions();
    const { session, scan } = await prepare(sessions, await inflated(bytes, format), format, "mini-book");
    expect(expectReply(post(sessions, scan), "scanned").chunkCount).toBeGreaterThan(1);
    expect(post(sessions, { op: "dispose", session })).toEqual({ op: "disposed" });
    const gone = { op: "error", name: "Error", message: `no workbook parse session ${session} is open` };
    expect(post(sessions, { op: "chunk", session, index: 0 })).toEqual(gone);
    expect(post(sessions, scan)).toEqual(gone);
    expect(post(sessions, { op: "dispose", session })).toEqual({ op: "disposed" });
  });

  it("answers a request for a session never opened with an error, not a throw", () => {
    expect(createParseSessions().handle({ op: "chunk", session: 7, index: 0 })).toEqual({ op: "error", name: "Error", message: "no workbook parse session 7 is open" });
  });

  it("answers a request it does not know with an error", () => {
    const unknown = { op: "parse" } as unknown as SheetParseRequest;
    expect(createParseSessions().handle(unknown)).toEqual({ op: "error", name: "Error", message: 'unknown workbook parse request "parse"' });
  });
});

describe("error replies", () => {
  it("carry a malformed worksheet's parse error — the class and message an inline read throws", async () => {
    const { bytes, format } = fixture("mini-book.xlsx");
    const corrupt = async (): Promise<Map<string, Uint8Array>> => {
      const parts = await inflated(bytes, format);
      const name = "xl/worksheets/sheet2.xml";
      const text = new TextDecoder().decode(parts.get(name));
      // An end tag that matches nothing: inside sheetData, which only the
      // worksheet's scan (phase 3, in the worker) reads.
      expect(text).toContain("</row>");
      parts.set(name, new TextEncoder().encode(text.replace("</row>", "</rows>")));
      return parts;
    };
    const thrown = await readSpreadsheetBytes(new Uint8Array(0), format, "mini-book", await corrupt()).then(
      () => undefined,
      (error: unknown) => error,
    );
    expect(thrown).toBeInstanceOf(SpreadsheetOpenError);
    const sessions = createParseSessions();
    const { scan } = await prepare(sessions, await corrupt(), format, "mini-book");
    const reply = expectReply(post(sessions, scan), "error");
    expect(reply).toEqual({ op: "error", name: "SpreadsheetOpenError", message: (thrown as Error).message });
    const rebuilt = errorFromReply(reply);
    expect(rebuilt).toBeInstanceOf(SpreadsheetOpenError);
    expect(rebuilt.message).toBe((thrown as Error).message);
  });

  /** mini-book.ods with its content.xml edited — inside a table, which only
   * the worker's phases read (phase 1 checks the whole part; phase 3 reads
   * the table's values). */
  async function corruptOds(edit: (text: string) => string): Promise<Map<string, Uint8Array>> {
    const parts = await inflated(fixture("mini-book.ods").bytes, "ods");
    const text = new TextDecoder().decode(parts.get("content.xml"));
    const edited = edit(text);
    expect(edited).not.toBe(text);
    parts.set("content.xml", new TextEncoder().encode(edited));
    return parts;
  }

  it.each<{ what: string; edit: (text: string) => string; at: "prepared" | "scanned" }>([
    // An end tag that matches nothing, inside the first table: phase 1's
    // scan of content.xml refuses it, at the single pass's offset.
    { what: "a malformed table", edit: (text) => text.replace("</table:table-row>", "</table:table-rows>"), at: "prepared" },
    // A number that is none: phase 3's scan of the table refuses it, naming the cell.
    { what: "a float cell with no number", edit: (text) => text.replace(/(office:value-type="float"[^>]*?\soffice:value=")[^"]*"/, '$1twelve"'), at: "scanned" },
  ])("carry $what's .ods parse error — the class and message the single pass and an inline read throw", async ({ edit, at }) => {
    const { parseOds } = await import("./ods/reader");
    const single = await parseOds(new Uint8Array(0), await corruptOds(edit)).then(
      () => undefined,
      (error: unknown) => error,
    );
    expect(single).toBeInstanceOf(SpreadsheetOpenError);
    const thrown = await readSpreadsheetBytes(new Uint8Array(0), "ods", "mini-book", await corruptOds(edit)).then(
      () => undefined,
      (error: unknown) => error,
    );
    expect(thrown).toBeInstanceOf(SpreadsheetOpenError);
    expect((thrown as Error).message).toBe((single as Error).message);
    const sessions = createParseSessions();
    const parts = await corruptOds(edit);
    const prepared = post(sessions, { op: "prepare", format: "ods", sources: sourceCopies(parts, "ods") });
    let reply = prepared;
    if (at === "scanned") {
      const { sources, session } = expectReply(prepared, "prepared");
      const { jobs, context } = await readWorkbookShell(parts, sources, "mini-book");
      reply = post(sessions, { op: "scan", session, jobs, context });
    } else {
      expect((thrown as Error).message).toMatch(/not well-formed XML — Closing tag .*\(at offset \d+\) \(content\.xml\)$/);
    }
    expect(reply).toEqual({ op: "error", name: "SpreadsheetOpenError", message: (thrown as Error).message });
    const rebuilt = errorFromReply(expectReply(reply, "error"));
    expect(rebuilt).toBeInstanceOf(SpreadsheetOpenError);
    expect(rebuilt.message).toBe((thrown as Error).message);
  });

  // mini-book.xlsx: Summary 32 cells, Data 1,500, Hidden 3 — so a ceiling
  // one below the total admits every sheet on its own and refuses the
  // workbook (mini-book.ods holds the same sheets).
  it.each(["mini-book.xlsx", "spec121-links.xlsx", "mini-book.ods"])("carry the WORKBOOK_TOO_LARGE_MESSAGE refusal, counted across every sheet of the session (%s)", async (name) => {
    const { bytes, format } = fixture(name);
    const { results } = await readThroughSession(createParseSessions(), await inflated(bytes, format), format, name);
    const perSheet = results.map((result) => result.cells);
    const total = perSheet.reduce((sum, cells) => sum + cells, 0);
    expect(perSheet.filter((cells) => cells > 0).length).toBeGreaterThan(1);
    expect(Math.max(...perSheet)).toBeLessThan(total);

    ceiling.cells = total - 1;
    const sessions = createParseSessions();
    const { scan } = await prepare(sessions, await inflated(bytes, format), format, name);
    const reply = expectReply(post(sessions, scan), "error");
    expect(reply).toEqual({ op: "error", name: "SpreadsheetOpenError", message: WORKBOOK_TOO_LARGE_MESSAGE });
    const rebuilt = errorFromReply(reply);
    expect(rebuilt).toBeInstanceOf(SpreadsheetOpenError);
    expect(rebuilt.message).toBe(WORKBOOK_TOO_LARGE_MESSAGE);
    // The inline read refuses it alike…
    await expect(readSpreadsheetBytes(bytes, format, name)).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);

    // …and at exactly its total, the workbook reads.
    ceiling.cells = total;
    const { model } = await readThroughSession(createParseSessions(), await inflated(bytes, format), format, name);
    expect(model.sheets.map((sheet) => sheet.rows.length)).toEqual((await readSpreadsheetBytes(bytes, format, name)).sheets.map((sheet) => sheet.rows.length));
  });

  it("carry an XmlScanError's offset, and rebuild it identically — class, message and offset", () => {
    const thrown = new XmlScanError("unexpected end of input", 42);
    const reply = errorReply(thrown);
    expect(reply).toEqual({ op: "error", name: "XmlScanError", message: "unexpected end of input (at offset 42)", offset: 42 });
    const rebuilt = errorFromReply(structuredClone(reply));
    expect(rebuilt).toBeInstanceOf(XmlScanError);
    expect(rebuilt.name).toBe("XmlScanError");
    expect(rebuilt.message).toBe(thrown.message);
    expect((rebuilt as XmlScanError).offset).toBe(42);
  });

  it("rebuild the readers' SpreadsheetOpenError as that class, whose name stays Error", () => {
    const thrown = new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
    const reply = errorReply(thrown);
    expect(reply).toEqual({ op: "error", name: "SpreadsheetOpenError", message: WORKBOOK_TOO_LARGE_MESSAGE });
    const rebuilt = errorFromReply(structuredClone(reply));
    expect(rebuilt).toBeInstanceOf(SpreadsheetOpenError);
    expect(rebuilt.name).toBe(thrown.name);
    expect(rebuilt.message).toBe(thrown.message);
  });

  it("rebuild any other error as an Error carrying its name and message", () => {
    const rebuilt = errorFromReply(structuredClone(errorReply(new TypeError("rows is not iterable"))));
    expect(rebuilt).toBeInstanceOf(Error);
    expect(rebuilt).not.toBeInstanceOf(SpreadsheetOpenError);
    expect(rebuilt.name).toBe("TypeError");
    expect(rebuilt.message).toBe("rows is not iterable");
    expect(errorReply("a thrown string")).toEqual({ op: "error", name: "Error", message: "a thrown string" });
  });

  it("answer a format the phases do not read with the phases' own error", () => {
    expect(createParseSessions().handle({ op: "prepare", format: "csv", sources: [] })).toEqual({
      op: "error",
      name: "Error",
      message: "a .csv workbook is read by readCsvBytes, not as a package",
    });
  });
});
