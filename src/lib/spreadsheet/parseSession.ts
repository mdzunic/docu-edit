// The parse worker's sessions (SPEC-137 "The parse protocol", "Where the work
// runs"): the message handler workers/spreadsheetParse.worker.ts hosts — and
// parseSession.spec.ts drives directly — for the two DOMParser-free phases of
// a workbook read (readPhases.ts). One session is one open:
//
// - `prepare` decodes the source parts the main thread posted copies of
//   (xlsx: the shared strings and the worksheets; ods: content.xml) and runs
//   phase 1, prepareWorkbookSources. The session keeps the texts and the
//   whole WorkbookSources — the shared-string table never leaves the worker —
//   and replies with the rest (an ods envelope and its table offsets), which
//   the main thread's shell phase needs.
// - `scan` runs phase 3, scanSheet, for every job the shell phase made, in
//   order, against ONE cell budget for the whole workbook: the
//   WORKBOOK_TOO_LARGE_MESSAGE refusal counts across all sheets here exactly
//   as it does inline. The rows are kept as the RowChunks the scans emitted;
//   the texts and the sources are dropped, so a session scans once.
// - `chunk` hands the chunks over one per request, by index, releasing the
//   session's reference as each is sent, so no single reply deserialises
//   more than one chunk on the main thread and the worker's memory falls as
//   the main thread pulls.
// - `dispose` drops whatever the session still holds.
//
// Nothing thrown escapes handle(): every error becomes an { op: "error" }
// reply (errorReply), and errorFromReply rebuilds it on the main thread as
// the same class with the same message — XmlScanError with its offset, the
// readers' SpreadsheetOpenError (their parse errors and the cell-ceiling
// refusal) — so every open-funnel message reads exactly as an inline read's.
// Both directions live here, beside the protocol they marshal for.
//
// Imports are spreadsheet siblings and core only (boundary.spec.ts), and
// nothing here touches the DOM: in the worker, DOMParser, document and
// window do not exist.
import { XmlScanError } from "../core/xmlScanner";
import { SpreadsheetOpenError } from "./errors";
import { MAX_WORKBOOK_CELLS, type SpreadsheetFormat } from "./model";
import { prepareWorkbookSources, scanSheet } from "./open";
import { partText } from "./parsed";
import type { RowChunk, ScanBudget, ScanContext, SheetScanJob, SheetScanResult, WorkbookSources } from "./readPhases";

/** What the main thread asks the parse worker (workerRpc tags each with the
 * correlation id the worker echoes). */
export type SheetParseRequest =
  | { op: "prepare"; format: SpreadsheetFormat; sources: { name: string; bytes: Uint8Array }[] } // copies, not the parts themselves
  | { op: "scan"; session: number; jobs: SheetScanJob[]; context: ScanContext }
  | { op: "chunk"; session: number; index: number }
  | { op: "dispose"; session: number };

/** What the parse worker answers. */
export type SheetParseResponse =
  | { op: "prepared"; session: number; sources: Omit<WorkbookSources, "sharedStrings"> }
  | { op: "scanned"; results: SheetScanResult[]; chunkCount: number }
  | { op: "chunk"; chunk: RowChunk }
  | { op: "disposed" }
  | { op: "error"; name: string; message: string; offset?: number };

/** An error, as it crosses the worker boundary. */
export type SheetParseErrorReply = Extract<SheetParseResponse, { op: "error" }>;

/** The name an error reply gives SpreadsheetOpenError, which sets no `name`
 * of its own (its instances read "Error"). */
const SPREADSHEET_OPEN_ERROR = "SpreadsheetOpenError";

/** A thrown error as the reply that carries it across the boundary: its
 * class by name, its message verbatim and, for an XmlScanError, its offset. */
export function errorReply(error: unknown): SheetParseErrorReply {
  if (error instanceof XmlScanError) {
    return { op: "error", name: error.name, message: error.message, offset: error.offset };
  }
  if (error instanceof SpreadsheetOpenError) {
    return { op: "error", name: SPREADSHEET_OPEN_ERROR, message: error.message };
  }
  if (error instanceof Error) {
    return { op: "error", name: error.name, message: error.message };
  }
  return { op: "error", name: "Error", message: String(error) };
}

/** The error an { op: "error" } reply carries, rebuilt as the class that was
 * thrown with the same message (and offset): XmlScanError, the readers'
 * SpreadsheetOpenError, else an Error carrying the thrown error's name. */
export function errorFromReply(reply: SheetParseErrorReply): Error {
  if (reply.name === "XmlScanError") {
    // The constructor appends " (at offset N)" to what it is given; the
    // reply carries the finished message, so it is taken as it stands.
    const error = new XmlScanError("", reply.offset ?? 0);
    error.message = reply.message;
    return error;
  }
  if (reply.name === SPREADSHEET_OPEN_ERROR) {
    return new SpreadsheetOpenError(reply.message);
  }
  const error = new Error(reply.message);
  error.name = reply.name;
  return error;
}

/** One open's state in the worker. */
interface ParseSession {
  /** The decoded source parts, until the scan reads them. */
  texts?: Map<string, string>;
  /** Phase 1's output, until the scan uses it. */
  sources?: WorkbookSources;
  /** The scan's chunks in emission order, each released once sent. */
  chunks?: (RowChunk | undefined)[];
}

/** The parse worker's message handler and the sessions it holds. */
export interface ParseSessions {
  /** Answers one request — never throws (see the module header). */
  handle(request: SheetParseRequest): SheetParseResponse;
}

export function createParseSessions(): ParseSessions {
  const open = new Map<number, ParseSession>();
  let nextSession = 1;

  function sessionOf(id: number): ParseSession {
    const session = open.get(id);
    if (!session) {
      throw new Error(`no workbook parse session ${id} is open`);
    }
    return session;
  }

  function prepare(format: SpreadsheetFormat, posted: readonly { name: string; bytes: Uint8Array }[]): SheetParseResponse {
    const texts = new Map<string, string>();
    for (const { name, bytes } of posted) {
      texts.set(name, partText(bytes));
    }
    const sources = prepareWorkbookSources(texts, format);
    const session = nextSession;
    nextSession += 1;
    open.set(session, { texts, sources });
    const { sharedStrings: _sharedStrings, ...rest } = sources;
    return { op: "prepared", session, sources: rest };
  }

  function scan(id: number, jobs: readonly SheetScanJob[], context: ScanContext): SheetParseResponse {
    const session = sessionOf(id);
    const { texts, sources } = session;
    if (!texts || !sources) {
      throw new Error(`workbook parse session ${id} was already scanned`);
    }
    // Read once, here: the session lets go of the texts and the shared
    // strings whether or not the scan succeeds.
    session.texts = undefined;
    session.sources = undefined;
    const budget: ScanBudget = { cellsLeft: MAX_WORKBOOK_CELLS };
    const chunks: RowChunk[] = [];
    const results: SheetScanResult[] = [];
    for (const job of jobs) {
      results.push(scanSheet(job, texts.get(job.partName) ?? "", sources, context, (chunk) => chunks.push(chunk), budget));
    }
    session.chunks = chunks;
    return { op: "scanned", results, chunkCount: chunks.length };
  }

  function chunk(id: number, index: number): SheetParseResponse {
    const session = sessionOf(id);
    const sent = session.chunks?.[index];
    if (!sent) {
      throw new Error(`workbook parse session ${id} has no chunk ${index} to send`);
    }
    session.chunks![index] = undefined;
    return { op: "chunk", chunk: sent };
  }

  return {
    handle(request) {
      try {
        switch (request.op) {
          case "prepare":
            return prepare(request.format, request.sources);
          case "scan":
            return scan(request.session, request.jobs, request.context);
          case "chunk":
            return chunk(request.session, request.index);
          case "dispose":
            open.delete(request.session);
            return { op: "disposed" };
          default:
            // Requests arrive by postMessage, typed by convention only.
            throw new Error(`unknown workbook parse request ${JSON.stringify((request as { op?: unknown }).op)}`);
        }
      } catch (error) {
        return errorReply(error);
      }
    },
  };
}
