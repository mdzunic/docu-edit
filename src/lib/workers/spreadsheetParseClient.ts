// Main-thread client for the workbook parse worker (SPEC-137 "The parse
// protocol", "Where the work runs"). The xlsx family (.xlsx/.xlsm/.xltx/
// .xltm) has the read phases; an .ods is still read in one pass, inline,
// until its reader is split the same way. One open is one worker and one
// session of parseSession.ts's protocol:
//
//   prepare   copies of the source parts (sourcePartNames — xlsx: the shared
//             strings and the worksheets; ods: content.xml); the worker
//             decodes them and runs phase 1, keeping the shared strings;
//   (here)    phase 2, readWorkbookShell, with the sources the worker posted
//             back: every small part, DOMParser and all;
//   scan      the jobs and their context; the worker runs phase 3 for every
//             sheet, counting cells against MAX_WORKBOOK_CELLS across all;
//   chunk ×n  the rows, one SHEET_ROWS_CHUNK band per request with a
//             yieldToMainThread() between pulls — so no single message
//             deserialises more than one chunk here — reporting
//             (index + 1) / chunkCount as the open's progress;
//   dispose   then phase 4, assembleWorkbook, here — the format's completion
//             finishing what needs the package or DOMParser, as inline.
//
// The worker is created for the open and terminated when the open settles,
// on every path. The read degrades, never fails: without a Worker (Vitest,
// jsdom) the same phases run inline, yielding every SCAN_YIELD_ROWS rows; a
// worker that fails — an `error` or `messageerror` event, a worker that
// cannot be constructed or posted to, one that died between requests — is
// torn down and the open retried once inline. An error the READ raised in
// the worker is another matter: it arrives as a reply and is rethrown as the
// same class with the same message (parseSession.ts errorFromReply), never
// retried — an inline read would only raise it again — so every open-funnel
// message is unchanged.
//
// src/lib/workers is app-layer infrastructure, so this client may import
// lib/spreadsheet; the product layer may not import it, which is why
// upgradeSpreadsheetModel takes its reader as a parameter (the store passes
// this one).
import { yieldToMainThread } from "../core/idleScheduler";
import type { SpreadsheetFormat, SpreadsheetModel } from "../spreadsheet/model";
import { assembleWorkbook, readPhased, readSpreadsheetBytes, readWorkbookShell, sourcePartNames, type RowChunk } from "../spreadsheet/open";
import { errorFromReply, type SheetParseRequest, type SheetParseResponse } from "../spreadsheet/parseSession";
import { createWorkerRpcClient, isWorkerSupported, type WorkerRpcClient } from "./workerRpc";

/** Rows an inline read scans between yields to the main thread (no Worker,
 * or the retry after a worker failed). */
export const SCAN_YIELD_ROWS = 2_000;

export interface OffThreadReadOptions {
  /** The open's progress, 0..1, as the rows arrive (chunks pulled / chunk count). */
  onProgress?: (fraction: number) => void;
}

type ParseClient = WorkerRpcClient<SheetParseRequest, SheetParseResponse>;

/** The worker failed — as opposed to the read failing in it, which arrives
 * as an error reply. The open is retried inline. */
class WorkerFailure extends Error {
  constructor(cause: unknown) {
    super("the workbook parse worker failed", { cause });
    this.name = "WorkerFailure";
  }
}

/** Reads an inflated workbook package into a new SpreadsheetModel, parsing
 * its O(cells) parts off the main thread (see the module header) — the same
 * model readSpreadsheetBytes reads. */
export async function readSpreadsheetOffThread(
  parts: Map<string, Uint8Array>,
  format: SpreadsheetFormat,
  title: string,
  options: OffThreadReadOptions = {},
): Promise<SpreadsheetModel> {
  if (format !== "xlsx") {
    // Only the xlsx family has the read phases. An .ods is read in
    // readSpreadsheetBytes' single pass, inline; a delimited-text file is no
    // package (csv/parse.ts reads it — the open funnel never sends one here,
    // and readSpreadsheetBytes says so).
    const model = await readSpreadsheetBytes(new Uint8Array(0), format, title, parts);
    options.onProgress?.(1);
    return model;
  }
  if (!isWorkerSupported()) {
    return readInline(parts, format, title, options);
  }
  let constructed = false;
  const client: ParseClient = createWorkerRpcClient<SheetParseRequest, SheetParseResponse>(() => {
    // workerRpc drops a worker that fired `error` and builds a fresh one for
    // the next call — which would hold none of this open's session. A
    // second construction is the first worker's failure (one that died
    // while no request was in flight: during the shell phase, between pulls).
    if (constructed) {
      throw new Error("the workbook parse worker stopped between requests");
    }
    constructed = true;
    return new Worker(new URL("../../workers/spreadsheetParse.worker.ts", import.meta.url), { type: "module" });
  });
  try {
    return await readInWorker(client, parts, format, title, options);
  } catch (error) {
    if (!(error instanceof WorkerFailure)) {
      throw error;
    }
    client.terminate();
    return await readInline(parts, format, title, options);
  } finally {
    client.terminate();
  }
}

/** The protocol, start to finish (see the module header). */
async function readInWorker(
  client: ParseClient,
  parts: Map<string, Uint8Array>,
  format: SpreadsheetFormat,
  title: string,
  options: OffThreadReadOptions,
): Promise<SpreadsheetModel> {
  // Copies, never the parts themselves: the parts — and the `opaque` the
  // shell phase encodes from them — keep their own bytes. bytes.slice()
  // because a stored zip entry is a view into the whole package, which
  // posting the view would clone entire; each copy, this open's alone, is
  // then transferred rather than cloned a second time.
  const sources: { name: string; bytes: Uint8Array }[] = [];
  for (const name of sourcePartNames(parts, format)) {
    const bytes = parts.get(name);
    if (bytes) {
      sources.push({ name, bytes: bytes.slice() });
    }
  }
  const transfer = sources.map((source) => source.bytes.buffer as ArrayBuffer);
  const prepared = await request(client, "prepared", { op: "prepare", format, sources }, transfer);
  const { session } = prepared;
  const { shell, jobs, context } = await readWorkbookShell(parts, prepared.sources, title);
  const { results, chunkCount } = await request(client, "scanned", { op: "scan", session, jobs, context });
  const strings = new Map<string, string>();
  const chunks: RowChunk[] = [];
  for (let index = 0; index < chunkCount; index += 1) {
    if (index > 0) {
      await yieldToMainThread();
    }
    const { chunk } = await request(client, "chunk", { op: "chunk", session, index });
    internStrings(chunk, strings);
    chunks.push(chunk);
    options.onProgress?.((index + 1) / chunkCount);
  }
  if (chunkCount === 0) {
    options.onProgress?.(1);
  }
  // Every row is here: a worker failing now costs the read nothing, and it
  // is terminated when the open settles either way.
  await client.call({ op: "dispose", session }).catch(() => undefined);
  return assembleWorkbook(shell, results, chunks);
}

/** The same phases inline, handing the main thread back every
 * SCAN_YIELD_ROWS rows. */
async function readInline(parts: Map<string, Uint8Array>, format: SpreadsheetFormat, title: string, options: OffThreadReadOptions): Promise<SpreadsheetModel> {
  return readPhased(parts, format, title, { yieldEveryRows: SCAN_YIELD_ROWS, onProgress: options.onProgress });
}

/** One request answered with `op`; else rejected — with the read's own
 * error when the reply carries one, with a WorkerFailure when the worker
 * failed (the call rejected) or answered out of protocol. */
async function request<Op extends SheetParseResponse["op"]>(
  client: ParseClient,
  op: Op,
  message: SheetParseRequest,
  transfer?: Transferable[],
): Promise<Extract<SheetParseResponse, { op: Op }>> {
  let reply: SheetParseResponse;
  try {
    reply = await client.call(message, transfer);
  } catch (failure) {
    throw new WorkerFailure(failure);
  }
  if (reply.op === "error") {
    throw errorFromReply(reply);
  }
  if (reply.op !== op) {
    throw new WorkerFailure(new Error(`the workbook parse worker answered "${message.op}" with "${reply.op}"`));
  }
  return reply as Extract<SheetParseResponse, { op: Op }>;
}

/** Structured clone gives every string a chunk carries its own copy on this
 * side, where the worker's cells held one instance per shared string
 * (SPEC-137 "Memory and the cell ceiling"): each string value is swapped for
 * the first equal one this open has seen, so equal values share one
 * instance again and the copies die young. One map lookup per string cell. */
function internStrings(chunk: RowChunk, strings: Map<string, string>): void {
  for (const row of chunk.rows) {
    for (const cell of row.cells) {
      const value = cell.value;
      if (typeof value !== "string") {
        continue;
      }
      const seen = strings.get(value);
      if (seen === undefined) {
        strings.set(value, value);
      } else {
        cell.value = seen;
      }
    }
  }
}
