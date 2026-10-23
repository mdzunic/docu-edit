// The workbook parse worker (SPEC-137 "Where the work runs"): the O(cells)
// half of a workbook read — decoding the worksheet parts or content.xml, the
// shared-string table, every sheet's scan into row chunks — runs here, off the
// main thread. parseSession.ts owns the protocol and the sessions; this file
// only hosts its message handler. Like the phases it runs, nothing it loads
// touches DOMParser (open.workerSafety.spec.ts runs them with it stubbed
// away), which a worker does not have.
import { createParseSessions, type SheetParseRequest } from "../lib/spreadsheet/parseSession";

// See wordCount.worker.ts's comment on why `self` is cast rather than the
// tsconfig gaining a "webworker" lib entry.
const ctx = self as unknown as {
  addEventListener: (type: "message", listener: (event: MessageEvent<SheetParseRequest & { id: number }>) => void) => void;
  postMessage: (message: unknown) => void;
};

const sessions = createParseSessions();

ctx.addEventListener("message", (event) => {
  const { id } = event.data;
  // See wordCount.worker.ts's comment on why targetOrigin does not apply
  // to a dedicated worker's postMessage.
  // eslint-disable-next-line unicorn/require-post-message-target-origin
  ctx.postMessage({ id, ...sessions.handle(event.data) });
});
