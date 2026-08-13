// The word-count worker (SPEC-047 AC-12): computes documentStats off the
// main thread. Framework-free, importing straight from src/lib like every
// other consumer — the walk itself (src/lib/wordCount.ts) already does no
// DOM work, so it needs no adaptation to run here.
import { documentStats } from "../lib/wordCount";
import type { DocumentModel } from "../lib/documentModel";

interface StatsRequest {
  id: number;
  model: DocumentModel;
  includeNotes: boolean;
}

// Cast rather than adding "webworker" to tsconfig's `lib`: that lib
// declares a `self` incompatible with the "DOM" lib's `Window` this project
// already depends on everywhere else, so every worker file in SPEC-047
// narrows the ambient `self` locally instead of changing a global compiler
// setting for three files.
const ctx = self as unknown as {
  addEventListener: (type: "message", listener: (event: MessageEvent<StatsRequest>) => void) => void;
  postMessage: (message: unknown) => void;
};

ctx.addEventListener("message", (event) => {
  const { id, model, includeNotes } = event.data;
  const stats = documentStats(model, { includeNotes });
  // A dedicated worker's postMessage has no targetOrigin parameter (that's
  // Window.postMessage's cross-origin concern) — the same non-applicability
  // idleScheduler.ts's MessagePort.postMessage call notes.
  // eslint-disable-next-line unicorn/require-post-message-target-origin
  ctx.postMessage({ id, stats });
});
