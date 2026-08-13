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
  onmessage: ((event: MessageEvent<StatsRequest>) => void) | null;
  postMessage: (message: unknown) => void;
};

ctx.onmessage = (event) => {
  const { id, model, includeNotes } = event.data;
  const stats = documentStats(model, { includeNotes });
  ctx.postMessage({ id, stats });
};
