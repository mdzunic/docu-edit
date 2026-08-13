// The spell-check worker (SPEC-047 AC-13): checkText plus the dictionary
// load it needs, both off the main thread. Per-run skip rules (noProof,
// code style, tracked deletions), custom-dictionary/ignore filtering and
// viewport scheduling stay one layer up (views/document/useSpellCheck.ts) —
// this only ever sees text this worker already knows must be checked.
import { checkText } from "../lib/spellcheck/checker";
import { loadDictionary } from "../lib/spellcheck/dictionaries";
import type { Dictionary } from "../lib/spellcheck/dictionary";

interface CheckItem {
  key: string;
  lang: string;
  text: string;
}

interface CheckRequest {
  id: number;
  items: CheckItem[];
}

// See wordCount.worker.ts's comment on why `self` is cast rather than the
// tsconfig gaining a "webworker" lib entry.
const ctx = self as unknown as {
  addEventListener: (type: "message", listener: (event: MessageEvent<CheckRequest>) => void) => void;
  postMessage: (message: unknown) => void;
};

// One dictionary cache for the worker's lifetime — a language loaded for an
// earlier request is already resident for the next one, matching the
// composable's own "on demand, once" contract (SPEC-035 AC-17), just kept
// here instead of on the main thread.
const dictionaries = new Map<string, Promise<Dictionary | null>>();

function dictionaryFor(lang: string): Promise<Dictionary | null> {
  let cached = dictionaries.get(lang);
  if (!cached) {
    cached = loadDictionary(lang) ?? Promise.resolve(null);
    dictionaries.set(lang, cached);
  }
  return cached;
}

ctx.addEventListener("message", (event) => {
  const { id, items } = event.data;
  void Promise.all(
    items.map(async (item) => {
      const dictionary = await dictionaryFor(item.lang);
      return { key: item.key, matches: dictionary ? checkText(item.text, dictionary) : [] };
    }),
  ).then((results) => {
    // See wordCount.worker.ts's comment on why targetOrigin does not apply
    // to a dedicated worker's postMessage.
    // eslint-disable-next-line unicorn/require-post-message-target-origin
    ctx.postMessage({ id, results });
  });
});
