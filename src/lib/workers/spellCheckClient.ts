// Main-thread client for the spell-check worker (SPEC-047 AC-13): the
// dictionary scan itself (checkText) and the dictionary load it needs move
// off-thread here. DOM measurement (issue rects, from useSpellCheck.ts) has
// no worker equivalent and stays on the main thread — see
// specs/047-large-document-performance.md's "What leaves the main thread".
import { checkText, type SpellIssue } from "../spellcheck/checker";
import { loadDictionary } from "../spellcheck/dictionaries";
import type { Dictionary } from "../spellcheck/dictionary";
import { createWorkerRpcClient, isWorkerSupported, type WorkerRpcClient } from "./workerRpc";

export interface SpellCheckItem {
  /** Caller-chosen correlation key (e.g. a run index) — echoed back on the
   * matching result so callers can reassemble a batch without relying on
   * array order. */
  key: string;
  lang: string;
  text: string;
}

export interface SpellCheckItemResult {
  key: string;
  matches: SpellIssue[];
}

interface CheckRequest {
  items: SpellCheckItem[];
}

interface CheckResponse {
  results: SpellCheckItemResult[];
}

let client: WorkerRpcClient<CheckRequest, CheckResponse> | null = null;

function getClient(): WorkerRpcClient<CheckRequest, CheckResponse> {
  if (!client) {
    client = createWorkerRpcClient<CheckRequest, CheckResponse>(
      () => new Worker(new URL("../../workers/spellCheck.worker.ts", import.meta.url), { type: "module" }),
    );
  }
  return client;
}

const fallbackDictionaries = new Map<string, Promise<Dictionary | null>>();

function fallbackDictionaryFor(lang: string): Promise<Dictionary | null> {
  let cached = fallbackDictionaries.get(lang);
  if (!cached) {
    cached = loadDictionary(lang) ?? Promise.resolve(null);
    fallbackDictionaries.set(lang, cached);
  }
  return cached;
}

/** Checks a batch of texts (each already known to have a bundled
 * dictionary — callers filter with isDictionaryAvailable first) against
 * their language's dictionary. Empty matches for a text whose dictionary
 * somehow still fails to resolve, never a thrown error — matches AC-19's
 * "skip, don't error" contract. */
export async function checkTextsInWorker(items: SpellCheckItem[]): Promise<SpellCheckItemResult[]> {
  if (items.length === 0) {
    return [];
  }
  if (!isWorkerSupported()) {
    return Promise.all(
      items.map(async (item) => {
        const dictionary = await fallbackDictionaryFor(item.lang);
        return { key: item.key, matches: dictionary ? checkText(item.text, dictionary) : [] };
      }),
    );
  }
  const response = await getClient().call({ items });
  return response.results;
}
