// Loads (and caches) the bundled dictionary for a language tag (SPEC-035
// AC-17: on demand, when a language is first encountered — never at
// startup). Each language's word list is its own module, so a dynamic
// `import()` here is also a real code-split point: opening an
// English-only document never downloads the German/French/Spanish/Czech
// word lists.
import { isDictionaryAvailable, normalizeLanguageTag } from "../../language";
import { createDictionary, type Dictionary } from "../dictionary";

type Loader = () => Promise<{ WORDS: readonly string[] }>;

const LOADERS: Record<string, Loader> = {
  "en-US": () => import("./enUS"),
  "en-GB": () => import("./enGB"),
  "de-DE": () => import("./deDE"),
  "fr-FR": () => import("./frFR"),
  "es-ES": () => import("./esES"),
  "cs-CZ": () => import("./csCZ"),
};

const cache = new Map<string, Promise<Dictionary>>();

/** Loads the dictionary for a language tag, or null when no dictionary is
 * bundled for it (SPEC-035 AC-19) — the caller's job is to skip spell
 * checking for that text, not to error or show anything. */
export function loadDictionary(tag: string): Promise<Dictionary> | null {
  if (!isDictionaryAvailable(tag)) {
    return null;
  }
  const normalized = normalizeLanguageTag(tag);
  let cached = cache.get(normalized);
  if (!cached) {
    cached = LOADERS[normalized]().then((module) => createDictionary(module.WORDS));
    cache.set(normalized, cached);
  }
  return cached;
}
