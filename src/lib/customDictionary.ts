// The custom spelling dictionary (SPEC-035): a per-user, per-device word
// list, one entry per (word, language) pair — "Add to dictionary" adds a
// word for its current language (AC-7), and the same word can be a valid
// addition in one language but not another. Persisted in localStorage with
// the same tolerant-parse-with-fallback convention as lib/autocorrect.ts
// (AC-8: persists across sessions, editable from a settings dialog).
import { normalizeLanguageTag } from "./language";

export const CUSTOM_DICTIONARY_STORAGE_KEY = "docu-edit:custom-dictionary";

export interface CustomDictionaryEntry {
  word: string;
  /** Normalised language tag (SPEC-035), e.g. "en-US". */
  lang: string;
}

function load(): CustomDictionaryEntry[] {
  try {
    const raw = localStorage.getItem(CUSTOM_DICTIONARY_STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    const entries: CustomDictionaryEntry[] = [];
    for (const item of parsed) {
      if (
        item &&
        typeof item === "object" &&
        typeof (item as { word?: unknown }).word === "string" &&
        typeof (item as { lang?: unknown }).lang === "string" &&
        (item as { word: string }).word.trim().length > 0
      ) {
        entries.push({
          word: (item as { word: string }).word,
          lang: normalizeLanguageTag((item as { lang: string }).lang),
        });
      }
    }
    return entries;
  } catch {
    return [];
  }
}

let entries: CustomDictionaryEntry[] = load();
const listeners = new Set<() => void>();

function persist(): void {
  try {
    localStorage.setItem(CUSTOM_DICTIONARY_STORAGE_KEY, JSON.stringify(entries));
  } catch {
    // Persistence is best-effort (private browsing, full quota); the
    // in-memory list still applies for the rest of this session.
  }
  for (const listener of listeners) {
    listener();
  }
}

/** Notifies `listener` whenever the dictionary changes (add/remove, in this
 * tab or reloaded from storage) — returns an unsubscribe function. Plain
 * callback pub-sub, not a Vue ref: this module has no Vue import (lib/ is
 * framework-free per CLAUDE.md); composables wrap this in a ref themselves. */
export function subscribeCustomDictionary(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Every entry, in insertion order. */
export function customDictionaryEntries(): readonly CustomDictionaryEntry[] {
  return entries;
}

export function hasCustomWord(word: string, lang: string): boolean {
  const target = word.trim().toLowerCase();
  const normalizedLang = normalizeLanguageTag(lang);
  return entries.some((entry) => entry.word.toLowerCase() === target && entry.lang === normalizedLang);
}

/** Adds `word` for `lang` (SPEC-035 AC-7); a no-op (returns false) for an
 * empty word or one already present for that language (AC-8.2's "adding"
 * rejects a duplicate/blank the same way the settings dialog does). */
export function addCustomWord(word: string, lang: string): boolean {
  const trimmed = word.trim();
  if (trimmed.length === 0) {
    return false;
  }
  const normalizedLang = normalizeLanguageTag(lang);
  if (hasCustomWord(trimmed, normalizedLang)) {
    return false;
  }
  entries = [...entries, { word: trimmed, lang: normalizedLang }];
  persist();
  return true;
}

export function removeCustomWord(word: string, lang: string): void {
  const target = word.trim().toLowerCase();
  const normalizedLang = normalizeLanguageTag(lang);
  const next = entries.filter(
    (entry) => !(entry.word.toLowerCase() === target && entry.lang === normalizedLang),
  );
  if (next.length !== entries.length) {
    entries = next;
    persist();
  }
}

/** Test-only reset: re-reads from localStorage (isolating specs from each
 * other, and re-running the tolerant parse after a test pokes storage
 * directly). */
export function resetCustomDictionaryForTests(): void {
  entries = load();
}

/** Typed pass-through for a CustomDictionaryEntry[] arriving from a `.vue`
 * component's untyped Array prop (same espree constraint as
 * documentModel.ts's asCharacterFormatKey). */
export function asCustomDictionaryEntries(value: unknown): CustomDictionaryEntry[] {
  return value as CustomDictionaryEntry[];
}
