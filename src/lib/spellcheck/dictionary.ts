// A loaded spell-check dictionary: a case-insensitive word set plus the
// word list itself, which lib/spellcheck/suggestions.ts scans to rank
// suggestions. Deliberately not a "real" Hunspell affix-rule engine — the
// product's package.json is protected from new dependencies (see
// CLAUDE.md), so this is a curated whole-word list per language instead.
// Format-compatible with a drop-in replacement: anything that can produce a
// flat array of correctly-spelled words works here.
export interface Dictionary {
  has(word: string): boolean;
  readonly words: readonly string[];
}

export function createDictionary(words: readonly string[]): Dictionary {
  const lookup = new Set(words.map((word) => word.toLowerCase()));
  return {
    has(word: string): boolean {
      return lookup.has(word.toLowerCase());
    },
    words,
  };
}
