// A loaded spell-check dictionary: a case-insensitive word set plus the
// word list itself, which lib/spellcheck/suggestions.ts scans to rank
// suggestions. Not a Hunspell engine — matching is a set lookup against a
// vendored word list that already carries every inflected form, so no suffix
// rules are needed or wanted here. Guessing at suffixes would only ADD
// acceptances, and a checker that infers "boxs" from "box" stops catching a
// real class of typo.
//
// What a word list cannot express is the shapes the tokenizer deliberately
// keeps whole (see tokenize.ts): "don't", "client's" and "well-known" arrive
// as single tokens, and `wordlist-english` contains no apostrophes and no
// hyphens at all. English dictionaries therefore opt into
// `wordShapes: "english"`, which resolves those compounds down to the words
// they are built from. Without it every contraction in a document is
// underlined — which is the failure this option exists to prevent.
export interface Dictionary {
  has(word: string): boolean;
  readonly words: readonly string[];
}

export interface DictionaryOptions {
  /** Accept the compound token shapes English writing produces: contractions
   * ("don't"), possessives ("client's", "employees'") and hyphenated
   * compounds ("well-known"). Omit for languages that form these differently. */
  wordShapes?: "english";
}

// The clitics a token can end in once the tokenizer has kept it whole. "n't"
// leaves a real stem behind ("do", "is", "did"); "can't", "won't" and
// "shan't" leave "ca"/"wo"/"sha" and are carried in enExtra.ts instead.
const CLITICS = ["n't", "'s", "'re", "'ve", "'ll", "'d", "'m"] as const;

function createEnglishMatcher(lookup: ReadonlySet<string>): (word: string) => boolean {
  const withClitic = (word: string): boolean => {
    for (const clitic of CLITICS) {
      if (!word.endsWith(clitic) || word.length <= clitic.length) continue;
      if (lookup.has(word.slice(0, -clitic.length))) return true;
    }
    // Plural possessive: "the employees' rights" ends in a bare apostrophe,
    // which none of the clitics above match.
    if (word.endsWith("s'") && lookup.has(word.slice(0, -1))) return true;
    return false;
  };

  const part = (word: string): boolean => lookup.has(word) || withClitic(word);

  return (word: string): boolean => {
    // The tokenizer preserves typographic apostrophes; normalise so "don't"
    // and "don’t" are one case rather than two entries in every list.
    const normalized = word.replace(/’/g, "'");
    if (part(normalized)) return true;
    // "well-known", "read-only": a compound is correct when every part is.
    if (!normalized.includes("-")) return false;
    const parts = normalized.split("-");
    return parts.length > 1 && parts.every((piece) => piece.length > 0 && part(piece));
  };
}

export function createDictionary(
  words: readonly string[],
  options: DictionaryOptions = {},
): Dictionary {
  const lookup = new Set(words.map((word) => word.toLowerCase()));
  const matches =
    options.wordShapes === "english"
      ? createEnglishMatcher(lookup)
      : (word: string): boolean => lookup.has(word);
  return {
    has(word: string): boolean {
      return matches(word.toLowerCase());
    },
    words,
  };
}
