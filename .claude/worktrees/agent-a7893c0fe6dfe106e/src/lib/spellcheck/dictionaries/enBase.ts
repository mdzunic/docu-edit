// Common English vocabulary shared by en-US and en-GB (SPEC-035), vendored
// from `wordlist-english` (MIT) — see CLAUDE.md's dependency note.
//
// That package publishes SCOWL's word lists split exactly the way this
// directory already was: a shared list plus per-dialect lists, so
// "colour"/"color" stay on their own sides and AC-10 holds by construction.
//
// The size suffixes are SCOWL "sizes", and they are ADDITIVE, not cumulative:
// `english-words-40.json` holds the words that size 40 adds to size 35, so a
// usable dictionary is the union of every tier up to the one you want. Taking
// them up to 40 gives ~44k words for ~470 KB of raw JSON — enough that
// ordinary prose comes back clean, while size 50 and beyond mostly add rare
// and archaic vocabulary that costs payload and weakens error detection.
// Dictionaries load lazily (see ./index.ts), so this never reaches the
// startup path.
import shared10 from "wordlist-english/english-words-10.json";
import shared20 from "wordlist-english/english-words-20.json";
import shared35 from "wordlist-english/english-words-35.json";
import shared40 from "wordlist-english/english-words-40.json";
import { SHARED_EXTRA } from "./enExtra";

export const BASE_WORDS: readonly string[] = [
  ...shared10,
  ...shared20,
  ...shared35,
  ...shared40,
  ...SHARED_EXTRA,
];
