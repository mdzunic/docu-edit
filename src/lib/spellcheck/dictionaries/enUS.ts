// American-English spellings (SPEC-035 AC-10/AC-10.1's "the dictionary
// matches the text's language"): the shared vocabulary plus the forms that
// are valid in en-US and flagged in en-GB, and vice versa via enGB.ts.
//
// SCOWL's `american-words-*` lists hold exactly that difference — "color" but
// not "colour", "organize" but not "organise" — so the split needs no curation
// of ours. See enBase.ts on why the tiers are unioned rather than picked.
import american10 from "wordlist-english/american-words-10.json";
import american20 from "wordlist-english/american-words-20.json";
import american35 from "wordlist-english/american-words-35.json";
import american40 from "wordlist-english/american-words-40.json";
import { BASE_WORDS } from "./enBase";
import { US_EXTRA } from "./enExtra";

export const WORDS: readonly string[] = [
  ...BASE_WORDS,
  ...american10,
  ...american20,
  ...american35,
  ...american40,
  ...US_EXTRA,
];
