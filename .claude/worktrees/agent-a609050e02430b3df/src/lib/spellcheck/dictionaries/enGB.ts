// British-English spellings (SPEC-035 AC-10/AC-10.1) — see enUS.ts.
import british10 from "wordlist-english/british-words-10.json";
import british20 from "wordlist-english/british-words-20.json";
import british35 from "wordlist-english/british-words-35.json";
import british40 from "wordlist-english/british-words-40.json";
import { BASE_WORDS } from "./enBase";
import { GB_EXTRA } from "./enExtra";

export const WORDS: readonly string[] = [
  ...BASE_WORDS,
  ...british10,
  ...british20,
  ...british35,
  ...british40,
  ...GB_EXTRA,
];
