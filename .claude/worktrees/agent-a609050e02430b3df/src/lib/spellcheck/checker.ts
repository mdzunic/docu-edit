// Flags misspelled words in a run of text (SPEC-035). Deliberately narrow:
// this only knows about a plain string and a Dictionary — language
// resolution, per-run skip rules (noProof/code style), custom-dictionary and
// session-ignore filtering, and viewport scheduling all live one layer up
// (composables/useSpellCheck.ts), so this stays unit-testable without a
// document model or DOM at all.
import type { Dictionary } from "./dictionary";
import { tokenizeForSpellCheck } from "./tokenize";

export interface SpellIssue {
  /** Offset into the checked text, inclusive. */
  start: number;
  /** Offset into the checked text, exclusive. */
  end: number;
  word: string;
}

/** Every word in `text` the dictionary does not recognise, in source order.
 * URLs/emails are already excluded by the tokenizer (AC-12); the caller
 * filters ignored/custom-dictionary words before rendering underlines. */
export function checkText(text: string, dictionary: Dictionary): SpellIssue[] {
  const issues: SpellIssue[] = [];
  for (const token of tokenizeForSpellCheck(text)) {
    if (!dictionary.has(token.text)) {
      issues.push({ start: token.start, end: token.end, word: token.text });
    }
  }
  return issues;
}
