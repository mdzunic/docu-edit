// Suggestion ranking and capitalisation-pattern handling (SPEC-035 AC-2,
// AC-3.1's "up to five suggestions"; AC-4/AC-4.1's format- and
// case-preserving replacement).
import type { Dictionary } from "./dictionary";

export const MAX_SUGGESTIONS = 5;

// A suggestion more than this far from the misspelling is not a suggestion
// at all — it just adds noise. It is also the budget that makes ranking
// affordable: the dictionary is tens of thousands of words, and every one of
// them is a candidate for every right-click.
const MAX_EDIT_DISTANCE = 2;

/** Damerau-Levenshtein edit distance (insertion, deletion, substitution,
 * adjacent transposition) — transposition matters for typos like "wrold",
 * the single most common misspelling shape, which plain Levenshtein scores
 * as distance 2 instead of 1 and would rank behind less relevant words.
 *
 * Bounded by `max`: any distance greater than that is reported as exactly
 * `max + 1` rather than computed. Two things follow from the bound, and both
 * matter at dictionary scale. Only cells within `max` of the diagonal can
 * hold a value inside the budget, so each row computes a band of `2·max + 1`
 * cells instead of the whole width; and once an entire row exceeds the
 * budget no later row can come back under it, so the candidate is abandoned
 * there. The result is exact for every distance the caller can actually use.
 */
function editDistance(a: string, b: string, max: number): number {
  const al = a.length;
  const bl = b.length;
  if (Math.abs(al - bl) > max) return max + 1;
  if (al === 0) return bl;
  if (bl === 0) return al;

  const over = max + 1;
  let prev2: number[] | null = null;
  let prev: number[] = Array.from({ length: bl + 1 }, () => 0);
  let cur: number[] = Array.from({ length: bl + 1 }, () => 0);
  for (let j = 0; j <= bl; j++) prev[j] = j;

  for (let i = 1; i <= al; i++) {
    cur.fill(over);
    cur[0] = i;
    const from = Math.max(1, i - max);
    const to = Math.min(bl, i + max);
    let rowMin = over;
    for (let j = from; j <= to; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      let value = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      if (prev2 && i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        value = Math.min(value, prev2[j - 2] + cost);
      }
      cur[j] = value;
      if (value < rowMin) rowMin = value;
    }
    if (rowMin > max) return over;
    const spare: number[] = prev2 ?? Array.from({ length: bl + 1 }, () => 0);
    prev2 = prev;
    prev = cur;
    cur = spare;
  }
  return prev[bl];
}

/** Up to `MAX_SUGGESTIONS` dictionary words ranked by edit distance to
 * `word` (closest first; ties broken by the dictionary's own order, which
 * is frequency-ish for the bundled lists). Case-insensitive: callers apply
 * `applyCasePattern` to the result for display/replacement. */
export function getSuggestions(word: string, dictionary: Dictionary): string[] {
  const lower = word.toLowerCase();
  const scored: Array<{ candidate: string; distance: number }> = [];
  for (const candidate of dictionary.words) {
    // Reject on length before touching the matrix: each edit changes the
    // length by at most one, so a candidate more than MAX_EDIT_DISTANCE
    // longer or shorter can never score within budget. This is the cheapest
    // of the three filters and discards most of the list on its own.
    if (Math.abs(candidate.length - lower.length) > MAX_EDIT_DISTANCE) {
      continue;
    }
    const distance = editDistance(lower, candidate.toLowerCase(), MAX_EDIT_DISTANCE);
    if (distance <= MAX_EDIT_DISTANCE) {
      scored.push({ candidate, distance });
    }
  }
  scored.sort((a, b) => a.distance - b.distance);
  const seen = new Set<string>();
  const results: string[] = [];
  for (const { candidate } of scored) {
    const key = candidate.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    results.push(candidate);
    if (results.length >= MAX_SUGGESTIONS) break;
  }
  return results;
}

export type CasePattern = "lower" | "upper" | "capitalized" | "mixed";

/** The capitalisation pattern of a misspelled word (SPEC-035 AC-4.1): only
 * "all letters uppercase" and "first letter uppercase, rest lowercase" are
 * canonical patterns a replacement can be shaped to reproduce; anything
 * else (mixed case, no letters) is reported as "mixed" and left as the
 * dictionary spells it (see the spec's own ambiguity note on mixed case). */
export function detectCasePattern(word: string): CasePattern {
  const letters = Array.from(word).filter((ch) => /\p{L}/u.test(ch));
  if (letters.length === 0) {
    return "lower";
  }
  const joined = letters.join("");
  if (joined === joined.toUpperCase() && joined !== joined.toLowerCase()) {
    return "upper";
  }
  if (joined === joined.toLowerCase()) {
    return "lower";
  }
  const first = letters[0];
  const rest = letters.slice(1).join("");
  if (first === first.toUpperCase() && rest === rest.toLowerCase()) {
    return "capitalized";
  }
  return "mixed";
}

/** `replacement` reshaped to match `pattern` (SPEC-035 AC-4.1) — "mixed"
 * has no single well-defined target shape, so the suggestion's own
 * (dictionary) casing is kept as-is. */
export function applyCasePattern(replacement: string, pattern: CasePattern): string {
  switch (pattern) {
    case "upper":
      return replacement.toUpperCase();
    case "capitalized":
      return replacement.length > 0
        ? replacement[0].toUpperCase() + replacement.slice(1).toLowerCase()
        : replacement;
    case "lower":
      return replacement.toLowerCase();
    default:
      return replacement;
  }
}
