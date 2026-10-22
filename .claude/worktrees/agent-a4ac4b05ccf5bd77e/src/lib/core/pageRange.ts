// Parsing a custom number range — "2", "1-3", "1-2, 5" — against a total
// (SPEC-034's page-range input, moved here by SPEC-104 so the deck print
// dialog can parse *slide* numbers with the very same rules). The noun is the
// only thing that varies: lib/printRange.ts keeps its page-worded wrappers, so
// every SPEC-034 error string is unchanged, and the deck dialog passes
// "slide". Framework-free and product-neutral.

/** A parsed range: the 1-based numbers selected, ascending, or a user-facing
 * validation error. Exactly one of the two is set. */
export interface NumberRangeResult {
  pages?: number[];
  error?: string;
}

/** The hint every range error ends with, e.g. `Use page numbers like "1-3, 5"`. */
export function rangeHint(noun = "page"): string {
  return `Use ${noun} numbers like "1-3, 5"`;
}

function capitalised(text: string): string {
  return text.length > 0 ? text[0].toUpperCase() + text.slice(1) : text;
}

/** Parses a custom range string against `total` items named `noun`. Errors on
 * empty input, non-numeric tokens, ranges whose start is after their end, and
 * numbers outside 1..total. */
export function parseNumberRange(input: string, total: number, noun = "page"): NumberRangeResult {
  const hint = rangeHint(noun);
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return { error: `Enter a ${noun} range. ${hint}` };
  }
  const selected = new Set<number>();
  for (const token of trimmed.split(",")) {
    const part = token.trim();
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(part);
    if (!match) {
      return { error: `"${part}" is not a ${noun} number. ${hint}` };
    }
    const start = Number.parseInt(match[1], 10);
    const end = match[2] === undefined ? start : Number.parseInt(match[2], 10);
    if (end < start) {
      return { error: `"${part}" starts after it ends` };
    }
    if (start < 1 || end > total) {
      return {
        error:
          total === 1
            ? `This document only has 1 ${noun}`
            : `${capitalised(noun)}s go from 1 to ${total}`,
      };
    }
    for (let number = start; number <= end; number += 1) {
      selected.add(number);
    }
  }
  return { pages: Array.from(selected).toSorted((a, b) => a - b) };
}
