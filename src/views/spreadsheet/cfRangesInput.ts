// "Applies to" text parsing (SPEC-129 "The side panel" — ConditionalFormatsPane's
// editable `cf-rule-applies-to`): A1 ranges separated by spaces or commas, the
// way a person types them rather than the single-space-only sqref grammar
// address.ts's parseSqref reads from a file. Framework-free and pure.
import { formatRange, parseRange, type GridRange } from "../../lib/spreadsheet/address";

export const APPLIES_TO_EMPTY_MESSAGE = "Enter at least one range";
export const APPLIES_TO_INVALID_MESSAGE = "Not a valid range";

/** `ranges` as the field shows them: "B2:B13 D1". */
export function formatAppliesTo(ranges: readonly GridRange[]): string {
  return ranges.map(formatRange).join(" ");
}

export type AppliesToResult = { ranges: GridRange[] } | { error: string };

/** The ranges `text` names, space- or comma-separated — "B2:B13, D1" and
 * "B2:B13 D1" both read two ranges. Empty entries (repeated separators) are
 * skipped; any entry that is not a valid A1 range is a refusal naming it,
 * and nothing is saved (ConditionalFormatsPane's rule, per the spec). */
export function parseAppliesTo(text: string): AppliesToResult {
  const entries = text
    .split(/[\s,]+/)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
  if (entries.length === 0) {
    return { error: APPLIES_TO_EMPTY_MESSAGE };
  }
  const ranges: GridRange[] = [];
  for (const entry of entries) {
    const range = parseRange(entry);
    if (!range) {
      return { error: `${APPLIES_TO_INVALID_MESSAGE}: "${entry}"` };
    }
    ranges.push(range);
  }
  return { ranges };
}
