// Which formulas depend on row visibility (SPEC-127 "SUBTOTAL and
// AGGREGATE"): a formula "calls a visibility-sensitive function" when its
// tokens hold a call to a registered function marked `visibilitySensitive`
// (SUBTOTAL, AGGREGATE). The answer is memoised per formula string, and is
// both how SUBTOTAL and AGGREGATE recognise a nested cell to skip and how
// recalc.ts's visibilitySensitiveCells finds what hiding or showing rows
// must recalculate.
import { getFunction } from "./registry";
import { tokenize } from "./tokens";

/** Formula strings remembered before the memo starts over. */
const MEMO_LIMIT = 50_000;

const memo = new Map<string, boolean>();

/** Whether a canonical formula calls SUBTOTAL, AGGREGATE or another
 * visibility-sensitive function, anywhere in it. */
export function callsVisibilitySensitive(formula: string): boolean {
  let calls = memo.get(formula);
  if (calls === undefined) {
    calls = tokenize(formula).some((token) => token.kind === "function" && getFunction(token.text)?.visibilitySensitive === true);
    if (memo.size >= MEMO_LIMIT) {
      memo.clear();
    }
    memo.set(formula, calls);
  }
  return calls;
}
