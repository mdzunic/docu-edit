// Reading a model's reply (SPEC-115). A reply is untrusted text: it is parsed
// with JSON.parse inside a try and nothing else — no eval, no Function, no
// dynamic import, no fetch of anything it names — and then checked by the
// deterministic validators before anything is shown or built. Every failure
// is a DocumentGenerationError naming the offending path, which is what the
// one repair sends back.
import type { DeckDropRecord } from "../generate/deckDrops";
import { deckFromDescription } from "../generate/deckDescription";
import { DocumentGenerationError } from "../generate/errors";
import { planToDescription, validatePlan, validateRedesign, type DeckPlan, type DeckRedesign } from "./plan";

export type CheckedReply<T> = { ok: true; value: T; dropped: DeckDropRecord[] } | { ok: false; error: DocumentGenerationError };

/** The one way a reply is read: JSON.parse inside a try. */
export function parseReply(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new DocumentGenerationError(`reply: not valid JSON (${error instanceof Error ? error.message : String(error)})`);
  }
}

function checked<T>(run: () => { value: T; dropped: DeckDropRecord[] }): CheckedReply<T> {
  try {
    return { ok: true, ...run() };
  } catch (error) {
    if (error instanceof DocumentGenerationError) {
      return { ok: false, error };
    }
    throw error;
  }
}

/** A New-flow reply: parseable JSON, a plan validatePlan accepts, and one
 * whose description deckFromDescription builds — all before the plan step
 * shows it. */
export function checkPlanReply(text: string): CheckedReply<DeckPlan> {
  return checked(() => {
    const { plan, dropped } = validatePlan(parseReply(text));
    deckFromDescription(planToDescription(plan));
    return { value: plan, dropped };
  });
}

/** A Redesign reply: parseable JSON and an object with a `slides` array —
 * everything else validateRedesign normalises to the open deck. */
export function checkRedesignReply(text: string, slideCount: number): CheckedReply<DeckRedesign> {
  return checked(() => {
    const { redesign, dropped } = validateRedesign(parseReply(text), slideCount);
    return { value: redesign, dropped };
  });
}
