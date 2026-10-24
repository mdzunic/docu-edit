// Typed helpers for the deck designer's .vue files (SPEC-115). Their scripts
// are espree-linted — runtime `defineProps`, no TS annotations or casts (see
// AppShell.vue) — so the coercions of their untyped Object props and the
// DOM-event readers live here, as deckSetupFields.ts does for DeckSetupDialog.
import type { DeckDesigner, RedesignHost } from "./useDeckDesigner";

/** The `controller` prop every designer step component takes. */
export function asDeckDesigner(value: unknown): DeckDesigner {
  return value as DeckDesigner;
}

/** The `host` prop of a Redesign dialog: the open deck's model and undo
 * choreography, or null for the New flow. */
export function asRedesignHost(value: unknown): RedesignHost | null {
  return (value ?? null) as RedesignHost | null;
}

/** The string value of an `<input>`/`<textarea>`/`<select>` event. */
export function fieldValue(event: Event): string {
  const target = event.target;
  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement) {
    return target.value;
  }
  return "";
}

/** Whether an `<input type="checkbox">` event left it ticked. */
export function fieldChecked(event: Event): boolean {
  return event.target instanceof HTMLInputElement && event.target.checked;
}

/** The plan step's reorder shortcut: Alt+↑ is -1, Alt+↓ is +1, anything
 * else 0. */
export function reorderDelta(event: KeyboardEvent): number {
  if (!event.altKey || event.ctrlKey || event.metaKey) {
    return 0;
  }
  return event.key === "ArrowUp" ? -1 : event.key === "ArrowDown" ? 1 : 0;
}

/** A DeckDropRecord (an untyped list item in a .vue template) as the "What
 * we changed" list shows it: where, capitalised, and why. */
export function dropLine(record: unknown): { at: string; why: string } {
  const { location, reason } = (record ?? {}) as { location?: unknown; reason?: unknown };
  const where = String(location ?? "");
  return { at: where === "deck" ? "Deck" : where.replace(/^slide /, "Slide "), why: String(reason ?? "") };
}
