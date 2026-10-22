// The typed coercion the PresentationToolbar's espree-linted script cannot
// express inline (no TS casts in .vue): the untyped `state` prop back to the
// command layer's SelectionFormatState.
import type { SelectionFormatState } from "../../lib/presentation/text";

export function asFormatState(value: unknown): SelectionFormatState | null {
  return (value ?? null) as SelectionFormatState | null;
}

/** aria-pressed for a toggle: "mixed" for an indeterminate selection (AC-24). */
export function togglePressed(state: SelectionFormatState | null, field: string): "true" | "false" | "mixed" {
  const value = state ? (state as unknown as Record<string, unknown>)[field] : undefined;
  if (value === "mixed") return "mixed";
  return value === true ? "true" : "false";
}

export function alignIsActive(state: SelectionFormatState | null, value: string): boolean {
  return state?.align === value;
}
