// Typed helpers for HeaderFooterDialog.vue's runtime bindings (SPEC-093): the
// espree convention bars type syntax in `.vue` scripts, so the checkbox/text
// event readers live here (the deckSetupFields.ts pattern).
export { eventValue } from "./deckSetupFields";

/** The `checked` of a checkbox change event, or false. */
export function eventChecked(event: Event): boolean {
  const target = event.target;
  return target instanceof HTMLInputElement ? target.checked : false;
}

/** One row of the date-format `<select>`: the catalogue pattern (value) and its
 * preview against today's instant (label). */
export interface DateFormatOption {
  value: string;
  label: string;
}

/** Typed pass-through for the dialog's runtime option list prop. */
export function asDateFormatOptions(value: unknown): DateFormatOption[] {
  return value as DateFormatOption[];
}
