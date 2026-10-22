// Typed helpers for DeckSetupDialog.vue (SPEC-090). The dialog's `.vue` script
// is espree-linted (runtime `defineProps`, no TS annotations), so the pieces
// that need a real static type — the metadata field table and the DOM-event
// readers — live here, the same split PropertiesDialog uses via
// `lib/customProperties` and `documentModel`.
import type { DocumentProperties } from "../../lib/core/documentProperties";

/** The six editable metadata fields, in display order — each mapped to the
 * `DocumentProperties` key it edits, its label and its testid (AC-18). */
export interface DeckPropertyField {
  key: keyof DocumentProperties;
  label: string;
  testid: string;
  multiline?: boolean;
}

export const DECK_PROPERTY_FIELDS: DeckPropertyField[] = [
  { key: "title", label: "Title", testid: "deck-properties-title" },
  { key: "subject", label: "Subject", testid: "deck-properties-subject" },
  { key: "keywords", label: "Keywords", testid: "deck-properties-keywords" },
  { key: "comments", label: "Comments", testid: "deck-properties-comments", multiline: true },
  { key: "author", label: "Author", testid: "deck-properties-author" },
  { key: "lastModifiedBy", label: "Last modified by", testid: "deck-properties-last-modified-by" },
];

/** The string value of an `<input>`/`<textarea>`/`<select>` change event. */
export function eventValue(event: Event): string {
  const target = event.target;
  if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement) {
    return target.value;
  }
  return "";
}
