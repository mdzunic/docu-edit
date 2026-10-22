// Typed helpers for DeckPrintDialog.vue (SPEC-104) — the `.vue` script is
// espree-linted, so its option tables and DOM-event readers live here, the
// deckSetupFields.ts split.
import { PAPER_SIZE_NAMES } from "../../lib/core/pageSetup";
import { DECK_PRINT_LAYOUTS } from "../../lib/presentation/printLayout";
import { PNG_SCALES } from "../../lib/presentation/deckExport";

/** The layout select's options (Slides, Notes pages, Handout 1…9 per page). */
export const LAYOUT_OPTIONS = DECK_PRINT_LAYOUTS.map(({ value, label }) => ({ value, label }));

/** The paper select's options, from core/pageSetup's catalogue (A4 first). */
export const PAPER_OPTIONS = PAPER_SIZE_NAMES.filter((name) => name !== "Custom").map((name) => ({
  value: name,
  label: name,
}));

/** The PNG scale select's options: 1×, 2×, 4×. */
export const SCALE_OPTIONS = PNG_SCALES.map((scale) => ({ value: String(scale), label: `${scale}×` }));

/** The two output tabs. */
export const DECK_PRINT_TABS = [
  { id: "pdf", label: "PDF and print", testid: "deck-print-tab-pdf" },
  { id: "images", label: "Images", testid: "deck-print-tab-images" },
];

/** Whether a checkbox change event left its box checked. */
export function eventChecked(event: Event): boolean {
  return event.target instanceof HTMLInputElement && event.target.checked;
}
