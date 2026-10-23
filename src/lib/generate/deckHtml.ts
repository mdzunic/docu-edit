// HTML export through the generation API (SPEC-106, resolving SPEC-088's
// deferred follow-up): a thin, validated wrapper over
// presentation/htmlExport.ts's DOM-free `exportPresentationHtml`, called
// unchanged — the same single self-contained document with no external
// requests — so a generated deck reaches HTML without a browser. SPEC-088
// declares no default for `includeNotes` (its options are required); this API
// chooses `false` and passes it through explicitly. SPEC-088 throws a plain
// Error on a deck with no slides; this API refuses that deck first, with a
// `DocumentGenerationError` naming it.
import { exportPresentationHtml } from "../presentation/htmlExport";
import type { PresentationModel } from "../presentation/model";
import { DocumentGenerationError } from "./errors";
import { expectArray, expectBoolean, expectPlainObject, rejectUnknownKeys } from "./validate";

export interface DeckHtmlOptions {
  /** Embed the speaker-notes overlay (toggled with `n` in the viewer).
   * Defaults to `false`. */
  includeNotes?: boolean;
}

/** A deck as one self-contained `.html` document (SPEC-088's export). */
export function exportDeckHtml(model: PresentationModel, options: DeckHtmlOptions = {}): string {
  const obj = expectPlainObject(options, "options");
  rejectUnknownKeys(obj, ["includeNotes"], "options");
  const includeNotes = obj.includeNotes === undefined ? false : expectBoolean(obj.includeNotes, "options.includeNotes");
  const slides = expectArray(expectPlainObject(model, "deck").slides, "deck.slides");
  if (slides.length === 0) {
    throw new DocumentGenerationError("deck: an empty deck (no slides) cannot be exported to HTML — add a slide first");
  }
  return exportPresentationHtml(model, { includeNotes });
}
