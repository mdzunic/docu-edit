// The deck history panel's typed transforms (SPEC-103), beside
// DeckHistoryPanel.vue the way transitionStyles.ts sits beside the
// transitions panel: the component's script is linted by espree, so its
// runtime-declared `controller` arrives untyped and every label it shows is
// computed here, framework-free and unit-tested.
import { formatTimestamp } from "../../lib/format";
import { now } from "../../lib/clock";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import type { DeckVersionListEntry } from "../../lib/presentationVersions";
import type { DeckHistory } from "../../views/presentation/useDeckHistory";
import { filmstripThumbScale, stageFields, type StageField } from "./slideRender";

/** The panel's copy about what a restore's undo brings back (SPEC-103 AC-43):
 * usePresentationUndo's snapshot holds the slides and the deck settings it
 * clones, never the package — so the way all the way back is the version the
 * restore saved first. */
export const RESTORE_UNDO_NOTE =
  "Restoring replaces the whole deck. Ctrl+Z afterwards brings back only the slides and their settings — " +
  "to undo a restore completely, restore the “Before restore” version it saved first.";

/** How wide the preview's read-only stage draws, in px. */
export const PREVIEW_STAGE_WIDTH_PX = 320;

export function asDeckHistory(value: unknown): DeckHistory {
  return value as DeckHistory;
}

export function asVersionEntries(value: unknown): DeckVersionListEntry[] {
  return (value ?? []) as DeckVersionListEntry[];
}

/** "1 slide", "12 slides". */
export function slideCountLabel(count: number): string {
  return count === 1 ? "1 slide" : `${count} slides`;
}

/** The slide-count change against the previous version. */
export function slideDeltaLabel(delta: number): string {
  if (delta === 0) {
    return "no slide change";
  }
  const amount = Math.abs(delta);
  return `${delta > 0 ? "+" : "−"}${amount} ${amount === 1 ? "slide" : "slides"}`;
}

/** What took the snapshot, when it is worth saying. */
export function triggerLabel(entry: DeckVersionListEntry): string {
  switch (entry.trigger) {
    case "pre-restore":
      return "Before restore";
    case "post-restore":
      return "After restore";
    default:
      return "";
  }
}

export function versionTime(entry: DeckVersionListEntry): string {
  return formatTimestamp(new Date(entry.timestamp));
}

/** "Download .pptx" / "Download .odp". */
export function downloadLabel(format: string): string {
  return `Download .${format === "odp" ? "odp" : "pptx"}`;
}

/** The version a preview shows, by its name or its time. */
export function previewTitle(entries: DeckVersionListEntry[], versionId: string | null): string {
  const entry = entries.find((e) => e.id === versionId);
  return entry ? (entry.name ?? versionTime(entry)) : "";
}

export function previewSlide(model: PresentationModel | null, index: number): Slide | null {
  return model?.slides[index] ?? null;
}

export function previewScale(model: PresentationModel): number {
  return filmstripThumbScale(model.slideWidthCm, PREVIEW_STAGE_WIDTH_PX);
}

export function previewFields(model: PresentationModel, index: number): StageField[] {
  return stageFields(model, index, now());
}
