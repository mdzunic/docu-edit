// The deck history panel's typed transforms (SPEC-103), beside
// DeckHistoryPanel.vue the way transitionStyles.ts sits beside the
// transitions panel: the component's script is linted by espree, so its
// runtime-declared `controller` arrives untyped and every label it shows is
// computed here, framework-free and unit-tested.
//
// SPEC-138 moved the panel itself into the product-neutral HistoryPanel.vue;
// deckPanelController hands it the deck's two facts — the slide count and
// change, and the .pptx/.odp download — beside the controller.
import { computed } from "vue";
import { now } from "../../lib/clock";
import type { VersionIndexEntry } from "../../lib/core/versionChain";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import type { DeckVersionListEntry } from "../../lib/presentationVersions";
import type { DeckHistory } from "../../views/presentation/useDeckHistory";
import type { HistoryPanelController } from "../historyPanelRender";
import { triggerLabel, versionTime } from "../historyPanelRender";
import { filmstripThumbScale, stageFields, type StageField } from "./slideRender";

export { triggerLabel, versionTime };

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

/** "Download .pptx" / "Download .odp". */
export function downloadLabel(format: string): string {
  return `Download .${format === "odp" ? "odp" : "pptx"}`;
}

/** "4 slides · +2 slides" — a deck version's counts as the panel lists them. */
export function deckStatLabel(entry: VersionIndexEntry): string {
  return `${slideCountLabel(entry.slides ?? 0)} · ${slideDeltaLabel((entry as Partial<DeckVersionListEntry>).slideDelta ?? 0)}`;
}

/** The deck's controller as HistoryPanel reads it: its own fields, plus the
 * two the deck states — the stat line and the download's label. */
export function deckPanelController(c: DeckHistory): HistoryPanelController {
  return {
    ...c,
    statLabel: deckStatLabel,
    downloadLabels: computed(() => ({ button: downloadLabel(c.format.value), title: downloadLabel(c.format.value) })),
  };
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
