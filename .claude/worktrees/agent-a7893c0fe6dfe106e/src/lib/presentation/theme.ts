// Deck-theme commands (SPEC-094): applying a theme to a presentation, the
// per-format dirty rules, and the ref-usage predicate the odp writer and the
// panel share. Framework-free; the undo/commit choreography lives in
// `src/views/presentation/useDeckTheme.ts`.
import { DEFAULT_THEME, isThemeColor, isThemeFont, type DocumentTheme } from "../core/theme";
import type { DeckTheme, PresentationModel, Slide, SlideFrame } from "./model";
import { markSlideEdited } from "./text";

/** The theme a deck resolves against: its own, else Office (SPEC-094 AC-7). */
export function deckThemeOf(model: Pick<PresentationModel, "theme"> | null | undefined): DocumentTheme {
  return model?.theme ?? DEFAULT_THEME;
}

/** A deep, plain-data clone of a `DeckTheme` (for undo snapshots and safe
 * mutation — the theme is small, so a JSON round-trip is fine). */
export function cloneDeckTheme(theme: DeckTheme): DeckTheme {
  return JSON.parse(JSON.stringify(theme)) as DeckTheme;
}

function frameUsesTheme(frame: SlideFrame): boolean {
  if ((frame.kind === "text" || frame.kind === "shape") && frame.fill?.type === "solid" && isThemeColor(frame.fill.color)) {
    return true;
  }
  if ("stroke" in frame && frame.stroke && isThemeColor(frame.stroke.color)) {
    return true;
  }
  const paragraphs =
    frame.kind === "text" ? frame.paragraphs : frame.kind === "shape" ? frame.paragraphs ?? [] : [];
  for (const para of paragraphs) {
    for (const run of para.runs) {
      if (isThemeColor(run.color) || isThemeFont(run.fontFamily)) {
        return true;
      }
    }
  }
  if (frame.kind === "table") {
    for (const row of frame.rows) {
      for (const cell of row.cells) {
        if (cell.fill?.type === "solid" && isThemeColor(cell.fill.color)) {
          return true;
        }
        for (const para of cell.paragraphs) {
          for (const run of para.runs) {
            if (isThemeColor(run.color) || isThemeFont(run.fontFamily)) {
              return true;
            }
          }
        }
      }
    }
  }
  if (frame.kind === "group") {
    return frame.children.some(frameUsesTheme);
  }
  return false;
}

/** Whether any of the slide's own content carries a scheme reference — the
 * slides the odp writer must re-serialise on a theme change (their automatic
 * styles are theme-parented), and the panel's "this slide tracks the theme". */
export function slideUsesTheme(slide: Slide): boolean {
  if (slide.background?.type === "solid" && isThemeColor(slide.background.color)) {
    return true;
  }
  return slide.frames.some(frameUsesTheme);
}

/** Applies a theme's scheme + fonts to the deck (SPEC-094): `model.theme`
 * takes the new values (keeping the package wiring — `colorMap`, `partName`)
 * and is marked `dirty` so the writers rewrite the theme part in place.
 *
 * No slide ref changes — a reference tracks the theme by construction, and an
 * explicit literal keeps exactly the value it had. Per-format dirtying: pptx
 * dirties NO slide (the slide parts' refs are unchanged bytes); odp marks
 * every ref-carrying slide edited, because its pages inline the resolved
 * literal beside the `de-theme-*` parent link. */
export function applyTheme(model: PresentationModel, theme: DocumentTheme): void {
  const next: DeckTheme = {
    name: theme.name,
    colors: { ...theme.colors },
    fonts: { ...theme.fonts },
    dirty: true,
  };
  if (model.theme?.colorMap) {
    next.colorMap = { ...model.theme.colorMap };
  }
  if (model.theme?.partName) {
    next.partName = model.theme.partName;
  }
  model.theme = next;
  if (model.format === "odp") {
    for (const slide of model.slides) {
      if (slideUsesTheme(slide)) {
        markSlideEdited(slide, "odp");
      }
    }
  }
}

/** Sets one scheme slot on the deck's theme (the Customise inputs), applying
 * through `applyTheme` so the same dirty rules hold. */
export function setThemeColor(model: PresentationModel, slot: keyof DocumentTheme["colors"], hex: string): void {
  const current = deckThemeOf(model);
  applyTheme(model, { ...current, colors: { ...current.colors, [slot]: hex } });
}

/** Sets the major/minor theme font (the Customise inputs). */
export function setThemeFont(model: PresentationModel, which: "major" | "minor", family: string): void {
  const current = deckThemeOf(model);
  applyTheme(model, { ...current, fonts: { ...current.fonts, [which]: family } });
}
