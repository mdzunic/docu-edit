// The deck search's results-list rows (SPEC-098): DeckSearchResults.vue's
// model, split out of useDeckSearch.ts like sheetFindRows.ts. One group per
// slide with a hit (its 1-based number for the chip), then one entry per hit:
// the ±40-character snippet split around the match so the list can emphasise
// it, the frame's name (frameAccessibleName, SPEC-082) and a scope tag
// ("Notes", "Table r2c3").
import type { PresentationModel } from "../../lib/presentation/model";
import { hitsBySlide, type DeckHit } from "../../lib/presentation/deckSearch";
import { frameAccessibleName } from "./useFrameManipulation";

export interface DeckResultEntry {
  /** Index into the full hit list (what activating it makes current). */
  index: number;
  before: string;
  match: string;
  after: string;
  frameName: string;
  scopeTag: string;
}

export interface DeckResultGroup {
  slideIndex: number;
  slideNumber: number;
  hits: DeckResultEntry[];
}

/** The scope tag a hit's row carries. */
export function scopeTag(hit: DeckHit): string {
  if (hit.scope === "notes") return "Notes";
  if (hit.scope === "cell") return `Table r${(hit.row ?? 0) + 1}c${(hit.column ?? 0) + 1}`;
  return "";
}

export function deckResultGroups(model: PresentationModel | null, hits: readonly DeckHit[]): DeckResultGroup[] {
  return hitsBySlide(hits).map((group) => ({
    slideIndex: group.slideIndex,
    slideNumber: group.slideIndex + 1,
    hits: group.hits.map(({ hit, index }) => {
      const frame = hit.frameId ? model?.slides[hit.slideIndex]?.frames.find((each) => each.id === hit.frameId) : undefined;
      const matchEnd = hit.snippetOffset + hit.end - hit.start;
      return {
        index,
        before: hit.snippet.slice(0, hit.snippetOffset),
        match: hit.snippet.slice(hit.snippetOffset, matchEnd),
        after: hit.snippet.slice(matchEnd),
        frameName: frame ? frameAccessibleName(frame) : "Speaker notes",
        scopeTag: scopeTag(hit),
      };
    }),
  }));
}

/** Narrows DeckSearchResults.vue's rows (the .vue parser bars casts). */
export function asDeckResultGroups(value: unknown): DeckResultGroup[] {
  return value as DeckResultGroup[];
}
