// The deck side's drop report (SPEC-106): SPEC-008's `DropRecord` SHAPE —
// what was dropped, where, and why — re-declared here over its own closed
// union. `src/lib/opaque.ts`'s `DropKind` is the text product's closed union
// (it has no member for a shape, a transition or a fenced block), and
// importing it would pull documentModel.ts across the product boundary, so
// the deck builder and slide Markdown report through this type instead.
// Plain data, no imports.

/** Every construct the deck path can drop, closed. The first group is what a
 * deck carries and Markdown cannot (`writeSlideMarkdown`), the second what
 * Markdown spells and the slide grammar does not (`parseSlideMarkdown`);
 * `runFormatting` serves both directions and the builder's run mapping. */
export type DeckDropKind =
  | "shape"
  | "group"
  | "foreignFrame"
  | "table"
  | "chart"
  | "media"
  | "transition"
  | "animation"
  | "background"
  | "fill"
  | "stroke"
  | "runFormatting"
  | "slideGeometry"
  | "blockQuote"
  | "codeFence"
  | "rawHtml"
  | "heading"
  | "remoteImage"
  | "frontMatterKey";

/** One dropped construct. `location` names the slide by 1-based number
 * ("slide 4"), or "front matter" / "deck" for what belongs to neither. */
export interface DeckDropRecord {
  kind: DeckDropKind;
  location: string;
  reason: string;
}

/** The 1-based location label a slide's drop records carry. */
export function slideLocation(slideIndex: number): string {
  return `slide ${slideIndex + 1}`;
}
