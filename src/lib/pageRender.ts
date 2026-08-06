// Typed shapes for the paginated canvas's precomputed render entries
// (SPEC-028): DocumentView's pagedViews computed builds them, DocPage.vue
// renders them. Kept in a lib so the espree-parsed `.vue` scripts share the
// types through a typed default seed (see styles.noStyleEntries for the
// same pattern).
import type { Block } from "./documentModel";
import type { MarkSegment } from "./marks";

/** One run's precomputed rendering: link fields resolved, style object
 * built, and — with formatting marks on — the text split into mark
 * segments. */
export interface RenderRun {
  key: string;
  /** Non-empty renders an <a>; empty renders a <span>. */
  href: string;
  target?: string;
  rel?: string;
  styleObj: Record<string, string>;
  text: string;
  segments: MarkSegment[] | null;
}

/** One page item's precomputed rendering — a paragraph (possibly a slice of
 * one), a table, an image, a break block or a section-break marker. */
export interface RenderEntry {
  key: string;
  /** The v-memo array: every value the entry's DOM depends on. */
  memo: unknown[];
  kind: "paragraph" | "table" | "image" | "break" | "marker";
  /** Element tag for paragraph/break/marker entries ("p", "h2", "div"). */
  tag?: string;
  className?: string;
  styleObj?: Record<string, string>;
  /** contenteditable="false" (breaks and markers — atoms, not text). */
  uneditable?: boolean;
  dataSection?: number;
  dataBlock?: number;
  /** Character offset this slice's text starts at (split paragraphs). */
  sliceStart?: string;
  /** "1" when the paragraph continues on a later page. */
  sliceMore?: string;
  breakKind?: string;
  sectionMarker?: string;
  markerLabel?: string;
  listMarker?: string;
  listLevel?: number;
  blockProp?: Block;
  sectionProp?: number;
  blockIndexProp?: number;
  imageSrc?: string;
  imageSelected?: boolean;
  runs?: RenderRun[];
  needsBr?: boolean;
}

/** Typed empty seed for DocPage's entries prop (espree bars the
 * annotation in `.vue` scripts). */
export function noRenderEntries(): RenderEntry[] {
  return [];
}

/** Typed pass-through for the entries prop arriving as a runtime-declared
 * Array (see styles.asStyleEntries for the same pattern). */
export function asRenderEntries(value: unknown): RenderEntry[] {
  return value as RenderEntry[];
}
