// Typed shapes for the paginated canvas's precomputed render entries
// (SPEC-028): DocumentView's pagedViews computed builds them, DocPage.vue
// renders them. Kept in a lib so the espree-parsed `.vue` scripts share the
// types through a typed default seed (see styles.noStyleEntries for the
// same pattern).
import type { Block, ListDefinition, TrackChangesDisplayMode } from "./documentModel";
import type { MarkSegment } from "./marks";
import type { PageColumns } from "./pagination";

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
  /** Note reference mark (SPEC-031): the model placeholder text the
   * rendered label stands for — becomes the span's data-caret-atom, so
   * caret mapping counts it as one character (see lib/domCaret.ts). */
  atomText?: string;
  /** The mark's note id (SPEC-031) — click-to-jump's target (AC-15). */
  noteId?: string;
  /** The mark's hover tooltip: the note's text (SPEC-031 AC-14). */
  title?: string;
  /** This run's effective language (SPEC-035 AC-14) — the run's own
   * `RunFormatting.lang`, else the document default, resolved once here so
   * the renderer never re-derives it. Bound as the DOM `lang` attribute for
   * assistive technology; always set, unlike the model field it derives
   * from (see lib/language.ts's effectiveLanguage). */
  lang: string;
  /** SPEC-039 AC-9's "No markup"/"Original" collapse: when set, this run
   * renders NO visible text (an empty, `contenteditable="false"` atom) but
   * still carries its full model text via `data-caret-atom` (this value),
   * so lib/domCaret.ts's offset mapping counts it exactly as if the text
   * were showing (see lib/renderFormatting.ts's trackChangeRunHidden). */
  caretAtom?: string;
}

/** One page item's precomputed rendering — a paragraph (possibly a slice of
 * one), a table, an image, a break block, a table of contents or a
 * section-break marker. */
export interface RenderEntry {
  key: string;
  /** The v-memo array: every value the entry's DOM depends on. */
  memo: unknown[];
  kind: "paragraph" | "table" | "image" | "break" | "toc" | "marker";
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
  /** The model's multilevel definitions (SPEC-032) a table entry's cell
   * paragraphs derive defined-list markers/indents from. */
  listDefinitions?: ListDefinition[];
  imageSrc?: string;
  imageSelected?: boolean;
  /** Staleness flags a TOC entry's update control reacts to (SPEC-033
   * AC-10, AC-11). */
  tocStaleHeadings?: boolean;
  tocStalePagination?: boolean;
  runs?: RenderRun[];
  needsBr?: boolean;
  /** A body heading that owns content and so shows a collapse chevron
   * (SPEC-037 AC-16); only on the heading's first slice. */
  outlineToggle?: boolean;
  /** Whether that heading is currently collapsed in the body. */
  outlineCollapsed?: boolean;
  /** Which column of a multi-column page this entry flows into (SPEC-030);
   * absent on single-column pages. */
  column?: number;
  /** The review pane's display mode (SPEC-039 AC-9) — a table entry's cell
   * runs need it passed down since DocTableBlock renders its own runs
   * rather than through RenderRun, and an image entry needs it to hide a
   * tracked-inserted/deleted image in "Original"/"No markup". */
  trackChangesDisplayMode?: TrackChangesDisplayMode;
  /** Whether this paragraph carries a tracked change (SPEC-039 AC-5's
   * margin change bar) — suppressed in "No markup"/"Original" display
   * mode, which render as if the document already reflects one side of
   * every change. */
  hasTrackedChange?: boolean;
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

/** Typed pass-through for a single entry prop (espree bars the annotation). */
export function asRenderEntry(value: unknown): RenderEntry {
  return value as RenderEntry;
}

/** Typed pass-through for DocPage's `columns` prop (SPEC-030) arriving as a
 * runtime-declared Object — the page's multi-column geometry, or null on a
 * single-column page. */
export function asPageColumns(value: unknown): PageColumns | null {
  return (value ?? null) as PageColumns | null;
}
