// Typed shapes for the paginated canvas's precomputed render entries
// (SPEC-028): DocumentView's pagedViews computed builds them, DocPage.vue
// renders them. Kept in a lib so the espree-parsed `.vue` scripts share the
// types through a typed default seed (see styles.noStyleEntries for the
// same pattern).
import { isTable, type Block, type ListDefinition, type Shape, type TrackChangesDisplayMode } from "./documentModel";
import { headingTag, paragraphClassName, paragraphStyle, runKey, runStyle } from "./renderFormatting";
import { markersForBlocks } from "./lists";
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
  /** Note reference mark (SPEC-031) or mail-merge chip (SPEC-046): the
   * model placeholder text the rendered label stands for — becomes the
   * span's data-caret-atom, so caret mapping counts it as one character
   * (see lib/domCaret.ts). */
  atomText?: string;
  /** The atom span's CSS class — "doc-note-ref" when unset, "doc-merge-chip"
   * for a merge chip (SPEC-046). */
  atomClass?: string;
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
  kind: "paragraph" | "table" | "image" | "shape" | "chart" | "break" | "toc" | "marker";
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
   * paragraphs (or, for a shape entry, its text frame's list paragraphs)
   * derive defined-list markers/indents from. */
  listDefinitions?: ListDefinition[];
  imageSrc?: string;
  imageSelected?: boolean;
  /** Index into the document-order drawing list (SPEC-051): the `<index>`
   * in the shape's ARIA proxy's `data-testid="drawing-<index>"`. */
  drawingIndex?: number;
  shapeSelected?: boolean;
  /** Whether this shape's text frame is in edit mode (SPEC-051 AC-17,
   * AC-18) — only true shapeSelected's own entry ever sets this. */
  shapeEditing?: boolean;
  /** A chart entry's own selection state (SPEC-052) — mirrors shapeSelected;
   * a chart shares the same document-order `drawingIndex` counter shapes do
   * (SPEC-051's "Rendering and test surface"), never its own. */
  chartSelected?: boolean;
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

/** A shape's text frame rendered as DocBlockEntries entries (SPEC-051
 * AC-17..AC-21): paragraphs (each with its own self-contained list
 * numbering — markersForBlocks scoped to just this text frame, not shared
 * with the body) and tables, the "same Block content... including lists and
 * tables" AC-18 requires. Character formatting renders via the same
 * runStyle every other run does; hyperlink/note/merge-chip/track-changes
 * presentation and formatting marks are out of scope for a shape's text (a
 * plain styled span renders their text, just not the special-cased markup —
 * see the implement summary). Framework-free like the rest of this module,
 * so DocShapeBlock.vue and its own tests share one source of truth for what
 * a shape's text looks like. */
export function shapeTextRenderEntries(shape: Shape, listDefinitions: readonly ListDefinition[] = []): RenderEntry[] {
  const blocks = shape.text?.blocks ?? [];
  const markers = markersForBlocks(blocks, listDefinitions);
  const entries: RenderEntry[] = [];
  blocks.forEach((block, blockIndex) => {
    if (isTable(block)) {
      entries.push({
        key: `shape-text:${blockIndex}`,
        memo: [block, blockIndex],
        kind: "table",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: blockIndex,
        listDefinitions: [...listDefinitions],
      });
      return;
    }
    const paragraph = block.type === "paragraph" ? block : { type: "paragraph" as const, runs: [] };
    const styleObj = { ...paragraphStyle(paragraph, listDefinitions) };
    const renderRuns: RenderRun[] = paragraph.runs.map((run, runIndex) => ({
      key: runKey(run, runIndex),
      href: "",
      styleObj: runStyle(run.formatting),
      text: run.text,
      segments: null,
      lang: run.formatting.lang ?? "",
    }));
    entries.push({
      key: `shape-text:${blockIndex}`,
      memo: [block, blockIndex, markers[blockIndex]],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: paragraphClassName(paragraph),
      styleObj,
      dataSection: 0,
      dataBlock: blockIndex,
      listMarker: markers[blockIndex] || undefined,
      listLevel: paragraph.list ? paragraph.list.level : undefined,
      runs: renderRuns,
      needsBr: renderRuns.length === 0,
    });
  });
  return entries;
}
