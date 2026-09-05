// The pagination render pipeline (SPEC-028, SPEC-029, SPEC-030, SPEC-031).
// Given the pagination controller's layout and the model, this builds the exact
// render data DocPage mounts: one `PageView` per laid-out page (body entries in
// the page's item order, plus header/footer bands and note areas), and the
// section-flattened `flatBlocks` list that carries the v-memo keys keeping
// large-document typing cheap.
//
// Everything here is read-only derivation from the session — it reacts to
// `doc` (replaced on every edit), pagination, note labels, and the selection/
// remount atoms — so it describes exactly what the canvas shows and never edits
// the model.
import { computed, type ComputedRef, type Ref } from "vue";
import {
  asImage,
  asParagraphBlock,
  asToc,
  asTof,
  createEmptyHeaderFooterContent,
  documentThemeOf,
  isBreak,
  isChart,
  isConnector,
  isImage,
  isMergeChipField,
  isShape,
  isTable,
  isToc,
  isTof,
  type ActiveHeaderFooterTarget,
  type ActiveNoteTarget,
  type Block,
  type DocumentModel,
  type NoteContent,
  type Paragraph,
  type Run,
  type TrackChangesDisplayMode,
} from "../../lib/documentModel";
import { resolveColor } from "../../lib/core/theme";
import { hasMixedPageSizes, pageBoxStyle } from "../../lib/printRange";
import { chipLabel } from "../../lib/mergeChips";
import {
  sectionColumnFlowWidthPx,
  type PageColumns,
  type PageItem,
  type PageLayout,
} from "../../lib/pagination";
import { listMarkers } from "../../lib/lists";
import { noNoteAreaItems, type NoteArea, type NoteAreaItem } from "../../lib/notesLayout";
import { hoistFloatingDrawingEntries, noRenderEntries, type RenderEntry, type RenderRun } from "../../lib/pageRender";
import { noteById } from "../../lib/notes";
import { headerFooterForPage } from "../../lib/pageNumbering";
import { fieldDisplayText, isReadOnlyField, UNSUPPORTED_FIELD_MESSAGE, type FieldContext } from "../../lib/fields";
import { now } from "../../lib/clock";
import {
  displayRunFormatting,
  isExternalLink,
  headingTag,
  paragraphClassName,
  paragraphHasTrackedChange,
  pageBorderStyle,
  paragraphStyle,
  runKey,
  runLinkHref,
  runStyle,
  trackChangeRunHidden,
  trackChangeRunStyle,
} from "../../lib/renderFormatting";
import { effectiveLanguage } from "../../lib/core/language";
import { markSegments, tabOnlySegments, type MarkSegment } from "../../lib/core/marks";
import type { ParagraphTabLayout } from "../../lib/tabLayout";
import {
  borderRuns,
  fragmentEdges,
  type BorderRunFlags,
  type InheritedPaintOf,
} from "../../lib/paragraphBorders";
import { paragraphPaintResolver } from "../../lib/styles";
import { splitRuns } from "../../lib/textEditing";
import type { CaretPosition } from "../../lib/textEditing";
import { tocHeadingsStale, tocPageNumbersStale } from "../../lib/toc";
import { tofEntriesStale, tofPageNumbersStale } from "../../lib/tof";
import { blockKey, headingContentLength } from "../../lib/outline";
import { mediaDataUri } from "../../lib/core/media";
import type { PaginationController } from "../../composables/usePagination";
import type { CellRange } from "../../lib/core/tables";
import type { ImageLocation } from "../../lib/images";
import type { ShapeLocation } from "../../lib/shapes";
import type { ChartLocation } from "../../lib/charts";
import type { EditorContext } from "./context";

/** One flattened block row — the section-flattened list a single v-for carries
 * (the v-memo keys that keep typing responsive on very large documents). */
export interface FlatEntry {
  key: string;
  block: Block;
  paragraph: Paragraph;
  isTable: boolean;
  isImage: boolean;
  isShape: boolean;
  isChart: boolean;
  isConnector: boolean;
  isToc: boolean;
  isTof: boolean;
  imageSelected: boolean;
  shapeSelected: boolean;
  /** Whether this exact shape is the one in text-editing mode (SPEC-051
   * AC-17, AC-18) — true only when it's also shapeSelected. */
  shapeEditing: boolean;
  /** This chart's own selection state (SPEC-052) — mirrors shapeSelected. */
  chartSelected: boolean;
  /** This shape/chart's 0-based index among every Shape/Chart block in the
   * document, in document order (SPEC-051's `data-testid="drawing-<index>"`,
   * shared by charts per SPEC-052 AC-24). */
  drawingIndex: number;
  section: number;
  blockIndex: number;
  marker: string | undefined;
  cellSel: string;
  imgSel: string;
}

/** A rendered note-area group (the footnote/endnote band shape). */
export interface NoteAreaView {
  kind: string;
  notes: Array<{
    noteId: string;
    kind: string;
    label: string;
    editable: boolean;
    entries: RenderEntry[];
  }>;
}

/** One rendered page — the object DocumentView's v-for hands to DocPage. */
export interface PageView {
  pageStyle: Record<string, string>;
  contentStyle: Record<string, string>;
  /** The section's page border as an absolutely-positioned box (SPEC-074
   * AC-22), or null when this page takes none. */
  pageBorderStyle: Record<string, string> | null;
  /** The section's page background (SPEC-074 AC-24), or null. Rendered as
   * its own layer behind the content so the print path can drop it (AC-25). */
  pageBackgroundStyle: Record<string, string> | null;
  entries: RenderEntry[];
  columns: PageColumns | undefined;
  virtual: boolean;
  dimmed: boolean;
  headerEditable: boolean;
  footerEditable: boolean;
  headerStyle: Record<string, string> | null;
  footerStyle: Record<string, string> | null;
  headerEntries: RenderEntry[];
  footerEntries: RenderEntry[];
  footnoteArea: (NoteAreaView & { style: Record<string, string> }) | null;
  endnoteArea: NoteAreaView | null;
}

export interface PaginationViewDeps extends EditorContext {
  pagination: PaginationController;
  noteLabelsMap: Ref<Map<string, string>>;
  noteTooltipsMap: Ref<Map<string, string>>;
  pageNumbers: Ref<number[]>;
  marksOn: Ref<boolean>;
  virtualizing: Ref<boolean>;
  viewportPageRange: Ref<{ start: number; end: number }>;
  caretPage: Ref<number>;
  anchorPage: Ref<number>;
  findMatchPage: Ref<number>;
  activeHeaderFooter: Ref<ActiveHeaderFooterTarget | null>;
  activeNote: Ref<ActiveNoteTarget | null>;
  cellSelection: Ref<CellRange | null>;
  imageSelection: Ref<ImageLocation | null>;
  /** The selected shape (SPEC-051), if any — mirrors imageSelection. */
  shapeSelection: Ref<ShapeLocation | null>;
  /** A Shift-click multi-selection of two or more shapes (SPEC-051 AC-34) —
   * every member renders selected, not just shapeSelection's own. Empty
   * outside multi-select mode. */
  shapeMultiSelection: Ref<ShapeLocation[]>;
  /** Whether shapeSelection's own shape is in text-editing mode (SPEC-051
   * AC-17, AC-18) — only that one shape's entry renders its text frame
   * `contenteditable`, real DOM the user's typing can reach. */
  editingShapeText: Ref<boolean>;
  /** The selected chart (SPEC-052), if any — mirrors shapeSelection; no
   * multi-select array (no chart group-selection gesture). */
  chartSelection: Ref<ChartLocation | null>;
  /** Body headings collapsed in the document body (SPEC-037 AC-16), keyed by
   * their "section:block" address — drives the chevron direction. */
  collapsedHeadings: Ref<ReadonlySet<string>>;
  remount: Ref<{ key: string; count: number; epoch: number }>;
  realizedSticky: WeakSet<object>;
  hfRemountEpoch: Ref<number>;
  noteRemountEpoch: Ref<number>;
  tocPageNumberFor: (model: DocumentModel, anchor: string) => number | null;
  /** The review pane's display-mode selector (SPEC-039 AC-9); drives
   * whether body runs show inline tracked-change markup. */
  trackChangesDisplayMode: Ref<TrackChangesDisplayMode>;
  /** Whether "View → Field shading" is on (SPEC-053 AC-11): every field run
   * paints its grey shading regardless of the caret. Shading never prints or
   * exports (the class is dropped in print CSS). */
  fieldShadingOn: Ref<boolean>;
  /** The reactive caret mirror (SPEC-053 AC-10): a field run shades when the
   * caret sits on it. Unlike `getCaret()` this is a ref, so a caret move
   * re-renders the paragraph it lands in (which `getCaret()`, being
   * non-reactive, would not). Read only through usePaginationView's
   * `fieldCaret` computed, never directly by the page render. */
  caretPosition: Ref<CaretPosition | null>;
}

/** The paint-time render of a SPEC-053 field run (everything but merge chips
 * and note refs, handled by their own branches): its computed value as a
 * caret atom, grey-shaded when field shading is on or the caret sits on it.
 * `page`/`numpages` resolve through the same `evaluateField` path, given the
 * page context. */
function fieldRenderRun(
  run: Run,
  runIndex: number,
  ctx: FieldContext,
  styleObj: Record<string, string>,
  lang: string,
  caretOnField: boolean,
  shadingOn: boolean,
): RenderRun {
  const field = run.field!;
  // Show the saved cached value on open, never a live recompute (AC-19/20);
  // only pagination-dependent fields evaluate live. See fieldDisplayText.
  const value = fieldDisplayText(field, ctx);
  const shaded = shadingOn || caretOnField;
  const readOnly = isReadOnlyField(field);
  const locked = "locked" in field && field.locked === true;
  const baseClass = `doc-field${readOnly ? " doc-field-readonly" : ""}${locked ? " doc-field-locked" : ""}`;
  const refTarget = field.kind === "ref" || field.kind === "pageRef" ? field.target : undefined;
  return {
    key: runKey(run, runIndex),
    href: "",
    styleObj,
    // A zero-width space keeps an empty-valued field a clickable one-char atom.
    text: value.length > 0 ? value : "​",
    segments: null,
    atomText: run.text,
    atomClass: shaded ? `${baseClass} doc-field-shaded` : baseClass,
    title: readOnly ? UNSUPPORTED_FIELD_MESSAGE : undefined,
    refTarget,
    lang,
  };
}

const SECTION_MARKER_LABELS: Record<string, string> = {
  continuous: "Section Break (Continuous)",
  nextPage: "Section Break (Next Page)",
};
const BREAK_LABELS: Record<string, string> = { page: "Page Break", column: "Column Break" };

/** The rendered data-URI src of an image block's media part (SPEC-024). */
function mediaSrc(model: DocumentModel | null, block: Block): string {
  if (!isImage(block) || !model) {
    return "";
  }
  return mediaDataUri(model.media?.find((part) => part.mediaId === asImage(block).mediaId));
}

/** Annotates a run's tab segments with their resolved box widths and leaders
 * (SPEC-054): the advance map is keyed by full-paragraph character index, so
 * the run's own start offset in the paragraph is added to each segment's
 * local offset. Mutates and returns the segments. */
function annotateTabSegments(
  segments: MarkSegment[],
  runGlobalStart: number,
  layout: ParagraphTabLayout,
): MarkSegment[] {
  let offset = 0;
  for (const segment of segments) {
    if (segment.kind === "tab") {
      const advance = layout.advances.get(runGlobalStart + offset);
      if (advance) {
        segment.widthPx = advance.widthPx;
        segment.leader = advance.leader;
      }
    }
    offset += segment.text.length;
  }
  return segments;
}

/** One page item's precomputed render entry for DocPage — or null for the
 * continuation slice of a block whose real metrics have not arrived yet. */
/** `borderRuns` per section, cached on the blocks array's identity
 * (SPEC-074 AC-12). Every command clones the containers it edits, so a
 * section whose blocks array is unchanged genuinely has unchanged runs — the
 * same identity rule `v-memo` relies on. A WeakMap so a replaced section is
 * collected with its entry.
 *
 * The runs also depend on the style catalogue, because a paragraph can
 * inherit its box from its style (AC-17) without carrying anything of its
 * own — and editing a style leaves every blocks array untouched. So the
 * entry records the catalogue it was computed against (a fresh array on
 * every style edit, see commands.customStyles.modifyStyle) and is recomputed
 * when that changes. */
interface CachedBorderRuns {
  styles: DocumentModel["styles"];
  runs: BorderRunFlags[];
}

const borderRunCache = new WeakMap<readonly Block[], CachedBorderRuns>();

function sectionBorderRuns(
  model: DocumentModel,
  blocks: readonly Block[],
  inheritedPaintOf: InheritedPaintOf,
): BorderRunFlags[] {
  const cached = borderRunCache.get(blocks);
  if (cached && cached.styles === model.styles) {
    return cached.runs;
  }
  const runs = borderRuns(blocks, inheritedPaintOf);
  borderRunCache.set(blocks, { styles: model.styles, runs });
  return runs;
}

/** Each page's 0-based ordinal within its own section. */
function ordinalsWithinSection(pages: readonly PageLayout[]): number[] {
  const ordinals: number[] = [];
  let currentSection = -1;
  let ordinal = 0;
  for (const page of pages) {
    if (page.sectionIndex !== currentSection) {
      currentSection = page.sectionIndex;
      ordinal = 0;
    }
    ordinals.push(ordinal);
    ordinal += 1;
  }
  return ordinals;
}

function pageBorderStyleFor(
  model: DocumentModel,
  page: PageLayout,
  pageIndexInSection: number,
): Record<string, string> | null {
  return pageBorderStyle(model.sections[page.sectionIndex]?.pageBorders, page.margins, pageIndexInSection ?? 0, documentThemeOf(model));
}

function pageBackgroundStyleFor(model: DocumentModel, page: PageLayout): Record<string, string> | null {
  const fill = model.sections[page.sectionIndex]?.pageBackground;
  return fill === undefined ? null : { backgroundColor: resolveColor(fill, documentThemeOf(model)) };
}

/** Every key paragraphBorderStyle can emit (SPEC-074). They go into the
 * v-memo signature because each one is decided by things that are NOT on the
 * block: the border run this paragraph belongs to (AC-12 — a neighbour
 * gaining or losing identical borders changes what THIS paragraph paints)
 * and where its fragment sits in a page split (AC-14..16). Listing all of
 * them rather than the visible subset means a run that merges or breaks
 * always repaints, even when only its collapsed margins change. */
const BORDER_MEMO_KEYS = [
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "paddingTop",
  "paddingBottom",
  "paddingLeft",
  "paddingRight",
  "marginTop",
  "marginBottom",
  "backgroundColor",
] as const;

function renderEntryFor(
  model: DocumentModel,
  flatEntry: FlatEntry,
  item: PageItem,
  marks: boolean,
  deps: PaginationViewDeps,
  fieldCaret: CaretPosition | null,
  inheritedPaintOf: InheritedPaintOf,
): RenderEntry | null {
  const block = flatEntry.block;
  const column = item.columnIndex;
  // The document theme joins every colour-bearing entry's v-memo signature
  // (SPEC-077 AC-15): applying a theme replaces model.theme but deliberately
  // keeps every block's identity (commands.theme.ts), so without this key
  // v-memo would skip the repaint that re-resolves the references.
  const themeSig = documentThemeOf(model);
  if (isBreak(block)) {
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, marks, column],
      kind: "break",
      tag: "div",
      className: "doc-break",
      uneditable: true,
      dataSection: flatEntry.section,
      dataBlock: flatEntry.blockIndex,
      breakKind: block.kind,
      markerLabel: BREAK_LABELS[block.kind] ?? "Break",
      column,
    } as RenderEntry;
  }
  if (flatEntry.isTable || flatEntry.isImage || flatEntry.isShape || flatEntry.isChart) {
    // SPEC-076: a table splits at row boundaries, so its "lines" are its
    // model rows and a page slice is the row window [startLine, endLine).
    // Only the first slice may render before the table has been measured —
    // a continuation with no metrics would repeat rows already drawn.
    const tableRowStart = flatEntry.isTable ? item.startLine ?? 0 : 0;
    const tableRowEnd = flatEntry.isTable && item.endLine !== undefined ? item.endLine : -1;
    return {
      key: tableRowStart > 0 ? `${flatEntry.key}@${tableRowStart}` : flatEntry.key,
      memo: [
        block,
        flatEntry.section,
        flatEntry.blockIndex,
        tableRowStart,
        tableRowEnd,
        flatEntry.marker,
        flatEntry.cellSel,
        flatEntry.imgSel,
        flatEntry.shapeSelected,
        flatEntry.shapeEditing,
        flatEntry.chartSelected,
        column,
        model.listDefinitions,
        deps.trackChangesDisplayMode.value,
        themeSig,
      ],
      kind: flatEntry.isTable ? "table" : flatEntry.isImage ? "image" : flatEntry.isShape ? "shape" : "chart",
      blockProp: block,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      listDefinitions: flatEntry.isTable || flatEntry.isShape ? model.listDefinitions ?? [] : undefined,
      // SPEC-076: the catalogue a banded table resolves its style through,
      // and the row window this page slice draws.
      styles: flatEntry.isTable ? model.styles ?? [] : undefined,
      rowStart: flatEntry.isTable ? tableRowStart : undefined,
      rowEnd: flatEntry.isTable ? tableRowEnd : undefined,
      dataTableSection: flatEntry.isTable ? flatEntry.section : undefined,
      dataTableBlock: flatEntry.isTable ? flatEntry.blockIndex : undefined,
      // SPEC-047: resolve the image's data URI here — only realized pages
      // reach renderEntryFor, so a still-deferred part is force-encoded (via
      // mediaSrc -> mediaDataUri) just for the images actually painting, not
      // the whole document's media up front in flatBlocks (which would put
      // the entire ~1s encode back on the open path).
      imageSrc: flatEntry.isImage ? mediaSrc(model, block) : undefined,
      imageSelected: flatEntry.isImage ? flatEntry.imageSelected : undefined,
      shapeSelected: flatEntry.isShape ? flatEntry.shapeSelected : undefined,
      shapeEditing: flatEntry.isShape ? flatEntry.shapeEditing : undefined,
      chartSelected: flatEntry.isChart ? flatEntry.chartSelected : undefined,
      drawingIndex: flatEntry.isShape || flatEntry.isChart ? flatEntry.drawingIndex : undefined,
      column,
      trackChangesDisplayMode: deps.trackChangesDisplayMode.value,
    } as RenderEntry;
  }
  if (flatEntry.isConnector) {
    // SPEC-070: a connector renders as a zero-footprint anchor (like a
    // floating shape's) — the visible SVG path is painted by the page-level
    // connector overlay (useConnectors), which measures this anchor's origin.
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, column],
      kind: "connector",
      // Zero-footprint anchor entry, shaped exactly like a break: no blockProp
      // (so componentFor's plain <div> never stringifies a block into an
      // attribute) and no runs. Its data-section/data-block let useConnectors
      // measure its page origin; sectionProp/blockIndexProp let
      // hoistFloatingDrawingEntries pair it with its anchor paragraph.
      tag: "div",
      className: "doc-connector-anchor",
      uneditable: true,
      dataSection: flatEntry.section,
      dataBlock: flatEntry.blockIndex,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      column,
    } as RenderEntry;
  }
  if (flatEntry.isToc) {
    const toc = asToc(block);
    const staleHeadings = tocHeadingsStale(model, toc);
    const stalePagination = tocPageNumbersStale(toc, (anchor) => deps.tocPageNumberFor(model, anchor));
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, staleHeadings, stalePagination, column],
      kind: "toc",
      uneditable: true,
      blockProp: block,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      tocStaleHeadings: staleHeadings,
      tocStalePagination: stalePagination,
      column,
    } as RenderEntry;
  }
  if (flatEntry.isTof) {
    const tof = asTof(block);
    const staleEntries = tofEntriesStale(model, tof);
    const stalePagination = tofPageNumbersStale(tof, (anchor) => deps.tocPageNumberFor(model, anchor));
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, staleEntries, stalePagination, column],
      kind: "tof",
      uneditable: true,
      blockProp: block,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      tofStaleEntries: staleEntries,
      tofStalePagination: stalePagination,
      column,
    } as RenderEntry;
  }
  const paragraph = flatEntry.paragraph;
  let sliceStart: number | undefined;
  let sliceEnd: number | undefined;
  if (item.startLine !== undefined) {
    const widthPx = sectionColumnFlowWidthPx(model, item.sectionIndex);
    const offsets = deps.pagination.isMeasured(block, widthPx)
      ? deps.pagination.lineOffsetsFor(block)
      : null;
    if (!offsets) {
      if (item.startLine > 0) {
        return null;
      }
    } else {
      sliceStart = offsets[item.startLine] ?? 0;
      sliceEnd =
        item.endLine !== undefined && item.endLine < offsets.length
          ? offsets[item.endLine]
          : undefined;
      if (sliceStart === 0 && sliceEnd === undefined) {
        sliceStart = undefined;
      }
    }
  }
  let runs = paragraph.runs;
  if (sliceStart !== undefined || sliceEnd !== undefined) {
    const from = sliceStart ?? 0;
    const tail = splitRuns(paragraph.runs, from)[1];
    runs = sliceEnd !== undefined ? splitRuns(tail, sliceEnd - from)[0] : tail;
  }
  const continuation = sliceStart !== undefined && sliceStart > 0;
  // Paragraph borders and shading (SPEC-074): which edges THIS fragment
  // paints depends on the border run it belongs to (AC-12) and on where the
  // fragment sits inside a page-split paragraph (AC-14..16).
  const styleObj = {
    ...paragraphStyle(
      paragraph,
      documentThemeOf(model),
      model.listDefinitions ?? [],
      deps.trackChangesDisplayMode.value,
      fragmentEdges(
        sectionBorderRuns(model, model.sections[flatEntry.section].blocks, inheritedPaintOf),
        flatEntry.blockIndex,
        !continuation,
        sliceEnd === undefined,
      ),
      // What the paragraph's style contributes (AC-17); its own wins.
      inheritedPaintOf(paragraph),
    ),
  };
  if (continuation) {
    // A continuation slice starts mid-paragraph: no space-before, no
    // first-line indent, no repeated list marker.
    styleObj.marginTop = "0";
    styleObj.textIndent = "0";
  }
  if (sliceEnd !== undefined) {
    styleObj.marginBottom = "0";
  }
  const sliceSig = `${sliceStart ?? ""}:${sliceEnd ?? ""}`;
  // A body heading that owns content shows a collapse chevron (SPEC-037
  // AC-16) — only on its first slice, never on a continuation.
  const headingKey = blockKey(flatEntry.section, flatEntry.blockIndex);
  const outlineToggle =
    paragraph.headingLevel !== undefined &&
    !continuation &&
    headingContentLength(model, flatEntry.section, flatEntry.blockIndex) > 0;
  const outlineCollapsed = outlineToggle && deps.collapsedHeadings.value.has(headingKey);
  // SPEC-053 field shading: the caret's offset within this body paragraph, so
  // the field run it lands on paints shaded (AC-10). Run start offsets let a
  // field branch below know its own atom's offset. `fieldCaret` is the
  // caret mirror narrowed by usePaginationView to "inside a field-bearing
  // body paragraph" — never the raw reactive caret, which would make every
  // caret move rebuild every page (see the computed's comment).
  const caretOffsetHere =
    fieldCaret && fieldCaret.section === flatEntry.section && fieldCaret.block === flatEntry.blockIndex
      ? fieldCaret.offset
      : null;
  const runStarts: number[] = [];
  {
    let acc = 0;
    for (const r of runs) {
      runStarts.push(acc);
      acc += r.text.length;
    }
  }
  // SPEC-054 tab layout: resolve the paragraph's tab advances and bar-rule
  // positions once, keyed by full-paragraph character index. Only paragraphs
  // that actually carry a tab or their own stops pay for the resolution; a
  // slice's run offsets are shifted by its start below.
  const ownStops = paragraph.formatting?.tabStops;
  const paragraphHasTab = paragraph.runs.some((run) => run.text.includes("\t"));
  const tabLayout =
    paragraphHasTab || (ownStops !== undefined && ownStops.length > 0)
      ? deps.pagination.tabLayoutFor(model, paragraph)
      : null;
  const sliceBase = sliceStart ?? 0;
  const fieldCtx: FieldContext = {
    model,
    now: now(),
    fileName: model.title ?? "",
    pageOfBookmark: (name) => deps.tocPageNumberFor(model, name),
  };
  let noteSig = "";
  const renderRuns = runs.map((run, runIndex) => {
    // An inline/display equation (SPEC-056): the laid-out equation is
    // substituted for the one-character placeholder at paint time, as a caret
    // atom (data-caret-atom) so one Backspace removes the whole equation and
    // caret arithmetic counts it as a single character.
    if (run.equation) {
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: { ...runStyle(run.formatting, documentThemeOf(model)), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
        text: "",
        segments: null,
        atomText: run.text,
        equation: run.equation,
        lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
      };
    }
    // A mail-merge chip (SPEC-046): renders its canonical `{{...}}` label
    // as a caret atom standing for the one-character model placeholder —
    // the same atomText/data-caret-atom mechanism as a note reference mark
    // just below, so one Backspace removes the whole chip (AC-14).
    if (run.field && isMergeChipField(run.field)) {
      const label = chipLabel(run.field);
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: { ...runStyle(run.formatting, documentThemeOf(model)), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
        text: label,
        segments: null,
        atomText: run.text,
        atomClass: "doc-merge-chip",
        lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
      };
    }
    // A note reference mark (SPEC-031): renders its computed number as a
    // caret atom standing for the one-character model placeholder, with the
    // note's text as its tooltip (AC-14).
    if (run.note) {
      const label = deps.noteLabelsMap.value.get(run.note.noteId) ?? run.text;
      const title = deps.noteTooltipsMap.value.get(run.note.noteId) ?? "";
      noteSig += `${run.note.noteId}=${label}|${title};`;
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: { ...runStyle(run.formatting, documentThemeOf(model)), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
        text: label,
        segments: null,
        atomText: run.text,
        noteId: run.note.noteId,
        title,
        lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
      };
    }
    // A SPEC-053 computed field (date, author, cross-reference, page/count in
    // the body, …): its value substituted at paint time, shaded when field
    // shading is on or the caret sits on it (AC-6, AC-10, AC-11).
    if (run.field) {
      const start = runStarts[runIndex];
      const caretOnField =
        caretOffsetHere !== null && caretOffsetHere >= start && caretOffsetHere <= start + run.text.length;
      return fieldRenderRun(
        run,
        runIndex,
        { ...fieldCtx, here: { section: flatEntry.section, block: flatEntry.blockIndex, offset: start } },
        { ...runStyle(run.formatting, documentThemeOf(model)), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
        effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
        caretOnField,
        deps.fieldShadingOn.value,
      );
    }
    // SPEC-039 AC-9: "No markup"/"Original" render as if every change were
    // already accepted/rejected — a hidden run (a pending deletion in "No
    // markup", a pending insertion in "Original") shows no text at all, but
    // stays a caret atom carrying its full model text (see
    // trackChangeRunHidden/DocBlockEntries.vue's caretAtom branch).
    const hidden = trackChangeRunHidden(run, deps.trackChangesDisplayMode.value);
    const displayFormatting = displayRunFormatting(run, deps.trackChangesDisplayMode.value);
    const href = hidden ? undefined : runLinkHref(displayFormatting);
    const external = href ? isExternalLink(displayFormatting) : false;
    // SPEC-054: a run holding a tab is split at its tabs so each renders in a
    // sized `.doc-tab` box (marks or no marks); with marks on but no tab, the
    // familiar space/pilcrow marking still applies.
    let segments: MarkSegment[] | null = null;
    if (!hidden && tabLayout && run.text.includes("\t")) {
      segments = annotateTabSegments(
        marks ? markSegments(run.text) : tabOnlySegments(run.text),
        sliceBase + runStarts[runIndex],
        tabLayout,
      );
    } else if (!hidden && marks) {
      segments = markSegments(run.text);
    }
    const marked = segments && segments.some((segment) => segment.kind !== "text");
    return {
      key: runKey(run, runIndex),
      href: href || "",
      target: external ? "_blank" : undefined,
      rel: external ? "noopener noreferrer" : undefined,
      styleObj: { ...runStyle(displayFormatting, documentThemeOf(model)), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
      text: hidden ? "" : run.text,
      segments: marked ? segments : null,
      caretAtom: hidden ? run.text : undefined,
      lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
    };
  });
  // Field shading (AC-10, AC-11) is caret- and toggle-driven state, not part
  // of the model — the memo below must carry it, or v-memo skips repainting a
  // field's shading when only the caret moved or the View toggle flipped.
  // Only paragraphs that actually hold a field vary this signal, so plain text
  // keeps its cheap memo.
  const fieldShadeSig = runs.some((run) => run.field && !isMergeChipField(run.field))
    ? `${deps.fieldShadingOn.value ? 1 : 0}:${caretOffsetHere ?? ""}`
    : "";
  // SPEC-054: a signature of the resolved tab geometry so v-memo repaints when
  // an advance width or bar position changes without the block object itself
  // changing (a default-interval or font-load reflow). Empty for the tab-free
  // majority, keeping their memo cheap.
  const tabSig = tabLayout
    ? `${tabLayout.bars.map((bar) => Math.round(bar)).join(",")}|${[...tabLayout.advances.entries()]
        .map(([index, advance]) => `${index}:${Math.round(advance.widthPx)}`)
        .join(",")}`
    : "";
  if (tabLayout && tabLayout.bars.length > 0) {
    // The bar rules are absolutely positioned inside the paragraph box.
    styleObj.position = "relative";
  }
  const borderSig = BORDER_MEMO_KEYS.map((key) => styleObj[key] ?? "").join("|");
  return {
    key: continuation ? `${flatEntry.key}@${sliceStart}` : flatEntry.key,
    memo: [
      block,
      flatEntry.section,
      flatEntry.blockIndex,
      flatEntry.marker,
      sliceSig,
      marks,
      column,
      noteSig,
      outlineToggle,
      outlineCollapsed,
      deps.trackChangesDisplayMode.value,
      fieldShadeSig,
      tabSig,
      // The painted edges depend on the NEIGHBOURS' borders too (SPEC-074
      // AC-12), so a paragraph must repaint when the run it belongs to
      // changes shape even though its own block object did not.
      borderSig,
      themeSig,
    ],
    kind: "paragraph",
    tag: headingTag(paragraph.headingLevel),
    className: paragraphClassName(paragraph),
    styleObj,
    dataSection: flatEntry.section,
    dataBlock: flatEntry.blockIndex,
    sliceStart: continuation ? String(sliceStart) : undefined,
    sliceMore: sliceEnd !== undefined ? "1" : undefined,
    listMarker: continuation ? undefined : flatEntry.marker || undefined,
    listLevel: paragraph.list ? paragraph.list.level : undefined,
    runs: renderRuns,
    barRules: tabLayout && tabLayout.bars.length > 0 ? tabLayout.bars : undefined,
    needsBr: runs.length === 0,
    outlineToggle: outlineToggle || undefined,
    outlineCollapsed: outlineCollapsed || undefined,
    column,
    hasTrackedChange:
      (deps.trackChangesDisplayMode.value === "all" || deps.trackChangesDisplayMode.value === "simple") &&
      paragraphHasTrackedChange(paragraph),
  } as RenderEntry;
}

/** A header/footer's content rendered as DocPage entries (SPEC-029 AC-3..6):
 * the same shapes renderEntryFor builds, minus line-slicing, with any
 * page-number/page-count field run substituted for its live value. */
function headerFooterRenderEntries(
  blocks: Block[],
  pageNumber: number,
  totalPages: number,
  marks: boolean,
  deps: PaginationViewDeps,
): RenderEntry[] {
  const model = deps.doc.value;
  const entries = noRenderEntries();
  blocks.forEach((block, blockIndex) => {
    if (isTable(block)) {
      entries.push({
        key: `hf:${blockIndex}`,
        memo: [block, blockIndex, marks, documentThemeOf(model)],
        kind: "table",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: blockIndex,
      } as RenderEntry);
      return;
    }
    if (isImage(block)) {
      entries.push({
        key: `hf:${blockIndex}`,
        memo: [block, blockIndex],
        kind: "image",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: blockIndex,
        imageSrc: mediaSrc(model, block),
        imageSelected: false,
      } as RenderEntry);
      return;
    }
    const paragraph = asParagraphBlock(block);
    const styleObj = { ...paragraphStyle(paragraph, documentThemeOf(model)) };
    const renderRuns = paragraph.runs.map((run, runIndex) => {
      if (run.field && isMergeChipField(run.field)) {
        // A merge chip in a header/footer (SPEC-046): resolution happens on
        // the model before the writers, so a header still holding a chip
        // here just means an unresolved template — show the canonical
        // label, same as the body.
        const label = chipLabel(run.field);
        return {
          key: runKey(run, runIndex),
          href: "",
          styleObj: runStyle(run.formatting, documentThemeOf(model)),
          text: label,
          segments: null,
          atomText: run.text,
          atomClass: "doc-merge-chip",
          lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
        };
      }
      if (run.field && model) {
        // A SPEC-053 field in a header/footer — page/count show the live page
        // (AC-35: correct on every page); other kinds show their evaluated
        // value. Shaded only via the View toggle (the caret is never in a
        // header/footer paint pass).
        return fieldRenderRun(
          run,
          runIndex,
          { model, now: now(), fileName: model.title ?? "", pageNumber, totalPages, pageOfBookmark: (name) => deps.tocPageNumberFor(model, name) },
          runStyle(run.formatting, documentThemeOf(model)),
          effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
          false,
          deps.fieldShadingOn.value,
        );
      }
      const text = run.text;
      const segments = marks ? markSegments(text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting, documentThemeOf(model)),
        text,
        segments: marked ? segments : null,
        lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
      };
    });
    // A header/footer field shades only via the View toggle (AC-11); track it
    // in the memo so flipping the toggle repaints, when the band holds a field.
    const hfFieldShadeSig = paragraph.runs.some((run) => run.field && !isMergeChipField(run.field))
      ? `${deps.fieldShadingOn.value ? 1 : 0}`
      : "";
    entries.push({
      key: `hf:${blockIndex}`,
      memo: [block, blockIndex, pageNumber, totalPages, marks, deps.hfRemountEpoch.value, hfFieldShadeSig, documentThemeOf(model)],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: paragraphClassName(paragraph),
      styleObj,
      dataSection: 0,
      dataBlock: blockIndex,
      runs: renderRuns,
      needsBr: renderRuns.length === 0,
    } as RenderEntry);
  });
  return entries;
}

/** One note's blocks (or page slices of them) as DocPage entries (SPEC-031). */
function noteBlocksEntries(
  note: NoteContent,
  items: NoteAreaItem[],
  pageNumber: number,
  totalPages: number,
  marks: boolean,
  deps: PaginationViewDeps,
): RenderEntry[] {
  const entries = noRenderEntries();
  if (!note) {
    return entries;
  }
  const model = deps.doc.value;
  for (const item of items) {
    const block = note.blocks[item.blockIndex];
    if (!block) {
      continue;
    }
    if (isTable(block)) {
      entries.push({
        key: `note:${note.noteId}:${item.blockIndex}`,
        memo: [block, item.blockIndex, marks],
        kind: "table",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: item.blockIndex,
      } as RenderEntry);
      continue;
    }
    if (isImage(block)) {
      entries.push({
        key: `note:${note.noteId}:${item.blockIndex}`,
        memo: [block, item.blockIndex],
        kind: "image",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: item.blockIndex,
        imageSrc: mediaSrc(model, block),
        imageSelected: false,
      } as RenderEntry);
      continue;
    }
    const paragraph = asParagraphBlock(block);
    let sliceStart: number | undefined;
    let sliceEnd: number | undefined;
    if (item.startLine !== undefined) {
      const offsets = deps.pagination.lineOffsetsFor(block);
      if (!offsets) {
        // Unmeasured split: the first slice renders whole, continuations wait.
        if (item.startLine > 0) {
          continue;
        }
      } else {
        sliceStart = offsets[item.startLine] ?? 0;
        sliceEnd =
          item.endLine !== undefined && item.endLine < offsets.length
            ? offsets[item.endLine]
            : undefined;
        if (sliceStart === 0 && sliceEnd === undefined) {
          sliceStart = undefined;
        }
      }
    }
    let runs = paragraph.runs;
    if (sliceStart !== undefined || sliceEnd !== undefined) {
      const from = sliceStart ?? 0;
      const tail = splitRuns(paragraph.runs, from)[1];
      runs = sliceEnd !== undefined ? splitRuns(tail, sliceEnd - from)[0] : tail;
    }
    const continuation = sliceStart !== undefined && sliceStart > 0;
    const renderRuns = runs.map((run, runIndex) => {
      if (run.field && isMergeChipField(run.field)) {
        // A merge chip in a note (SPEC-046) — resolution happens on the model
        // before the writers, so one still here means an unresolved template;
        // show the canonical label, same as the body and header/footer.
        return {
          key: runKey(run, runIndex),
          href: "",
          styleObj: runStyle(run.formatting, documentThemeOf(model)),
          text: chipLabel(run.field),
          segments: null,
          atomText: run.text,
          atomClass: "doc-merge-chip",
          lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
        };
      }
      if (run.field && model) {
        // A SPEC-053 field in a footnote/endnote (AC-33 lists notes among the
        // surfaces fields work in): its value substituted at paint time, the
        // same branch headerFooterRenderEntries takes. Without this a note
        // field painted its raw one-character placeholder, so a note field
        // could never show a value — nor show F9 having recomputed it.
        // Shaded only via the View toggle: the note caret lives in the note's
        // own synthetic coordinates, not the body caret this pass can read.
        return fieldRenderRun(
          run,
          runIndex,
          { model, now: now(), fileName: model.title ?? "", pageNumber, totalPages, pageOfBookmark: (name) => deps.tocPageNumberFor(model, name) },
          runStyle(run.formatting, documentThemeOf(model)),
          effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
          false,
          deps.fieldShadingOn.value,
        );
      }
      const segments = marks ? markSegments(run.text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting, documentThemeOf(model)),
        text: run.text,
        segments: marked ? segments : null,
        lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
      };
    });
    entries.push({
      key: `note:${note.noteId}:${item.blockIndex}${continuation ? `@${sliceStart}` : ""}`,
      // pageNumber/totalPages join the memo for the same reason the header's
      // entry carries them: a `page` field in a note has to repaint when
      // repagination moves the note to another page.
      memo: [
        block,
        item.blockIndex,
        `${sliceStart ?? ""}:${sliceEnd ?? ""}`,
        pageNumber,
        totalPages,
        marks,
        deps.fieldShadingOn.value,
        deps.noteRemountEpoch.value,
      ],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: paragraphClassName(paragraph),
      styleObj: { ...paragraphStyle(paragraph, documentThemeOf(model)) },
      dataSection: 0,
      dataBlock: item.blockIndex,
      sliceStart: continuation ? String(sliceStart) : undefined,
      sliceMore: sliceEnd !== undefined ? "1" : undefined,
      runs: renderRuns,
      needsBr: renderRuns.length === 0,
    } as RenderEntry);
  }
  return entries;
}

/** One page's footnote or endnote area as DocPage's prop shape (SPEC-031). */
function noteAreaView(
  model: DocumentModel,
  area: NoteArea | null,
  pageNumber: number,
  totalPages: number,
  marks: boolean,
  deps: PaginationViewDeps,
): NoteAreaView | null {
  if (!area) {
    return null;
  }
  const activeId = deps.activeNote.value?.noteId ?? null;
  const labels = deps.noteLabelsMap.value;
  const groups: Array<{ noteId: string; items: NoteAreaItem[] }> = [];
  let current: { noteId: string; items: NoteAreaItem[] } | null = null;
  for (const item of area.items) {
    if (!current || current.noteId !== item.noteId) {
      current = { noteId: item.noteId, items: noNoteAreaItems() };
      groups.push(current);
    }
    current.items.push(item);
  }
  const notes: NoteAreaView["notes"] = [];
  for (const group of groups) {
    const note = noteById(model, group.noteId);
    if (!note) {
      continue;
    }
    const first = group.items[0];
    const starts = first.blockIndex === 0 && (first.startLine === undefined || first.startLine === 0);
    const isActive = deps.editable.value && activeId === group.noteId;
    if (isActive && !starts) {
      continue; // The active note renders whole where it starts.
    }
    const items: NoteAreaItem[] = isActive
      ? note.blocks.map((_, blockIndex) => ({ noteId: note.noteId, blockIndex }))
      : group.items;
    notes.push({
      noteId: note.noteId,
      kind: note.kind,
      label: starts ? (labels.get(note.noteId) ?? "") : "",
      editable: isActive,
      entries: noteBlocksEntries(note, items, pageNumber, totalPages, marks, deps),
    });
  }
  return notes.length > 0 ? { kind: area.kind, notes } : null;
}

/** The footnote area's band geometry (SPEC-031 AC-2). */
function footnoteAreaStyle(
  margins: { topCm: number; rightCm: number; bottomCm: number; leftCm: number },
  heightPx: number,
): Record<string, string> {
  return {
    left: `${margins.leftCm}cm`,
    right: `${margins.rightCm}cm`,
    bottom: `${margins.bottomCm}cm`,
    height: `${heightPx}px`,
  };
}

/** The header/footer band geometry for a page (SPEC-029). */
function headerFooterBandStyle(
  margins: { topCm: number; rightCm: number; bottomCm: number; leftCm: number },
  part: string,
): Record<string, string> {
  const left = `${margins.leftCm}cm`;
  const right = `${margins.rightCm}cm`;
  if (part === "header") {
    return { left, right, height: `${margins.topCm}cm`, top: "0" };
  }
  return { left, right, height: `${margins.bottomCm}cm`, bottom: "0" };
}

export function usePaginationView(deps: PaginationViewDeps): { pagedViews: ComputedRef<PageView[]> } {
  /** Sections flattened to one list so a single v-for can carry v-memo. */
  const flatBlocks = computed<FlatEntry[]>(() => {
    if (!deps.doc.value) {
      return [];
    }
    const bump = deps.remount.value;
    const markers = listMarkers(deps.doc.value);
    const range = deps.cellSelection.value;
    const selectedImage = deps.imageSelection.value;
    const selectedShape = deps.shapeSelection.value;
    const multiSelectedShapes = deps.shapeMultiSelection.value;
    const shapeTextEditing = deps.editingShapeText.value;
    const selectedChart = deps.chartSelection.value;
    let drawingCounter = 0;
    return deps.doc.value.sections.flatMap((section, sectionIndex) =>
      section.blocks.map((block, blockIndex) => {
        const address = `${sectionIndex}:${blockIndex}`;
        const suffix = bump.key === address ? `:${bump.count}` : "";
        const table = isTable(block);
        const image = isImage(block);
        const shape = isShape(block);
        const chart = isChart(block);
        const connector = isConnector(block);
        const toc = isToc(block);
        const tof = isTof(block);
        const cellSel =
          table && range && range.section === sectionIndex && range.block === blockIndex
            ? `${range.anchor.row}:${range.anchor.col}:${range.focus.row}:${range.focus.col}`
            : "";
        const imgSelected =
          image && !!selectedImage && selectedImage.section === sectionIndex && selectedImage.block === blockIndex;
        const shapeIsPrimarySelection =
          !!selectedShape && selectedShape.section === sectionIndex && selectedShape.block === blockIndex;
        const shapeSelected =
          shape &&
          (shapeIsPrimarySelection ||
            multiSelectedShapes.some((loc) => loc.section === sectionIndex && loc.block === blockIndex));
        const chartSelected =
          chart && !!selectedChart && selectedChart.section === sectionIndex && selectedChart.block === blockIndex;
        return {
          key: `${bump.epoch}:${address}${suffix}`,
          block,
          paragraph: asParagraphBlock(block),
          isTable: table,
          isImage: image,
          isShape: shape,
          isChart: chart,
          isConnector: connector,
          isToc: toc,
          isTof: tof,
          imageSelected: imgSelected,
          shapeSelected,
          shapeEditing: shape && shapeIsPrimarySelection && shapeTextEditing,
          chartSelected,
          drawingIndex: shape || chart ? drawingCounter++ : -1,
          section: sectionIndex,
          blockIndex,
          marker: markers[sectionIndex][blockIndex],
          cellSel,
          imgSel: imgSelected ? "1" : "",
        };
      }),
    );
  });

  /** SPEC-053 AC-10 field shading is the page render's only caret-dependent
   * input, and it must not track the raw caret mirror: every keystroke and
   * click replaces `caretPosition`, and a tracked read of it inside
   * renderEntryFor rebuilt every realized page's entries and re-rendered every
   * DocPage for a shading state that cannot have changed — a second full page
   * pass inside the open's first-layout task (SPEC-047 AC-16's longest block)
   * and extra full-page flushes on every keystroke. This computed is the caret
   * only while it sits in a body paragraph that holds a shadeable field run,
   * and null otherwise; a computed that keeps returning null never notifies
   * pagedViews, so plain-text caret moves cost nothing here. Cell carets stay
   * null, as before: table-cell fields are DocTableBlock's own paint. */
  const fieldCaret = computed<CaretPosition | null>(() => {
    const caret = deps.caretPosition.value;
    const model = deps.doc.value;
    if (!caret || !model || caret.cell) {
      return null;
    }
    const block = model.sections[caret.section]?.blocks[caret.block];
    if (!block || block.type !== "paragraph") {
      return null;
    }
    return block.runs.some((run) => run.field && !isMergeChipField(run.field)) ? caret : null;
  });

  /** The rendered pages: the layout's items joined with flatBlocks' per-block
   * render data, styles resolved from each page's own geometry. Under
   * virtualization only realized pages build entries; the rest render as
   * empty boxes with the layout's exact geometry. */
  const pagedViews = computed<PageView[]>(() => {
    const model = deps.doc.value;
    if (!model) {
      return [];
    }
    const fieldCaretNow = fieldCaret.value;
    const marks = deps.marksOn.value;
    // Only a mixed-size document names its page boxes (see pageBoxStyle).
    const namedPages = hasMixedPageSizes(model);
    const byAddress = new Map<string, FlatEntry>();
    for (const entry of flatBlocks.value) {
      byAddress.set(`${entry.section}:${entry.blockIndex}`, entry);
    }
    const lastSection = model.sections.length - 1;
    const virtual = deps.virtualizing.value;
    const range = deps.viewportPageRange.value;
    const caretIndex = deps.caretPage.value - 1;
    const anchorIndex = deps.anchorPage.value;
    const findIndex = deps.findMatchPage.value;
    const active = deps.activeHeaderFooter.value;
    const numbers = deps.pageNumbers.value;
    const totalPages = deps.pagination.pages.value.length;
    const pageNoteAreas = deps.pagination.noteAreas.value;
    // Borders and shading a paragraph inherits from its style (SPEC-074
    // AC-17). One resolver per pass, memoised per style id inside.
    const inheritedPaintOf = paragraphPaintResolver(model);
    // Each page's ordinal within its own section — what a page border's
    // "first page"/"all but the first" restriction is measured against
    // (SPEC-074 AC-22).
    const pageIndexInSection = ordinalsWithinSection(deps.pagination.pages.value);
    return deps.pagination.pages.value.map((page, pageIndex) => {
      const realized =
        !virtual ||
        (pageIndex >= range.start && pageIndex <= range.end) ||
        pageIndex === caretIndex ||
        pageIndex === anchorIndex ||
        pageIndex === findIndex ||
        deps.realizedSticky.has(page);
      if (!realized) {
        return {
          pageStyle: pageBoxStyle(page, namedPages),
          contentStyle: {
            padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
          },
          pageBorderStyle: pageBorderStyleFor(model, page, pageIndexInSection[pageIndex]),
          pageBackgroundStyle: pageBackgroundStyleFor(model, page),
          entries: [],
          columns: page.columns,
          virtual: true,
          dimmed: false,
          headerEditable: false,
          footerEditable: false,
          headerStyle: headerFooterBandStyle(page.margins, "header"),
          footerStyle: headerFooterBandStyle(page.margins, "footer"),
          headerEntries: [],
          footerEntries: [],
          footnoteArea: null,
          endnoteArea: null,
        };
      }
      deps.realizedSticky.add(page);
      const entries: RenderEntry[] = [];
      for (const item of page.items) {
        const flatEntry = byAddress.get(`${item.sectionIndex}:${item.blockIndex}`);
        if (!flatEntry) {
          continue;
        }
        const entry = renderEntryFor(model, flatEntry, item, marks, deps, fieldCaretNow, inheritedPaintOf);
        if (!entry) {
          continue;
        }
        entries.push(entry);
        const sectionBlocks = model.sections[item.sectionIndex].blocks;
        if (
          marks &&
          item.sectionIndex < lastSection &&
          item.blockIndex === sectionBlocks.length - 1 &&
          entry.sliceMore === undefined
        ) {
          const kind =
            model.sections[item.sectionIndex + 1].start === "continuous" ? "continuous" : "nextPage";
          entries.push({
            key: `${flatEntry.key}:marker`,
            memo: [kind, marks, flatEntry.key],
            kind: "marker",
            tag: "div",
            className: "doc-section-marker",
            uneditable: true,
            sectionMarker: kind,
            markerLabel: SECTION_MARKER_LABELS[kind],
            column: entry.column,
          } as RenderEntry);
        }
      }
      // Header/footer bands (SPEC-029 AC-9, AC-10): every page of a section
      // renders the same resolved content — the one page `active` addresses
      // renders it as the live editable surface, every other page read-only.
      const pageNumber = numbers[pageIndex] ?? pageIndex + 1;
      const headerEditable = !!active && active.part === "header" && active.pageIndex === pageIndex;
      const footerEditable = !!active && active.part === "footer" && active.pageIndex === pageIndex;
      const headerContent = headerFooterForPage(model, deps.pagination.pages.value, pageIndex, pageNumber, "header");
      const footerContent = headerFooterForPage(model, deps.pagination.pages.value, pageIndex, pageNumber, "footer");
      const headerBlocks = headerContent?.blocks ?? (headerEditable ? createEmptyHeaderFooterContent().blocks : null);
      const footerBlocks = footerContent?.blocks ?? (footerEditable ? createEmptyHeaderFooterContent().blocks : null);
      // Note areas (SPEC-031): the footnote band pins to the content bottom;
      // the endnote area flows after the body content (AC-2, AC-4).
      const footnotePlan = pageNoteAreas.footnotes[pageIndex] ?? null;
      const footnoteView = footnotePlan ? noteAreaView(model, footnotePlan, pageNumber, totalPages, marks, deps) : null;
      const endnoteView = noteAreaView(model, pageNoteAreas.endnotes[pageIndex] ?? null, pageNumber, totalPages, marks, deps);
      return {
        pageStyle: pageBoxStyle(page, namedPages),
        contentStyle: {
          padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
        },
        pageBorderStyle: pageBorderStyleFor(model, page, pageIndexInSection[pageIndex]),
        pageBackgroundStyle: pageBackgroundStyleFor(model, page),
        // Floating drawings render just before their anchor paragraph —
        // the paragraph-top reference their offsets are measured from
        // (SPEC-078, hoistFloatingDrawingEntries).
        entries: hoistFloatingDrawingEntries(entries, model),
        columns: page.columns,
        virtual: false,
        dimmed: !!active,
        headerEditable,
        footerEditable,
        headerStyle: headerFooterBandStyle(page.margins, "header"),
        footerStyle: headerFooterBandStyle(page.margins, "footer"),
        headerEntries: headerBlocks ? headerFooterRenderEntries(headerBlocks, pageNumber, totalPages, marks, deps) : [],
        footerEntries: footerBlocks ? headerFooterRenderEntries(footerBlocks, pageNumber, totalPages, marks, deps) : [],
        footnoteArea:
          footnotePlan && footnoteView
            ? { ...footnoteView, style: footnoteAreaStyle(page.margins, footnotePlan.heightPx) }
            : null,
        endnoteArea: endnoteView,
      };
    });
  });

  return { pagedViews };
}
