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
  createEmptyHeaderFooterContent,
  isBreak,
  isImage,
  isTable,
  isToc,
  type Block,
  type DocumentModel,
  type NoteContent,
  type Paragraph,
  type TrackChangesDisplayMode,
} from "../../lib/documentModel";
import {
  sectionColumnFlowWidthPx,
  type PageColumns,
  type PageItem,
} from "../../lib/pagination";
import { listMarkers } from "../../lib/lists";
import { noNoteAreaItems, type NoteArea, type NoteAreaItem } from "../../lib/notesLayout";
import { noRenderEntries, type RenderEntry } from "../../lib/pageRender";
import { noteById } from "../../lib/notes";
import { fieldValueText, headerFooterForPage } from "../../lib/pageNumbering";
import { headingTag, isExternalLink, paragraphStyle, runKey, runLinkHref, runStyle, trackChangeRunStyle } from "../../lib/renderFormatting";
import { effectiveLanguage } from "../../lib/language";
import { markSegments } from "../../lib/marks";
import { splitRuns } from "../../lib/textEditing";
import { tocHeadingsStale, tocPageNumbersStale } from "../../lib/toc";
import { blockKey, headingContentLength } from "../../lib/outline";
import { mediaDataUri } from "../../lib/images";
import type { PaginationController } from "../../composables/usePagination";
import type { CellRange } from "../../lib/tables";
import type { ImageLocation } from "../../lib/images";

/** One flattened block row — the section-flattened list a single v-for carries
 * (the v-memo keys that keep typing responsive on very large documents). */
export interface FlatEntry {
  key: string;
  block: Block;
  paragraph: Paragraph;
  isTable: boolean;
  isImage: boolean;
  isToc: boolean;
  imageSrc: string;
  imageSelected: boolean;
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

export interface PaginationViewDeps {
  doc: Ref<DocumentModel | null>;
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
  activeHeaderFooter: Ref<any>;
  activeNote: Ref<any>;
  editable: Ref<boolean>;
  cellSelection: Ref<CellRange | null>;
  imageSelection: Ref<ImageLocation | null>;
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

/** One page item's precomputed render entry for DocPage — or null for the
 * continuation slice of a block whose real metrics have not arrived yet. */
function renderEntryFor(
  model: DocumentModel,
  flatEntry: FlatEntry,
  item: PageItem,
  marks: boolean,
  deps: PaginationViewDeps,
): RenderEntry | null {
  const block = flatEntry.block;
  const column = item.columnIndex;
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
  if (flatEntry.isTable || flatEntry.isImage) {
    return {
      key: flatEntry.key,
      memo: [
        block,
        flatEntry.section,
        flatEntry.blockIndex,
        flatEntry.marker,
        flatEntry.cellSel,
        flatEntry.imgSel,
        column,
        model.listDefinitions,
        deps.trackChangesDisplayMode.value,
      ],
      kind: flatEntry.isTable ? "table" : "image",
      blockProp: block,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      listDefinitions: flatEntry.isTable ? model.listDefinitions ?? [] : undefined,
      imageSrc: flatEntry.isImage ? flatEntry.imageSrc : undefined,
      imageSelected: flatEntry.isImage ? flatEntry.imageSelected : undefined,
      column,
      trackChangesDisplayMode: flatEntry.isTable ? deps.trackChangesDisplayMode.value : undefined,
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
  const styleObj = { ...paragraphStyle(paragraph, model.listDefinitions ?? []) };
  const continuation = sliceStart !== undefined && sliceStart > 0;
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
  let noteSig = "";
  const renderRuns = runs.map((run, runIndex) => {
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
        styleObj: { ...runStyle(run.formatting), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
        text: label,
        segments: null,
        atomText: run.text,
        noteId: run.note.noteId,
        title,
        lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
      };
    }
    const href = runLinkHref(run.formatting);
    const external = href ? isExternalLink(run.formatting) : false;
    const segments = marks ? markSegments(run.text) : null;
    const marked = segments && segments.some((segment) => segment.kind !== "text");
    return {
      key: runKey(run, runIndex),
      href: href || "",
      target: external ? "_blank" : undefined,
      rel: external ? "noopener noreferrer" : undefined,
      styleObj: { ...runStyle(run.formatting), ...trackChangeRunStyle(run, deps.trackChangesDisplayMode.value) },
      text: run.text,
      segments: marked ? segments : null,
      lang: effectiveLanguage(run.formatting.lang, model.defaultStyles.language),
    };
  });
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
    ],
    kind: "paragraph",
    tag: headingTag(paragraph.headingLevel),
    className: "doc-paragraph",
    styleObj,
    dataSection: flatEntry.section,
    dataBlock: flatEntry.blockIndex,
    sliceStart: continuation ? String(sliceStart) : undefined,
    sliceMore: sliceEnd !== undefined ? "1" : undefined,
    listMarker: continuation ? undefined : flatEntry.marker || undefined,
    listLevel: paragraph.list ? paragraph.list.level : undefined,
    runs: renderRuns,
    needsBr: runs.length === 0,
    outlineToggle: outlineToggle || undefined,
    outlineCollapsed: outlineCollapsed || undefined,
    column,
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
        memo: [block, blockIndex, marks],
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
    const styleObj = { ...paragraphStyle(paragraph) };
    const renderRuns = paragraph.runs.map((run, runIndex) => {
      const text = run.field
        ? fieldValueText(run.field.kind, run.field.format, pageNumber, totalPages)
        : run.text;
      const segments = marks ? markSegments(text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting),
        text,
        segments: marked ? segments : null,
        lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
      };
    });
    entries.push({
      key: `hf:${blockIndex}`,
      memo: [block, blockIndex, pageNumber, totalPages, marks, deps.hfRemountEpoch.value],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: "doc-paragraph",
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
      const segments = marks ? markSegments(run.text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting),
        text: run.text,
        segments: marked ? segments : null,
        lang: effectiveLanguage(run.formatting.lang, model?.defaultStyles.language),
      };
    });
    entries.push({
      key: `note:${note.noteId}:${item.blockIndex}${continuation ? `@${sliceStart}` : ""}`,
      memo: [block, item.blockIndex, `${sliceStart ?? ""}:${sliceEnd ?? ""}`, marks, deps.noteRemountEpoch.value],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: "doc-paragraph",
      styleObj: { ...paragraphStyle(paragraph) },
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
      entries: noteBlocksEntries(note, items, marks, deps),
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
    return deps.doc.value.sections.flatMap((section, sectionIndex) =>
      section.blocks.map((block, blockIndex) => {
        const address = `${sectionIndex}:${blockIndex}`;
        const suffix = bump.key === address ? `:${bump.count}` : "";
        const table = isTable(block);
        const image = isImage(block);
        const toc = isToc(block);
        const cellSel =
          table && range && range.section === sectionIndex && range.block === blockIndex
            ? `${range.anchor.row}:${range.anchor.col}:${range.focus.row}:${range.focus.col}`
            : "";
        const imgSelected =
          image && !!selectedImage && selectedImage.section === sectionIndex && selectedImage.block === blockIndex;
        return {
          key: `${bump.epoch}:${address}${suffix}`,
          block,
          paragraph: asParagraphBlock(block),
          isTable: table,
          isImage: image,
          isToc: toc,
          imageSrc: image ? mediaSrc(deps.doc.value, block) : "",
          imageSelected: imgSelected,
          section: sectionIndex,
          blockIndex,
          marker: markers[sectionIndex][blockIndex],
          cellSel,
          imgSel: imgSelected ? "1" : "",
        };
      }),
    );
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
    const marks = deps.marksOn.value;
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
          pageStyle: {
            width: `${page.pageSize.widthCm}cm`,
            height: `${page.pageSize.heightCm}cm`,
          },
          contentStyle: {
            padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
          },
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
        const entry = renderEntryFor(model, flatEntry, item, marks, deps);
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
      const footnoteView = footnotePlan ? noteAreaView(model, footnotePlan, marks, deps) : null;
      const endnoteView = noteAreaView(model, pageNoteAreas.endnotes[pageIndex] ?? null, marks, deps);
      return {
        pageStyle: {
          width: `${page.pageSize.widthCm}cm`,
          height: `${page.pageSize.heightCm}cm`,
        },
        contentStyle: {
          padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
        },
        entries,
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
