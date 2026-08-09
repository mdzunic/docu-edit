// Footnote/endnote placement (SPEC-031): a layout wrapper over the pure
// pagination engine that puts every footnote at the bottom of the page its
// reference renders on (AC-12, AC-25), splits over-long notes across pages
// instead of overflowing (AC-13), and lays the endnote area out after the
// last content of the document (AC-4). Framework-free like pagination.ts:
// the caller supplies measured metrics for body blocks (the engine's own
// provider) and for note blocks, so the planner is pure math — unit-testable
// with fake metrics.
//
// Footnotes and pagination are mutually dependent — reserving room for a
// note shortens its page, which can move body content (and references) to
// the next page, which moves notes. The planner runs a small fixed-point
// iteration: paginate, assign references to pages, plan the note areas,
// re-paginate with the areas' heights reserved, until the reserves settle
// (or a pass cap breaks a pathological oscillation — the final areas are
// always planned against the final pages, so notes render on the page of
// their reference even then).
import {
  sectionGeometryPx,
  sectionColumnFlowWidthPx,
  linesFitting,
  paginate,
  sliceHeight,
  type BlockMetrics,
  type MetricsProvider,
  type PageLayout,
} from "./pagination";
import type { Block, DocumentModel, NoteContent, NoteKind } from "./documentModel";
import { collectNoteRefs, noteById, type NoteRefLocation } from "./notes";

/** The gutter (px) a note's content is indented by in the note area — the
 * note number renders inside it. Note paragraphs are measured at the page
 * content width minus this, so the planner and the renderer agree. */
export const NOTE_LABEL_GUTTER_PX = 24;

/** Height (px) of the separator rule and its spacing above a note area
 * (AC-5) — included in the area's reserved height. */
export const NOTE_SEPARATOR_HEIGHT_PX = 16;

/** Vertical gap (px) between two notes in an area. */
export const NOTE_GAP_PX = 4;

/** The share of a page's content height the footnote area may grow to
 * before the rest of the note continues on the next page (AC-13). */
const MAX_FOOTNOTE_AREA_FRACTION = 0.7;

/** Fixed-point passes before accepting the current layout (see header). */
const MAX_NOTE_LAYOUT_PASSES = 4;

/** Synthetic note-only pages appended after the body — a runaway guard, far
 * above anything a real document produces per pass. */
const MAX_EXTRA_NOTE_PAGES = 200;

const EPS = 0.5;

/** One note block's share of a page's note area — mirrors PageItem's
 * whole-vs-slice convention. The note's number label renders on the item
 * that starts the note (blockIndex 0, first line). */
export interface NoteAreaItem {
  noteId: string;
  blockIndex: number;
  startLine?: number;
  endLine?: number;
}

/** One page's footnote or endnote area. `heightPx` includes the separator. */
export interface NoteArea {
  kind: NoteKind;
  heightPx: number;
  items: NoteAreaItem[];
}

/** Supplies measured metrics for a note's blocks at the note area's content
 * width — the pagination controller backs this with the same measurement
 * cache as body blocks. */
export interface NoteMetricsProvider {
  noteMetricsFor(
    note: NoteContent,
    blockIndex: number,
    block: Block,
    contentWidthPx: number,
  ): BlockMetrics;
}

/** The combined result: the final pages (including any synthetic note-only
 * pages appended at the end) plus each page's note areas by page index. */
export interface NotesLayout {
  pages: PageLayout[];
  footnoteAreas: (NoteArea | null)[];
  endnoteAreas: (NoteArea | null)[];
  /** The page each footnote's area starts on, by noteId (AC-25's check). */
  footnotePageByNote: Map<string, number>;
  /** The page each footnote reference renders on, by index into
   * collectNoteRefs(model, "footnote") — the per-page numbering restart's
   * grouping input (AC-19). */
  footnoteRefPages: number[];
}

/** The width (px) note content is measured and rendered at on a section's
 * pages: the full content width minus the number gutter. */
export function noteContentWidthPx(model: DocumentModel, sectionIndex: number): number {
  return Math.max(1, sectionGeometryPx(model, sectionIndex).contentWidthPx - NOTE_LABEL_GUTTER_PX);
}

/** Typed `null` seed for note-area parameters in `.vue` scripts, which
 * espree bars from writing the annotation (see pageRender.asPageColumns). */
export function noNoteArea(): NoteArea | null {
  return null;
}

/** Typed empty seed for note-area item lists in `.vue` scripts. */
export function noNoteAreaItems(): NoteAreaItem[] {
  return [];
}

// --- reference → page assignment --------------------------------------------

interface Placement {
  page: number;
  startChar: number;
  endChar: number;
}

// Char-windowed page placements per "section:block" — the same derivation
// the pagination controller builds for the caret page indicator, rebuilt
// here from a given layout so every fixed-point pass sees fresh pages.
function buildPlacements(
  model: DocumentModel,
  pages: PageLayout[],
  lineOffsetsOf: (block: Block) => number[] | null,
): Map<string, Placement[]> {
  const placements = new Map<string, Placement[]>();
  for (let p = 0; p < pages.length; p += 1) {
    for (const item of pages[p].items) {
      const key = `${item.sectionIndex}:${item.blockIndex}`;
      const list = placements.get(key) ?? [];
      if (item.startLine === undefined) {
        list.push({ page: p, startChar: 0, endChar: Number.POSITIVE_INFINITY });
      } else {
        const block = model.sections[item.sectionIndex]?.blocks[item.blockIndex];
        const offsets = block ? lineOffsetsOf(block) : null;
        if (!offsets) {
          // Split without known line offsets yet: it renders whole on its
          // first page (see usePagination.buildPlacements, the same rule).
          if (item.startLine === 0) {
            list.push({ page: p, startChar: 0, endChar: Number.POSITIVE_INFINITY });
          }
        } else {
          const startChar = offsets[item.startLine] ?? 0;
          const endChar =
            item.endLine !== undefined && item.endLine < offsets.length
              ? offsets[item.endLine]
              : Number.POSITIVE_INFINITY;
          list.push({ page: p, startChar, endChar });
        }
      }
      placements.set(key, list);
    }
  }
  return placements;
}

function assignRefsToPages(
  model: DocumentModel,
  pages: PageLayout[],
  refs: NoteRefLocation[],
  lineOffsetsOf: (block: Block) => number[] | null,
): number[] {
  const placements = buildPlacements(model, pages, lineOffsetsOf);
  const last = Math.max(0, pages.length - 1);
  return refs.map((ref) => {
    const list = placements.get(`${ref.section}:${ref.block}`);
    if (!list || list.length === 0) {
      return last;
    }
    for (const placement of list) {
      if (ref.offset >= placement.startChar && ref.offset < placement.endChar) {
        return placement.page;
      }
    }
    return list[list.length - 1].page;
  });
}

// --- note area filling -------------------------------------------------------

interface QueueEntry {
  note: NoteContent;
  blockIndex: number;
  startLine: number;
}

// Fills one page's note area from the front of the queue: whole blocks
// while they fit, then a line-boundary split (AC-13). Force-places at least
// one line (or one unsplittable block) into an empty area so a note taller
// than every page still makes progress. Mutates `queue` (shifting consumed
// entries) and returns the placed items and their stacked height.
function fillArea(
  queue: QueueEntry[],
  capacityPx: number,
  contentWidthPx: number,
  noteMetrics: NoteMetricsProvider,
): { items: NoteAreaItem[]; heightPx: number } {
  const items: NoteAreaItem[] = [];
  let used = 0;
  while (queue.length > 0) {
    const entry = queue[0];
    const block = entry.note.blocks[entry.blockIndex];
    if (!block) {
      queue.shift();
      continue;
    }
    const m = noteMetrics.noteMetricsFor(entry.note, entry.blockIndex, block, contentWidthPx);
    const lineCount = m.lines.length;
    const gap = items.length === 0 ? 0 : NOTE_GAP_PX;
    const remaining =
      entry.startLine === 0 ? m.heightPx : sliceHeight(m, entry.startLine, lineCount);
    if (used + gap + remaining <= capacityPx + EPS) {
      items.push(
        entry.startLine === 0
          ? { noteId: entry.note.noteId, blockIndex: entry.blockIndex }
          : {
              noteId: entry.note.noteId,
              blockIndex: entry.blockIndex,
              startLine: entry.startLine,
              endLine: lineCount,
            },
      );
      used += gap + remaining;
      if (entry.blockIndex + 1 < entry.note.blocks.length) {
        queue[0] = { ...entry, blockIndex: entry.blockIndex + 1, startLine: 0 };
      } else {
        queue.shift();
      }
      continue;
    }
    // Does not fit whole: split at a line boundary, or defer to the next
    // page's area when this one already holds something.
    const available = capacityPx - used - gap;
    const cannotSplit = lineCount <= 1;
    let fit = cannotSplit ? 0 : linesFitting(m, entry.startLine, available);
    if (fit <= 0) {
      if (items.length > 0) {
        break;
      }
      if (cannotSplit) {
        // Nothing fits even in an empty area: place the unsplittable block
        // whole (clipped) rather than looping forever.
        items.push({ noteId: entry.note.noteId, blockIndex: entry.blockIndex });
        used += remaining;
        if (entry.blockIndex + 1 < entry.note.blocks.length) {
          queue[0] = { ...entry, blockIndex: entry.blockIndex + 1, startLine: 0 };
        } else {
          queue.shift();
        }
        continue;
      }
      fit = 1;
    }
    items.push({
      noteId: entry.note.noteId,
      blockIndex: entry.blockIndex,
      startLine: entry.startLine,
      endLine: entry.startLine + fit,
    });
    used += gap + sliceHeight(m, entry.startLine, entry.startLine + fit);
    queue[0] = { ...entry, startLine: entry.startLine + fit };
    break;
  }
  return { items, heightPx: used };
}

interface FootnotePlan {
  areas: (NoteArea | null)[];
  reserves: number[];
  carry: QueueEntry[];
  pageByNote: Map<string, number>;
}

function planFootnoteAreas(
  model: DocumentModel,
  pages: PageLayout[],
  refs: NoteRefLocation[],
  refPages: number[],
  noteMetrics: NoteMetricsProvider,
): FootnotePlan {
  const notesPerPage: NoteContent[][] = pages.map(() => []);
  const seen = new Set<string>();
  refs.forEach((ref, index) => {
    if (seen.has(ref.noteId)) {
      return;
    }
    seen.add(ref.noteId);
    const note = noteById(model, ref.noteId);
    if (!note) {
      return;
    }
    const page = Math.min(Math.max(0, refPages[index] ?? 0), pages.length - 1);
    notesPerPage[page].push(note);
  });
  const areas: (NoteArea | null)[] = pages.map(() => null);
  const reserves: number[] = pages.map(() => 0);
  const pageByNote = new Map<string, number>();
  let queue: QueueEntry[] = [];
  for (let p = 0; p < pages.length; p += 1) {
    for (const note of notesPerPage[p]) {
      queue.push({ note, blockIndex: 0, startLine: 0 });
    }
    if (queue.length === 0) {
      continue;
    }
    const geometry = sectionGeometryPx(model, pages[p].sectionIndex);
    const capacity =
      geometry.contentHeightPx * MAX_FOOTNOTE_AREA_FRACTION - NOTE_SEPARATOR_HEIGHT_PX;
    const width = noteContentWidthPx(model, pages[p].sectionIndex);
    const { items, heightPx } = fillArea(queue, capacity, width, noteMetrics);
    if (items.length > 0) {
      areas[p] = { kind: "footnote", heightPx: heightPx + NOTE_SEPARATOR_HEIGHT_PX, items };
      reserves[p] = heightPx + NOTE_SEPARATOR_HEIGHT_PX;
      for (const item of items) {
        if (!pageByNote.has(item.noteId)) {
          pageByNote.set(item.noteId, p);
        }
      }
    }
  }
  return { areas, reserves, carry: queue, pageByNote };
}

// --- used page height (for the endnote area's starting position) ------------

// Replicates the engine's cursor arithmetic (placeWhole/placeSlice) over a
// laid-out page so the endnote planner knows where the body content ends.
function usedHeightPx(model: DocumentModel, metrics: MetricsProvider, page: PageLayout): number {
  const states = new Map<number, { cursor: number; pending: number; count: number }>();
  for (const item of page.items) {
    const block = model.sections[item.sectionIndex]?.blocks[item.blockIndex];
    if (!block) {
      continue;
    }
    const column = item.columnIndex ?? 0;
    const state = states.get(column) ?? { cursor: 0, pending: 0, count: 0 };
    const width = sectionColumnFlowWidthPx(model, item.sectionIndex);
    const m = metrics.metricsFor(item.sectionIndex, item.blockIndex, block, width);
    if (item.startLine === undefined) {
      const gap = state.count === 0 ? m.marginTopPx : Math.max(state.pending, m.marginTopPx);
      state.cursor += gap + m.heightPx;
      state.pending = m.marginBottomPx;
    } else {
      const marginTop = item.startLine === 0 ? m.marginTopPx : 0;
      const gap = state.count === 0 ? marginTop : Math.max(state.pending, marginTop);
      const endLine = item.endLine ?? m.lines.length;
      state.cursor += gap + sliceHeight(m, item.startLine, endLine);
      state.pending = endLine >= m.lines.length ? m.marginBottomPx : 0;
    }
    state.count += 1;
    states.set(column, state);
  }
  let used = 0;
  for (const state of states.values()) {
    used = Math.max(used, state.cursor);
  }
  return used;
}

// --- the combined layout -----------------------------------------------------

function sameReserves(a: number[], b: number[]): boolean {
  if (a.length !== b.length) {
    return false;
  }
  return a.every((value, index) => Math.abs(value - b[index]) <= EPS);
}

function syntheticPage(model: DocumentModel, sectionIndex: number): PageLayout {
  const section = Math.min(Math.max(0, sectionIndex), model.sections.length - 1);
  return {
    sectionIndex: section,
    pageSize: { ...(model.sections[section]?.pageSize ?? model.defaultStyles.pageSize) },
    margins: { ...(model.sections[section]?.margins ?? model.defaultStyles.margins) },
    items: [],
  };
}

/** Lays the whole document out with its note areas (see the header
 * comment). Without any note references this is exactly `paginate`. */
export function layoutWithNotes(
  model: DocumentModel,
  metrics: MetricsProvider,
  noteMetrics: NoteMetricsProvider,
  lineOffsetsOf: (block: Block) => number[] | null,
  isHidden: (sectionIndex: number, blockIndex: number) => boolean = () => false,
): NotesLayout {
  const footnoteRefs = collectNoteRefs(model, "footnote");
  const endnoteRefs = collectNoteRefs(model, "endnote");

  let pages = paginate(model, metrics, undefined, isHidden);
  let refPages: number[] = [];
  let plan: FootnotePlan = { areas: pages.map(() => null), reserves: [], carry: [], pageByNote: new Map() };

  if (footnoteRefs.length > 0) {
    let reserves: number[] = [];
    for (let pass = 0; pass < MAX_NOTE_LAYOUT_PASSES; pass += 1) {
      refPages = assignRefsToPages(model, pages, footnoteRefs, lineOffsetsOf);
      plan = planFootnoteAreas(model, pages, footnoteRefs, refPages, noteMetrics);
      if (sameReserves(plan.reserves, reserves)) {
        break;
      }
      reserves = plan.reserves;
      pages = paginate(model, metrics, (page) => reserves[page] ?? 0, isHidden);
    }
    // The last pagination may have shifted references: the rendered areas
    // must match the final pages, so re-plan once against them.
    refPages = assignRefsToPages(model, pages, footnoteRefs, lineOffsetsOf);
    plan = planFootnoteAreas(model, pages, footnoteRefs, refPages, noteMetrics);
  }

  const allPages = [...pages];
  const footnoteAreas: (NoteArea | null)[] = [...plan.areas];
  // Footnote content that outgrew the last page continues on note-only
  // pages appended after the body (AC-13).
  let carry = plan.carry;
  let extra = 0;
  while (carry.length > 0 && extra < MAX_EXTRA_NOTE_PAGES) {
    const page = syntheticPage(model, allPages[allPages.length - 1].sectionIndex);
    const geometry = sectionGeometryPx(model, page.sectionIndex);
    const width = noteContentWidthPx(model, page.sectionIndex);
    const capacity = geometry.contentHeightPx - NOTE_SEPARATOR_HEIGHT_PX;
    const { items, heightPx } = fillArea(carry, capacity, width, noteMetrics);
    if (items.length === 0) {
      break;
    }
    allPages.push(page);
    footnoteAreas.push({ kind: "footnote", heightPx: heightPx + NOTE_SEPARATOR_HEIGHT_PX, items });
    for (const item of items) {
      if (!plan.pageByNote.has(item.noteId)) {
        plan.pageByNote.set(item.noteId, allPages.length - 1);
      }
    }
    extra += 1;
  }

  // Endnotes: one area after the last content of the document (AC-4),
  // continuing onto appended pages when it outgrows the remaining space.
  const endnoteAreas: (NoteArea | null)[] = allPages.map(() => null);
  if (endnoteRefs.length > 0) {
    const seen = new Set<string>();
    let queue: QueueEntry[] = [];
    for (const ref of endnoteRefs) {
      if (seen.has(ref.noteId)) {
        continue;
      }
      seen.add(ref.noteId);
      const note = noteById(model, ref.noteId);
      if (note) {
        queue.push({ note, blockIndex: 0, startLine: 0 });
      }
    }
    let pageIndex = allPages.length - 1;
    const lastEndnotePage = pageIndex + MAX_EXTRA_NOTE_PAGES;
    while (queue.length > 0 && pageIndex <= lastEndnotePage) {
      let page = allPages[pageIndex];
      if (!page) {
        page = syntheticPage(model, allPages[allPages.length - 1].sectionIndex);
        allPages.push(page);
        footnoteAreas.push(null);
        endnoteAreas.push(null);
      }
      const geometry = sectionGeometryPx(model, page.sectionIndex);
      const width = noteContentWidthPx(model, page.sectionIndex);
      const used = usedHeightPx(model, metrics, page);
      const footnoteReserve = footnoteAreas[pageIndex]?.heightPx ?? 0;
      const capacity =
        geometry.contentHeightPx - used - footnoteReserve - NOTE_SEPARATOR_HEIGHT_PX;
      if (capacity > EPS) {
        const { items, heightPx } = fillArea(queue, capacity, width, noteMetrics);
        if (items.length > 0) {
          endnoteAreas[pageIndex] = {
            kind: "endnote",
            heightPx: heightPx + NOTE_SEPARATOR_HEIGHT_PX,
            items,
          };
        }
      }
      pageIndex += 1;
    }
  }

  return {
    pages: allPages,
    footnoteAreas,
    endnoteAreas,
    footnotePageByNote: plan.pageByNote,
    footnoteRefPages: refPages,
  };
}
