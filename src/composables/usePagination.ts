// The pagination controller (SPEC-028): owns the measurement cache and the
// measure → layout loop that turns the pure engine (lib/pagination) plus
// measurements into the pages the canvas renders.
//
// Every model change re-runs the LAYOUT synchronously (cheap math), so the
// rendered pages always agree with the model's block indices; MEASUREMENT
// is what stays incremental and asynchronous: metrics are cached per block
// object (commands replace only the blocks they touch — see commands.ts),
// so a keystroke re-measures one block, and a 200-page document settles
// within the AC-50 budget without ever blocking the UI (AC-49..51).
//
// Two measurement backends share the cache:
//  - blocks rendered on a realized page are measured off their own DOM
//    elements (lib/pageMetrics) — exact, and the only option for tables,
//    images and break markers;
//  - unrendered paragraphs (their pages are virtualized empty boxes — see
//    DocumentView's realization) are measured with canvas text metrics
//    (lib/textMetrics), which also yields their line-start offsets for
//    free. This is what makes the bulk pass over a 200-page document cheap:
//    no DOM is mounted to measure it.
// Unrendered non-paragraphs keep their estimates ("deferred"): the settle
// loop finishes once a full pass makes no progress, and DocumentView
// re-schedules when realization changes so they are measured the moment
// they render.
import { ref, shallowRef, type Ref, type ShallowRef } from "vue";
import {
  isBreak,
  isImage,
  isTable,
  type Block,
  type DocumentModel,
  type NoteKind,
  type Paragraph,
} from "../lib/documentModel";
import {
  estimateMetrics,
  reconcilePages,
  sectionColumnFlowWidthPx,
  type BlockMetrics,
  type MetricsProvider,
  type PageLayout,
} from "../lib/pagination";
import { layoutWithNotes, type NoteArea, type NoteMetricsProvider } from "../lib/notesLayout";
import { paragraphWithNoteLabels, renderedOffsetToModel } from "../lib/notes";
import { renderedOffsetsToModel } from "../lib/domCaret";
import {
  combineSliceMetrics,
  lineStartOffsets,
  measureBlockElement,
} from "../lib/pageMetrics";
import {
  createCanvasTextMeasurer,
  measureParagraphText,
  type BaseFont,
  type CanvasTextMeasurer,
} from "../lib/textMetrics";
import { sliceStartOf } from "../lib/domCaret";
import { scheduleEagerWork, type IdleHandle } from "../lib/idleScheduler";
import type { CaretPosition } from "../lib/textEditing";

interface CacheEntry {
  metrics: BlockMetrics;
  /** Character offset each line starts at. Known immediately for
   * canvas-measured paragraphs (the wrap decided them); computed lazily
   * with Range hit-testing for DOM-measured blocks the layout actually
   * splits across pages. */
  lineOffsets?: number[];
}

interface PendingTarget {
  section: number;
  block: number;
  ref: Block;
}

/** One block's page placements, char-windowed for split paragraphs — the
 * caret page indicator's lookup table (AC-3). */
interface PagePlacement {
  page: number;
  startChar: number;
  endChar: number;
}

/** How long measurement must run before the "Paginating…" indicator shows —
 * brief re-measures (typing) never flash it. */
const BUSY_DELAY_MS = 200;

/** Per-slot measurement budget: slots chain on macrotasks (near-100% duty,
 * which is what gets a 200-page settle inside AC-50's budget), so this is
 * also the longest an input event can wait behind one slot. */
const SLOT_BUDGET_MS = 24;

/** How long after the user's last interaction bulk measurement stays
 * paused — their keystroke/scroll/click (and whatever work it triggers,
 * e.g. a find refresh) always wins the main thread. */
const INTERACTION_PAUSE_MS = 250;

/** Retry cadence while bulk work is paused for interaction. */
const PAUSED_RETRY_MS = 50;

/** Per-relayout budget for resolving DOM-measured split blocks' line-start
 * character offsets (Range hit-testing is the most expensive read this
 * module does). Canvas-measured paragraphs never need this. */
const OFFSETS_BUDGET_MS = 12;

/** Full passes over the pending list with zero progress before the loop
 * settles: what remains is deferred (unrendered non-paragraphs keeping
 * their estimates until realization renders them). Two passes, not one, so
 * a Vue patch that is about to render them gets a chance to land. */
const MAX_STALLED_PASSES = 2;

/** One relayout's note areas (SPEC-031), aligned with `pages` by index. */
export interface PageNoteAreas {
  footnotes: (NoteArea | null)[];
  endnotes: (NoteArea | null)[];
}

export interface PaginationController {
  pages: ShallowRef<PageLayout[]>;
  /** Per-page footnote/endnote areas (SPEC-031). */
  noteAreas: ShallowRef<PageNoteAreas>;
  /** The page each footnote reference renders on, aligned with
   * collectNoteRefs(model, "footnote") — noteLabels' per-page restart
   * input (SPEC-031 AC-19). */
  footnoteRefPages: ShallowRef<number[]>;
  /** True while a long measurement pass runs (drives "Paginating…"). */
  busy: Ref<boolean>;
  /** "measuring" while any rendered page still uses estimated metrics;
   * "clean" once the layout is fully measured — the canvas exposes it as
   * data-pagination-state so tests get a positive settled signal. */
  state: Ref<"clean" | "measuring">;
  /** Whether unrendered paragraphs can be measured without mounting them
   * (canvas metrics available) — the gate for page virtualization: without
   * it every page must realize so the DOM path can measure it. */
  canMeasureOffscreen: boolean;
  /** Re-lays-out synchronously and schedules measurement of anything the
   * cache misses. Call on every model change (and when realization renders
   * previously virtual pages, so their blocks get exact DOM metrics). */
  schedule: () => void;
  /** Drops every cached measurement (formatting-marks toggle, font load —
   * anything that reflows text without replacing block objects). */
  invalidateAll: () => void;
  /** Whether a block has real (non-estimated) metrics at its current
   * width — split rendering is only allowed then. */
  isMeasured: (block: Block, contentWidthPx: number) => boolean;
  /** The character offset each of the block's lines starts at, or null
   * until measured — the renderer's slice windows and the caret page
   * refinement both read it. */
  lineOffsetsFor: (block: Block) => number[] | null;
  /** The 0-based page the caret renders on. */
  pageOfCaret: (caret: CaretPosition) => number;
  stop: () => void;
}

function sameArea(a: NoteArea | null, b: NoteArea | null): boolean {
  if (a === b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }
  return (
    a.kind === b.kind &&
    Math.abs(a.heightPx - b.heightPx) < 0.5 &&
    a.items.length === b.items.length &&
    a.items.every((item, index) => {
      const other = b.items[index];
      return (
        item.noteId === other.noteId &&
        item.blockIndex === other.blockIndex &&
        item.startLine === other.startLine &&
        item.endLine === other.endLine
      );
    })
  );
}

function sameAreaLists(a: (NoteArea | null)[], b: (NoteArea | null)[]): boolean {
  return a.length === b.length && a.every((area, index) => sameArea(area, b[index]));
}

export function usePagination(
  docOf: () => DocumentModel | null,
  rootOf: () => HTMLElement | null,
  noteLabelOf: (noteId: string, kind: NoteKind) => string = () => "",
): PaginationController {
  const pages = shallowRef<PageLayout[]>([]);
  const noteAreas = shallowRef<PageNoteAreas>({ footnotes: [], endnotes: [] });
  const footnoteRefPages = shallowRef<number[]>([]);
  const busy = ref(false);
  const state = ref<"clean" | "measuring">("clean");

  let cache = new WeakMap<Block, CacheEntry>();
  let pending: PendingTarget[] = [];
  let placements = new Map<string, PagePlacement[]>();
  let slot: IdleHandle | null = null;
  let running = false;
  let stalledFrames = 0;
  let stalledPasses = 0;
  let busyTimer: ReturnType<typeof setTimeout> | null = null;
  let pausedTimer: ReturnType<typeof setTimeout> | null = null;
  // One `[data-section][data-block]` walk of the rendered canvas, indexed by
  // address, so slice lookup is O(1) instead of a full-tree querySelectorAll
  // per measured block — a large pass that way cost O(n^2) DOM walks, which
  // blew the AC-50 2s budget at ~200 pages. Invalidated whenever a relayout
  // assigns new pages (the only time the rendered slice set can change);
  // within a pass the DOM is stable, so one walk serves every block.
  let sliceIndex: Map<string, Element[]> | null = null;
  // Measurement cursor into `pending` between re-layouts: measuring does
  // not need a fresh layout every frame (elements are found in whatever is
  // currently rendered), so mid-pass re-layouts are throttled — they are
  // what makes a long pass expensive, not the measuring itself.
  let cursor = 0;
  let framesSinceLayout = 0;
  let measuredThisPass = 0;
  let lastInteraction = 0;

  // The canvas measurer's DOM probe host: the pages wrapper (the content
  // root's parent), OUTSIDE the contenteditable root so probe text can
  // never leak into canvas-vs-model text comparisons, but inside the app's
  // font context.
  const measurer: CanvasTextMeasurer | null = createCanvasTextMeasurer(
    () => rootOf()?.parentElement ?? null,
  );
  let baseFont: BaseFont | null = null;

  function baseFontOf(root: HTMLElement): BaseFont {
    if (!baseFont) {
      const style = root.ownerDocument.defaultView?.getComputedStyle(root);
      baseFont = {
        family: style?.fontFamily || "sans-serif",
        sizePx: style ? Number.parseFloat(style.fontSize) || 16 : 16,
      };
    }
    return baseFont;
  }

  function noteInteraction(): void {
    lastInteraction = performance.now();
  }
  if (typeof document !== "undefined") {
    document.addEventListener("keydown", noteInteraction, { capture: true, passive: true });
    document.addEventListener("pointerdown", noteInteraction, { capture: true, passive: true });
    document.addEventListener("input", noteInteraction, { capture: true, passive: true });
  }

  function entryFor(block: Block, contentWidthPx: number): CacheEntry | null {
    const entry = cache.get(block);
    return entry && Math.abs(entry.metrics.contentWidthPx - contentWidthPx) < 0.5 ? entry : null;
  }

  function relayout(model: DocumentModel): { unresolved: number; resolved: number } {
    pending = [];
    const seen = new Set<Block>();
    const provider: MetricsProvider = {
      metricsFor(sectionIndex, blockIndex, block, contentWidthPx) {
        const entry = entryFor(block, contentWidthPx);
        if (entry) {
          return entry.metrics;
        }
        if (!seen.has(block)) {
          seen.add(block);
          pending.push({ section: sectionIndex, block: blockIndex, ref: block });
        }
        return estimateMetrics(block, contentWidthPx, model.defaultStyles.fontSizePt);
      },
    };
    // Note content (SPEC-031) is measured synchronously with canvas metrics
    // (notes are small and never DOM-measured — their areas render from the
    // same cache these measurements land in).
    const noteProvider: NoteMetricsProvider = {
      noteMetricsFor(_note, _blockIndex, block, contentWidthPx) {
        const entry = entryFor(block, contentWidthPx);
        if (entry) {
          return entry.metrics;
        }
        const root = rootOf();
        if (root && measurer && isParagraphBlock(block)) {
          const measured = measureParagraphText(
            block,
            contentWidthPx,
            baseFontOf(root),
            model.showFormattingMarks === true,
            measurer,
          );
          if (measured) {
            cache.set(block, { metrics: measured.metrics, lineOffsets: measured.lineOffsets });
            return measured.metrics;
          }
        }
        return estimateMetrics(block, contentWidthPx, model.defaultStyles.fontSizePt);
      },
    };
    const notesLayout = layoutWithNotes(
      model,
      provider,
      noteProvider,
      (block) => cache.get(block)?.lineOffsets ?? null,
    );
    const layout = notesLayout.pages;
    if (
      !sameAreaLists(noteAreas.value.footnotes, notesLayout.footnoteAreas) ||
      !sameAreaLists(noteAreas.value.endnotes, notesLayout.endnoteAreas)
    ) {
      noteAreas.value = {
        footnotes: notesLayout.footnoteAreas,
        endnotes: notesLayout.endnoteAreas,
      };
    }
    if (
      footnoteRefPages.value.length !== notesLayout.footnoteRefPages.length ||
      footnoteRefPages.value.some((page, index) => page !== notesLayout.footnoteRefPages[index])
    ) {
      footnoteRefPages.value = notesLayout.footnoteRefPages;
    }
    const offsets = ensureLineOffsets(model, layout);
    buildPlacements(model, layout);
    const reconciled = reconcilePages(pages.value, layout);
    // Skip the assignment when nothing changed — during a long measurement
    // pass most re-layouts refine only later pages, and every assignment
    // costs a full render-entry rebuild downstream.
    if (
      reconciled.length !== pages.value.length ||
      reconciled.some((entry, index) => entry !== pages.value[index])
    ) {
      pages.value = reconciled;
      sliceIndex = null; // The rendered slice set is about to change.
    }
    return { unresolved: pending.length + offsets.missing, resolved: offsets.resolved };
  }

  // Character offsets for every split block the new layout contains that
  // still lacks them (DOM-measured blocks only — canvas measurement stores
  // offsets at measure time), read from the block's CURRENT rendering
  // (whole, or sliced by the previous layout — either holds all its
  // lines). Budgeted: Range hit-testing is expensive, so leftovers are
  // counted and resolved on later frames — an unresolved split renders
  // whole meanwhile (see renderEntryFor).
  function ensureLineOffsets(
    model: DocumentModel,
    layout: PageLayout[],
  ): { missing: number; resolved: number } {
    const root = rootOf();
    if (!root) {
      return { missing: 0, resolved: 0 };
    }
    const started = performance.now();
    let missing = 0;
    let resolved = 0;
    for (const page of layout) {
      for (const item of page.items) {
        if (item.startLine === undefined && item.endLine === undefined) {
          continue;
        }
        const block = model.sections[item.sectionIndex]?.blocks[item.blockIndex];
        if (!block) {
          continue;
        }
        const entry = cache.get(block);
        if (!entry || entry.lineOffsets) {
          continue;
        }
        if (resolved > 0 && performance.now() - started > OFFSETS_BUDGET_MS) {
          missing += 1;
          continue;
        }
        const slices = sliceElements(root, item.sectionIndex, item.blockIndex);
        if (slices.length === 0) {
          continue;
        }
        const offsets: number[] = [];
        for (const element of slices) {
          const base = sliceStartOf(element);
          const measured = measureBlockElement(element, entry.metrics.contentWidthPx, true);
          // DOM-measured offsets are in rendered text space; note reference
          // labels (SPEC-031) may be longer than their model placeholder.
          for (const offset of renderedOffsetsToModel(
            element,
            lineStartOffsets(element, measured.lines),
          )) {
            offsets.push(base + offset);
          }
        }
        entry.lineOffsets = offsets;
        resolved += 1;
      }
    }
    return { missing, resolved };
  }

  function buildPlacements(model: DocumentModel, layout: PageLayout[]): void {
    placements = new Map();
    for (let p = 0; p < layout.length; p += 1) {
      for (const item of layout[p].items) {
        const key = `${item.sectionIndex}:${item.blockIndex}`;
        const list = placements.get(key) ?? [];
        if (item.startLine === undefined) {
          list.push({ page: p, startChar: 0, endChar: Number.POSITIVE_INFINITY });
        } else {
          const block = model.sections[item.sectionIndex]?.blocks[item.blockIndex];
          const offsets = block ? cache.get(block)?.lineOffsets : undefined;
          if (!offsets) {
            // Unmeasured splits render whole on their first page.
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
  }

  function sliceElements(root: HTMLElement, section: number, block: number): Element[] {
    if (sliceIndex === null) {
      // One walk, grouped by address. Row elements inside tables repeat the
      // parent block's data-block (SPEC-022) and are not slices, so they are
      // excluded the same way the per-address query was.
      const index = new Map<string, Element[]>();
      for (const el of Array.from(root.querySelectorAll("[data-section][data-block]:not([data-row])"))) {
        const address = `${el.getAttribute("data-section")}:${el.getAttribute("data-block")}`;
        const group = index.get(address);
        if (group) {
          group.push(el);
        } else {
          index.set(address, [el]);
        }
      }
      for (const group of index.values()) {
        group.sort((a, b) => sliceStartOf(a) - sliceStartOf(b));
      }
      sliceIndex = index;
    }
    return sliceIndex.get(`${section}:${block}`) ?? [];
  }

  function isParagraphBlock(block: Block): block is Paragraph {
    return !isTable(block) && !isImage(block) && !isBreak(block);
  }

  function measureOne(root: HTMLElement, model: DocumentModel, target: PendingTarget): boolean {
    const block = model.sections[target.section]?.blocks[target.block];
    if (block !== target.ref) {
      return true; // the model moved on; the new layout re-listed what matters
    }
    const width = sectionColumnFlowWidthPx(model, target.section);
    const elements = sliceElements(root, target.section, target.block);
    if (elements.length > 0) {
      const splittable = isParagraphBlock(block);
      const metrics = combineSliceMetrics(
        elements.map((element) => measureBlockElement(element, width, splittable)),
      );
      cache.set(block, { metrics });
      return true;
    }
    // Not rendered (its page is a virtualized empty box): paragraphs are
    // measured with canvas text metrics; anything else keeps its estimate
    // until realization renders it. Note reference marks measure at their
    // rendered label text (SPEC-031), with line offsets mapped back to
    // model space.
    if (measurer && isParagraphBlock(block)) {
      const rendered = paragraphWithNoteLabels(block, noteLabelOf);
      const measured = measureParagraphText(
        rendered,
        width,
        baseFontOf(root),
        model.showFormattingMarks === true,
        measurer,
      );
      if (measured) {
        const lineOffsets =
          rendered === block
            ? measured.lineOffsets
            : measured.lineOffsets.map((offset) => renderedOffsetToModel(block, noteLabelOf, offset));
        cache.set(block, { metrics: measured.metrics, lineOffsets });
        return true;
      }
    }
    return false;
  }

  function finish(): void {
    running = false;
    stalledFrames = 0;
    stalledPasses = 0;
    if (busyTimer !== null) {
      clearTimeout(busyTimer);
      busyTimer = null;
    }
    busy.value = false;
    state.value = "clean";
  }

  // Measurement chains in budgeted macrotask slots: input events and
  // rendering interleave between slots, which is what keeps a 200-page
  // pass from ever freezing or janking the UI (AC-49/51), while the
  // near-100% duty cycle keeps the settle inside AC-50's budget.
  function measureLoop(): void {
    slot = scheduleEagerWork((budgetMs) => {
      const root = rootOf();
      const model = docOf();
      if (!root || !model) {
        finish();
        return;
      }
      // Bulk work yields to live interaction: while the user is typing or
      // clicking, a large pass waits (their input owns the thread); a tiny
      // pass (the one block a keystroke re-measures) still runs immediately
      // so split paragraphs update under the caret.
      if (
        pending.length - cursor > 3 &&
        performance.now() - lastInteraction < INTERACTION_PAUSE_MS
      ) {
        pausedTimer = setTimeout(() => {
          pausedTimer = null;
          measureLoop();
        }, PAUSED_RETRY_MS);
        return;
      }
      const started = performance.now();
      const startCursor = cursor;
      let measured = 0;
      while (cursor < pending.length) {
        // The budget break holds regardless of progress — skipped (deferred)
        // targets count too — so one idle slot can never turn into an
        // unbounded walk; at least one target advances per slot.
        if (cursor > startCursor && performance.now() - started > budgetMs) {
          break;
        }
        if (measureOne(root, model, pending[cursor])) {
          measured += 1;
        }
        cursor += 1;
      }
      measuredThisPass += measured;
      framesSinceLayout += 1;
      // Re-layouts are the expensive part of a pass (page fill + offset
      // resolution + a render-entry rebuild downstream), so they happen
      // only when the pending list is exhausted — with a generous safety
      // valve — never every frame.
      const exhausted = cursor >= pending.length;
      let relaidOut = false;
      if (exhausted || framesSinceLayout >= 60) {
        const { unresolved, resolved } = relayout(model);
        relaidOut = true;
        cursor = 0;
        framesSinceLayout = 0;
        if (unresolved === 0) {
          finish();
          return;
        }
        // A full pass that neither measured a block nor resolved offsets
        // made no progress: what's left is deferred until realization
        // renders it. Settle rather than spin — DocumentView re-schedules
        // when realization changes.
        if (exhausted) {
          stalledPasses = measuredThisPass === 0 && resolved === 0 ? stalledPasses + 1 : 0;
          measuredThisPass = 0;
          if (stalledPasses >= MAX_STALLED_PASSES) {
            finish();
            return;
          }
        }
      }
      // No progress at all — nothing measured, no movement through the
      // pending list, no relayout this frame: retry a few frames, then give
      // up until the next schedule() rather than spinning.
      stalledFrames =
        measured === 0 && cursor === startCursor && !relaidOut ? stalledFrames + 1 : 0;
      if (stalledFrames > 8) {
        finish();
        return;
      }
      measureLoop();
    }, SLOT_BUDGET_MS);
  }

  function startMeasuring(): void {
    state.value = "measuring";
    if (busyTimer === null && !busy.value) {
      busyTimer = setTimeout(() => {
        busyTimer = null;
        busy.value = true;
      }, BUSY_DELAY_MS);
    }
    if (!running) {
      running = true;
      stalledFrames = 0;
      stalledPasses = 0;
      measuredThisPass = 0;
      measureLoop();
    }
  }

  function schedule(): void {
    const model = docOf();
    if (!model) {
      pages.value = [];
      finish();
      return;
    }
    const { unresolved } = relayout(model);
    cursor = 0;
    framesSinceLayout = 0;
    if (unresolved === 0) {
      finish();
      return;
    }
    if (typeof requestAnimationFrame !== "function") {
      // Non-browser environment: keep the estimated layout.
      finish();
      return;
    }
    startMeasuring();
  }

  function invalidateAll(): void {
    cache = new WeakMap();
    baseFont = null;
    measurer?.reset();
    schedule();
  }

  function isMeasured(block: Block, contentWidthPx: number): boolean {
    return entryFor(block, contentWidthPx) !== null;
  }

  function lineOffsetsFor(block: Block): number[] | null {
    return cache.get(block)?.lineOffsets ?? null;
  }

  function pageOfCaret(caret: CaretPosition): number {
    const list = placements.get(`${caret.section}:${caret.block}`);
    if (!list || list.length === 0) {
      return Math.max(0, pages.value.length - 1);
    }
    for (const placement of list) {
      if (caret.offset >= placement.startChar && caret.offset < placement.endChar) {
        return placement.page;
      }
    }
    return list[list.length - 1].page;
  }

  function stop(): void {
    if (typeof document !== "undefined") {
      document.removeEventListener("keydown", noteInteraction, { capture: true });
      document.removeEventListener("pointerdown", noteInteraction, { capture: true });
      document.removeEventListener("input", noteInteraction, { capture: true });
    }
    slot?.cancel();
    slot = null;
    if (pausedTimer !== null) {
      clearTimeout(pausedTimer);
      pausedTimer = null;
    }
    finish();
  }

  // Web fonts arriving reflow every line: measurements taken before they
  // loaded are wrong, so drop them all once.
  if (typeof document !== "undefined" && document.fonts?.ready) {
    void document.fonts.ready.then(() => {
      if (docOf()) {
        invalidateAll();
      }
    });
  }

  return {
    pages,
    noteAreas,
    footnoteRefPages,
    busy,
    state,
    canMeasureOffscreen: measurer !== null,
    schedule,
    invalidateAll,
    isMeasured,
    lineOffsetsFor,
    pageOfCaret,
    stop,
  };
}
