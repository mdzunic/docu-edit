// The pagination controller (SPEC-028): owns the measurement cache and the
// measure → layout loop that turns the pure engine (lib/pagination) plus
// DOM measurements (lib/pageMetrics) into the pages the canvas renders.
//
// Every model change re-runs the LAYOUT synchronously (cheap math), so the
// rendered pages always agree with the model's block indices; MEASUREMENT
// is what stays incremental and asynchronous: metrics are cached per block
// object (commands replace only the blocks they touch — see commands.ts),
// so a keystroke re-measures one block, in a requestAnimationFrame batch,
// and a 200-page document opens with estimated pages that refine over a
// few frames without ever blocking the UI (AC-34, AC-35).
import { ref, shallowRef, type Ref, type ShallowRef } from "vue";
import {
  isBreak,
  isImage,
  isTable,
  type Block,
  type DocumentModel,
} from "../lib/documentModel";
import {
  estimateMetrics,
  paginate,
  reconcilePages,
  sectionGeometryPx,
  type BlockMetrics,
  type MetricsProvider,
  type PageLayout,
} from "../lib/pagination";
import {
  combineSliceMetrics,
  lineStartOffsets,
  measureBlockElement,
} from "../lib/pageMetrics";
import { sliceStartOf } from "../lib/domCaret";
import { scheduleIdleWork, type IdleHandle } from "../lib/idleScheduler";
import type { CaretPosition } from "../lib/textEditing";

interface CacheEntry {
  metrics: BlockMetrics;
  /** Character offset each line starts at — computed lazily, only for
   * blocks the layout actually splits across pages. */
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

/** Per-slot measurement budget when no idle deadline is available. */
const FRAME_BUDGET_MS = 4;

/** Ceiling per idle slot even when the browser reports more idle time. */
const IDLE_BUDGET_MS = 12;

/** How long after the user's last interaction bulk measurement stays
 * paused — their keystroke/scroll/click (and whatever work it triggers,
 * e.g. a find refresh) always wins the main thread. */
const INTERACTION_PAUSE_MS = 400;

/** Per-relayout budget for resolving split blocks' line-start character
 * offsets (Range hit-testing is the most expensive read this module does —
 * a 200-page pass resolves them across several frames, not in one). */
const OFFSETS_BUDGET_MS = 12;

export interface PaginationController {
  pages: ShallowRef<PageLayout[]>;
  /** True while a long measurement pass runs (drives "Paginating…"). */
  busy: Ref<boolean>;
  /** "measuring" while any rendered page still uses estimated metrics;
   * "clean" once the layout is fully measured — the canvas exposes it as
   * data-pagination-state so tests get a positive settled signal. */
  state: Ref<"clean" | "measuring">;
  /** Re-lays-out synchronously and schedules measurement of anything the
   * cache misses. Call on every model change. */
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

export function usePagination(
  docOf: () => DocumentModel | null,
  rootOf: () => HTMLElement | null,
): PaginationController {
  const pages = shallowRef<PageLayout[]>([]);
  const busy = ref(false);
  const state = ref<"clean" | "measuring">("clean");

  let cache = new WeakMap<Block, CacheEntry>();
  let pending: PendingTarget[] = [];
  let placements = new Map<string, PagePlacement[]>();
  let slot: IdleHandle | null = null;
  let running = false;
  let stalledFrames = 0;
  let busyTimer: ReturnType<typeof setTimeout> | null = null;
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
  let lastInteraction = 0;

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

  // Re-lays-out from the current cache. Returns the number of blocks still
  // unresolved (unmeasured, or split without line offsets yet) so the
  // measure loop knows whether the layout has fully settled.
  function relayout(model: DocumentModel): number {
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
    const layout = paginate(model, provider);
    const missingOffsets = ensureLineOffsets(model, layout);
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
    return pending.length + missingOffsets;
  }

  // Character offsets for every split block the new layout contains, read
  // from the block's CURRENT rendering (whole, or sliced by the previous
  // layout — either holds all its lines). Budgeted: Range hit-testing is
  // expensive, so leftovers are counted and resolved on later frames — an
  // unresolved split renders whole meanwhile (see renderEntryFor).
  function ensureLineOffsets(model: DocumentModel, layout: PageLayout[]): number {
    const root = rootOf();
    if (!root) {
      return 0;
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
          for (const offset of lineStartOffsets(element, measured.lines)) {
            offsets.push(base + offset);
          }
        }
        entry.lineOffsets = offsets;
        resolved += 1;
      }
    }
    return missing;
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

  function measureOne(root: HTMLElement, model: DocumentModel, target: PendingTarget): boolean {
    const block = model.sections[target.section]?.blocks[target.block];
    if (block !== target.ref) {
      return true; // the model moved on; the new layout re-listed what matters
    }
    const elements = sliceElements(root, target.section, target.block);
    if (elements.length === 0) {
      return false;
    }
    const width = sectionGeometryPx(model, target.section).contentWidthPx;
    const splittable = !isTable(block) && !isImage(block) && !isBreak(block);
    const metrics = combineSliceMetrics(
      elements.map((element) => measureBlockElement(element, width, splittable)),
    );
    cache.set(block, { metrics });
    return true;
  }

  function finish(): void {
    running = false;
    stalledFrames = 0;
    if (busyTimer !== null) {
      clearTimeout(busyTimer);
      busyTimer = null;
    }
    busy.value = false;
    state.value = "clean";
  }

  // Measurement runs in idle time where the browser offers it — user input
  // and rendering preempt it naturally, which is what keeps a 200-page
  // pass from ever freezing or janking the UI (AC-34).
  function measureLoop(): void {
    slot = scheduleIdleWork((budgetMs) => {
      const root = rootOf();
      const model = docOf();
      if (!root || !model) {
        finish();
        return;
      }
      // Bulk work yields to live interaction: while the user is typing or
      // clicking, a large pass waits (their input owns the thread — AC-34);
      // a tiny pass (the one block a keystroke re-measures) still runs
      // immediately so split paragraphs update under the caret.
      if (
        pending.length - cursor > 3 &&
        performance.now() - lastInteraction < INTERACTION_PAUSE_MS
      ) {
        measureLoop();
        return;
      }
      const started = performance.now();
      let measured = 0;
      while (cursor < pending.length) {
        if (measured > 0 && performance.now() - started > budgetMs) {
          break;
        }
        if (measureOne(root, model, pending[cursor])) {
          measured += 1;
        }
        cursor += 1;
      }
      framesSinceLayout += 1;
      // Re-layouts are the expensive part of a pass (page fill + offset
      // resolution + a render-entry rebuild downstream), so they happen
      // only when the pending list is exhausted — with a generous safety
      // valve — never every frame.
      const exhausted = cursor >= pending.length;
      let unresolved = -1;
      if (exhausted || framesSinceLayout >= 60) {
        unresolved = relayout(model);
        cursor = 0;
        framesSinceLayout = 0;
        if (unresolved === 0) {
          finish();
          return;
        }
      }
      // No progress at all (elements not rendered yet, or only offsets
      // remain and a relayout did not run this frame): retry a few frames,
      // then give up until the next schedule() rather than spinning.
      stalledFrames = measured === 0 && unresolved === -1 ? stalledFrames + 1 : 0;
      if (stalledFrames > 8) {
        finish();
        return;
      }
      measureLoop();
    }, { timeoutMs: 200, fallbackBudgetMs: FRAME_BUDGET_MS, maxBudgetMs: IDLE_BUDGET_MS });
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
    const unresolved = relayout(model);
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
    busy,
    state,
    schedule,
    invalidateAll,
    isMeasured,
    lineOffsetsFor,
    pageOfCaret,
    stop,
  };
}
