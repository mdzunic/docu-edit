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

/** Per-frame measurement budget: long enough to make fast progress on a
 * 200-page open, short enough to leave the frame responsive (AC-34). */
const FRAME_BUDGET_MS = 8;

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
  let rafHandle = 0;
  let running = false;
  let stalledFrames = 0;
  let busyTimer: ReturnType<typeof setTimeout> | null = null;

  function entryFor(block: Block, contentWidthPx: number): CacheEntry | null {
    const entry = cache.get(block);
    return entry && Math.abs(entry.metrics.contentWidthPx - contentWidthPx) < 0.5 ? entry : null;
  }

  function relayout(model: DocumentModel): void {
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
    ensureLineOffsets(model, layout);
    buildPlacements(model, layout);
    pages.value = reconcilePages(pages.value, layout);
  }

  // Character offsets for every split block the new layout contains, read
  // from the block's CURRENT rendering (whole, or sliced by the previous
  // layout — either holds all its lines).
  function ensureLineOffsets(model: DocumentModel, layout: PageLayout[]): void {
    const root = rootOf();
    if (!root) {
      return;
    }
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
      }
    }
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
    return Array.from(
      root.querySelectorAll(
        `[data-section="${section}"][data-block="${block}"]:not([data-row])`,
      ),
    ).sort((a, b) => sliceStartOf(a) - sliceStartOf(b));
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

  function measureLoop(): void {
    rafHandle = requestAnimationFrame(() => {
      const root = rootOf();
      const model = docOf();
      if (!root || !model) {
        finish();
        return;
      }
      const started = performance.now();
      let measured = 0;
      for (const target of pending) {
        if (measured > 0 && performance.now() - started > FRAME_BUDGET_MS) {
          break;
        }
        if (measureOne(root, model, target)) {
          measured += 1;
        }
      }
      relayout(model);
      if (pending.length === 0) {
        finish();
        return;
      }
      // No element found for anything pending (render not flushed yet):
      // retry a few frames, then give up until the next schedule() rather
      // than spinning forever.
      stalledFrames = measured === 0 ? stalledFrames + 1 : 0;
      if (stalledFrames > 5) {
        finish();
        return;
      }
      measureLoop();
    });
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
    relayout(model);
    if (pending.length === 0) {
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
    if (rafHandle && typeof cancelAnimationFrame === "function") {
      cancelAnimationFrame(rafHandle);
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
