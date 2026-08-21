import { nextTick, onBeforeUnmount, onMounted, ref, watch, type Ref } from "vue";
import { collapsedBands, planCollapse } from "../lib/toolbarOverflow";

export interface UseToolbarOverflowOptions {
  /** The strip element whose available width is measured (AC-14). */
  strip: () => HTMLElement | null;
  /** The band numbers present, left to right (e.g. [1..8]). */
  bandOrder: number[];
  /** Px reserved on the trailing edge for open side panes, so opening a pane
   * collapses the strip exactly as narrowing the window does (AC-15). */
  reserved: Ref<number>;
  /** Extra slack before a collapsed band re-expands (AC hysteresis). */
  hysteresis?: number;
}

/** Drives whole-band overflow collapse from the strip's real available width.
 * Band widths are measured once every band is expanded and cached; a
 * ResizeObserver on the strip re-plans on width changes, and settle-timer
 * re-measures cover the UI font loading in after first paint (the observer
 * does not re-fire once the strip clips its own overflow). */
export function useToolbarOverflow({
  strip,
  bandOrder,
  reserved,
  hysteresis = 24,
}: UseToolbarOverflowOptions) {
  const collapsedCount = ref(0);
  // True while the strip force-expands every band so their natural widths can
  // be measured; DocumentToolbar disables the collapse teleports meanwhile.
  const measuring = ref(false);

  const bandWidths: Record<number, number> = {};
  let gap = 0;
  let overflowWidth = 44;

  function plan() {
    const el = strip();
    if (!el) {
      return;
    }
    const available = el.clientWidth - Math.max(0, reserved.value);
    collapsedCount.value = planCollapse({
      bandWidths,
      bandOrder,
      gap,
      available,
      overflowWidth,
      previousCollapsed: collapsedCount.value,
      hysteresis,
    });
  }

  async function measure() {
    const el = strip();
    // Nothing to measure without a laid-out strip: in jsdom (unit tests) every
    // offsetWidth is 0, so skip the measuring re-render churn and the timers
    // entirely rather than flapping the `measuring` flag on every mount.
    if (!el || el.clientWidth === 0) {
      return;
    }
    measuring.value = true;
    await nextTick();
    const styles = getComputedStyle(el);
    gap = Number.parseFloat(styles.columnGap) || 0;
    for (const band of bandOrder) {
      const bandEl = el.querySelector(`[data-band="${band}"]`);
      bandWidths[band] = bandEl instanceof HTMLElement ? bandEl.offsetWidth : 0;
    }
    const overflowEl = el.querySelector('[data-testid="toolbar-overflow-button"]');
    if (overflowEl instanceof HTMLElement && overflowEl.offsetWidth > 0) {
      overflowWidth = overflowEl.offsetWidth + gap;
    }
    measuring.value = false;
    await nextTick();
    plan();
  }

  let observer: ResizeObserver | null = null;
  const timers: number[] = [];

  onMounted(() => {
    void measure();
    const el = strip();
    if (el && typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver(() => plan());
      observer.observe(el);
    }
    // The UI font can settle in after first paint, changing band widths; the
    // ResizeObserver never re-fires for that (the strip clips its own
    // overflow), so re-measure a few times and once fonts are ready.
    for (const delay of [250, 700, 1400]) {
      timers.push(window.setTimeout(() => void measure(), delay));
    }
    const fonts = (document as Document & { fonts?: { ready?: Promise<unknown> } }).fonts;
    if (fonts?.ready) {
      void fonts.ready.then(() => measure());
    }
  });

  onBeforeUnmount(() => {
    observer?.disconnect();
    for (const timer of timers) {
      clearTimeout(timer);
    }
  });

  // Opening/closing a pane changes the reserved width, not the strip's own
  // width, so re-plan explicitly (AC-15).
  watch(reserved, () => plan());

  function isCollapsed(band: number): boolean {
    return collapsedBands(collapsedCount.value).has(band);
  }

  return { collapsedCount, measuring, isCollapsed, remeasure: measure, plan };
}
