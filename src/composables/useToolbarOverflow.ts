import { nextTick, onBeforeUnmount, onMounted, ref, watch, type Ref } from "vue";
import { collapsedBands, fitsWholly, planCollapse } from "../lib/toolbarOverflow";

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
  /** Whether the user wants control captions at all (the Toolbar Redesign's
   * "icons + labels" / "icons only" preference). When false the strip never
   * measures or shows the captioned geometry. */
  captionsWanted?: () => boolean;
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
  captionsWanted = () => true,
}: UseToolbarOverflowOptions) {
  const collapsedCount = ref(0);
  // Whether the strip is currently showing its control captions. The redesign
  // degrades captions BEFORE it collapses a band: losing the word under an
  // icon keeps every command in the strip, whereas collapsing a band hides
  // commands behind a menu. So the strip is measured twice — once captioned,
  // once not — and the captioned geometry is used only when it fits whole.
  const labelsOn = ref(false);
  // True while the strip force-expands every band so their natural widths can
  // be measured; DocumentToolbar disables the collapse teleports meanwhile.
  const measuring = ref(false);

  // Two measured geometries of the same bands: with captions and without.
  const captionedWidths: Record<number, number> = {};
  const bandWidths: Record<number, number> = {};
  // While measuring, this forces the captions on for the first pass and off
  // for the second, independently of the `labelsOn` the plan will settle on.
  // DocumentToolbar reads it so the two geometries are really rendered.
  const measuringCaptions = ref(true);
  // Until the first measure pass lands, captionedWidths is empty and would
  // read as "everything fits" — turning captions on for a frame and then off
  // again on a narrow strip. Hold them off until there is something to
  // compare.
  let measured = false;
  let gap = 0;
  let overflowWidth = 44;
  let captionedOverflowWidth = 44;

  function plan(fresh = false) {
    const el = strip();
    if (!el) {
      return;
    }
    const available = el.clientWidth - Math.max(0, reserved.value);
    // Step one: captions, but only if every band still fits with them on.
    labelsOn.value =
      measured &&
      captionsWanted() &&
      fitsWholly({
        bandWidths: captionedWidths,
        bandOrder,
        gap,
        available,
        overflowWidth: captionedOverflowWidth,
      });
    // Step two: collapse bands against whichever geometry is now rendered.
    collapsedCount.value = planCollapse({
      bandWidths: labelsOn.value ? captionedWidths : bandWidths,
      bandOrder,
      gap,
      available,
      overflowWidth: labelsOn.value ? captionedOverflowWidth : overflowWidth,
      // A fresh plan follows a re-measure (widths changed, not the available
      // width), so it takes the exact minimal fit — hysteresis exists to damp
      // resize flapping, and applying it here would lock in a pre-font
      // over-collapse that later measurements should undo (AC-36).
      previousCollapsed: fresh ? 0 : collapsedCount.value,
      hysteresis,
    });
  }

  function readInto(el: HTMLElement, into: Record<number, number>) {
    for (const band of bandOrder) {
      const bandEl = el.querySelector(`[data-band="${band}"]`);
      into[band] = bandEl instanceof HTMLElement ? bandEl.offsetWidth : 0;
    }
  }

  /** The trailing overflow button's footprint including its margins and the
   * gap before it, or null while it is hidden (nothing is collapsed yet, so
   * the browser reports zero and the previous measurement stands). */
  function readOverflowWidth(el: HTMLElement): number | null {
    const overflowEl = el.querySelector('[data-testid="toolbar-overflow-button"]');
    if (!(overflowEl instanceof HTMLElement) || overflowEl.offsetWidth === 0) {
      return null;
    }
    const ovfStyles = getComputedStyle(overflowEl);
    const margins =
      (Number.parseFloat(ovfStyles.marginInlineStart) || 0) +
      (Number.parseFloat(ovfStyles.marginInlineEnd) || 0);
    return overflowEl.offsetWidth + margins + gap;
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
    // Pass one: every band expanded AND captioned.
    measuringCaptions.value = true;
    await nextTick();
    const styles = getComputedStyle(el);
    gap = Number.parseFloat(styles.columnGap) || 0;
    readInto(el, captionedWidths);
    captionedOverflowWidth = readOverflowWidth(el) ?? captionedOverflowWidth;
    // Pass two: the same bands with every caption hidden.
    measuringCaptions.value = false;
    await nextTick();
    readInto(el, bandWidths);
    overflowWidth = readOverflowWidth(el) ?? overflowWidth;
    measured = true;
    measuring.value = false;
    await nextTick();
    plan(true);
  }

  let observer: ResizeObserver | null = null;
  const timers: number[] = [];

  /** The first measurement, then the re-measures the UI font needs. */
  function start(): void {
    void measure();
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
  }

  onMounted(() => {
    const el = strip();
    if (!el) {
      return;
    }
    if (typeof ResizeObserver === "undefined") {
      // No layout (jsdom / unit tests): skip the settle timers and font hook
      // entirely — there is nothing to measure and the churn slows the suite.
      // measure() itself also bails on a zero-width strip.
      if (el.clientWidth !== 0) {
        start();
      }
      return;
    }
    // The first measurement waits for the observer's first callback, which
    // the browser delivers straight after the first layout: reading the
    // strip's width here, in the task that mounted the view, would force
    // that layout — the whole view's (SPEC-137: ~30 ms of a large
    // workbook's open) — inside this task. A strip with no width then (not
    // laid out) is left alone, as before.
    let started = false;
    observer = new ResizeObserver(() => {
      if (started) {
        plan();
        return;
      }
      started = true;
      if (el.clientWidth === 0) {
        observer?.disconnect();
        return;
      }
      start();
    });
    observer.observe(el);
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

  return {
    collapsedCount,
    measuring,
    measuringCaptions,
    labelsOn,
    isCollapsed,
    remeasure: measure,
    plan,
  };
}
