// One axis of the grid's viewport (SPEC-116, SPEC-117 "Geometry", SPEC-137
// "The grid's ceiling"), framework-free — useSheetViewport holds the state
// and runs these:
//
// - mapAxis / scrollFor: the scroll mapping. Above MAX_SCROLL_EXTENT_PX the
//   spacer stops growing and offsets map onto the sheet proportionally.
// - axisWindow: the items a window renders — the frozen items, then the
//   scrolled ones from `overscan` before the first visible to `overscan`
//   past the last, filling the window's capacity.
// - SPEC-137: a scrolled item's slot (the key its elements take) is its
//   index less the axis's ring origin, mod the window's capacity. A "far"
//   scroll step — one that moves the content further than the overscan
//   reaches on either axis — turns the ring by the rows (columns) it moved,
//   so the item at the scrolled area's edge keeps its slot and every element
//   stays in place, patched.
// - SPEC-137: while the grid is settling — from a far step until no scroll
//   event has arrived for SETTLE_MS (ScrollSettler) — it renders the
//   viewport alone. A far step lands past everything the overscan had drawn,
//   so drawing it only doubles the step's patch (~1,300 cell elements
//   instead of ~600 at 1280 × 720: a fling's every frame, a jump's one). A
//   window may be asked for its slots at a wider overscan than it renders
//   (`slotOverscan`); with the ring turning, the viewport's elements keep
//   their slots from frame to frame. Settled, the overscan comes back a
//   step a frame (OVERSCAN_RAMP), so its own elements mount a few at a
//   time rather than ~700 in one task.
import type { AxisWindow } from "./sheetRenderPlan";

export const MAX_SCROLL_EXTENT_PX = 15_000_000;

/** How long the grid settles after a far step's last scroll event before
 * its overscan is drawn again (see the module header). */
export const SETTLE_MS = 150;

/** The overscan drawn on each frame once the grid has settled, as a
 * fraction of the full overscan. */
export const OVERSCAN_RAMP: readonly number[] = [0.25, 0.5, 1];

/** One axis's scroll mapping: unscaled content `total` px seen through
 * `view` unscaled px at zoom `z`, scrolled to `scroll` (scaled px). */
export interface AxisMapping {
  /** The spacer's length in scaled px. */
  spacerPx: number;
  /** The unscaled content offset of the scrolled area's leading edge. */
  offset: number;
  /** The scaled scroll range (spacer − view). */
  range: number;
  proportional: boolean;
  total: number;
  view: number;
}

export function mapAxis(total: number, view: number, scroll: number, z: number): AxisMapping {
  const scaled = total * z;
  const spacerPx = Math.min(scaled, MAX_SCROLL_EXTENT_PX);
  const range = Math.max(0, spacerPx - view * z);
  const proportional = scaled > spacerPx;
  let offset: number;
  if (proportional) {
    offset = range > 0 ? (Math.min(Math.max(scroll, 0), range) / range) * Math.max(0, total - view) : 0;
  } else {
    offset = Math.max(0, scroll) / z;
  }
  return { spacerPx, offset, range, proportional, total, view };
}

/** The scaled scroll position that shows content `offset` at the leading
 * edge — mapAxis's inverse. */
export function scrollFor(mapping: AxisMapping, offset: number, z: number): number {
  const maxOffset = Math.max(0, mapping.total - mapping.view);
  const clamped = Math.min(Math.max(offset, 0), maxOffset);
  if (!mapping.proportional) {
    return clamped * z;
  }
  return maxOffset > 0 ? (clamped / maxOffset) * mapping.range : 0;
}

/** The scroll offset that shows the item spanning content px begin..end
 * inside a scrolled area of `size` px past a frozen band of `frozenSize`,
 * currently scrolled to `offset` — unchanged when it is already in view. */
export function revealOffset(begin: number, end: number, frozenSize: number, size: number, offset: number): number {
  const from = begin - frozenSize;
  const to = end - frozenSize;
  if (from < offset) {
    return from;
  }
  if (to > offset + size) {
    return Math.min(from, to - size);
  }
  return offset;
}

/** A row or column in a window: its index, its leading edge relative to its
 * band (frozen items: the sheet origin; scrolled items: the scrolled area's
 * edge) and its size, unscaled px. */
export interface AxisItem {
  index: number;
  start: number;
  size: number;
  /** SPEC-137: the item's slot — a frozen item's index, a scrolled one's
   * index less the ring origin, mod the window's capacity. */
  slot: number;
}

/** What axisWindow reads of one axis. */
export interface AxisWindowInput {
  count: number;
  frozenSize: number;
  view: number;
  offset: number;
  extent: number;
  overscan: number;
  /** The overscan whose capacity the slots are taken modulo (≥ overscan). */
  slotOverscan: number;
  /** The slot ring's origin (see the module header). */
  ring: number;
  defaultSize: number;
  startOf: (i: number) => number;
  sizeOf: (i: number) => number;
  at: (pos: number) => number;
}

function frozenItems(count: number, view: number, startOf: (i: number) => number, sizeOf: (i: number) => number): AxisItem[] {
  const out: AxisItem[] = [];
  for (let i = 0; i < count; i += 1) {
    const begin = startOf(i);
    if (begin >= view) {
      break;
    }
    out.push({ index: i, start: begin, size: sizeOf(i), slot: i });
  }
  return out;
}

/** One axis's window at `overscan`, its slots taken modulo the capacity at
 * `slotOverscan` (see the module header). */
export function axisWindow(a: AxisWindowInput): AxisWindow {
  const top = a.at(a.frozenSize + a.offset);
  const bottom = a.at(a.offset + a.view);
  const capacityAt = (overscan: number): { from: number; capacity: number } => {
    const from = Math.max(a.count, top - overscan);
    const needed = Math.min(a.extent - 1, bottom + overscan) - from + 1;
    return { from, capacity: Math.max(1, needed, Math.ceil(Math.max(0, a.view - a.frozenSize) / a.defaultSize) + 1 + 2 * overscan) };
  };
  const { from, capacity } = capacityAt(a.overscan);
  const slots = a.slotOverscan > a.overscan ? Math.max(capacity, capacityAt(a.slotOverscan).capacity) : capacity;
  const last = Math.max(from - 1, Math.min(a.extent - 1, from + capacity - 1));
  const first = Math.max(a.count, Math.min(from, last - capacity + 1));
  const scrolled: AxisItem[] = [];
  for (let i = first; i <= last; i += 1) {
    const slot = (i - a.ring) % slots;
    scrolled.push({ index: i, start: a.startOf(i) - (a.offset + a.frozenSize), size: a.sizeOf(i), slot: slot < 0 ? slot + slots : slot });
  }
  return {
    frozen: frozenItems(a.count, a.view, a.startOf, a.sizeOf),
    scrolled,
    firstVisible: Math.max(a.count, top),
    lastVisible: Math.max(a.count, Math.min(a.extent - 1, bottom)),
    capacity: slots,
  };
}

/** How far apart OVERSCAN_RAMP's steps run: about a frame, each in a task
 * of its own (a frame callback could share the settle timer's frame). */
const RAMP_STEP_MS = 16;

/** How much of the overscan the grid draws (see the module header):
 * `step` is told of every scroll step and whether it was far; `onChange`
 * hears the fraction — 0 from a far step, then OVERSCAN_RAMP's, about a
 * frame apart, once SETTLE_MS have passed without a scroll event. */
export class ScrollSettler {
  /** The fraction of the overscan drawn now (1: all of it). */
  drawn = 1;
  private timer: ReturnType<typeof setTimeout> | undefined;
  /** Bumped by every step that restarts the wait, so a ramp it interrupts stops. */
  private generation = 0;

  constructor(private readonly onChange: (fraction: number) => void) {}

  /** A far step drops the overscan; every step while it is not all drawn
   * puts the ramp's start SETTLE_MS further off. */
  step(far: boolean): void {
    if (far) {
      this.set(0);
    }
    if (this.drawn < 1) {
      this.generation += 1;
      const generation = this.generation;
      clearTimeout(this.timer);
      this.timer = setTimeout(() => {
        this.timer = undefined;
        this.ramp(generation);
      }, SETTLE_MS);
    }
  }

  dispose(): void {
    clearTimeout(this.timer);
    this.generation += 1;
  }

  /** The next OVERSCAN_RAMP step up from what is drawn, then the next
   * RAMP_STEP_MS later. */
  private ramp(generation: number): void {
    const next = OVERSCAN_RAMP.find((fraction) => fraction > this.drawn);
    if (generation !== this.generation || next === undefined) {
      return;
    }
    this.set(next);
    this.timer = setTimeout(() => {
      this.timer = undefined;
      this.ramp(generation);
    }, RAMP_STEP_MS);
  }

  private set(fraction: number): void {
    if (fraction !== this.drawn) {
      this.drawn = fraction;
      this.onChange(fraction);
    }
  }
}
