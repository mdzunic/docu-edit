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
// - SPEC-137: while the grid is settling — through a fling: two far steps
//   within SETTLE_MS of each other, until no scroll event has arrived for
//   SETTLE_MS (ScrollSettler) — it renders the viewport alone. Each far step
//   lands past everything the overscan had drawn, so drawing it only doubles
//   every frame's patch (~1,300 cell elements instead of ~600 at 1280 × 720).
//   A window may be asked for its slots at a wider overscan than it renders
//   (`slotOverscan`); with the ring turning, the viewport's elements keep
//   their slots from frame to frame, and when the overscan comes back only
//   its own elements mount. A single jump (a name-box move, Ctrl+End) draws
//   the full window.
import type { AxisWindow } from "./sheetRenderPlan";

export const MAX_SCROLL_EXTENT_PX = 15_000_000;

/** How long the grid settles after a fling's last scroll event before its
 * overscan is drawn again (see the module header). */
export const SETTLE_MS = 150;

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

/** Tells a fling from a jump (see the module header): `step` is told of
 * every scroll step and whether it was far; `onChange` hears when the grid
 * starts and stops settling. */
export class ScrollSettler {
  settling = false;
  private timer: ReturnType<typeof setTimeout> | undefined;
  /** When the last far step came (performance.now()); -Infinity: none yet. */
  private lastFar = Number.NEGATIVE_INFINITY;

  constructor(private readonly onChange: (settling: boolean) => void) {}

  /** A second far step within SETTLE_MS of the last starts a fling; every
   * step while settling puts its end SETTLE_MS further off. */
  step(far: boolean): void {
    if (far) {
      const now = performance.now();
      if (now - this.lastFar <= SETTLE_MS && !this.settling) {
        this.settling = true;
        this.onChange(true);
      }
      this.lastFar = now;
    }
    if (this.settling) {
      clearTimeout(this.timer);
      this.timer = setTimeout(() => {
        this.timer = undefined;
        this.settling = false;
        this.onChange(false);
      }, SETTLE_MS);
    }
  }

  dispose(): void {
    clearTimeout(this.timer);
  }
}
