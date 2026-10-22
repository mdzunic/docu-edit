// Typed helpers for the sheet tab bar's components (SPEC-121): SheetTabs,
// SheetTabMenu and SheetUnhideDialog. Their .vue scripts are linted with
// espree, which bars TS-only syntax (casts, type predicates, generics), so
// the runtime props are narrowed here, and the drag-reorder geometry and the
// menu placement live here as plain, unit-testable functions.

/** One sheet tab: the sheet's index in the model, its id and name, and its
 * tab colour resolved to a CSS colour ("" for none — the view resolves theme
 * references, never the tab bar). */
export interface SheetTabItem {
  index: number;
  id: string;
  name: string;
  color: string;
}

/** Narrows SheetTabs' runtime-typed `tabs` prop. */
export function asSheetTabItems(value: unknown): SheetTabItem[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((entry) => {
    const tab = (entry ?? {}) as Partial<SheetTabItem>;
    return {
      index: typeof tab.index === "number" ? tab.index : -1,
      id: String(tab.id ?? ""),
      name: String(tab.name ?? ""),
      color: typeof tab.color === "string" ? tab.color : "",
    };
  });
}

/** One entry of the Unhide dialog: a hidden sheet's model index and name. */
export interface HiddenSheetEntry {
  index: number;
  name: string;
}

/** Narrows SheetUnhideDialog's runtime-typed `sheets` prop. */
export function asHiddenSheetEntries(value: unknown): HiddenSheetEntry[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((entry) => {
    const sheet = (entry ?? {}) as Partial<HiddenSheetEntry>;
    return { index: typeof sheet.index === "number" ? sheet.index : -1, name: String(sheet.name ?? "") };
  });
}

/** A coloured tab that is not on show is tinted with its colour; the one on
 * show keeps the selected look and carries the colour as a bottom bar. */
export function tabTint(color: string, active: boolean): Record<string, string> {
  return color && !active ? { background: `color-mix(in srgb, ${color} 25%, transparent)` } : {};
}

// ---- drag-reorder -----------------------------------------------------------

/** A pointer must travel further than this from its press before a press on
 * a tab becomes a drag (a click that wobbles is still a click). */
export const DRAG_THRESHOLD_PX = 4;

export function beyondDragThreshold(dx: number, dy: number): boolean {
  return Math.hypot(dx, dy) > DRAG_THRESHOLD_PX;
}

/** A tab's horizontal extent in viewport px. */
export interface TabSpan {
  left: number;
  right: number;
}

/** Every tab slot (a tab, or the rename input standing in for one) inside
 * `container`, in visible order — each carries `data-tab-position`. */
export function tabSlotElements(container: Element | null | undefined): HTMLElement[] {
  const nodes = container ? Array.from(container.querySelectorAll("[data-tab-position]")) : [];
  return nodes.filter((node): node is HTMLElement => node instanceof HTMLElement);
}

/** The tab with visible position `position` (never the rename input). */
export function tabElementAt(container: Element | null | undefined, position: number): HTMLElement | null {
  const node = container?.querySelector(`[role="tab"][data-tab-position="${position}"]`);
  return node instanceof HTMLElement ? node : null;
}

export function slotSpans(container: Element | null | undefined): TabSpan[] {
  return tabSlotElements(container).map((element) => {
    const rect = element.getBoundingClientRect();
    return { left: rect.left, right: rect.right };
  });
}

/** The gap nearest a pointer at viewport `x`: 0 is before the first tab,
 * `spans.length` after the last. Over a tab's left half the gap before it is
 * nearer, over its right half the gap after it. */
export function dropGapAt(x: number, spans: readonly TabSpan[]): number {
  const index = spans.findIndex((span) => x < (span.left + span.right) / 2);
  return index === -1 ? spans.length : index;
}

/** The viewport x of gap `gap` — midway between its two neighbours. */
export function gapX(gap: number, spans: readonly TabSpan[]): number {
  if (spans.length === 0) {
    return 0;
  }
  if (gap <= 0) {
    return spans[0].left;
  }
  if (gap >= spans.length) {
    return spans[spans.length - 1].right;
  }
  return (spans[gap - 1].right + spans[gap].left) / 2;
}

/** The `move` a drop of the tab at visible position `from` into gap `gap`
 * asks for — `index` the dragged sheet's model index, `before` the model
 * index of the tab it lands in front of, or -1 for after the last tab — or
 * null when the drop leaves the order as it was (either gap beside it). */
export function moveForDrop(
  tabs: readonly SheetTabItem[],
  from: number,
  gap: number,
): { index: number; before: number } | null {
  const tab = tabs[from];
  if (!tab || gap === from || gap === from + 1 || gap < 0 || gap > tabs.length) {
    return null;
  }
  return { index: tab.index, before: gap < tabs.length ? tabs[gap].index : -1 };
}

export interface ViewportBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export function insideRect(x: number, y: number, rect: ViewportBox): boolean {
  return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
}

// ---- menus ------------------------------------------------------------------

/** Where a `size` box opened at viewport `point` goes so that it stays on
 * screen: shifted left to fit, and flipped above the point when it would run
 * off the bottom (the tab bar sits at the foot of the view). */
export function clampToViewport(
  point: { x: number; y: number },
  size: { width: number; height: number },
  viewport: { width: number; height: number },
  margin = 8,
): { x: number; y: number } {
  const x = Math.max(margin, Math.min(point.x, viewport.width - size.width - margin));
  const fitsBelow = point.y + size.height <= viewport.height - margin;
  const y = fitsBelow
    ? Math.max(0, point.y)
    : Math.max(margin, Math.min(point.y - size.height, viewport.height - size.height - margin));
  return { x, y };
}

/** The enabled menu items inside a rendered menu, in DOM order — the ring
 * the arrow keys move through. */
export function enabledMenuItems(container: Element | null | undefined): HTMLElement[] {
  const nodes = container ? Array.from(container.querySelectorAll('[role="menuitem"]')) : [];
  return nodes.filter(
    (node): node is HTMLElement => node instanceof HTMLElement && !(node instanceof HTMLButtonElement && node.disabled),
  );
}
