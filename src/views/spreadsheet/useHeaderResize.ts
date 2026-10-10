// Resizing rows and columns at their header boundaries (SPEC-124 "Sizes",
// "Autofit"). The grid's pointer handler (useSheetPointer) asks this
// composable, in unscaled content px, whether a header press or hover lies
// within RESIZE_HIT_PX screen px of a boundary — the trailing edge of the
// row or column `boundaryAt` returns:
//
// - Hovering there shows the `col-resize` / `row-resize` cursor. A press
//   starts a drag: the guide line (`sheet-resize-guide`) follows the
//   boundary and the tooltip names the size — "Width: 8.43 (64 pixels)"
//   (Excel characters through SPEC-117's MDW formula) or "Height: 15.00 (20
//   pixels)". Escape cancels; the release commits ONE undo step through
//   useStructure ("Resize rows"/"Resize columns"); a size under 4 px hides
//   the row or column, keeping its size for unhide.
// - When the dragged header belongs to a whole-row or whole-column
//   selection, every selected row or column takes the new size; otherwise
//   only the dragged one.
// - A double-click on a boundary autofits ("Autofit") — every selected row
//   or column when the boundary belongs to the selection.
import { computed, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { Band, BandAxis } from "../../lib/spreadsheet/bands";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { fontMaxDigitWidth } from "../../lib/spreadsheet/styles";
import { bandSelected, selectionBands } from "./headerMenu";
import { resizeTooltip } from "./sheetSizes";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetZoom } from "./useSheetZoom";
import type { Structure } from "./useStructure";

/** How near (screen px) a boundary a header press or hover must be. */
export const RESIZE_HIT_PX = 3;

export interface HeaderResizeDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  structure: Pick<Structure, "resizeBands" | "autofitBands">;
  selection: Pick<SheetSelectionController, "selection">;
  sheetMetrics: Pick<SheetMetrics, "metrics">;
  zoom: Pick<SheetZoom, "percent">;
}

/** A drag in progress, in unscaled px: the item's leading edge (content),
 * the pointer where it pressed (content), the item's size then and now, and
 * the guide's position in the viewport. */
interface ResizeDrag {
  axis: BandAxis;
  index: number;
  edge: number;
  start: number;
  size0: number;
  size: number;
  guide: number;
  across: number;
}

export function useHeaderResize(deps: HeaderResizeDeps) {
  const drag = shallowRef<ResizeDrag | null>(null);
  const hover = ref<BandAxis | null>(null);

  const metrics = () => deps.sheetMetrics.metrics.value;
  const zoom = (): number => Math.max(0.1, deps.zoom.percent.value / 100);
  const leadingEdge = (axis: BandAxis, index: number): number => (axis === "col" ? metrics().colLeft(index) : metrics().rowTop(index));
  const sizeOf = (axis: BandAxis, index: number): number => (axis === "col" ? metrics().colWidth(index) : metrics().rowHeight(index));

  /** The row or column whose trailing boundary lies within RESIZE_HIT_PX
   * (screen px) of content position `pos`, or null. */
  function boundaryAt(axis: BandAxis, pos: number): number | null {
    const tolerance = RESIZE_HIT_PX / zoom();
    const at = axis === "col" ? metrics().colAt(Math.max(0, pos)) : metrics().rowAt(Math.max(0, pos));
    const leading = leadingEdge(axis, at);
    if (leading + sizeOf(axis, at) - pos <= tolerance) {
      return at;
    }
    return at > 0 && pos - leading <= tolerance ? at - 1 : null;
  }

  /** Every row or column a resize or autofit of `index` reaches. */
  function bandsFor(axis: BandAxis, index: number): Band[] {
    const selection = deps.selection.selection.value;
    return bandSelected(selection, axis, index) ? selectionBands(selection, axis) : [{ start: index, end: index }];
  }

  /** A press at `pos` (content px) on boundary `index`; `view` is the
   * pointer's viewport px on the axis, `across` on the other axis. */
  function begin(axis: BandAxis, index: number, pos: number, view: number, across: number): void {
    const size = sizeOf(axis, index);
    const edge = leadingEdge(axis, index);
    drag.value = { axis, index, edge, start: pos, size0: size, size, guide: view - (pos - edge - size), across };
  }

  function move(pos: number, view: number, across: number): void {
    const current = drag.value;
    if (!current) {
      return;
    }
    const size = Math.max(0, current.size0 + pos - current.start);
    // The guide sits on the dragged boundary: under the pointer, or at the
    // leading edge once the size reaches 0.
    drag.value = { ...current, size, guide: view - (pos - (current.edge + size)), across };
  }

  /** The release: one undo step when the size changed. */
  function finish(): void {
    const current = drag.value;
    drag.value = null;
    if (current && Math.round(current.size) !== current.size0) {
      deps.structure.resizeBands(current.axis, bandsFor(current.axis, current.index), Math.round(current.size));
    }
  }

  function cancel(): void {
    drag.value = null;
  }

  /** A double-click on boundary `index`: autofit. */
  function autofit(axis: BandAxis, index: number): void {
    deps.structure.autofitBands(axis, bandsFor(axis, index));
  }

  const mdw = (): number => {
    const m = deps.book.workbook.value;
    return fontMaxDigitWidth(m?.styles[0]?.font, m?.theme ?? DEFAULT_THEME);
  };

  /** The guide line (viewport px along the axis), while dragging. */
  const guide = computed(() => (drag.value ? { axis: drag.value.axis, at: drag.value.guide } : null));
  /** The size tooltip, beside the pointer, while dragging. */
  const tooltip = computed(() => {
    const current = drag.value;
    if (!current) {
      return null;
    }
    const text = resizeTooltip(current.axis, current.size, mdw());
    return current.axis === "col" ? { text, x: current.guide + 8, y: current.across + 12 } : { text, x: current.across + 12, y: current.guide + 8 };
  });
  /** The grid's cursor: resizing, or hovering a boundary. */
  const cursor = computed((): string | undefined => {
    const axis = drag.value?.axis ?? hover.value;
    return axis === "col" ? "col-resize" : axis === "row" ? "row-resize" : undefined;
  });

  function onKeydown(event: KeyboardEvent): void {
    if (drag.value && event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      cancel();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    boundaryAt,
    begin,
    move,
    finish,
    cancel,
    autofit,
    /** Whether a drag is in progress. */
    active: computed(() => drag.value !== null),
    setHover(axis: BandAxis | null): void {
      if (hover.value !== axis) {
        hover.value = axis;
      }
    },
    guide,
    tooltip,
    cursor,
    onKeydown,
  };
}

export type HeaderResize = ReturnType<typeof useHeaderResize>;
