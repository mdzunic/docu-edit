// The sheet drawing layer's view-level state (SPEC-131 "The drawing layer"):
// which drawing is selected, its keyboard, every drawing command, the alt-text
// dialog, the drawing context menu, Insert ▸ Image… and Insert ▸ Shapes, and
// the Arrange menu's actions. The grid on show (useDrawingGestures, one per
// SheetGrid) paints the layer from `shown`, runs the pointer gestures and lends
// this its focus, visible area and scrolling through `bindGrid`.
//
// - Selection is one drawing on the sheet on show (a sheet switch, an undo
//   that removes it or hiding its rows drops it). A press or focus selects it
//   and focuses its element; Escape — or a press on a cell (the grid's
//   gesture) — deselects and returns focus to the grid's focus holder with the
//   cell selection unchanged. Tab/Shift+Tab on a drawing cycle the selection
//   through the painted drawings in z-order, wrapping, revealing each.
// - Keys on a selected drawing: the arrows nudge it NUDGE_PX (Shift:
//   NUDGE_LARGE_PX) 100%-zoom px, clamped to the grid's top-left; Delete and
//   Backspace delete it. Modified keys pass (Ctrl/Cmd+Z undoes).
// - Every command is drawings.ts' and ONE SPEC-121 undo step on the sheet:
//   view state synced, `record(label)`, applyDrawingResult, `commit()` —
//   written through by the edit pipeline. A command that would change nothing
//   records nothing. Labels: "Insert image", "Insert shape", "Move {noun}",
//   "Resize {noun}", "Delete {noun}" (drawingNoun), Z_ORDER_LABELS' four and
//   "Alt text". Nothing runs while a cell is being edited (`busy`).
// - Insert ▸ Image… and Insert ▸ Shapes are useDrawingInsert's, through the
//   same step runner; neither is offered on a CSV or TSV workbook.
import { computed, nextTick, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { LIBRARY_IMAGE_MIME_TYPES, mediaDataUri, type MediaPart } from "../../lib/core/media";
import { DEFAULT_THEME } from "../../lib/core/theme";
import { applyDrawingResult, clampToGrid, deleteDrawing, drawingNoun, drawingRectPx, NUDGE_LARGE_PX, NUDGE_PX } from "../../lib/spreadsheet/drawings";
import { reorderDrawing, setDrawingAltText, setDrawingRect, visibleDrawings, Z_ORDER_LABELS, type DrawingCommandResult, type RectPx, type ZOrderMove } from "../../lib/spreadsheet/drawings";
import { isCsvFormat, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { ImageLoader } from "./drawingImport";
import type { ShownDrawing } from "./sheetDrawingLayout";
import { useDrawingInsert } from "./useDrawingInsert";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface SheetDrawingsDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  /** The active cell an image lands on. */
  selection: Pick<SheetSelectionController, "selection">;
  sheetMetrics: Pick<SheetMetrics, "metrics">;
  /** Syncs the view state into the model before a step is recorded. */
  syncInto?: (model: SpreadsheetModel) => void;
  /** Whether a cell is being edited: drawings are inert meanwhile. */
  busy?: () => boolean;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** drawingImport's Image fallback — default: the browser's. */
  loadImage?: ImageLoader;
}

/** What the grid on show lends the drawings (useDrawingGestures binds it). */
export interface DrawingGridHooks {
  /** Focus to the grid's focus holder (the in-cell editor). */
  focusGrid(): void;
  focusDrawing(id: string): void;
  /** The cell area's size, 100%-zoom px. */
  visiblePx(): { width: number; height: number };
  /** Scrolls a rectangle's top-left cell into view. */
  reveal(rect: RectPx): void;
}

/** The drawing context menu's and the Arrange menu's actions. */
export type DrawingAction = ZOrderMove | "alt-text" | "delete";

const sameRect = (a: RectPx, b: RectPx): boolean =>
  Math.abs(a.left - b.left) < 1e-6 && Math.abs(a.top - b.top) < 1e-6 && Math.abs(a.width - b.width) < 1e-6 && Math.abs(a.height - b.height) < 1e-6;

export function useSheetDrawings(deps: SheetDrawingsDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  /** The selected drawing, with the sheet it is on. */
  const picked = shallowRef<{ sheetId: string; id: string } | null>(null);
  const contextMenu = shallowRef<{ x: number; y: number } | null>(null);
  /** The open alt-text dialog: the drawing, its alt text and name. */
  const altText = shallowRef<{ id: string; initial: string; name: string } | null>(null);
  let grid: DrawingGridHooks | null = null;

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  const metrics = () => deps.sheetMetrics.metrics.value;
  const busy = (): boolean => deps.busy?.() === true;

  /** Whether the workbook can hold drawings: a grid sheet, not CSV/TSV. */
  const canInsert = computed(() => {
    const m = model();
    return m !== null && gridSheet() !== null && !isCsvFormat(m.format);
  });

  /** The drawings that paint on the sheet on show, back to front. */
  const shown = computed((): ShownDrawing[] => {
    const sheet = gridSheet();
    if (!sheet?.drawings?.length) {
      return [];
    }
    const m = metrics();
    return visibleDrawings(sheet, m).map((drawing) => ({ drawing, rect: drawingRectPx(drawing.anchor, m) }));
  });
  const entryOf = (id: string | null): ShownDrawing | null => shown.value.find((entry) => entry.drawing.id === id) ?? null;

  const selectedId = computed((): string | null => {
    const p = picked.value;
    const sheet = gridSheet();
    return p && sheet && p.sheetId === sheet.id && entryOf(p.id) ? p.id : null;
  });
  const selected = computed(() => entryOf(selectedId.value));
  const theme = computed(() => model()?.theme ?? DEFAULT_THEME);

  /** The selection's z-order position (what the menus enable). */
  const order = computed(() => {
    const drawings = gridSheet()?.drawings ?? [];
    const index = selectedId.value === null ? -1 : drawings.findIndex((drawing) => drawing.id === selectedId.value);
    return { selected: index >= 0, atFront: index >= 0 && index === drawings.length - 1, atBack: index === 0 };
  });

  // An image's data URI, once per media part; "" paints the placeholder.
  const sources = new WeakMap<MediaPart, string>();
  const mediaById = computed(() => new Map((model()?.media ?? []).map((part) => [part.mediaId, part])));
  function imageSrc(mediaId: string): string {
    const part = mediaById.value.get(mediaId);
    if (!part || !(LIBRARY_IMAGE_MIME_TYPES as readonly string[]).includes(part.contentType)) {
      return "";
    }
    let src = sources.get(part);
    if (src === undefined) {
      src = mediaDataUri(part);
      sources.set(part, src);
    }
    return src;
  }

  /** Selects a drawing on the sheet on show (and focuses its element). */
  function select(id: string, focus = true): void {
    const sheet = gridSheet();
    if (!sheet || busy()) {
      return;
    }
    if (picked.value?.sheetId !== sheet.id || picked.value.id !== id) {
      picked.value = { sheetId: sheet.id, id };
    }
    if (focus) {
      void nextTick(() => grid?.focusDrawing(id));
    }
  }

  /** Clears the selection; focus back to the grid unless `focusGrid` is false. */
  function deselect(focusGrid = true): void {
    picked.value = null;
    contextMenu.value = null;
    if (focusGrid) {
      grid?.focusGrid();
    }
  }

  /** One command as ONE undo step on the sheet on show; null: nothing ran. */
  function run(label: string, command: (m: SpreadsheetModel, sheetId: string) => DrawingCommandResult | null): DrawingCommandResult | null {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || busy()) {
      return null;
    }
    deps.syncInto?.(m);
    const result = command(m, sheet.id);
    if (!result) {
      return null;
    }
    deps.undo.record(label, { sheetId: sheet.id });
    applyDrawingResult(m, result);
    deps.undo.commit();
    return result;
  }

  /** A move or resize to `rect` (clamped to the grid's top-left). */
  function setRect(id: string, rect: RectPx, verb: "Move" | "Resize"): void {
    const entry = entryOf(id);
    const next = clampToGrid(rect);
    if (!entry || sameRect(next, entry.rect)) {
      return;
    }
    if (run(`${verb} ${drawingNoun(entry.drawing)}`, (m, sheetId) => setDrawingRect(m, sheetId, id, next, metrics()))) {
      select(id);
    }
  }

  function nudge(id: string, dx: number, dy: number): void {
    const entry = entryOf(id);
    if (entry) {
      setRect(id, { ...entry.rect, left: entry.rect.left + dx, top: entry.rect.top + dy }, "Move");
    }
  }

  function remove(id: string | null = selectedId.value): void {
    const entry = entryOf(id);
    if (entry && run(`Delete ${drawingNoun(entry.drawing)}`, (m, sheetId) => deleteDrawing(m, sheetId, entry.drawing.id))) {
      deselect();
    }
  }

  function reorder(move: ZOrderMove, id: string | null = selectedId.value): void {
    if (id !== null && entryOf(id)) {
      run(Z_ORDER_LABELS[move], (m, sheetId) => reorderDrawing(m, sheetId, id, move));
      select(id);
    }
  }

  /** Tab/Shift+Tab: the next/previous painted drawing in z-order, wrapping. */
  function cycle(delta: 1 | -1): void {
    const list = shown.value;
    if (list.length === 0) {
      return;
    }
    const index = list.findIndex((entry) => entry.drawing.id === selectedId.value);
    const from = index < 0 ? (delta > 0 ? -1 : 0) : index;
    const next = list[(from + delta + list.length) % list.length]!;
    grid?.reveal(next.rect);
    select(next.drawing.id);
  }

  const ARROWS: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };

  /** A key on a drawing's element (see the module header). */
  function onDrawingKeydown(id: string, event: KeyboardEvent): void {
    if (event.ctrlKey || event.metaKey || event.altKey || busy()) {
      return;
    }
    if (selectedId.value !== id) {
      select(id, false);
    }
    const arrow = ARROWS[event.key];
    const step = event.shiftKey ? NUDGE_LARGE_PX : NUDGE_PX;
    if (arrow) {
      nudge(id, arrow[0] * step, arrow[1] * step);
    } else if (event.key === "Delete" || event.key === "Backspace") {
      remove(id);
    } else if (event.key === "Escape") {
      deselect();
    } else if (event.key === "Tab") {
      cycle(event.shiftKey ? -1 : 1);
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
  }

  // --- alt text and the context menu ----------------------------------------------------
  function openAltText(id: string | null = selectedId.value): void {
    const entry = entryOf(id);
    if (entry && !busy()) {
      contextMenu.value = null;
      altText.value = { id: entry.drawing.id, initial: entry.drawing.altText ?? "", name: entry.drawing.name };
    }
  }

  /** The dialog's OK: the alt text written ("" removes it), one step. */
  function confirmAltText(text: string): void {
    const open = altText.value;
    altText.value = null;
    if (open) {
      run("Alt text", (m, sheetId) => setDrawingAltText(m, sheetId, open.id, text));
      select(open.id);
    }
  }

  function cancelAltText(): void {
    const open = altText.value;
    altText.value = null;
    if (open) {
      select(open.id);
    }
  }

  /** A right-click (or the Menu key) on a drawing: select it, open the menu. */
  function openContextMenu(id: string, x: number, y: number): void {
    if (!busy() && entryOf(id)) {
      select(id, false);
      contextMenu.value = { x, y };
    }
  }

  function closeContextMenu(): void {
    contextMenu.value = null;
  }

  /** The context menu's and the Arrange menu's actions on the selection. */
  function runAction(action: DrawingAction): void {
    contextMenu.value = null;
    if (action === "alt-text") {
      openAltText();
    } else if (action === "delete") {
      remove();
    } else {
      reorder(action);
    }
  }

  /** The grid on show lends its hooks; the returned function takes them back. */
  function bindGrid(hooks: DrawingGridHooks): () => void {
    grid = hooks;
    return () => {
      if (grid === hooks) {
        grid = null;
      }
    };
  }

  const insert = useDrawingInsert({
    model,
    canInsert: () => canInsert.value,
    busy,
    run,
    rectOf: (id) => entryOf(id)?.rect ?? null,
    select: (id) => select(id),
    deselect: () => deselect(),
    activeCell: () => deps.selection.selection.value.active,
    metrics,
    grid: () => grid,
    toast,
    loadImage: deps.loadImage,
  });

  // prettier-ignore
  return {
    shown, selectedId, selected, order, theme, canInsert, contextMenu, altText, busy, imageSrc,
    select, deselect, setRect, nudge, remove, reorder, cycle, onDrawingKeydown, bindGrid,
    openAltText, confirmAltText, cancelAltText, openContextMenu, closeContextMenu, runAction,
    ...insert,
  };
}

export type SheetDrawings = ReturnType<typeof useSheetDrawings>;

/** Narrows a component's runtime-typed `drawings` prop (the .vue parser bars
 * casts). */
export function asSheetDrawings(value: unknown): SheetDrawings {
  return value as SheetDrawings;
}
