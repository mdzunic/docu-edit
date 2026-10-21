// The Insert menu's drawings (SPEC-131 "Insert ▸ Image…", "Insert ▸ Shapes",
// "Paste"), composed by useSheetDrawings and run through its step runner, so
// each insert is ONE undo step ("Insert image", "Insert shape") and the new
// drawing lands selected.
//
// - Insert ▸ Image… clicks the hidden picker (SheetInsertSurfaces binds it); a
//   picked file — or a pasted image (useSheetClipboard) — is read by
//   drawingImport.ts (a refusal toasts and changes nothing: an unsupported
//   type, IMAGE_TOO_LARGE_MESSAGE, a broken file) and lands with its top-left
//   on the active cell at its natural size, scaled down to IMAGE_FIT_FRACTION
//   of the visible grid area (drawings.ts' insertImage), scrolled into view.
// - Insert ▸ Shapes opens the gallery popover; a pick arms the draw gesture
//   (`armed`, which useDrawingGestures reads) and clears the selection; the
//   gesture's box lands as the shape (insertShape) and disarms; Escape
//   anywhere disarms it first.
// Neither runs without `canInsert` (a grid sheet in an .xlsx or .ods
// workbook) or while a cell is being edited.
import { onScopeDispose, ref } from "vue";
import type { ShapeGeometry } from "../../lib/core/shapes";
import type { CellRef } from "../../lib/spreadsheet/address";
import { insertImage, insertShape, type DrawingCommandResult, type RectPx } from "../../lib/spreadsheet/drawings";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { prepareImage, type ImageLoader, type ImageSource } from "./drawingImport";
import type { DrawingGridHooks } from "./useSheetDrawings";

export interface DrawingInsertDeps {
  model: () => SpreadsheetModel | null;
  canInsert: () => boolean;
  busy: () => boolean;
  /** useSheetDrawings' step runner: one undo step, null when nothing ran. */
  run: (label: string, command: (model: SpreadsheetModel, sheetId: string) => DrawingCommandResult | null) => DrawingCommandResult | null;
  /** A drawing's rectangle on the sheet on show, once installed. */
  rectOf: (id: string) => RectPx | null;
  select: (id: string) => void;
  deselect: () => void;
  activeCell: () => CellRef;
  metrics: () => GridMetrics;
  grid: () => DrawingGridHooks | null;
  toast: (message: string) => void;
  loadImage?: ImageLoader;
}

/** The visible area an image is fitted to without a grid (an unmeasured
 * view's fallback desktop window, headers off). */
const FALLBACK_VISIBLE_PX = { width: 1224, height: 700 };

export function useDrawingInsert(deps: DrawingInsertDeps) {
  /** Insert ▸ Shapes' pick: the geometry the next grid press draws. */
  const armed = ref<ShapeGeometry | null>(null);
  const galleryOpen = ref(false);
  let picker: HTMLInputElement | null = null;
  const ready = (): boolean => deps.canInsert() && !deps.busy();

  function openGallery(): void {
    if (ready()) {
      galleryOpen.value = true;
    }
  }

  function closeGallery(): void {
    galleryOpen.value = false;
  }

  /** A gallery pick: the draw gesture armed, the selection cleared. */
  function pickShape(geometry: ShapeGeometry): void {
    galleryOpen.value = false;
    armed.value = geometry;
    deps.deselect();
  }

  function disarm(): void {
    armed.value = null;
  }

  /** The draw gesture's box (100%-zoom px): the shape lands, selected. */
  function drawShape(geometry: ShapeGeometry, rect: RectPx): void {
    armed.value = null;
    const result = deps.run("Insert shape", (m, sheetId) => insertShape(m, sheetId, geometry, rect, deps.metrics()));
    if (result?.drawingId) {
      deps.select(result.drawingId);
    }
  }

  function bindImagePicker(element: HTMLInputElement | null): void {
    picker = element;
  }

  /** Insert ▸ Image…: the file picker (its change inserts). */
  function pickImage(): void {
    if (ready()) {
      picker?.click();
    }
  }

  /** A picked or pasted image inserted at the active cell, selected; false
   * when refused (a toast) or nothing went in. */
  async function insertImageFile(file: ImageSource): Promise<boolean> {
    const m = deps.model();
    if (!m || !ready()) {
      return false;
    }
    const prepared = await prepareImage(file, m.format, deps.loadImage);
    if (typeof prepared === "string") {
      deps.toast(prepared);
      return false;
    }
    if (deps.model() !== m) {
      return false;
    }
    const at = deps.activeCell();
    const visiblePx = deps.grid()?.visiblePx() ?? FALLBACK_VISIBLE_PX;
    const result = deps.run("Insert image", (current, sheetId) => insertImage(current, sheetId, { ...prepared, at, visiblePx }, deps.metrics()));
    const rect = result?.drawingId ? deps.rectOf(result.drawingId) : null;
    if (!result?.drawingId || !rect) {
      return false;
    }
    deps.grid()?.reveal(rect);
    deps.select(result.drawingId);
    return true;
  }

  /** Escape disarms an armed draw gesture wherever focus is. */
  function onWindowKeydown(event: KeyboardEvent): void {
    if (event.key === "Escape" && armed.value !== null) {
      armed.value = null;
      event.preventDefault();
      event.stopPropagation();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onWindowKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onWindowKeydown, true);
    }
  });

  return { armed, galleryOpen, openGallery, closeGallery, pickShape, disarm, drawShape, bindImagePicker, pickImage, insertImageFile };
}
