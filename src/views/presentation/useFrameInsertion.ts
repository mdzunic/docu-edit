// Inserting new slide objects (SPEC-082): "Text box" and gallery shapes arm a
// draw gesture on the stage (drag draws the box, a click drops a default), and
// "Image…" opens a file picker. Every insert records one undo entry, selects
// the new frame, persists write-through, and (a new text box) opens straight
// into a text-editing session. The gesture arithmetic is the core drawing
// vocabulary (drawGestureBox); coordinate conversion is stageMath's.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { DRAW_DRAG_THRESHOLD_CM, drawGestureBox } from "../../lib/core/shapes";
import type { ShapeGeometry } from "../../lib/core/shapes";
import { decodeImageDimensions } from "../../lib/core/imageDimensions";
import type { MediaPart } from "../../lib/core/media";
import { bytesToBase64 } from "../../lib/core/packageParts";
import {
  DEFAULT_TEXT_BOX_SIZE_CM,
  insertImageFrame,
  insertImageRect,
  insertShapeFrame,
  insertTextBoxFrame,
  type FrameRect,
} from "../../lib/presentation/frames";
import type { FrameSelection } from "./useFrameSelection";
import { clientToCm } from "./stageMath";

/** The image formats "Insert → Image" accepts (AC-4). Distinct from the text
 * product's core set (which carries svg, not webp), so it is stated here. */
export const PRESENTATION_IMAGE_MIME_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp"] as const;
export const PRESENTATION_IMAGE_ACCEPT = PRESENTATION_IMAGE_MIME_TYPES.join(",");

/** The armed insertion tool: a free text box, or a catalogue shape. */
export type ArmedTool = { kind: "textbox" } | { kind: "shape"; geometry: ShapeGeometry };

export interface FrameInsertionDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  selection: FrameSelection;
  surface: () => HTMLElement | null;
  record: (label?: string) => void;
  commit: () => void;
  /** Opens the new text box into an editing session (AC-12). */
  startEditing: (frameId: string) => void;
  /** Reports a broken/unreadable image (AC-9). */
  reportError: (message: string) => void;
}

export function useFrameInsertion(deps: FrameInsertionDeps) {
  const armed = ref<ArmedTool | null>(null);
  /** The live draw preview rect (cm), for the stage overlay. */
  const preview = ref<FrameRect | null>(null);

  const isArmed = computed(() => armed.value !== null);

  function armTextBox(): void {
    armed.value = { kind: "textbox" };
  }
  function armShape(geometry: ShapeGeometry): void {
    armed.value = { kind: "shape", geometry };
  }
  function disarm(): void {
    armed.value = null;
    preview.value = null;
  }

  function pointerCm(event: PointerEvent): { xCm: number; yCm: number } | null {
    const rect = deps.surface()?.getBoundingClientRect();
    const model = deps.model.value;
    if (!rect || !model) return null;
    return clientToCm(event.clientX, event.clientY, rect, model.slideWidthCm, model.slideHeightCm);
  }

  /** The stage's pointerdown while a tool is armed: begins a draw gesture.
   * Returns true when it consumed the event (the caller then skips deselect). */
  function onSurfacePointerDown(event: PointerEvent): boolean {
    const tool = armed.value;
    const model = deps.model.value;
    const start = pointerCm(event);
    if (!tool || !model || !start) return false;
    event.stopPropagation();
    const geometry: ShapeGeometry = tool.kind === "shape" ? tool.geometry : "rect";

    const onMove = (move: PointerEvent): void => {
      const now = pointerCm(move);
      if (!now) return;
      const box = drawGestureBox(geometry, start, now, { shiftKey: move.shiftKey });
      preview.value = { xCm: box.leftCm, yCm: box.topCm, widthCm: box.widthCm, heightCm: box.heightCm };
    };
    const onUp = (up: PointerEvent): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const now = pointerCm(up) ?? start;
      commitDraw(tool, geometry, start, now, up.shiftKey);
      preview.value = null;
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    return true;
  }

  function commitDraw(
    tool: ArmedTool,
    geometry: ShapeGeometry,
    start: { xCm: number; yCm: number },
    pointer: { xCm: number; yCm: number },
    shiftKey: boolean,
  ): void {
    const model = deps.model.value;
    if (!model) return;
    const dragged =
      Math.abs(pointer.xCm - start.xCm) > DRAW_DRAG_THRESHOLD_CM ||
      Math.abs(pointer.yCm - start.yCm) > DRAW_DRAG_THRESHOLD_CM;
    let rect: FrameRect;
    if (tool.kind === "textbox" && !dragged) {
      // A click (no drag) with Text box armed drops the default 10×2.5cm box
      // at the pointer (AC-11).
      rect = { xCm: start.xCm, yCm: start.yCm, ...DEFAULT_TEXT_BOX_SIZE_CM };
    } else {
      const box = drawGestureBox(geometry, start, pointer, { shiftKey });
      rect = { xCm: box.leftCm, yCm: box.topCm, widthCm: box.widthCm, heightCm: box.heightCm };
    }
    deps.record(tool.kind === "textbox" ? "Insert text box" : "Insert shape");
    const id =
      tool.kind === "textbox"
        ? insertTextBoxFrame(model, deps.currentIndex.value, rect)
        : insertShapeFrame(model, deps.currentIndex.value, geometry, rect);
    deps.selection.select(id);
    deps.commit();
    disarm();
    if (tool.kind === "textbox" && id) {
      // Defer the editing session past the placing click's own focus handling:
      // clicking the (non-focusable) surface blurs the active element, so a
      // contenteditable focused synchronously here would lose focus the moment
      // the click completes (AC-12).
      setTimeout(() => deps.startEditing(id), 0);
    }
  }

  // --- image picker ---------------------------------------------------------

  async function insertImageFile(file: File): Promise<void> {
    const model = deps.model.value;
    if (!model) return;
    try {
      const buffer = new Uint8Array(await file.arrayBuffer());
      const contentType = file.type || "image/png";
      const natural = decodeImageDimensions(buffer, contentType);
      const media: MediaPart = { mediaId: "", contentType, data: bytesToBase64(buffer) };
      const rect = insertImageRect(model, natural);
      deps.record("Insert image");
      const id = insertImageFrame(model, deps.currentIndex.value, media, rect);
      if (!id) {
        deps.reportError("That image couldn't be inserted.");
        return;
      }
      deps.selection.select(id);
      deps.commit();
    } catch {
      deps.reportError("That image couldn't be read.");
    }
  }

  return {
    armed,
    isArmed,
    preview,
    armTextBox,
    armShape,
    disarm,
    onSurfacePointerDown,
    insertImageFile,
    accept: PRESENTATION_IMAGE_ACCEPT,
  };
}
