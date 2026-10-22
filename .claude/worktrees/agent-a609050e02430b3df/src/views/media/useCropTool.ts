// The crop tool (SPEC-112): the aspect presets, the crop rectangle's
// handles — focusable, moved 1 px by an arrow key and 10 px with Shift, or
// dragged — and the rectangle's announcement as "x, y, width × height".
// The rectangle lives in source pixels in the straightened frame and is
// clamped into the straighten's inscribed area (core/imageEdits.ts); the
// crop tool's preview shows exactly that area, so the overlay maps the
// rectangle onto the preview with one offset and one scale. A drag is one
// undo entry ("Crop"); so is each key press.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { removeKind, upsertEntry } from "../../lib/core/imageEditHistory";
import {
  cropAnnouncement,
  cropAspectRatio,
  cropForAspect,
  moveCropHandle,
  type CropAspect,
  type CropHandle,
  type PixelRect,
} from "../../lib/core/imageEdits";
import type { EditSession } from "./editSession";

/** One arrow press moves a handle this far; with Shift, ten times as far. */
export const CROP_STEP_PX = 1;
export const CROP_SHIFT_STEP_PX = 10;

const HANDLE_NAMES: Record<CropHandle, string> = {
  move: "Crop area",
  n: "Top crop edge",
  s: "Bottom crop edge",
  e: "Right crop edge",
  w: "Left crop edge",
  ne: "Top-right crop corner",
  nw: "Top-left crop corner",
  se: "Bottom-right crop corner",
  sw: "Bottom-left crop corner",
};

export interface CropTool {
  aspect: Ref<CropAspect>;
  /** The effective crop rectangle (straightened frame, source pixels). */
  rect: ComputedRef<PixelRect | null>;
  /** The area the rectangle is clamped into — what the crop preview shows. */
  bounds: ComputedRef<PixelRect | null>;
  announcement: ComputedRef<string>;
  setAspect: (aspect: CropAspect) => void;
  handleName: (handle: CropHandle) => string;
  /** Arrow keys on a focused handle; true when handled. */
  handleKey: (handle: CropHandle, event: KeyboardEvent) => boolean;
  startDrag: (handle: CropHandle) => void;
  /** Moves the dragged handle by (dx, dy) source pixels from where the drag
   * began. */
  dragBy: (dx: number, dy: number) => void;
  endDrag: () => void;
  reset: () => void;
}

export function useCropTool(session: EditSession): CropTool {
  const aspect = ref<CropAspect>("free");
  const rect = computed(() => session.geometry.value?.source ?? null);
  const bounds = computed(() => session.geometry.value?.inscribed ?? null);
  const ratio = computed(() => (session.naturalPx.value ? cropAspectRatio(aspect.value, session.naturalPx.value) : null));
  const announcement = computed(() => (rect.value ? cropAnnouncement(rect.value) : ""));

  function cropped(next: PixelRect): ReturnType<typeof upsertEntry> {
    return upsertEntry(session.history.value, { kind: "crop", ...next });
  }

  function setAspect(next: CropAspect): void {
    aspect.value = next;
    const area = bounds.value;
    if (next === "free" || !area) {
      return;
    }
    session.apply("Crop", cropped(cropForAspect(ratio.value, area)));
  }

  function moveBy(handle: CropHandle, from: PixelRect, dx: number, dy: number): PixelRect | null {
    const area = bounds.value;
    return area ? moveCropHandle(from, handle, dx, dy, ratio.value, area) : null;
  }

  function handleKey(handle: CropHandle, event: KeyboardEvent): boolean {
    const step = event.shiftKey ? CROP_SHIFT_STEP_PX : CROP_STEP_PX;
    const delta: Record<string, [number, number]> = {
      ArrowLeft: [-step, 0],
      ArrowRight: [step, 0],
      ArrowUp: [0, -step],
      ArrowDown: [0, step],
    };
    const move = delta[event.key];
    const from = rect.value;
    if (!move || !from || event.altKey || event.metaKey || event.ctrlKey) {
      return false;
    }
    const next = moveBy(handle, from, move[0], move[1]);
    if (next) {
      session.apply("Crop", cropped(next));
    }
    return true;
  }

  let drag: { handle: CropHandle; from: PixelRect } | null = null;

  function startDrag(handle: CropHandle): void {
    if (rect.value) {
      drag = { handle, from: rect.value };
      session.begin("Crop");
    }
  }

  function dragBy(dx: number, dy: number): void {
    if (!drag) return;
    const next = moveBy(drag.handle, drag.from, dx, dy);
    if (next) {
      session.update(cropped(next));
    }
  }

  function endDrag(): void {
    if (drag) {
      drag = null;
      session.end();
    }
  }

  return {
    aspect,
    rect,
    bounds,
    announcement,
    setAspect,
    handleName: (handle) => HANDLE_NAMES[handle],
    handleKey,
    startDrag,
    dragBy,
    endDrag,
    reset: () => session.apply("Reset crop", removeKind(session.history.value, "crop")),
  };
}
