// The assembly preview (SPEC-109): the sheet at fit-to-window, drawn by
// media/composeCanvas's composeToCanvas — the function the export calls, at
// a smaller pixel size — so what the user sees is what is written. Re-drawn
// whenever the plan, the sheet, the background or the window changes; a
// newer render cancels an older one at its next boundary.
//
// It is also where overlay items move by hand: a press on the canvas picks
// the topmost placed item under the pointer (selecting its tray entry), and
// in the overlay layout a drag moves it — Shift held keeps it on the axis it
// has moved along most — while the arrow keys nudge the selected item by
// 1 mm (Shift: 1 cm).
import { computed, onScopeDispose, ref, shallowRef, watch, type ComputedRef, type Ref } from "vue";
import { composeToCanvas, decodeSource, prepareSources } from "../../lib/media/composeCanvas";
import type { DecodedImage } from "../../lib/media/renderEdits";
import type { OverlayTransform, PlacedItem } from "../../lib/core/assemblyLayout";
import type { PixelSize } from "../../lib/core/imageEdits";
import type { MediaPart } from "../../lib/core/media";
import type { SheetGeometry } from "../../lib/core/sheetLayout";
import { NUDGE_CM, NUDGE_SHIFT_CM, type AssemblyController } from "./useAssembly";
import { composeInputs } from "./useAssemblyOutput";

/** The preview canvas's largest side, in device pixels. */
export const PREVIEW_MAX_EDGE_PX = 2048;

/** The width the preview takes before (or without) a measured container. */
export const PREVIEW_FALLBACK_WIDTH_PX = 480;

/** The sheet's on-screen size: as large as fits `box` (CSS px). */
export function previewDisplaySize(geometry: SheetGeometry, box: PixelSize): PixelSize {
  const ratio = geometry.heightCm / geometry.widthCm;
  if (box.widthPx <= 0 || box.heightPx <= 0) {
    return { widthPx: PREVIEW_FALLBACK_WIDTH_PX, heightPx: Math.round(PREVIEW_FALLBACK_WIDTH_PX * ratio) };
  }
  const scale = Math.min(box.widthPx / geometry.widthCm, box.heightPx / geometry.heightCm);
  return { widthPx: Math.max(1, Math.floor(geometry.widthCm * scale)), heightPx: Math.max(1, Math.floor(geometry.heightCm * scale)) };
}

/** The canvas's own pixel size for a display size at `pixelRatio`, its
 * longest side capped at PREVIEW_MAX_EDGE_PX. */
export function previewPixelSize(display: PixelSize, pixelRatio: number): PixelSize {
  const ratio = pixelRatio > 0 ? pixelRatio : 1;
  const longest = Math.max(display.widthPx, display.heightPx) * ratio;
  const scale = longest > PREVIEW_MAX_EDGE_PX ? (PREVIEW_MAX_EDGE_PX / longest) * ratio : ratio;
  return { widthPx: Math.max(1, Math.round(display.widthPx * scale)), heightPx: Math.max(1, Math.round(display.heightPx * scale)) };
}

/** The topmost placed item containing a sheet point (the last drawn wins). */
export function hitTest(placed: readonly PlacedItem[], xCm: number, yCm: number): PlacedItem | null {
  for (let i = placed.length - 1; i >= 0; i -= 1) {
    const item = placed[i];
    if (xCm >= item.xCm && xCm <= item.xCm + item.widthCm && yCm >= item.yCm && yCm <= item.yCm + item.heightCm) {
      return item;
    }
  }
  return null;
}

/** A drag's distance with Shift's axis lock: the smaller component zeroed. */
export function constrainDrag(dxCm: number, dyCm: number, lockAxis: boolean): { dxCm: number; dyCm: number } {
  if (!lockAxis) {
    return { dxCm, dyCm };
  }
  return Math.abs(dxCm) >= Math.abs(dyCm) ? { dxCm, dyCm: 0 } : { dxCm: 0, dyCm };
}

/** An arrow key's nudge in centimetres, or null for any other key. */
export function nudgeFor(key: string, shift: boolean): { dxCm: number; dyCm: number } | null {
  const step = shift ? NUDGE_SHIFT_CM : NUDGE_CM;
  const moves: Record<string, [number, number]> = {
    ArrowLeft: [-step, 0],
    ArrowRight: [step, 0],
    ArrowUp: [0, -step],
    ArrowDown: [0, step],
  };
  const move = moves[key];
  return move ? { dxCm: move[0], dyCm: move[1] } : null;
}

export interface AssemblyPreviewDeps {
  assembly: AssemblyController;
  /** Test seam: the decode step (the session cache by default). */
  decode?: (part: MediaPart, placedPx?: PixelSize) => Promise<DecodedImage>;
}

export interface AssemblyPreviewController {
  /** Function refs for the template: the canvas and its container. */
  attachCanvas: (el: unknown) => void;
  attachContainer: (el: unknown) => void;
  display: ComputedRef<PixelSize | null>;
  pixels: ComputedRef<PixelSize | null>;
  /** Completed renders (tests and the busy state read it). */
  renders: Ref<number>;
  render: () => Promise<void>;
  onPointerDown: (event: PointerEvent) => void;
  onPointerMove: (event: PointerEvent) => void;
  onPointerUp: (event: PointerEvent) => void;
  onKeydown: (event: KeyboardEvent) => void;
}

export function useAssemblyPreview(deps: AssemblyPreviewDeps): AssemblyPreviewController {
  const { assembly } = deps;
  const decode = deps.decode ?? decodeSource;
  const canvas = shallowRef<HTMLCanvasElement | null>(null);
  const box = ref<PixelSize>({ widthPx: 0, heightPx: 0 });
  const renders = ref(0);
  let observer: ResizeObserver | null = null;
  let generation = 0;

  const display = computed(() => (assembly.geometry.value ? previewDisplaySize(assembly.geometry.value, box.value) : null));
  const pixels = computed(() =>
    display.value ? previewPixelSize(display.value, typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1) : null,
  );
  const inputs = computed(() => composeInputs(assembly.tray.entries.value));

  async function render(): Promise<void> {
    generation += 1;
    const current = generation;
    const el = canvas.value;
    const geometry = assembly.geometry.value;
    const size = pixels.value;
    if (!el || !geometry || !size) {
      return;
    }
    const placed = assembly.placed.value;
    const stale = () => current !== generation;
    const sx = size.widthPx / geometry.widthCm;
    const sy = size.heightPx / geometry.heightCm;
    const prepared = await prepareSources(placed, inputs.value, sx, sy, decode, { cancelled: stale });
    if (stale()) {
      return;
    }
    if (el.width !== size.widthPx || el.height !== size.heightPx) {
      el.width = size.widthPx;
      el.height = size.heightPx;
    }
    if (await composeToCanvas(el, placed, { geometry, background: assembly.background.value }, prepared.sources, { cancelled: stale })) {
      renders.value += 1;
    }
  }

  watch([assembly.placed, assembly.geometry, assembly.background, pixels, canvas, inputs], () => {
    void render();
  });

  function attachCanvas(el: unknown): void {
    canvas.value = el instanceof HTMLCanvasElement ? el : null;
  }

  let observed: HTMLElement | null = null;

  // A function ref runs on every render of its component, so this attaches
  // only when the element really changes — a fresh observer per render would
  // report the same size every frame, and each report would re-render.
  function attachContainer(el: unknown): void {
    const next = el instanceof HTMLElement ? el : null;
    if (next === observed) {
      return;
    }
    observer?.disconnect();
    observer = null;
    observed = next;
    if (!next || typeof ResizeObserver !== "function") {
      return;
    }
    observer = new ResizeObserver((entries) => {
      const rect = entries[0]?.contentRect;
      if (!rect) {
        return;
      }
      const size = { widthPx: Math.floor(rect.width), heightPx: Math.floor(rect.height) };
      if (size.widthPx !== box.value.widthPx || size.heightPx !== box.value.heightPx) {
        box.value = size;
      }
    });
    observer.observe(next);
  }

  /** The sheet point under a pointer event, in centimetres. */
  function pointOf(event: PointerEvent): { xCm: number; yCm: number } | null {
    const el = canvas.value;
    const geometry = assembly.geometry.value;
    if (!el || !geometry) {
      return null;
    }
    const rect = el.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      return null;
    }
    return {
      xCm: ((event.clientX - rect.left) / rect.width) * geometry.widthCm,
      yCm: ((event.clientY - rect.top) / rect.height) * geometry.heightCm,
    };
  }

  let drag: { pointerId: number; key: string; start: OverlayTransform; xCm: number; yCm: number } | null = null;

  function onPointerDown(event: PointerEvent): void {
    const point = pointOf(event);
    if (!point || event.button !== 0) {
      return;
    }
    const hit = hitTest(assembly.placed.value, point.xCm, point.yCm);
    const entry = hit ? assembly.tray.entries.value[hit.index] : undefined;
    if (!entry) {
      return;
    }
    assembly.select(entry.key);
    if (assembly.isOverlay.value) {
      event.preventDefault();
      drag = { pointerId: event.pointerId, key: entry.key, start: { ...entry.transform }, ...point };
      canvas.value?.setPointerCapture?.(event.pointerId);
    }
  }

  function onPointerMove(event: PointerEvent): void {
    if (!drag || event.pointerId !== drag.pointerId) {
      return;
    }
    const point = pointOf(event);
    if (point) {
      const { dxCm, dyCm } = constrainDrag(point.xCm - drag.xCm, point.yCm - drag.yCm, event.shiftKey);
      assembly.dragTo(drag.key, drag.start, dxCm, dyCm);
    }
  }

  function onPointerUp(event: PointerEvent): void {
    if (drag && event.pointerId === drag.pointerId) {
      canvas.value?.releasePointerCapture?.(event.pointerId);
      drag = null;
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    const entry = assembly.overlayItem.value;
    const move = nudgeFor(event.key, event.shiftKey);
    if (!entry || !move || event.altKey || event.ctrlKey || event.metaKey) {
      return;
    }
    event.preventDefault();
    assembly.nudge(entry.key, move.dxCm, move.dyCm);
  }

  onScopeDispose(() => {
    observer?.disconnect();
    generation += 1;
  });

  return {
    attachCanvas,
    attachContainer,
    display,
    pixels,
    renders,
    render,
    onPointerDown,
    onPointerMove,
    onPointerUp,
    onKeydown,
  };
}
