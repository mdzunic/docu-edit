// The image viewer's fit/zoom/pan state (SPEC-107), reusing SPEC-063's zoom
// vocabulary — lib/zoom.ts's clamp band and preset ladder — rather than
// inventing a second one. Pure state + arithmetic: the component owns the
// DOM (ResizeObserver, pointer events) and feeds sizes/gestures in.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import { clampZoom, stepZoomDown, stepZoomUp } from "../../lib/zoom";

/** Breathing room the fit computation leaves around the image, per side. */
export const FIT_MARGIN_PX = 16;

export interface ImageViewport {
  /** "fit" tracks the window live; "manual" holds a chosen percentage. */
  mode: Ref<"fit" | "manual">;
  /** The effective zoom percentage (fit-derived or manual). */
  percent: ComputedRef<number>;
  scale: ComputedRef<number>;
  pan: Ref<{ x: number; y: number }>;
  setContainerSize: (width: number, height: number) => void;
  setImageSize: (width: number, height: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  /** 100% ("actual size"), the Cmd/Ctrl+0 target per SPEC-063. */
  zoomActual: () => void;
  zoomFit: () => void;
  /** Double-click: fit ⇄ 100%. */
  toggleFitActual: () => void;
  panBy: (dx: number, dy: number) => void;
  /** Keyboard vocabulary (Cmd/Ctrl +, -, 0). True when handled. */
  handleZoomKey: (event: KeyboardEvent) => boolean;
}

export function useImageViewport(): ImageViewport {
  const mode = ref<"fit" | "manual">("fit");
  const manualPercent = ref(100);
  const containerSize = ref({ width: 0, height: 0 });
  const imageSize = ref({ width: 0, height: 0 });
  const pan = ref({ x: 0, y: 0 });

  const fitPercent = computed(() => {
    const cw = containerSize.value.width - FIT_MARGIN_PX * 2;
    const ch = containerSize.value.height - FIT_MARGIN_PX * 2;
    const iw = imageSize.value.width;
    const ih = imageSize.value.height;
    if (!(cw > 0) || !(ch > 0) || !(iw > 0) || !(ih > 0)) {
      return 100;
    }
    return clampZoom(Math.round(Math.min(cw / iw, ch / ih) * 100));
  });

  const percent = computed(() => (mode.value === "fit" ? fitPercent.value : manualPercent.value));
  const scale = computed(() => percent.value / 100);

  function setContainerSize(width: number, height: number): void {
    containerSize.value = { width, height };
  }

  function setImageSize(width: number, height: number): void {
    imageSize.value = { width, height };
  }

  function toManual(next: number): void {
    manualPercent.value = clampZoom(next);
    mode.value = "manual";
  }

  function zoomIn(): void {
    toManual(stepZoomUp(percent.value));
  }

  function zoomOut(): void {
    toManual(stepZoomDown(percent.value));
  }

  function zoomActual(): void {
    toManual(100);
  }

  function zoomFit(): void {
    mode.value = "fit";
    pan.value = { x: 0, y: 0 };
  }

  function toggleFitActual(): void {
    if (mode.value === "fit") {
      zoomActual();
    } else {
      zoomFit();
    }
  }

  function panBy(dx: number, dy: number): void {
    pan.value = { x: pan.value.x + dx, y: pan.value.y + dy };
  }

  function handleZoomKey(event: KeyboardEvent): boolean {
    if (!(event.metaKey || event.ctrlKey) || event.altKey) {
      return false;
    }
    if (event.key === "+" || event.key === "=") {
      zoomIn();
      return true;
    }
    if (event.key === "-") {
      zoomOut();
      return true;
    }
    if (event.key === "0") {
      zoomActual();
      return true;
    }
    return false;
  }

  return {
    mode,
    percent,
    scale,
    pan,
    setContainerSize,
    setImageSize,
    zoomIn,
    zoomOut,
    zoomActual,
    zoomFit,
    toggleFitActual,
    panBy,
    handleZoomKey,
  };
}
