// DOM/math helpers for the image resize-handle drag (SPEC-024 AC-10..12).
// Kept out of DocumentView.vue for the same reason tableResize.ts is: typed
// DOM element access needs real type annotations, which ESLint's espree
// parser rejects in `.vue` scripts.
import { MIN_IMAGE_SIZE_CM } from "./images";
import { CM_PER_PX } from "./tableResize";

export { CM_PER_PX };

/** The eight resize handles a selected image shows (AC-8): the four
 * corners resize proportionally (AC-10); the four edges resize one
 * dimension (AC-11) unless Shift constrains them too (AC-12). */
export type ImageHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export const IMAGE_HANDLES: ImageHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

/** Typed pass-through for a handle name coming from a `.vue` template's
 * plain string (espree bars the type annotation — see documentModel.ts's
 * asCharacterFormatKey for the same pattern). */
export function asImageHandle(value: string): ImageHandle {
  return value as ImageHandle;
}

export interface ImageResizeStart {
  widthCm: number;
  heightCm: number;
}

/** The image's own size ratio at drag start — computed once (not
 * re-derived from the live/possibly-already-resized dimensions) so a drag
 * that overshoots and comes back reproduces the exact starting size. */
function aspectRatioOf(start: ImageResizeStart): number {
  return start.heightCm > 0 ? start.widthCm / start.heightCm : 1;
}

/** The new size for a handle drag of `deltaXCm`/`deltaYCm` from the drag's
 * start size (AC-10..12). A corner handle always preserves aspect ratio,
 * driven by whichever axis moved further (so a diagonal drag feels
 * natural); an edge handle resizes only its own axis unless
 * `constrainAspect` (Shift held, AC-12) is set, in which case it drives the
 * other axis from the same aspect ratio too. Clamped to MIN_IMAGE_SIZE_CM. */
export function resizedImageSize(
  start: ImageResizeStart,
  handle: ImageHandle,
  deltaXCm: number,
  deltaYCm: number,
  constrainAspect: boolean,
): { widthCm: number; heightCm: number } {
  const isCorner = handle === "nw" || handle === "ne" || handle === "sw" || handle === "se";
  const xSign = handle.includes("e") ? 1 : handle.includes("w") ? -1 : 0;
  const ySign = handle.includes("s") ? 1 : handle.includes("n") ? -1 : 0;

  let width = start.widthCm + xSign * deltaXCm;
  let height = start.heightCm + ySign * deltaYCm;
  const ratio = aspectRatioOf(start);

  if (isCorner) {
    if (Math.abs(deltaXCm) >= Math.abs(deltaYCm)) {
      height = width / ratio;
    } else {
      width = height * ratio;
    }
  } else if (constrainAspect) {
    if (xSign !== 0) {
      height = width / ratio;
    } else if (ySign !== 0) {
      width = height * ratio;
    }
  }

  return {
    widthCm: Math.max(MIN_IMAGE_SIZE_CM, width),
    heightCm: Math.max(MIN_IMAGE_SIZE_CM, height),
  };
}

/** Live-previews a resize directly on the DOM (never the model mid-drag —
 * the same rule tableResize.ts's previewColumnWidths follows). */
export function previewImageSize(el: HTMLElement, widthCm: number, heightCm: number): void {
  el.style.width = `${widthCm}cm`;
  el.style.height = `${heightCm}cm`;
}

export interface ImageResizeDrag {
  section: number;
  block: number;
  handle: ImageHandle;
  startClientX: number;
  startClientY: number;
  start: ImageResizeStart;
  imageEl: HTMLElement;
}

/** Typed `null` seed for an image-resize-drag ref/variable in `.vue`
 * scripts (espree bars the annotation — see domCaret.ts's noElement). */
export function noImageResizeDrag(): ImageResizeDrag | null {
  return null;
}
