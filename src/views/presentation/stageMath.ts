// Pure geometry for the slide stage's direct manipulation (SPEC-082): pointer
// coordinate conversion, resize-handle arithmetic, and centre/edge alignment
// guides with snapping. Framework-free so the view specs can exercise the
// gesture maths without mounting a component — the composables
// (useFrameInsertion / useFrameManipulation) call into it and own the DOM.
import { cmToPx } from "../../lib/core/units";
import type { FrameRect } from "../../lib/presentation/frames";

/** The eight resize handles, named by compass point (nw = top-left). Corner
 * handles (four-letter) keep aspect under the Shift/image rule. */
export type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export const RESIZE_HANDLES: ResizeHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

export function isCornerHandle(handle: ResizeHandle): boolean {
  return handle.length === 2;
}

/** Converts a client (viewport) point to slide-space centimetres, using the
 * surface element's on-screen rect — which already folds in the fit scale, so
 * no separate scale is needed. */
export function clientToCm(
  clientX: number,
  clientY: number,
  surface: DOMRect,
  slideWidthCm: number,
  slideHeightCm: number,
): { xCm: number; yCm: number } {
  const pxPerCmX = surface.width / slideWidthCm;
  const pxPerCmY = surface.height / slideHeightCm;
  return {
    xCm: pxPerCmX > 0 ? (clientX - surface.left) / pxPerCmX : 0,
    yCm: pxPerCmY > 0 ? (clientY - surface.top) / pxPerCmY : 0,
  };
}

/** A centimetre length as on-screen pixels at the surface's scale. */
export function cmToStagePx(cm: number, surface: DOMRect, slideWidthCm: number): number {
  return slideWidthCm > 0 ? cm * (surface.width / slideWidthCm) : cmToPx(cm);
}

const MIN_SIZE_CM = 0.5;

/** The new rect a resize handle produces from a pointer position (cm), keeping
 * the opposite edge/corner fixed and enforcing the 0.5cm minimum. When
 * `aspect` is set (Shift on a corner, or an image), a corner drag preserves
 * the original width/height ratio. */
export function resizeRect(
  handle: ResizeHandle,
  orig: FrameRect,
  pointer: { xCm: number; yCm: number },
  aspect: boolean,
): FrameRect {
  const left = orig.xCm;
  const top = orig.yCm;
  const right = orig.xCm + orig.widthCm;
  const bottom = orig.yCm + orig.heightCm;

  let newLeft = left;
  let newTop = top;
  let newRight = right;
  let newBottom = bottom;

  if (handle.includes("w")) newLeft = pointer.xCm;
  if (handle.includes("e")) newRight = pointer.xCm;
  if (handle.includes("n")) newTop = pointer.yCm;
  if (handle.includes("s")) newBottom = pointer.yCm;

  let width = Math.max(newRight - newLeft, MIN_SIZE_CM);
  let height = Math.max(newBottom - newTop, MIN_SIZE_CM);

  if (aspect && isCornerHandle(handle) && orig.widthCm > 0 && orig.heightCm > 0) {
    const ratio = orig.widthCm / orig.heightCm;
    // Grow to the larger of the two proposed dimensions, keeping ratio.
    if (width / ratio >= height) {
      height = width / ratio;
    } else {
      width = height * ratio;
    }
  }

  // Anchor: the edge opposite the moved one stays put.
  const x = handle.includes("w") ? right - width : left;
  const y = handle.includes("n") ? bottom - height : top;
  return { xCm: x, yCm: y, widthCm: width, heightCm: height };
}

export interface StageGuides {
  /** Slide x-positions to draw a vertical guide at (cm). */
  vertical: number[];
  /** Slide y-positions to draw a horizontal guide at (cm). */
  horizontal: number[];
}

/** Snaps a dragged rect to the slide's centrelines and zero-margin edges when
 * within `thresholdCm`, returning the snapped rect and the guides that fired
 * (AC-19/20). At most one snap per axis, centre taking priority over edges. */
export function applyGuides(
  rect: FrameRect,
  slideWidthCm: number,
  slideHeightCm: number,
  thresholdCm: number,
): { rect: FrameRect; guides: StageGuides } {
  const guides: StageGuides = { vertical: [], horizontal: [] };
  let { xCm, yCm } = rect;

  const centerX = xCm + rect.widthCm / 2;
  const xTargets: { line: number; x: number; dist: number }[] = [
    { line: slideWidthCm / 2, x: slideWidthCm / 2 - rect.widthCm / 2, dist: Math.abs(centerX - slideWidthCm / 2) },
    { line: 0, x: 0, dist: Math.abs(xCm) },
    { line: slideWidthCm, x: slideWidthCm - rect.widthCm, dist: Math.abs(xCm + rect.widthCm - slideWidthCm) },
  ];
  const xSnap = xTargets.filter((t) => t.dist <= thresholdCm).sort((a, b) => a.dist - b.dist)[0];
  if (xSnap) {
    xCm = xSnap.x;
    guides.vertical.push(xSnap.line);
  }

  const centerY = yCm + rect.heightCm / 2;
  const yTargets: { line: number; y: number; dist: number }[] = [
    { line: slideHeightCm / 2, y: slideHeightCm / 2 - rect.heightCm / 2, dist: Math.abs(centerY - slideHeightCm / 2) },
    { line: 0, y: 0, dist: Math.abs(yCm) },
    { line: slideHeightCm, y: slideHeightCm - rect.heightCm, dist: Math.abs(yCm + rect.heightCm - slideHeightCm) },
  ];
  const ySnap = yTargets.filter((t) => t.dist <= thresholdCm).sort((a, b) => a.dist - b.dist)[0];
  if (ySnap) {
    yCm = ySnap.y;
    guides.horizontal.push(ySnap.line);
  }

  return { rect: { ...rect, xCm, yCm }, guides };
}
