// Pure geometry for the slide stage's direct manipulation (SPEC-082): pointer
// coordinate conversion, resize-handle arithmetic, and centre/edge alignment
// guides with snapping. Framework-free so the view specs can exercise the
// gesture maths without mounting a component — the composables
// (useFrameInsertion / useFrameManipulation) call into it and own the DOM.
import { cmToPx } from "../../lib/core/units";
import { guideCandidates, snapRect, type Rect } from "../../lib/core/arrange";
import { unrotatePoint } from "../../lib/core/shapeGeometry";
import type { FrameRect } from "../../lib/presentation/frames";

/** How far a pointer travels (cm) before a press counts as a drag rather
 * than a click — 3px at 96dpi. Hoisted from useFrameManipulation's drag
 * closure (SPEC-096) so the marquee gesture shares the same threshold.
 * Deliberately NOT expressed in terms of core/shapes' DRAW_DRAG_THRESHOLD_CM:
 * "a drag has begun" and "a drawn box is worth keeping" may diverge. */
export const DRAG_ENGAGE_CM = (3 * 2.54) / 96;

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
  /** What each fired guide belongs to, aligned with `vertical`/`horizontal`
   * (SPEC-096 AC — `data-guide-kind`). Absent entries render as "slide". */
  verticalKinds?: ("slide" | "frame")[];
  horizontalKinds?: ("slide" | "frame")[];
}

/** Snaps a dragged rect to the slide's centrelines and zero-margin edges —
 * and, when `others` is given (SPEC-096), to any other top-level frame's
 * edges and centres — when within `thresholdCm`, returning the snapped rect
 * and the guides that fired (AC-19/20). At most one snap per axis, the
 * slide's centre/edges taking priority over frame lines. Re-expressed over
 * core/arrange's `guideCandidates`/`snapRect`; the four-argument behaviour
 * is unchanged. */
export function applyGuides(
  rect: FrameRect,
  slideWidthCm: number,
  slideHeightCm: number,
  thresholdCm: number,
  others: Rect[] = [],
): { rect: FrameRect; guides: StageGuides } {
  const container: Rect = { xCm: 0, yCm: 0, widthCm: slideWidthCm, heightCm: slideHeightCm };
  const snapped = snapRect(rect, guideCandidates(others, container), thresholdCm);
  const guides: StageGuides = { vertical: [], horizontal: [], verticalKinds: [], horizontalKinds: [] };
  for (const fired of snapped.fired) {
    // A line coinciding with the slide's own centre/edges reads as a slide
    // guide; everything else came from a frame.
    const slideLines =
      fired.axis === "x" ? [0, slideWidthCm / 2, slideWidthCm] : [0, slideHeightCm / 2, slideHeightCm];
    const kind = slideLines.includes(fired.line) ? "slide" : "frame";
    if (fired.axis === "x") {
      guides.vertical.push(fired.line);
      guides.verticalKinds!.push(kind);
    } else {
      guides.horizontal.push(fired.line);
      guides.horizontalKinds!.push(kind);
    }
  }
  return { rect: { ...rect, xCm: snapped.rect.xCm, yCm: snapped.rect.yCm }, guides };
}

// --- rotated resize (SPEC-096) ----------------------------------------------

/** The corner (or edge midpoint) opposite `handle`, in the rect's own
 * unrotated coordinates — the point that stays fixed in slide space. */
function anchorPoint(handle: ResizeHandle, rect: FrameRect): { x: number; y: number } {
  const x = handle.includes("w")
    ? rect.xCm + rect.widthCm
    : handle.includes("e")
      ? rect.xCm
      : rect.xCm + rect.widthCm / 2;
  const y = handle.includes("n")
    ? rect.yCm + rect.heightCm
    : handle.includes("s")
      ? rect.yCm
      : rect.yCm + rect.heightCm / 2;
  return { x, y };
}

function rotateVec(dx: number, dy: number, deg: number): { x: number; y: number } {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return { x: dx * cos - dy * sin, y: dx * sin + dy * cos };
}

/** Resize of a possibly-rotated frame ("Rotation maths"): the pointer maps
 * into the frame's unrotated local space with `unrotatePoint` about the
 * original centre, `resizeRect` runs unchanged there, and the result is
 * translated so the corner opposite the dragged handle stays fixed in slide
 * space. At 0° this is bit-identical to `resizeRect`. */
export function resizeRotatedRect(
  handle: ResizeHandle,
  orig: FrameRect,
  rotationDeg: number,
  pointer: { xCm: number; yCm: number },
  aspect: boolean,
): FrameRect {
  if (!rotationDeg) {
    return resizeRect(handle, orig, pointer, aspect);
  }
  const rel = unrotatePoint(
    { x: pointer.xCm - orig.xCm, y: pointer.yCm - orig.yCm },
    orig.widthCm,
    orig.heightCm,
    rotationDeg,
  );
  const local = resizeRect(handle, orig, { xCm: orig.xCm + rel.x, yCm: orig.yCm + rel.y }, aspect);
  // The anchor's slide-space position before the drag…
  const anchor = anchorPoint(handle, orig);
  const c0 = { x: orig.xCm + orig.widthCm / 2, y: orig.yCm + orig.heightCm / 2 };
  const spun = rotateVec(anchor.x - c0.x, anchor.y - c0.y, rotationDeg);
  const anchorSlide = { x: c0.x + spun.x, y: c0.y + spun.y };
  // …pins the new rect: newCentre = anchor + R(θ)·(localCentre − localAnchor).
  const localAnchor = anchorPoint(handle, local);
  const localCentre = { x: local.xCm + local.widthCm / 2, y: local.yCm + local.heightCm / 2 };
  const d = rotateVec(localCentre.x - localAnchor.x, localCentre.y - localAnchor.y, rotationDeg);
  return {
    xCm: anchorSlide.x + d.x - local.widthCm / 2,
    yCm: anchorSlide.y + d.y - local.heightCm / 2,
    widthCm: local.widthCm,
    heightCm: local.heightCm,
  };
}
