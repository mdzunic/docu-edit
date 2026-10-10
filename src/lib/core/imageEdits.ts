// The non-destructive image edit stack (SPEC-112) — THIS FILE is the
// normative source for `ImageEdit`, for the order edits are applied in and
// for the geometry a stack resolves to. Pure and framework-free: every
// function here is given only the source's natural pixel size, never its
// bytes, so every geometric consequence of a stack is unit-testable without
// a canvas, and the preview and the export (lib/media/renderEdits.ts) read
// one answer.
//
// A stack is a canonical program, not a log. `normaliseStack` sorts edits
// into the fixed application order
//   straighten → crop → rotate → flip → resize → adjust → redact → annotate
// whatever order the user worked in (all redactions before all
// annotations, each in insertion order), and collapses repeats of the
// single-valued kinds:
// - straighten, crop, resize: the last one wins (a second crop replaces the
//   first);
// - rotate: quarter turns compose (two "rotate right"s are 180°);
// - flip: per axis, by parity (flipping twice is no flip);
// - adjust: merged channel by channel, later values winning;
// while redact and annotate accumulate in insertion order. An entry that
// changes nothing (0° of straighten, no rotation, all-neutral adjust) is
// dropped, so two stacks that draw the same picture normalise alike.
//
// Coordinate spaces:
// - straighten rotates the source about its centre inside the source's own
//   W × H frame (positive degrees clockwise, the canvas convention);
// - crop is a rectangle in that straightened frame, in source pixels, so
//   adjusting the angle later keeps the crop meaningful rather than
//   re-basing it. It is clamped into the largest axis-aligned rectangle of
//   the source's aspect that the rotated source still fully covers (the
//   "inscribed" area) — which is also the source rectangle of a stack that
//   straightens without cropping, so no empty corner ever reaches the output;
// - redact and annotate coordinates are post-geometry OUTPUT pixels (after
//   straighten/crop/rotate/flip/resize).
import { ADJUST_CHANNELS, clampAdjust, isNeutralAdjust, type AdjustSettings } from "./imageAdjust";

export type AnnotationShape =
  | { shape: "arrow"; x1: number; y1: number; x2: number; y2: number }
  | { shape: "line"; x1: number; y1: number; x2: number; y2: number }
  | { shape: "rectangle"; x: number; y: number; width: number; height: number }
  | { shape: "ellipse"; x: number; y: number; width: number; height: number }
  | { shape: "freehand"; points: Array<{ x: number; y: number }> }
  | { shape: "highlight"; x: number; y: number; width: number; height: number }
  | { shape: "text"; x: number; y: number; text: string; fontSize: number };

export type AnnotationKind = AnnotationShape["shape"];

export type ImageEdit =
  | { kind: "straighten"; degrees: number }
  | { kind: "crop"; x: number; y: number; width: number; height: number }
  | { kind: "rotate"; quarterTurns: 0 | 1 | 2 | 3 }
  | { kind: "flip"; axis: "horizontal" | "vertical" }
  | { kind: "resize"; widthPx: number; heightPx: number; resample: "nearest" | "bilinear" }
  | ({ kind: "adjust" } & AdjustSettings)
  | { kind: "redact"; x: number; y: number; width: number; height: number; color: string }
  | { kind: "annotate"; shape: AnnotationShape; color: string; strokeWidth: number; opacity: number };

export type EditKind = ImageEdit["kind"];

/** Every edit kind, in application order. */
export const EDIT_KINDS: readonly EditKind[] = [
  "straighten",
  "crop",
  "rotate",
  "flip",
  "resize",
  "adjust",
  "redact",
  "annotate",
];

/** The annotation shapes, in the order the annotate tool offers them. */
export const ANNOTATION_KINDS: readonly AnnotationKind[] = [
  "arrow",
  "line",
  "rectangle",
  "ellipse",
  "freehand",
  "highlight",
  "text",
];

/** The longest a normalised stack may be: redactions and annotations past
 * it are dropped (the newest first), and the editor refuses to add one. */
export const MAX_EDIT_STACK = 500;

/** Straighten's range, in degrees either way. */
export const MAX_STRAIGHTEN_DEGREES = 45;

/** Either output edge's ceiling — SPEC-109's MAX_OUTPUT_EDGE_PX. */
export const MAX_EDIT_EDGE_PX = 16384;

export const DEFAULT_REDACT_COLOR = "#000000";

/** A freehand stroke keeps at most this many points. */
export const MAX_FREEHAND_POINTS = 2000;

export interface PixelSize {
  widthPx: number;
  heightPx: number;
}

export interface PixelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A 2-D affine transform in the canvas `setTransform(a, b, c, d, e, f)`
 * convention: x′ = a·x + c·y + e, y′ = b·x + d·y + f. */
export interface AffineMatrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

export interface EditGeometry {
  /** The output size after every geometric edit. */
  widthPx: number;
  heightPx: number;
  /** The source rectangle the output samples, in the straightened frame,
   * before rotation/flip. */
  source: PixelRect;
  quarterTurns: 0 | 1 | 2 | 3;
  flipH: boolean;
  flipV: boolean;
  straightenDegrees: number;
  /** The area a crop is clamped into: the whole source unless straightened. */
  inscribed: PixelRect;
  /** The output size before any resize (the source rectangle, rotated). */
  preResizePx: PixelSize;
  /** The cumulative transform from natural source pixels to output pixels:
   * drawing the natural image through it onto a widthPx × heightPx canvas
   * IS the geometric half of the render. */
  transform: AffineMatrix;
}

// --- sanitising ------------------------------------------------------------

function finite(value: unknown, fallback = 0): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function edge(value: unknown): number {
  return Math.min(MAX_EDIT_EDGE_PX, Math.max(1, Math.round(finite(value, 1))));
}

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

/** A "#rrggbb" colour, lower-cased; anything else falls back. */
export function normaliseColor(value: unknown, fallback = DEFAULT_REDACT_COLOR): string {
  return typeof value === "string" && HEX_COLOR.test(value) ? value.toLowerCase() : fallback;
}

function rectOf(edit: { x: unknown; y: unknown; width: unknown; height: unknown }): PixelRect {
  return {
    x: Math.round(finite(edit.x)),
    y: Math.round(finite(edit.y)),
    width: Math.max(1, Math.round(finite(edit.width, 1))),
    height: Math.max(1, Math.round(finite(edit.height, 1))),
  };
}

function point(x: unknown, y: unknown): { x: number; y: number } {
  return { x: finite(x), y: finite(y) };
}

/** An annotation shape with finite coordinates, or null when it is not one. */
function sanitiseShape(shape: AnnotationShape | undefined): AnnotationShape | null {
  if (!shape || typeof shape !== "object") {
    return null;
  }
  switch (shape.shape) {
    case "arrow":
    case "line": {
      const from = point(shape.x1, shape.y1);
      const to = point(shape.x2, shape.y2);
      return { shape: shape.shape, x1: from.x, y1: from.y, x2: to.x, y2: to.y };
    }
    case "rectangle":
    case "ellipse":
    case "highlight":
      return {
        shape: shape.shape,
        x: finite(shape.x),
        y: finite(shape.y),
        width: Math.max(1, finite(shape.width, 1)),
        height: Math.max(1, finite(shape.height, 1)),
      };
    case "freehand": {
      const points = Array.isArray(shape.points) ? shape.points.slice(0, MAX_FREEHAND_POINTS) : [];
      return points.length > 0 ? { shape: "freehand", points: points.map((p) => point(p?.x, p?.y)) } : null;
    }
    case "text": {
      const text = typeof shape.text === "string" ? shape.text : "";
      return text.length > 0
        ? { shape: "text", x: finite(shape.x), y: finite(shape.y), text, fontSize: Math.max(4, finite(shape.fontSize, 24)) }
        : null;
    }
    default:
      return null;
  }
}

/** The adjust channels that are set and not neutral, clamped. */
function adjustSettingsOf(edit: AdjustSettings): AdjustSettings {
  const settings: AdjustSettings = {};
  for (const channel of ADJUST_CHANNELS) {
    if (channel === "grayscale" || channel === "invert") {
      if (edit[channel] === true) {
        settings[channel] = true;
      }
    } else {
      const value = Math.round(clampAdjust(finite(edit[channel])));
      if (value !== 0) {
        settings[channel] = value;
      }
    }
  }
  return settings;
}

// --- normalisation ---------------------------------------------------------

/** The canonical program for any list of edits (see the module header):
 * application order, single-valued kinds collapsed, identities dropped,
 * redact/annotate accumulated in insertion order and bounded by
 * MAX_EDIT_STACK. Pure and total — malformed entries are dropped, never
 * thrown on — and every returned edit is a fresh object. */
export function normaliseStack(edits: readonly ImageEdit[]): ImageEdit[] {
  let straighten = 0;
  let crop: PixelRect | null = null;
  let quarterTurns = 0;
  let flipH = false;
  let flipV = false;
  let resize: Extract<ImageEdit, { kind: "resize" }> | null = null;
  // Channel-by-channel merge: a later value for a channel replaces the
  // earlier one, including a later neutral value.
  const adjust: Record<string, number | boolean> = {};
  const redactions: ImageEdit[] = [];
  const annotations: ImageEdit[] = [];

  for (const edit of Array.isArray(edits) ? edits : []) {
    if (!edit || typeof edit !== "object") {
      continue;
    }
    switch (edit.kind) {
      case "straighten":
        straighten = Math.min(MAX_STRAIGHTEN_DEGREES, Math.max(-MAX_STRAIGHTEN_DEGREES, finite(edit.degrees)));
        break;
      case "crop":
        crop = rectOf(edit);
        crop.x = Math.max(0, crop.x);
        crop.y = Math.max(0, crop.y);
        break;
      case "rotate":
        quarterTurns += Math.round(finite(edit.quarterTurns));
        break;
      case "flip":
        if (edit.axis === "horizontal") {
          flipH = !flipH;
        } else if (edit.axis === "vertical") {
          flipV = !flipV;
        }
        break;
      case "resize":
        resize = {
          kind: "resize",
          widthPx: edge(edit.widthPx),
          heightPx: edge(edit.heightPx),
          resample: edit.resample === "nearest" ? "nearest" : "bilinear",
        };
        break;
      case "adjust":
        for (const channel of ADJUST_CHANNELS) {
          const value = edit[channel];
          if (value !== undefined) {
            adjust[channel] = value;
          }
        }
        break;
      case "redact": {
        const rect = rectOf(edit);
        redactions.push({ kind: "redact", ...rect, color: normaliseColor(edit.color) });
        break;
      }
      case "annotate": {
        const shape = sanitiseShape(edit.shape);
        if (shape) {
          annotations.push({
            kind: "annotate",
            shape,
            color: normaliseColor(edit.color),
            strokeWidth: Math.min(200, Math.max(0.5, finite(edit.strokeWidth, 4))),
            opacity: Math.min(1, Math.max(0, finite(edit.opacity, 1))),
          });
        }
        break;
      }
      default:
        break;
    }
  }

  const stack: ImageEdit[] = [];
  // Rounded to 1/100°: the UI's precision, and what keeps a stored angle
  // from differing from the shown one in the last bit.
  const degrees = Math.round(straighten * 100) / 100;
  if (degrees !== 0) {
    stack.push({ kind: "straighten", degrees });
  }
  if (crop) {
    stack.push({ kind: "crop", ...crop });
  }
  const turns = (((quarterTurns % 4) + 4) % 4) as 0 | 1 | 2 | 3;
  if (turns !== 0) {
    stack.push({ kind: "rotate", quarterTurns: turns });
  }
  if (flipH) {
    stack.push({ kind: "flip", axis: "horizontal" });
  }
  if (flipV) {
    stack.push({ kind: "flip", axis: "vertical" });
  }
  if (resize) {
    stack.push(resize);
  }
  const settings = adjustSettingsOf(adjust as AdjustSettings);
  if (!isNeutralAdjust(settings)) {
    stack.push({ kind: "adjust", ...settings });
  }
  // Past the bound, the newest objects go first — annotations before any
  // redaction, so a cover is never dropped to keep a doodle.
  const room = Math.max(0, MAX_EDIT_STACK - stack.length);
  stack.push(...redactions.slice(0, room));
  stack.push(...annotations.slice(0, Math.max(0, room - redactions.length)));
  return stack;
}

// --- geometry --------------------------------------------------------------

const EPSILON = 1e-9;

/** The largest axis-aligned rectangle of the source's own aspect, centred,
 * that a `degrees` rotation of a W × H source about its centre still fully
 * covers — in whole pixels, never overhanging the exact area. 0° is the
 * whole source. */
export function inscribedRect(naturalPx: PixelSize, degrees: number): PixelRect {
  const width = Math.max(1, Math.round(finite(naturalPx.widthPx, 1)));
  const height = Math.max(1, Math.round(finite(naturalPx.heightPx, 1)));
  const radians = (finite(degrees) * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  // A centred w × h box, rotated back by −θ, must fit the W × H source: its
  // bounding box w·cos + h·sin ≤ W and w·sin + h·cos ≤ H, with w:h = W:H.
  const scale = Math.min(1, width / (width * cos + height * sin), height / (width * sin + height * cos));
  const span = (full: number): [number, number] => {
    // `+ 0` turns the −0 a ceil of −ε gives into 0.
    const lo = Math.ceil((full - full * scale) / 2 - EPSILON) + 0;
    const hi = Math.floor((full + full * scale) / 2 + EPSILON);
    return hi > lo ? [lo, hi] : [Math.floor(full / 2), Math.floor(full / 2) + 1];
  };
  const [x0, x1] = span(width);
  const [y0, y1] = span(height);
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
}

/** `rect` clamped into `bounds` (whole pixels, at least 1 × 1). */
export function clampRect(rect: PixelRect, bounds: PixelRect): PixelRect {
  const bx1 = bounds.x + bounds.width;
  const by1 = bounds.y + bounds.height;
  let x0 = Math.min(Math.max(Math.round(rect.x), bounds.x), bx1 - 1);
  let y0 = Math.min(Math.max(Math.round(rect.y), bounds.y), by1 - 1);
  const x1 = Math.max(x0 + 1, Math.min(Math.round(rect.x + rect.width), bx1));
  const y1 = Math.max(y0 + 1, Math.min(Math.round(rect.y + rect.height), by1));
  x0 = Math.max(bounds.x, Math.min(x0, x1 - 1));
  y0 = Math.max(bounds.y, Math.min(y0, y1 - 1));
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
}

const IDENTITY: AffineMatrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** `left ∘ right`: apply `right` first, then `left`. */
export function multiplyMatrix(left: AffineMatrix, right: AffineMatrix): AffineMatrix {
  return {
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  };
}

/** Applies a matrix to a point. */
export function applyMatrix(matrix: AffineMatrix, x: number, y: number): { x: number; y: number } {
  return { x: matrix.a * x + matrix.c * y + matrix.e, y: matrix.b * x + matrix.d * y + matrix.f };
}

function translate(x: number, y: number): AffineMatrix {
  return { a: 1, b: 0, c: 0, d: 1, e: x, f: y };
}

/** A scale matrix. */
export function scaleMatrix(sx: number, sy: number): AffineMatrix {
  return { a: sx, b: 0, c: 0, d: sy, e: 0, f: 0 };
}

/** Quarter turns clockwise of a `width` × `height` frame, mapped into the
 * turned frame's own origin. Exact: every coefficient is 0 or ±1. */
function quarterTurnMatrix(turns: 0 | 1 | 2 | 3, width: number, height: number): AffineMatrix {
  switch (turns) {
    case 1:
      return { a: 0, b: 1, c: -1, d: 0, e: height, f: 0 };
    case 2:
      return { a: -1, b: 0, c: 0, d: -1, e: width, f: height };
    case 3:
      return { a: 0, b: -1, c: 1, d: 0, e: 0, f: width };
    default:
      return IDENTITY;
  }
}

/** What a stack does to a source of `naturalPx`: the output size, the source
 * rectangle sampled, the orientation and the cumulative transform. Pure and
 * total; the empty stack is the identity (the whole source, unchanged). */
export function resolveEditGeometry(edits: readonly ImageEdit[], naturalPx: PixelSize): EditGeometry {
  const stack = normaliseStack(edits);
  const width = Math.max(1, Math.round(finite(naturalPx.widthPx, 1)));
  const height = Math.max(1, Math.round(finite(naturalPx.heightPx, 1)));
  let degrees = 0;
  let crop: PixelRect | null = null;
  let quarterTurns: 0 | 1 | 2 | 3 = 0;
  let flipH = false;
  let flipV = false;
  let resize: PixelSize | null = null;
  for (const edit of stack) {
    if (edit.kind === "straighten") {
      degrees = edit.degrees;
    } else if (edit.kind === "crop") {
      crop = edit;
    } else if (edit.kind === "rotate") {
      quarterTurns = edit.quarterTurns;
    } else if (edit.kind === "flip") {
      flipH = flipH || edit.axis === "horizontal";
      flipV = flipV || edit.axis === "vertical";
    } else if (edit.kind === "resize") {
      resize = { widthPx: edit.widthPx, heightPx: edit.heightPx };
    }
  }
  const inscribed = inscribedRect({ widthPx: width, heightPx: height }, degrees);
  const source = clampRect(crop ?? inscribed, inscribed);
  const turned = quarterTurns % 2 === 1;
  const preResizePx = turned
    ? { widthPx: source.height, heightPx: source.width }
    : { widthPx: source.width, heightPx: source.height };
  const output = resize ?? preResizePx;

  // natural → straightened (about the centre) → cropped → turned → flipped → resized
  let transform = IDENTITY;
  if (degrees !== 0) {
    const radians = (degrees * Math.PI) / 180;
    const cos = Math.cos(radians);
    const sin = Math.sin(radians);
    const rotation: AffineMatrix = { a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 };
    transform = multiplyMatrix(translate(width / 2, height / 2), multiplyMatrix(rotation, translate(0 - width / 2, 0 - height / 2)));
  }
  transform = multiplyMatrix(translate(0 - source.x, 0 - source.y), transform);
  transform = multiplyMatrix(quarterTurnMatrix(quarterTurns, source.width, source.height), transform);
  if (flipH) {
    transform = multiplyMatrix({ a: -1, b: 0, c: 0, d: 1, e: preResizePx.widthPx, f: 0 }, transform);
  }
  if (flipV) {
    transform = multiplyMatrix({ a: 1, b: 0, c: 0, d: -1, e: 0, f: preResizePx.heightPx }, transform);
  }
  if (output.widthPx !== preResizePx.widthPx || output.heightPx !== preResizePx.heightPx) {
    transform = multiplyMatrix(
      scaleMatrix(output.widthPx / preResizePx.widthPx, output.heightPx / preResizePx.heightPx),
      transform,
    );
  }
  return {
    widthPx: output.widthPx,
    heightPx: output.heightPx,
    source,
    quarterTurns,
    flipH,
    flipV,
    straightenDegrees: degrees,
    inscribed,
    preResizePx,
    transform,
  };
}

// --- crop helpers ------------------------------------------------------------

/** The crop tool's aspect presets. */
export const CROP_ASPECTS = ["free", "original", "1:1", "4:3", "3:2", "16:9"] as const;

export type CropAspect = (typeof CROP_ASPECTS)[number];

export const CROP_ASPECT_LABELS: Record<CropAspect, string> = {
  free: "Free",
  original: "Original",
  "1:1": "1:1",
  "4:3": "4:3",
  "3:2": "3:2",
  "16:9": "16:9",
};

/** A preset's width ÷ height, or null for Free. */
export function cropAspectRatio(aspect: CropAspect, naturalPx: PixelSize): number | null {
  switch (aspect) {
    case "original":
      return Math.max(1, naturalPx.widthPx) / Math.max(1, naturalPx.heightPx);
    case "1:1":
      return 1;
    case "4:3":
      return 4 / 3;
    case "3:2":
      return 3 / 2;
    case "16:9":
      return 16 / 9;
    default:
      return null;
  }
}

/** The largest rectangle of `ratio` centred in `bounds` (all of `bounds`
 * for a null ratio). */
export function cropForAspect(ratio: number | null, bounds: PixelRect): PixelRect {
  if (ratio === null || !(ratio > 0)) {
    return { ...bounds };
  }
  let width = bounds.width;
  let height = Math.round(width / ratio);
  if (height > bounds.height) {
    height = bounds.height;
    width = Math.min(bounds.width, Math.round(height * ratio));
  }
  width = Math.max(1, width);
  height = Math.max(1, height);
  return {
    x: bounds.x + Math.floor((bounds.width - width) / 2),
    y: bounds.y + Math.floor((bounds.height - height) / 2),
    width,
    height,
  };
}

/** A crop handle: a compass edge or corner, or the whole rectangle. */
export type CropHandle = "move" | "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

export const CROP_HANDLES: readonly CropHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w", "move"];

/** Moves one crop handle by (dx, dy) source pixels inside `bounds`. With a
 * `ratio` the aspect is held: the edge or corner opposite the handle stays
 * anchored and the other dimension follows the driven one. */
export function moveCropHandle(
  rect: PixelRect,
  handle: CropHandle,
  dx: number,
  dy: number,
  ratio: number | null,
  bounds: PixelRect,
): PixelRect {
  const start = clampRect(rect, bounds);
  const bx1 = bounds.x + bounds.width;
  const by1 = bounds.y + bounds.height;
  if (handle === "move") {
    const x = Math.min(Math.max(start.x + Math.round(dx), bounds.x), bx1 - start.width);
    const y = Math.min(Math.max(start.y + Math.round(dy), bounds.y), by1 - start.height);
    return { ...start, x, y };
  }
  const west = handle.includes("w");
  const east = handle.includes("e");
  const north = handle.includes("n");
  const south = handle.includes("s");
  let left = start.x;
  let top = start.y;
  let right = start.x + start.width;
  let bottom = start.y + start.height;
  if (ratio === null || !(ratio > 0)) {
    if (west) left = Math.min(Math.max(left + Math.round(dx), bounds.x), right - 1);
    if (east) right = Math.max(Math.min(right + Math.round(dx), bx1), left + 1);
    if (north) top = Math.min(Math.max(top + Math.round(dy), bounds.y), bottom - 1);
    if (south) bottom = Math.max(Math.min(bottom + Math.round(dy), by1), top + 1);
    return { x: left, y: top, width: right - left, height: bottom - top };
  }
  // Aspect-locked: anchor the opposite side(s). A pure edge handle grows
  // away from its opposite edge and towards the bottom/right.
  const anchorX = west ? right : left;
  const anchorY = north ? bottom : top;
  const widthDriven = east || west ? (north || south ? Math.abs(dx) >= Math.abs(dy) : true) : false;
  let width = start.width;
  let height = start.height;
  if (widthDriven) {
    width = start.width + Math.round(west ? -dx : dx);
    height = Math.round(width / ratio);
  } else {
    height = start.height + Math.round(north ? -dy : dy);
    width = Math.round(height * ratio);
  }
  const maxWidth = west ? anchorX - bounds.x : bx1 - anchorX;
  const maxHeight = north ? anchorY - bounds.y : by1 - anchorY;
  if (width > maxWidth) {
    width = maxWidth;
    height = Math.round(width / ratio);
  }
  if (height > maxHeight) {
    height = maxHeight;
    width = Math.round(height * ratio);
  }
  width = Math.max(1, Math.min(width, maxWidth));
  height = Math.max(1, Math.min(height, maxHeight));
  return {
    x: west ? anchorX - width : anchorX,
    y: north ? anchorY - height : anchorY,
    width,
    height,
  };
}

/** How a crop rectangle is announced: "x, y, width × height". */
export function cropAnnouncement(rect: PixelRect): string {
  return `${rect.x}, ${rect.y}, ${rect.width} × ${rect.height}`;
}

// --- objects (redact / annotate) --------------------------------------------

/** Whether an edit is one of the accumulating objects drawn over the image. */
export function isObjectEdit(edit: ImageEdit): edit is Extract<ImageEdit, { kind: "redact" | "annotate" }> {
  return edit.kind === "redact" || edit.kind === "annotate";
}

function translateShape(shape: AnnotationShape, dx: number, dy: number): AnnotationShape {
  switch (shape.shape) {
    case "arrow":
    case "line":
      return { ...shape, x1: shape.x1 + dx, y1: shape.y1 + dy, x2: shape.x2 + dx, y2: shape.y2 + dy };
    case "freehand":
      return { ...shape, points: shape.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
    default:
      return { ...shape, x: shape.x + dx, y: shape.y + dy };
  }
}

/** A redaction or annotation moved by (dx, dy) output pixels; any other edit
 * comes back unchanged. */
export function translateEdit(edit: ImageEdit, dx: number, dy: number): ImageEdit {
  if (edit.kind === "redact") {
    return { ...edit, x: edit.x + dx, y: edit.y + dy };
  }
  if (edit.kind === "annotate") {
    return { ...edit, shape: translateShape(edit.shape, dx, dy) };
  }
  return edit;
}

// --- descriptions ----------------------------------------------------------

const SHAPE_NOUNS: Record<AnnotationKind, string> = {
  arrow: "arrow",
  line: "line",
  rectangle: "rectangle",
  ellipse: "ellipse",
  freehand: "freehand drawing",
  highlight: "highlight",
  text: "text",
};

/** The noun an annotation shape reads as ("arrow", "freehand drawing"). */
export function annotationNoun(shape: AnnotationKind): string {
  return SHAPE_NOUNS[shape];
}

function listWords(words: string[]): string {
  if (words.length <= 1) {
    return words.join("");
  }
  return `${words.slice(0, -1).join(", ")} and ${words[words.length - 1]}`;
}

/** The history list's (and the undo label's) name for one edit: "Crop",
 * "Rotate right", "Adjust brightness", "Add arrow". */
export function describeEdit(edit: ImageEdit): string {
  switch (edit.kind) {
    case "straighten":
      return `Straighten ${edit.degrees > 0 ? "+" : ""}${edit.degrees}°`;
    case "crop":
      return "Crop";
    case "rotate":
      return edit.quarterTurns === 1 ? "Rotate right" : edit.quarterTurns === 3 ? "Rotate left" : edit.quarterTurns === 2 ? "Rotate 180°" : "Rotate";
    case "flip":
      return edit.axis === "horizontal" ? "Flip horizontal" : "Flip vertical";
    case "resize":
      return `Resize to ${edit.widthPx} × ${edit.heightPx}`;
    case "adjust": {
      const channels = ADJUST_CHANNELS.filter((channel) =>
        channel === "grayscale" || channel === "invert" ? edit[channel] === true : typeof edit[channel] === "number" && edit[channel] !== 0,
      );
      return channels.length > 0 ? `Adjust ${listWords([...channels])}` : "Adjust";
    }
    case "redact":
      return "Redact";
    case "annotate":
      return `Add ${annotationNoun(edit.shape.shape)}`;
    default:
      return "Edit";
  }
}

function signed(value: number): string {
  return value > 0 ? `+${value}` : String(value);
}

/** The history row's second line — the edit's values — or "". */
export function describeEditDetail(edit: ImageEdit): string {
  switch (edit.kind) {
    case "crop":
    case "redact":
      return cropAnnouncement(edit);
    case "adjust":
      return ADJUST_CHANNELS.flatMap((channel) => {
        const value = edit[channel];
        if (typeof value === "number" && value !== 0) {
          return [`${channel} ${signed(value)}`];
        }
        return value === true ? [channel] : [];
      }).join(", ");
    case "resize":
      return edit.resample === "nearest" ? "nearest neighbour" : "bilinear";
    case "annotate":
      return edit.shape.shape === "text" ? `“${edit.shape.text}”` : "";
    default:
      return "";
  }
}
