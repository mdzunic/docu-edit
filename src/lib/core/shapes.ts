// The product-neutral half of SPEC-051's drawing shapes: the closed
// geometry catalogue, the visual properties a shape paints with, and the
// pure draw/resize/nudge arithmetic. Everything here is independent of the
// document model — a Slides product places the same shapes on a slide
// instead of anchoring them to a paragraph. The document-specific half
// (anchoring, wrap, z-order, grouping, the editing commands) stays in
// lib/shapes.ts and lib/documentModel.ts.

/** The closed set of geometries the editor draws, resizes and round-trips
 * natively — see "Shape catalogue" in specs/051-drawing-shapes-and-text-boxes.md.
 * Anything else arriving from a file is preserved per SPEC-008, not
 * approximated into one of these. */
export type ShapeGeometry =
  | "line"
  | "arrow"
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "diamond"
  | "star5"
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "wedgeRectCallout"
  | "wedgeRoundRectCallout";

/** A shape's outline (border) — colour, width in points and dash style. */
export interface ShapeOutline {
  color: string;
  widthPt: number;
  dash: "solid" | "dash" | "dot";
}

/** Typed pass-through for a shape geometry coming from a `.vue` template's
 * plain string (espree bars the annotation in `.vue` scripts). */
export function asShapeGeometry(value: string): ShapeGeometry {
  return value as ShapeGeometry;
}

/** The closed catalogue geometries, in "Shape catalogue" order — the gallery
 * (AC-1) and the docx/odf readers' geometry lookup both walk this list. */
export const SHAPE_CATALOGUE: { geometry: ShapeGeometry; group: "Lines" | "Basic shapes" | "Arrows" | "Callouts"; label: string }[] = [
  { geometry: "line", group: "Lines", label: "Line" },
  { geometry: "arrow", group: "Lines", label: "Arrow" },
  { geometry: "rect", group: "Basic shapes", label: "Rectangle" },
  { geometry: "roundRect", group: "Basic shapes", label: "Rounded rectangle" },
  { geometry: "ellipse", group: "Basic shapes", label: "Ellipse" },
  { geometry: "triangle", group: "Basic shapes", label: "Triangle" },
  { geometry: "diamond", group: "Basic shapes", label: "Diamond" },
  { geometry: "star5", group: "Basic shapes", label: "Star (5-point)" },
  { geometry: "rightArrow", group: "Arrows", label: "Right arrow" },
  { geometry: "leftArrow", group: "Arrows", label: "Left arrow" },
  { geometry: "upArrow", group: "Arrows", label: "Up arrow" },
  { geometry: "downArrow", group: "Arrows", label: "Down arrow" },
  { geometry: "wedgeRectCallout", group: "Callouts", label: "Rectangular callout" },
  { geometry: "wedgeRoundRectCallout", group: "Callouts", label: "Rounded callout" },
];

/** Geometries with zero area, drawn as a segment rather than a filled/outlined
 * region (SPEC-051 AC-6, AC-8, "Selection and hit testing"). */
export const LINE_GEOMETRIES: ReadonlySet<ShapeGeometry> = new Set(["line", "arrow"]);

/** What painting a shape needs — the visual subset of the document model's
 * `Shape` block (which extends this with anchoring/wrap/text). Any host
 * (page canvas, slide canvas, gallery icon) can paint one of these. */
export interface ShapeVisual {
  geometry: ShapeGeometry;
  /** "#rrggbb"; absent means no fill. */
  fill?: string;
  /** Absent means no outline. */
  outline?: ShapeOutline;
  /** Whole degrees clockwise, 0-359; absent means 0. */
  rotationDeg?: number;
  /** Present when the source geometry was outside the catalogue (SPEC-008
   * preservation, AC-60): the shape paints as a dashed placeholder at its
   * true bounds instead of `geometry`. */
  unsupportedGeometryXml?: string;
}

// --- drawing defaults and gesture arithmetic (SPEC-051 AC-4..AC-8) ----------

/** The default click-placed size for an area geometry (AC-5): 3cm x 2cm. */
export const DEFAULT_AREA_SIZE_CM = { widthCm: 3, heightCm: 2 };

/** The default click-placed size for a Line/Arrow (AC-6): a horizontal 3cm
 * segment — a box with zero height (see shapeGeometry.ts's lineSegment). */
export const DEFAULT_LINE_SIZE_CM = { widthCm: 3, heightCm: 0 };

/** The size draw mode uses for a single click with `geometry` chosen
 * (AC-5, AC-6). */
export function defaultClickSizeCm(geometry: ShapeGeometry): { widthCm: number; heightCm: number } {
  return LINE_GEOMETRIES.has(geometry) ? DEFAULT_LINE_SIZE_CM : DEFAULT_AREA_SIZE_CM;
}

/** The fill a gallery-drawn area shape starts with — nothing in the spec
 * mandates a colour (AC-45 only pins alt text absent), but a shape with no
 * fill and no outline paints nothing at all, which is not a usable starting
 * point for a shape someone just drew. Matches the accent blue Word/Writer
 * both default new shapes to. Never applied to a text box (AC-9 requires
 * fill/outline absent) or to a Line/Arrow (no area to fill). */
export const DEFAULT_NEW_SHAPE_FILL = "#4472c4";

/** The outline a gallery-drawn shape (area or Line/Arrow) starts with —
 * same reasoning as DEFAULT_NEW_SHAPE_FILL. */
export const DEFAULT_NEW_SHAPE_OUTLINE: ShapeOutline = { color: "#2e5395", widthPt: 1, dash: "solid" };

/** Shift-constrained drag size (AC-7, AC-8): an area geometry snaps to a
 * square (a circle for the ellipse — same "square bounding box" rule, the
 * renderer inscribes the ellipse in it); a Line/Arrow snaps its own
 * direction to the nearest 45-degree increment, i.e. its bounding box
 * becomes a square (45-degree diagonal) or collapses one axis to 0
 * (horizontal/vertical), whichever the dragged ratio is closer to. */
export function constrainedDragSizeCm(
  geometry: ShapeGeometry,
  widthCm: number,
  heightCm: number,
): { widthCm: number; heightCm: number } {
  const w = Math.abs(widthCm);
  const h = Math.abs(heightCm);
  if (LINE_GEOMETRIES.has(geometry)) {
    const angle = Math.atan2(h, w);
    const step = Math.PI / 4;
    const snapped = Math.round(angle / step) * step;
    if (snapped === 0 || snapped === Math.PI / 2) {
      return snapped === 0 ? { widthCm: w, heightCm: 0 } : { widthCm: 0, heightCm: h };
    }
    const side = Math.max(w, h);
    return { widthCm: side, heightCm: side };
  }
  const side = Math.max(w, h);
  return { widthCm: side, heightCm: side };
}

/** How far the pointer must travel before a draw gesture counts as a drag
 * rather than a click (AC-4, AC-5) — 3px at 96dpi. */
export const DRAW_DRAG_THRESHOLD_CM = 3 * (2.54 / 96);

/** The box a draw gesture currently describes (AC-4..AC-8), all in
 * page-relative cm: the rectangle dragged from `start` to `pointer`, or —
 * while the pointer has barely moved, and always for a text box — the
 * default click size placed at the pointer.
 *
 * The live preview and the shape that actually lands both come from here.
 * They have to: a preview drawn from its own copy of this arithmetic is
 * free to promise a box the insertion then does not deliver. */
export function drawGestureBox(
  geometry: ShapeGeometry,
  start: { xCm: number; yCm: number },
  pointer: { xCm: number; yCm: number },
  opts: { shiftKey?: boolean; forceDefaultSize?: boolean } = {},
): { leftCm: number; topCm: number; widthCm: number; heightCm: number } {
  const dxCm = pointer.xCm - start.xCm;
  const dyCm = pointer.yCm - start.yCm;
  const dragged =
    Math.abs(dxCm) > DRAW_DRAG_THRESHOLD_CM || Math.abs(dyCm) > DRAW_DRAG_THRESHOLD_CM;
  if (!dragged || opts.forceDefaultSize) {
    return { leftCm: pointer.xCm, topCm: pointer.yCm, ...defaultClickSizeCm(geometry) };
  }
  const size = opts.shiftKey
    ? constrainedDragSizeCm(geometry, dxCm, dyCm)
    : { widthCm: Math.abs(dxCm), heightCm: Math.abs(dyCm) };
  return {
    // Dragging up or left puts the box's own corner at the pointer.
    leftCm: dxCm < 0 ? pointer.xCm : start.xCm,
    topCm: dyCm < 0 ? pointer.yCm : start.yCm,
    ...size,
  };
}

/** The smallest display size a resize may leave a shape (resize-handle drag
 * on release, and the shape toolbar's size fields). Clamped like
 * MIN_IMAGE_SIZE_CM so a drag can never collapse a shape to nothing (a
 * zero-height Line stays allowed — that is a real horizontal line, not a
 * collapsed shape). */
export const MIN_SHAPE_SIZE_CM = 0.2;

/** 1pt in cm — arrow-key nudge distance (AC-40). 1pt = 1/72in = 2.54/72cm. */
export const NUDGE_PT_CM = 2.54 / 72;
/** 0.1pt in cm — Ctrl/Cmd+arrow fine nudge distance (AC-41). */
export const NUDGE_FINE_PT_CM = NUDGE_PT_CM / 10;

export type NudgeDirection = "up" | "down" | "left" | "right";

/** Typed pass-through for a nudge direction coming from a `.vue` template's
 * plain string (espree bars the annotation in `.vue` scripts). */
export function asNudgeDirection(value: string): NudgeDirection {
  return value as NudgeDirection;
}

/** Rotation in whole degrees, wrapped into 0..359 (AC-16) — a negative
 * or >=360 input (a drag that overshoots, or the toolbar field) wraps
 * around rather than clamping, matching Word/Writer's own rotation field. */
export function normalizeRotationDeg(deg: number): number {
  const whole = Math.round(deg);
  return ((whole % 360) + 360) % 360;
}
