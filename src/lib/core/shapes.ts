// The product-neutral half of SPEC-051's drawing shapes: the closed
// geometry catalogue, the visual properties a shape paints with, and the
// pure draw/resize/nudge arithmetic. Everything here is independent of the
// document model — a Slides product places the same shapes on a slide
// instead of anchoring them to a paragraph. The document-specific half
// (anchoring, wrap, z-order, grouping, the editing commands) stays in
// lib/shapes.ts and lib/documentModel.ts.
import { SHAPE_DEFINITIONS } from "./shapeCatalogue";
import type { ShapePath } from "./shapePath";
import { applyTransform, hexToRgb, rgbToHex } from "./theme";

/** The closed set of geometries the editor draws, resizes and round-trips
 * natively — the "expanded catalogue" of specs/069-shape-catalogue-expansion.md,
 * superseding SPEC-051's fourteen. Anything still outside this set arriving
 * from a file is preserved per SPEC-008, not approximated into one of these.
 *
 * The single source of truth for each geometry's group, label, names,
 * outline, adjustments and text rect is `SHAPE_DEFINITIONS` in
 * shapeCatalogue.ts (SPEC-069 AC-15a); this union is only the closed set of
 * keys those definitions may use. */
export type ShapeGeometry =
  // Lines
  | "line"
  | "arrow"
  // Basic shapes
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "diamond"
  | "star5"
  | "parallelogram"
  | "trapezoid"
  | "pentagon"
  | "hexagon"
  | "octagon"
  | "plus"
  | "can"
  | "cube"
  | "heart"
  | "cloud"
  | "donut"
  | "chord"
  | "plaque"
  | "frame"
  // Block arrows
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "leftRightArrow"
  | "upDownArrow"
  | "quadArrow"
  | "homePlate"
  | "chevron"
  | "notchedRightArrow"
  | "stripedRightArrow"
  // Flowchart
  | "flowChartProcess"
  | "flowChartAlternateProcess"
  | "flowChartDecision"
  | "flowChartInputOutput"
  | "flowChartPredefinedProcess"
  | "flowChartInternalStorage"
  | "flowChartDocument"
  | "flowChartMultidocument"
  | "flowChartTerminator"
  | "flowChartPreparation"
  | "flowChartManualInput"
  | "flowChartManualOperation"
  | "flowChartConnector"
  | "flowChartOffpageConnector"
  | "flowChartPunchedCard"
  | "flowChartMagneticDisk"
  | "flowChartDelay"
  | "flowChartExtract"
  | "flowChartMerge"
  | "flowChartSort"
  // Stars and banners
  | "star4"
  | "star6"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  | "irregularSeal1"
  | "irregularSeal2"
  | "ribbon"
  | "ribbon2"
  | "verticalScroll"
  | "horizontalScroll"
  | "wave"
  | "doubleWave"
  // Callouts
  | "wedgeRectCallout"
  | "wedgeRoundRectCallout"
  | "wedgeEllipseCallout"
  | "cloudCallout"
  | "borderCallout1"
  | "borderCallout2"
  | "borderCallout3"
  | "accentCallout1";

/** The catalogue's groups, in gallery order (SPEC-069 "The expanded
 * catalogue"). */
export type ShapeGroup =
  | "Lines"
  | "Basic shapes"
  | "Block arrows"
  | "Flowchart"
  | "Stars and banners"
  | "Callouts";

/** A line-end decoration (SPEC-070 "Arrowheads"). `.docx` writes it as
 * `a:headEnd`/`a:tailEnd @type`, `.odt` as `draw:marker-start`/
 * `draw:marker-end` referencing a standard marker name. Introduced for
 * connectors; SPEC-073 extends it to plain lines and shape outlines using
 * this exact closed union — no new kinds (D1/AC-13). */
export type ArrowheadKind = "none" | "triangle" | "open" | "stealth" | "diamond" | "oval";

/** Arrowhead size (SPEC-073 D1). "medium" is the default a control applies
 * when turning an arrowhead on with no size chosen (AC-14). */
export type ArrowheadSize = "small" | "medium" | "large";

/** One end's arrowhead on a shape outline (SPEC-073). Absent means no
 * arrowhead (AC-15); connectors keep SPEC-070's own end-arrow default. */
export interface ShapeArrowhead {
  kind: ArrowheadKind;
  size: ArrowheadSize;
}

/** How much each arrowhead size scales the stroke width the head geometry is
 * derived from (core/connectors.ts's arrowheadOutline ratios). */
export const ARROWHEAD_SIZE_FACTORS: Record<ArrowheadSize, number> = {
  small: 0.75,
  medium: 1,
  large: 1.5,
};

/** A shape's outline (border). SPEC-051's flat colour/width/dash widened by
 * SPEC-073 ("The outline model"): the three original dash values keep their
 * spelling and meaning, the union only extends (AC-2). `alpha` follows D5
 * (0..1 opacity, absent = opaque). */
export interface ShapeOutline {
  color: string;
  widthPt: number;
  dash: "solid" | "dash" | "dot" | "dashDot" | "longDash" | "longDashDot";
  alpha?: number;
  /** Absent = miter, both formats' default. */
  join?: "round" | "bevel" | "miter";
  /** Absent = flat. */
  cap?: "flat" | "round" | "square";
  compound?: "single" | "double" | "thickThin" | "thinThick";
  startArrow?: ShapeArrowhead;
  endArrow?: ShapeArrowhead;
}

/** A gradient fill's definition (SPEC-073 "The fill model"). */
export interface ShapeGradient {
  type: "linear" | "radial";
  /** Whole degrees clockwise from "left to right"; linear only. */
  angleDeg?: number;
  /** Two or more, ordered by position, first at 0 and last at 1. */
  stops: Array<{ position: number; color: string; alpha?: number }>;
}

/** A shape's fill (SPEC-073 AC-1): the full union replacing SPEC-051's flat
 * `"#rrggbb"` string. `alpha` is 0..1 opacity per D5. A picture fill's
 * `mediaId` references the same media store `Image` blocks use (SPEC-024),
 * so the two share one part and one MediaRecord (AC-4). */
export type ShapeFill =
  | { kind: "none" }
  | { kind: "solid"; color: string; alpha?: number }
  | { kind: "gradient"; gradient: ShapeGradient }
  | { kind: "picture"; mediaId: string; mode: "stretch" | "tile"; alpha?: number }
  | { kind: "pattern"; preset: string; foreground: string; background: string };

/** An outer shadow (SPEC-073 AC-8): colour at `alpha` opacity, blurred by
 * `blurPt`, offset `distancePt` towards `angleDeg` (degrees clockwise, 0 =
 * right, 90 = down). */
export interface ShapeShadow {
  color: string;
  alpha: number;
  blurPt: number;
  distancePt: number;
  angleDeg: number;
}

/** The "Lighter 40%" tint of a #rrggbb colour (lumMod 0.6 then lumOff 0.4)
 * — SPEC-073 D2, reusing core/theme.ts's pure helpers, no new maths. */
export function lighterTint(hex: string): string {
  const l = applyTransform(hexToRgb(hex), "lumMod", 0.6);
  return rgbToHex(applyTransform(l, "lumOff", 0.4));
}

/** The "Lighter 80%" tint (lumMod 0.2 then lumOff 0.8) — the subtleFilled
 * preset's fill (D3). */
function lighterTint80(hex: string): string {
  const l = applyTransform(hexToRgb(hex), "lumMod", 0.2);
  return rgbToHex(applyTransform(l, "lumOff", 0.8));
}

/** Accepts the legacy `"#rrggbb"` string form and the `ShapeFill` union;
 * returns the union (SPEC-073 AC-30, "Migration"). Anything else — including
 * a malformed object — is `undefined` (no fill). Called on the document-store
 * read path, not scattered through render code. */
export function normaliseShapeFill(value: unknown): ShapeFill | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value === "string") {
    return /^#[0-9a-fA-F]{6}$/.test(value) ? { kind: "solid", color: value } : undefined;
  }
  if (typeof value === "object") {
    const kind = (value as { kind?: unknown }).kind;
    if (kind === "none" || kind === "solid" || kind === "gradient" || kind === "picture" || kind === "pattern") {
      return value as ShapeFill;
    }
  }
  return undefined;
}

/** Typed `null` seed for the picture-fill media resolver a host injects into
 * the shape block component (espree bars the annotation in `.vue` scripts —
 * the noImageLocation pattern). */
export function noShapeMediaSrc(): ((mediaId: string) => string) | null {
  return null;
}

/** Whether a fill paints anything at all — `none` and a malformed union
 * member do not; every other kind does. Hit-testing and clip-path logic key
 * off this the way they used to key off a truthy colour string. */
export function shapeFillIsVisible(fill: ShapeFill | undefined): boolean {
  return fill !== undefined && fill.kind !== "none";
}

/** The single flat colour a fill reduces to where only one colour can be
 * painted (PDF text extraction, library thumbnails): solid → its colour,
 * gradient → its first stop, pattern → its foreground; none/picture have no
 * meaningful flat colour. */
export function shapeFillBaseColor(fill: ShapeFill | undefined): string | undefined {
  if (!fill) {
    return undefined;
  }
  if (fill.kind === "solid") {
    return fill.color;
  }
  if (fill.kind === "gradient") {
    return fill.gradient.stops[0]?.color;
  }
  if (fill.kind === "pattern") {
    return fill.foreground;
  }
  return undefined;
}

/** The seeded gradient a fill control builds when switching to "gradient"
 * with no gradient yet (SPEC-073 AC-22, D2): two linear stops at 90°, the
 * base colour to its lighter tint. */
export function seededGradient(base: string): ShapeGradient {
  return {
    type: "linear",
    angleDeg: 90,
    stops: [
      { position: 0, color: base },
      { position: 1, color: lighterTint(base) },
    ],
  };
}

/** Typed pass-through for a shape geometry coming from a `.vue` template's
 * plain string (espree bars the annotation in `.vue` scripts). */
export function asShapeGeometry(value: string): ShapeGeometry {
  return value as ShapeGeometry;
}

/** The closed catalogue geometries, in "The expanded catalogue" order — the
 * gallery (AC-2) and the docx/odf readers' geometry lookup both walk this
 * list. SPEC-069 AC-15a: derived from the one `SHAPE_DEFINITIONS` table
 * rather than maintained beside it, so a new geometry is a single entry
 * there. The projected shape (geometry/group/label) is unchanged, so every
 * existing importer and the shipped shape tests keep working. */
export const SHAPE_CATALOGUE: { geometry: ShapeGeometry; group: ShapeGroup; label: string }[] =
  SHAPE_DEFINITIONS.map((def) => ({ geometry: def.geometry, group: def.group, label: def.label }));

/** Geometries with zero area, drawn as a segment rather than a filled/outlined
 * region (SPEC-051 AC-6, AC-8, "Selection and hit testing"). Derived from
 * the `isLine` flag on `SHAPE_DEFINITIONS` (SPEC-069 AC-15a). */
export const LINE_GEOMETRIES: ReadonlySet<ShapeGeometry> = new Set(
  SHAPE_DEFINITIONS.filter((def) => def.isLine).map((def) => def.geometry),
);

/** What painting a shape needs — the visual subset of the document model's
 * `Shape` block (which extends this with anchoring/wrap/text). Any host
 * (page canvas, slide canvas, gallery icon) can paint one of these. */
export interface ShapeVisual {
  geometry: ShapeGeometry;
  /** Absent means no fill (SPEC-073 widened this from a "#rrggbb" string). */
  fill?: ShapeFill;
  /** Absent means no outline. */
  outline?: ShapeOutline;
  /** Outer shadow (SPEC-073); absent means none — the shadow pass is then
   * skipped entirely (AC-40). */
  shadow?: ShapeShadow;
  /** Whole degrees clockwise, 0-359; absent means 0. */
  rotationDeg?: number;
  /** Geometry adjustment values (SPEC-069), each a 0..1-style fraction in
   * the order this geometry's `ShapeDefinition` declares them. Absent means
   * every adjustment is at its default. Consumed by the definition's
   * `outline`/`textRect` and clamped for rendering (AC-7a). */
  adjustments?: readonly number[];
  /** Present when the source geometry was outside the catalogue (SPEC-008
   * preservation, AC-60): the shape paints as a dashed placeholder at its
   * true bounds instead of `geometry`. */
  unsupportedGeometryXml?: string;
  /** A custom outline (SPEC-071). When present it is painted instead of
   * `geometry` — its real Bézier/line outline, never the placeholder. */
  path?: ShapePath;
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
export const DEFAULT_NEW_SHAPE_FILL: ShapeFill = { kind: "solid", color: "#4472c4" };

/** The accent colour the defaults and the D3 style presets derive from —
 * `DEFAULT_NEW_SHAPE_FILL`'s own colour (SPEC-077 will swap this for a theme
 * colour; until then it is this literal). */
const ACCENT = DEFAULT_NEW_SHAPE_FILL.color;

/** The outline a gallery-drawn shape (area or Line/Arrow) starts with —
 * same reasoning as DEFAULT_NEW_SHAPE_FILL. */
export const DEFAULT_NEW_SHAPE_OUTLINE: ShapeOutline = { color: "#2e5395", widthPt: 1, dash: "solid" };

/** The standard outer shadow (SPEC-073 D4): black at 40% opacity, 4pt blur,
 * offset 3pt downward — Office's "Offset: Bottom". The shadow toggle applies
 * it (AC-20) and the `outlinedShadow` preset carries it. */
export const DEFAULT_SHAPE_SHADOW: ShapeShadow = {
  color: "#000000",
  alpha: 0.4,
  blurPt: 4,
  distancePt: 3,
  angleDeg: 90,
};

/** The nine D4 shadow presets (AC-21): eight directional variants of the
 * standard shadow differing only in `angleDeg`, plus "No shadow" clearing it. */
export const SHAPE_SHADOW_PRESETS: { id: string; label: string; shadow?: ShapeShadow }[] = [
  { id: "none", label: "No shadow" },
  { id: "bottom", label: "Bottom", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 90 } },
  { id: "bottomRight", label: "Bottom right", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 45 } },
  { id: "right", label: "Right", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 0 } },
  { id: "topRight", label: "Top right", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 315 } },
  { id: "top", label: "Top", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 270 } },
  { id: "topLeft", label: "Top left", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 225 } },
  { id: "left", label: "Left", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 180 } },
  { id: "bottomLeft", label: "Bottom left", shadow: { ...DEFAULT_SHAPE_SHADOW, angleDeg: 135 } },
];

/** One of the five style presets (SPEC-073 D3): a coordinated (fill,
 * outline, shadow) triple applied in one click and one undo entry (AC-18).
 * An absent member clears the property. */
export interface ShapeStylePreset {
  id: "filled" | "subtleFilled" | "outlined" | "outlinedShadow" | "noFill";
  label: string;
  fill?: ShapeFill;
  outline?: ShapeOutline;
  shadow?: ShapeShadow;
}

/** The exactly-five preset gallery (AC-18/AC-19), each colour derived from
 * the accent (`DEFAULT_NEW_SHAPE_FILL`) via core/theme.ts transforms — a
 * single source, ready for SPEC-077's theme accent. With today's literal
 * accent `#4472c4` these resolve to D3's table: subtleFilled's fill is the
 * accent at Lighter 80% (`#dae3f3`). */
export const SHAPE_STYLE_PRESETS: ShapeStylePreset[] = [
  {
    id: "filled",
    label: "Filled",
    fill: { kind: "solid", color: ACCENT },
    outline: { ...DEFAULT_NEW_SHAPE_OUTLINE },
  },
  {
    id: "subtleFilled",
    label: "Subtle fill",
    fill: { kind: "solid", color: lighterTint80(ACCENT) },
    outline: { color: ACCENT, widthPt: 1, dash: "solid" },
  },
  {
    id: "outlined",
    label: "Outlined",
    outline: { color: ACCENT, widthPt: 1.5, dash: "solid" },
  },
  {
    id: "outlinedShadow",
    label: "Outlined + shadow",
    outline: { color: ACCENT, widthPt: 1.5, dash: "solid" },
    shadow: { ...DEFAULT_SHAPE_SHADOW },
  },
  { id: "noFill", label: "No fill" },
];

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
