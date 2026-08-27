// SPEC-069: the single source of truth for every catalogue geometry. One
// `ShapeDefinition` per geometry owns its group, label, gallery aliases, the
// OOXML `prst` and ODF `draw:type` names, its outline in a unit square, its
// declared adjustments, and where text lays out inside it. Everything else —
// `SHAPE_CATALOGUE`, `shapePolygon`, the `PRST_*`/`GEOMETRY_*` name maps and
// the ODF reverse map — is derived from this table rather than maintained
// beside it (AC-15a), so adding a geometry is a single entry here.
//
// Product-neutral core (boundary.spec.ts): no document model, no Vue, no
// docx/odf reader. Type-only imports of `ShapeGeometry`/`ShapeGroup` from
// shapes.ts and `Point` from shapeGeometry.ts are erased at build, so this
// module has no runtime dependency on either and no import cycle.
import type { ShapeGeometry, ShapeGroup } from "./shapes";
import type { Point } from "./shapeGeometry";

/** An outline in a unit square (0..1 of width/height): a closed polygon of
 * vertices. Geometries that curve return a densely-sampled polygon, so the
 * same list paints, hit-tests, clips and draws the gallery icon (AC-3a) and
 * `shapePolygon` keeps returning `Point[]`. */
export type ShapeOutlinePath = Point[];

/** One declared adjustment: its accessible label, default and clamp range
 * (all 0..1-style fractions in the geometry's own convention), and the axis
 * a drag handle moves it along. */
export interface ShapeAdjustmentSpec {
  label: string;
  default: number;
  min: number;
  max: number;
  axis: "x" | "y" | "radial";
}

/** Where text lays out inside a shape, as unit-square fractions. */
export interface ShapeTextRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ShapeDefinition {
  geometry: ShapeGeometry;
  group: ShapeGroup;
  label: string;
  /** Extra search terms for the gallery's filter (AC-2b). */
  aliases?: string[];
  /** OOXML preset name (`a:prstGeom/@prst`). */
  prst: string;
  /** ODF enhanced-geometry type (`draw:type`). */
  drawType: string;
  /** Outline in a unit square, given the shape's (already resolved)
   * adjustments — see `resolvedShapeAdjustments`. */
  outline(adjustments: readonly number[]): ShapeOutlinePath;
  /** Extra fill-only sub-regions punched out of the outline (donut hole,
   * picture frame): each a closed unit-square ring painted with the
   * even-odd rule. Absent for a solid shape. */
  holes?(adjustments: readonly number[]): ShapeOutlinePath[];
  /** Declared adjustments, in the order of `Shape.adjustments` and the
   * format's own adjustment list. */
  adjustments?: ShapeAdjustmentSpec[];
  /** Where text lays out; defaults to the full box. */
  textRect?(adjustments: readonly number[]): ShapeTextRect;
  /** True for a zero-area geometry drawn as a segment (line/arrow). */
  isLine?: boolean;
  /** The format(s) whose closest preset only approximates this geometry, so
   * a cross-format save to that format writes the preset AND reports a
   * `shapeGeometry` DropRecord (AC-9). Absent means both formats carry it. */
  approximateIn?: ("docx" | "odt")[];
}

// --- unit-square geometry helpers -------------------------------------------

const TAU = Math.PI * 2;

function clamp01(value: number): number {
  return Math.min(1, Math.max(0, value));
}

/** Rounds off floating dust (e.g. -1e-17) and clamps a sampled point into
 * the unit square so the "within bounds" invariant (AC-12b) holds exactly. */
function bound(p: Point): Point {
  return { x: clamp01(p.x), y: clamp01(p.y) };
}

function deg2rad(deg: number): number {
  return (deg * Math.PI) / 180;
}

/** Points evenly spaced round an ellipse, y measured downward (screen
 * convention): angle 0 is the right, 90 the bottom. */
function ringPoints(cx: number, cy: number, rx: number, ry: number, steps = 40, startDeg = 0): Point[] {
  const pts: Point[] = [];
  for (let i = 0; i < steps; i++) {
    const a = deg2rad(startDeg + (360 * i) / steps);
    pts.push(bound({ x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) }));
  }
  return pts;
}

/** Points along an elliptical arc from `startDeg` to `endDeg` inclusive. */
function arcPoints(cx: number, cy: number, rx: number, ry: number, startDeg: number, endDeg: number, steps = 12): Point[] {
  const pts: Point[] = [];
  for (let i = 0; i <= steps; i++) {
    const t = startDeg + ((endDeg - startDeg) * i) / steps;
    const a = deg2rad(t);
    pts.push(bound({ x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) }));
  }
  return pts;
}

/** A regular n-gon inscribed in the unit square, first vertex at `startDeg`
 * (default straight up). */
function regularPoly(n: number, startDeg = -90): Point[] {
  const pts: Point[] = [];
  for (let i = 0; i < n; i++) {
    const a = deg2rad(startDeg) + (i * TAU) / n;
    pts.push(bound({ x: 0.5 + 0.5 * Math.cos(a), y: 0.5 + 0.5 * Math.sin(a) }));
  }
  return pts;
}

/** An n-point star: `n` outer points at radius 0.5, `n` inner at
 * `0.5 * innerRatio`, first outer point at `startDeg` (default straight up). */
function starPoly(n: number, innerRatio: number, startDeg = -90): Point[] {
  const pts: Point[] = [];
  const outerR = 0.5;
  const innerR = 0.5 * innerRatio;
  for (let i = 0; i < 2 * n; i++) {
    const r = i % 2 === 0 ? outerR : innerR;
    const a = deg2rad(startDeg) + (i * Math.PI) / n;
    pts.push(bound({ x: 0.5 + r * Math.cos(a), y: 0.5 + r * Math.sin(a) }));
  }
  return pts;
}

/** Rescales an arbitrary point cloud to exactly fill the unit square,
 * preserving shape — used for parametric outlines (heart) authored in their
 * own coordinate space. */
function normalizeToBox(raw: Point[]): Point[] {
  const xs = raw.map((p) => p.x);
  const ys = raw.map((p) => p.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const w = maxX - minX || 1;
  const h = maxY - minY || 1;
  return raw.map((p) => bound({ x: (p.x - minX) / w, y: (p.y - minY) / h }));
}

const rectPoly = (): Point[] => [
  { x: 0, y: 0 },
  { x: 1, y: 0 },
  { x: 1, y: 1 },
  { x: 0, y: 1 },
];

/** The 10-vertex 5-point star (SPEC-051) — reproduced exactly so
 * `shapePolygon("star5")` is byte-for-byte what the shipped tests pin. */
function star5Points(): Point[] {
  const pts: Point[] = [];
  const outerR = 0.5;
  const innerR = outerR * 0.382;
  for (let i = 0; i < 10; i++) {
    const r = i % 2 === 0 ? outerR : innerR;
    const angle = -Math.PI / 2 + (i * TAU) / 10;
    pts.push({ x: 0.5 + r * Math.cos(angle), y: 0.5 + r * Math.sin(angle) });
  }
  return pts;
}

// The two SPEC-051 callouts, reproduced so shapePolygon() stays exactly the
// straight-edged 7-vertex approximation roundedCalloutRadius/Outline (still
// in shapeGeometry.ts, still pinned by tests) are measured against.
const CALLOUT_BODY = 0.75;
const WEDGE_TAIL = { left: 0.05, tip: 0.15, right: 0.35 };
const ROUND_TAIL = { left: 0.25, tip: 0.35, right: 0.55 };

function calloutPoly(feetLeft: number, tipX: number, feetRight: number, tipY: number): Point[] {
  return [
    { x: 0, y: 0 },
    { x: 1, y: 0 },
    { x: 1, y: CALLOUT_BODY },
    { x: feetRight, y: CALLOUT_BODY },
    { x: tipX, y: tipY },
    { x: feetLeft, y: CALLOUT_BODY },
    { x: 0, y: CALLOUT_BODY },
  ];
}

const CALLOUT_TAIL_ADJ: ShapeAdjustmentSpec[] = [
  { label: "tail x", default: 0, min: 0, max: 1, axis: "x" },
  { label: "tail y", default: 1, min: CALLOUT_BODY, max: 1, axis: "y" },
];

const calloutTextRect = (): ShapeTextRect => ({ x: 0, y: 0, w: 1, h: CALLOUT_BODY });

// --- block-arrow builders ---------------------------------------------------

/** A single-headed straight arrow pointing right, parametrised by where the
 * head begins (`neck`) and the shaft's half-thickness (`half`). */
function rightArrowPoly(neck: number, half: number): Point[] {
  return [
    { x: 0, y: 0.5 - half },
    { x: neck, y: 0.5 - half },
    { x: neck, y: 0 },
    { x: 1, y: 0.5 },
    { x: neck, y: 1 },
    { x: neck, y: 0.5 + half },
    { x: 0, y: 0.5 + half },
  ];
}

const ARROW_ADJ: ShapeAdjustmentSpec[] = [
  { label: "head length", default: 0.4, min: 0.1, max: 0.9, axis: "x" },
  { label: "shaft thickness", default: 0.2, min: 0.05, max: 0.45, axis: "y" },
];

function mirrorX(pts: Point[]): Point[] {
  return pts.map((p) => ({ x: 1 - p.x, y: p.y }));
}

function swapAxes(pts: Point[]): Point[] {
  return pts.map((p) => ({ x: p.y, y: p.x }));
}

// ============================================================================
// Definitions, grouped as in "The expanded catalogue".
// ============================================================================

const LINES: ShapeDefinition[] = [
  { geometry: "line", group: "Lines", label: "Line", prst: "line", drawType: "line", isLine: true, outline: rectPoly },
  { geometry: "arrow", group: "Lines", label: "Arrow", aliases: ["straight"], prst: "straightConnector1", drawType: "line", isLine: true, outline: rectPoly },
];

const BASIC: ShapeDefinition[] = [
  { geometry: "rect", group: "Basic shapes", label: "Rectangle", aliases: ["box", "square"], prst: "rect", drawType: "rectangle", outline: rectPoly },
  {
    geometry: "roundRect",
    group: "Basic shapes",
    label: "Rounded rectangle",
    aliases: ["box"],
    prst: "roundRect",
    drawType: "round-rectangle",
    // Hit polygon stays the rectangle (the rounding is cosmetic, painted by
    // the canvas from the corner-radius adjustment) — as SPEC-051 pinned.
    outline: rectPoly,
    adjustments: [{ label: "corner radius", default: 0.1667, min: 0, max: 0.5, axis: "x" }],
  },
  { geometry: "ellipse", group: "Basic shapes", label: "Ellipse", aliases: ["circle", "oval"], prst: "ellipse", drawType: "ellipse", outline: () => ringPoints(0.5, 0.5, 0.5, 0.5, 48) },
  { geometry: "triangle", group: "Basic shapes", label: "Triangle", aliases: ["isosceles"], prst: "triangle", drawType: "isosceles-triangle", outline: () => [{ x: 0.5, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }] },
  { geometry: "diamond", group: "Basic shapes", label: "Diamond", aliases: ["rhombus"], prst: "diamond", drawType: "diamond", outline: () => [{ x: 0.5, y: 0 }, { x: 1, y: 0.5 }, { x: 0.5, y: 1 }, { x: 0, y: 0.5 }] },
  { geometry: "star5", group: "Basic shapes", label: "Star (5-point)", aliases: ["star"], prst: "star5", drawType: "star5", outline: star5Points },
  {
    geometry: "parallelogram",
    group: "Basic shapes",
    label: "Parallelogram",
    prst: "parallelogram",
    drawType: "parallelogram",
    adjustments: [{ label: "slant", default: 0.25, min: 0, max: 0.5, axis: "x" }],
    outline: (a) => {
      const s = a[0] ?? 0.25;
      return [{ x: s, y: 0 }, { x: 1, y: 0 }, { x: 1 - s, y: 1 }, { x: 0, y: 1 }];
    },
  },
  {
    geometry: "trapezoid",
    group: "Basic shapes",
    label: "Trapezoid",
    prst: "trapezoid",
    drawType: "trapezoid",
    adjustments: [{ label: "slant", default: 0.25, min: 0, max: 0.5, axis: "x" }],
    outline: (a) => {
      const s = a[0] ?? 0.25;
      return [{ x: s, y: 0 }, { x: 1 - s, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }];
    },
  },
  { geometry: "pentagon", group: "Basic shapes", label: "Pentagon", prst: "pentagon", drawType: "regular-pentagon", outline: () => regularPoly(5, -90) },
  { geometry: "hexagon", group: "Basic shapes", label: "Hexagon", prst: "hexagon", drawType: "hexagon", outline: () => [{ x: 0.25, y: 0 }, { x: 0.75, y: 0 }, { x: 1, y: 0.5 }, { x: 0.75, y: 1 }, { x: 0.25, y: 1 }, { x: 0, y: 0.5 }] },
  {
    geometry: "octagon",
    group: "Basic shapes",
    label: "Octagon",
    prst: "octagon",
    drawType: "octagon",
    outline: () => {
      const c = 1 / (2 + Math.SQRT2);
      return [
        { x: c, y: 0 }, { x: 1 - c, y: 0 }, { x: 1, y: c }, { x: 1, y: 1 - c },
        { x: 1 - c, y: 1 }, { x: c, y: 1 }, { x: 0, y: 1 - c }, { x: 0, y: c },
      ];
    },
  },
  {
    geometry: "plus",
    group: "Basic shapes",
    label: "Cross",
    aliases: ["plus"],
    prst: "plus",
    drawType: "cross",
    adjustments: [{ label: "thickness", default: 0.35, min: 0.1, max: 0.49, axis: "x" }],
    outline: (a) => {
      const t = a[0] ?? 0.35;
      return [
        { x: t, y: 0 }, { x: 1 - t, y: 0 }, { x: 1 - t, y: t }, { x: 1, y: t },
        { x: 1, y: 1 - t }, { x: 1 - t, y: 1 - t }, { x: 1 - t, y: 1 }, { x: t, y: 1 },
        { x: t, y: 1 - t }, { x: 0, y: 1 - t }, { x: 0, y: t }, { x: t, y: t },
      ];
    },
  },
  {
    geometry: "can",
    group: "Basic shapes",
    label: "Cylinder",
    aliases: ["can", "cylinder", "drum"],
    prst: "can",
    drawType: "can",
    adjustments: [{ label: "top height", default: 0.15, min: 0.05, max: 0.45, axis: "y" }],
    outline: (a) => {
      const e = a[0] ?? 0.15;
      return [...arcPoints(0.5, e, 0.5, e, 180, 360, 16), ...arcPoints(0.5, 1 - e, 0.5, e, 0, 180, 16)];
    },
    textRect: (a) => ({ x: 0, y: 2 * (a[0] ?? 0.15), w: 1, h: 1 - 3 * (a[0] ?? 0.15) }),
  },
  {
    geometry: "cube",
    group: "Basic shapes",
    label: "Cube",
    prst: "cube",
    drawType: "cube",
    adjustments: [{ label: "depth", default: 0.25, min: 0.05, max: 0.5, axis: "x" }],
    outline: (a) => {
      const d = a[0] ?? 0.25;
      return [{ x: 0, y: d }, { x: d, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 - d }, { x: 1 - d, y: 1 }, { x: 0, y: 1 }];
    },
    textRect: (a) => ({ x: 0, y: a[0] ?? 0.25, w: 1 - (a[0] ?? 0.25), h: 1 - (a[0] ?? 0.25) }),
  },
  {
    geometry: "heart",
    group: "Basic shapes",
    label: "Heart",
    prst: "heart",
    drawType: "heart",
    outline: () => {
      const raw: Point[] = [];
      const steps = 44;
      for (let i = 0; i < steps; i++) {
        const t = deg2rad((360 * i) / steps);
        const x = 16 * Math.pow(Math.sin(t), 3);
        const y = 13 * Math.cos(t) - 5 * Math.cos(2 * t) - 2 * Math.cos(3 * t) - Math.cos(4 * t);
        raw.push({ x, y: -y });
      }
      return normalizeToBox(raw);
    },
  },
  {
    geometry: "cloud",
    group: "Basic shapes",
    label: "Cloud",
    prst: "cloud",
    drawType: "cloud",
    outline: () => cloudRing(0.5, 0.5, 0.38, 0.34, 8),
  },
  {
    geometry: "donut",
    group: "Basic shapes",
    label: "Donut",
    aliases: ["ring"],
    prst: "donut",
    drawType: "ring",
    adjustments: [{ label: "hole size", default: 0.5, min: 0.1, max: 0.9, axis: "radial" }],
    outline: () => ringPoints(0.5, 0.5, 0.5, 0.5, 48),
    holes: (a) => {
      const r = 0.5 * (a[0] ?? 0.5);
      return [ringPoints(0.5, 0.5, r, r, 48)];
    },
  },
  {
    geometry: "chord",
    group: "Basic shapes",
    label: "Chord",
    prst: "chord",
    drawType: "chord",
    outline: () => arcPoints(0.5, 0.5, 0.5, 0.5, 45, 315, 28),
  },
  {
    geometry: "plaque",
    group: "Basic shapes",
    label: "Plaque",
    prst: "plaque",
    drawType: "plaque",
    outline: () => {
      const c = 0.2;
      return [
        { x: c, y: 0 }, { x: 1 - c, y: 0 },
        ...arcPoints(1, 0, c, c, 180, 90, 6),
        { x: 1, y: 1 - c },
        ...arcPoints(1, 1, c, c, 270, 180, 6),
        { x: c, y: 1 },
        ...arcPoints(0, 1, c, c, 0, -90, 6),
        { x: 0, y: c },
        ...arcPoints(0, 0, c, c, 90, 0, 6),
      ];
    },
  },
  {
    geometry: "frame",
    group: "Basic shapes",
    label: "Frame",
    prst: "frame",
    drawType: "frame",
    adjustments: [{ label: "frame width", default: 0.125, min: 0.02, max: 0.45, axis: "x" }],
    outline: rectPoly,
    holes: (a) => {
      const w = a[0] ?? 0.125;
      return [[{ x: w, y: w }, { x: 1 - w, y: w }, { x: 1 - w, y: 1 - w }, { x: w, y: 1 - w }]];
    },
    textRect: (a) => ({ x: 0, y: 0, w: 1, h: a[0] ?? 0.125 }),
  },
];

/** A scalloped ring (cloud silhouette): base ellipse radius modulated by a
 * small sinusoidal bump, kept inside the box. */
function cloudRing(cx: number, cy: number, rx: number, ry: number, bumps: number): Point[] {
  const pts: Point[] = [];
  const steps = bumps * 8;
  const amp = 0.11;
  for (let i = 0; i < steps; i++) {
    const a = (TAU * i) / steps;
    const scallop = 1 - amp + amp * Math.abs(Math.sin((bumps * a) / 2));
    pts.push(bound({ x: cx + rx * scallop * Math.cos(a), y: cy + ry * scallop * Math.sin(a) }));
  }
  return pts;
}

const BLOCK_ARROWS: ShapeDefinition[] = [
  { geometry: "rightArrow", group: "Block arrows", label: "Right arrow", prst: "rightArrow", drawType: "right-arrow", adjustments: ARROW_ADJ, outline: (a) => rightArrowPoly(1 - (a[0] ?? 0.4), a[1] ?? 0.2), textRect: (a) => ({ x: 0, y: 0.5 - (a[1] ?? 0.2), w: 1 - (a[0] ?? 0.4), h: 2 * (a[1] ?? 0.2) }) },
  { geometry: "leftArrow", group: "Block arrows", label: "Left arrow", prst: "leftArrow", drawType: "left-arrow", adjustments: ARROW_ADJ, outline: (a) => mirrorX(rightArrowPoly(1 - (a[0] ?? 0.4), a[1] ?? 0.2)) },
  { geometry: "upArrow", group: "Block arrows", label: "Up arrow", prst: "upArrow", drawType: "up-arrow", adjustments: ARROW_ADJ, outline: (a) => swapAxes(mirrorX(rightArrowPoly(1 - (a[0] ?? 0.4), a[1] ?? 0.2))) },
  { geometry: "downArrow", group: "Block arrows", label: "Down arrow", prst: "downArrow", drawType: "down-arrow", adjustments: ARROW_ADJ, outline: (a) => swapAxes(rightArrowPoly(1 - (a[0] ?? 0.4), a[1] ?? 0.2)) },
  {
    geometry: "leftRightArrow",
    group: "Block arrows",
    label: "Left-right arrow",
    prst: "leftRightArrow",
    drawType: "left-right-arrow",
    adjustments: [{ label: "head length", default: 0.25, min: 0.1, max: 0.45, axis: "x" }, { label: "shaft thickness", default: 0.15, min: 0.05, max: 0.45, axis: "y" }],
    outline: (a) => {
      const h = a[0] ?? 0.25;
      const t = a[1] ?? 0.15;
      return [
        { x: 0, y: 0.5 }, { x: h, y: 0 }, { x: h, y: 0.5 - t }, { x: 1 - h, y: 0.5 - t },
        { x: 1 - h, y: 0 }, { x: 1, y: 0.5 }, { x: 1 - h, y: 1 }, { x: 1 - h, y: 0.5 + t },
        { x: h, y: 0.5 + t }, { x: h, y: 1 },
      ];
    },
  },
  {
    geometry: "upDownArrow",
    group: "Block arrows",
    label: "Up-down arrow",
    prst: "upDownArrow",
    drawType: "up-down-arrow",
    adjustments: [{ label: "head length", default: 0.25, min: 0.1, max: 0.45, axis: "y" }, { label: "shaft thickness", default: 0.15, min: 0.05, max: 0.45, axis: "x" }],
    outline: (a) => {
      const h = a[0] ?? 0.25;
      const t = a[1] ?? 0.15;
      return swapAxes([
        { x: 0, y: 0.5 }, { x: h, y: 0 }, { x: h, y: 0.5 - t }, { x: 1 - h, y: 0.5 - t },
        { x: 1 - h, y: 0 }, { x: 1, y: 0.5 }, { x: 1 - h, y: 1 }, { x: 1 - h, y: 0.5 + t },
        { x: h, y: 0.5 + t }, { x: h, y: 1 },
      ]);
    },
  },
  {
    geometry: "quadArrow",
    group: "Block arrows",
    label: "Quad arrow",
    prst: "quadArrow",
    drawType: "quad-arrow",
    outline: () => {
      const t = 0.12;
      const h = 0.28;
      return [
        { x: 0.5, y: 0 }, { x: 0.5 + h, y: h }, { x: 0.5 + t, y: h }, { x: 0.5 + t, y: 0.5 - t },
        { x: 1 - h, y: 0.5 - t }, { x: 1 - h, y: 0.5 - h }, { x: 1, y: 0.5 }, { x: 1 - h, y: 0.5 + h },
        { x: 1 - h, y: 0.5 + t }, { x: 0.5 + t, y: 0.5 + t }, { x: 0.5 + t, y: 1 - h }, { x: 0.5 + h, y: 1 - h },
        { x: 0.5, y: 1 }, { x: 0.5 - h, y: 1 - h }, { x: 0.5 - t, y: 1 - h }, { x: 0.5 - t, y: 0.5 + t },
        { x: h, y: 0.5 + t }, { x: h, y: 0.5 + h }, { x: 0, y: 0.5 }, { x: h, y: 0.5 - h },
        { x: h, y: 0.5 - t }, { x: 0.5 - t, y: 0.5 - t }, { x: 0.5 - t, y: h }, { x: 0.5 - h, y: h },
      ];
    },
  },
  {
    geometry: "homePlate",
    group: "Block arrows",
    label: "Pentagon arrow",
    aliases: ["home plate", "pentagon arrow"],
    prst: "homePlate",
    drawType: "pentagon",
    adjustments: [{ label: "point depth", default: 0.3, min: 0, max: 0.7, axis: "x" }],
    outline: (a) => {
      const d = a[0] ?? 0.3;
      return [{ x: 0, y: 0 }, { x: 1 - d, y: 0 }, { x: 1, y: 0.5 }, { x: 1 - d, y: 1 }, { x: 0, y: 1 }];
    },
  },
  {
    geometry: "chevron",
    group: "Block arrows",
    label: "Chevron",
    prst: "chevron",
    drawType: "chevron",
    adjustments: [{ label: "notch depth", default: 0.3, min: 0, max: 0.7, axis: "x" }],
    outline: (a) => {
      const d = a[0] ?? 0.3;
      return [{ x: 0, y: 0 }, { x: 1 - d, y: 0 }, { x: 1, y: 0.5 }, { x: 1 - d, y: 1 }, { x: 0, y: 1 }, { x: d, y: 0.5 }];
    },
  },
  {
    geometry: "notchedRightArrow",
    group: "Block arrows",
    label: "Notched right arrow",
    prst: "notchedRightArrow",
    drawType: "notched-right-arrow",
    adjustments: ARROW_ADJ,
    outline: (a) => {
      const neck = 1 - (a[0] ?? 0.4);
      const half = a[1] ?? 0.2;
      return [
        { x: 0, y: 0.5 - half }, { x: neck, y: 0.5 - half }, { x: neck, y: 0 }, { x: 1, y: 0.5 },
        { x: neck, y: 1 }, { x: neck, y: 0.5 + half }, { x: 0, y: 0.5 + half }, { x: half, y: 0.5 },
      ];
    },
  },
  {
    geometry: "stripedRightArrow",
    group: "Block arrows",
    label: "Striped right arrow",
    prst: "stripedRightArrow",
    drawType: "striped-right-arrow",
    adjustments: ARROW_ADJ,
    outline: (a) => rightArrowPoly(1 - (a[0] ?? 0.4), a[1] ?? 0.2),
  },
];

const FLOWCHART: ShapeDefinition[] = [
  { geometry: "flowChartProcess", group: "Flowchart", label: "Process", aliases: ["rectangle"], prst: "flowChartProcess", drawType: "flowchart-process", outline: rectPoly },
  { geometry: "flowChartAlternateProcess", group: "Flowchart", label: "Alternate process", prst: "flowChartAlternateProcess", drawType: "flowchart-alternate-process", outline: rectPoly, adjustments: [{ label: "corner radius", default: 0.1667, min: 0, max: 0.5, axis: "x" }] },
  { geometry: "flowChartDecision", group: "Flowchart", label: "Decision", aliases: ["diamond", "if"], prst: "flowChartDecision", drawType: "flowchart-decision", outline: () => [{ x: 0.5, y: 0 }, { x: 1, y: 0.5 }, { x: 0.5, y: 1 }, { x: 0, y: 0.5 }] },
  { geometry: "flowChartInputOutput", group: "Flowchart", label: "Input/Output", aliases: ["data", "parallelogram"], prst: "flowChartInputOutput", drawType: "flowchart-data", outline: () => [{ x: 0.2, y: 0 }, { x: 1, y: 0 }, { x: 0.8, y: 1 }, { x: 0, y: 1 }] },
  { geometry: "flowChartPredefinedProcess", group: "Flowchart", label: "Predefined process", aliases: ["subroutine"], prst: "flowChartPredefinedProcess", drawType: "flowchart-predefined-process", outline: rectPoly, textRect: () => ({ x: 0.12, y: 0, w: 0.76, h: 1 }) },
  { geometry: "flowChartInternalStorage", group: "Flowchart", label: "Internal storage", prst: "flowChartInternalStorage", drawType: "flowchart-internal-storage", outline: rectPoly, textRect: () => ({ x: 0.12, y: 0.12, w: 0.88, h: 0.88 }) },
  { geometry: "flowChartDocument", group: "Flowchart", label: "Document", prst: "flowChartDocument", drawType: "flowchart-document", outline: () => [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 0.82 }, ...waveRun(1, 0, 0.82, 0.98, 12), { x: 0, y: 0.9 }], textRect: () => ({ x: 0, y: 0, w: 1, h: 0.82 }) },
  { geometry: "flowChartMultidocument", group: "Flowchart", label: "Multidocument", prst: "flowChartMultidocument", drawType: "flowchart-multidocument", outline: () => [{ x: 0, y: 0.12 }, { x: 0.12, y: 0.12 }, { x: 0.12, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 0.7 }, { x: 0.88, y: 0.7 }, { x: 0.88, y: 0.82 }, ...waveRun(0.88, 0, 0.82, 0.98, 10), { x: 0, y: 0.9 }] },
  { geometry: "flowChartTerminator", group: "Flowchart", label: "Terminator", aliases: ["start", "end", "stadium"], prst: "flowChartTerminator", drawType: "flowchart-terminator", outline: () => [...arcPoints(0.75, 0.5, 0.25, 0.5, -90, 90, 12), ...arcPoints(0.25, 0.5, 0.25, 0.5, 90, 270, 12)], textRect: () => ({ x: 0.15, y: 0, w: 0.7, h: 1 }) },
  { geometry: "flowChartPreparation", group: "Flowchart", label: "Preparation", prst: "flowChartPreparation", drawType: "flowchart-preparation", outline: () => [{ x: 0.15, y: 0 }, { x: 0.85, y: 0 }, { x: 1, y: 0.5 }, { x: 0.85, y: 1 }, { x: 0.15, y: 1 }, { x: 0, y: 0.5 }], textRect: () => ({ x: 0.15, y: 0, w: 0.7, h: 1 }) },
  { geometry: "flowChartManualInput", group: "Flowchart", label: "Manual input", prst: "flowChartManualInput", drawType: "flowchart-manual-input", outline: () => [{ x: 0, y: 0.25 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }] },
  { geometry: "flowChartManualOperation", group: "Flowchart", label: "Manual operation", prst: "flowChartManualOperation", drawType: "flowchart-manual-operation", outline: () => [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 0.8, y: 1 }, { x: 0.2, y: 1 }] },
  { geometry: "flowChartConnector", group: "Flowchart", label: "Connector", aliases: ["circle"], prst: "flowChartConnector", drawType: "flowchart-connector", outline: () => ringPoints(0.5, 0.5, 0.5, 0.5, 40) },
  { geometry: "flowChartOffpageConnector", group: "Flowchart", label: "Off-page connector", prst: "flowChartOffpageConnector", drawType: "flowchart-off-page-connector", outline: () => [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 0.6 }, { x: 0.5, y: 1 }, { x: 0, y: 0.6 }], textRect: () => ({ x: 0, y: 0, w: 1, h: 0.6 }) },
  { geometry: "flowChartPunchedCard", group: "Flowchart", label: "Card", prst: "flowChartPunchedCard", drawType: "flowchart-card", outline: () => [{ x: 0.15, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { x: 0, y: 0.15 }] },
  { geometry: "flowChartMagneticDisk", group: "Flowchart", label: "Magnetic disk", aliases: ["database", "cylinder"], prst: "flowChartMagneticDisk", drawType: "flowchart-magnetic-disk", outline: () => [...arcPoints(0.5, 0.15, 0.5, 0.15, 180, 360, 16), ...arcPoints(0.5, 0.85, 0.5, 0.15, 0, 180, 16)], textRect: () => ({ x: 0, y: 0.3, w: 1, h: 0.55 }) },
  { geometry: "flowChartDelay", group: "Flowchart", label: "Delay", prst: "flowChartDelay", drawType: "flowchart-delay", outline: () => [{ x: 0, y: 0 }, { x: 0.5, y: 0 }, ...arcPoints(0.5, 0.5, 0.5, 0.5, -90, 90, 16), { x: 0, y: 1 }], textRect: () => ({ x: 0, y: 0, w: 0.7, h: 1 }) },
  { geometry: "flowChartExtract", group: "Flowchart", label: "Extract", aliases: ["triangle"], prst: "flowChartExtract", drawType: "flowchart-extract", outline: () => [{ x: 0.5, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }] },
  { geometry: "flowChartMerge", group: "Flowchart", label: "Merge", prst: "flowChartMerge", drawType: "flowchart-merge", outline: () => [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 0.5, y: 1 }] },
  { geometry: "flowChartSort", group: "Flowchart", label: "Sort", prst: "flowChartSort", drawType: "flowchart-sort", outline: () => [{ x: 0.5, y: 0 }, { x: 1, y: 0.5 }, { x: 0.5, y: 1 }, { x: 0, y: 0.5 }] },
];

/** A row of half-period sine bumps from x=x0 to x=x1 at baseline y, dipping
 * to `low` — the wavy bottom edge of a flowchart document. */
function waveRun(x0: number, x1: number, yHigh: number, yLow: number, steps: number): Point[] {
  const pts: Point[] = [];
  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    const x = x0 + (x1 - x0) * t;
    const y = yHigh + (yLow - yHigh) * (0.5 - 0.5 * Math.cos(TAU * t));
    pts.push(bound({ x, y }));
  }
  return pts;
}

const STARS_BANNERS: ShapeDefinition[] = [
  { geometry: "star4", group: "Stars and banners", label: "Star (4-point)", prst: "star4", drawType: "star4", adjustments: starAdj(), outline: (a) => starPoly(4, a[0] ?? 0.2) },
  { geometry: "star6", group: "Stars and banners", label: "Star (6-point)", prst: "star6", drawType: "star6", adjustments: starAdj(0.5), outline: (a) => starPoly(6, a[0] ?? 0.5) },
  { geometry: "star8", group: "Stars and banners", label: "Star (8-point)", prst: "star8", drawType: "star8", adjustments: starAdj(0.55), outline: (a) => starPoly(8, a[0] ?? 0.55) },
  { geometry: "star10", group: "Stars and banners", label: "Star (10-point)", prst: "star10", drawType: "star10", adjustments: starAdj(0.6), outline: (a) => starPoly(10, a[0] ?? 0.6) },
  { geometry: "star12", group: "Stars and banners", label: "Star (12-point)", prst: "star12", drawType: "star12", adjustments: starAdj(0.65), outline: (a) => starPoly(12, a[0] ?? 0.65) },
  { geometry: "star16", group: "Stars and banners", label: "Star (16-point)", prst: "star16", drawType: "star16", adjustments: starAdj(0.7), outline: (a) => starPoly(16, a[0] ?? 0.7) },
  { geometry: "star24", group: "Stars and banners", label: "Star (24-point)", prst: "star24", drawType: "star24", adjustments: starAdj(0.75), outline: (a) => starPoly(24, a[0] ?? 0.75) },
  { geometry: "star32", group: "Stars and banners", label: "Star (32-point)", prst: "star32", drawType: "star32", adjustments: starAdj(0.8), outline: (a) => starPoly(32, a[0] ?? 0.8) },
  { geometry: "irregularSeal1", group: "Stars and banners", label: "Explosion 1", aliases: ["seal", "burst", "bang"], prst: "irregularSeal1", drawType: "bang", outline: () => sealPoly(12, 7) },
  { geometry: "irregularSeal2", group: "Stars and banners", label: "Explosion 2", aliases: ["seal", "burst", "bang"], prst: "irregularSeal2", drawType: "bang2", approximateIn: ["odt"], outline: () => sealPoly(16, 11) },
  { geometry: "ribbon", group: "Stars and banners", label: "Down ribbon", aliases: ["banner"], prst: "ribbon", drawType: "down-ribbon", outline: () => ribbonPoly(false) },
  { geometry: "ribbon2", group: "Stars and banners", label: "Up ribbon", aliases: ["banner"], prst: "ribbon2", drawType: "up-ribbon", outline: () => ribbonPoly(true) },
  { geometry: "verticalScroll", group: "Stars and banners", label: "Vertical scroll", prst: "verticalScroll", drawType: "vertical-scroll", outline: () => scrollPoly(true) },
  { geometry: "horizontalScroll", group: "Stars and banners", label: "Horizontal scroll", prst: "horizontalScroll", drawType: "horizontal-scroll", outline: () => scrollPoly(false) },
  { geometry: "wave", group: "Stars and banners", label: "Wave", prst: "wave", drawType: "wave", adjustments: [{ label: "amplitude", default: 0.1, min: 0.02, max: 0.2, axis: "y" }], outline: (a) => wavePoly(a[0] ?? 0.1, 1) },
  { geometry: "doubleWave", group: "Stars and banners", label: "Double wave", prst: "doubleWave", drawType: "double-wave", adjustments: [{ label: "amplitude", default: 0.08, min: 0.02, max: 0.15, axis: "y" }], outline: (a) => wavePoly(a[0] ?? 0.08, 2) },
];

function starAdj(def = 0.38): ShapeAdjustmentSpec[] {
  return [{ label: "point depth", default: def, min: 0.05, max: 0.95, axis: "radial" }];
}

/** A jagged explosion outline: outer/inner radii alternating with a fixed
 * pseudo-random jitter so the two seals look different but stay stable. */
function sealPoly(points: number, seed: number): Point[] {
  const pts: Point[] = [];
  for (let i = 0; i < 2 * points; i++) {
    const jitter = 0.06 * Math.sin(seed * (i + 1) * 1.7);
    const r = (i % 2 === 0 ? 0.5 : 0.28) + jitter;
    const a = -Math.PI / 2 + (i * Math.PI) / points;
    pts.push(bound({ x: 0.5 + r * Math.cos(a), y: 0.5 + r * Math.sin(a) }));
  }
  return pts;
}

/** A ribbon banner: a central plate with two folded tails, `up` flips the
 * fold above the plate rather than below. */
function ribbonPoly(up: boolean): Point[] {
  const base: Point[] = [
    { x: 0, y: 0.15 }, { x: 0.15, y: 0.15 }, { x: 0.15, y: 0 }, { x: 0.85, y: 0 },
    { x: 0.85, y: 0.15 }, { x: 1, y: 0.15 }, { x: 0.85, y: 0.35 }, { x: 1, y: 0.55 },
    { x: 0.85, y: 0.55 }, { x: 0.85, y: 0.75 }, { x: 0.15, y: 0.75 }, { x: 0.15, y: 0.55 },
    { x: 0, y: 0.55 }, { x: 0.15, y: 0.35 },
  ];
  return up ? base.map((p) => ({ x: p.x, y: 1 - p.y })) : base;
}

/** A scroll (paper roll) silhouette: a plate with one curled corner. */
function scrollPoly(vertical: boolean): Point[] {
  const base: Point[] = [
    { x: 0.12, y: 0.08 }, { x: 0.92, y: 0.08 },
    ...arcPoints(0.92, 0.2, 0.08, 0.12, -90, 90, 6),
    { x: 0.88, y: 0.92 }, { x: 0.08, y: 0.92 },
    ...arcPoints(0.08, 0.8, 0.08, 0.12, 90, 270, 6),
  ];
  return vertical ? base : swapAxes(base);
}

/** A wave/banner: top and bottom sine edges over `periods` periods, sheared
 * so the leading edge tilts (as ECMA's wave does). */
function wavePoly(amp: number, periods: number): Point[] {
  const steps = periods * 12;
  const top: Point[] = [];
  const bottom: Point[] = [];
  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    const x = t;
    const s = amp * Math.sin(TAU * periods * t);
    top.push(bound({ x, y: amp + s }));
    bottom.push(bound({ x, y: 1 - amp + s }));
  }
  return [...top, ...bottom.reverse()];
}

const CALLOUTS: ShapeDefinition[] = [
  {
    geometry: "wedgeRectCallout",
    group: "Callouts",
    label: "Rectangular callout",
    aliases: ["speech", "bubble"],
    prst: "wedgeRectCallout",
    drawType: "rectangular-callout",
    adjustments: [
      { label: "tail x", default: WEDGE_TAIL.tip, min: 0, max: 1, axis: "x" },
      { label: "tail y", default: 1, min: CALLOUT_BODY, max: 1, axis: "y" },
    ],
    outline: (a) => calloutPoly(WEDGE_TAIL.left, a[0] ?? WEDGE_TAIL.tip, WEDGE_TAIL.right, a[1] ?? 1),
    textRect: calloutTextRect,
  },
  {
    geometry: "wedgeRoundRectCallout",
    group: "Callouts",
    label: "Rounded callout",
    aliases: ["speech", "bubble"],
    prst: "wedgeRoundRectCallout",
    drawType: "round-rectangular-callout",
    adjustments: [
      { label: "tail x", default: ROUND_TAIL.tip, min: 0, max: 1, axis: "x" },
      { label: "tail y", default: 1, min: CALLOUT_BODY, max: 1, axis: "y" },
    ],
    outline: (a) => calloutPoly(ROUND_TAIL.left, a[0] ?? ROUND_TAIL.tip, ROUND_TAIL.right, a[1] ?? 1),
    textRect: calloutTextRect,
  },
  {
    geometry: "wedgeEllipseCallout",
    group: "Callouts",
    label: "Oval callout",
    aliases: ["speech", "bubble", "round"],
    prst: "wedgeEllipseCallout",
    drawType: "round-callout",
    adjustments: CALLOUT_TAIL_ADJ,
    // A full ellipse body (broken at the bottom to admit the tail) plus the
    // tail tip — so the box centre lands inside the painted region.
    outline: (a) => [...arcPoints(0.5, 0.4, 0.5, 0.4, 110, 430, 30), { x: a[0] ?? 0.2, y: a[1] ?? 1 }],
    textRect: () => ({ x: 0.1, y: 0.05, w: 0.8, h: 0.6 }),
  },
  {
    geometry: "cloudCallout",
    group: "Callouts",
    label: "Cloud callout",
    aliases: ["speech", "thought"],
    prst: "cloudCallout",
    drawType: "cloud-callout",
    outline: () => cloudRing(0.5, 0.38, 0.42, 0.32, 8),
    textRect: () => ({ x: 0.15, y: 0.1, w: 0.7, h: 0.5 }),
  },
  {
    geometry: "borderCallout1",
    group: "Callouts",
    label: "Line callout 1",
    aliases: ["leader"],
    prst: "borderCallout1",
    drawType: "line-callout-1",
    outline: rectPoly,
    approximateIn: ["odt"],
  },
  {
    geometry: "borderCallout2",
    group: "Callouts",
    label: "Line callout 2",
    aliases: ["leader"],
    prst: "borderCallout2",
    drawType: "line-callout-2",
    outline: rectPoly,
    approximateIn: ["odt"],
  },
  {
    geometry: "borderCallout3",
    group: "Callouts",
    label: "Line callout 3",
    aliases: ["leader"],
    prst: "borderCallout3",
    drawType: "line-callout-3",
    outline: rectPoly,
    approximateIn: ["odt"],
  },
  {
    geometry: "accentCallout1",
    group: "Callouts",
    label: "Accent callout 1",
    aliases: ["leader"],
    prst: "accentCallout1",
    drawType: "accent-callout-1",
    outline: rectPoly,
    approximateIn: ["odt"],
  },
];

/** The one table SPEC-069 AC-15a makes the sole source of the catalogue. */
export const SHAPE_DEFINITIONS: ShapeDefinition[] = [
  ...LINES,
  ...BASIC,
  ...BLOCK_ARROWS,
  ...FLOWCHART,
  ...STARS_BANNERS,
  ...CALLOUTS,
];

const BY_GEOMETRY = new Map<ShapeGeometry, ShapeDefinition>(SHAPE_DEFINITIONS.map((d) => [d.geometry, d]));

/** The definition for `geometry`, or undefined for a non-catalogue one. */
export function shapeDefinition(geometry: ShapeGeometry): ShapeDefinition | undefined {
  return BY_GEOMETRY.get(geometry);
}

/** The default adjustment vector for a definition (every handle at its
 * declared default), the value `shapePolygon` and a fresh shape use. */
export function shapeAdjustmentDefaults(def: ShapeDefinition): number[] {
  return (def.adjustments ?? []).map((spec) => spec.default);
}

// Adjustment values are normalised to fractions at the reader boundary
// (SPEC-069 "Adjustment values"), the same choice the model makes for
// lengths (cm) and rotation (degrees). Both formats express an adjustment as
// an integer in a shape-relative scale — OOXML guide formulas and ODF
// draw:modifiers — so a single hundred-thousandths scale round-trips a
// fraction through our own reader/writer for every geometry. (Word's true
// per-geometry semantics vary; matching them exactly is beyond this spec,
// which only promises a faithful round-trip and clamped rendering.)
const ADJ_SCALE = 100000;

/** A 0..1-style fraction as the integer guide value the formats store. */
export function adjustmentToGuideValue(fraction: number): number {
  return Math.round(fraction * ADJ_SCALE);
}

/** A stored integer guide value back to its 0..1-style fraction. */
export function guideValueToAdjustment(value: number): number {
  return value / ADJ_SCALE;
}

/** The `a:gd@name` / positional name a geometry's i-th adjustment writes,
 * matching Word's own convention (`adj` for a lone one, else `adj1`..). */
export function adjustmentGuideName(index: number, count: number): string {
  return count <= 1 ? "adj" : `adj${index + 1}`;
}

/** A shape's adjustments resolved for rendering: stored values where present,
 * defaults for any missing entry, each clamped to its declared range
 * (AC-7a). Never mutates the stored array. */
export function resolvedShapeAdjustments(geometry: ShapeGeometry, stored?: readonly number[]): number[] {
  const def = shapeDefinition(geometry);
  if (!def) {
    return stored ? [...stored] : [];
  }
  return (def.adjustments ?? []).map((spec, i) => {
    const value = stored?.[i];
    if (value === undefined || Number.isNaN(value)) {
      return spec.default;
    }
    return Math.min(spec.max, Math.max(spec.min, value));
  });
}

// --- SVG path geometry (SPEC-088) -------------------------------------------
// A DOM-free `d` string for a shape's outline, so the HTML exporter can draw a
// shape as an inline <svg><path> from the SAME ShapeDefinition outline the
// canvas painter (core/shapeCanvas) consumes — polygon geometries become M/L/Z
// paths, ellipse and the rounded-rectangle cases their arc equivalents, holes
// (donut, frame) additional even-odd subpaths, lines a corner-to-corner
// segment. Kept in core beside the catalogue (not shapeGeometry.ts, which
// imports this module) so it has no import cycle; the maths it needs
// (roundRect radius, rounded-callout outline) is small and reproduced here.

/** Rounds a coordinate to at most 3 decimals for a compact, deterministic path
 * string (float dust like -1e-15 collapses to 0). */
function svgNum(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

/** A closed polygon (already in px) as an `M/L…Z` path. */
function polygonPathD(points: ShapeOutlinePath): string {
  if (points.length === 0) {
    return "";
  }
  const [first, ...rest] = points;
  const parts = [`M ${svgNum(first.x)} ${svgNum(first.y)}`];
  for (const p of rest) {
    parts.push(`L ${svgNum(p.x)} ${svgNum(p.y)}`);
  }
  parts.push("Z");
  return parts.join(" ");
}

function scaleOutline(outline: ShapeOutlinePath, widthPx: number, heightPx: number): ShapeOutlinePath {
  return outline.map((p) => ({ x: p.x * widthPx, y: p.y * heightPx }));
}

/** A full ellipse inscribed in the box, as two SVG arcs (matching the canvas's
 * `ctx.ellipse`). */
function ellipsePathD(widthPx: number, heightPx: number): string {
  const rx = Math.max(0, widthPx / 2);
  const ry = Math.max(0, heightPx / 2);
  const cy = heightPx / 2;
  return (
    `M 0 ${svgNum(cy)} ` +
    `A ${svgNum(rx)} ${svgNum(ry)} 0 1 0 ${svgNum(widthPx)} ${svgNum(cy)} ` +
    `A ${svgNum(rx)} ${svgNum(ry)} 0 1 0 0 ${svgNum(cy)} Z`
  );
}

/** A rounded rectangle with quarter-circle corners (matching the canvas's
 * `roundRectPath`): `radiusFraction` of the shorter side, clamped to half. */
function roundRectPathD(widthPx: number, heightPx: number, radiusFraction: number): string {
  const base = Math.min(widthPx, heightPx) * radiusFraction;
  const r = Math.max(0, Math.min(base, widthPx / 2, heightPx / 2));
  const a = (x: number, y: number) => `A ${svgNum(r)} ${svgNum(r)} 0 0 1 ${svgNum(x)} ${svgNum(y)}`;
  return [
    `M ${svgNum(r)} 0`,
    `L ${svgNum(widthPx - r)} 0`,
    a(widthPx, r),
    `L ${svgNum(widthPx)} ${svgNum(heightPx - r)}`,
    a(widthPx - r, heightPx),
    `L ${svgNum(r)} ${svgNum(heightPx)}`,
    a(0, heightPx - r),
    `L 0 ${svgNum(r)}`,
    a(r, 0),
    "Z",
  ].join(" ");
}

// The rounded-callout body/tail fractions — the same constants
// core/shapeGeometry.ts walks its outline with (kept in sync; the wedge tail's
// straight edges come from the catalogue's own CALLOUT_BODY/ROUND_TAIL above).
const ROUND_CALLOUT_BODY = CALLOUT_BODY;
const ROUND_CALLOUT_TAIL = ROUND_TAIL;

/** `wedgeRoundRectCallout` as arcs (matching the canvas's `roundedCalloutPath`):
 * a rounded-rectangle body whose bottom edge detours into the sharp tail. */
function roundedCalloutPathD(widthPx: number, heightPx: number): string {
  const body = heightPx * ROUND_CALLOUT_BODY;
  const radius = Math.max(
    0,
    Math.min(
      Math.min(widthPx, body) * 0.1667,
      ROUND_CALLOUT_TAIL.left * widthPx,
      (1 - ROUND_CALLOUT_TAIL.right) * widthPx,
      widthPx / 2,
      body / 2,
    ),
  );
  const arc = (x: number, y: number) => `A ${svgNum(radius)} ${svgNum(radius)} 0 0 1 ${svgNum(x)} ${svgNum(y)}`;
  return [
    `M ${svgNum(radius)} 0`,
    `L ${svgNum(widthPx - radius)} 0`,
    arc(widthPx, radius),
    `L ${svgNum(widthPx)} ${svgNum(body - radius)}`,
    arc(widthPx - radius, body),
    `L ${svgNum(ROUND_CALLOUT_TAIL.right * widthPx)} ${svgNum(body)}`,
    `L ${svgNum(ROUND_CALLOUT_TAIL.tip * widthPx)} ${svgNum(heightPx)}`,
    `L ${svgNum(ROUND_CALLOUT_TAIL.left * widthPx)} ${svgNum(body)}`,
    `L ${svgNum(radius)} ${svgNum(body)}`,
    arc(0, body - radius),
    `L 0 ${svgNum(radius)}`,
    arc(radius, 0),
    "Z",
  ].join(" ");
}

/** The SVG path `d` for a shape's outline, in pixels, at the given resolved-or-
 * raw adjustments — the exporter's equivalent of `paintShape`'s canvas path.
 * Polygon geometries return `M/L…Z`; `ellipse` and the rounded-rectangle
 * cases return arc equivalents; holes (donut/frame) are appended as extra
 * subpaths (fill with `fill-rule="evenodd"`); lines return a corner-to-corner
 * segment. Returns "" for a geometry with no catalogue definition. */
export function shapeSvgPath(
  geometry: ShapeGeometry,
  widthPx: number,
  heightPx: number,
  adjustments?: readonly number[],
): string {
  const def = shapeDefinition(geometry);
  if (def?.isLine) {
    return `M 0 0 L ${svgNum(widthPx)} ${svgNum(heightPx)}`;
  }
  if (geometry === "ellipse") {
    return ellipsePathD(widthPx, heightPx);
  }
  const resolved = resolvedShapeAdjustments(geometry, adjustments);
  if (geometry === "roundRect" || geometry === "flowChartAlternateProcess") {
    return roundRectPathD(widthPx, heightPx, resolved[0] ?? 0.1667);
  }
  if (geometry === "wedgeRoundRectCallout") {
    return roundedCalloutPathD(widthPx, heightPx);
  }
  if (!def) {
    return "";
  }
  const main = polygonPathD(scaleOutline(def.outline(resolved), widthPx, heightPx));
  const holes = (def.holes?.(resolved) ?? []).map((hole) => polygonPathD(scaleOutline(hole, widthPx, heightPx)));
  return [main, ...holes].filter((d) => d.length > 0).join(" ");
}
