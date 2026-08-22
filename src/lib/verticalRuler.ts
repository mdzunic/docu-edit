// Geometry for the vertical ruler down the left edge of the page (SPEC-064
// AC-1..6): the unit-labelled tick scheme, and the mapping between a band
// boundary's physical position and the model value a drag on it sets.
// Framework-free so the tick maths and the clamps are testable without
// mounting the component, mirroring lib/ruler.ts for the horizontal ruler.
//
// All positions are centimetres measured from the page's PHYSICAL TOP edge,
// matching CSS `top` on the ruler's markers. Tick LABELS, by contrast, count
// outward from the top of the text area (the top margin), so the number the
// ruler shows is the distance into — or out of — the printable page, exactly
// as a word processor's vertical ruler reads (AC-4a).
import { formatLength, toCm, type MeasurementUnit } from "./core/units";
import { RULER_DRAG_STEP_CM } from "./ruler";

/** The page geometry the vertical ruler positions its markers against. */
export interface VerticalRulerContext {
  pageHeightCm: number;
  marginTopCm: number;
  marginBottomCm: number;
  /** Header band distance from the page's top edge (SPEC-029). */
  headerDistanceCm: number;
  /** Footer band distance from the page's bottom edge (SPEC-029). */
  footerDistanceCm: number;
}

/** The draggable band boundaries on the vertical ruler. */
export type VerticalMarkerKind = "margin-top" | "margin-bottom" | "header" | "footer";

/** The model field a band-boundary drag updates, with its new value in cm. */
export interface VerticalRulerUpdate {
  kind: "marginTopCm" | "marginBottomCm" | "headerDistanceCm" | "footerDistanceCm";
  valueCm: number;
}

/** One tick on the ruler: its position from the page top (cm), whether it is a
 * labelled major tick, and its label (the empty string on minor ticks). */
export interface VerticalTick {
  cm: number;
  major: boolean;
  label: string;
}

/** A caret-table row boundary the ruler can drag, measured from the page's top
 * edge: the row's top and bottom, plus its index (AC-5). */
export interface VerticalRowBoundary {
  row: number;
  topCm: number;
  bottomCm: number;
}

/** Typed pass-through for a VerticalRowBoundary[] coming from the ruler's
 * `unknown[]` prop (espree bars the annotation in `.vue` scripts — same idiom
 * as documentModel.asTabStops). */
export function asRowBoundaries(value: unknown): VerticalRowBoundary[] {
  return value as VerticalRowBoundary[];
}

function snapStep(cm: number): number {
  return Math.round(cm / RULER_DRAG_STEP_CM) * RULER_DRAG_STEP_CM;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/** The tick scheme per unit (SPEC-064 "Tick scheme"), resolved for a zoom.
 * `majorCm` labels; `minorCm` is the finest step drawn. Ticks are geometry,
 * derived from the fixed unit ratios and the zoom, never from a design token
 * (SPEC-058's third exemption). Below 50 % only majors draw; between 50 % and
 * 150 % the minor step draws; above 150 % the next (finer) subdivision draws.
 * Inches additionally coarsen to quarter-inch minors below 75 % zoom. */
export function verticalTickScheme(unit: MeasurementUnit, zoom = 1): { majorCm: number; minorCm: number } {
  // Major step and the two finer steps per unit, all in centimetres.
  // Major / minor / next-subdivision / coarse-minor steps per unit, in cm.
  const scheme =
    unit === "cm"
      ? { majorCm: 1, minorCm: 0.5, subCm: 0.25, coarseMinorCm: 0.5 }
      : unit === "mm"
        ? { majorCm: 1, minorCm: 0.5, subCm: 0.25, coarseMinorCm: 0.5 } // 10mm / 5mm / 2.5mm
        : unit === "in"
          ? {
              majorCm: toCm(1, "in"),
              minorCm: toCm(1 / 8, "in"),
              subCm: toCm(1 / 16, "in"),
              coarseMinorCm: toCm(1 / 4, "in"),
            }
          : {
              majorCm: toCm(72, "pt"),
              minorCm: toCm(36, "pt"),
              subCm: toCm(18, "pt"),
              coarseMinorCm: toCm(36, "pt"),
            };
  if (zoom < 0.5) {
    // Only major ticks draw — the minor step collapses onto the major.
    return { majorCm: scheme.majorCm, minorCm: scheme.majorCm };
  }
  if (zoom > 1.5) {
    return { majorCm: scheme.majorCm, minorCm: scheme.subCm };
  }
  if (zoom < 0.75) {
    return { majorCm: scheme.majorCm, minorCm: scheme.coarseMinorCm };
  }
  return { majorCm: scheme.majorCm, minorCm: scheme.minorCm };
}

/** The ruler's ticks, positioned from the page top and labelled outward from
 * the text-area top (the numbering origin, AC-4a). The origin tick is labelled
 * `"0"`; every major tick above or below it counts the absolute distance into
 * the page, so a reader sees how far a boundary sits from the printable top. */
export function verticalRulerTicks(
  pageHeightCm: number,
  marginTopCm: number,
  unit: MeasurementUnit,
  zoom = 1,
): VerticalTick[] {
  const { majorCm, minorCm } = verticalTickScheme(unit, zoom);
  if (!(pageHeightCm > 0) || !(minorCm > 0)) {
    return [];
  }
  const ticks: VerticalTick[] = [];
  // Ticks are laid on a grid anchored at the origin (marginTopCm), so a
  // labelled major lands an exact unit from the text-area top.
  const firstN = Math.ceil((0 - marginTopCm) / minorCm - 1e-6);
  const lastN = Math.floor((pageHeightCm - marginTopCm) / minorCm + 1e-6);
  for (let n = firstN; n <= lastN; n += 1) {
    const distanceFromOrigin = n * minorCm;
    const cm = marginTopCm + distanceFromOrigin;
    const ratio = distanceFromOrigin / majorCm;
    const major = Math.abs(ratio - Math.round(ratio)) < 1e-6;
    ticks.push({
      cm,
      major,
      label: major ? formatLength(Math.abs(distanceFromOrigin), unit) : "",
    });
  }
  return ticks;
}

/** A band boundary's physical position on the ruler, in cm from the page top. */
export function verticalMarkerPositionCm(kind: VerticalMarkerKind, ctx: VerticalRulerContext): number {
  switch (kind) {
    case "margin-top":
      return ctx.marginTopCm;
    case "margin-bottom":
      return ctx.pageHeightCm - ctx.marginBottomCm;
    case "header":
      return ctx.headerDistanceCm;
    case "footer":
      return ctx.pageHeightCm - ctx.footerDistanceCm;
  }
}

/** The model value a band boundary dragged to `positionCm` (cm from the page
 * top) implies, snapped to the drag grid and clamped inside the page. Top and
 * bottom margins may not cross each other; header/footer distances stay within
 * the page. The command layer re-snaps onto the exact twip grid before storing
 * (AC-2a). */
export function valueFromVerticalPosition(
  kind: VerticalMarkerKind,
  positionCm: number,
  ctx: VerticalRulerContext,
): VerticalRulerUpdate {
  const position = clamp(positionCm, 0, ctx.pageHeightCm);
  switch (kind) {
    case "margin-top":
      return {
        kind: "marginTopCm",
        valueCm: clamp(snapStep(position), 0, Math.max(0, ctx.pageHeightCm - ctx.marginBottomCm)),
      };
    case "margin-bottom":
      return {
        kind: "marginBottomCm",
        valueCm: clamp(snapStep(ctx.pageHeightCm - position), 0, Math.max(0, ctx.pageHeightCm - ctx.marginTopCm)),
      };
    case "header":
      return { kind: "headerDistanceCm", valueCm: clamp(snapStep(position), 0, ctx.pageHeightCm) };
    case "footer":
      return { kind: "footerDistanceCm", valueCm: clamp(snapStep(ctx.pageHeightCm - position), 0, ctx.pageHeightCm) };
  }
}
