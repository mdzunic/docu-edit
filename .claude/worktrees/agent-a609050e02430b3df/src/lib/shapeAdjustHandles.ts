// SPEC-069 adjustment handles: the pure geometry mapping a shape's declared
// adjustments to draggable handle positions and back, plus the keyboard
// nudge. Framework-free (no Vue, no DOM) so it unit-tests directly; the DOM
// wiring lives in useShapes.ts and the handle elements in DocumentView.
import { resolvedShapeAdjustments, shapeDefinition, type ShapeAdjustmentSpec } from "./core/shapeCatalogue";
import type { ShapeGeometry } from "./core/shapes";

/** A handle as the overlay renders it: its identity/label/value plus the
 * absolute pixel style useShapes computed. */
export interface AdjustmentHandleView {
  index: number;
  label: string;
  value: number;
  style: Record<string, string>;
}

/** Typed pass-through for the runtime Array prop ShapeAdjustHandles.vue
 * receives (the .vue espree parser bars a PropType annotation there). */
export function asAdjustmentHandleViews(value: unknown): AdjustmentHandleView[] {
  return value as AdjustmentHandleView[];
}

export interface AdjustmentHandle {
  /** Index into Shape.adjustments / the definition's adjustment list. */
  index: number;
  /** Accessible label, e.g. "corner radius". */
  label: string;
  /** Current (clamped) value. */
  value: number;
  /** Handle position in unit-square fractions (0..1 of width/height). */
  xFrac: number;
  yFrac: number;
  axis: ShapeAdjustmentSpec["axis"];
}

/** Where a handle sits, in unit-square fractions, for an adjustment's axis
 * and current value: an x-axis handle rides the top edge, a y-axis handle
 * the left edge, and a radial handle sits straight up from the centre at a
 * distance proportional to its value. */
function handlePoint(spec: ShapeAdjustmentSpec, value: number): { xFrac: number; yFrac: number } {
  if (spec.axis === "x") {
    return { xFrac: value, yFrac: 0 };
  }
  if (spec.axis === "y") {
    return { xFrac: 0, yFrac: value };
  }
  return { xFrac: 0.5, yFrac: 0.5 - value * 0.5 };
}

/** The handles a shape shows (AC-5a): one per declared adjustment, at its
 * current (clamped) value. Empty for a geometry with no adjustments. */
export function adjustmentHandles(geometry: ShapeGeometry, stored?: readonly number[]): AdjustmentHandle[] {
  const def = shapeDefinition(geometry);
  if (!def?.adjustments || def.adjustments.length === 0) {
    return [];
  }
  const values = resolvedShapeAdjustments(geometry, stored);
  return def.adjustments.map((spec, index) => {
    const value = values[index];
    const { xFrac, yFrac } = handlePoint(spec, value);
    return { index, label: spec.label, value, xFrac, yFrac, axis: spec.axis };
  });
}

function clampToSpec(spec: ShapeAdjustmentSpec, value: number): number {
  return Math.min(spec.max, Math.max(spec.min, value));
}

/** The adjustment value a drag to a unit-square point maps to, clamped to
 * the declared range (AC-5c). Inverse of `handlePoint` for the axis. */
export function adjustmentValueFromPoint(
  geometry: ShapeGeometry,
  index: number,
  xFrac: number,
  yFrac: number,
): number | null {
  const spec = shapeDefinition(geometry)?.adjustments?.[index];
  if (!spec) {
    return null;
  }
  if (spec.axis === "x") {
    return clampToSpec(spec, xFrac);
  }
  if (spec.axis === "y") {
    return clampToSpec(spec, yFrac);
  }
  const dist = Math.hypot(xFrac - 0.5, yFrac - 0.5);
  return clampToSpec(spec, dist * 2);
}

/** The full resolved adjustment vector with entry `index` nudged by
 * `steps` * the keyboard step (1% of the range, or 5% with Shift), clamped
 * to range (AC-5c, "Adjustment handles" keyboard). */
export function nudgedAdjustments(
  geometry: ShapeGeometry,
  stored: readonly number[] | undefined,
  index: number,
  steps: number,
  coarse: boolean,
): number[] {
  const def = shapeDefinition(geometry);
  const values = resolvedShapeAdjustments(geometry, stored);
  const spec = def?.adjustments?.[index];
  if (!spec) {
    return values;
  }
  const step = (coarse ? 0.05 : 0.01) * (spec.max - spec.min);
  values[index] = clampToSpec(spec, values[index] + steps * step);
  return values;
}
