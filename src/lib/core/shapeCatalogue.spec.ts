// SPEC-069: the single-source shape table. These pin the AC-12 geometry
// invariants over every catalogue geometry, the AC-15a derivations, and the
// adjustment codec, all straight off SHAPE_DEFINITIONS so a new entry is
// covered automatically.
import { describe, expect, it } from "vitest";
import {
  SHAPE_DEFINITIONS,
  adjustmentGuideName,
  adjustmentToGuideValue,
  guideValueToAdjustment,
  resolvedShapeAdjustments,
  shapeAdjustmentDefaults,
  shapeDefinition,
} from "./shapeCatalogue";
import { LINE_GEOMETRIES, MIN_SHAPE_SIZE_CM, SHAPE_CATALOGUE } from "./shapes";
import { scalePolygon, shapePolygon, type Point } from "./shapeGeometry";

/** Signed polygon area (shoelace), in the units the points are given in. */
function polygonArea(points: Point[]): number {
  let area = 0;
  for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
    area += points[j].x * points[i].y - points[i].x * points[j].y;
  }
  return Math.abs(area) / 2;
}

const AREA_DEFINITIONS = SHAPE_DEFINITIONS.filter((def) => !def.isLine);

describe("SHAPE_DEFINITIONS — geometry invariants (AC-12)", () => {
  for (const def of AREA_DEFINITIONS) {
    describe(def.geometry, () => {
      const outline = def.outline(shapeAdjustmentDefaults(def));

      it("is a closed polygon of at least three vertices (AC-12a)", () => {
        expect(outline.length).toBeGreaterThanOrEqual(3);
      });

      it("stays within the unit square (AC-12b)", () => {
        for (const p of outline) {
          expect(p.x, `${def.geometry} x`).toBeGreaterThanOrEqual(0);
          expect(p.x, `${def.geometry} x`).toBeLessThanOrEqual(1);
          expect(p.y, `${def.geometry} y`).toBeGreaterThanOrEqual(0);
          expect(p.y, `${def.geometry} y`).toBeLessThanOrEqual(1);
        }
        for (const hole of def.holes?.(shapeAdjustmentDefaults(def)) ?? []) {
          for (const p of hole) {
            expect(p.x).toBeGreaterThanOrEqual(0);
            expect(p.x).toBeLessThanOrEqual(1);
            expect(p.y).toBeGreaterThanOrEqual(0);
            expect(p.y).toBeLessThanOrEqual(1);
          }
        }
      });

      it("is non-degenerate at the minimum shape size (AC-12c)", () => {
        const scaled = scalePolygon(outline, MIN_SHAPE_SIZE_CM, MIN_SHAPE_SIZE_CM);
        expect(polygonArea(scaled)).toBeGreaterThan(0);
      });
    });
  }
});

describe("SHAPE_DEFINITIONS — derivations (AC-15a)", () => {
  it("is the sole source of SHAPE_CATALOGUE (same geometries, groups, labels, in order)", () => {
    expect(SHAPE_CATALOGUE).toEqual(
      SHAPE_DEFINITIONS.map((def) => ({ geometry: def.geometry, group: def.group, label: def.label })),
    );
  });

  it("is the sole source of LINE_GEOMETRIES", () => {
    expect([...LINE_GEOMETRIES].sort()).toEqual(
      SHAPE_DEFINITIONS.filter((def) => def.isLine).map((def) => def.geometry).sort(),
    );
  });

  it("gives every geometry a prstGeom name, a draw:type name and a label (AC-1b, AC-1c)", () => {
    for (const def of SHAPE_DEFINITIONS) {
      expect(def.prst, def.geometry).toBeTruthy();
      expect(def.drawType, def.geometry).toBeTruthy();
      expect(def.label, def.geometry).toBeTruthy();
    }
  });

  it("gives every prst a unique geometry so the OOXML map round-trips", () => {
    const prsts = SHAPE_DEFINITIONS.map((def) => def.prst);
    expect(new Set(prsts).size).toBe(prsts.length);
  });

  it("gives every area geometry a unique draw:type so the ODF reverse map round-trips", () => {
    const drawTypes = AREA_DEFINITIONS.map((def) => def.drawType);
    expect(new Set(drawTypes).size).toBe(drawTypes.length);
  });

  it("shapePolygon returns the definition outline at defaults for every area geometry", () => {
    for (const def of AREA_DEFINITIONS) {
      if (def.geometry === "ellipse") {
        expect(shapePolygon("ellipse")).toBeNull();
        continue;
      }
      expect(shapePolygon(def.geometry)).toEqual(def.outline(shapeAdjustmentDefaults(def)));
    }
  });
});

describe("adjustment codec", () => {
  it("round-trips a fraction through the guide value scale", () => {
    for (const fraction of [0, 0.1667, 0.25, 0.5, 0.8]) {
      expect(guideValueToAdjustment(adjustmentToGuideValue(fraction))).toBeCloseTo(fraction, 5);
    }
  });

  it("names a lone adjustment 'adj' and multiple 'adj1'..'adjN'", () => {
    expect(adjustmentGuideName(0, 1)).toBe("adj");
    expect(adjustmentGuideName(0, 2)).toBe("adj1");
    expect(adjustmentGuideName(1, 2)).toBe("adj2");
  });

  it("fills missing entries with defaults and clamps out-of-range for rendering (AC-7a)", () => {
    const roundRect = shapeDefinition("roundRect")!;
    // No stored values -> all defaults.
    expect(resolvedShapeAdjustments("roundRect")).toEqual([roundRect.adjustments![0].default]);
    // Out of range -> clamped to the declared max.
    expect(resolvedShapeAdjustments("roundRect", [5])).toEqual([roundRect.adjustments![0].max]);
    expect(resolvedShapeAdjustments("roundRect", [-1])).toEqual([roundRect.adjustments![0].min]);
  });

  it("leaves a geometry with no declared adjustments empty", () => {
    expect(resolvedShapeAdjustments("rect")).toEqual([]);
  });
});
