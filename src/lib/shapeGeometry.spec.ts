// SPEC-051: pure geometry for the Shape catalogue — see shapeGeometry.ts's
// own header comment. Framework-free (no canvas), so every function here is
// exercised directly against known coordinates.
import { describe, expect, it } from "vitest";
import {
  distanceToPolygonOutline,
  distanceToSegment,
  HIT_TOLERANCE_CM,
  lineSegment,
  pointHitsShape,
  pointInEllipse,
  pointInPolygon,
  roundedCalloutOutline,
  roundedCalloutRadius,
  roundRectRadius,
  scalePolygon,
  shapePolygon,
  unrotatePoint,
  type Point,
} from "./shapeGeometry";
import { LINE_GEOMETRIES, SHAPE_CATALOGUE, type ShapeGeometry } from "./documentModel";

// Geometries with a real polygon representation: every catalogue geometry
// except the two line geometries (a segment, not a region — see
// lineSegment) and ellipse (no polygon at all — see pointInEllipse), both
// of which get their own dedicated describe blocks below.
const POLYGON_GEOMETRIES = SHAPE_CATALOGUE.map((entry) => entry.geometry).filter(
  (geometry) => geometry !== "line" && geometry !== "arrow" && geometry !== "ellipse",
);

// --- shapePolygon ------------------------------------------------------------

describe("shapePolygon", () => {
  it("returns null for ellipse, line and arrow (no polygon representation)", () => {
    expect(shapePolygon("ellipse")).toBeNull();
    expect(shapePolygon("line")).toBeNull();
    expect(shapePolygon("arrow")).toBeNull();
  });

  it("returns a closed, non-empty polygon in unit-square fractions for every other catalogue geometry", () => {
    for (const geometry of POLYGON_GEOMETRIES) {
      const polygon = shapePolygon(geometry);
      expect(polygon, geometry).not.toBeNull();
      expect(polygon!.length).toBeGreaterThanOrEqual(3);
      for (const point of polygon!) {
        expect(point.x).toBeGreaterThanOrEqual(0);
        expect(point.x).toBeLessThanOrEqual(1);
        expect(point.y).toBeGreaterThanOrEqual(0);
        expect(point.y).toBeLessThanOrEqual(1);
      }
    }
  });

  it("rect is the unit square, corners in order", () => {
    expect(shapePolygon("rect")).toEqual([
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: 1, y: 1 },
      { x: 0, y: 1 },
    ]);
  });

  it("roundRect shares rect's polygon (the rounding is cosmetic, not geometric)", () => {
    expect(shapePolygon("roundRect")).toEqual(shapePolygon("rect"));
  });

  it("triangle points up, base at the bottom", () => {
    expect(shapePolygon("triangle")).toEqual([
      { x: 0.5, y: 0 },
      { x: 1, y: 1 },
      { x: 0, y: 1 },
    ]);
  });

  it("diamond's four vertices sit at the midpoints of each edge", () => {
    expect(shapePolygon("diamond")).toEqual([
      { x: 0.5, y: 0 },
      { x: 1, y: 0.5 },
      { x: 0.5, y: 1 },
      { x: 0, y: 0.5 },
    ]);
  });

  it("star5 has 10 points (5 outer, 5 inner) centred on (0.5, 0.5)", () => {
    const star = shapePolygon("star5")!;
    expect(star).toHaveLength(10);
    // Alternating radius: even indices are the outer points (r=0.5 from
    // centre), odd indices the inner ones (r=0.5*0.382).
    const dist = (p: Point) => Math.hypot(p.x - 0.5, p.y - 0.5);
    expect(dist(star[0])).toBeCloseTo(0.5, 6);
    expect(dist(star[1])).toBeCloseTo(0.5 * 0.382, 6);
    // The first (outer) point is straight up from centre.
    expect(star[0].x).toBeCloseTo(0.5, 6);
    expect(star[0].y).toBeCloseTo(0, 6);
  });

  it("rightArrow/leftArrow/upArrow/downArrow each tip toward their named direction", () => {
    expect(shapePolygon("rightArrow")!.some((p) => p.x === 1 && p.y === 0.5)).toBe(true);
    expect(shapePolygon("leftArrow")!.some((p) => p.x === 0 && p.y === 0.5)).toBe(true);
    expect(shapePolygon("upArrow")!.some((p) => p.x === 0.5 && p.y === 0)).toBe(true);
    expect(shapePolygon("downArrow")!.some((p) => p.x === 0.5 && p.y === 1)).toBe(true);
  });

  it("both callouts are a body with a downward tail, the rounded one's tail set in from the corner", () => {
    const wedge = shapePolygon("wedgeRectCallout")!;
    const round = shapePolygon("wedgeRoundRectCallout")!;
    // Same seven-vertex construction either way: a rectangular body
    // (y 0..0.75) with a downward tail poking to y=1.
    for (const polygon of [wedge, round]) {
      expect(polygon).toHaveLength(7);
      expect(Math.max(...polygon.map((p) => p.y))).toBe(1);
      expect(polygon.some((p) => p.y === 0)).toBe(true);
      expect(polygon.filter((p) => p.y === 0.75)).toHaveLength(4);
    }
    // They are NOT the same polygon: the rounded one paints its body's
    // bottom-left corner as an arc, so its tail is shifted clear of that
    // corner — a tail hard against it leaves nothing to round. The tail
    // itself is unchanged, just moved (same feet spacing, same tip offset).
    expect(round).not.toEqual(wedge);
    const shift = round[5].x - wedge[5].x;
    expect(shift).toBeGreaterThan(0);
    expect(round[4].x - wedge[4].x).toBeCloseTo(shift, 8);
    expect(round[3].x - wedge[3].x).toBeCloseTo(shift, 8);
  });
});

// --- roundedCalloutRadius / roundedCalloutOutline -----------------------------

describe("roundedCalloutRadius", () => {
  it("applies roundRect's 16.67% rule to the body, not the whole box", () => {
    // 4 wide x 4 tall: the body is 3 tall, so the shorter side is 3.
    expect(roundedCalloutRadius(4, 4)).toBeCloseTo(3 * 0.1667, 8);
  });

  it("never reaches past the bottom edge either side of the tail", () => {
    // A very tall, narrow callout: the unclamped 16.67% radius would be
    // 0.1667 of the width, still inside the tail's near foot at 0.25 —
    // clamped cases are the wide ones, where the body dwarfs the width.
    for (const [width, height] of [
      [10, 2],
      [2, 40],
      [3, 2],
      [1, 1],
    ]) {
      const radius = roundedCalloutRadius(width, height);
      const polygon = shapePolygon("wedgeRoundRectCallout")!;
      const tailLeft = polygon[5].x * width;
      const tailRight = polygon[3].x * width;
      expect(radius, `${width}x${height}`).toBeGreaterThan(0);
      expect(radius, `${width}x${height}`).toBeLessThanOrEqual(tailLeft);
      expect(radius, `${width}x${height}`).toBeLessThanOrEqual(width - tailRight);
    }
  });

  it("is zero for a degenerate box rather than negative", () => {
    expect(roundedCalloutRadius(0, 0)).toBe(0);
  });
});

describe("roundedCalloutOutline", () => {
  it("walks the body clockwise with four corner arcs and a sharp tail", () => {
    const { start, steps, radius } = roundedCalloutOutline(20, 20);
    const corners = steps.filter((step) => step.corner);
    expect(corners).toHaveLength(4);
    expect(corners.map((step) => step.corner)).toEqual([
      { x: 20, y: 0 },
      { x: 20, y: 15 },
      { x: 0, y: 15 },
      { x: 0, y: 0 },
    ]);
    // Closed: the last leg lands back on the starting tangent point.
    expect(steps[steps.length - 1].to).toEqual(start);
    // The tail's three vertices are straight legs, tip on the bottom edge.
    expect(steps.some((step) => !step.corner && step.to.y === 20)).toBe(true);
    expect(radius).toBeCloseTo(roundedCalloutRadius(20, 20), 8);
  });

  it("puts every corner's endpoint exactly one radius along the next edge", () => {
    const { steps, radius } = roundedCalloutOutline(12, 8);
    for (const step of steps) {
      if (!step.corner) {
        continue;
      }
      const distance = Math.hypot(step.to.x - step.corner.x, step.to.y - step.corner.y);
      expect(distance).toBeCloseTo(radius, 8);
    }
  });
});

// --- roundRectRadius / lineSegment / scalePolygon -----------------------------

describe("roundRectRadius", () => {
  it("is ~16.67% of the shorter side", () => {
    expect(roundRectRadius(10, 6)).toBeCloseTo(6 * 0.1667, 8);
    expect(roundRectRadius(4, 10)).toBeCloseTo(4 * 0.1667, 8);
  });

  it("is symmetric in width/height", () => {
    expect(roundRectRadius(5, 8)).toBe(roundRectRadius(8, 5));
  });
});

describe("lineSegment", () => {
  it("runs from the top-left corner to the bottom-right corner", () => {
    expect(lineSegment(10, 6)).toEqual([
      { x: 0, y: 0 },
      { x: 10, y: 6 },
    ]);
  });

  it("collapses to a horizontal segment when height is 0", () => {
    expect(lineSegment(5, 0)).toEqual([
      { x: 0, y: 0 },
      { x: 5, y: 0 },
    ]);
  });

  it("collapses to a vertical segment when width is 0", () => {
    expect(lineSegment(0, 5)).toEqual([
      { x: 0, y: 0 },
      { x: 0, y: 5 },
    ]);
  });
});

describe("scalePolygon", () => {
  it("multiplies every point by width/height independently", () => {
    const unit = [
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: 0.5, y: 1 },
    ];
    expect(scalePolygon(unit, 10, 4)).toEqual([
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 5, y: 4 },
    ]);
  });
});

// --- unrotatePoint -------------------------------------------------------------

describe("unrotatePoint", () => {
  it("is the identity when rotationDeg is 0 (fast path, no rotation math)", () => {
    const point = { x: 3, y: 7 };
    expect(unrotatePoint(point, 10, 6, 0)).toBe(point);
  });

  it("leaves the box centre fixed under any rotation", () => {
    const centre = { x: 5, y: 3 };
    for (const deg of [30, 90, 180, 270]) {
      const result = unrotatePoint(centre, 10, 6, deg);
      expect(result.x).toBeCloseTo(5, 8);
      expect(result.y).toBeCloseTo(3, 8);
    }
  });

  it("maps a page-space point back to its shape-local point for a 90-degree rotation", () => {
    // width=10, height=6, rotated 90deg: the page-space point (8, 3) maps
    // back to the unrotated local point (5, 0) — the top edge's midpoint.
    const result = unrotatePoint({ x: 8, y: 3 }, 10, 6, 90);
    expect(result.x).toBeCloseTo(5, 6);
    expect(result.y).toBeCloseTo(0, 6);
  });

  it("a 360-degree rotation returns to the original point", () => {
    const result = unrotatePoint({ x: 2, y: 4 }, 10, 6, 360);
    expect(result.x).toBeCloseTo(2, 6);
    expect(result.y).toBeCloseTo(4, 6);
  });
});

// --- pointInPolygon / pointInEllipse -------------------------------------------

describe("pointInPolygon", () => {
  const square: Point[] = [
    { x: 0, y: 0 },
    { x: 10, y: 0 },
    { x: 10, y: 10 },
    { x: 0, y: 10 },
  ];

  it("is true for a point well inside", () => {
    expect(pointInPolygon({ x: 5, y: 5 }, square)).toBe(true);
  });

  it("is false for a point well outside", () => {
    expect(pointInPolygon({ x: 20, y: 20 }, square)).toBe(false);
    expect(pointInPolygon({ x: -5, y: 5 }, square)).toBe(false);
  });

  it("handles a concave polygon (rightArrow's notch) correctly", () => {
    const polygon = scalePolygon(shapePolygon("rightArrow")!, 10, 10);
    // The notch just above the shaft, outside the arrow, between the two
    // "wings" — x=1 (inside the shaft's x-range), y=1 (above the shaft,
    // which spans y=3..7 at width 10).
    expect(pointInPolygon({ x: 1, y: 1 }, polygon)).toBe(false);
    // Inside the shaft itself.
    expect(pointInPolygon({ x: 1, y: 5 }, polygon)).toBe(true);
  });
});

describe("pointInEllipse", () => {
  it("is true at the centre", () => {
    expect(pointInEllipse({ x: 5, y: 3 }, 10, 6)).toBe(true);
  });

  it("is true exactly on the boundary (right edge)", () => {
    expect(pointInEllipse({ x: 10, y: 3 }, 10, 6)).toBe(true);
  });

  it("is false just outside the boundary", () => {
    expect(pointInEllipse({ x: 10.5, y: 3 }, 10, 6)).toBe(false);
  });

  it("is false for a corner of the bounding box", () => {
    expect(pointInEllipse({ x: 0, y: 0 }, 10, 6)).toBe(false);
  });

  it("is false for a degenerate (zero-size) ellipse", () => {
    expect(pointInEllipse({ x: 0, y: 0 }, 0, 6)).toBe(false);
    expect(pointInEllipse({ x: 0, y: 0 }, 10, 0)).toBe(false);
  });
});

// --- distanceToSegment / distanceToPolygonOutline ------------------------------

describe("distanceToSegment", () => {
  it("is 0 for a point on the segment", () => {
    expect(distanceToSegment({ x: 5, y: 5 }, { x: 0, y: 0 }, { x: 10, y: 10 })).toBeCloseTo(0, 8);
  });

  it("is the perpendicular distance for a point off a horizontal segment", () => {
    expect(distanceToSegment({ x: 5, y: 3 }, { x: 0, y: 0 }, { x: 10, y: 0 })).toBeCloseTo(3, 8);
  });

  it("clamps to the nearest endpoint beyond the segment's ends", () => {
    expect(distanceToSegment({ x: -3, y: 0 }, { x: 0, y: 0 }, { x: 10, y: 0 })).toBeCloseTo(3, 8);
    expect(distanceToSegment({ x: 13, y: 0 }, { x: 0, y: 0 }, { x: 10, y: 0 })).toBeCloseTo(3, 8);
  });

  it("falls back to point-to-point distance for a zero-length segment", () => {
    expect(distanceToSegment({ x: 3, y: 4 }, { x: 0, y: 0 }, { x: 0, y: 0 })).toBeCloseTo(5, 8);
  });
});

describe("distanceToPolygonOutline", () => {
  const square: Point[] = [
    { x: 0, y: 0 },
    { x: 10, y: 0 },
    { x: 10, y: 10 },
    { x: 0, y: 10 },
  ];

  it("is 0 on an edge", () => {
    expect(distanceToPolygonOutline({ x: 5, y: 0 }, square)).toBeCloseTo(0, 8);
  });

  it("is the distance to the nearest edge for an interior point", () => {
    // Nearest edge to (2, 5) is the left edge, x=0 — distance 2.
    expect(distanceToPolygonOutline({ x: 2, y: 5 }, square)).toBeCloseTo(2, 8);
  });

  it("closes the polygon (distance from the last vertex back to the first is considered)", () => {
    const triangle: Point[] = [
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 0, y: 10 },
    ];
    // Nearest the closing edge (from (0,10) back to (0,0), i.e. x=0) —
    // closer to it than to the hypotenuse or the bottom edge.
    expect(distanceToPolygonOutline({ x: 0.5, y: 8 }, triangle)).toBeCloseTo(0.5, 8);
  });
});

// --- HIT_TOLERANCE_CM -----------------------------------------------------------

describe("HIT_TOLERANCE_CM", () => {
  it("is 3pt converted to cm", () => {
    expect(HIT_TOLERANCE_CM).toBeCloseTo((3 * 2.54) / 72, 10);
  });
});

// --- pointHitsShape --------------------------------------------------------------

describe("pointHitsShape — every catalogue area geometry (filled/unfilled/outline/text)", () => {
  const width = 10;
  const height = 6;

  for (const geometry of POLYGON_GEOMETRIES) {
    describe(geometry, () => {
      const polygon = scalePolygon(shapePolygon(geometry)!, width, height);
      const centre = { x: width / 2, y: height / 2 };
      const farOutside = { x: width * 3, y: height * 3 };
      const onOutline = polygon[0];

      it("a filled shape is hit at its centre", () => {
        expect(
          pointHitsShape(centre, {
            geometry,
            widthCm: width,
            heightCm: height,
            filled: true,
            hasText: false,
            outlineToleranceCm: HIT_TOLERANCE_CM,
          }),
        ).toBe(true);
      });

      it("a filled shape is not hit far outside its bounds", () => {
        expect(
          pointHitsShape(farOutside, {
            geometry,
            widthCm: width,
            heightCm: height,
            filled: true,
            hasText: false,
            outlineToleranceCm: HIT_TOLERANCE_CM,
          }),
        ).toBe(false);
      });

      it("an unfilled, textless shape is NOT hit at its centre (interior, no fill)", () => {
        expect(
          pointHitsShape(centre, {
            geometry,
            widthCm: width,
            heightCm: height,
            filled: false,
            hasText: false,
            outlineToleranceCm: HIT_TOLERANCE_CM,
          }),
        ).toBe(false);
      });

      it("an unfilled shape IS hit exactly on its outline", () => {
        expect(
          pointHitsShape(onOutline, {
            geometry,
            widthCm: width,
            heightCm: height,
            filled: false,
            hasText: false,
            outlineToleranceCm: HIT_TOLERANCE_CM,
          }),
        ).toBe(true);
      });

      it("an unfilled shape with hasText IS hit at its centre (the text block's bounds)", () => {
        expect(
          pointHitsShape(centre, {
            geometry,
            widthCm: width,
            heightCm: height,
            filled: false,
            hasText: true,
            outlineToleranceCm: HIT_TOLERANCE_CM,
          }),
        ).toBe(true);
      });
    });
  }
});

describe("pointHitsShape — ellipse", () => {
  const width = 10;
  const height = 6;

  it("a filled ellipse is hit at its centre", () => {
    expect(
      pointHitsShape(
        { x: 5, y: 3 },
        { geometry: "ellipse", widthCm: width, heightCm: height, filled: true, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
      ),
    ).toBe(true);
  });

  it("a filled ellipse is not hit at a bounding-box corner (outside the inscribed ellipse)", () => {
    expect(
      pointHitsShape(
        { x: 0, y: 0 },
        { geometry: "ellipse", widthCm: width, heightCm: height, filled: true, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
      ),
    ).toBe(false);
  });

  it("an unfilled ellipse is hit near its right-edge boundary", () => {
    expect(
      pointHitsShape(
        { x: width, y: height / 2 },
        { geometry: "ellipse", widthCm: width, heightCm: height, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
      ),
    ).toBe(true);
  });

  it("an unfilled ellipse is not hit at its centre", () => {
    expect(
      pointHitsShape(
        { x: 5, y: 3 },
        { geometry: "ellipse", widthCm: width, heightCm: height, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
      ),
    ).toBe(false);
  });

  it("an unfilled ellipse with hasText IS hit at its centre", () => {
    expect(
      pointHitsShape(
        { x: 5, y: 3 },
        { geometry: "ellipse", widthCm: width, heightCm: height, filled: false, hasText: true, outlineToleranceCm: HIT_TOLERANCE_CM },
      ),
    ).toBe(true);
  });

  it("degenerate (zero) dimensions never hit, filled or not", () => {
    const opts = { geometry: "ellipse" as ShapeGeometry, widthCm: 0, heightCm: 0, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM };
    expect(pointHitsShape({ x: 0, y: 0 }, { ...opts, filled: true })).toBe(false);
    expect(pointHitsShape({ x: 0, y: 0 }, { ...opts, filled: false })).toBe(false);
  });
});

describe("pointHitsShape — line and arrow (segment hit testing, LINE_GEOMETRIES)", () => {
  it("LINE_GEOMETRIES contains exactly line and arrow", () => {
    expect(LINE_GEOMETRIES.has("line")).toBe(true);
    expect(LINE_GEOMETRIES.has("arrow")).toBe(true);
    expect(LINE_GEOMETRIES.has("rect")).toBe(false);
  });

  for (const geometry of ["line", "arrow"] as const) {
    describe(geometry, () => {
      it("is hit exactly on the diagonal segment", () => {
        expect(
          pointHitsShape(
            { x: 5, y: 3 },
            { geometry, widthCm: 10, heightCm: 6, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
          ),
        ).toBe(true);
      });

      it("is hit within tolerance just off the segment", () => {
        const tolerance = HIT_TOLERANCE_CM;
        expect(
          pointHitsShape(
            { x: 5, y: 3 + tolerance / 2 },
            { geometry, widthCm: 10, heightCm: 6, filled: false, hasText: false, outlineToleranceCm: tolerance },
          ),
        ).toBe(true);
      });

      it("is not hit far off the segment", () => {
        expect(
          pointHitsShape(
            { x: 5, y: 5.9 },
            { geometry, widthCm: 10, heightCm: 6, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
          ),
        ).toBe(false);
      });

      it("`filled: true` makes no difference — a line is never a filled region", () => {
        expect(
          pointHitsShape(
            { x: 5, y: 5.9 },
            { geometry, widthCm: 10, heightCm: 6, filled: true, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
          ),
        ).toBe(false);
      });

      it("handles a degenerate horizontal line (heightCm 0)", () => {
        expect(
          pointHitsShape(
            { x: 5, y: 0 },
            { geometry, widthCm: 10, heightCm: 0, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
          ),
        ).toBe(true);
        expect(
          pointHitsShape(
            { x: 5, y: 1 },
            { geometry, widthCm: 10, heightCm: 0, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM },
          ),
        ).toBe(false);
      });
    });
  }
});

describe("pointHitsShape — rotation", () => {
  it("rotates the hit point into the shape's local frame before testing (90deg rect)", () => {
    // width=10, height=6, rotated 90deg: page-space point (8,3) maps to the
    // local top-edge midpoint (5,0) — on the outline.
    const opts = { geometry: "rect" as ShapeGeometry, widthCm: 10, heightCm: 6, rotationDeg: 90, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM };
    expect(pointHitsShape({ x: 8, y: 3 }, opts)).toBe(true);
  });

  it("the same page-space point does not hit the unrotated shape's outline", () => {
    const opts = { geometry: "rect" as ShapeGeometry, widthCm: 10, heightCm: 6, filled: false, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM };
    expect(pointHitsShape({ x: 8, y: 3 }, opts)).toBe(false);
  });

  it("a filled shape's centre is hit regardless of rotation (the centre never moves)", () => {
    for (const rotationDeg of [0, 45, 90, 180, 270]) {
      const opts = { geometry: "rect" as ShapeGeometry, widthCm: 10, heightCm: 6, rotationDeg, filled: true, hasText: false, outlineToleranceCm: HIT_TOLERANCE_CM };
      expect(pointHitsShape({ x: 5, y: 3 }, opts)).toBe(true);
    }
  });
});
