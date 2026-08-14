// SPEC-051 AC-1: the "Insert -> Shape" gallery shows exactly the catalogue's
// geometries. Now that a gallery cell shows the geometry itself rather than
// its name, "the right shape" is a property of the emitted path — these
// tests pin it to the same vertex lists the page canvas paints with, so an
// icon cannot quietly stop matching what picking it draws.
import { describe, expect, it } from "vitest";
import { LINE_GEOMETRIES, SHAPE_CATALOGUE } from "./documentModel";
import { roundRectRadius, shapePolygon } from "./shapeGeometry";
import { SHAPE_ICON_SIZE, shapeIconPath } from "./shapeIcons";

const GEOMETRIES = SHAPE_CATALOGUE.map((entry) => entry.geometry);

/** Every vertex in a path's `d`, as [x, y] numbers. An arc's leading
 * radii/flags are stripped first — they are written like a coordinate pair
 * but are not a point on the outline. */
function pathPoints(d: string): [number, number][] {
  const outline = d.replace(/A [\d.]+,[\d.]+ [01] [01] [01] /g, "");
  return [...outline.matchAll(/(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/g)].map(
    (match) => [Number(match[1]), Number(match[2])] as [number, number],
  );
}

/** The inset and side of the box the geometries are scaled into, recovered
 * from the one icon whose own polygon spans the full unit square. */
function drawingBox(): { inset: number; side: number } {
  const xs = pathPoints(shapeIconPath("rect")).map(([x]) => x);
  return { inset: Math.min(...xs), side: Math.max(...xs) - Math.min(...xs) };
}

describe("shapeIconPath", () => {
  it("emits a path for every catalogue geometry", () => {
    for (const geometry of GEOMETRIES) {
      expect(shapeIconPath(geometry), geometry).toMatch(/^M /);
    }
  });

  it("keeps every icon inside the viewBox", () => {
    for (const geometry of GEOMETRIES) {
      for (const [x, y] of pathPoints(shapeIconPath(geometry))) {
        expect(x, `${geometry} x`).toBeGreaterThanOrEqual(0);
        expect(x, `${geometry} x`).toBeLessThanOrEqual(SHAPE_ICON_SIZE);
        expect(y, `${geometry} y`).toBeGreaterThanOrEqual(0);
        expect(y, `${geometry} y`).toBeLessThanOrEqual(SHAPE_ICON_SIZE);
      }
    }
  });

  it("draws each polygon geometry from shapePolygon's own vertices, in order", () => {
    const { inset, side } = drawingBox();
    for (const geometry of GEOMETRIES) {
      const polygon = shapePolygon(geometry);
      // roundRect and the rounded callout are the two the canvas paints
      // with corner arcs rather than straight through their own vertex
      // lists — each asserted on its own below.
      if (!polygon || geometry === "roundRect" || geometry === "wedgeRoundRectCallout") {
        continue;
      }
      const points = pathPoints(shapeIconPath(geometry));
      expect(points, geometry).toHaveLength(polygon.length);
      points.forEach(([x, y], index) => {
        expect(x, `${geometry} vertex ${index} x`).toBeCloseTo(inset + polygon[index].x * side, 1);
        expect(y, `${geometry} vertex ${index} y`).toBeCloseTo(inset + polygon[index].y * side, 1);
      });
      expect(shapeIconPath(geometry).endsWith(" Z"), `${geometry} closes`).toBe(true);
    }
  });

  it("rounds the rounded rectangle by the same radius the canvas uses", () => {
    const { side } = drawingBox();
    const arcs = [...shapeIconPath("roundRect").matchAll(/A ([\d.]+),([\d.]+)/g)];
    expect(arcs).toHaveLength(4);
    for (const arc of arcs) {
      expect(Number(arc[1])).toBeCloseTo(roundRectRadius(side, side), 2);
      expect(Number(arc[2])).toBeCloseTo(roundRectRadius(side, side), 2);
    }
  });

  it("tells the two callouts apart — the rounded one is the one with corner arcs", () => {
    const wedge = shapeIconPath("wedgeRectCallout");
    const round = shapeIconPath("wedgeRoundRectCallout");
    expect(round).not.toBe(wedge);
    expect(wedge).not.toContain("A ");
    const arcs = [...round.matchAll(/A ([\d.]+),([\d.]+)/g)];
    expect(arcs).toHaveLength(4);
    // Big enough to read at icon size — the whole point of the two cells
    // not looking identical in the gallery.
    for (const arc of arcs) {
      expect(Number(arc[1])).toBeGreaterThan(1.5);
    }
    // Both still hang a tail off the bottom of the drawing box.
    const { inset, side } = drawingBox();
    for (const d of [wedge, round]) {
      expect(pathPoints(d).some(([, y]) => y === inset + side)).toBe(true);
    }
  });

  it("draws the ellipse as a closed pair of half-arcs, not a polygon", () => {
    const d = shapeIconPath("ellipse");
    const { inset, side } = drawingBox();
    expect(d).not.toContain("L ");
    const arcs = [...d.matchAll(/A ([\d.]+),([\d.]+)/g)];
    expect(arcs).toHaveLength(2);
    for (const arc of arcs) {
      expect(Number(arc[1])).toBeCloseTo(side / 2, 2);
      expect(Number(arc[2])).toBeCloseTo(side / 2, 2);
    }
    // Both arcs meet on the ellipse's horizontal axis, at the box's sides.
    for (const [x, y] of pathPoints(d)) {
      expect(y).toBeCloseTo(inset + side / 2, 2);
      expect(x === inset || x === inset + side).toBe(true);
    }
    expect(d.endsWith(" Z")).toBe(true);
  });

  it("leaves a line open and caps an arrow with a single filled head", () => {
    // The segment subpath stays open so a `fill` on the element paints the
    // arrowhead alone (ShapeGallery.vue strokes both).
    const line = shapeIconPath("line");
    expect(line).not.toContain("Z");
    expect(line.match(/M /g)).toHaveLength(1);

    const arrow = shapeIconPath("arrow");
    // The head hangs off the far end of that same corner-to-corner segment.
    expect(arrow.startsWith(line)).toBe(true);
    expect(arrow.match(/M /g)).toHaveLength(2);
    expect(arrow.match(/Z/g)).toHaveLength(1);
  });

  it("gives every non-line geometry one closed subpath, and only those", () => {
    for (const geometry of GEOMETRIES) {
      const d = shapeIconPath(geometry);
      if (LINE_GEOMETRIES.has(geometry)) {
        expect(d.startsWith(shapeIconPath("line")), geometry).toBe(true);
        continue;
      }
      expect(d.match(/M /g), geometry).toHaveLength(1);
      expect(d.endsWith(" Z"), geometry).toBe(true);
    }
  });
});
