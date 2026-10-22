import { describe, expect, it } from "vitest";
import {
  buildConnectorPaint,
  connectorDashArray,
  connectorPaintBounds,
  connectorStrokeWidthPx,
  DEFAULT_CONNECTOR_STROKE,
  scaleRoutedPath,
} from "./connectorPaint";
import { arrowheadOutline, routeConnector, routedPathEnds, translateRoutedPath } from "./connectors";
import type { ArrowheadKind } from "./shapes";

const from = { point: { x: 10, y: 40 }, normal: { x: 1, y: 0 } };
const to = { point: { x: 110, y: 40 }, normal: { x: -1, y: 0 } };
const straight = routeConnector("straight", from, to);

describe("arrowheadOutline", () => {
  const along = { x: 1, y: 0 };

  it("is null for `none` and for an absent kind", () => {
    expect(arrowheadOutline("none", { x: 0, y: 0 }, along, 2)).toBeNull();
    expect(arrowheadOutline(undefined, { x: 0, y: 0 }, along, 2)).toBeNull();
  });

  it("is null for a degenerate direction or a zero stroke", () => {
    expect(arrowheadOutline("triangle", { x: 0, y: 0 }, { x: 0, y: 0 }, 2)).toBeNull();
    expect(arrowheadOutline("triangle", { x: 0, y: 0 }, along, 0)).toBeNull();
  });

  it("puts a triangle's tip at the endpoint and its base back along the line", () => {
    const head = arrowheadOutline("triangle", { x: 100, y: 50 }, along, 2);
    expect(head?.filled).toBe(true);
    expect(head?.points[0]).toEqual({ x: 100, y: 50 });
    // Base 6.3 x stroke back, 3.5 x stroke either side (the marker geometry
    // the screen overlay used before arrowheads became real paths).
    expect(head?.points[1].x).toBeCloseTo(100 - 12.6);
    expect(head?.points[1].y).toBeCloseTo(50 + 7);
    expect(head?.points[2].y).toBeCloseTo(50 - 7);
  });

  it("gives every kind a distinct outline, and only `open` is stroked", () => {
    const kinds: ArrowheadKind[] = ["triangle", "open", "stealth", "diamond", "oval"];
    const shapes = kinds.map((kind) => arrowheadOutline(kind, { x: 0, y: 0 }, along, 2));
    expect(shapes.every((shape) => shape !== null)).toBe(true);
    expect(shapes.map((shape) => shape!.filled)).toEqual([true, false, true, true, true]);
    const serialized = shapes.map((shape) => JSON.stringify(shape!.points));
    expect(new Set(serialized).size).toBe(kinds.length);
  });

  it("rotates with the direction it is given", () => {
    const down = arrowheadOutline("triangle", { x: 0, y: 0 }, { x: 0, y: 1 }, 2);
    expect(down?.points[1].y).toBeCloseTo(-12.6);
    expect(Math.abs(down!.points[1].x)).toBeCloseTo(7);
  });
});

describe("routedPathEnds", () => {
  it("points each end outward along the line", () => {
    const ends = routedPathEnds(straight);
    expect(ends?.start.along).toEqual({ x: -1, y: 0 });
    expect(ends?.end.along).toEqual({ x: 1, y: 0 });
  });

  it("takes a curve's direction from its control points, not its endpoints", () => {
    const curved = routeConnector("curved", from, to);
    const ends = routedPathEnds(curved);
    // The control points sit along each site's outward normal.
    expect(ends?.start.along.x).toBeCloseTo(-1);
    expect(ends?.end.along.x).toBeCloseTo(1);
  });

  it("is null for a path with no length to take a direction from", () => {
    expect(routedPathEnds({ kind: "straight", points: [{ x: 1, y: 1 }] })).toBeNull();
  });
});

describe("buildConnectorPaint", () => {
  it("carries the outline's colour, width and dash", () => {
    const paint = buildConnectorPaint("0:1", straight, {
      outline: { color: "#ff0000", widthPt: 3, dash: "dash" },
      endArrow: "triangle",
    });
    expect(paint.key).toBe("0:1");
    expect(paint.color).toBe("#ff0000");
    expect(paint.widthPx).toBeCloseTo(4);
    expect(paint.dash).toBe("6 4");
    expect(paint.d).toContain("M 10 40");
  });

  it("falls back to the default stroke when there is no outline", () => {
    expect(buildConnectorPaint("k", straight, {}).color).toBe(DEFAULT_CONNECTOR_STROKE);
    expect(connectorStrokeWidthPx(undefined)).toBeCloseTo(96 / 72);
    expect(connectorDashArray("dot")).toBe("1 3");
    expect(connectorDashArray("solid")).toBe("");
  });

  it("emits one head per non-`none` arrowhead, start first", () => {
    expect(buildConnectorPaint("k", straight, {}).heads).toHaveLength(0);
    expect(buildConnectorPaint("k", straight, { endArrow: "triangle" }).heads).toHaveLength(1);
    const both = buildConnectorPaint("k", straight, { startArrow: "triangle", endArrow: "triangle" });
    expect(both.heads).toHaveLength(2);
    // Each head's tip is its own end of the line, and the two point apart.
    expect(both.heads[0].d.startsWith("M 10 40")).toBe(true);
    expect(both.heads[1].d.startsWith("M 110 40")).toBe(true);
    expect(both.heads.every((head) => head.filled)).toBe(true);
    // An oval head is a circle tangent to the endpoint, so its own `d` starts
    // elsewhere on the circle but still touches it.
    const oval = buildConnectorPaint("k", straight, { startArrow: "oval" });
    expect(oval.heads[0].d).toContain("10 40");
  });

  it("marks an `open` head as stroked, not filled, and leaves it unclosed", () => {
    const [head] = buildConnectorPaint("k", straight, { endArrow: "open" }).heads;
    expect(head.filled).toBe(false);
    expect(head.d.endsWith("Z")).toBe(false);
  });
});

describe("connectorPaintBounds", () => {
  it("grows a horizontal line's zero-height box by the stroke and the head", () => {
    const bounds = connectorPaintBounds(straight, {
      outline: { color: "#000000", widthPt: 3, dash: "solid" },
      endArrow: "triangle",
    });
    expect(bounds.left).toBeCloseTo(10 - 2);
    expect(bounds.right).toBeCloseTo(110 + 2);
    // 4px stroke → 3.5 x 4 = 14 half-width, plus the 2px stroke pad.
    expect(bounds.top).toBeCloseTo(40 - 16);
    expect(bounds.bottom).toBeCloseTo(40 + 16);
  });

  it("includes a curve's control points, which the curve bends toward", () => {
    const curved = routeConnector(
      "curved",
      { point: { x: 0, y: 0 }, normal: { x: 0, y: -1 } },
      { point: { x: 100, y: 0 }, normal: { x: 0, y: -1 } },
    );
    expect(connectorPaintBounds(curved, {}).top).toBeLessThan(-10);
  });
});

describe("scaleRoutedPath / translateRoutedPath", () => {
  it("scales points and control points together", () => {
    const curved = routeConnector("curved", from, to);
    const scaled = scaleRoutedPath(curved, 2);
    expect(scaled.points[0]).toEqual({ x: 20, y: 80 });
    expect(scaled.controls![0].x).toBeCloseTo(curved.controls![0].x * 2);
  });

  it("moves a path into another frame without changing its shape", () => {
    const moved = translateRoutedPath(straight, -10, -40);
    expect(moved.points[0]).toEqual({ x: 0, y: 0 });
    expect(moved.points[1]).toEqual({ x: 100, y: 0 });
  });
});
