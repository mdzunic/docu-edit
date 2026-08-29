import { describe, expect, it } from "vitest";
import { parseSvgPath } from "./svgPath";
import { pointsToSvg, routeConnector, routedPathToSvg } from "../core/connectors";

describe("parseSvgPath", () => {
  it("reads the poly-line form routedPathToSvg emits", () => {
    const subpaths = parseSvgPath("M 10 20 L 30 20 L 30 60 L 50 60");
    expect(subpaths).not.toBeNull();
    expect(subpaths).toHaveLength(1);
    expect(subpaths![0].start).toEqual({ x: 10, y: 20 });
    expect(subpaths![0].segments.map((segment) => segment.to)).toEqual([
      { x: 30, y: 20 },
      { x: 30, y: 60 },
      { x: 50, y: 60 },
    ]);
    expect(subpaths![0].closed).toBe(false);
  });

  it("reads the cubic form a curved connector emits", () => {
    const subpaths = parseSvgPath("M 0 0 C 10 0 20 10 20 10");
    expect(subpaths).toHaveLength(1);
    const [segment] = subpaths![0].segments;
    expect(segment.c1).toEqual({ x: 10, y: 0 });
    expect(segment.c2).toEqual({ x: 20, y: 10 });
    expect(segment.to).toEqual({ x: 20, y: 10 });
  });

  it("reads a closed arrowhead polygon", () => {
    const subpaths = parseSvgPath("M 100 50 L 90 46 L 90 54 Z");
    expect(subpaths![0].closed).toBe(true);
    expect(subpaths![0].segments).toHaveLength(2);
  });

  it("handles relative commands, H/V and separator-free numbers", () => {
    const subpaths = parseSvgPath("M0 0l10 0v5h-4z");
    expect(subpaths![0].segments.map((segment) => segment.to)).toEqual([
      { x: 10, y: 0 },
      { x: 10, y: 5 },
      { x: 6, y: 5 },
    ]);
    expect(subpaths![0].closed).toBe(true);
  });

  it("treats repeated coordinates after a move as line-tos, per the SVG spec", () => {
    const subpaths = parseSvgPath("M 0 0 5 5 10 0");
    expect(subpaths![0].segments.map((segment) => segment.to)).toEqual([
      { x: 5, y: 5 },
      { x: 10, y: 0 },
    ]);
  });

  it("returns separate sub-paths for a multi-move `d`", () => {
    const subpaths = parseSvgPath("M 0 0 L 1 1 M 5 5 L 6 6");
    expect(subpaths).toHaveLength(2);
    expect(subpaths![1].start).toEqual({ x: 5, y: 5 });
  });

  it("is null for commands it does not model, so the caller skips the path", () => {
    // Arcs, quadratics and smooth curves would be drawn wrong, not partially.
    expect(parseSvgPath("M 0 0 A 5 5 0 0 1 10 10")).toBeNull();
    expect(parseSvgPath("M 0 0 Q 5 5 10 0")).toBeNull();
    expect(parseSvgPath("M 0 0 S 5 5 10 0")).toBeNull();
  });

  it("is null for malformed data (a draw before any move, a short command)", () => {
    expect(parseSvgPath("L 10 10")).toBeNull();
    expect(parseSvgPath("M 0 0 L 10")).toBeNull();
    expect(parseSvgPath("10 10")).toBeNull();
  });

  it("parses an empty `d` as no sub-paths rather than a failure", () => {
    expect(parseSvgPath("")).toEqual([]);
    expect(parseSvgPath("   ")).toEqual([]);
  });

  it("round-trips every routing kind's own serialization", () => {
    for (const kind of ["straight", "elbow", "curved"] as const) {
      const routed = routeConnector(
        kind,
        { point: { x: 10, y: 10 }, normal: { x: 1, y: 0 } },
        { point: { x: 90, y: 70 }, normal: { x: -1, y: 0 } },
      );
      const subpaths = parseSvgPath(routedPathToSvg(routed));
      expect(subpaths, kind).not.toBeNull();
      expect(subpaths![0].start).toEqual(routed.points[0]);
      const last = subpaths![0].segments[subpaths![0].segments.length - 1];
      expect(last.to).toEqual(routed.points[routed.points.length - 1]);
    }
  });

  it("round-trips pointsToSvg's closed and open forms", () => {
    const closed = parseSvgPath(pointsToSvg([{ x: 0, y: 0 }, { x: 3, y: 1 }, { x: 0, y: 2 }], true));
    expect(closed![0].closed).toBe(true);
    const open = parseSvgPath(pointsToSvg([{ x: 0, y: 0 }, { x: 3, y: 1 }, { x: 0, y: 2 }], false));
    expect(open![0].closed).toBe(false);
  });
});
