import { describe, expect, it } from "vitest";
import {
  CONNECTION_SITE_COUNT,
  clampSite,
  clipCentreSite,
  connectionSitePoint,
  connectionSites,
  elbowAvoidsBox,
  nearestSite,
  routeConnector,
  routedPathToSvg,
  segmentPolygonCrossing,
  type ConnectorBox,
} from "./connectors";

const BOX: ConnectorBox = { x: 10, y: 20, w: 100, h: 40 };

describe("connectionSites", () => {
  it("derives nine sites in the fixed on-disk index order", () => {
    const sites = connectionSites(BOX);
    expect(sites).toHaveLength(CONNECTION_SITE_COUNT);
    expect(sites[0].point).toEqual({ x: 60, y: 20 }); // top mid
    expect(sites[1].point).toEqual({ x: 110, y: 40 }); // right mid
    expect(sites[2].point).toEqual({ x: 60, y: 60 }); // bottom mid
    expect(sites[3].point).toEqual({ x: 10, y: 40 }); // left mid
    expect(sites[4].point).toEqual({ x: 10, y: 20 }); // top-left
    expect(sites[6].point).toEqual({ x: 110, y: 60 }); // bottom-right
    expect(sites[8].point).toEqual({ x: 60, y: 40 }); // centre
  });

  it("gives each edge site an outward normal", () => {
    const sites = connectionSites(BOX);
    expect(sites[0].normal).toEqual({ x: 0, y: -1 });
    expect(sites[1].normal).toEqual({ x: 1, y: 0 });
    expect(sites[2].normal).toEqual({ x: 0, y: 1 });
    expect(sites[3].normal).toEqual({ x: -1, y: 0 });
  });

  it("rotates sites about the box centre", () => {
    const rotated = connectionSitePoint({ ...BOX, rotationDeg: 90 }, 1); // right-mid
    // centre is (60,40); right-mid (110,40) rotates 90° CW to bottom (60,90).
    expect(rotated.x).toBeCloseTo(60, 6);
    expect(rotated.y).toBeCloseTo(90, 6);
  });
});

describe("clampSite / nearestSite", () => {
  it("falls unmappable indices back to the centre", () => {
    expect(clampSite(-1)).toBe(8);
    expect(clampSite(9)).toBe(8);
    expect(clampSite(2.5)).toBe(8);
    expect(clampSite(3)).toBe(3);
  });

  it("finds the nearest site to a point", () => {
    expect(nearestSite(BOX, { x: 60, y: 19 })).toBe(0); // just above top mid
    expect(nearestSite(BOX, { x: 111, y: 41 })).toBe(1); // near right mid
  });
});

describe("routeConnector", () => {
  it("straight routes a single segment", () => {
    const path = routeConnector("straight", { point: { x: 0, y: 0 } }, { point: { x: 100, y: 50 } });
    expect(path.points).toEqual([
      { x: 0, y: 0 },
      { x: 100, y: 50 },
    ]);
    expect(routedPathToSvg(path)).toBe("M 0 0 L 100 50");
  });

  it("elbow routes orthogonal segments leaving each horizontal site", () => {
    const path = routeConnector(
      "elbow",
      { point: { x: 0, y: 0 }, normal: { x: 1, y: 0 } },
      { point: { x: 100, y: 60 }, normal: { x: -1, y: 0 } },
      0.5,
    );
    // vertical crossing at the midpoint x=50
    expect(path.points).toEqual([
      { x: 0, y: 0 },
      { x: 50, y: 0 },
      { x: 50, y: 60 },
      { x: 100, y: 60 },
    ]);
    // every segment is axis-aligned
    for (let i = 1; i < path.points.length; i++) {
      const a = path.points[i - 1];
      const b = path.points[i];
      expect(a.x === b.x || a.y === b.y).toBe(true);
    }
  });

  it("elbow honours the adjustment fraction", () => {
    const path = routeConnector(
      "elbow",
      { point: { x: 0, y: 0 }, normal: { x: 1, y: 0 } },
      { point: { x: 100, y: 60 }, normal: { x: -1, y: 0 } },
      0.25,
    );
    expect(path.points[1].x).toBe(25);
  });

  it("elbow avoids running back through either shape", () => {
    const from = { point: { x: 100, y: 20 }, normal: { x: 1, y: 0 } };
    const to = { point: { x: 300, y: 20 }, normal: { x: -1, y: 0 } };
    const path = routeConnector("elbow", from, to, 0.5);
    const boxA: ConnectorBox = { x: 0, y: 0, w: 100, h: 40 };
    const boxB: ConnectorBox = { x: 300, y: 0, w: 100, h: 40 };
    expect(elbowAvoidsBox(path.points, boxA)).toBe(true);
    expect(elbowAvoidsBox(path.points, boxB)).toBe(true);
  });

  it("curved matches endpoint directions with cubic controls", () => {
    const path = routeConnector(
      "curved",
      { point: { x: 0, y: 0 }, normal: { x: 1, y: 0 } },
      { point: { x: 100, y: 0 }, normal: { x: -1, y: 0 } },
    );
    expect(path.controls).toBeDefined();
    const [c1, c2] = path.controls!;
    // control 1 extends rightward from the start, control 2 leftward from end
    expect(c1.x).toBeGreaterThan(0);
    expect(c1.y).toBe(0);
    expect(c2.x).toBeLessThan(100);
    expect(routedPathToSvg(path)).toMatch(/^M 0 0 C /);
  });
});

describe("centre-site clipping", () => {
  const square: { x: number; y: number }[] = [
    { x: 0, y: 0 },
    { x: 40, y: 0 },
    { x: 40, y: 40 },
    { x: 0, y: 40 },
  ];

  it("crosses the outline nearest the start", () => {
    const hit = segmentPolygonCrossing({ x: 20, y: 20 }, { x: 200, y: 20 }, square);
    expect(hit).toEqual({ x: 40, y: 20 });
  });

  it("clips a centre endpoint to the outline toward the other end", () => {
    const clipped = clipCentreSite({ x: 20, y: 20 }, { x: 200, y: 20 }, square);
    expect(clipped.point).toEqual({ x: 40, y: 20 });
    expect(clipped.normal).toEqual({ x: 1, y: 0 });
  });

  it("falls back to the centre when the ray misses", () => {
    const clipped = clipCentreSite({ x: 20, y: 20 }, { x: 20, y: 20 }, square);
    expect(clipped.point).toEqual({ x: 20, y: 20 });
  });
});
