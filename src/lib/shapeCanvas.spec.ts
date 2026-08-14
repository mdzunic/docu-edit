// SPEC-051 regression: paintShape() must never invent an outline a Shape
// doesn't have. A Line/Arrow with no fill of its own relies entirely on
// shape.outline to be visible at all, so "no outline" has to mean nothing
// is painted, not a hardcoded fallback.
import { describe, it, expect } from "vitest";
import { paintShape } from "./shapeCanvas";
import type { Shape } from "./documentModel";

/** A minimal CanvasRenderingContext2D fake that records the calls paintShape
 * makes — real canvas painting isn't available in jsdom, and this project
 * has no canvas-mocking dependency, so this stands in for just the surface
 * paintShape actually uses. */
function fakeCtx(): CanvasRenderingContext2D {
  const calls: string[] = [];
  const ctx = {
    calls,
    save: () => calls.push("save"),
    restore: () => calls.push("restore"),
    translate: () => {},
    rotate: () => {},
    beginPath: () => calls.push("beginPath"),
    closePath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    arcTo: () => calls.push("arcTo"),
    stroke: () => calls.push("stroke"),
    fill: () => calls.push("fill"),
    ellipse: () => {},
    strokeRect: () => calls.push("strokeRect"),
    setLineDash: () => {},
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function baseLine(overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "line",
    widthCm: 3,
    heightCm: 0,
    wrap: "inline",
    ...overrides,
  };
}

describe("paintShape — Line/Arrow outline (AC-15)", () => {
  it("paints nothing when outline is absent (removed via 'No outline')", () => {
    const ctx = fakeCtx() as unknown as { calls: string[] };
    paintShape(ctx as unknown as CanvasRenderingContext2D, baseLine(), 100, 0);
    expect(ctx.calls).not.toContain("stroke");
  });

  it("strokes with the shape's own outline when present", () => {
    const ctx = fakeCtx() as unknown as { calls: string[] };
    paintShape(
      ctx as unknown as CanvasRenderingContext2D,
      baseLine({ outline: { color: "#ff0000", widthPt: 2, dash: "solid" } }),
      100,
      0,
    );
    expect(ctx.calls).toContain("stroke");
  });

  it("an Arrow with no outline also paints nothing (no arrowhead fallback)", () => {
    const ctx = fakeCtx() as unknown as { calls: string[] };
    paintShape(ctx as unknown as CanvasRenderingContext2D, baseLine({ geometry: "arrow" }), 100, 0);
    expect(ctx.calls).not.toContain("stroke");
    expect(ctx.calls).not.toContain("fill");
  });
});

/** The calls paintShape makes for a filled 3cm x 2cm callout of `geometry`. */
function paintCalls(geometry: Shape["geometry"]): string[] {
  const shape: Shape = { type: "shape", geometry, widthCm: 3, heightCm: 2, wrap: "square", fill: "#ffffff" };
  const ctx = fakeCtx() as unknown as { calls: string[] };
  paintShape(ctx as unknown as CanvasRenderingContext2D, shape, 120, 80);
  return ctx.calls;
}

describe("paintShape — the two callouts are painted differently", () => {
  it("rounds wedgeRoundRectCallout's body corners", () => {
    expect(paintCalls("wedgeRoundRectCallout").filter((call) => call === "arcTo")).toHaveLength(4);
  });

  it("leaves wedgeRectCallout's corners square", () => {
    expect(paintCalls("wedgeRectCallout")).not.toContain("arcTo");
  });
});
