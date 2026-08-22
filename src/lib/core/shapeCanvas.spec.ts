// SPEC-051 regression: paintShape() must never invent an outline a Shape
// doesn't have. A Line/Arrow with no fill of its own relies entirely on
// shape.outline to be visible at all, so "no outline" has to mean nothing
// is painted, not a hardcoded fallback.
import { describe, it, expect } from "vitest";
import { paintShape, paintShapeIntoCanvas } from "./shapeCanvas";
import type { Shape } from "../documentModel";

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

// SPEC-063 AC-20: the backing store scales with devicePixelRatio * zoom while
// the CSS size stays the model px, so the .pages transform can scale the
// canvas without blurring. jsdom reports devicePixelRatio == 1, isolating the
// zoom factor. (width/height/style are set before getContext, which jsdom
// returns null for, so the sizing is observable without a real 2d context.)
function rectShape(): Shape {
  return { type: "shape", geometry: "rect", widthCm: 4, heightCm: 3, wrap: "inline" };
}

describe("paintShapeIntoCanvas — zoom-scaled backing store (AC-20)", () => {
  it("keeps the CSS size but multiplies the backing store by the zoom", () => {
    const at100 = document.createElement("canvas");
    paintShapeIntoCanvas(at100, rectShape(), 200, 100, 1);
    const at200 = document.createElement("canvas");
    paintShapeIntoCanvas(at200, rectShape(), 200, 100, 2);

    // CSS size is unchanged by zoom — the transform provides the on-screen scale.
    expect(at100.style.width).toBe("200px");
    expect(at200.style.width).toBe("200px");
    // Backing store doubles from 100% to 200%.
    expect(at200.width).toBe(at100.width * 2);
    expect(at200.height).toBe(at100.height * 2);
  });

  it("defaults to no extra scaling when the zoom is omitted", () => {
    const explicit = document.createElement("canvas");
    paintShapeIntoCanvas(explicit, rectShape(), 200, 100, 1);
    const implicit = document.createElement("canvas");
    paintShapeIntoCanvas(implicit, rectShape(), 200, 100);
    expect(implicit.width).toBe(explicit.width);
  });
});
