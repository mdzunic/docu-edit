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
  const shape: Shape = { type: "shape", geometry, widthCm: 3, heightCm: 2, wrap: "square", fill: { kind: "solid", color: "#ffffff" } };
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

describe("paintShape — path (SPEC-071)", () => {
  function pathCtx() {
    const calls: string[] = [];
    const ctx = {
      calls,
      save: () => calls.push("save"),
      restore: () => calls.push("restore"),
      translate: () => {},
      rotate: () => {},
      beginPath: () => calls.push("beginPath"),
      closePath: () => calls.push("closePath"),
      moveTo: () => calls.push("moveTo"),
      lineTo: () => calls.push("lineTo"),
      bezierCurveTo: () => calls.push("bezierCurveTo"),
      arcTo: () => calls.push("arcTo"),
      stroke: () => calls.push("stroke"),
      fill: (rule?: string) => calls.push(`fill:${rule ?? ""}`),
      ellipse: () => {},
      strokeRect: () => calls.push("strokeRect"),
      setLineDash: () => {},
    };
    return ctx;
  }

  const triShape = (overrides: Partial<Shape> = {}): Shape => ({
    type: "shape",
    geometry: "rect",
    widthCm: 4,
    heightCm: 4,
    wrap: "inline",
    path: {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 1 },
            { op: "line", x: 1, y: 1 },
            { op: "cubic", x1: 0.75, y1: 0.5, x2: 0.25, y2: 0.5, x: 0.5, y: 0 },
            { op: "close" },
          ],
          closed: true,
        },
      ],
    },
    ...overrides,
  });

  it("paints the real outline (bezier), never the dashed placeholder", () => {
    const ctx = pathCtx();
    paintShape(ctx as unknown as CanvasRenderingContext2D, triShape({ outline: { color: "#000", widthPt: 1, dash: "solid" } }), 100, 100);
    expect(ctx.calls).toContain("bezierCurveTo");
    expect(ctx.calls).toContain("stroke");
    expect(ctx.calls).not.toContain("strokeRect");
  });

  it("fills a closed path with the non-zero winding rule", () => {
    const ctx = pathCtx();
    paintShape(ctx as unknown as CanvasRenderingContext2D, triShape({ fill: { kind: "solid", color: "#abcdef" } }), 100, 100);
    expect(ctx.calls).toContain("fill:nonzero");
  });

  it("an unfilled open path only strokes", () => {
    const ctx = pathCtx();
    const open = triShape({ fill: undefined, outline: { color: "#000", widthPt: 1, dash: "solid" } });
    open.path!.subpaths[0].closed = false;
    open.path!.subpaths[0].steps = open.path!.subpaths[0].steps.filter((s) => s.op !== "close");
    paintShape(ctx as unknown as CanvasRenderingContext2D, open, 100, 100);
    expect(ctx.calls).toContain("stroke");
    expect(ctx.calls.some((c) => c.startsWith("fill"))).toBe(false);
  });
});

// --- SPEC-073: fills, alpha, shadow, arrowheads ------------------------------

/** A recording fake covering the SPEC-073 painting surface: property writes
 * (fillStyle/strokeStyle/shadow*) and gradient/pattern factories, plus the
 * call log the older fake records. */
function richCtx(scale = 2) {
  const calls: string[] = [];
  const sets: Record<string, unknown[]> = {};
  const record = (key: string, value: unknown) => {
    (sets[key] ??= []).push(value);
  };
  const gradient = {
    stops: [] as Array<[number, string]>,
    addColorStop(position: number, color: string) {
      this.stops.push([position, color]);
    },
  };
  const target: Record<string, unknown> = {
    calls,
    sets,
    gradient,
    save: () => calls.push("save"),
    restore: () => calls.push("restore"),
    translate: () => calls.push("translate"),
    rotate: () => {},
    beginPath: () => calls.push("beginPath"),
    closePath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    arcTo: () => {},
    ellipse: () => {},
    stroke: () => calls.push("stroke"),
    fill: (rule?: string) => calls.push(rule ? `fill:${rule}` : "fill"),
    clip: (rule?: string) => calls.push(rule ? `clip:${rule}` : "clip"),
    fillRect: () => calls.push("fillRect"),
    strokeRect: () => calls.push("strokeRect"),
    drawImage: () => calls.push("drawImage"),
    setLineDash: (segments: number[]) => record("lineDash", segments),
    getTransform: () => ({ a: scale, b: 0 }),
    createLinearGradient: (...args: number[]) => {
      calls.push("createLinearGradient");
      record("linearGradientArgs", args);
      return gradient;
    },
    createRadialGradient: (...args: number[]) => {
      calls.push("createRadialGradient");
      record("radialGradientArgs", args);
      return gradient;
    },
    createPattern: () => {
      calls.push("createPattern");
      return {};
    },
  };
  const proxy = new Proxy(target, {
    set(obj, key, value) {
      record(String(key), value);
      obj[String(key)] = value;
      return true;
    },
  });
  return proxy as unknown as CanvasRenderingContext2D & {
    calls: string[];
    sets: Record<string, unknown[]>;
    gradient: { stops: Array<[number, string]> };
  };
}

function areaShape(overrides: Partial<Shape> = {}): Shape {
  return { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "inline", ...overrides };
}

describe("SPEC-073 fills (AC-1, AC-5, AC-7)", () => {
  it("a solid fill with alpha paints through an rgba() fillStyle", () => {
    const ctx = richCtx();
    paintShape(ctx, areaShape({ fill: { kind: "solid", color: "#336699", alpha: 0.4 } }), 100, 80);
    expect(ctx.sets.fillStyle?.some((v) => v === "rgba(51, 102, 153, 0.4)")).toBe(true);
    expect(ctx.calls).toContain("fill");
  });

  it("a linear gradient builds a CanvasGradient with every stop (alpha as rgba)", () => {
    const ctx = richCtx();
    paintShape(
      ctx,
      areaShape({
        fill: {
          kind: "gradient",
          gradient: {
            type: "linear",
            angleDeg: 90,
            stops: [
              { position: 0, color: "#4472c4" },
              { position: 1, color: "#8faadc", alpha: 0.5 },
            ],
          },
        },
      }),
      100,
      80,
    );
    expect(ctx.calls).toContain("createLinearGradient");
    expect(ctx.gradient.stops).toEqual([
      [0, "#4472c4"],
      [1, "rgba(143, 170, 220, 0.5)"],
    ]);
  });

  it("a radial gradient builds a radial CanvasGradient", () => {
    const ctx = richCtx();
    paintShape(
      ctx,
      areaShape({
        fill: {
          kind: "gradient",
          gradient: {
            type: "radial",
            stops: [
              { position: 0, color: "#000000" },
              { position: 1, color: "#ffffff" },
            ],
          },
        },
      }),
      100,
      80,
    );
    expect(ctx.calls).toContain("createRadialGradient");
  });

  it("a pattern fill clips to the geometry, paints the background and strokes the hatch vectors", () => {
    const ctx = richCtx();
    paintShape(ctx, areaShape({ fill: { kind: "pattern", preset: "ltHorz", foreground: "#4472c4", background: "#ffffff" } }), 100, 80);
    expect(ctx.calls).toContain("clip");
    expect(ctx.calls).toContain("fillRect");
    expect(ctx.calls).toContain("stroke");
  });

  it("a fill of kind none paints nothing", () => {
    const ctx = richCtx();
    paintShape(ctx, areaShape({ fill: { kind: "none" } }), 100, 80);
    expect(ctx.calls).not.toContain("fill");
  });

  it("an unresolved picture fill paints nothing; a resolved one clips and draws", () => {
    const unresolved = richCtx();
    paintShape(unresolved, areaShape({ fill: { kind: "picture", mediaId: "m1", mode: "stretch" } }), 100, 80);
    expect(unresolved.calls).not.toContain("drawImage");
    const resolved = richCtx();
    paintShape(
      resolved,
      areaShape({ fill: { kind: "picture", mediaId: "m1", mode: "stretch" } }),
      100,
      80,
      { pictures: () => ({}) as unknown as CanvasImageSource },
    );
    expect(resolved.calls).toContain("clip");
    expect(resolved.calls).toContain("drawImage");
  });
});

describe("SPEC-073 outline alpha/join/cap and arrowheads (AC-6, AC-14)", () => {
  it("outline alpha strokes through an rgba() strokeStyle; join/cap apply", () => {
    const ctx = richCtx();
    paintShape(
      ctx,
      areaShape({ outline: { color: "#ff0000", widthPt: 2, dash: "solid", alpha: 0.5, join: "round", cap: "round" } }),
      100,
      80,
    );
    expect(ctx.sets.strokeStyle?.some((v) => v === "rgba(255, 0, 0, 0.5)")).toBe(true);
    expect(ctx.sets.lineJoin?.at(-1)).toBe("round");
    expect(ctx.sets.lineCap?.at(-1)).toBe("round");
  });

  it("explicit start/end arrowheads on a line paint filled heads", () => {
    const ctx = richCtx();
    paintShape(
      ctx,
      baseLine({
        outline: {
          color: "#000000",
          widthPt: 1,
          dash: "solid",
          startArrow: { kind: "triangle", size: "medium" },
          endArrow: { kind: "diamond", size: "large" },
        },
      }),
      100,
      0,
    );
    // Line stroke + two filled heads.
    expect(ctx.calls.filter((c) => c === "fill")).toHaveLength(2);
    expect(ctx.calls).toContain("stroke");
  });
});

describe("SPEC-073 shadow pass (AC-9, AC-39, AC-40)", () => {
  it("is skipped entirely for shapes with no shadow (AC-40)", () => {
    const ctx = richCtx();
    paintShape(ctx, areaShape({ fill: { kind: "solid", color: "#4472c4" } }), 100, 80);
    expect(ctx.sets.shadowColor).toBeUndefined();
    expect(ctx.sets.shadowBlur).toBeUndefined();
  });

  it("multiplies blur and offset by the context's scale — computed in the devicePixelRatio*zoom space, not scaled as a bitmap (AC-39)", () => {
    const PX_PER_PT = 96 / 72;
    const shadow = { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 };
    const at1 = richCtx(1);
    paintShape(at1, areaShape({ fill: { kind: "solid", color: "#4472c4" }, shadow }), 100, 80);
    const at3 = richCtx(3);
    paintShape(at3, areaShape({ fill: { kind: "solid", color: "#4472c4" }, shadow }), 100, 80);
    expect(at1.sets.shadowBlur?.[0]).toBeCloseTo(4 * PX_PER_PT);
    expect(at3.sets.shadowBlur?.[0]).toBeCloseTo(4 * PX_PER_PT * 3);
    expect(at1.sets.shadowColor?.[0]).toBe("rgba(0, 0, 0, 0.4)");
    // 90° = downward: the y offset carries the distance.
    expect(at3.sets.shadowOffsetY?.[0]).toBeCloseTo(3 * PX_PER_PT * 3);
  });
});
