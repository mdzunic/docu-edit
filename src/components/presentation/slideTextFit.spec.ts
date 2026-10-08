// SPEC-097 — the DOM half of autofit: the text body's inline style (insets,
// anchor, columns, wrap and the two scale custom properties), measuring an
// off-flow clone at a scale, and the ladder walk over that measurer. jsdom
// lays nothing out, so heights come from a stubbed `offsetHeight` that reads
// the clone's own `--fit-scale` — the contract the stage relies on.
import { afterEach, describe, expect, it, vi } from "vitest";
import type { ShapeFrame, TextFrame } from "../../lib/presentation/model";
import { cmToPx } from "../../lib/core/units";
import { cmPx } from "./slideRender";
import {
  applyFitStyle,
  bodyStyle,
  FRAME_ID_ATTRIBUTE,
  frameTextHost,
  FULL_SCALE,
  measureBodyHeightCm,
  measureFrameFit,
  TEXT_BODY_CLASS,
  textWidthCm,
} from "./slideTextFit";

function textFrame(extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t1", kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 4, paragraphs: [{ runs: [{ text: "x" }] }], ...extra };
}

function shapeFrame(extra: Partial<ShapeFrame> = {}): ShapeFrame {
  return { id: "s1", kind: "shape", geometry: "rect", xCm: 0, yCm: 0, widthCm: 6, heightCm: 3, ...extra };
}

/** A rendered frame root + text body in the document, as the stage emits it. */
function renderedBody(frameId = "t1"): HTMLElement {
  const root = document.createElement("div");
  root.setAttribute(FRAME_ID_ATTRIBUTE, frameId);
  const body = document.createElement("div");
  body.className = TEXT_BODY_CLASS;
  body.textContent = "Some text";
  root.appendChild(body);
  document.body.appendChild(root);
  return body;
}

/** Stubs layout: a `.text-body` measures `heightAtFullCm × --fit-scale`. */
function stubLayout(heightAtFullCm: number): void {
  vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockImplementation(function (this: HTMLElement) {
    if (!this.classList.contains(TEXT_BODY_CLASS)) return 0;
    const scale = Number(this.style.getPropertyValue("--fit-scale") || "1");
    return cmToPx(heightAtFullCm * scale);
  });
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.textContent = "";
});

describe("bodyStyle (AC-11..15)", () => {
  it("renders the stored defaults: pptx insets, top anchor for text, unscaled", () => {
    const style = bodyStyle(textFrame());
    expect(style.padding).toBe(`${cmPx(0.127)} ${cmPx(0.254)} ${cmPx(0.127)} ${cmPx(0.254)}`);
    expect(style["justify-content"]).toBe("flex-start");
    expect(style["--fit-scale"]).toBe("1");
    expect(style["--line-scale"]).toBe("1");
    expect(style["column-count"]).toBeUndefined();
    expect(style["white-space"]).toBeUndefined();
  });

  it("anchors shape text middle by default and honours a stored anchor", () => {
    expect(bodyStyle(shapeFrame())["justify-content"]).toBe("center");
    expect(bodyStyle(textFrame({ body: { anchor: "bottom" } }))["justify-content"]).toBe("flex-end");
    expect(bodyStyle(shapeFrame({ body: { anchor: "top" } }))["justify-content"]).toBe("flex-start");
  });

  it("carries custom insets, columns, no-wrap and the scale", () => {
    const frame = textFrame({
      body: { insetsCm: { left: 1, top: 0.5, right: 0, bottom: 2 }, columns: { count: 2, gapCm: 0.5 }, wrap: false },
    });
    const style = bodyStyle(frame, { fontScalePct: 62.5, lineScalePct: 90 });
    expect(style.padding).toBe(`${cmPx(0.5)} ${cmPx(0)} ${cmPx(2)} ${cmPx(1)}`);
    expect(style["column-count"]).toBe("2");
    expect(style["column-gap"]).toBe(cmPx(0.5));
    expect(style.display).toBe("block");
    expect(style["white-space"]).toBe("pre");
    expect(style["--fit-scale"]).toBe("0.625");
    expect(style["--line-scale"]).toBe("0.9");
  });

  it("applyFitStyle assigns the same declarations to an element", () => {
    const el = document.createElement("div");
    applyFitStyle(el, textFrame(), { fontScalePct: 85, lineScalePct: 100 });
    expect(el.style.getPropertyValue("--fit-scale")).toBe("0.85");
    expect(el.style.getPropertyValue("justify-content")).toBe("flex-start");
  });
});

describe("measuring (AC-37/38/40)", () => {
  it("finds a rendered body by its frame id, else null", () => {
    const body = renderedBody("frame-9");
    expect(frameTextHost("frame-9")).toBe(body);
    expect(frameTextHost("missing")).toBeNull();
  });

  it("measures an off-flow clone at the requested scale and removes it", () => {
    stubLayout(5);
    const body = renderedBody();
    const parent = body.parentElement as HTMLElement;
    expect(measureBodyHeightCm(body, 8, 100, 100)).toBeCloseTo(5, 6);
    expect(measureBodyHeightCm(body, 8, 50, 80)).toBeCloseTo(2.5, 6);
    // Only the original remains; the original's own style is untouched.
    expect(parent.children).toHaveLength(1);
    expect(body.style.getPropertyValue("--fit-scale")).toBe("");
  });

  it("the clone releases an inset body's far edges so its height is the content's", () => {
    const body = renderedBody();
    body.style.setProperty("inset", "0");
    let seen: { right: string; bottom: string; width: string } | null = null;
    vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockImplementation(function (this: HTMLElement) {
      if (this !== body && this.classList.contains(TEXT_BODY_CLASS)) {
        seen = { right: this.style.right, bottom: this.style.bottom, width: this.style.width };
      }
      return 10;
    });
    measureBodyHeightCm(body, 3, 100, 100);
    expect(seen).toEqual({ right: "auto", bottom: "auto", width: cmPx(3) });
  });

  it("returns null when nothing lays out (jsdom) — callers keep what is stored", () => {
    const body = renderedBody();
    expect(measureFrameFit(textFrame({ body: { autofit: "shrink" } }), body)).toBeNull();
    expect(measureFrameFit(textFrame(), null)).toBeNull();
  });

  it("walks the shrink ladder to the first step that fits the inset-reduced height", () => {
    // 4cm frame − 0.254cm insets = 3.746cm available; content is 5cm at 100%.
    stubLayout(5);
    const fit = measureFrameFit(textFrame({ body: { autofit: "shrink" } }), renderedBody());
    // 5 × 0.70 = 3.5 fits; 5 × 0.775 = 3.875 does not.
    expect(fit).toEqual({ scale: { fontScalePct: 70, lineScalePct: 90 }, contentHeightCm: 5 });
  });

  it("reports full scale and the natural height for a non-shrink body (resizeShape's input)", () => {
    stubLayout(2.5);
    const fit = measureFrameFit(shapeFrame({ body: { autofit: "resizeShape" } }), renderedBody("s1"));
    expect(fit).toEqual({ scale: FULL_SCALE, contentHeightCm: 2.5 });
  });

  it("text width is the frame width less the side insets", () => {
    expect(textWidthCm(textFrame())).toBeCloseTo(10 - 0.508, 6);
    expect(textWidthCm(textFrame({ body: { insetsCm: { left: 1, top: 0, right: 2, bottom: 0 } } }))).toBe(7);
  });
});
