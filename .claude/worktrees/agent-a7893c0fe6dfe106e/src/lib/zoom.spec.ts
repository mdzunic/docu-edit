import { describe, expect, it } from "vitest";
import {
  clampZoom,
  clientToContentPx,
  contentPxToClient,
  DEFAULT_ZOOM,
  FIT_WIDTH_PADDING,
  fitPageZoom,
  fitWidthZoom,
  gridColumns,
  INPUT_MAX,
  INPUT_MIN,
  parseZoomInput,
  PRESET_MAX,
  PRESET_MIN,
  stepZoomDown,
  stepZoomUp,
  ZOOM_PRESETS,
} from "./zoom";

describe("clampZoom", () => {
  it("clamps to the input band by default (AC-9)", () => {
    expect(clampZoom(10)).toBe(INPUT_MIN);
    expect(clampZoom(500)).toBe(INPUT_MAX);
    expect(clampZoom(100)).toBe(100);
  });

  it("accepts custom bounds", () => {
    expect(clampZoom(300, PRESET_MIN, PRESET_MAX)).toBe(PRESET_MAX);
    expect(clampZoom(10, PRESET_MIN, PRESET_MAX)).toBe(PRESET_MIN);
  });
});

describe("stepZoomUp / stepZoomDown (AC-6, AC-7)", () => {
  it("steps the preset ladder from an on-ladder value", () => {
    expect(stepZoomUp(100)).toBe(125);
    expect(stepZoomUp(125)).toBe(150);
    expect(stepZoomDown(100)).toBe(90);
    expect(stepZoomDown(90)).toBe(75);
  });

  it("clamps at the ladder ends (AC-6)", () => {
    expect(stepZoomUp(200)).toBe(PRESET_MAX);
    expect(stepZoomUp(400)).toBe(PRESET_MAX);
    expect(stepZoomDown(50)).toBe(PRESET_MIN);
    expect(stepZoomDown(25)).toBe(PRESET_MIN);
  });

  it("snaps to the surrounding presets when off-ladder (AC-7)", () => {
    expect(stepZoomUp(110)).toBe(125);
    expect(stepZoomDown(110)).toBe(100);
    expect(stepZoomUp(62)).toBe(75);
    expect(stepZoomDown(62)).toBe(50);
  });

  it("walks the whole ladder up then down", () => {
    let z = PRESET_MIN;
    const up: number[] = [z];
    for (let i = 0; i < ZOOM_PRESETS.length + 2; i += 1) {
      z = stepZoomUp(z);
      up.push(z);
    }
    expect(up[up.length - 1]).toBe(PRESET_MAX);
    expect(up).toContain(90);
    expect(up).toContain(125);
  });
});

describe("parseZoomInput (AC-8, AC-9, AC-10)", () => {
  it("trims whitespace and a trailing percent", () => {
    expect(parseZoomInput("  250  ")).toBe(250);
    expect(parseZoomInput("88%")).toBe(88);
    expect(parseZoomInput(" 120 % ")).toBe(120);
  });

  it("clamps out-of-range values to the nearest bound (AC-9)", () => {
    expect(parseZoomInput("10")).toBe(INPUT_MIN);
    expect(parseZoomInput("500")).toBe(INPUT_MAX);
    expect(parseZoomInput("25")).toBe(25);
    expect(parseZoomInput("400")).toBe(400);
  });

  it("returns null for empty or non-numeric input (AC-10)", () => {
    expect(parseZoomInput("")).toBeNull();
    expect(parseZoomInput("   ")).toBeNull();
    expect(parseZoomInput("abc")).toBeNull();
    expect(parseZoomInput("12.5")).toBeNull();
    expect(parseZoomInput("%")).toBeNull();
  });
});

describe("fitWidthZoom (AC-26)", () => {
  it("solves pageWidth * zoom == availableWidth - 48", () => {
    // pageWidth 800px, available 1400px → (1400-48)/800 = 1.69 → 169%
    const z = fitWidthZoom(1400, 800);
    expect((800 * z) / 100).toBeCloseTo(1400 - FIT_WIDTH_PADDING, 6);
  });

  it("clamps to the input band and guards a zero page width", () => {
    expect(fitWidthZoom(100, 0)).toBe(DEFAULT_ZOOM);
    expect(fitWidthZoom(100, 2000)).toBe(INPUT_MIN);
  });
});

describe("fitPageZoom (AC-28)", () => {
  it("solves pageHeight * zoom == visibleHeight", () => {
    const z = fitPageZoom(600, 1000);
    expect(z).toBeCloseTo(60, 5);
  });

  it("guards a zero page height", () => {
    expect(fitPageZoom(600, 0)).toBe(DEFAULT_ZOOM);
  });
});

describe("gridColumns (AC-36)", () => {
  it("floors availableWidth / (pageWidth*zoom + gap), min one", () => {
    // page 800px at 50% → 400px; +24 gap → 424; 1400/424 = 3.30 → 3 cols
    expect(gridColumns(1400, 800, 0.5)).toBe(3);
    // at 100% one page (800+24=824) fits once in 1400
    expect(gridColumns(1400, 800, 1)).toBe(1);
    // never below one even when a single page overflows
    expect(gridColumns(500, 800, 1)).toBe(1);
  });
});

describe("clientToContentPx / contentPxToClient round-trip (AC-23)", () => {
  const rect = { left: 120, top: 60 };
  for (const zoom of [0.5, 1, 2]) {
    it(`is the identity at zoom ${zoom}`, () => {
      for (const [cx, cy] of [
        [0, 0],
        [400, 250],
        [1000, 3000],
      ]) {
        const content = clientToContentPx(cx, cy, rect, zoom);
        const back = contentPxToClient(content.x, content.y, rect, zoom);
        expect(back.x).toBeCloseTo(cx, 6);
        expect(back.y).toBeCloseTo(cy, 6);
      }
    });
  }

  it("divides measured offsets by zoom (AC-22)", () => {
    // a point 100px right of .pages' scaled edge is 200 model-px at 50%
    expect(clientToContentPx(220, 60, rect, 0.5).x).toBe(200);
    // a 200 model-px position lands 100px right of the edge at 50%
    expect(contentPxToClient(200, 0, rect, 0.5).x).toBe(220);
  });
});
