// SPEC-069 AC-5: the pure adjustment-handle geometry — one handle per
// declared adjustment, placed by axis, mapped back from a drag point, and
// nudged by keyboard, all clamped to the declared range.
import { describe, expect, it } from "vitest";
import {
  adjustmentHandles,
  adjustmentValueFromPoint,
  nudgedAdjustments,
} from "./shapeAdjustHandles";
import { shapeDefinition } from "./core/shapeCatalogue";

describe("adjustmentHandles", () => {
  it("returns nothing for a geometry with no adjustments", () => {
    expect(adjustmentHandles("rect")).toEqual([]);
    expect(adjustmentHandles("triangle")).toEqual([]);
  });

  it("gives one handle per declared adjustment, at the default value (AC-5a)", () => {
    const handles = adjustmentHandles("rightArrow");
    const def = shapeDefinition("rightArrow")!;
    expect(handles).toHaveLength(def.adjustments!.length);
    handles.forEach((handle, i) => {
      expect(handle.index).toBe(i);
      expect(handle.label).toBe(def.adjustments![i].label);
      expect(handle.value).toBeCloseTo(def.adjustments![i].default, 6);
    });
  });

  it("places an x-axis handle on the top edge at its value", () => {
    const [handle] = adjustmentHandles("roundRect", [0.3]);
    expect(handle.axis).toBe("x");
    expect(handle.xFrac).toBeCloseTo(0.3, 6);
    expect(handle.yFrac).toBe(0);
  });

  it("places a radial handle straight up from the centre by its value", () => {
    const [handle] = adjustmentHandles("star6", [0.6]);
    expect(handle.axis).toBe("radial");
    expect(handle.xFrac).toBeCloseTo(0.5, 6);
    expect(handle.yFrac).toBeCloseTo(0.5 - 0.6 * 0.5, 6);
  });

  it("uses the clamped value for an out-of-range stored adjustment", () => {
    const def = shapeDefinition("roundRect")!;
    const [handle] = adjustmentHandles("roundRect", [5]);
    expect(handle.value).toBe(def.adjustments![0].max);
  });
});

describe("adjustmentValueFromPoint", () => {
  it("maps an x-axis drag to the x fraction, clamped to range (AC-5c)", () => {
    const def = shapeDefinition("roundRect")!;
    expect(adjustmentValueFromPoint("roundRect", 0, 0.25, 0)).toBeCloseTo(0.25, 6);
    expect(adjustmentValueFromPoint("roundRect", 0, 2, 0)).toBe(def.adjustments![0].max);
    expect(adjustmentValueFromPoint("roundRect", 0, -1, 0)).toBe(def.adjustments![0].min);
  });

  it("maps a radial drag to twice the distance from the centre", () => {
    // A point 0.25 above centre -> distance 0.25 -> value 0.5.
    expect(adjustmentValueFromPoint("star6", 0, 0.5, 0.25)).toBeCloseTo(0.5, 6);
  });

  it("returns null for a geometry/index with no adjustment", () => {
    expect(adjustmentValueFromPoint("rect", 0, 0.5, 0.5)).toBeNull();
  });
});

describe("nudgedAdjustments", () => {
  it("nudges the focused entry by 1% of the range, 5% with shift", () => {
    const def = shapeDefinition("roundRect")!;
    const range = def.adjustments![0].max - def.adjustments![0].min;
    const base = def.adjustments![0].default;
    expect(nudgedAdjustments("roundRect", undefined, 0, 1, false)[0]).toBeCloseTo(base + 0.01 * range, 6);
    expect(nudgedAdjustments("roundRect", undefined, 0, 1, true)[0]).toBeCloseTo(base + 0.05 * range, 6);
    expect(nudgedAdjustments("roundRect", undefined, 0, -1, false)[0]).toBeCloseTo(base - 0.01 * range, 6);
  });

  it("clamps at the range bounds", () => {
    const def = shapeDefinition("roundRect")!;
    expect(nudgedAdjustments("roundRect", [def.adjustments![0].max], 0, 1, true)[0]).toBe(def.adjustments![0].max);
    expect(nudgedAdjustments("roundRect", [def.adjustments![0].min], 0, -1, true)[0]).toBe(def.adjustments![0].min);
  });
});
