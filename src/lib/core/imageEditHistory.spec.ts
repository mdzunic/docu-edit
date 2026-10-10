// SPEC-112: the history list's helpers — immutable, one entry per
// single-valued kind, objects accumulating, and a resize that keeps its
// scale when the size under it changes.
import { describe, expect, it } from "vitest";
import {
  canAddObject,
  keepResizeScale,
  removeAt,
  replaceAt,
  rotateBy,
  setAdjustChannel,
  toggleFlip,
  upsertEntry,
} from "./imageEditHistory";
import { MAX_EDIT_STACK, type ImageEdit } from "./imageEdits";

const PHOTO = { widthPx: 160, heightPx: 120 };

describe("imageEditHistory", () => {
  it("updates single-valued entries in place, never mutating the input", () => {
    const history = Object.freeze([
      { kind: "crop", x: 0, y: 0, width: 10, height: 10 },
      { kind: "flip", axis: "horizontal" },
    ]) as readonly ImageEdit[];
    const next = upsertEntry(history, { kind: "crop", x: 1, y: 1, width: 5, height: 5 });
    expect(next).toEqual([
      { kind: "crop", x: 1, y: 1, width: 5, height: 5 },
      { kind: "flip", axis: "horizontal" },
    ]);
    expect(upsertEntry(history, { kind: "straighten", degrees: 2 })).toHaveLength(3);
  });

  it("turns by quarter turns, removing a full turn", () => {
    let history: ImageEdit[] = [];
    history = rotateBy(history, 1);
    expect(history).toEqual([{ kind: "rotate", quarterTurns: 1 }]);
    history = rotateBy(history, 1);
    expect(history).toEqual([{ kind: "rotate", quarterTurns: 2 }]);
    history = rotateBy(history, -1);
    history = rotateBy(history, -1);
    expect(history).toEqual([]);
    expect(rotateBy([], -1)).toEqual([{ kind: "rotate", quarterTurns: 3 }]);
  });

  it("toggles flips per axis", () => {
    const once = toggleFlip([], "vertical");
    expect(once).toEqual([{ kind: "flip", axis: "vertical" }]);
    expect(toggleFlip(once, "horizontal")).toHaveLength(2);
    expect(toggleFlip(once, "vertical")).toEqual([]);
  });

  it("merges adjust channels and drops the entry at neutral", () => {
    let history = setAdjustChannel([], "brightness", 30);
    history = setAdjustChannel(history, "invert", true);
    expect(history).toEqual([{ kind: "adjust", brightness: 30, invert: true }]);
    history = setAdjustChannel(history, "brightness", 0);
    expect(history).toEqual([{ kind: "adjust", invert: true }]);
    expect(setAdjustChannel(history, "invert", false)).toEqual([]);
  });

  it("replaces and removes by index", () => {
    const a: ImageEdit = { kind: "redact", x: 0, y: 0, width: 1, height: 1, color: "#000000" };
    const b: ImageEdit = { kind: "redact", x: 5, y: 5, width: 1, height: 1, color: "#000000" };
    expect(replaceAt([a, a], 1, b)).toEqual([a, b]);
    expect(removeAt([a, b], 0)).toEqual([b]);
  });

  it("refuses objects past the stack bound", () => {
    const full: ImageEdit[] = Array.from({ length: MAX_EDIT_STACK }, () => ({ kind: "redact", x: 0, y: 0, width: 1, height: 1, color: "#000000" }));
    expect(canAddObject(full.slice(1))).toBe(true);
    expect(canAddObject(full)).toBe(false);
  });

  it("keeps a resize's scale when a crop or a turn changes the size under it", () => {
    const resized: ImageEdit[] = [{ kind: "resize", widthPx: 80, heightPx: 60, resample: "bilinear" }];
    // A turn swaps it: 80 × 60 becomes 60 × 80.
    const turned = keepResizeScale(resized, rotateBy(resized, 1), PHOTO);
    expect(turned).toContainEqual({ kind: "resize", widthPx: 60, heightPx: 80, resample: "bilinear" });
    // A 16:9 crop (160 × 90) keeps the half scale: 80 × 45.
    const cropped = keepResizeScale(resized, upsertEntry(resized, { kind: "crop", x: 0, y: 15, width: 160, height: 90 }), PHOTO);
    expect(cropped).toContainEqual({ kind: "resize", widthPx: 80, heightPx: 45, resample: "bilinear" });
    // An explicit new resize is left as the user set it.
    const explicit = upsertEntry(resized, { kind: "resize", widthPx: 10, heightPx: 10, resample: "bilinear" });
    expect(keepResizeScale(resized, explicit, PHOTO)).toEqual(explicit);
  });
});
