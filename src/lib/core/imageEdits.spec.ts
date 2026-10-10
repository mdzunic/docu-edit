// SPEC-112: the edit stack, pure — normalisation order and collapsing for
// every kind pair, geometry through crop + straighten + rotate + flip +
// resize in combination (exact output sizes and source rectangles, worked
// by hand in the comments), the empty stack as identity, and the
// over-long stack bound. No canvas anywhere: resolveEditGeometry is given
// only the natural size.
import { describe, expect, it } from "vitest";
import {
  applyMatrix,
  CROP_ASPECTS,
  cropAnnouncement,
  cropAspectRatio,
  cropForAspect,
  describeEdit,
  describeEditDetail,
  EDIT_KINDS,
  inscribedRect,
  MAX_EDIT_STACK,
  moveCropHandle,
  normaliseStack,
  resizeObjectEdit,
  resolveEditGeometry,
  translateEdit,
  type EditKind,
  type ImageEdit,
} from "./imageEdits";

const PHOTO = { widthPx: 160, heightPx: 120 };

/** Two distinct, non-identity samples of every kind. */
const SAMPLES: Record<EditKind, [ImageEdit, ImageEdit]> = {
  straighten: [
    { kind: "straighten", degrees: 5 },
    { kind: "straighten", degrees: -3 },
  ],
  crop: [
    { kind: "crop", x: 0, y: 15, width: 160, height: 90 },
    { kind: "crop", x: 10, y: 10, width: 50, height: 40 },
  ],
  rotate: [
    { kind: "rotate", quarterTurns: 1 },
    { kind: "rotate", quarterTurns: 2 },
  ],
  flip: [
    { kind: "flip", axis: "horizontal" },
    { kind: "flip", axis: "vertical" },
  ],
  resize: [
    { kind: "resize", widthPx: 80, heightPx: 45, resample: "bilinear" },
    { kind: "resize", widthPx: 40, heightPx: 30, resample: "nearest" },
  ],
  adjust: [
    { kind: "adjust", brightness: 30, contrast: 10 },
    { kind: "adjust", contrast: -20, invert: true },
  ],
  redact: [
    { kind: "redact", x: 1, y: 2, width: 3, height: 4, color: "#000000" },
    { kind: "redact", x: 5, y: 6, width: 7, height: 8, color: "#ff0000" },
  ],
  annotate: [
    { kind: "annotate", shape: { shape: "arrow", x1: 0, y1: 0, x2: 10, y2: 10 }, color: "#ff0000", strokeWidth: 4, opacity: 1 },
    { kind: "annotate", shape: { shape: "text", x: 3, y: 4, text: "Hi", fontSize: 24 }, color: "#000000", strokeWidth: 2, opacity: 0.5 },
  ],
};

/** What a same-kind pair [first, second] collapses to. */
const COLLAPSED: Record<EditKind, ImageEdit[]> = {
  // The last one wins.
  straighten: [{ kind: "straighten", degrees: -3 }],
  crop: [{ kind: "crop", x: 10, y: 10, width: 50, height: 40 }],
  // Quarter turns compose: 1 + 2 = 3.
  rotate: [{ kind: "rotate", quarterTurns: 3 }],
  // One per axis, horizontal first.
  flip: [
    { kind: "flip", axis: "horizontal" },
    { kind: "flip", axis: "vertical" },
  ],
  resize: [{ kind: "resize", widthPx: 40, heightPx: 30, resample: "nearest" }],
  // Merged channel by channel; the later contrast wins.
  adjust: [{ kind: "adjust", brightness: 30, contrast: -20, invert: true }],
  // Accumulate in insertion order.
  redact: [SAMPLES.redact[0], SAMPLES.redact[1]],
  annotate: [SAMPLES.annotate[0], SAMPLES.annotate[1]],
};

describe("normaliseStack", () => {
  it("lists the kinds in application order", () => {
    expect(EDIT_KINDS).toEqual(["straighten", "crop", "rotate", "flip", "resize", "adjust", "redact", "annotate"]);
  });

  describe("orders every pair of different kinds, whatever order they arrive in", () => {
    const pairs: Array<[EditKind, EditKind]> = [];
    for (const [i, first] of EDIT_KINDS.entries()) {
      for (const second of EDIT_KINDS.slice(i + 1)) {
        pairs.push([first, second]);
      }
    }
    it.each(pairs)("%s before %s", (first, second) => {
      const a = SAMPLES[first][0];
      const b = SAMPLES[second][0];
      expect(normaliseStack([b, a])).toEqual([a, b]);
      expect(normaliseStack([a, b])).toEqual([a, b]);
    });
  });

  describe("collapses (or accumulates) every same-kind pair", () => {
    it.each(EDIT_KINDS.map((kind) => [kind]))("%s", (kind) => {
      expect(normaliseStack(SAMPLES[kind])).toEqual(COLLAPSED[kind]);
    });
  });

  it("drops identities: 0°, a full turn, a double flip and a neutral adjust", () => {
    expect(
      normaliseStack([
        { kind: "straighten", degrees: 0 },
        { kind: "rotate", quarterTurns: 3 },
        { kind: "rotate", quarterTurns: 1 },
        { kind: "flip", axis: "vertical" },
        { kind: "flip", axis: "vertical" },
        { kind: "adjust", brightness: 40 },
        { kind: "adjust", brightness: 0 },
      ]),
    ).toEqual([]);
  });

  it("keeps the canonical order for a full stack worked in reverse", () => {
    const worked = EDIT_KINDS.toReversed().map((kind) => SAMPLES[kind][0]);
    expect(normaliseStack(worked).map((edit) => edit.kind)).toEqual(EDIT_KINDS);
  });

  it("is pure and total: frozen input, malformed entries, the same answer twice", () => {
    const input = Object.freeze([
      Object.freeze({ kind: "adjust", brightness: 30 }),
      Object.freeze({ kind: "bogus" }),
      null,
      Object.freeze({ kind: "straighten", degrees: 90 }),
      Object.freeze({ kind: "crop", x: -5, y: 2.4, width: 0, height: Number.NaN }),
      Object.freeze({ kind: "redact", x: 1, y: 1, width: 2, height: 2, color: "red" }),
      Object.freeze({ kind: "annotate", shape: { shape: "text", x: 0, y: 0, text: "", fontSize: 10 }, color: "#000", strokeWidth: 1, opacity: 1 }),
    ]) as unknown as ImageEdit[];
    const once = normaliseStack(input);
    expect(normaliseStack(input)).toEqual(once);
    expect(once).toEqual([
      // Clamped to the ±45° range.
      { kind: "straighten", degrees: 45 },
      // Negative origin clamped, sizes at least one pixel.
      { kind: "crop", x: 0, y: 2, width: 1, height: 1 },
      { kind: "adjust", brightness: 30 },
      // A colour that is not #rrggbb falls back to black.
      { kind: "redact", x: 1, y: 1, width: 2, height: 2, color: "#000000" },
      // Empty text annotates nothing.
    ]);
  });

  it("bounds an over-long stack at MAX_EDIT_STACK, keeping the single-valued edits", () => {
    const redactions: ImageEdit[] = Array.from({ length: MAX_EDIT_STACK + 100 }, (_, i) => ({
      kind: "redact",
      x: i,
      y: 0,
      width: 1,
      height: 1,
      color: "#000000",
    }));
    const stack = normaliseStack([...redactions, SAMPLES.crop[0]]);
    expect(stack).toHaveLength(MAX_EDIT_STACK);
    expect(stack[0]).toEqual(SAMPLES.crop[0]);
    // The oldest redactions survive; the newest past the bound are dropped.
    expect(stack[1]).toMatchObject({ kind: "redact", x: 0 });
    expect(stack[MAX_EDIT_STACK - 1]).toMatchObject({ kind: "redact", x: MAX_EDIT_STACK - 2 });
  });
});

describe("resolveEditGeometry", () => {
  it("treats the empty stack as the identity", () => {
    const geometry = resolveEditGeometry([], PHOTO);
    expect(geometry).toMatchObject({
      widthPx: 160,
      heightPx: 120,
      source: { x: 0, y: 0, width: 160, height: 120 },
      quarterTurns: 0,
      flipH: false,
      flipV: false,
      straightenDegrees: 0,
      inscribed: { x: 0, y: 0, width: 160, height: 120 },
    });
    expect(geometry.transform).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
  });

  it("crops to 16:9 and turns right: 160 × 90 from y = 15, then 90 × 160", () => {
    const crop = cropForAspect(cropAspectRatio("16:9", PHOTO), resolveEditGeometry([], PHOTO).inscribed);
    // 160 / (16/9) = 90 rows, centred: (120 − 90) / 2 = 15.
    expect(crop).toEqual({ x: 0, y: 15, width: 160, height: 90 });
    const cropped = resolveEditGeometry([{ kind: "crop", ...crop }], PHOTO);
    expect([cropped.widthPx, cropped.heightPx]).toEqual([160, 90]);
    const turned = resolveEditGeometry([{ kind: "crop", ...crop }, { kind: "rotate", quarterTurns: 1 }], PHOTO);
    expect([turned.widthPx, turned.heightPx]).toEqual([90, 160]);
    expect(turned.source).toEqual({ x: 0, y: 15, width: 160, height: 90 });
  });

  it("maps the crop's corners exactly through rotate + flip + resize", () => {
    const geometry = resolveEditGeometry(
      [
        { kind: "resize", widthPx: 45, heightPx: 80, resample: "bilinear" },
        { kind: "flip", axis: "horizontal" },
        { kind: "rotate", quarterTurns: 1 },
        { kind: "crop", x: 0, y: 15, width: 160, height: 90 },
      ],
      PHOTO,
    );
    expect([geometry.widthPx, geometry.heightPx]).toEqual([45, 80]);
    expect(geometry.preResizePx).toEqual({ widthPx: 90, heightPx: 160 });
    // Top-left of the crop (0, 15): cropped (0, 0) → turned right in the
    // 160 × 90 frame (90 − 0, 0) = (90, 0) → flipped in 90 × 160 (0, 0) →
    // halved (0, 0).
    expect(applyMatrix(geometry.transform, 0, 15)).toEqual({ x: 0, y: 0 });
    // Bottom-right (160, 105): (160, 90) → (0, 160) → (90, 160) → (45, 80).
    expect(applyMatrix(geometry.transform, 160, 105)).toEqual({ x: 45, y: 80 });
    // Top-right (160, 15): (160, 0) → (90, 160) → (0, 160) → (0, 80).
    expect(applyMatrix(geometry.transform, 160, 15)).toEqual({ x: 0, y: 80 });
  });

  it("straightens into the inscribed area: 10° on 160 × 120 keeps 130 × 98", () => {
    // |cos 10°| = 0.98481, |sin 10°| = 0.17365; the scale is
    // min(160 / (157.57 + 20.84), 120 / (27.78 + 118.18)) = 0.82214, so the
    // exact area is 131.54 × 98.66 centred: x 14.23..145.77 → 15..145,
    // y 10.67..109.33 → 11..109.
    expect(inscribedRect(PHOTO, 10)).toEqual({ x: 15, y: 11, width: 130, height: 98 });
    const geometry = resolveEditGeometry([{ kind: "straighten", degrees: 10 }], PHOTO);
    expect(geometry.source).toEqual({ x: 15, y: 11, width: 130, height: 98 });
    expect([geometry.widthPx, geometry.heightPx]).toEqual([130, 98]);
    // Rotation is about the centre: the source centre lands on the output's.
    const centre = applyMatrix(geometry.transform, 80, 60);
    expect(centre.x).toBeCloseTo(65, 9);
    expect(centre.y).toBeCloseTo(49, 9);
    // A whole-image crop is clamped into the inscribed area.
    const clamped = resolveEditGeometry(
      [{ kind: "crop", x: 0, y: 0, width: 160, height: 120 }, { kind: "straighten", degrees: 10 }],
      PHOTO,
    );
    expect(clamped.source).toEqual({ x: 15, y: 11, width: 130, height: 98 });
  });

  it("is the whole square at 0° and its centred 70 × 70 at 45°", () => {
    expect(inscribedRect({ widthPx: 100, heightPx: 100 }, 0)).toEqual({ x: 0, y: 0, width: 100, height: 100 });
    // 100 / (70.71 + 70.71) = 0.70711 → 14.64..85.36 → 15..85.
    expect(inscribedRect({ widthPx: 100, heightPx: 100 }, 45)).toEqual({ x: 15, y: 15, width: 70, height: 70 });
  });

  it("resolves crop + straighten + rotate + flip + resize together", () => {
    // −5° on 160 × 120: scale min(160 / 169.85, 120 / 133.49) = 0.89896 →
    // 143.83 × 107.88 centred → x 8.08..151.92 → 9..151, y 6.06..113.94 →
    // 7..113: the inscribed area is (9, 7, 142, 106). The crop (20, 20,
    // 100 × 60) lies inside it, turns left to 60 × 100 and resizes to 30 × 50.
    const geometry = resolveEditGeometry(
      [
        { kind: "flip", axis: "vertical" },
        { kind: "resize", widthPx: 30, heightPx: 50, resample: "nearest" },
        { kind: "rotate", quarterTurns: 3 },
        { kind: "crop", x: 20, y: 20, width: 100, height: 60 },
        { kind: "straighten", degrees: -5 },
      ],
      PHOTO,
    );
    expect(geometry).toMatchObject({
      widthPx: 30,
      heightPx: 50,
      source: { x: 20, y: 20, width: 100, height: 60 },
      inscribed: { x: 9, y: 7, width: 142, height: 106 },
      preResizePx: { widthPx: 60, heightPx: 100 },
      quarterTurns: 3,
      flipH: false,
      flipV: true,
      straightenDegrees: -5,
    });
  });

  it("is deterministic and leaves its input alone", () => {
    const stack = Object.freeze(EDIT_KINDS.map((kind) => Object.freeze(SAMPLES[kind][0]))) as readonly ImageEdit[];
    const first = resolveEditGeometry(stack, PHOTO);
    expect(resolveEditGeometry(stack, PHOTO)).toEqual(first);
    expect(resolveEditGeometry(normaliseStack(stack), PHOTO)).toEqual(first);
  });

  it("clamps a crop wholly outside the image to one pixel inside it", () => {
    const geometry = resolveEditGeometry([{ kind: "crop", x: 500, y: 500, width: 10, height: 10 }], PHOTO);
    expect(geometry.source).toEqual({ x: 159, y: 119, width: 1, height: 1 });
    expect([geometry.widthPx, geometry.heightPx]).toEqual([1, 1]);
  });
});

describe("crop helpers", () => {
  const bounds = { x: 0, y: 0, width: 160, height: 120 };

  it("fits every aspect preset into the bounds", () => {
    const fitted = Object.fromEntries(
      CROP_ASPECTS.map((aspect) => [aspect, cropForAspect(cropAspectRatio(aspect, PHOTO), bounds)]),
    );
    expect(fitted).toEqual({
      free: { x: 0, y: 0, width: 160, height: 120 },
      original: { x: 0, y: 0, width: 160, height: 120 },
      // 120 × 120 centred: (160 − 120) / 2 = 20.
      "1:1": { x: 20, y: 0, width: 120, height: 120 },
      "4:3": { x: 0, y: 0, width: 160, height: 120 },
      // 160 / 1.5 = 106.67 → 107 rows, (120 − 107) / 2 = 6.5 → 6.
      "3:2": { x: 0, y: 6, width: 160, height: 107 },
      "16:9": { x: 0, y: 15, width: 160, height: 90 },
    });
  });

  it("moves edges and corners by whole pixels, freely, inside the bounds", () => {
    const rect = { x: 10, y: 10, width: 100, height: 50 };
    expect(moveCropHandle(rect, "e", 1, 0, null, bounds)).toEqual({ x: 10, y: 10, width: 101, height: 50 });
    expect(moveCropHandle(rect, "w", -10, 0, null, bounds)).toEqual({ x: 0, y: 10, width: 110, height: 50 });
    expect(moveCropHandle(rect, "nw", 10, 10, null, bounds)).toEqual({ x: 20, y: 20, width: 90, height: 40 });
    // Past the image edge: stopped at it.
    expect(moveCropHandle(rect, "s", 0, 500, null, bounds)).toEqual({ x: 10, y: 10, width: 100, height: 110 });
    // Never inverted: at least one pixel.
    expect(moveCropHandle(rect, "e", -500, 0, null, bounds)).toEqual({ x: 10, y: 10, width: 1, height: 50 });
    // The body moves whole, clamped.
    expect(moveCropHandle(rect, "move", 1, -1, null, bounds)).toEqual({ x: 11, y: 9, width: 100, height: 50 });
    expect(moveCropHandle(rect, "move", 100, 100, null, bounds)).toEqual({ x: 60, y: 70, width: 100, height: 50 });
  });

  it("holds a locked aspect, anchoring the opposite corner", () => {
    const rect = { x: 0, y: 15, width: 160, height: 90 };
    const ratio = 16 / 9;
    // The south-east corner in 16 px: 144 wide → 81 tall, top-left fixed.
    expect(moveCropHandle(rect, "se", -16, 0, ratio, bounds)).toEqual({ x: 0, y: 15, width: 144, height: 81 });
    // The north-west corner by 16: anchored bottom-right at (160, 105).
    expect(moveCropHandle(rect, "nw", 16, 0, ratio, bounds)).toEqual({ x: 16, y: 24, width: 144, height: 81 });
    // Growing past the bounds is stopped at them, the aspect kept.
    expect(moveCropHandle(rect, "se", 50, 0, ratio, bounds)).toEqual({ x: 0, y: 15, width: 160, height: 90 });
  });

  it("announces a rectangle as 'x, y, width × height'", () => {
    expect(cropAnnouncement({ x: 0, y: 15, width: 160, height: 90 })).toBe("0, 15, 160 × 90");
  });
});

describe("descriptions", () => {
  it("names every edit as the history list (and the undo label) reads it", () => {
    expect(EDIT_KINDS.map((kind) => describeEdit(SAMPLES[kind][0]))).toEqual([
      "Straighten +5°",
      "Crop",
      "Rotate right",
      "Flip horizontal",
      "Resize to 80 × 45",
      "Adjust brightness and contrast",
      "Redact",
      "Add arrow",
    ]);
    expect(describeEdit({ kind: "adjust", brightness: 30 })).toBe("Adjust brightness");
    expect(describeEdit({ kind: "rotate", quarterTurns: 3 })).toBe("Rotate left");
    expect(describeEdit({ kind: "rotate", quarterTurns: 2 })).toBe("Rotate 180°");
    expect(describeEdit(SAMPLES.annotate[1])).toBe("Add text");
    expect(describeEditDetail(SAMPLES.crop[0])).toBe("0, 15, 160 × 90");
    expect(describeEditDetail({ kind: "adjust", brightness: 30, invert: true })).toBe("brightness +30, invert");
  });

  it("moves objects without touching other edits", () => {
    expect(translateEdit(SAMPLES.redact[0], 2, 3)).toMatchObject({ x: 3, y: 5 });
    expect(translateEdit(SAMPLES.annotate[0], 1, -1)).toMatchObject({ shape: { x1: 1, y1: -1, x2: 11, y2: 9 } });
    expect(translateEdit(SAMPLES.crop[0], 5, 5)).toBe(SAMPLES.crop[0]);
  });

  it("resizes objects from the keyboard, never below one pixel", () => {
    expect(resizeObjectEdit(SAMPLES.redact[0], 10, -1)).toMatchObject({ x: 1, y: 2, width: 13, height: 3 });
    expect(resizeObjectEdit(SAMPLES.redact[0], -50, -50)).toMatchObject({ width: 1, height: 1 });
    expect(resizeObjectEdit(SAMPLES.annotate[0], 5, 0)).toMatchObject({ shape: { x1: 0, y1: 0, x2: 15, y2: 10 } });
    // Text grows with Up/Right and shrinks with Down/Left.
    expect(resizeObjectEdit(SAMPLES.annotate[1], 0, -1)).toMatchObject({ shape: { fontSize: 25 } });
    expect(resizeObjectEdit(SAMPLES.crop[0], 5, 5)).toBe(SAMPLES.crop[0]);
  });
});
