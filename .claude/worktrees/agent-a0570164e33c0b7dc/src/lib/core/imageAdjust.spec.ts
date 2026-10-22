// SPEC-112: the adjust channels asserted against HAND-COMPUTED values, not
// snapshots — one table per channel at 0, ±50 and ±100, each expectation
// worked through in the comment beside it (one half-up rounding per
// channel, then the 0..255 clamp). Alpha is untouched throughout.
import { describe, expect, it } from "vitest";
import {
  applyAdjustments,
  brightness,
  contrast,
  contrastFactor,
  exposure,
  exposureFactor,
  grayscale,
  invert,
  isNeutralAdjust,
  luma,
  saturation,
} from "./imageAdjust";

/** One RGBA pixel with a distinctive alpha, so "untouched" is checkable. */
function pixel(r: number, g: number, b: number, a = 77): Uint8ClampedArray {
  return new Uint8ClampedArray([r, g, b, a]);
}

function rgb(data: Uint8ClampedArray): number[] {
  return [data[0], data[1], data[2]];
}

describe("brightness (additive v · 2.55)", () => {
  it.each([
    // 0 → +0
    [0, [100, 150, 200], [100, 150, 200]],
    // +50 → +127.5: 227.5→228, 277.5→255, 327.5→255
    [50, [100, 150, 200], [228, 255, 255]],
    // +50 on a darker pixel: 137.5→138, 187.5→188, 247.5→248
    [50, [10, 60, 120], [138, 188, 248]],
    // −50 → −127.5: −27.5→0, 22.5→23, 72.5→73
    [-50, [100, 150, 200], [0, 23, 73]],
    // +100 → +255: everything white
    [100, [100, 150, 200], [255, 255, 255]],
    // −100 → −255: everything black
    [-100, [100, 150, 200], [0, 0, 0]],
  ])("%i on %j → %j", (value, input, expected) => {
    const source = pixel(input[0], input[1], input[2]);
    const out = brightness(source, value);
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });
});

describe("contrast (f = 259(c + 255) / 255(259 − c), about 128)", () => {
  it("states the factor", () => {
    expect(contrastFactor(0)).toBe(1);
    // 259·305 / 255·209 = 78995 / 53295
    expect(contrastFactor(50)).toBeCloseTo(78995 / 53295, 12);
    // 259·205 / 255·309 = 53095 / 78795
    expect(contrastFactor(-50)).toBeCloseTo(53095 / 78795, 12);
    // 259·355 / 255·159 = 91945 / 40545
    expect(contrastFactor(100)).toBeCloseTo(91945 / 40545, 12);
    // 259·155 / 255·359 = 40145 / 91545
    expect(contrastFactor(-100)).toBeCloseTo(40145 / 91545, 12);
  });

  it.each([
    [0, [100, 150, 200]],
    // f = 1.48222: −28·f+128 = 86.50→86, 22·f+128 = 160.61→161, 72·f+128 = 234.72→235
    [50, [86, 161, 235]],
    // f = 0.67384: 109.13→109, 142.82→143, 176.52→177
    [-50, [109, 143, 177]],
    // f = 2.26773: 64.50→65, 177.89→178, 291.28→255
    [100, [65, 178, 255]],
    // f = 0.43853: 115.72→116, 137.65→138, 159.57→160
    [-100, [116, 138, 160]],
  ])("%i on [100, 150, 200] → %j", (value, expected) => {
    const out = contrast(pixel(100, 150, 200), value);
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });
});

describe("saturation (towards Rec. 709 luma)", () => {
  it("uses Rec. 709 luma", () => {
    // 0.2126·100 + 0.7152·150 + 0.0722·200 = 21.26 + 107.28 + 14.44
    expect(luma(100, 150, 200)).toBeCloseTo(142.98, 10);
  });

  it.each([
    [0, [100, 150, 200]],
    // ×1.5 from L = 142.98: 78.51→79, 153.51→154, 228.51→229
    [50, [79, 154, 229]],
    // ×0.5: 121.49→121, 146.49→146, 171.49→171
    [-50, [121, 146, 171]],
    // ×2: 57.02→57, 157.02→157, 257.02→255
    [100, [57, 157, 255]],
    // ×0: the luma itself, 142.98→143
    [-100, [143, 143, 143]],
  ])("%i on [100, 150, 200] → %j", (value, expected) => {
    const out = saturation(pixel(100, 150, 200), value);
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });
});

describe("exposure (× 2 ** (v / 50))", () => {
  it("states the multiplier", () => {
    expect([0, 50, -50, 100, -100].map(exposureFactor)).toEqual([1, 2, 0.5, 4, 0.25]);
  });

  it.each([
    [0, [10, 40, 60]],
    // ×2
    [50, [20, 80, 120]],
    // ×0.5
    [-50, [5, 20, 30]],
    // ×4
    [100, [40, 160, 240]],
    // ×0.25: 2.5→3, 10, 15
    [-100, [3, 10, 15]],
  ])("%i on [10, 40, 60] → %j", (value, expected) => {
    const out = exposure(pixel(10, 40, 60), value);
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });
});

describe("grayscale (the same luma)", () => {
  it.each([
    // 142.98 → 143
    [[100, 150, 200], [143, 143, 143]],
    // 2.126 + 28.608 + 4.332 = 35.066 → 35
    [[10, 40, 60], [35, 35, 35]],
    [[0, 0, 0], [0, 0, 0]],
    [[255, 255, 255], [255, 255, 255]],
  ])("%j → %j", (input, expected) => {
    const out = grayscale(pixel(input[0], input[1], input[2]));
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });
});

describe("invert (255 − channel)", () => {
  it.each([
    [[100, 150, 200], [155, 105, 55]],
    [[0, 255, 1], [255, 0, 254]],
  ])("%j → %j", (input, expected) => {
    const out = invert(pixel(input[0], input[1], input[2]));
    expect(rgb(out)).toEqual(expected);
    expect(out[3]).toBe(77);
  });

  it("round-trips: invert(invert(x)) is x, for every byte value", () => {
    const x = new Uint8ClampedArray(256 * 4);
    for (let v = 0; v < 256; v += 1) {
      x.set([v, 255 - v, (v * 7) % 256, v], v * 4);
    }
    expect(invert(invert(x))).toEqual(x);
  });
});

describe("purity and order", () => {
  it("never mutates its input unless handed it as `out`", () => {
    const source = pixel(100, 150, 200);
    brightness(source, 50);
    contrast(source, 50);
    saturation(source, 50);
    exposure(source, 50);
    grayscale(source);
    invert(source);
    expect(Array.from(source)).toEqual([100, 150, 200, 77]);
    // In place when asked.
    const same = brightness(source, 10, source);
    expect(same).toBe(source);
    expect(rgb(source)).toEqual([126, 176, 226]);
  });

  it("applies the channels brightness → contrast → saturation → exposure → grayscale → invert", () => {
    const settings = { brightness: 10, contrast: 20, saturation: -30, exposure: 15, grayscale: true, invert: true };
    const source = pixel(100, 150, 200);
    const stepwise = invert(
      grayscale(exposure(saturation(contrast(brightness(source, 10), 20), -30), 15)),
    );
    expect(applyAdjustments(source, settings)).toEqual(stepwise);
    expect(Array.from(source)).toEqual([100, 150, 200, 77]);
    // The order matters: invert-then-brighten is a different picture.
    expect(rgb(brightness(invert(source), 50))).not.toEqual(rgb(invert(brightness(source, 50))));
  });

  it("treats neutral settings as no change", () => {
    expect(isNeutralAdjust({})).toBe(true);
    expect(isNeutralAdjust({ brightness: 0, grayscale: false })).toBe(true);
    expect(isNeutralAdjust({ invert: true })).toBe(false);
    const source = pixel(1, 2, 3, 4);
    const out = applyAdjustments(source, { brightness: 0 });
    expect(out).toEqual(source);
    expect(out).not.toBe(source);
  });
});
