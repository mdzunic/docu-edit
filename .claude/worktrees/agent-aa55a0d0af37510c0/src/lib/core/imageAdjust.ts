// The image editor's pixel maths (SPEC-112) — THIS FILE is the normative
// source for the six adjust channels. Each is stated as arithmetic over an
// RGBA `Uint8ClampedArray` rather than delegated to a CSS filter string,
// whose results vary between engines: the on-screen preview (a downscaled
// buffer) and every full-resolution export run these same functions, so a
// brightness of +30 is the same +30 everywhere.
//
// Every function is pure — it reads `pixels` and writes a fresh array —
// unless the caller passes `out`, which may be `pixels` itself for an
// in-place pass over a large buffer (the renderer does, band by band).
// Alpha (every fourth byte) is copied through untouched. Each channel rounds
// its result half-up (Math.round) before the array clamps it to 0..255, so a
// hand-computed expectation is one rounding per channel, never the typed
// array's round-half-to-even.

/** The six adjust channels, in the order `applyAdjustments` applies them. */
export const ADJUST_CHANNELS = ["brightness", "contrast", "saturation", "exposure", "grayscale", "invert"] as const;

export type AdjustChannel = (typeof ADJUST_CHANNELS)[number];

/** The slider channels' range: −100..100, 0 neutral. */
export const ADJUST_MIN = -100;
export const ADJUST_MAX = 100;

/** The adjust edit's settings (the `adjust` ImageEdit without its kind). An
 * absent channel is neutral. */
export interface AdjustSettings {
  brightness?: number;
  contrast?: number;
  saturation?: number;
  exposure?: number;
  grayscale?: boolean;
  invert?: boolean;
}

/** Rec. 709 luma weights — saturation interpolates towards this, and
 * grayscale writes it. */
export const LUMA_R = 0.2126;
export const LUMA_G = 0.7152;
export const LUMA_B = 0.0722;

/** Rec. 709 luma of one pixel. */
export function luma(r: number, g: number, b: number): number {
  return LUMA_R * r + LUMA_G * g + LUMA_B * b;
}

/** A slider value clamped to −100..100 (non-finite reads as 0). */
export function clampAdjust(value: number): number {
  return Number.isFinite(value) ? Math.min(ADJUST_MAX, Math.max(ADJUST_MIN, value)) : 0;
}

/** The contrast multiplier for a −100..100 setting:
 * f = (259 · (c + 255)) / (255 · (259 − c)). */
export function contrastFactor(value: number): number {
  const c = clampAdjust(value);
  return (259 * (c + 255)) / (255 * (259 - c));
}

/** The exposure multiplier for a −100..100 setting: 2 ** (v / 50). */
export function exposureFactor(value: number): number {
  return 2 ** (clampAdjust(value) / 50);
}

function target(pixels: Uint8ClampedArray, out?: Uint8ClampedArray): Uint8ClampedArray {
  return out ?? new Uint8ClampedArray(pixels.length);
}

/** Maps every pixel's R, G and B through `channel`, copying alpha. */
function perChannel(
  pixels: Uint8ClampedArray,
  out: Uint8ClampedArray | undefined,
  channel: (value: number) => number,
): Uint8ClampedArray {
  const result = target(pixels, out);
  for (let i = 0; i < pixels.length; i += 4) {
    result[i] = Math.round(channel(pixels[i]));
    result[i + 1] = Math.round(channel(pixels[i + 1]));
    result[i + 2] = Math.round(channel(pixels[i + 2]));
    result[i + 3] = pixels[i + 3];
  }
  return result;
}

/** Brightness: an additive term of `v · 2.55` on R, G and B. */
export function brightness(pixels: Uint8ClampedArray, value: number, out?: Uint8ClampedArray): Uint8ClampedArray {
  const term = clampAdjust(value) * 2.55;
  return perChannel(pixels, out, (channel) => channel + term);
}

/** Contrast: `f · (channel − 128) + 128` with f from `contrastFactor`. */
export function contrast(pixels: Uint8ClampedArray, value: number, out?: Uint8ClampedArray): Uint8ClampedArray {
  const factor = contrastFactor(value);
  return perChannel(pixels, out, (channel) => factor * (channel - 128) + 128);
}

/** Exposure: each channel multiplied by `2 ** (v / 50)`. */
export function exposure(pixels: Uint8ClampedArray, value: number, out?: Uint8ClampedArray): Uint8ClampedArray {
  const factor = exposureFactor(value);
  return perChannel(pixels, out, (channel) => channel * factor);
}

/** Inversion: `255 − channel`. */
export function invert(pixels: Uint8ClampedArray, out?: Uint8ClampedArray): Uint8ClampedArray {
  return perChannel(pixels, out, (channel) => 255 - channel);
}

/** Saturation: each channel interpolated towards (v < 0) or away from
 * (v > 0) the pixel's Rec. 709 luma — `L + (channel − L) · (1 + v / 100)`,
 * so −100 is the luma itself and +100 doubles every channel's distance
 * from it. */
export function saturation(pixels: Uint8ClampedArray, value: number, out?: Uint8ClampedArray): Uint8ClampedArray {
  const scale = 1 + clampAdjust(value) / 100;
  const result = target(pixels, out);
  for (let i = 0; i < pixels.length; i += 4) {
    const r = pixels[i];
    const g = pixels[i + 1];
    const b = pixels[i + 2];
    const l = luma(r, g, b);
    result[i] = Math.round(l + (r - l) * scale);
    result[i + 1] = Math.round(l + (g - l) * scale);
    result[i + 2] = Math.round(l + (b - l) * scale);
    result[i + 3] = pixels[i + 3];
  }
  return result;
}

/** Grayscale: R, G and B all become the pixel's Rec. 709 luma. */
export function grayscale(pixels: Uint8ClampedArray, out?: Uint8ClampedArray): Uint8ClampedArray {
  const result = target(pixels, out);
  for (let i = 0; i < pixels.length; i += 4) {
    const l = Math.round(luma(pixels[i], pixels[i + 1], pixels[i + 2]));
    result[i] = l;
    result[i + 1] = l;
    result[i + 2] = l;
    result[i + 3] = pixels[i + 3];
  }
  return result;
}

/** Whether a channel's setting changes nothing. */
export function isNeutralChannel(settings: AdjustSettings, channel: AdjustChannel): boolean {
  const value = settings[channel];
  if (channel === "grayscale" || channel === "invert") {
    return value !== true;
  }
  return typeof value !== "number" || clampAdjust(value) === 0;
}

/** Whether every channel is neutral (the settings change nothing). */
export function isNeutralAdjust(settings: AdjustSettings): boolean {
  return ADJUST_CHANNELS.every((channel) => isNeutralChannel(settings, channel));
}

/** Every non-neutral channel applied in ADJUST_CHANNELS order. Pure unless
 * `out` is given; with nothing to apply it returns a copy (or `out`
 * holding one). */
export function applyAdjustments(
  pixels: Uint8ClampedArray,
  settings: AdjustSettings,
  out?: Uint8ClampedArray,
): Uint8ClampedArray {
  let current = pixels;
  let result = out;
  const step = (apply: (source: Uint8ClampedArray, into: Uint8ClampedArray | undefined) => Uint8ClampedArray): void => {
    current = apply(current, result);
    // After the first pass the working buffer is our own (or the caller's
    // `out`), so every later pass may write over it in place.
    result = current;
  };
  if (!isNeutralChannel(settings, "brightness")) {
    step((source, into) => brightness(source, settings.brightness ?? 0, into));
  }
  if (!isNeutralChannel(settings, "contrast")) {
    step((source, into) => contrast(source, settings.contrast ?? 0, into));
  }
  if (!isNeutralChannel(settings, "saturation")) {
    step((source, into) => saturation(source, settings.saturation ?? 0, into));
  }
  if (!isNeutralChannel(settings, "exposure")) {
    step((source, into) => exposure(source, settings.exposure ?? 0, into));
  }
  if (settings.grayscale === true) {
    step((source, into) => grayscale(source, into));
  }
  if (settings.invert === true) {
    step((source, into) => invert(source, into));
  }
  if (current === pixels) {
    const copy = target(pixels, out);
    if (copy !== pixels) {
      copy.set(pixels);
    }
    return copy;
  }
  return current;
}
