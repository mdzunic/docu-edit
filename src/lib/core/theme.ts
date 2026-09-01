// The shared document-theme model (SPEC-094; SPEC-077 owns the concept): a
// named scheme of twelve colours and two fonts, the `ModelColor`/`ModelFont`
// widenings every colour/font field takes, and the single resolution point
// (`resolveColor`/`resolveFont`) — nothing else interprets a `ModelColor`.
// Product-neutral: both the word processor and the presentation product import
// this file, so it lives in `core` (the only layer `lib/presentation` may
// reach). Plain serialisable data only.
//
// The colour maths (`applyTransform` and the sRGB↔HSL helpers) is hoisted here
// from `presentation/pptx/color.ts` (SPEC-079), which re-imports it — the
// transform semantics are ECMA-376 §20.1.2.3 (shade c·val, tint c·val+(1−val),
// lumMod/lumOff via HSL luminance), unchanged by the move.

/** The twelve scheme slots of ECMA-376's `a:clrScheme`, in scheme order. */
export type ThemeColorSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

/** Every slot, in scheme order — the panel's Customise inputs iterate this. */
export const THEME_COLOR_SLOTS: ThemeColorSlot[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

export interface DocumentTheme {
  name: string;
  /** Slot → lowercase `#rrggbb`. */
  colors: Record<ThemeColorSlot, string>;
  /** Heading (`major`) and body (`minor`) latin typefaces. */
  fonts: { major: string; minor: string };
}

/** A scheme reference plus transforms, each a 0..1 fraction (wire `val`
 * divided by 100000). Transforms apply in the canonical order lumMod, lumOff,
 * shade, tint — the order the pptx writer emits (SPEC-094 AC-17) and the only
 * order the reader captures a reference in. */
export interface ThemeColorRef {
  theme: ThemeColorSlot;
  lumMod?: number;
  lumOff?: number;
  shade?: number;
  tint?: number;
}

/** Widens every colour field: a literal `#rrggbb` or a scheme reference. */
export type ModelColor = string | ThemeColorRef;

/** Widens every font-family field: a literal family or a scheme reference. */
export type ModelFont = string | { theme: "major" | "minor" };

/** Whether a `ModelColor` is a scheme reference (vs. a literal). */
export function isThemeColor(color: ModelColor | undefined): color is ThemeColorRef {
  return typeof color === "object" && color !== null;
}

/** Whether a `ModelFont` is a scheme reference (vs. a literal family). */
export function isThemeFont(font: ModelFont | undefined): font is { theme: "major" | "minor" } {
  return typeof font === "object" && font !== null;
}

// --- colour maths (hoisted from presentation/pptx/color.ts, SPEC-079) -------

export function clamp01(value: number): number {
  return Math.min(1, Math.max(0, value));
}

export function hexToRgb(hex: string): [number, number, number] {
  const n = hex.replace("#", "");
  return [parseInt(n.slice(0, 2), 16) / 255, parseInt(n.slice(2, 4), 16) / 255, parseInt(n.slice(4, 6), 16) / 255];
}

export function rgbToHex(rgb: [number, number, number]): string {
  return `#${rgb.map((c) => Math.round(clamp01(c) * 255).toString(16).padStart(2, "0")).join("")}`;
}

export function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) {
    return [0, 0, l];
  }
  const s = d / (1 - Math.abs(2 * l - 1));
  let h: number;
  if (max === r) {
    h = 60 * (((g - b) / d) % 6);
  } else if (max === g) {
    h = 60 * ((b - r) / d + 2);
  } else {
    h = 60 * ((r - g) / d + 4);
  }
  if (h < 0) {
    h += 360;
  }
  return [h, s, l];
}

export function hslToRgb([h, s, l]: [number, number, number]): [number, number, number] {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let rgb: [number, number, number];
  if (h < 60) rgb = [c, x, 0];
  else if (h < 120) rgb = [x, c, 0];
  else if (h < 180) rgb = [0, c, x];
  else if (h < 240) rgb = [0, x, c];
  else if (h < 300) rgb = [x, 0, c];
  else rgb = [c, 0, x];
  return [rgb[0] + m, rgb[1] + m, rgb[2] + m];
}

/** Applies one DrawingML colour transform (`shade`/`tint`/`lumMod`/`lumOff`)
 * to an sRGB colour, `val` already normalised to a 0..1 fraction. Other
 * transforms are ignored (returned unchanged). */
export function applyTransform(rgb: [number, number, number], name: string, val: number): [number, number, number] {
  if (name === "shade") {
    return [rgb[0] * val, rgb[1] * val, rgb[2] * val];
  }
  if (name === "tint") {
    return [rgb[0] * val + (1 - val), rgb[1] * val + (1 - val), rgb[2] * val + (1 - val)];
  }
  if (name === "lumMod" || name === "lumOff") {
    const [h, s, l] = rgbToHsl(rgb);
    const l2 = clamp01(name === "lumMod" ? l * val : l + val);
    return hslToRgb([h, s, l2]);
  }
  return rgb;
}

// --- the single resolution point --------------------------------------------

const TRANSFORM_ORDER = ["lumMod", "lumOff", "shade", "tint"] as const;

/** Resolves a `ModelColor` to a lowercase `#rrggbb` against a theme — the
 * single resolution point; nothing else interprets a ModelColor. A literal
 * passes through untouched; a reference looks its slot up in the scheme and
 * applies its transforms in the canonical order. */
export function resolveColor(color: ModelColor, theme: DocumentTheme): string {
  if (!isThemeColor(color)) {
    return color;
  }
  let rgb = hexToRgb(theme.colors[color.theme]);
  for (const name of TRANSFORM_ORDER) {
    const val = color[name];
    if (val !== undefined) {
      rgb = applyTransform(rgb, name, val);
    }
  }
  return rgbToHex(rgb);
}

/** Resolves a `ModelFont` to a concrete family name against a theme. */
export function resolveFont(font: ModelFont, theme: DocumentTheme): string {
  if (!isThemeFont(font)) {
    return font;
  }
  return font.theme === "major" ? theme.fonts.major : theme.fonts.minor;
}

/** Structural equality over `ModelColor`s (two equal refs are the same colour
 * even as distinct objects — run coalescing and toolbars fold by this). */
export function sameModelColor(a: ModelColor | undefined, b: ModelColor | undefined): boolean {
  if (a === b) {
    return true;
  }
  if (!isThemeColor(a) || !isThemeColor(b)) {
    return false;
  }
  return (
    a.theme === b.theme && a.lumMod === b.lumMod && a.lumOff === b.lumOff && a.shade === b.shade && a.tint === b.tint
  );
}

/** Structural equality over `ModelFont`s. */
export function sameModelFont(a: ModelFont | undefined, b: ModelFont | undefined): boolean {
  if (a === b) {
    return true;
  }
  if (!isThemeFont(a) || !isThemeFont(b)) {
    return false;
  }
  return a.theme === b.theme;
}

// --- built-in themes ---------------------------------------------------------

/** Office — the default theme every unthemed document resolves against. */
export const DEFAULT_THEME: DocumentTheme = {
  name: "Office",
  colors: {
    dk1: "#000000",
    lt1: "#ffffff",
    dk2: "#44546a",
    lt2: "#e7e6e6",
    accent1: "#4472c4",
    accent2: "#ed7d31",
    accent3: "#a5a5a5",
    accent4: "#ffc000",
    accent5: "#5b9bd5",
    accent6: "#70ad47",
    hlink: "#0563c1",
    folHlink: "#954f72",
  },
  fonts: { major: "Calibri Light", minor: "Calibri" },
};

/** The six built-in themes the gallery offers, Office first. */
export const BUILT_IN_THEMES: DocumentTheme[] = [
  DEFAULT_THEME,
  {
    name: "Facet",
    colors: {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#2c3c43",
      lt2: "#ebebeb",
      accent1: "#90c226",
      accent2: "#54a021",
      accent3: "#e6b91e",
      accent4: "#e76618",
      accent5: "#c42f1a",
      accent6: "#918655",
      hlink: "#99ca3c",
      folHlink: "#b26b02",
    },
    fonts: { major: "Trebuchet MS", minor: "Trebuchet MS" },
  },
  {
    name: "Ion",
    colors: {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#1b587c",
      lt2: "#ccdaea",
      accent1: "#b01513",
      accent2: "#ea6312",
      accent3: "#e6b729",
      accent4: "#6aac90",
      accent5: "#5f9c9d",
      accent6: "#9b75b2",
      hlink: "#6b9f25",
      folHlink: "#b26b02",
    },
    fonts: { major: "Century Gothic", minor: "Century Gothic" },
  },
  {
    name: "Retrospect",
    colors: {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#637052",
      lt2: "#ccddea",
      accent1: "#e48312",
      accent2: "#bd582c",
      accent3: "#865640",
      accent4: "#9b8357",
      accent5: "#c2bc80",
      accent6: "#94a088",
      hlink: "#2998e3",
      folHlink: "#8c8c8c",
    },
    fonts: { major: "Calibri Light", minor: "Calibri" },
  },
  {
    name: "Slice",
    colors: {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#242852",
      lt2: "#dcdee8",
      accent1: "#1cade4",
      accent2: "#2683c6",
      accent3: "#27ced7",
      accent4: "#42ba97",
      accent5: "#3e8853",
      accent6: "#62a39f",
      hlink: "#6b9f25",
      folHlink: "#b26b02",
    },
    fonts: { major: "Century Gothic", minor: "Verdana" },
  },
  {
    name: "Wisp",
    colors: {
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#766f54",
      lt2: "#e3eed4",
      accent1: "#a53010",
      accent2: "#de7e18",
      accent3: "#9f8351",
      accent4: "#728653",
      accent5: "#92aa4c",
      accent6: "#6aac91",
      hlink: "#fb4a18",
      folHlink: "#ba8422",
    },
    fonts: { major: "Century Gothic", minor: "Corbel" },
  },
];

// --- picker variants ---------------------------------------------------------

/** One lightness variant a theme-swatch column offers under its base colour —
 * PowerPoint's five (Lighter 80/60/40, Darker 25/50), as lumMod/lumOff
 * fractions. */
export interface ThemeSwatchVariant {
  /** Suffix of the swatch's accessible name, e.g. "lighter 40%". */
  label: string;
  lumMod?: number;
  lumOff?: number;
}

/** The five lumMod/lumOff variants each theme swatch offers (SPEC-094 AC),
 * lightest first — matching PowerPoint's theme-colour columns. */
export const THEME_SWATCH_VARIANTS: ThemeSwatchVariant[] = [
  { label: "lighter 80%", lumMod: 0.2, lumOff: 0.8 },
  { label: "lighter 60%", lumMod: 0.4, lumOff: 0.6 },
  { label: "lighter 40%", lumMod: 0.6, lumOff: 0.4 },
  { label: "darker 25%", lumMod: 0.75 },
  { label: "darker 50%", lumMod: 0.5 },
];

/** The ten slots a picker's theme row shows, in PowerPoint's column order. */
export const THEME_SWATCH_SLOTS: ThemeColorSlot[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
];

const SLOT_LABELS: Record<ThemeColorSlot, string> = {
  dk1: "Dark 1",
  lt1: "Light 1",
  dk2: "Dark 2",
  lt2: "Light 2",
  accent1: "Accent 1",
  accent2: "Accent 2",
  accent3: "Accent 3",
  accent4: "Accent 4",
  accent5: "Accent 5",
  accent6: "Accent 6",
  hlink: "Hyperlink",
  folHlink: "Followed Hyperlink",
};

/** The human name of a scheme slot ("accent1" → "Accent 1"). */
export function themeSlotLabel(slot: ThemeColorSlot): string {
  return SLOT_LABELS[slot];
}
