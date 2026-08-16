// Maps the CSS font families the canvas actually renders with onto the
// bundled faces the PDF embeds (SPEC-034 AC-20). The app never ships the
// document's own font files, so export substitutes the metric-compatible
// bundled family — Carlito for Calibri, Caladea for Cambria, else the
// Liberation family of the same classification (sans/serif/mono) × regular/
// bold/italic — exactly like the screen falls back for uninstalled fonts
// (SPEC-048). Because both the screen and this classifier key off the same
// first family, the PDF embeds the same face the screen rendered.
// Framework-free per the src/lib convention.
import { substituteFor } from "../fonts/substitution";

export type FaceKey =
  | "sans-regular"
  | "sans-bold"
  | "sans-italic"
  | "sans-bolditalic"
  | "serif-regular"
  | "serif-bold"
  | "serif-italic"
  | "serif-bolditalic"
  | "mono-regular"
  | "mono-bold"
  | "mono-italic"
  | "mono-bolditalic"
  | "calibri-regular"
  | "calibri-bold"
  | "calibri-italic"
  | "calibri-bolditalic"
  | "cambria-regular"
  | "cambria-bold"
  | "cambria-italic"
  | "cambria-bolditalic";

const SERIF_NAMES = [
  "times",
  "georgia",
  "garamond",
  "cambria",
  "book antiqua",
  "palatino",
  "liberation serif",
  "tinos",
  "dejavu serif",
  "noto serif",
];

const MONO_NAMES = [
  "courier",
  "consolas",
  "menlo",
  "monaco",
  "liberation mono",
  "cousine",
  "dejavu sans mono",
  "noto sans mono",
  "source code",
  "fira code",
];

/** The first family name of a CSS font-family list, unquoted, lowercase. */
export function firstFamilyName(cssFontFamily: string): string {
  const first = cssFontFamily.split(",")[0] ?? "";
  return first.trim().replace(/^["']|["']$/g, "").toLowerCase();
}

/** Classifies a CSS font-family into the Liberation family that stands in
 * for it: monospace → Mono, serif → Serif, everything else → Sans (the
 * document default Calibri, Arial, Helvetica, system-ui … are all sans). */
export function classifyFamily(cssFontFamily: string): "sans" | "serif" | "mono" {
  const name = firstFamilyName(cssFontFamily);
  if (name === "monospace" || MONO_NAMES.some((mono) => name.includes(mono)) || name.includes("mono")) {
    return "mono";
  }
  if (name === "serif" || SERIF_NAMES.some((serif) => name.includes(serif))) {
    return "serif";
  }
  return "sans";
}

/** The metric-twin face group for a family that has a bundled twin — Calibri
 * → Carlito, Cambria → Caladea — so the PDF embeds the same twin the screen
 * rendered rather than the generic Liberation classification (which would put
 * Calibri on Arial metrics). Returns null for families we classify generically. */
function twinGroup(cssFontFamily: string): "calibri" | "cambria" | null {
  const sub = substituteFor(firstFamilyName(cssFontFamily));
  if (sub?.family === "Carlito") {
    return "calibri";
  }
  if (sub?.family === "Caladea") {
    return "cambria";
  }
  return null;
}

/** The face a rendered run maps to, from its computed style. */
export function faceKeyFor(cssFontFamily: string, bold: boolean, italic: boolean): FaceKey {
  const family = twinGroup(cssFontFamily) ?? classifyFamily(cssFontFamily);
  const style = bold && italic ? "bolditalic" : bold ? "bold" : italic ? "italic" : "regular";
  return `${family}-${style}`;
}

const FACE_FILES: Record<FaceKey, string> = {
  "sans-regular": "LiberationSans-Regular.ttf",
  "sans-bold": "LiberationSans-Bold.ttf",
  "sans-italic": "LiberationSans-Italic.ttf",
  "sans-bolditalic": "LiberationSans-BoldItalic.ttf",
  "serif-regular": "LiberationSerif-Regular.ttf",
  "serif-bold": "LiberationSerif-Bold.ttf",
  "serif-italic": "LiberationSerif-Italic.ttf",
  "serif-bolditalic": "LiberationSerif-BoldItalic.ttf",
  "mono-regular": "LiberationMono-Regular.ttf",
  "mono-bold": "LiberationMono-Bold.ttf",
  "mono-italic": "LiberationMono-Italic.ttf",
  "mono-bolditalic": "LiberationMono-BoldItalic.ttf",
  "calibri-regular": "Carlito-Regular.ttf",
  "calibri-bold": "Carlito-Bold.ttf",
  "calibri-italic": "Carlito-Italic.ttf",
  "calibri-bolditalic": "Carlito-BoldItalic.ttf",
  "cambria-regular": "Caladea-Regular.ttf",
  "cambria-bold": "Caladea-Bold.ttf",
  "cambria-italic": "Caladea-Italic.ttf",
  "cambria-bolditalic": "Caladea-BoldItalic.ttf",
};

/** The bundled font file (under /fonts/) that carries a face. */
export function faceFileName(face: FaceKey): string {
  return FACE_FILES[face];
}

/** Every face key, for tests and preloading. */
export function allFaceKeys(): FaceKey[] {
  return Object.keys(FACE_FILES) as FaceKey[];
}
