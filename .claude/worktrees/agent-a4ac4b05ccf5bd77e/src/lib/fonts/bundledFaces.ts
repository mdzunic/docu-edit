// The bundled font faces the app registers as @font-face so the browser can
// render the metric-compatible substitutes (SPEC-048 AC-1). These are the
// same files the PDF exporter embeds (lib/core/pdfFonts), listed here by family
// × weight × slant for the CSS Font Loading API. Framework-free (no Vue, no
// FontFace construction — that is the browser-only composable's job).
import { BUNDLED_FAMILIES, type BundledFamily } from "../core/fontSubstitution";

export interface BundledFace {
  family: BundledFamily;
  weight: "normal" | "bold";
  style: "normal" | "italic";
  /** File under public/fonts/ (served at `${BASE_URL}fonts/<file>`). */
  file: string;
}

/** "Liberation Sans" → "LiberationSans" — the family's file-name stem. */
function stem(family: BundledFamily): string {
  return family.replace(/\s+/g, "");
}

const STYLE_SUFFIX: Record<string, string> = {
  "normal-normal": "Regular",
  "bold-normal": "Bold",
  "normal-italic": "Italic",
  "bold-italic": "BoldItalic",
};

function facesFor(family: BundledFamily): BundledFace[] {
  const combos: Array<{ weight: "normal" | "bold"; style: "normal" | "italic" }> = [
    { weight: "normal", style: "normal" },
    { weight: "bold", style: "normal" },
    { weight: "normal", style: "italic" },
    { weight: "bold", style: "italic" },
  ];
  return combos.map(({ weight, style }) => ({
    family,
    weight,
    style,
    file: `${stem(family)}-${STYLE_SUFFIX[`${weight}-${style}`]}.ttf`,
  }));
}

/** Every bundled face (20: five families × four styles). */
export const BUNDLED_FACES: readonly BundledFace[] = BUNDLED_FAMILIES.flatMap(facesFor);
