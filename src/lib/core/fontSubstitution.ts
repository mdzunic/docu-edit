// SPEC-048 — metric-compatible font substitution, the single source of
// truth every rendering surface routes a run's family through. The document
// model always keeps the *original* family name (Calibri, Cambria, Arial…);
// substitution is a rendering decision made here and nowhere else, so the
// writers keep emitting the original name (AC-3/4/5).
//
// A document written in Calibri on Windows must reflow identically on a
// machine that has never heard of Calibri. We bundle metric-compatible twins
// — Carlito for Calibri, Caladea for Cambria, Liberation Sans/Serif/Mono for
// Arial/Times New Roman/Courier New — and render/measure/embed with the twin
// whenever the original is unavailable (AC-1, AC-2, AC-6). A family we have
// no twin for falls back to a documented default (AC-9).
//
// Framework-free (no Vue imports) per the src/lib convention.

/** The families we ship as bundled font files (public/fonts/) and register
 * as @font-face faces / embed into exported PDFs. */
export type BundledFamily =
  | "Liberation Sans"
  | "Liberation Serif"
  | "Liberation Mono"
  | "Carlito"
  | "Caladea";

export const BUNDLED_FAMILIES: readonly BundledFamily[] = [
  "Liberation Sans",
  "Liberation Serif",
  "Liberation Mono",
  "Carlito",
  "Caladea",
];

/** The face a genuinely missing font falls back to (AC-9). Documented here so
 * the Fonts panel and the fallback render agree on one answer. Liberation
 * Sans is the neutral, always-bundled default; a missing serif/mono family
 * still falls back through its own generic keyword after this. */
export const DOCUMENTED_DEFAULT_FAMILY: BundledFamily = "Liberation Sans";

type Generic = "sans-serif" | "serif" | "monospace";

export interface Substitute {
  /** The bundled metric-compatible twin. */
  readonly family: BundledFamily;
  /** The CSS generic the twin itself falls back to. */
  readonly generic: Generic;
}

/** Original Microsoft family (lower-cased) → its bundled metric twin. */
const SUBSTITUTES: Readonly<Record<string, Substitute>> = {
  calibri: { family: "Carlito", generic: "sans-serif" },
  cambria: { family: "Caladea", generic: "serif" },
  arial: { family: "Liberation Sans", generic: "sans-serif" },
  "arial narrow": { family: "Liberation Sans", generic: "sans-serif" },
  "times new roman": { family: "Liberation Serif", generic: "serif" },
  "courier new": { family: "Liberation Mono", generic: "monospace" },
};

/** The first family of a CSS font-family list, unquoted and lower-cased. */
function firstName(family: string): string {
  const first = family.split(",")[0] ?? "";
  return first.trim().replace(/^["']|["']$/g, "").toLowerCase();
}

/** The bundled twin for a family, or undefined when we ship none for it. */
export function substituteFor(family: string): Substitute | undefined {
  return SUBSTITUTES[firstName(family)];
}

/** Whether a family is one of the fonts we bundle ourselves (so it renders
 * directly, never itself substituted). */
export function isBundledFamily(family: string): boolean {
  const name = firstName(family);
  return BUNDLED_FAMILIES.some((bundled) => bundled.toLowerCase() === name);
}

const SERIF_HINTS = ["times", "georgia", "garamond", "cambria", "serif", "book antiqua", "palatino"];
const MONO_HINTS = ["courier", "consolas", "menlo", "monaco", "mono"];

/** The CSS generic a family reads as, for the documented-default fallback. */
function genericOf(family: string): Generic {
  const name = firstName(family);
  if (name === "monospace" || MONO_HINTS.some((hint) => name.includes(hint))) {
    return "monospace";
  }
  if (name === "serif" || SERIF_HINTS.some((hint) => name.includes(hint))) {
    return "serif";
  }
  return "sans-serif";
}

/** Quotes a family for a CSS font-family list: a name with spaces must be
 * quoted or the whole declaration is invalid. Generic keywords pass through. */
function quote(family: string): string {
  const trimmed = family.trim();
  if (trimmed === "sans-serif" || trimmed === "serif" || trimmed === "monospace") {
    return trimmed;
  }
  return /^[a-zA-Z][a-zA-Z0-9-]*$/.test(trimmed) ? trimmed : `"${trimmed.replace(/"/g, '\\"')}"`;
}

/** The CSS font-family stack a run's family renders and is measured with.
 * The original name comes first, so an installed original still wins (AC-2's
 * "when the original is unavailable"); the bundled twin — or the documented
 * default for a family we have no twin for — backs it, so a machine without
 * the original reflows on the metric-compatible face. Returns "" for an
 * unset family, so callers leave the inherited base font in place.
 *
 * Because both the screen renderer and the canvas/DOM measurer build their
 * font from this same stack, measure and paint resolve to the same face — and
 * so does the PDF, whose face classifier keys off the same first family. */
export function fontStack(family: string | undefined): string {
  if (!family || family.trim().length === 0) {
    return "";
  }
  const original = quote(family);
  if (isBundledFamily(family)) {
    return `${original}, ${genericOf(family)}`;
  }
  const sub = substituteFor(family);
  if (sub) {
    return `${original}, ${sub.family}, ${sub.generic}`;
  }
  // No twin: fall back to the documented default, then the family's generic.
  return `${original}, ${DOCUMENTED_DEFAULT_FAMILY}, ${genericOf(family)}`;
}
