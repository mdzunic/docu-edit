// SPEC-048 — the "Fonts" panel's data model (AC-7, AC-8, AC-10). Lists every
// family the document uses and classifies each as available (installed
// locally, bundled, or embedded in the file), substituted (rendered on a
// bundled metric twin, which is named), or missing (no twin either — falls
// back to the documented default). Framework-free: local-availability
// detection is injected, so the panel supplies a browser check and tests
// supply a deterministic fake.
import type { DocumentModel } from "../documentModel";
import { documentFontFamilies } from "../fonts";
import { isBundledFamily, substituteFor } from "./substitution";

export type FontStatus = "available" | "substituted" | "missing";

export interface FontReportEntry {
  /** The family exactly as the document names it. */
  family: string;
  status: FontStatus;
  /** The bundled twin actually rendered, present only for "substituted". */
  substitute?: string;
}

export interface FontReportContext {
  /** Whether the family is installed on this machine (the panel passes a
   * `document.fonts.check` probe; tests pass a fake). */
  isInstalled: (family: string) => boolean;
  /** Lower-cased families the document embeds its own faces for (AC-11/13) —
   * those render from the file itself, so they count as available. */
  embeddedFamilies?: ReadonlySet<string>;
}

/** Classifies one family for the panel. */
export function fontStatusOf(family: string, ctx: FontReportContext): FontReportEntry {
  const lower = family.toLowerCase();
  const embedded = ctx.embeddedFamilies?.has(lower) ?? false;
  // Available: the document carries the face, the OS has it, or we bundle it.
  if (embedded || ctx.isInstalled(family) || isBundledFamily(family)) {
    return { family, status: "available" };
  }
  const sub = substituteFor(family);
  if (sub) {
    return { family, status: "substituted", substitute: sub.family };
  }
  // No twin and not installed: renders on the documented default (AC-9/10).
  return { family, status: "missing" };
}

/** The full Fonts-panel list, one entry per family the document uses,
 * alphabetical (documentFontFamilies already de-duplicates and sorts). */
export function documentFontReport(model: DocumentModel, ctx: FontReportContext): FontReportEntry[] {
  return documentFontFamilies(model).map((family) => fontStatusOf(family, ctx));
}
