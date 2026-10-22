// The assembly sheet's setup rules (SPEC-108) — the framework-free half of
// the useSheetSetup controller: the panel's labels and choices, the starting
// setup a surface opens with, why a candidate setup is refused (worded in the
// user's measurement unit), which margin preset a margin set is, and the live
// output estimate. The geometry itself is core/sheetLayout.ts; nothing here
// imports Vue, so every rule is unit-testable on its own.
import { MARGIN_PRESET_NAMES, MARGIN_PRESETS, type Preferences, assemblySheetDefaults } from "./preferences";
import { formatFileSize } from "./format";
import type { Margins } from "./core/pageSetup";
import { CM_PER_INCH, formatLength, type MeasurementUnit } from "./core/units";
import {
  defaultSheetBackground,
  FIT_TO_CONTENT,
  marginsLeaveContent,
  MAX_SHEET_CM,
  MIN_CONTENT_CM,
  paperSizeOf,
  resolveSheet,
  SHEET_DEFAULTS,
  SHEET_DPI_CHOICES,
  SHEET_PAPER_NAMES,
  sheetPixelSize,
  sheetSizeAllowed,
  type ItemPx,
  type SheetGeometry,
  type SheetOutput,
  type SheetPlacement,
  type SheetSetup,
} from "./core/sheetLayout";

export type MarginEdge = "top" | "right" | "bottom" | "left";

/** Each margin field's key in `Margins`. */
export const MARGIN_EDGE_KEYS: Record<MarginEdge, keyof Margins> = {
  top: "topCm",
  right: "rightCm",
  bottom: "bottomCm",
  left: "leftCm",
};

/** The fit select's labels. */
export const SHEET_FIT_LABELS: Readonly<Record<string, string>> = {
  contain: "Fit inside (contain)",
  cover: "Fill and crop (cover)",
  stretch: "Stretch to fill",
  actual: "Actual size",
};

/** The alignment select's labels. */
export const SHEET_ALIGN_LABELS: Readonly<Record<string, string>> = {
  "top-left": "Top left",
  top: "Top",
  "top-right": "Top right",
  left: "Left",
  center: "Center",
  right: "Right",
  "bottom-left": "Bottom left",
  bottom: "Bottom",
  "bottom-right": "Bottom right",
};

/** The background select's named choices; any other "#rrggbb" reads as
 * "custom" and is edited through the colour input beside it. */
export const SHEET_BACKGROUND_CHOICES = [
  { value: "#ffffff", label: "White" },
  { value: "#000000", label: "Black" },
  { value: "transparent", label: "Transparent" },
  { value: "custom", label: "Custom colour" },
];

/** The colour "Custom colour" starts from. */
export const CUSTOM_BACKGROUND_START = "#808080";

export function cloneSetup(setup: SheetSetup): SheetSetup {
  return { ...setup, margins: { ...setup.margins } };
}

export function marginsAllEqual(margins: Margins): boolean {
  return (
    margins.topCm === margins.rightCm &&
    margins.topCm === margins.bottomCm &&
    margins.topCm === margins.leftCm
  );
}

/** The setup a surface opens with when the session has none: the user's
 * assembly preferences (else their page-setup defaults), the output's
 * default background, and Normal margins if the remembered ones do not fit
 * the remembered paper. Linked when all four margins are equal. */
export function initialSheetSetup(prefs: Preferences, output: SheetOutput): { setup: SheetSetup; linked: boolean } {
  const chosen = assemblySheetDefaults(prefs);
  const paper = SHEET_PAPER_NAMES.find((name) => name === chosen.paper && name !== "Custom") ?? SHEET_DEFAULTS.paper;
  const setup: SheetSetup = {
    ...cloneSetup(SHEET_DEFAULTS),
    paper,
    margins: chosen.margins,
    dpi: SHEET_DPI_CHOICES.includes(chosen.dpi) ? chosen.dpi : SHEET_DEFAULTS.dpi,
    background: defaultSheetBackground(output),
  };
  if (paper !== FIT_TO_CONTENT && !marginsLeaveContent(paperSizeOf(setup), setup.margins)) {
    setup.margins = { ...SHEET_DEFAULTS.margins };
  }
  return { setup, linked: marginsAllEqual(setup.margins) };
}

/** The sheet for a setup, or null while "Fit to content" has no item size. */
export function sheetGeometryFor(setup: SheetSetup, itemPx: ItemPx | null): SheetGeometry | null {
  if (setup.paper === FIT_TO_CONTENT && !itemPx) {
    return null;
  }
  return resolveSheet(setup, itemPx ?? undefined);
}

/** "21 × 29.7 cm" in `unit`. */
export function sizeInUnit(widthCm: number, heightCm: number, unit: MeasurementUnit): string {
  return `${formatLength(widthCm, unit)} × ${formatLength(heightCm, unit)} ${unit}`;
}

/** Why `candidate` cannot be the setup — worded in `unit` — or null when it
 * can. The panel shows this beside the input and keeps the setup as it was. */
export function sheetSetupRefusal(candidate: SheetSetup, itemPx: ItemPx | null, unit: MeasurementUnit): string | null {
  const margins = candidate.margins;
  const maxSide = `${formatLength(MAX_SHEET_CM, unit)} ${unit}`;
  if (Object.values(margins).some((value) => value < 0)) {
    return "Margins cannot be negative.";
  }
  if (candidate.paper === FIT_TO_CONTENT) {
    const sheet = sheetGeometryFor(candidate, itemPx);
    if (!sheet) {
      return `"${FIT_TO_CONTENT}" needs the image's size, which could not be read.`;
    }
    if (!sheetSizeAllowed(sheet.widthCm, sheet.heightCm)) {
      return `That sheet would be ${sizeInUnit(sheet.widthCm, sheet.heightCm, unit)}; a side can be at most ${maxSide}.`;
    }
    return null;
  }
  const paper = paperSizeOf(candidate);
  if (!sheetSizeAllowed(paper.widthCm, paper.heightCm)) {
    return `A sheet side must be more than 0 and at most ${maxSide}.`;
  }
  if (!marginsLeaveContent(paper, margins)) {
    const across = Math.max(0, paper.widthCm - margins.leftCm - margins.rightCm);
    const down = Math.max(0, paper.heightCm - margins.topCm - margins.bottomCm);
    return (
      `These margins leave ${sizeInUnit(across, down, unit)} for the content on a ` +
      `${sizeInUnit(paper.widthCm, paper.heightCm, unit)} sheet; it needs at least ` +
      `${formatLength(MIN_CONTENT_CM, unit)} ${unit} each way.`
    );
  }
  return null;
}

/** The MARGIN_PRESETS name `margins` match exactly, else "Custom". */
export function marginPresetOf(margins: Margins): string {
  const match = MARGIN_PRESET_NAMES.find((name) => {
    const preset = MARGIN_PRESETS[name];
    return (
      preset.topCm === margins.topCm &&
      preset.rightCm === margins.rightCm &&
      preset.bottomCm === margins.bottomCm &&
      preset.leftCm === margins.leftCm
    );
  });
  return match ?? "Custom";
}

/** The live output estimate: the sheet's pixels at `dpi`, and an upper bound
 * on the bytes. A PDF carries only the placed image's samples — never more
 * than the source has, never more than the placed size at this dpi (the
 * encoder's own ceiling); a raster output is the whole sheet. Either way the
 * bound is four uncompressed bytes a sample plus the file's own overhead. */
export function sheetOutputEstimate(
  geometry: SheetGeometry,
  placement: SheetPlacement | null,
  dpi: number,
  output: SheetOutput,
): string {
  const pixels = sheetPixelSize(geometry, dpi);
  let samples = pixels.widthPx * pixels.heightPx;
  if (output === "pdf") {
    samples = 0;
    if (placement) {
      const across = Math.min(Math.ceil(placement.crop.widthPx), Math.round((placement.widthCm / CM_PER_INCH) * dpi));
      const down = Math.min(Math.ceil(placement.crop.heightPx), Math.round((placement.heightCm / CM_PER_INCH) * dpi));
      samples = Math.max(1, across) * Math.max(1, down);
    }
  }
  const bytes = samples * 4 + 2048;
  return `${pixels.widthPx} × ${pixels.heightPx} px at ${dpi} dpi · up to about ${formatFileSize(bytes)}`;
}
