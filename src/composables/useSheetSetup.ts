// The assembly sheet's controller (SPEC-108): the state and commands behind
// SheetSetupPanel.vue, which every assembly surface embeds (the Convert to
// PDF dialog today; SPEC-109..111 next). The geometry itself is
// core/sheetLayout.ts — this composable only holds a setup, validates each
// input against it, and formats lengths in the user's SPEC-064 measurement
// unit while the setup stays in centimetres.
//
// Choices persist twice: for the session, per surface, in this module (so
// reopening a dialog finds the sheet as it was left); and, for the three
// that are really preferences — paper, margins, dpi — in lib/preferences.ts.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import { usePreferences } from "./usePreferences";
import { assemblySheetDefaults, MARGIN_PRESET_NAMES, MARGIN_PRESETS } from "../lib/preferences";
import { formatFileSize } from "../lib/format";
import type { Margins, Orientation } from "../lib/core/pageSetup";
import { orientationOf } from "../lib/core/pageSetup";
import { CM_PER_INCH, formatLength, parseLengthToCm, type MeasurementUnit } from "../lib/core/units";
import {
  backgroundRefusal,
  defaultSheetBackground,
  FIT_TO_CONTENT,
  isSheetBackground,
  marginsLeaveContent,
  MAX_SHEET_CM,
  MIN_CONTENT_CM,
  paperSizeOf,
  placeOnSheet,
  resolveSheet,
  SHEET_ALIGNS,
  SHEET_DEFAULTS,
  SHEET_DPI_CHOICES,
  SHEET_FITS,
  SHEET_PAPER_NAMES,
  sheetPixelSize,
  sheetSizeAllowed,
  type ItemPx,
  type SheetAlign,
  type SheetFit,
  type SheetGeometry,
  type SheetOutput,
  type SheetPaper,
  type SheetPlacement,
  type SheetSetup,
} from "../lib/core/sheetLayout";

export type MarginEdge = "top" | "right" | "bottom" | "left";

export const MARGIN_EDGES: readonly MarginEdge[] = ["top", "right", "bottom", "left"];

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

export interface SheetSetupDeps {
  /** The session key: each assembly surface keeps its own choices. */
  surface: string;
  output: SheetOutput;
  /** The item's natural pixel size, once known ("Fit to content" and the
   * placement need it). */
  itemPx: Ref<ItemPx | null>;
}

export interface SheetSetupController {
  output: SheetOutput;
  setup: Ref<SheetSetup>;
  itemPx: Ref<ItemPx | null>;
  unit: ComputedRef<MeasurementUnit>;
  /** The resolved sheet, or null while "Fit to content" has no item size. */
  geometry: ComputedRef<SheetGeometry | null>;
  placement: ComputedRef<SheetPlacement | null>;
  /** The reason the last input was rejected ("" when it was accepted). */
  error: Ref<string>;
  /** Which control that rejected input came from ("" when none). */
  errorField: Ref<string>;
  marginsLinked: Ref<boolean>;
  /** The MARGIN_PRESETS name the margins match, else "Custom". */
  marginPreset: ComputedRef<string>;
  /** "custom" when the background is a colour without a named choice. */
  backgroundChoice: ComputedRef<string>;
  /** Shown before the user commits when the output cannot honour the
   * background ("transparent" on a PDF page). */
  backgroundNotice: ComputedRef<string>;
  /** The live output estimate: pixels at the chosen dpi and a byte bound. */
  estimate: ComputedRef<string>;
  isCustom: ComputedRef<boolean>;
  isFitToContent: ComputedRef<boolean>;
  paperOptions: readonly SheetPaper[];
  fitOptions: readonly SheetFit[];
  alignOptions: readonly SheetAlign[];
  dpiChoices: readonly number[];
  marginPresetOptions: readonly string[];
  /** A centimetre length in the active unit, without the suffix. */
  display: (cm: number) => string;
  /** "21 × 29.7 cm" in the active unit. */
  displaySize: (widthCm: number, heightCm: number) => string;
  setPaper: (paper: string) => boolean;
  setOrientation: (orientation: string) => boolean;
  setCustomSize: (axis: "width" | "height", input: string) => boolean;
  setMarginPreset: (name: string) => boolean;
  setMargin: (edge: MarginEdge, input: string) => boolean;
  setMarginsLinked: (linked: boolean) => boolean;
  setFit: (fit: string) => boolean;
  setAlign: (align: string) => boolean;
  setBackground: (value: string) => boolean;
  setDpi: (dpi: number | string) => boolean;
}

interface SessionState {
  setup: SheetSetup;
  linked: boolean;
}

const sessions = new Map<string, SessionState>();

/** Forgets every surface's session choices — a fresh session, and tests. */
export function resetSheetSetupSessions(): void {
  sessions.clear();
}

function cloneSetup(setup: SheetSetup): SheetSetup {
  return { ...setup, margins: { ...setup.margins } };
}

function allEqual(margins: Margins): boolean {
  return (
    margins.topCm === margins.rightCm &&
    margins.topCm === margins.bottomCm &&
    margins.topCm === margins.leftCm
  );
}

const EDGE_KEYS: Record<MarginEdge, keyof Margins> = {
  top: "topCm",
  right: "rightCm",
  bottom: "bottomCm",
  left: "leftCm",
};

export function useSheetSetup(deps: SheetSetupDeps): SheetSetupController {
  const { preferences, setPreference } = usePreferences();
  const unit = computed<MeasurementUnit>(() => preferences.measurementUnit);

  function initialState(): SessionState {
    const saved = sessions.get(deps.surface);
    if (saved) {
      return { setup: cloneSetup(saved.setup), linked: saved.linked };
    }
    const prefs = assemblySheetDefaults({ ...preferences, assemblyMargins: preferences.assemblyMargins ? { ...preferences.assemblyMargins } : null });
    const paper = SHEET_PAPER_NAMES.find((name) => name === prefs.paper && name !== "Custom") ?? SHEET_DEFAULTS.paper;
    const setup: SheetSetup = {
      ...cloneSetup(SHEET_DEFAULTS),
      paper,
      margins: prefs.margins,
      dpi: SHEET_DPI_CHOICES.includes(prefs.dpi) ? prefs.dpi : SHEET_DEFAULTS.dpi,
      background: defaultSheetBackground(deps.output),
    };
    // A remembered margin set the default paper cannot hold is not forced on
    // it: the sheet starts from the Normal preset instead.
    if (paper !== FIT_TO_CONTENT && !marginsLeaveContent(paperSizeOf(setup), setup.margins)) {
      setup.margins = { ...SHEET_DEFAULTS.margins };
    }
    return { setup, linked: allEqual(setup.margins) };
  }

  const start = initialState();
  const setup = ref<SheetSetup>(start.setup);
  const marginsLinked = ref(start.linked);
  const error = ref("");
  const errorField = ref("");

  // Every committed change is the session's state from then on.
  watch(
    [setup, marginsLinked],
    () => sessions.set(deps.surface, { setup: cloneSetup(setup.value), linked: marginsLinked.value }),
    { deep: true, immediate: true },
  );

  const display = (cm: number): string => formatLength(cm, unit.value);
  const displaySize = (widthCm: number, heightCm: number): string =>
    `${display(widthCm)} × ${display(heightCm)} ${unit.value}`;

  const isCustom = computed(() => setup.value.paper === "Custom");
  const isFitToContent = computed(() => setup.value.paper === FIT_TO_CONTENT);

  function geometryFor(candidate: SheetSetup): SheetGeometry | null {
    if (candidate.paper === FIT_TO_CONTENT && !deps.itemPx.value) {
      return null;
    }
    return resolveSheet(candidate, deps.itemPx.value ?? undefined);
  }

  const geometry = computed(() => geometryFor(setup.value));
  const placement = computed(() => {
    const sheet = geometry.value;
    const item = deps.itemPx.value;
    if (!sheet || !item) {
      return null;
    }
    return placeOnSheet(item, sheet.content, setup.value.fit, setup.value.align, setup.value.dpi);
  });

  /** Why `candidate` cannot be the setup, or null when it can. */
  function refusal(candidate: SheetSetup): string | null {
    const margins = candidate.margins;
    if (Object.values(margins).some((value) => value < 0)) {
      return "Margins cannot be negative.";
    }
    if (candidate.paper === FIT_TO_CONTENT) {
      const sheet = geometryFor(candidate);
      if (!sheet) {
        return `"${FIT_TO_CONTENT}" needs the image's size, which could not be read.`;
      }
      if (!sheetSizeAllowed(sheet.widthCm, sheet.heightCm)) {
        return `That sheet would be ${displaySize(sheet.widthCm, sheet.heightCm)}; a side can be at most ${display(MAX_SHEET_CM)} ${unit.value}.`;
      }
      return null;
    }
    const paper = paperSizeOf(candidate);
    if (!sheetSizeAllowed(paper.widthCm, paper.heightCm)) {
      return `A sheet side must be more than 0 and at most ${display(MAX_SHEET_CM)} ${unit.value}.`;
    }
    if (!marginsLeaveContent(paper, margins)) {
      const left = Math.max(0, paper.widthCm - margins.leftCm - margins.rightCm);
      const down = Math.max(0, paper.heightCm - margins.topCm - margins.bottomCm);
      return (
        `These margins leave ${displaySize(left, down)} for the content on a ` +
        `${displaySize(paper.widthCm, paper.heightCm)} sheet; it needs at least ` +
        `${display(MIN_CONTENT_CM)} ${unit.value} each way.`
      );
    }
    return null;
  }

  /** Commits `candidate` when it is valid; otherwise records why not and
   * leaves the setup exactly as it was — never a silent clamp. */
  function commit(candidate: SheetSetup, field: string): boolean {
    const reason = refusal(candidate);
    if (reason !== null) {
      error.value = reason;
      errorField.value = field;
      return false;
    }
    error.value = "";
    errorField.value = "";
    setup.value = candidate;
    return true;
  }

  function reject(reason: string, field: string): boolean {
    error.value = reason;
    errorField.value = field;
    return false;
  }

  function parseLength(input: string, field: string): number | null {
    const cm = parseLengthToCm(input, unit.value);
    if (cm === null) {
      reject(`"${input.trim()}" is not a length. Type a number, e.g. 2 or 2${unit.value}.`, field);
    }
    return cm;
  }

  function savePaperPreference(paper: SheetPaper): void {
    if (paper !== "Custom") {
      setPreference("assemblyPaper", paper);
    }
  }

  function setPaper(paper: string): boolean {
    const name = SHEET_PAPER_NAMES.find((candidate) => candidate === paper);
    if (!name) {
      return reject(`"${paper}" is not a paper size.`, "paper");
    }
    const candidate = cloneSetup(setup.value);
    if (name === "Custom" && candidate.paper !== "Custom") {
      // Custom starts from the sheet as it stands, so choosing it changes
      // nothing until a size is typed.
      const current = geometry.value ?? resolveSheet(SHEET_DEFAULTS);
      candidate.widthCm = current.widthCm;
      candidate.heightCm = current.heightCm;
      candidate.orientation = orientationOf(current);
    }
    candidate.paper = name;
    const ok = commit(candidate, "paper");
    if (ok) {
      savePaperPreference(name);
    }
    return ok;
  }

  function setOrientation(orientation: string): boolean {
    if (orientation !== "portrait" && orientation !== "landscape") {
      return reject(`"${orientation}" is not an orientation.`, "orientation");
    }
    const candidate = cloneSetup(setup.value);
    candidate.orientation = orientation as Orientation;
    return commit(candidate, "orientation");
  }

  function setCustomSize(axis: "width" | "height", input: string): boolean {
    const field = axis;
    const cm = parseLength(input, field);
    if (cm === null) {
      return false;
    }
    if (!(cm > 0) || cm > MAX_SHEET_CM) {
      return reject(`A sheet side must be more than 0 and at most ${display(MAX_SHEET_CM)} ${unit.value}.`, field);
    }
    const candidate = cloneSetup(setup.value);
    const current = paperSizeOf(candidate);
    const size = axis === "width" ? { widthCm: cm, heightCm: current.heightCm } : { widthCm: current.widthCm, heightCm: cm };
    candidate.paper = "Custom";
    candidate.widthCm = size.widthCm;
    candidate.heightCm = size.heightCm;
    // A typed size is the shape the user means: the orientation follows it.
    if (size.widthCm !== size.heightCm) {
      candidate.orientation = orientationOf(size);
    }
    return commit(candidate, field);
  }

  function commitMargins(margins: Margins, field: string): boolean {
    const candidate = cloneSetup(setup.value);
    candidate.margins = margins;
    const ok = commit(candidate, field);
    if (ok) {
      setPreference("assemblyMargins", { ...margins });
    }
    return ok;
  }

  function setMarginPreset(name: string): boolean {
    const preset = MARGIN_PRESETS[name];
    if (!preset) {
      // "Custom" is a reading of the four fields, not something to apply.
      return false;
    }
    const ok = commitMargins({ ...preset }, "margin-preset");
    if (ok && !allEqual(preset)) {
      marginsLinked.value = false;
    }
    return ok;
  }

  function setMargin(edge: MarginEdge, input: string): boolean {
    const field = `margin-${edge}`;
    const cm = parseLength(input, field);
    if (cm === null) {
      return false;
    }
    if (cm < 0) {
      return reject("Margins cannot be negative.", field);
    }
    const margins = marginsLinked.value
      ? { topCm: cm, rightCm: cm, bottomCm: cm, leftCm: cm }
      : { ...setup.value.margins, [EDGE_KEYS[edge]]: cm };
    return commitMargins(margins, field);
  }

  function setMarginsLinked(linked: boolean): boolean {
    if (!linked) {
      marginsLinked.value = false;
      return true;
    }
    // Linking makes all four equal to the top margin.
    const top = setup.value.margins.topCm;
    const ok = commitMargins({ topCm: top, rightCm: top, bottomCm: top, leftCm: top }, "margin-link");
    if (ok) {
      marginsLinked.value = true;
    }
    return ok;
  }

  function setFit(fit: string): boolean {
    const value = SHEET_FITS.find((candidate) => candidate === fit);
    if (!value) {
      return reject(`"${fit}" is not a fit.`, "fit");
    }
    return commit({ ...cloneSetup(setup.value), fit: value }, "fit");
  }

  function setAlign(align: string): boolean {
    const value = SHEET_ALIGNS.find((candidate) => candidate === align);
    if (!value) {
      return reject(`"${align}" is not an alignment.`, "align");
    }
    return commit({ ...cloneSetup(setup.value), align: value }, "align");
  }

  function setBackground(value: string): boolean {
    const next = value === "custom" ? "#808080" : value.toLowerCase();
    if (!isSheetBackground(next)) {
      return reject(`"${value}" is not a colour. Use #rrggbb or transparent.`, "background");
    }
    // JPEG refuses transparency outright; PDF accepts the choice and says,
    // through backgroundNotice, that it will paint white.
    const refusedOutright = deps.output === "jpeg" ? backgroundRefusal(next, deps.output) : null;
    if (refusedOutright !== null) {
      return reject(refusedOutright, "background");
    }
    return commit({ ...cloneSetup(setup.value), background: next }, "background");
  }

  function setDpi(dpi: number | string): boolean {
    const value = Number(dpi);
    if (!SHEET_DPI_CHOICES.includes(value)) {
      return reject(`Resolution must be one of ${SHEET_DPI_CHOICES.join(", ")} dpi.`, "dpi");
    }
    const ok = commit({ ...cloneSetup(setup.value), dpi: value }, "dpi");
    if (ok) {
      setPreference("assemblyDpi", value);
    }
    return ok;
  }

  const marginPreset = computed(() => {
    const margins = setup.value.margins;
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
  });

  const backgroundChoice = computed(() => {
    const value = setup.value.background;
    return SHEET_BACKGROUND_CHOICES.some((choice) => choice.value === value) ? value : "custom";
  });

  const backgroundNotice = computed(() => backgroundRefusal(setup.value.background, deps.output) ?? "");

  const estimate = computed(() => {
    const sheet = geometry.value;
    const placed = placement.value;
    if (!sheet) {
      return "";
    }
    const dpi = setup.value.dpi;
    const pixels = sheetPixelSize(sheet, dpi);
    // A PDF carries only the placed image's samples — never more than the
    // source has, never more than the placed size at this dpi (the encoder's
    // own ceiling); a raster output is the whole sheet. Either way the bound
    // is four uncompressed bytes a sample plus the file's own overhead.
    let samples = pixels.widthPx * pixels.heightPx;
    if (deps.output === "pdf") {
      samples = 0;
      if (placed) {
        const across = Math.min(Math.ceil(placed.crop.widthPx), Math.round((placed.widthCm / CM_PER_INCH) * dpi));
        const down = Math.min(Math.ceil(placed.crop.heightPx), Math.round((placed.heightCm / CM_PER_INCH) * dpi));
        samples = Math.max(1, across) * Math.max(1, down);
      }
    }
    const bytes = samples * 4 + 2048;
    return `${pixels.widthPx} × ${pixels.heightPx} px at ${dpi} dpi · up to about ${formatFileSize(bytes)}`;
  });

  return {
    output: deps.output,
    setup,
    itemPx: deps.itemPx,
    unit,
    geometry,
    placement,
    error,
    errorField,
    marginsLinked,
    marginPreset,
    backgroundChoice,
    backgroundNotice,
    estimate,
    isCustom,
    isFitToContent,
    paperOptions: SHEET_PAPER_NAMES,
    fitOptions: SHEET_FITS,
    alignOptions: SHEET_ALIGNS,
    dpiChoices: SHEET_DPI_CHOICES,
    marginPresetOptions: [...MARGIN_PRESET_NAMES, "Custom"],
    display,
    displaySize,
    setPaper,
    setOrientation,
    setCustomSize,
    setMarginPreset,
    setMargin,
    setMarginsLinked,
    setFit,
    setAlign,
    setBackground,
    setDpi,
  };
}
