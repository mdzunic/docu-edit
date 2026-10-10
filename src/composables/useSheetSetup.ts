// The assembly sheet's controller (SPEC-108): the state and commands behind
// SheetSetupPanel.vue, which every assembly surface embeds (the Convert to
// PDF dialog today; SPEC-109..111 next). The geometry is core/sheetLayout.ts
// and the rules (labels, refusals, the estimate) are lib/assemblySheet.ts —
// this composable holds a setup, validates each input against it, and shows
// lengths in the user's SPEC-064 measurement unit while the setup stays in
// centimetres. An input that fails is rejected with its reason; the setup is
// left exactly as it was, never silently clamped.
//
// Choices persist twice: for the session, per surface, in this module (so
// reopening a dialog finds the sheet as it was left); and, for the three
// that are really preferences — paper, margins, dpi — in lib/preferences.ts.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import { usePreferences } from "./usePreferences";
import { MARGIN_PRESET_NAMES, MARGIN_PRESETS } from "../lib/preferences";
import {
  cloneSetup,
  CUSTOM_BACKGROUND_START,
  initialSheetSetup,
  MARGIN_EDGE_KEYS,
  marginPresetOf,
  marginsAllEqual,
  SHEET_BACKGROUND_CHOICES,
  sheetGeometryFor,
  sheetOutputEstimate,
  sheetSetupRefusal,
  type MarginEdge,
} from "../lib/assemblySheet";
import type { Margins } from "../lib/core/pageSetup";
import { orientationOf } from "../lib/core/pageSetup";
import { formatLength, parseLengthToCm, type MeasurementUnit } from "../lib/core/units";
import {
  backgroundRefusal,
  isSheetBackground,
  MAX_SHEET_CM,
  paperSizeOf,
  placeOnSheet,
  resolveSheet,
  SHEET_ALIGNS,
  SHEET_DEFAULTS,
  SHEET_DPI_CHOICES,
  SHEET_FITS,
  SHEET_PAPER_NAMES,
  type ItemPx,
  type SheetAlign,
  type SheetFit,
  type SheetGeometry,
  type SheetOutput,
  type SheetPaper,
  type SheetPlacement,
  type SheetSetup,
} from "../lib/core/sheetLayout";

export { SHEET_ALIGN_LABELS, SHEET_BACKGROUND_CHOICES, SHEET_FIT_LABELS, type MarginEdge } from "../lib/assemblySheet";

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

export function useSheetSetup(deps: SheetSetupDeps): SheetSetupController {
  const { preferences, setPreference } = usePreferences();
  const unit = computed<MeasurementUnit>(() => preferences.measurementUnit);

  const saved = sessions.get(deps.surface);
  const start = saved
    ? { setup: cloneSetup(saved.setup), linked: saved.linked }
    : initialSheetSetup(
        {
          ...preferences,
          assemblyMargins: preferences.assemblyMargins ? { ...preferences.assemblyMargins } : null,
        },
        deps.output,
      );
  const setup = ref<SheetSetup>(start.setup);
  const marginsLinked = ref(start.linked);
  const error = ref("");
  const errorField = ref("");

  // Every committed change is the session's state from then on.
  watch(
    [setup, marginsLinked],
    () => sessions.set(deps.surface, { setup: cloneSetup(setup.value), linked: marginsLinked.value }),
    { deep: true, immediate: true, flush: "sync" },
  );

  const display = (cm: number): string => formatLength(cm, unit.value);
  const isCustom = computed(() => setup.value.paper === "Custom");
  const isFitToContent = computed(() => setup.value.paper === "Fit to content");
  const geometry = computed(() => sheetGeometryFor(setup.value, deps.itemPx.value));
  const placement = computed(() => {
    const sheet = geometry.value;
    const item = deps.itemPx.value;
    const { fit, align, dpi } = setup.value;
    return sheet && item ? placeOnSheet(item, sheet.content, fit, align, dpi) : null;
  });

  function reject(reason: string, field: string): boolean {
    error.value = reason;
    errorField.value = field;
    return false;
  }

  /** Commits `candidate` when it is valid; otherwise records why not. */
  function commit(candidate: SheetSetup, field: string): boolean {
    const reason = sheetSetupRefusal(candidate, deps.itemPx.value, unit.value);
    if (reason !== null) {
      return reject(reason, field);
    }
    error.value = "";
    errorField.value = "";
    setup.value = candidate;
    return true;
  }

  function parseLength(input: string, field: string): number | null {
    const cm = parseLengthToCm(input, unit.value);
    if (cm === null) {
      reject(`"${input.trim()}" is not a length. Type a number, e.g. 2 or 2${unit.value}.`, field);
    }
    return cm;
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
    // "Custom" is not a preference: its size belongs to this one sheet.
    if (ok && name !== "Custom") {
      setPreference("assemblyPaper", name);
    }
    return ok;
  }

  function setOrientation(orientation: string): boolean {
    if (orientation !== "portrait" && orientation !== "landscape") {
      return reject(`"${orientation}" is not an orientation.`, "orientation");
    }
    return commit({ ...cloneSetup(setup.value), orientation }, "orientation");
  }

  function setCustomSize(axis: "width" | "height", input: string): boolean {
    const cm = parseLength(input, axis);
    if (cm === null) {
      return false;
    }
    if (!(cm > 0) || cm > MAX_SHEET_CM) {
      return reject(`A sheet side must be more than 0 and at most ${display(MAX_SHEET_CM)} ${unit.value}.`, axis);
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
    return commit(candidate, axis);
  }

  function commitMargins(margins: Margins, field: string): boolean {
    const ok = commit({ ...cloneSetup(setup.value), margins }, field);
    if (ok) {
      setPreference("assemblyMargins", { ...margins });
    }
    return ok;
  }

  function setMarginPreset(name: string): boolean {
    const preset = MARGIN_PRESETS[name];
    // "Custom" is a reading of the four fields, not something to apply.
    if (!preset) {
      return false;
    }
    const ok = commitMargins({ ...preset }, "margin-preset");
    if (ok && !marginsAllEqual(preset)) {
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
      : { ...setup.value.margins, [MARGIN_EDGE_KEYS[edge]]: cm };
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
    return value ? commit({ ...cloneSetup(setup.value), fit: value }, "fit") : reject(`"${fit}" is not a fit.`, "fit");
  }

  function setAlign(align: string): boolean {
    const value = SHEET_ALIGNS.find((candidate) => candidate === align);
    return value
      ? commit({ ...cloneSetup(setup.value), align: value }, "align")
      : reject(`"${align}" is not an alignment.`, "align");
  }

  function setBackground(value: string): boolean {
    const next = value === "custom" ? CUSTOM_BACKGROUND_START : value.toLowerCase();
    if (!isSheetBackground(next)) {
      return reject(`"${value}" is not a colour. Use #rrggbb or transparent.`, "background");
    }
    // JPEG refuses transparency outright; PDF accepts the choice and says,
    // through backgroundNotice, that it will paint white.
    const refused = deps.output === "jpeg" ? backgroundRefusal(next, deps.output) : null;
    return refused !== null
      ? reject(refused, "background")
      : commit({ ...cloneSetup(setup.value), background: next }, "background");
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

  const backgroundChoice = computed(() => {
    const value = setup.value.background;
    return SHEET_BACKGROUND_CHOICES.some((choice) => choice.value === value) ? value : "custom";
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
    marginPreset: computed(() => marginPresetOf(setup.value.margins)),
    backgroundChoice,
    backgroundNotice: computed(() => backgroundRefusal(setup.value.background, deps.output) ?? ""),
    estimate: computed(() =>
      geometry.value ? sheetOutputEstimate(geometry.value, placement.value, setup.value.dpi, deps.output) : "",
    ),
    isCustom,
    isFitToContent,
    paperOptions: SHEET_PAPER_NAMES,
    fitOptions: SHEET_FITS,
    alignOptions: SHEET_ALIGNS,
    dpiChoices: SHEET_DPI_CHOICES,
    marginPresetOptions: [...MARGIN_PRESET_NAMES, "Custom"],
    display,
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
