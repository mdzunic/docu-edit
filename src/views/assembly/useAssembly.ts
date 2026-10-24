// The assembly workspace's controller (SPEC-109) — everything /assemble
// does, for AssembleView.vue to wire: the tray (useAssemblyTray), the layout
// settings and SPEC-108's sheet setup (useSheetSetup, whose fit and align
// apply per cell), the planned placements (core/assemblyLayout's
// planAssembly — the one source of "where does each item go"), overlay
// editing (inspector fields, preview drags, arrow-key nudges), the output
// (useAssemblyOutput; SPEC-110's PDF pages through usePdfAssembly) and the
// plan's persistence (useAssemblyPersistence).
//
// The tray order is the layout order: `items` is the tray, mapped. For a PDF
// the preview's placements are the current page's.
import { computed, reactive, ref, type ComputedRef, type Ref } from "vue";
import { useSheetSetup, type SheetSetupController } from "../../composables/useSheetSetup";
import {
  BLEND_MODES,
  LAYOUT_KINDS,
  MAX_ASSEMBLY_ITEMS,
  MAX_GAP_CM,
  naturalAssemblyPx,
  naturalColumns,
  planAssembly,
  type AssemblyItem,
  type AssemblyPlan,
  type BlendMode,
  type LayoutKind,
  type OverlayTransform,
  type PlacedItem,
} from "../../lib/core/assemblyLayout";
import { isSheetBackground, type SheetGeometry } from "../../lib/core/sheetLayout";
import { formatLength, parseLengthToCm } from "../../lib/core/units";
import { outputBackground } from "../../lib/assembly";
import {
  ASSEMBLY_FORMATS,
  ASSEMBLY_LAYOUT_DEFAULTS,
  ASSEMBLY_OUTPUT_DEFAULTS,
  type AssemblyFormat,
  type AssemblyLayoutSettings,
  type AssemblyOutputSettings,
} from "../../lib/assemblyPersist";
import { useAssemblyTray, type AssemblyTrayController, type TrayEntry } from "./useAssemblyTray";
import { BROWSER_ASSEMBLY_STEPS, useAssemblyOutput, type AssemblyOutputController, type AssemblySteps } from "./useAssemblyOutput";
import { useAssemblyPersistence } from "./useAssemblyPersistence";
import { pdfNaturalPx } from "../../lib/media/imagesToPdf";
import {
  BROWSER_PDF_ASSEMBLY_STEPS,
  usePdfAssembly,
  type PdfAssemblyController,
  type PdfAssemblySteps,
} from "../assemble/usePdfAssembly";

/** The sheet setup's session key for this surface. */
export const ASSEMBLY_SURFACE = "assemble";

/** An arrow key's nudge, and Shift+arrow's. */
export const NUDGE_CM = 0.1;
export const NUDGE_SHIFT_CM = 1;

/** The column choices the inspector offers (any whole number up to the
 * tray limit is accepted). */
export const COLUMN_CHOICES: readonly (number | "auto")[] = ["auto", 1, 2, 3, 4, 5, 6, 8, 10];

export interface AssemblyDeps {
  /** Test seam: the browser-bound output steps. */
  steps?: Partial<AssemblySteps>;
  /** Test seam: the PDF output's steps. */
  pdfSteps?: Partial<PdfAssemblySteps>;
}

export interface AssemblyController {
  tray: AssemblyTrayController;
  sheet: SheetSetupController;
  layout: AssemblyLayoutSettings;
  output: AssemblyOutputSettings;
  out: AssemblyOutputController;
  /** SPEC-110: the PDF's pages and output. */
  pdf: PdfAssemblyController;
  isPdf: ComputedRef<boolean>;
  items: ComputedRef<AssemblyItem[]>;
  plan: ComputedRef<AssemblyPlan>;
  placed: ComputedRef<PlacedItem[]>;
  geometry: ComputedRef<SheetGeometry | null>;
  /** The colour painted behind everything (JPEG's matte for transparency). */
  background: ComputedRef<string>;
  isOverlay: ComputedRef<boolean>;
  /** The selected entry, when the layout is overlay — what the inspector's
   * transform fields edit. */
  overlayItem: ComputedRef<TrayEntry | null>;
  /** The restored-tray notice ("" when none). */
  notice: Ref<string>;
  /** Why the last inspector input was rejected ("" when accepted). */
  error: Ref<string>;
  /** The gap in the user's unit, without the suffix. */
  gapDisplay: ComputedRef<string>;
  setLayout: (kind: string) => boolean;
  setColumns: (value: string) => boolean;
  setGap: (input: string) => boolean;
  setCellBackground: (value: string) => boolean;
  setFormat: (value: string) => boolean;
  setQuality: (value: string) => boolean;
  setTitle: (value: string) => void;
  setMatte: (value: string) => boolean;
  setTransformField: (field: keyof OverlayTransform, value: string) => boolean;
  /** Moves an overlay entry by a distance on the sheet. */
  nudge: (key: string, dxCm: number, dyCm: number) => void;
  /** Places an overlay entry at `start` moved by a distance (a drag). */
  dragTo: (key: string, start: OverlayTransform, dxCm: number, dyCm: number) => void;
  select: (key: string) => void;
  reset: () => void;
  flush: () => void;
}

export function useAssembly(deps: AssemblyDeps = {}): AssemblyController {
  const tray = useAssemblyTray();
  const layout = reactive<AssemblyLayoutSettings>({ ...ASSEMBLY_LAYOUT_DEFAULTS });
  const output = reactive<AssemblyOutputSettings>({ ...ASSEMBLY_OUTPUT_DEFAULTS });
  const error = ref("");

  const isPdf = computed(() => output.format === "pdf");
  // A PDF lays out by its own grid unless every image is on one page.
  const isOverlay = computed(() => layout.layout === "overlay" && (!isPdf.value || output.pageMode === "single-page"));
  const items = computed<AssemblyItem[]>(() =>
    tray.entries.value.map((entry) => ({
      id: entry.id,
      widthPx: entry.widthPx,
      heightPx: entry.heightPx,
      ...(isOverlay.value ? { transform: entry.transform } : {}),
    })),
  );
  // "Fit to content" sizes the sheet to the arrangement's own pixels — for a
  // PDF, one page's arrangement.
  const naturalPx = computed(() => {
    const natural: AssemblyPlan = { ...layout, fit: "contain", align: "center" };
    return isPdf.value
      ? pdfNaturalPx(items.value, { ...natural, pageMode: output.pageMode, nUp: output.nUp })
      : naturalAssemblyPx(items.value, natural);
  });
  const sheet = useSheetSetup({ surface: ASSEMBLY_SURFACE, output: "png", itemPx: naturalPx });

  const plan = computed<AssemblyPlan>(() => ({
    ...layout,
    // A sheet sized to the content was sized for naturalColumns' count:
    // "auto" deciding again on that box could choose another.
    columns: sheet.isFitToContent.value ? naturalColumns(layout.columns, items.value.length) : layout.columns,
    fit: sheet.setup.value.fit,
    align: sheet.setup.value.align,
  }));
  const geometry = computed(() => sheet.geometry.value);
  const pdf = usePdfAssembly({
    entries: tray.entries,
    items,
    plan,
    sheet,
    output,
    selectedKey: tray.selectedKey,
    steps: { ...BROWSER_PDF_ASSEMBLY_STEPS, ...deps.pdfSteps },
  });
  const placed = computed(() => {
    if (isPdf.value) {
      return pdf.currentPlaced.value;
    }
    return geometry.value ? planAssembly(items.value, plan.value, geometry.value, sheet.setup.value.dpi) : [];
  });
  const background = computed(() => outputBackground(sheet.setup.value.background, output.format, output.matte));
  const overlayItem = computed(() => (isOverlay.value ? tray.selected.value : null));

  const persistence = useAssemblyPersistence({ tray, layout, output, sheet });
  const out = useAssemblyOutput({
    entries: tray.entries,
    placed,
    sheet,
    output,
    background,
    steps: { ...BROWSER_ASSEMBLY_STEPS, ...deps.steps },
  });

  function reject(reason: string): boolean {
    error.value = reason;
    return false;
  }

  function accept(): boolean {
    error.value = "";
    return true;
  }

  function setLayout(kind: string): boolean {
    const value = LAYOUT_KINDS.find((candidate) => candidate === kind);
    if (!value) {
      return reject(`"${kind}" is not a layout.`);
    }
    layout.layout = value as LayoutKind;
    return accept();
  }

  function setColumns(value: string): boolean {
    if (value === "auto") {
      layout.columns = "auto";
      return accept();
    }
    const columns = Number(value);
    if (!Number.isInteger(columns) || columns < 1 || columns > MAX_ASSEMBLY_ITEMS) {
      return reject(`Columns must be "auto" or a whole number from 1 to ${MAX_ASSEMBLY_ITEMS}.`);
    }
    layout.columns = columns;
    return accept();
  }

  const unit = sheet.unit;
  const gapDisplay = computed(() => formatLength(layout.gapCm, unit.value));

  function setGap(input: string): boolean {
    const cm = parseLengthToCm(input, unit.value);
    if (cm === null) {
      return reject(`"${input.trim()}" is not a length. Type a number, e.g. 0.2 or 2mm.`);
    }
    if (cm < 0 || cm > MAX_GAP_CM) {
      return reject(`The gap must be from 0 to ${formatLength(MAX_GAP_CM, unit.value)} ${unit.value}.`);
    }
    layout.gapCm = cm;
    return accept();
  }

  function setCellBackground(value: string): boolean {
    const next = value === "custom" ? "#808080" : value.toLowerCase();
    if (!isSheetBackground(next)) {
      return reject(`"${value}" is not a colour. Use #rrggbb or transparent.`);
    }
    layout.cellBackground = next;
    return accept();
  }

  function setFormat(value: string): boolean {
    const format = ASSEMBLY_FORMATS.find((candidate) => candidate === value);
    if (!format) {
      return reject(`"${value}" is not an output format.`);
    }
    output.format = format as AssemblyFormat;
    return accept();
  }

  function setQuality(value: string): boolean {
    const quality = Number(value);
    if (!Number.isInteger(quality) || quality < 1 || quality > 100) {
      return reject("JPEG quality must be a whole number from 1 to 100.");
    }
    output.quality = quality;
    return accept();
  }

  function setMatte(value: string): boolean {
    if (!/^#[0-9a-f]{6}$/i.test(value)) {
      return reject(`"${value}" is not a colour.`);
    }
    output.matte = value.toLowerCase();
    return accept();
  }

  function setTransformField(field: keyof OverlayTransform, value: string): boolean {
    const entry = overlayItem.value;
    if (!entry) {
      return false;
    }
    if (field === "blend") {
      const blend = BLEND_MODES.find((candidate) => candidate === value);
      if (!blend) {
        return reject(`"${value}" is not a blend mode.`);
      }
      tray.setTransform(entry.key, { ...entry.transform, blend: blend as BlendMode });
      return accept();
    }
    const number = Number(value);
    if (value.trim() === "" || !Number.isFinite(number)) {
      return reject(`"${value.trim()}" is not a number.`);
    }
    tray.setTransform(entry.key, { ...entry.transform, [field]: number });
    return accept();
  }

  function dragTo(key: string, start: OverlayTransform, dxCm: number, dyCm: number): void {
    const box = geometry.value?.content;
    if (!box || box.widthCm <= 0 || box.heightCm <= 0) {
      return;
    }
    tray.setTransform(key, {
      ...start,
      xPercent: start.xPercent + (dxCm / box.widthCm) * 100,
      yPercent: start.yPercent + (dyCm / box.heightCm) * 100,
    });
  }

  function nudge(key: string, dxCm: number, dyCm: number): void {
    const entry = tray.entries.value.find((candidate) => candidate.key === key);
    if (entry) {
      dragTo(key, entry.transform, dxCm, dyCm);
    }
  }

  return {
    tray,
    sheet,
    layout,
    output,
    out,
    pdf,
    isPdf,
    items,
    plan,
    placed,
    geometry,
    background,
    isOverlay,
    overlayItem,
    notice: persistence.notice,
    error,
    gapDisplay,
    setLayout,
    setColumns,
    setGap,
    setCellBackground,
    setFormat,
    setQuality,
    setTitle: (value: string) => {
      output.title = value;
    },
    setMatte,
    setTransformField,
    nudge,
    dragTo,
    select: tray.select,
    reset: () => {
      persistence.reset();
      error.value = "";
    },
    flush: persistence.flush,
  };
}
