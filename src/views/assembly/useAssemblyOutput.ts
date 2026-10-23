// The assembly's output (SPEC-109): one PNG or JPEG of the whole sheet at
// SPEC-108's sheetPixelSize, saved to the library as a new media document or
// downloaded. The plan is refused BEFORE any work when the output breaks a
// bound — MAX_OUTPUT_EDGE_PX on either side, or the canvas ceiling this
// browser measurably allocates (probed once per session, and only for an
// area the guaranteed floor does not already cover) — with the limit named
// and the dpi that would fit offered.
//
// Producing never blocks the page: sources come from the session's decode
// cache, the drawing is media/composeCanvas's — the very function the
// preview calls — chunked by row with a yield between chunks, and Cancel
// stops at the next boundary with no file and no library record. The output
// canvas is released (width = 0) whatever happens. Source documents are only
// ever read: the result is always a NEW document.
import { computed, onScopeDispose, ref, type ComputedRef, type Ref } from "vue";
import { listDocuments, getMedia, storeMedia, uniqueTitle } from "../../lib/documentStore";
import { assembleMediaModel } from "../../lib/media/open";
import type { MediaModel } from "../../lib/media/model";
import { currentImagePart } from "../../lib/media/editModel";
import { canvasCeilingFor, GUARANTEED_CANVAS_PIXELS } from "../../lib/media/canvasCeiling";
import { releaseCanvas, type DecodedImage } from "../../lib/media/renderEdits";
import {
  composeToCanvas,
  decodeSource,
  encodeCanvas,
  prepareSources,
  type ComposeInput,
} from "../../lib/media/composeCanvas";
import type { PlacedItem } from "../../lib/core/assemblyLayout";
import type { PixelSize } from "../../lib/core/imageEdits";
import type { MediaPart } from "../../lib/core/media";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { sheetPixelSize } from "../../lib/core/sheetLayout";
import { sheetGeometryFor } from "../../lib/assemblySheet";
import {
  ASSEMBLY_CONTENT_TYPES,
  fittingDpi,
  jpegBackgroundNotice,
  outputEstimate,
  outputFileName,
  outputRefusal,
  type OutputPixels,
} from "../../lib/assembly";
import { DEFAULT_ASSEMBLY_TITLE, type AssemblyOutputSettings } from "../../lib/assemblyPersist";
import type { SheetSetupController } from "../../composables/useSheetSetup";
import type { TrayEntry } from "./useAssemblyTray";

/** The browser-bound steps, injectable so the pipeline is unit-testable. */
export interface AssemblySteps {
  createCanvas: (widthPx: number, heightPx: number) => HTMLCanvasElement;
  decode: (part: MediaPart, placedPx?: PixelSize) => Promise<DecodedImage>;
  encode: (canvas: HTMLCanvasElement, contentType: string, quality?: number) => Promise<Uint8Array>;
  /** The canvas ceiling to hold an area against (probes only when needed). */
  canvasCeiling: (areaPx: number) => number;
  /** Awaited between chunks of work. */
  between: () => Promise<void>;
  download: (bytes: Uint8Array, fileName: string, contentType: string) => void;
  release: (canvas: HTMLCanvasElement) => void;
}

export const BROWSER_ASSEMBLY_STEPS: AssemblySteps = {
  createCanvas: (widthPx, heightPx) => {
    const canvas = document.createElement("canvas");
    canvas.width = widthPx;
    canvas.height = heightPx;
    return canvas;
  },
  decode: decodeSource,
  encode: encodeCanvas,
  canvasCeiling: (areaPx) => canvasCeilingFor(areaPx),
  between: yieldToMainThread,
  download: triggerBlobDownload,
  release: releaseCanvas,
};

/** Typed pass-through for steps arriving through a `.vue` prop (runtime-
 * typed `Object`; espree bars the annotation in `.vue` scripts). */
export function asAssemblySteps(value: unknown): Partial<AssemblySteps> | undefined {
  return value ? (value as Partial<AssemblySteps>) : undefined;
}

/** How each tray entry is drawn: its current image's part (an edited
 * image's saved render) and the pixel size its placement is measured in. */
export function composeInputs(entries: readonly TrayEntry[]): (ComposeInput | null)[] {
  return entries.map((entry) => {
    const model = getMedia(entry.id);
    const part = model ? currentImagePart(model) : undefined;
    return part && entry.widthPx > 0 && entry.heightPx > 0 ? { part, widthPx: entry.widthPx, heightPx: entry.heightPx } : null;
  });
}

export interface AssemblyOutputDeps {
  entries: Ref<TrayEntry[]>;
  placed: ComputedRef<PlacedItem[]>;
  sheet: SheetSetupController;
  output: AssemblyOutputSettings;
  /** The colour actually painted behind everything. */
  background: ComputedRef<string>;
  steps: AssemblySteps;
}

export interface AssemblyOutputController {
  /** The output's pixel size, or null while the sheet is unresolved. */
  pixels: ComputedRef<OutputPixels | null>;
  /** The live estimate line. */
  estimate: ComputedRef<string>;
  /** Why the output cannot be produced ("" when it can). */
  refusal: ComputedRef<string>;
  /** The dpi a size refusal offers instead (null when none). */
  fixDpi: ComputedRef<number | null>;
  /** JPEG's notice for a transparent sheet ("" when none). */
  backgroundNotice: ComputedRef<string>;
  busy: Ref<boolean>;
  /** "Item {done} of {total}" while producing ("" otherwise). */
  progress: Ref<string>;
  /** The last failure ("" when none). */
  failure: Ref<string>;
  canProduce: ComputedRef<boolean>;
  save: () => Promise<MediaModel | null>;
  download: () => Promise<boolean>;
  cancel: () => void;
}

export function useAssemblyOutput(deps: AssemblyOutputDeps): AssemblyOutputController {
  const { sheet, output, steps } = deps;
  const busy = ref(false);
  const progress = ref("");
  const failure = ref("");
  let cancelled = false;

  const pixels = computed(() => {
    const geometry = sheet.geometry.value;
    return geometry ? sheetPixelSize(geometry, sheet.setup.value.dpi) : null;
  });
  const area = computed(() => (pixels.value ? pixels.value.widthPx * pixels.value.heightPx : 0));
  const ceiling = computed(() => steps.canvasCeiling(area.value));

  // The sheet at another dpi — for "Fit to content" a different sheet.
  const geometryAt = (dpi: number) => sheetGeometryFor({ ...sheet.setup.value, dpi }, sheet.itemPx.value);

  const refusal = computed(() => {
    const geometry = sheet.geometry.value;
    if (deps.entries.value.length === 0) {
      return "Add at least one image to the tray.";
    }
    const unreadable = deps.entries.value.filter((entry) => entry.widthPx <= 0 || entry.heightPx <= 0);
    if (unreadable.length > 0) {
      const names = unreadable.map((entry) => `“${entry.title}”`).join(", ");
      return `The size of ${names} could not be read; remove ${unreadable.length === 1 ? "it" : "them"} from the tray.`;
    }
    if (!geometry) {
      return "The sheet is not ready.";
    }
    return outputRefusal(geometry, sheet.setup.value.dpi, ceiling.value, geometryAt) ?? "";
  });
  const fixDpi = computed(() => {
    const geometry = sheet.geometry.value;
    return geometry && refusal.value && outputRefusal(geometry, sheet.setup.value.dpi, ceiling.value) !== null
      ? fittingDpi(geometryAt, ceiling.value)
      : null;
  });
  const estimate = computed(() =>
    pixels.value
      ? outputEstimate(
          pixels.value,
          sheet.setup.value.dpi,
          output.format,
          output.quality,
          area.value > GUARANTEED_CANVAS_PIXELS ? ceiling.value : null,
        )
      : "",
  );
  const backgroundNotice = computed(() => jpegBackgroundNotice(sheet.setup.value.background, output.format));
  const canProduce = computed(() => !busy.value && refusal.value === "");

  /** The encoded output, or null when refused or cancelled. Throws with a
   * user-facing reason when the browser cannot produce it. */
  async function produce(): Promise<Uint8Array | null> {
    const geometry = sheet.geometry.value;
    if (refusal.value || !geometry) {
      failure.value = refusal.value;
      return null;
    }
    const dpi = sheet.setup.value.dpi;
    const size = sheetPixelSize(geometry, dpi);
    // A snapshot: edits made while producing do not reach this file.
    const placed = deps.placed.value;
    const entries = [...deps.entries.value];
    const background = deps.background.value;
    const contentType = ASSEMBLY_CONTENT_TYPES[output.format];
    const quality = output.format === "jpeg" ? output.quality / 100 : undefined;
    const stop = () => cancelled;
    cancelled = false;
    busy.value = true;
    progress.value = `Item 0 of ${placed.length}`;
    let canvas: HTMLCanvasElement | null = null;
    try {
      const sx = size.widthPx / geometry.widthCm;
      const sy = size.heightPx / geometry.heightCm;
      const inputs = composeInputs(entries);
      const missing = entries.filter((_, index) => inputs[index] === null);
      if (missing.length > 0) {
        throw new Error(`${missing.map((entry) => `“${entry.title}”`).join(", ")} could not be read.`);
      }
      const prepared = await prepareSources(placed, inputs, sx, sy, steps.decode, {
        between: steps.between,
        cancelled: stop,
      });
      if (prepared.cancelled || cancelled) {
        return null;
      }
      if (prepared.failed.length > 0) {
        const names = prepared.failed.map((index) => `“${entries[index]?.title ?? ""}”`).join(", ");
        throw new Error(`${names} could not be read.`);
      }
      canvas = steps.createCanvas(size.widthPx, size.heightPx);
      const drawn = await composeToCanvas(canvas, placed, { geometry, background }, prepared.sources, {
        between: steps.between,
        cancelled: stop,
        onProgress: (done, total) => {
          progress.value = `Item ${done} of ${total}`;
        },
      });
      if (cancelled) {
        return null;
      }
      if (!drawn) {
        throw new Error("this browser could not draw the image.");
      }
      const bytes = await steps.encode(canvas, contentType, quality);
      return cancelled ? null : bytes;
    } finally {
      if (canvas) {
        steps.release(canvas);
      }
      busy.value = false;
      progress.value = "";
    }
  }

  async function run<T>(action: string, then: (bytes: Uint8Array) => Promise<T> | T): Promise<T | null> {
    if (busy.value) {
      return null;
    }
    failure.value = "";
    try {
      const bytes = await produce();
      return bytes ? await then(bytes) : null;
    } catch (error) {
      const reason = error instanceof Error && error.message ? error.message : String(error);
      failure.value = `Could not ${action}: ${reason}`;
      return null;
    }
  }

  const title = () => output.title.trim() || DEFAULT_ASSEMBLY_TITLE;

  // Leaving the workspace mid-way stops the work: no file, no record.
  onScopeDispose(() => {
    cancelled = busy.value || cancelled;
  });

  return {
    pixels,
    estimate,
    refusal,
    fixDpi,
    backgroundNotice,
    busy,
    progress,
    failure,
    canProduce,
    save: () =>
      run("save the image", async (bytes) => {
        const taken = listDocuments().map((entry) => entry.title);
        // Its `image` summary is read back from the bytes just produced.
        const model = await assembleMediaModel(uniqueTitle(title(), taken), output.format, bytes);
        storeMedia(model);
        return model;
      }),
    download: async () =>
      (await run("download the image", (bytes) => {
        steps.download(bytes, outputFileName(title(), output.format), ASSEMBLY_CONTENT_TYPES[output.format]);
        return true;
      })) ?? false,
    cancel: () => {
      if (busy.value) {
        cancelled = true;
      }
    },
  };
}
