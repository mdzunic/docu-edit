// The image editor's four outputs (SPEC-112), each explicit about what it
// touches and each a single write-through:
// - Save stores the stack on the record and refreshes the render, leaving
//   `sourceMediaId` and its bytes untouched;
// - Save as a copy creates a NEW media document carrying the same source
//   and stack;
// - Export downloads a flattened PNG or JPEG through SPEC-108's background
//   rules (a JPEG cannot be transparent, and says so);
// - Flatten and discard original — confirmed first, permanent, not undoable
//   — renders the stack and REPLACES the source part's bytes with it.
// Full-resolution work happens only here: the full decode is rendered
// through the same renderEditedImage the preview uses, chunked with
// yieldToMainThread() between stages and reported through `progress`
// (role="status"), and refused up front against the probed canvas ceiling.
// Every action renders and encodes BEFORE it writes anything, so a failure
// is reported with nothing mutated.
import { ref, type Ref } from "vue";
import { resolveEditGeometry, type ImageEdit } from "../../lib/core/imageEdits";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { backgroundRefusal, effectiveBackground, TRANSPARENT } from "../../lib/core/sheetLayout";
import { listDocuments, storeMedia, storeMediaEdit, uniqueTitle } from "../../lib/documentStore";
import { canvasCeilingRefusal } from "../../lib/media/canvasCeiling";
import {
  copiedWithStack,
  renderContentType,
  withFlattened,
  withSavedStack,
  type RenderedImage,
} from "../../lib/media/editModel";
import type { MediaModel } from "../../lib/media/model";
import type { EditSource, EncodeOptions, ImageEditSteps, RenderStage } from "../../lib/media/renderEdits";
import type { EditSession } from "./editSession";

/** The quality a JPEG render is stored at. */
export const RENDER_JPEG_QUALITY = 0.92;

const STAGE_LABELS: Record<RenderStage, string> = {
  geometry: "drawing the image",
  adjust: "adjusting colours",
  redact: "drawing redactions",
  annotate: "drawing annotations",
};

export interface EditActionsDeps {
  media: Ref<MediaModel | null>;
  session: EditSession;
  steps: ImageEditSteps;
  /** The full-resolution source (the session's single decode). */
  fullSource: () => Promise<EditSource>;
  /** An SVG's raster base made this session and not yet stored. */
  pendingRaster: () => RenderedImage | null;
  /** After Save stored `model`. */
  onSaved: (model: MediaModel) => void;
  /** After Flatten stored `model`: the session restarts on the new source. */
  onFlattened: (model: MediaModel) => Promise<void>;
}

export interface EditActions {
  busy: Ref<boolean>;
  /** The running action's stage ("Saving: adjusting colours…"), "" idle. */
  progress: Ref<string>;
  /** The last action's failure, "" when none. */
  failure: Ref<string>;
  /** The last action's outcome ("Saved."), "" when none. */
  notice: Ref<string>;
  /** The id of the copy Save as a copy last made, "" before one. */
  copyId: Ref<string>;
  flattenPending: Ref<boolean>;
  exportOpen: Ref<boolean>;
  exportFormat: Ref<"png" | "jpeg">;
  /** JPEG quality, 1–100. */
  exportQuality: Ref<number>;
  /** "#rrggbb" or "transparent" (SPEC-108's background vocabulary). */
  exportBackground: Ref<string>;
  exportBackgroundNotice: () => string;
  save: () => Promise<boolean>;
  saveCopy: () => Promise<boolean>;
  openExport: () => void;
  cancelExport: () => void;
  exportImage: () => Promise<boolean>;
  requestFlatten: () => void;
  cancelFlatten: () => void;
  confirmFlatten: () => Promise<boolean>;
}

function reason(error: unknown): string {
  return error instanceof Error && error.message ? error.message : String(error);
}

export function useImageEditActions(deps: EditActionsDeps): EditActions {
  const busy = ref(false);
  const progress = ref("");
  const failure = ref("");
  const notice = ref("");
  const copyId = ref("");
  const flattenPending = ref(false);
  const exportOpen = ref(false);
  const exportFormat = ref<"png" | "jpeg">("png");
  const exportQuality = ref(90);
  const exportBackground = ref(TRANSPARENT);

  /** Renders `stack` at full resolution and encodes it. */
  async function renderFull(stack: ImageEdit[], contentType: string, options: EncodeOptions, verb: string): Promise<RenderedImage> {
    const naturalPx = deps.session.naturalPx.value;
    if (!naturalPx) {
      throw new Error("The image is not ready.");
    }
    const geometry = resolveEditGeometry(stack, naturalPx);
    const size = { widthPx: geometry.widthPx, heightPx: geometry.heightPx };
    const refusal = canvasCeilingRefusal(size, deps.steps.canvasCeiling(size.widthPx * size.heightPx));
    if (refusal) {
      throw new Error(refusal);
    }
    progress.value = `${verb}: reading the original…`;
    const source = await deps.fullSource();
    const canvas = await deps.steps.render(source, stack, size, {
      chunked: true,
      onStage: (stage) => {
        progress.value = `${verb}: ${STAGE_LABELS[stage]}…`;
      },
    });
    try {
      await yieldToMainThread();
      progress.value = `${verb}: encoding…`;
      const bytes = await deps.steps.encode(canvas, contentType, options);
      return { bytes, contentType, ...size };
    } finally {
      deps.steps.release(canvas);
    }
  }

  async function run(failureVerb: string, work: (model: MediaModel, stack: ImageEdit[]) => Promise<void>): Promise<boolean> {
    const model = deps.media.value;
    if (busy.value || !model) {
      return false;
    }
    deps.session.end();
    busy.value = true;
    failure.value = "";
    notice.value = "";
    try {
      await work(model, deps.session.stack.value);
      return true;
    } catch (error) {
      failure.value = `Could not ${failureVerb}: ${reason(error)}`;
      return false;
    } finally {
      busy.value = false;
      progress.value = "";
    }
  }

  /** The render a saved stack carries (none for an empty stack), and the
   * base's natural size recorded beside it. */
  async function savedRender(model: MediaModel, stack: ImageEdit[], verb: string): Promise<RenderedImage | null> {
    return stack.length > 0 ? renderFull(stack, renderContentType(model), { quality: RENDER_JPEG_QUALITY }, verb) : null;
  }

  function withNatural(model: MediaModel, stack: ImageEdit[]): MediaModel {
    const naturalPx = deps.session.naturalPx.value;
    return naturalPx && stack.length > 0 ? { ...model, image: { ...naturalPx } } : model;
  }

  const save = () =>
    run("save the image", async (model, stack) => {
      const render = await savedRender(model, stack, "Saving");
      const raster = stack.length > 0 ? deps.pendingRaster() : null;
      const next = withNatural(withSavedStack(model, stack, render, raster), stack);
      progress.value = "Saving: writing to the library…";
      await storeMediaEdit(next);
      deps.onSaved(next);
      notice.value = "Saved.";
    });

  const saveCopy = () =>
    run("save a copy", async (model, stack) => {
      const render = await savedRender(model, stack, "Saving a copy");
      const title = uniqueTitle(`${model.title} (copy)`, listDocuments().map((entry) => entry.title));
      const copy = withNatural(copiedWithStack(model, stack, render, stack.length > 0 ? deps.pendingRaster() : null, title), stack);
      storeMedia(copy);
      copyId.value = copy.id;
      notice.value = `Saved a copy as “${title}”.`;
    });

  function exportBackgroundNotice(): string {
    return backgroundRefusal(exportBackground.value, exportFormat.value) ?? "";
  }

  const exportImage = () =>
    run("export the image", async (model, stack) => {
      const jpeg = exportFormat.value === "jpeg";
      const contentType = jpeg ? "image/jpeg" : "image/png";
      const rendered = await renderFull(
        stack,
        contentType,
        {
          quality: Math.min(100, Math.max(1, Math.round(exportQuality.value))) / 100,
          background: effectiveBackground(exportBackground.value, exportFormat.value),
        },
        "Exporting",
      );
      const title = model.title.trim() || "Untitled";
      deps.steps.download(rendered.bytes, `${title}.${jpeg ? "jpeg" : "png"}`, contentType);
      exportOpen.value = false;
      notice.value = `Exported ${rendered.widthPx} × ${rendered.heightPx} px ${jpeg ? "JPEG" : "PNG"}.`;
    });

  async function confirmFlatten(): Promise<boolean> {
    flattenPending.value = false;
    return run("flatten the image", async (model, stack) => {
      const rendered = await renderFull(stack, renderContentType(model), { quality: RENDER_JPEG_QUALITY }, "Flattening");
      const next = withFlattened(model, rendered);
      progress.value = "Flattening: writing to the library…";
      await storeMediaEdit(next);
      await deps.onFlattened(next);
      notice.value = "Flattened. The original has been replaced and cannot be restored.";
    });
  }

  return {
    busy,
    progress,
    failure,
    notice,
    copyId,
    flattenPending,
    exportOpen,
    exportFormat,
    exportQuality,
    exportBackground,
    exportBackgroundNotice,
    save,
    saveCopy,
    openExport: () => {
      exportOpen.value = true;
    },
    cancelExport: () => {
      exportOpen.value = false;
    },
    exportImage,
    requestFlatten: () => {
      if (!busy.value && deps.session.stack.value.length > 0) {
        flattenPending.value = true;
      }
    },
    cancelFlatten: () => {
      flattenPending.value = false;
    },
    confirmFlatten,
  };
}
