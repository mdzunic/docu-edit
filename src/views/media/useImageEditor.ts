// The image editor's session (SPEC-112): /media/:id's edit mode. Entering it
// decodes the stack's base once (an SVG is rasterised first, at its natural
// size or one the user picks), refuses up front an image over the probed
// canvas ceiling, and builds the preview's downscaled buffer; the history
// list starts from the saved stack. Every tool then mutates the history
// through the session's gesture API (editSession.ts), so undo/redo cover the
// whole session one gesture at a time (useEditUndo.ts). Nothing is written
// until Save, Save as a copy or Flatten (useImageEditActions.ts); during the
// session the preview — renderEditedImage at screen scale — is the current
// image, which `liveMedia` hands the details panel.
//
// Keyboard: Escape leaves (Done), Cmd/Ctrl+Z / Shift+Cmd/Ctrl+Z / Cmd/Ctrl+Y
// undo and redo (SPEC-015's bindings, from the shortcut registry), and
// Cmd/Ctrl+[ / Cmd/Ctrl+] rotate left and right.
import { computed, onScopeDispose, ref, shallowRef, watch, type Ref } from "vue";
import { describeEdit, normaliseStack, type ImageEdit, type PixelSize } from "../../lib/core/imageEdits";
import { gifFrameCount } from "../../lib/core/imageDimensions";
import { removeAt } from "../../lib/core/imageEditHistory";
import { findShortcut } from "../../lib/shortcuts";
import { canvasCeilingRefusal } from "../../lib/media/canvasCeiling";
import { mediaPartBytes } from "../../lib/media/download";
import { editBasePart, isEditableImage, savedStack, type RenderedImage } from "../../lib/media/editModel";
import type { MediaModel } from "../../lib/media/model";
import { BROWSER_EDIT_STEPS, createDecodeCache, type EditSource, type ImageEditSteps } from "../../lib/media/renderEdits";
import { createEditSession } from "./editSession";
import { useCropTool } from "./useCropTool";
import { useEditCommands } from "./useEditCommands";
import { sameHistory, useEditUndo } from "./useEditUndo";
import { useImageEditActions } from "./useImageEditActions";
import { useObjectTool } from "./useObjectTool";

export type EditTool = "crop" | "rotate" | "adjust" | "redact" | "annotate";

export const EDIT_TOOLS: readonly EditTool[] = ["crop", "rotate", "adjust", "redact", "annotate"];

export const EDIT_TOOL_LABELS: Record<EditTool, string> = {
  crop: "Crop & straighten",
  rotate: "Rotate, flip & resize",
  adjust: "Adjust",
  redact: "Redact",
  annotate: "Annotate",
};

/** The notice the editor shows while the stack holds a redaction. */
export const REDACTION_NOTICE =
  "The original is still stored. Choose Flatten and discard original to remove the hidden pixels.";

export const SVG_NOTICE = "Editing is raster: the SVG is edited as pixels, and the vector original is preserved.";

export const GIF_NOTICE = "This GIF is animated: only its first frame can be edited, and saving keeps that frame alone.";

/** The size an SVG that declares none is offered at (CSS's default). */
export const DEFAULT_SVG_SIZE: PixelSize = { widthPx: 300, heightPx: 150 };

export interface ImageEditorDeps {
  media: Ref<MediaModel | null>;
  /** After Save or Flatten stored a new model under this id. */
  onStored?: () => void;
  steps?: ImageEditSteps;
}

/** Whether a key event comes from a field where Cmd/Ctrl+Z edits text. */
function inTextField(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable || target.tagName === "TEXTAREA" || target.tagName === "SELECT") return true;
  return target.tagName === "INPUT" && !["range", "checkbox", "radio", "button", "color"].includes((target as HTMLInputElement).type);
}

function noBytes(): Uint8Array {
  return new Uint8Array();
}

export function useImageEditor(deps: ImageEditorDeps) {
  const steps = deps.steps ?? BROWSER_EDIT_STEPS;
  const active = ref(false);
  const phase = ref<"idle" | "loading" | "size" | "ready" | "refused">("idle");
  const refusal = ref("");
  const tool = ref<EditTool>("crop");
  const history = ref<ImageEdit[]>([]);
  const savedHistory = ref<ImageEdit[]>([]);
  const naturalPx = ref<PixelSize | null>(null);
  const preview = shallowRef<EditSource | null>(null);
  const svgSize = ref<PixelSize>({ ...DEFAULT_SVG_SIZE });
  const isSvg = ref(false);
  const isAnimatedGif = ref(false);
  const revertPending = ref(false);
  const discardPending = ref(false);
  let raster: RenderedImage | null = null;
  let decodes = createDecodeCache(steps.decode);
  let baseKey = "";
  let baseBytes = noBytes;
  let baseType = "";

  const undo = useEditUndo(history);
  const session = createEditSession(history, naturalPx, undo);
  const commands = useEditCommands(session);
  const crop = useCropTool(session);
  const objects = useObjectTool(session);
  const stack = session.stack;
  const geometry = session.geometry;

  async function fullSource(): Promise<EditSource> {
    const decoded = await decodes.get(baseKey, baseBytes, baseType);
    return { image: decoded.image, naturalPx: { widthPx: decoded.widthPx, heightPx: decoded.heightPx } };
  }

  /** Decodes the base, checks the ceiling and builds the preview buffer. */
  async function load(model: MediaModel): Promise<void> {
    const size = naturalPx.value ?? model.image ?? null;
    if (size) {
      const ceilingRefusal = canvasCeilingRefusal(size, steps.canvasCeiling(size.widthPx * size.heightPx));
      if (ceilingRefusal) throw new Error(ceilingRefusal);
    }
    const decoded = await decodes.get(baseKey, baseBytes, baseType);
    const decodedSize = { widthPx: decoded.widthPx, heightPx: decoded.heightPx };
    const ceilingRefusal = canvasCeilingRefusal(decodedSize, steps.canvasCeiling(decoded.widthPx * decoded.heightPx));
    if (ceilingRefusal) throw new Error(ceilingRefusal);
    naturalPx.value = decodedSize;
    preview.value = steps.preview(decoded);
    history.value = savedStack(model);
    savedHistory.value = history.value;
    undo.reset();
    phase.value = "ready";
  }

  async function begin(model: MediaModel, rasterSize?: PixelSize): Promise<void> {
    phase.value = "loading";
    try {
      raster = null;
      const base = editBasePart(model);
      if (!base) throw new Error("The image's bytes are missing.");
      if (isSvg.value && !model.baseMediaId) {
        // The SVG is rasterised once, here, and the raster becomes the
        // stack's base; the vector original stays the source.
        const png = await steps.rasteriseSvg(mediaPartBytes(base), rasterSize);
        raster = { bytes: png.data, contentType: "image/png", widthPx: png.widthPx, heightPx: png.heightPx };
        const bytes = png.data;
        baseKey = `${model.sourceMediaId}:raster`;
        baseBytes = () => bytes;
        baseType = "image/png";
      } else {
        baseKey = base.mediaId;
        baseBytes = () => mediaPartBytes(base);
        baseType = base.contentType;
      }
      naturalPx.value = null;
      await load(model);
    } catch (error) {
      phase.value = "refused";
      refusal.value = error instanceof Error && error.message ? error.message : String(error);
    }
  }

  /** Enters edit mode on the current media document. */
  async function enter(): Promise<void> {
    const model = deps.media.value;
    if (!model || !isEditableImage(model) || active.value) return;
    active.value = true;
    tool.value = "crop";
    refusal.value = "";
    actions.failure.value = "";
    actions.notice.value = "";
    isSvg.value = model.contentType === "image/svg+xml";
    const base = editBasePart(model);
    isAnimatedGif.value = model.format === "gif" && base !== undefined && gifFrameCount(mediaPartBytes(base)) > 1;
    if (isSvg.value && !model.baseMediaId && !model.image) {
      // No declared size: the user picks one before anything is drawn.
      phase.value = "size";
      return;
    }
    await begin(model, isSvg.value && model.image ? model.image : undefined);
  }

  /** Starts an unsized SVG at the picked size. */
  async function startSvg(): Promise<void> {
    const model = deps.media.value;
    if (!model || phase.value !== "size") return;
    const size = {
      widthPx: Math.max(1, Math.round(svgSize.value.widthPx)),
      heightPx: Math.max(1, Math.round(svgSize.value.heightPx)),
    };
    await begin(model, size);
  }

  function exit(): void {
    session.cancel();
    active.value = false;
    phase.value = "idle";
    discardPending.value = false;
    revertPending.value = false;
    actions.flattenPending.value = false;
    actions.exportOpen.value = false;
    preview.value = null;
    naturalPx.value = null;
    history.value = [];
    savedHistory.value = [];
    raster = null;
    undo.reset();
    decodes.clear();
    decodes = createDecodeCache(steps.decode);
  }

  /** Whether the session's stack differs from the one last saved. */
  const dirty = computed(() => !sameHistory(stack.value, normaliseStack(savedHistory.value)));

  /** Done / Escape: leaves at once, or asks first when edits are unsaved. */
  function done(): void {
    if (actions.busy.value) return;
    session.end();
    if (phase.value === "ready" && dirty.value) {
      discardPending.value = true;
      return;
    }
    exit();
  }

  const actions = useImageEditActions({
    media: deps.media,
    session,
    steps,
    fullSource,
    pendingRaster: () => raster,
    onSaved: (model) => {
      savedHistory.value = history.value;
      raster = null;
      deps.media.value = model;
      deps.onStored?.();
    },
    onFlattened: async (model) => {
      deps.media.value = model;
      deps.onStored?.();
      // Nothing left to undo into: the session restarts on the new source.
      decodes.clear();
      decodes = createDecodeCache(steps.decode);
      isSvg.value = false;
      isAnimatedGif.value = false;
      await begin(model);
    },
  });

  function removeEntry(index: number): void {
    const edit = history.value[index];
    if (edit) {
      session.apply(`Remove ${describeEdit(edit).toLowerCase()}`, removeAt(history.value, index));
    }
  }

  /** "Crop, Rotate right and Adjust brightness" — what a revert discards. */
  const revertSummary = computed(() => history.value.map(describeEdit).join(", "));

  function confirmRevert(): void {
    revertPending.value = false;
    session.apply("Revert to original", []);
  }

  function undoStep(): void {
    session.end();
    commands.release();
    undo.undo();
  }

  function redoStep(): void {
    session.end();
    commands.release();
    undo.redo();
  }

  function handleKeydown(event: KeyboardEvent): boolean {
    if (!active.value || event.defaultPrevented) return false;
    const target = event.target;
    if (target instanceof Element && target.closest('[role="dialog"]')) return false;
    const mod = event.metaKey || event.ctrlKey;
    if (event.key === "Escape" && !mod) {
      done();
      return true;
    }
    if (phase.value !== "ready" || actions.busy.value) return false;
    if (mod && !event.altKey && (event.key === "[" || event.key === "]" || event.code === "BracketLeft" || event.code === "BracketRight")) {
      commands.rotate(event.key === "]" || event.code === "BracketRight" ? 1 : -1);
      return true;
    }
    const shortcut = findShortcut(event);
    if ((shortcut?.id === "edit.undo" || shortcut?.id === "edit.redo") && !inTextField(target)) {
      if (shortcut.id === "edit.undo") undoStep();
      else redoStep();
      return true;
    }
    return false;
  }

  function onWindowKeydown(event: KeyboardEvent): void {
    if (handleKeydown(event)) event.preventDefault();
  }
  watch(active, (on) => {
    if (typeof window === "undefined") return;
    if (on) window.addEventListener("keydown", onWindowKeydown);
    else window.removeEventListener("keydown", onWindowKeydown);
  });
  onScopeDispose(() => {
    if (typeof window !== "undefined") window.removeEventListener("keydown", onWindowKeydown);
    decodes.clear();
  });

  // Leaving the document (a route change to another id) ends the session.
  watch(
    () => deps.media.value?.id,
    (id, previous) => {
      if (active.value && id !== previous) exit();
    },
  );

  /** The current image for the details panel: during a session, the stack
   * the preview shows over the base's natural size. */
  const liveMedia = computed(() => {
    const model = deps.media.value;
    if (!model || !active.value || !naturalPx.value) return model;
    return { ...model, image: { ...naturalPx.value }, edits: stack.value };
  });

  /** What the preview renders: the crop tool shows the whole croppable area
   * (the straightened, adjusted source), every other tool the full stack. */
  const previewStack = computed(() =>
    tool.value === "crop" ? stack.value.filter((edit) => edit.kind === "straighten" || edit.kind === "adjust") : stack.value,
  );

  return {
    active,
    phase,
    refusal,
    tool,
    history,
    stack,
    geometry,
    naturalPx,
    preview,
    previewStack,
    liveMedia,
    dirty,
    isSvg,
    isAnimatedGif,
    svgSize,
    hasRedaction: computed(() => stack.value.some((edit) => edit.kind === "redact")),
    revertPending,
    revertSummary,
    discardPending,
    undo,
    session,
    commands,
    crop,
    objects,
    actions,
    steps,
    enter,
    startSvg,
    done,
    exit,
    removeEntry,
    requestRevert: () => {
      if (history.value.length > 0) revertPending.value = true;
    },
    cancelRevert: () => {
      revertPending.value = false;
    },
    confirmRevert,
    cancelDiscard: () => {
      discardPending.value = false;
    },
    undoStep,
    redoStep,
    handleKeydown,
    setTool: (next: EditTool) => {
      session.end();
      commands.release();
      tool.value = next;
    },
  };
}

export type ImageEditor = ReturnType<typeof useImageEditor>;
