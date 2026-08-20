// Images (SPEC-024): the selected image's state (selection, measured handles,
// options dialog), insertion from a File (chosen/dropped/pasted, with
// header/footer support and SVG→PNG conversion for the ODT format), wrap-mode
// toggling, and the eight-handle resize drag. The `contenteditable` canvas
// routes mousedown/image-select here via the view.
import { nextTick, ref } from "vue";
import { caretSelection, documentStart, type ModelSelection } from "../../lib/commands";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import {
  asImageHandle,
  noImageResizeDrag,
  previewImageSize,
  resizedImageSize,
  type ImageResizeDrag,
} from "../../lib/imageResize";
import { CM_PER_PX } from "../../lib/core/units";
import {
  deleteImage,
  imageAt,
  insertImage,
  setImageAltText,
  setImageSize,
  setImageWrap,
  noImageCommand,
  type ImageLocation,
} from "../../lib/images";
import { pxToCm } from "../../lib/core/units";
import { asImageWrapMode, type DocumentFormat } from "../../lib/documentModel";
import { prepareImageInsert, ImageInsertError } from "../../lib/imageInsert";
import { withHeaderFooterModel } from "../../lib/commands";
import type { EditorContext, HeaderFooterBridge } from "./context";
import type { DocumentModel } from "../../lib/documentModel";
import type { CommandResult } from "../../lib/commands";

const HANDLE_SIZE = 8;

type ImageMutate = (
  model: DocumentModel,
  selection: ModelSelection,
  location: ImageLocation,
) => CommandResult;

// The bridge carries the header/footer branch of image insertion (SPEC-029
// AC-6); everything else images needs is the universal context.
export type ImagesDeps = EditorContext & HeaderFooterBridge;

export function useImages(deps: ImagesDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    currentSelection,
    applyEdit,
    clearCellSelection,
    addToast,
    activeHeaderFooter,
    currentHfSelection,
    applyHeaderFooterEdit,
    getCaret,
  } = deps;

  const imageSelection = ref<ImageLocation | null>(null);
  const imageHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  const imageOptionsOpen = ref(false);
  const imageOptionsInitial = ref({ altText: "", widthCm: 0, heightCm: 0, naturalWidthCm: 0, naturalHeightCm: 0 });
  let imageResizeDrag: ImageResizeDrag | null = noImageResizeDrag();
  let imageDragMoved = false;

  function clearImageSelection() {
    if (imageSelection.value) {
      imageSelection.value = null;
      imageHandles.value = null;
    }
  }

  /** The document's own effective save format — decides whether an inserted
   * SVG is converted to PNG (AC-2) or kept as-is (AC-3). A `.md` document
   * (SPEC-050) keeps SVG as-is: Markdown images are written as data URIs,
   * format-agnostic, so there is no docx-style "cannot store SVG" reason to
   * convert. */
  function effectiveFormat(): DocumentFormat {
    return doc.value?.format === "odt" || doc.value?.format === "md" ? doc.value.format : "docx";
  }

  function remeasureImageHandles() {
    const location = imageSelection.value;
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!location || !pageEl || !root) {
      return;
    }
    const imageEl = root.querySelector(
      `.doc-image-anchor[data-section="${location.section}"][data-block="${location.block}"]`,
    );
    imageHandles.value = imageEl ? measureImageGeometry(pageEl, imageEl) : null;
  }

  function handleImageSelect(section = 0, block = 0) {
    clearCellSelection();
    imageSelection.value = { section, block };
    void nextTick().then(remeasureImageHandles);
  }

  /** One resize handle's page-relative position, centred on its edge/corner
   * of the measured image rect (AC-8). */
  function imageHandleStyle(handle = "se"): Record<string, string> {
    const geometry = imageHandles.value;
    if (!geometry) {
      return {};
    }
    const xFrac = handle.includes("w") ? 0 : handle.includes("e") ? 1 : 0.5;
    const yFrac = handle.includes("n") ? 0 : handle.includes("s") ? 1 : 0.5;
    return {
      left: `${geometry.left + geometry.width * xFrac - HANDLE_SIZE / 2}px`,
      top: `${geometry.top + geometry.height * yFrac - HANDLE_SIZE / 2}px`,
      cursor: `${handle}-resize`,
    };
  }

  /** The floating image toolbar's fixed (viewport) position (AC-9). */
  function imageToolbarStyle(): { left: number; top: number } {
    const geometry = imageHandles.value;
    const pageEl = pageRef.value;
    if (!geometry || !pageEl) {
      return { left: 0, top: 0 };
    }
    const pageRect = pageEl.getBoundingClientRect();
    const viewportTop = pageRect.top + geometry.top;
    const above = viewportTop - 40;
    return {
      left: pageRect.left + geometry.left,
      top: above >= 0 ? above : pageRect.top + geometry.top + geometry.height + 8,
    };
  }

  /** Runs an image command against the currently selected image and commits
   * it — the shared apply step every toolbar/dialog/resize action funnels through. */
  function applyImageEdit(mutate: ImageMutate | null = noImageCommand(), label = "") {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const location = imageSelection.value;
    if (!mutate || !model || !root || !location) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = mutate(model, selection, location);
    applyEdit(result, selection, label);
  }

  function handleImageWrap(wrap = "inline") {
    const location = imageSelection.value;
    if (!location) {
      return;
    }
    applyImageEdit(
      (model, selection) => setImageWrap(model, selection, location, asImageWrapMode(wrap)),
      "wrap image",
    );
    // A wrap change moves the image between text flow and an absolute
    // offset; the selection handles must re-find it (same as useShapes).
    void nextTick().then(remeasureImageHandles);
  }

  function handleImageDelete() {
    const location = imageSelection.value;
    if (!location) {
      return;
    }
    applyImageEdit((model, selection) => deleteImage(model, selection, location), "delete image");
    clearImageSelection();
  }

  function openImageOptions() {
    const model = doc.value;
    const location = imageSelection.value;
    const image = model && location ? imageAt(model, location) : null;
    if (!image) {
      return;
    }
    imageOptionsInitial.value = {
      altText: image.altText || "",
      widthCm: Number(image.widthCm.toFixed(2)),
      heightCm: Number(image.heightCm.toFixed(2)),
      naturalWidthCm: Number(pxToCm(image.naturalWidthPx).toFixed(2)),
      naturalHeightCm: Number(pxToCm(image.naturalHeightPx).toFixed(2)),
    };
    imageOptionsOpen.value = true;
  }

  function handleImageOptionsApply(payload = { altText: "", widthCm: 0, heightCm: 0 }) {
    imageOptionsOpen.value = false;
    const location = imageSelection.value;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!location || !model || !root) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    let result = setImageAltText(model, selection, location, payload.altText);
    result = setImageSize(result.model, result.selection, location, payload.widthCm, payload.heightCm);
    applyEdit(result, selection, "image options");
    root.focus();
  }

  /** Reads a File chosen/dropped/pasted and inserts it at the caret (AC-1),
   * showing validation/conversion failures as a toast (AC-34). */
  async function insertImageFromFile(file = new File([], ""), atSelection: ModelSelection | null = null) {
    const active = activeHeaderFooter.value;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const selection = active
      ? (atSelection ?? currentHfSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 }))
      : (atSelection ??
          (root ? currentSelection(root) : null) ??
          (model ? caretSelection(getCaret() ?? documentStart()) : null));
    if (!model || !selection) {
      return;
    }
    let prepared: Awaited<ReturnType<typeof prepareImageInsert>>;
    try {
      prepared = await prepareImageInsert(file, effectiveFormat());
    } catch (error) {
      if (error instanceof ImageInsertError) {
        addToast(error.message);
        return;
      }
      throw error;
    }
    const current = doc.value;
    if (!current) {
      return;
    }
    if (active) {
      // Header/footer images (SPEC-029 AC-6): inserted into the active band's
      // synthetic model.
      const synthetic = withHeaderFooterModel(current, active.sectionIndex, active.part, active.kind);
      const result = insertImage(synthetic, selection, prepared.media, prepared.naturalWidthPx, prepared.naturalHeightPx);
      applyHeaderFooterEdit(result.model, result.selection, "insert image");
      return;
    }
    const result = insertImage(current, selection, prepared.media, prepared.naturalWidthPx, prepared.naturalHeightPx);
    applyEdit(result, selection, "insert image");
  }

  function handleInsertImageFile(file = new File([], "")) {
    void insertImageFromFile(file);
  }

  // --- image resize handles (SPEC-024 AC-8, AC-10..12) ----------------------

  function startImageResize(event: MouseEvent, handle = "se") {
    const location = imageSelection.value;
    const model = doc.value;
    const root = contentRoot.value;
    if (!location || !model || !root || !editable.value) {
      return;
    }
    const image = imageAt(model, location);
    const imageEl = root.querySelector(
      `.doc-image-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-image`,
    );
    if (!image || !(imageEl instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    imageDragMoved = false;
    imageResizeDrag = {
      section: location.section,
      block: location.block,
      handle: asImageHandle(handle),
      startClientX: event.clientX,
      startClientY: event.clientY,
      start: { widthCm: image.widthCm, heightCm: image.heightCm },
      imageEl,
    };
    window.addEventListener("mousemove", handleImageResizeMove);
    window.addEventListener("mouseup", finishImageResize);
  }

  function handleImageResizeMove(event: MouseEvent) {
    const drag = imageResizeDrag;
    if (!drag) {
      return;
    }
    imageDragMoved = true;
    const deltaXCm = (event.clientX - drag.startClientX) * CM_PER_PX;
    const deltaYCm = (event.clientY - drag.startClientY) * CM_PER_PX;
    const size = resizedImageSize(drag.start, drag.handle, deltaXCm, deltaYCm, event.shiftKey);
    previewImageSize(drag.imageEl, size.widthCm, size.heightCm);
  }

  /** Release (AC-10..12): the live-previewed size becomes the one model
   * update this whole drag makes. */
  function finishImageResize() {
    window.removeEventListener("mousemove", handleImageResizeMove);
    window.removeEventListener("mouseup", finishImageResize);
    const drag = imageResizeDrag;
    imageResizeDrag = null;
    if (!drag || !imageDragMoved) {
      return;
    }
    const widthCm = parseFloat(drag.imageEl.style.width) || drag.start.widthCm;
    const heightCm = parseFloat(drag.imageEl.style.height) || drag.start.heightCm;
    const location = { section: drag.section, block: drag.block };
    applyImageEdit(
      (model, selection) => setImageSize(model, selection, location, widthCm, heightCm),
      "resize image",
    );
  }

  return {
    imageSelection,
    imageHandles,
    imageOptionsOpen,
    imageOptionsInitial,
    clearImageSelection,
    handleImageSelect,
    imageHandleStyle,
    imageToolbarStyle,
    remeasureImageHandles,
    handleImageWrap,
    handleImageDelete,
    openImageOptions,
    handleImageOptionsApply,
    insertImageFromFile,
    handleInsertImageFile,
    startImageResize,
    handleImageResizeMove,
    finishImageResize,
  };
}
