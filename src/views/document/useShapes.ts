// Shapes and text boxes (SPEC-051): the selected shape's state (selection,
// measured handles), draw-mode insertion (gallery pick -> drag/click on the
// canvas), the shape toolbar's actions, the eight-handle resize drag, the
// rotation-handle drag, keyboard nudge/duplicate, and z-order/grouping.
// Mirrors useImages.ts's shape closely — see its own header comment for the
// pattern this follows.
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import { caretSelection, documentStart, type ModelSelection } from "../../lib/commands";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import { asImageHandle, previewImageSize, resizedImageSize } from "../../lib/imageResize";
import { CM_PER_PX } from "../../lib/tableResize";
import {
  constrainedDragSizeCm,
  defaultClickSizeCm,
  DEFAULT_NEW_SHAPE_FILL,
  DEFAULT_NEW_SHAPE_OUTLINE,
  deleteShape,
  duplicateShape,
  freshGroupOrdinal,
  groupMembers,
  groupShapes,
  insertShape,
  insertTextBox,
  moveGroup,
  nudgeShape,
  normalizeRotationDeg,
  reorderShapeZ,
  resizeGroup,
  rotateGroup,
  setShapeAltText,
  setShapeFill,
  setShapeOutline,
  setShapeRotation,
  setShapeSize,
  setShapeWrap,
  shapeAt,
  ungroupShapes,
  type NudgeDirection,
  type ShapeLocation,
  type ZOrderCommand,
} from "../../lib/shapes";
import {
  asImageWrapMode,
  isShape,
  LINE_GEOMETRIES,
  type DocumentModel,
  type ShapeGeometry,
  type ShapeOutline,
} from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { CommandResult } from "../../lib/commands";

const HANDLE_SIZE = 8;
const ROTATION_HANDLE_OFFSET_PX = 24;

type ShapeMutate = (model: DocumentModel, selection: ModelSelection, location: ShapeLocation) => CommandResult;

export interface ShapesDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  pageRef: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
  ) => void;
  clearCellSelection: () => void;
  getCaret: () => CaretPosition | null;
}

export function useShapes(deps: ShapesDeps) {
  const { doc, contentRoot, pageRef, editable, currentSelection, applyEdit, clearCellSelection, getCaret } = deps;

  const shapeSelection = ref<ShapeLocation | null>(null);
  const shapeHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  /** Double-clicking a shape (AC-17) puts the caret inside its text frame —
   * this just toggles the proxy's own contenteditable region on; DOM->model
   * sync for typed edits is not yet wired (see the implement summary). */
  const editingShapeText = ref(false);

  function handleShapeEditText(section = 0, block = 0) {
    clearCellSelection();
    shapeSelection.value = { section, block };
    editingShapeText.value = true;
    void nextTick().then(remeasureShapeHandles);
  }
  /** Armed by picking a gallery entry (AC-3): the geometry the next
   * drag/click on the canvas draws, or "textbox" — cleared once drawn. */
  const drawGeometry = ref<ShapeGeometry | "textbox" | null>(null);
  let resizeDrag: { location: ShapeLocation; handle: string; startClientX: number; startClientY: number; start: { widthCm: number; heightCm: number }; el: HTMLElement } | null = null;
  let resizeDragMoved = false;
  let rotateDrag: { location: ShapeLocation; centerClientX: number; centerClientY: number; startRotation: number; startAngle: number } | null = null;
  let drawDrag: { startClientX: number; startClientY: number; anchor: CaretPosition; anchorLeft: number; anchorTop: number } | null = null;

  function clearShapeSelection() {
    if (shapeSelection.value) {
      shapeSelection.value = null;
      shapeHandles.value = null;
    }
    editingShapeText.value = false;
  }

  function remeasureShapeHandles() {
    const location = shapeSelection.value;
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!location || !pageEl || !root) {
      return;
    }
    const el = root.querySelector(
      `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
    );
    shapeHandles.value = el ? measureImageGeometry(pageEl, el) : null;
  }

  function handleShapeSelect(section = 0, block = 0) {
    clearCellSelection();
    shapeSelection.value = { section, block };
    void nextTick().then(remeasureShapeHandles);
  }

  function shapeHandleStyle(handle = "se"): Record<string, string> {
    const geometry = shapeHandles.value;
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

  /** The rotation handle sits centred above the shape's top edge (AC-10). */
  function rotationHandleStyle(): Record<string, string> {
    const geometry = shapeHandles.value;
    if (!geometry) {
      return {};
    }
    return {
      left: `${geometry.left + geometry.width / 2 - HANDLE_SIZE / 2}px`,
      top: `${geometry.top - ROTATION_HANDLE_OFFSET_PX - HANDLE_SIZE / 2}px`,
    };
  }

  function shapeToolbarStyle(): { left: number; top: number } {
    const geometry = shapeHandles.value;
    const pageEl = pageRef.value;
    if (!geometry || !pageEl) {
      return { left: 0, top: 0 };
    }
    const pageRect = pageEl.getBoundingClientRect();
    const viewportTop = pageRect.top + geometry.top;
    const above = viewportTop - 44;
    return {
      left: pageRect.left + geometry.left,
      top: above >= 0 ? above : pageRect.top + geometry.top + geometry.height + 8,
    };
  }

  function applyShapeEdit(mutate: ShapeMutate | null, label = "") {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const location = shapeSelection.value;
    if (!mutate || !model || !root || !location) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = mutate(model, selection, location);
    applyEdit(result, selection, label);
  }

  function handleShapeDelete() {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => deleteShape(model, selection, location), "delete shape");
    clearShapeSelection();
  }

  function handleShapeDuplicate() {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = duplicateShape(model, selection, location);
    applyEdit(result, selection, "duplicate shape");
    shapeSelection.value = { section: location.section, block: location.block + 1 };
    void nextTick().then(remeasureShapeHandles);
  }

  function handleShapeNudge(direction: NudgeDirection, fine: boolean) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => nudgeShape(model, selection, location, direction, fine), "nudge shape");
    void nextTick().then(remeasureShapeHandles);
  }

  function handleShapeFill(fill: string | undefined) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => setShapeFill(model, selection, location, fill), "shape fill");
  }

  function handleShapeOutline(outline: ShapeOutline | undefined) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => setShapeOutline(model, selection, location, outline), "shape outline");
  }

  function handleShapeRotationInput(deg: number) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => setShapeRotation(model, selection, location, deg), "rotate shape");
    void nextTick().then(remeasureShapeHandles);
  }

  function handleShapeWrap(wrap: string) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => setShapeWrap(model, selection, location, asImageWrapMode(wrap)), "shape wrap");
  }

  function handleShapeAltText(altText: string) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => setShapeAltText(model, selection, location, altText), "shape alt text");
  }

  function handleShapeZOrder(command: ZOrderCommand) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit((model, selection) => reorderShapeZ(model, selection, location, command), "reorder shape");
  }

  /** Groups a list of selected drawing locations (AC-34..36) — called by
   * the multi-select UI (Shift-click) with every selected location. */
  function handleGroupSelected(locations: ShapeLocation[]) {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root || locations.length < 2) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = groupShapes(model, selection, locations);
    applyEdit(result, selection, "group shapes");
  }

  function handleUngroupSelected() {
    const location = shapeSelection.value;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!location || !model || !root) {
      return;
    }
    const shape = shapeAt(model, location);
    if (shape?.groupOrdinal === undefined) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = ungroupShapes(model, selection, shape.groupOrdinal);
    applyEdit(result, selection, "ungroup shapes");
  }

  function currentGroupMembers(): ShapeLocation[] {
    const location = shapeSelection.value;
    const model = doc.value;
    if (!location || !model) {
      return [];
    }
    const shape = shapeAt(model, location);
    if (shape?.groupOrdinal === undefined) {
      return [location];
    }
    return groupMembers(model, shape.groupOrdinal).map((m) => m.location);
  }

  // --- resize handles (AC-10..16) --------------------------------------

  function startShapeResize(event: MouseEvent, handle = "se") {
    const location = shapeSelection.value;
    const model = doc.value;
    const root = contentRoot.value;
    if (!location || !model || !root || !editable.value) {
      return;
    }
    const shape = shapeAt(model, location);
    const el = root.querySelector(
      `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
    );
    if (!shape || !(el instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    resizeDragMoved = false;
    resizeDrag = {
      location,
      handle,
      startClientX: event.clientX,
      startClientY: event.clientY,
      start: { widthCm: shape.widthCm, heightCm: shape.heightCm },
      el,
    };
    window.addEventListener("mousemove", handleShapeResizeMove);
    window.addEventListener("mouseup", finishShapeResize);
  }

  function handleShapeResizeMove(event: MouseEvent) {
    const drag = resizeDrag;
    if (!drag) {
      return;
    }
    resizeDragMoved = true;
    const deltaXCm = (event.clientX - drag.startClientX) * CM_PER_PX;
    const deltaYCm = (event.clientY - drag.startClientY) * CM_PER_PX;
    const size = resizedImageSize(drag.start, asImageHandle(drag.handle), deltaXCm, deltaYCm, event.shiftKey);
    previewImageSize(drag.el, size.widthCm, size.heightCm);
  }

  function finishShapeResize() {
    window.removeEventListener("mousemove", handleShapeResizeMove);
    window.removeEventListener("mouseup", finishShapeResize);
    const drag = resizeDrag;
    resizeDrag = null;
    if (!drag || !resizeDragMoved) {
      return;
    }
    const widthCm = parseFloat(drag.el.style.width) || drag.start.widthCm;
    const heightCm = parseFloat(drag.el.style.height) || drag.start.heightCm;
    applyShapeEdit(
      (model, selection) => setShapeSize(model, selection, drag.location, widthCm, heightCm),
      "resize shape",
    );
    void nextTick().then(remeasureShapeHandles);
  }

  // --- rotation handle (AC-16) -------------------------------------------

  function startShapeRotate(event: MouseEvent) {
    const location = shapeSelection.value;
    const model = doc.value;
    const geometry = shapeHandles.value;
    const pageEl = pageRef.value;
    if (!location || !model || !geometry || !pageEl || !editable.value) {
      return;
    }
    const shape = shapeAt(model, location);
    if (!shape) {
      return;
    }
    event.preventDefault();
    const pageRect = pageEl.getBoundingClientRect();
    const centerClientX = pageRect.left + geometry.left + geometry.width / 2;
    const centerClientY = pageRect.top + geometry.top + geometry.height / 2;
    const startAngle = Math.atan2(event.clientY - centerClientY, event.clientX - centerClientX);
    rotateDrag = { location, centerClientX, centerClientY, startRotation: shape.rotationDeg ?? 0, startAngle };
    window.addEventListener("mousemove", handleShapeRotateMove);
    window.addEventListener("mouseup", finishShapeRotate);
  }

  function handleShapeRotateMove(event: MouseEvent) {
    const drag = rotateDrag;
    if (!drag) {
      return;
    }
    const angle = Math.atan2(event.clientY - drag.centerClientY, event.clientX - drag.centerClientX);
    const deltaDeg = ((angle - drag.startAngle) * 180) / Math.PI;
    const next = normalizeRotationDeg(drag.startRotation + deltaDeg);
    applyShapeEdit(
      (model, selection) => setShapeRotation(model, selection, drag.location, next),
      "rotate shape",
    );
  }

  function finishShapeRotate() {
    window.removeEventListener("mousemove", handleShapeRotateMove);
    window.removeEventListener("mouseup", finishShapeRotate);
    rotateDrag = null;
    void nextTick().then(remeasureShapeHandles);
  }

  // --- draw mode: gallery pick -> drag/click insertion (AC-3..9) --------

  function armShapeDraw(geometry: ShapeGeometry | "textbox") {
    drawGeometry.value = geometry;
  }

  function disarmShapeDraw() {
    drawGeometry.value = null;
  }

  /** Finds the anchor CaretPosition + its page-relative top-left, in cm
   * (AC-4..6): the current caret's own paragraph, matching insertShape's
   * own caret-anchored insertion. */
  function currentAnchor(): { anchor: CaretPosition; left: number; top: number } | null {
    const root = contentRoot.value as HTMLElement | null;
    const pageEl = pageRef.value;
    if (!root || !pageEl) {
      return null;
    }
    const anchor = getCaret() ?? documentStart();
    const el = root.querySelector(`[data-section="${anchor.section}"][data-block="${anchor.block}"]`);
    if (!el) {
      return null;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    return { anchor, left: (rect.left - pageRect.left) * CM_PER_PX, top: (rect.top - pageRect.top) * CM_PER_PX };
  }

  function startShapeDraw(event: MouseEvent) {
    if (!drawGeometry.value || !editable.value) {
      return;
    }
    const pageEl = pageRef.value;
    const anchorInfo = currentAnchor();
    if (!pageEl || !anchorInfo) {
      return;
    }
    event.preventDefault();
    drawDrag = {
      startClientX: event.clientX,
      startClientY: event.clientY,
      anchor: anchorInfo.anchor,
      anchorLeft: anchorInfo.left,
      anchorTop: anchorInfo.top,
    };
    window.addEventListener("mouseup", finishShapeDraw);
  }

  function finishShapeDraw(event: MouseEvent) {
    window.removeEventListener("mouseup", finishShapeDraw);
    const drag = drawDrag;
    const geometry = drawGeometry.value;
    drawDrag = null;
    if (!drag || !geometry) {
      return;
    }
    const pageEl = pageRef.value;
    if (!pageEl) {
      return;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const dxPx = event.clientX - drag.startClientX;
    const dyPx = event.clientY - drag.startClientY;
    const dxCm = dxPx * CM_PER_PX;
    const dyCm = dyPx * CM_PER_PX;
    const dragged = Math.abs(dxPx) > 3 || Math.abs(dyPx) > 3;

    const isTextBox = geometry === "textbox";
    const realGeometry: ShapeGeometry = isTextBox ? "rect" : (geometry as ShapeGeometry);
    let widthCm: number;
    let heightCm: number;
    let leftCm: number;
    let topCm: number;
    if (dragged && !isTextBox) {
      const size = event.shiftKey ? constrainedDragSizeCm(realGeometry, dxCm, dyCm) : { widthCm: Math.abs(dxCm), heightCm: Math.abs(dyCm) };
      widthCm = size.widthCm;
      heightCm = size.heightCm;
      leftCm = (dxPx < 0 ? event.clientX - pageRect.left : drag.startClientX - pageRect.left) * CM_PER_PX;
      topCm = (dyPx < 0 ? event.clientY - pageRect.top : drag.startClientY - pageRect.top) * CM_PER_PX;
    } else {
      const size = defaultClickSizeCm(realGeometry);
      widthCm = size.widthCm;
      heightCm = size.heightCm;
      leftCm = (event.clientX - pageRect.left) * CM_PER_PX;
      topCm = (event.clientY - pageRect.top) * CM_PER_PX;
    }

    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = caretSelection(drag.anchor);
    const offsetXCm = leftCm - drag.anchorLeft;
    const offsetYCm = topCm - drag.anchorTop;
    const result = isTextBox
      ? insertTextBox(model, selection)
      : insertShape(model, selection, {
          geometry: realGeometry,
          widthCm,
          heightCm,
          offsetXCm,
          offsetYCm,
          outline: DEFAULT_NEW_SHAPE_OUTLINE,
          ...(LINE_GEOMETRIES.has(realGeometry) ? {} : { fill: DEFAULT_NEW_SHAPE_FILL }),
        });
    applyEdit(result, selection, isTextBox ? "insert text box" : "insert shape");
    disarmShapeDraw();
    const inserted = shapeInsertedLocation(result.model, drag.anchor);
    if (inserted) {
      shapeSelection.value = inserted;
      void nextTick().then(remeasureShapeHandles);
    }
  }

  /** The shape insertShape/insertTextBox just landed, found by scanning
   * near the anchor block for the freshest shape — insertShape's own
   * return selection points just after it (see its landing CaretPosition). */
  function shapeInsertedLocation(model: DocumentModel, anchor: CaretPosition): ShapeLocation | null {
    const blocks = model.sections[anchor.section]?.blocks ?? [];
    for (let i = 0; i < blocks.length; i++) {
      if (isShape(blocks[i])) {
        // Heuristic: the most recently inserted shape is adjacent to the
        // anchor block index (insertShape splits/lands right there).
        if (Math.abs(i - anchor.block) <= 2) {
          return { section: anchor.section, block: i };
        }
      }
    }
    return null;
  }

  function insertTextBoxAction() {
    armShapeDraw("textbox");
    // A text box has a fixed default size, so it inserts immediately at the
    // caret rather than waiting for a draw gesture (AC-9).
    const anchorInfo = currentAnchor();
    if (!anchorInfo) {
      disarmShapeDraw();
      return;
    }
    const model = doc.value;
    if (!model) {
      disarmShapeDraw();
      return;
    }
    const selection = caretSelection(anchorInfo.anchor);
    const result = insertTextBox(model, selection);
    applyEdit(result, selection, "insert text box");
    disarmShapeDraw();
    const inserted = shapeInsertedLocation(result.model, anchorInfo.anchor);
    if (inserted) {
      shapeSelection.value = inserted;
      void nextTick().then(remeasureShapeHandles);
    }
  }

  return {
    shapeSelection,
    shapeHandles,
    drawGeometry,
    editingShapeText,
    clearShapeSelection,
    handleShapeSelect,
    handleShapeEditText,
    shapeHandleStyle,
    rotationHandleStyle,
    shapeToolbarStyle,
    remeasureShapeHandles,
    handleShapeDelete,
    handleShapeDuplicate,
    handleShapeNudge,
    handleShapeFill,
    handleShapeOutline,
    handleShapeRotationInput,
    handleShapeWrap,
    handleShapeAltText,
    handleShapeZOrder,
    handleGroupSelected,
    handleUngroupSelected,
    currentGroupMembers,
    startShapeResize,
    handleShapeResizeMove,
    finishShapeResize,
    startShapeRotate,
    handleShapeRotateMove,
    finishShapeRotate,
    armShapeDraw,
    disarmShapeDraw,
    startShapeDraw,
    finishShapeDraw,
    insertTextBoxAction,
    freshGroupOrdinal: () => (doc.value ? freshGroupOrdinal(doc.value) : 1),
    moveGroup,
    resizeGroup,
    rotateGroup,
  };
}
