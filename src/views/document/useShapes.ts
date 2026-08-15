// Shapes and text boxes (SPEC-051): the selected shape's state (selection,
// measured handles), draw-mode insertion (gallery pick -> drag/click on the
// canvas), the shape toolbar's actions, the eight-handle resize drag, the
// rotation-handle drag, keyboard nudge/duplicate, and z-order/grouping.
// Mirrors useImages.ts's shape closely — see its own header comment for the
// pattern this follows.
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import {
  caretSelection,
  commitShapeModel,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentStart,
  insertLineBreak,
  insertText,
  isCollapsed,
  noPendingFormat,
  noSelection,
  splitBlock,
  withShapeModel,
  type ModelSelection,
} from "../../lib/commands";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import { asImageHandle, previewImageSize, resizedImageSize } from "../../lib/imageResize";
import { cmToPx } from "../../lib/images";
import { paintShapeIntoCanvas } from "../../lib/shapeCanvas";
import { CM_PER_PX } from "../../lib/tableResize";
import { modelRangeFromSelection, setDomCaret, setDomSelection } from "../../lib/domCaret";
import { blockLength, noCaret } from "../../lib/textEditing";
import { recordEdit as recordHistoryEdit } from "../../lib/undoHistory";
import { replaceDocumentModel } from "../../lib/documentStore";
import {
  DEFAULT_DRAWING_POSITION,
  DEFAULT_NEW_SHAPE_FILL,
  DEFAULT_NEW_SHAPE_OUTLINE,
  deleteShape,
  drawGestureBox,
  duplicateShape,
  ensureShapeTextFrame,
  freshGroupOrdinal,
  groupBoundingBoxCm,
  groupMembers,
  groupShapes,
  insertShape,
  insertTextBox,
  moveGroup,
  moveShape,
  nudgeShape,
  normalizeRotationDeg,
  reanchorShape,
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
  shapeHasVisibleText,
  shapesInPaintOrder,
  ungroupShapes,
  type NudgeDirection,
  type ParagraphSpan,
  type ShapeLocation,
  type ZOrderCommand,
} from "../../lib/shapes";
import {
  HIT_TOLERANCE_CM,
  pointHitsShape,
  rotatePoint,
  rotatedResizeCursor,
  rotatedBounds,
  unrotateDelta,
} from "../../lib/shapeGeometry";
import {
  asImageWrapMode,
  isShape,
  LINE_GEOMETRIES,
  type DocumentModel,
  type Shape,
  type ShapeGeometry,
  type ShapeOutline,
} from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { CommandResult } from "../../lib/commands";

function sameShapeLocation(a: ShapeLocation, b: ShapeLocation): boolean {
  return a.section === b.section && a.block === b.block;
}

const HANDLE_SIZE = 8;
const ROTATION_HANDLE_OFFSET_PX = 24;

/** Gap the floating toolbar keeps above the shape's top edge (AC-12): enough
 * to clear the rotation handle, which hangs ROTATION_HANDLE_OFFSET_PX above
 * that edge and is HANDLE_SIZE tall, so the handle stays visible and
 * clickable with the toolbar open (AC-16). */
const SHAPE_TOOLBAR_CLEARANCE_PX = ROTATION_HANDLE_OFFSET_PX + HANDLE_SIZE / 2 + 6;

/** Upper bound on the wrapped toolbar's own height — three rows of controls
 * at its max-width, i.e. more than it ever needs. Only the "is there room
 * above?" decision uses it (placing above needs no height at all, see
 * shapeToolbarStyle), so erring high merely flips a shape very near the top
 * of the viewport to a below placement a little sooner. */
const SHAPE_TOOLBAR_MAX_HEIGHT_PX = 96;

type ShapeMutate = (model: DocumentModel, selection: ModelSelection, location: ShapeLocation) => CommandResult;

export interface ShapesDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  pageRef: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  documentId: Readonly<Ref<string>>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
  ) => void;
  clearCellSelection: () => void;
  getCaret: () => CaretPosition | null;
  /** The shape text-editing pipeline's own commit path (applyShapeTextEdit)
   * bypasses the generic `applyEdit` above — like useNotes.ts's
   * applyNoteEdit, it must fold a synthetic withShapeModel result back into
   * the real document (commitShapeModel) before it can go through the
   * normal history/store pipeline, so it needs these directly. */
  markDirty: () => void;
  refreshUndoRedoState: () => void;
}

export function useShapes(deps: ShapesDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    documentId,
    currentSelection,
    applyEdit,
    clearCellSelection,
    getCaret,
    markDirty,
    refreshUndoRedoState,
  } = deps;

  const shapeSelection = ref<ShapeLocation | null>(null);
  /** Shift-click multi-selection (AC-34): empty for an ordinary single
   * selection, 2+ entries while a multi-object selection is active —
   * `shapeSelection` tracks the most recently (de)selected member so the
   * handles/toolbar still attach somewhere sensible. Shape-only (an Image
   * has its own separate selection ref); grouping across the two kinds is
   * out of scope for this gesture. */
  const shapeMultiSelection = ref<ShapeLocation[]>([]);
  const shapeHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  /** The anchor marker's position (AC-27) — page-relative px, same basis as
   * shapeHandles (measureImageGeometry). Null when nothing is selected, or
   * the selected shape's preceding block isn't a paragraph to mark (see
   * measureShapeAnchorMarker). */
  const shapeAnchorMarker = ref<{ left: number; top: number } | null>(null);
  /** A shape's text frame is a second live-editable surface (AC-17, AC-18),
   * the same one-live-surface pattern useNotes.ts follows for note bodies:
   * its own caret/selection-anchor/pending-format/composition state,
   * addressed through withShapeModel's synthetic single-section coordinates
   * (section 0) and folded back by applyShapeTextEdit (commitShapeModel). */
  const editingShapeText = ref(false);
  const shapeTextRemountEpoch = ref(0);
  let shapeTextCaret: CaretPosition | null = noCaret();
  let shapeTextSelectionAnchor: CaretPosition | null = noCaret();
  let shapeTextPendingFormat: ReturnType<typeof noPendingFormat> = null;
  let shapeTextComposing = false;
  let shapeTextCompositionRange: ModelSelection | null = noSelection();

  /** The one DOM element currently rendering as a shape's editable text
   * frame — DocShapeBlock.vue only sets data-shape-text-editable on the
   * shape whose entry carries shapeEditing (usePaginationView.ts), so at
   * most one exists at a time. */
  function shapeTextRootEl(): HTMLElement | null {
    const found = contentRoot.value?.querySelector("[data-shape-text-editable]");
    return found instanceof HTMLElement ? found : null;
  }

  function currentShapeTextSelection(): ModelSelection | null {
    const root = shapeTextRootEl();
    if (root) {
      const range = modelRangeFromSelection(root, document.getSelection());
      if (range) {
        return { anchor: range.start, focus: range.end };
      }
    }
    return shapeTextCaret ? { anchor: shapeTextSelectionAnchor ?? shapeTextCaret, focus: shapeTextCaret } : null;
  }

  function resetShapeTextEditingState(): void {
    editingShapeText.value = false;
    shapeTextCaret = noCaret();
    shapeTextSelectionAnchor = noCaret();
    shapeTextPendingFormat = noPendingFormat();
  }

  /** Leaves shape text editing (Escape — AC-17's sibling exit gesture),
   * keeping the shape itself selected. */
  function exitShapeTextEditing(): void {
    if (!editingShapeText.value) {
      return;
    }
    resetShapeTextEditingState();
    void nextTick().then(() => (contentRoot.value as HTMLElement | null)?.focus());
  }

  /** Double-clicking a shape (AC-17): puts the caret inside its text frame,
   * populating Shape.text first if the shape had none. A second dblclick on
   * the SAME already-editing shape is left alone — it is the browser's own
   * dblclick-selects-a-word gesture inside the now-live contenteditable
   * region, not a re-entry. */
  function handleShapeEditText(section = 0, block = 0) {
    const location: ShapeLocation = { section, block };
    if (editingShapeText.value && shapeSelection.value && sameShapeLocation(shapeSelection.value, location)) {
      return;
    }
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    clearCellSelection();
    shapeMultiSelection.value = [];
    shapeSelection.value = location;
    const shape = shapeAt(model, location);
    if (shape && !shape.text) {
      const root = contentRoot.value as HTMLElement | null;
      const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
      const result = ensureShapeTextFrame(model, selection, location);
      applyEdit(result, selection, "add shape text");
    }
    editingShapeText.value = true;
    const current = doc.value;
    const text = current ? shapeAt(current, location)?.text : undefined;
    const lastBlock = text ? Math.max(0, text.blocks.length - 1) : 0;
    const lastLength = text ? blockLength(text.blocks[lastBlock]) : 0;
    shapeTextCaret = { section: 0, block: lastBlock, offset: lastLength };
    shapeTextSelectionAnchor = noCaret();
    shapeTextPendingFormat = noPendingFormat();
    void nextTick().then(() => {
      remeasureShapeHandles();
      const root = shapeTextRootEl();
      if (root) {
        root.focus();
        setDomCaret(root, shapeTextCaret ?? { section: 0, block: 0, offset: 0 });
      }
    });
  }

  /** Folds a `withShapeModel` command result back into the real document and
   * commits it (applyNoteEdit's sibling — see useNotes.ts). */
  function applyShapeTextEdit(resultModel: DocumentModel, resultSelection: ModelSelection, label = ""): void {
    const location = shapeSelection.value;
    const before = doc.value;
    if (!location || !before || !resultSelection) {
      return;
    }
    const nextModel = commitShapeModel(before, location, resultModel);
    const bodyFiller = caretSelection(getCaret() ?? documentStart());
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: bodyFiller },
      after: { model: nextModel, selection: bodyFiller },
      label,
      groupable: false,
    });
    doc.value = replaceDocumentModel(documentId.value, nextModel);
    shapeTextCaret = resultSelection.focus;
    shapeTextSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
    markDirty();
    refreshUndoRedoState();
    void nextTick().then(() => {
      const root = shapeTextRootEl();
      if (!root || !shapeTextCaret) {
        return;
      }
      if (shapeTextSelectionAnchor) {
        setDomSelection(root, shapeTextSelectionAnchor, shapeTextCaret);
      } else {
        setDomCaret(root, shapeTextCaret);
      }
    });
  }

  function handleShapeTextBeforeInput(event: InputEvent) {
    const type = event.inputType;
    if (shapeTextComposing || event.isComposing || type === "insertCompositionText") {
      return;
    }
    event.preventDefault();
    const location = shapeSelection.value;
    const model = doc.value;
    if (!location || !model || !editable.value) {
      return;
    }
    const selection = currentShapeTextSelection();
    if (!selection) {
      return;
    }
    const synthetic = withShapeModel(model, location);
    let result: { model: DocumentModel; selection: ModelSelection } | null = null;
    let label = "";
    if (type === "insertText" || type === "insertFromPaste") {
      const text = event.data ?? (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
      if (text) {
        result = insertText(synthetic, selection, text, shapeTextPendingFormat);
        label = "typing";
      }
    } else if (type === "insertParagraph") {
      result = splitBlock(synthetic, selection);
      label = "new paragraph";
    } else if (type === "insertLineBreak") {
      result = insertLineBreak(synthetic, selection);
      label = "line break";
    } else if (type === "deleteContentBackward") {
      result = deleteBackward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteContentForward") {
      result = deleteForward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteByCut" || type === "deleteContent") {
      result = isCollapsed(selection) ? null : deleteRange(synthetic, selection);
      label = "delete";
    }
    if (result) {
      shapeTextPendingFormat = noPendingFormat();
      applyShapeTextEdit(result.model, result.selection, label);
    }
  }

  function handleShapeTextKeyDown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      event.stopPropagation();
      exitShapeTextEditing();
    }
  }

  function handleShapeTextCompositionStart() {
    if (!shapeTextRootEl() || !editable.value) {
      return;
    }
    shapeTextComposing = true;
    shapeTextCompositionRange = currentShapeTextSelection();
  }

  function handleShapeTextCompositionEnd(event: CompositionEvent) {
    shapeTextComposing = false;
    const range = shapeTextCompositionRange;
    shapeTextCompositionRange = null;
    const location = shapeSelection.value;
    const model = doc.value;
    if (!location || !model || !editable.value || !range) {
      return;
    }
    const data = event.data ?? "";
    if (!data) {
      return;
    }
    const synthetic = withShapeModel(model, location);
    const result = insertText(synthetic, range, data, null);
    applyShapeTextEdit(result.model, result.selection, "typing");
    shapeTextRemountEpoch.value += 1;
  }
  /** Armed by picking a gallery entry (AC-3): the geometry the next
   * drag/click on the canvas draws, or "textbox" — cleared once drawn. */
  const drawGeometry = ref<ShapeGeometry | "textbox" | null>(null);
  /** The rubber band an in-flight draw gesture is describing (AC-3): the
   * shape about to be inserted, plus its box in page-relative px for the
   * overlay to sit at. Null whenever no draw gesture is in flight. */
  const drawPreview = ref<{
    shape: Shape;
    left: number;
    top: number;
    width: number;
    height: number;
  } | null>(null);
  let resizeDrag: {
    location: ShapeLocation;
    handle: string;
    startClientX: number;
    startClientY: number;
    start: { widthCm: number; heightCm: number };
    el: HTMLElement;
    groupOrdinal: number | undefined;
    rotationDeg: number;
    currentWidthCm: number;
    currentHeightCm: number;
  } | null = null;
  let resizeDragMoved = false;
  let rotateDrag: {
    location: ShapeLocation;
    centerClientX: number;
    centerClientY: number;
    startRotation: number;
    startAngle: number;
    groupOrdinal: number | undefined;
    appliedDeg: number;
  } | null = null;
  let drawDrag: { startClientX: number; startClientY: number; anchor: CaretPosition; anchorLeft: number; anchorTop: number } | null = null;

  function clearShapeSelection() {
    if (shapeSelection.value) {
      shapeSelection.value = null;
      shapeHandles.value = null;
      shapeAnchorMarker.value = null;
    }
    shapeMultiSelection.value = [];
    resetShapeTextEditingState();
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
    shapeAnchorMarker.value = measureShapeAnchorMarker(location, pageEl, root);
  }

  /** The paragraph a shape at `location` is anchored to, as a block index in
   * the same section (AC-27, AC-28) — by convention (reanchorShape,
   * duplicateShape) that is the block immediately before the shape's own in
   * the block list, since both always land the shape right after its
   * anchor. insertShape is the one exception: a shape inserted at the very
   * start of its target paragraph (caret offset 0) lands BEFORE it instead,
   * so a shape with nothing paragraph-shaped before it falls back to
   * checking immediately after. Null when neither neighbour is a paragraph
   * (e.g. the shape sits directly next to a table on both sides). */
  function anchorParagraphBlockIndex(location: ShapeLocation): number | null {
    const model = doc.value;
    const blocks = model?.sections[location.section]?.blocks;
    if (!blocks) {
      return null;
    }
    if (blocks[location.block - 1]?.type === "paragraph") {
      return location.block - 1;
    }
    if (blocks[location.block + 1]?.type === "paragraph") {
      return location.block + 1;
    }
    return null;
  }

  /** The DOM position of the paragraph a selected shape is anchored to
   * (AC-27); null when there is none to mark (anchorParagraphBlockIndex) or
   * it isn't measurable. */
  function measureShapeAnchorMarker(
    location: ShapeLocation,
    pageEl: Element,
    root: Element,
  ): { left: number; top: number } | null {
    const anchorBlockIndex = anchorParagraphBlockIndex(location);
    if (anchorBlockIndex === null) {
      return null;
    }
    const el = root.querySelector(`[data-section="${location.section}"][data-block="${anchorBlockIndex}"]`);
    if (!el) {
      return null;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    return { left: rect.left - pageRect.left, top: rect.top - pageRect.top };
  }

  /** Every paragraph in a shape's own section, as re-anchoring candidates
   * ("Re-anchoring — the nearest paragraph") — shapes are always top-level
   * blocks (ShapeLocation has no cell component), so every candidate is
   * automatically "at the same container level" as the current anchor; no
   * further filtering is needed. Positions are page-relative cm, the same
   * basis startShapeMove measures the dragged shape's own top in. */
  function candidateParagraphsFor(location: ShapeLocation): ParagraphSpan[] {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const pageEl = pageRef.value;
    if (!model || !root || !pageEl) {
      return [];
    }
    const pageRect = pageEl.getBoundingClientRect();
    const blocks = model.sections[location.section]?.blocks ?? [];
    const spans: ParagraphSpan[] = [];
    blocks.forEach((block, index) => {
      if (block.type !== "paragraph") {
        return;
      }
      const el = root.querySelector(`[data-section="${location.section}"][data-block="${index}"]`);
      if (!el) {
        return;
      }
      const rect = el.getBoundingClientRect();
      spans.push({
        section: location.section,
        block: index,
        topCm: (rect.top - pageRect.top) * CM_PER_PX,
        bottomCm: (rect.bottom - pageRect.top) * CM_PER_PX,
      });
    });
    return spans;
  }

  /** "Selection and hit testing": the topmost shape whose *painted* region
   * (never its bounding box) contains a client-space point — DocumentView's
   * recovery path for a click that DocShapeBlock's own proxy couldn't catch
   * natively (an unfilled, textless shape goes `pointer-events: none` over
   * its whole box, see DocShapeBlock.vue's proxyStyle) and for a shape whose
   * only clickable region is its outline. Iterates every shape topmost-first
   * (shapesInPaintOrder) and stops at the first true pointHitsShape. */
  function shapeAtClientPoint(clientX: number, clientY: number): ShapeLocation | null {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return null;
    }
    for (const { location, shape } of shapesInPaintOrder(model)) {
      const el = root.querySelector(
        `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
      );
      if (!(el instanceof HTMLElement)) {
        continue;
      }
      const rect = el.getBoundingClientRect();
      const point = {
        x: (clientX - rect.left) * CM_PER_PX,
        y: (clientY - rect.top) * CM_PER_PX,
      };
      const hit = pointHitsShape(point, {
        geometry: shape.geometry,
        widthCm: shape.widthCm,
        heightCm: shape.heightCm,
        rotationDeg: shape.rotationDeg,
        filled: !!shape.fill,
        hasText: shapeHasVisibleText(shape),
        outlineToleranceCm: HIT_TOLERANCE_CM,
      });
      if (hit) {
        return location;
      }
    }
    return null;
  }

  /** A plain click replaces the selection with just this shape; a Shift
   * click toggles it into (or out of) a multi-object selection (AC-34) built
   * from whatever was already selected — starting from a single shape grows
   * shapeMultiSelection to that shape plus this one. Dropping back to one
   * (or zero) members collapses out of multi-select mode. */
  function handleShapeSelect(section = 0, block = 0, shiftKey = false) {
    clearCellSelection();
    const location: ShapeLocation = { section, block };
    // A click within the shape currently being edited (repositioning the
    // caret in its contenteditable text) must not interrupt editing; only
    // selecting something else does (AC-17's "double-click to edit" implies
    // a plain click elsewhere leaves it).
    if (editingShapeText.value && !(shapeSelection.value && sameShapeLocation(shapeSelection.value, location))) {
      resetShapeTextEditingState();
    }
    if (shiftKey && shapeSelection.value) {
      const current = shapeMultiSelection.value.length ? shapeMultiSelection.value : [shapeSelection.value];
      const index = current.findIndex((loc) => sameShapeLocation(loc, location));
      const next = index >= 0 ? current.filter((_, i) => i !== index) : [...current, location];
      shapeMultiSelection.value = next.length > 1 ? next : [];
      shapeSelection.value = next.length > 0 ? next[next.length - 1] : null;
    } else {
      shapeMultiSelection.value = [];
      shapeSelection.value = location;
    }
    void nextTick().then(remeasureShapeHandles);
  }

  /** Whether the current selection is eligible for "Group" (AC-34) — a
   * Shift-click multi-selection of two or more shapes, as opposed to a
   * single shape that merely already belongs to an existing group (that
   * shape offers Ungroup instead, via the `grouped` prop). */
  function canGroupSelection(): boolean {
    return shapeMultiSelection.value.length >= 2;
  }

  /** The rotation the selected shape is painted at. The handles are drawn
   * against the shape's UNROTATED box (the DOM proxy the geometry is
   * measured from never rotates — only the canvas the shape is painted on
   * does), so every handle position has to apply this rotation itself. */
  function selectedShapeRotationDeg(): number {
    const model = doc.value;
    const location = shapeSelection.value;
    const shape = model && location ? shapeAt(model, location) : null;
    return shape?.rotationDeg ?? 0;
  }

  /** A resize handle sits on the shape's own corner or edge midpoint, riding
   * round with it as it rotates (AC-10, AC-13) rather than staying on the
   * upright box the shape would occupy at 0°. */
  function shapeHandleStyle(handle = "se"): Record<string, string> {
    const geometry = shapeHandles.value;
    if (!geometry) {
      return {};
    }
    const xFrac = handle.includes("w") ? 0 : handle.includes("e") ? 1 : 0.5;
    const yFrac = handle.includes("n") ? 0 : handle.includes("s") ? 1 : 0.5;
    const rotationDeg = selectedShapeRotationDeg();
    const point = rotatePoint(
      { x: geometry.width * xFrac, y: geometry.height * yFrac },
      geometry.width,
      geometry.height,
      rotationDeg,
    );
    return {
      left: `${geometry.left + point.x - HANDLE_SIZE / 2}px`,
      top: `${geometry.top + point.y - HANDLE_SIZE / 2}px`,
      cursor: rotatedResizeCursor(handle, rotationDeg),
    };
  }

  /** The rotation handle sits off the middle of the shape's own top edge
   * (AC-10), so it too swings round with the shape — at 180° it hangs below
   * the shape, which is where that edge now is. */
  function rotationHandleStyle(): Record<string, string> {
    const geometry = shapeHandles.value;
    if (!geometry) {
      return {};
    }
    const point = rotatePoint(
      { x: geometry.width / 2, y: -ROTATION_HANDLE_OFFSET_PX },
      geometry.width,
      geometry.height,
      selectedShapeRotationDeg(),
    );
    return {
      left: `${geometry.left + point.x - HANDLE_SIZE / 2}px`,
      top: `${geometry.top + point.y - HANDLE_SIZE / 2}px`,
    };
  }

  /** Where the floating toolbar sits relative to the selected shape (AC-12).
   * `placement: "above"` means `top` is the line the toolbar's BOTTOM edge
   * rests on — ShapeToolbar.vue pulls itself up by its own height from
   * there. Anchoring the bottom edge rather than the top is what keeps the
   * toolbar off the shape: it wraps to two or three rows depending on width
   * and on which of Group/Ungroup apply, so no fixed height offset (this
   * used to subtract a flat 44px) can clear a shape whose toolbar is
   * taller than that.
   *
   * A rotated shape reaches past its own box on every side — and so do its
   * handles, which now ride round with it (AC-13) — so the clearance is
   * measured from the rotated extent, not from the upright box's edges. */
  function shapeToolbarStyle(): { left: number; top: number; placement: string } {
    const geometry = shapeHandles.value;
    const pageEl = pageRef.value;
    if (!geometry || !pageEl) {
      return { left: 0, top: 0, placement: "below" };
    }
    const bounds = rotatedBounds(geometry.width, geometry.height, selectedShapeRotationDeg());
    const pageRect = pageEl.getBoundingClientRect();
    const shapeTop = pageRect.top + geometry.top - bounds.padY;
    const bottomEdge = shapeTop - SHAPE_TOOLBAR_CLEARANCE_PX;
    const left = pageRect.left + geometry.left - bounds.padX;
    if (bottomEdge - SHAPE_TOOLBAR_MAX_HEIGHT_PX >= 0) {
      return { left, top: bottomEdge, placement: "above" };
    }
    return { left, top: pageRect.top + geometry.top + geometry.height + bounds.padY + 8, placement: "below" };
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
   * the multi-select UI (Shift-click) with every selected location. Leaves
   * the group's first member selected (single-select, not multi — every
   * member now shares a groupOrdinal, so the toolbar's `grouped` prop offers
   * Ungroup instead). */
  function handleGroupSelected(locations: ShapeLocation[]) {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root || locations.length < 2) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = groupShapes(model, selection, locations);
    applyEdit(result, selection, "group shapes");
    shapeMultiSelection.value = [];
    shapeSelection.value = locations[0] ?? null;
    void nextTick().then(remeasureShapeHandles);
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
    // A live Shift-click multi-selection (AC-34..36) is the Group action's
    // target set whenever one is active — takes priority over a single
    // shape's own existing group (that shape offers Ungroup instead).
    if (shapeMultiSelection.value.length >= 2) {
      return shapeMultiSelection.value;
    }
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
      groupOrdinal: shape.groupOrdinal,
      rotationDeg: shape.rotationDeg ?? 0,
      currentWidthCm: shape.widthCm,
      currentHeightCm: shape.heightCm,
    };
    window.addEventListener("mousemove", handleShapeResizeMove);
    window.addEventListener("mouseup", finishShapeResize);
  }

  /** Repaints a shape's own canvas at a size the model does not carry yet.
   * previewImageSize only stretches the wrapper element; on its own that
   * leaves a resize drag showing a box growing around a shape that stays
   * the size it started, so the drag gives no feel for what is being made.
   * Same painter and same sizing as the page's own render, so the stretched
   * preview is the shape, not a scaled bitmap of it. */
  function previewShapeCanvas(el: HTMLElement, shape: Shape, widthCm: number, heightCm: number) {
    const canvas = el.querySelector("canvas.doc-shape-canvas");
    if (!(canvas instanceof HTMLCanvasElement)) {
      return;
    }
    paintShapeIntoCanvas(
      canvas,
      { ...shape, widthCm, heightCm },
      Math.max(1, Math.round(cmToPx(widthCm))),
      Math.max(1, Math.round(cmToPx(Math.max(heightCm, 0.05)))),
    );
  }

  function handleShapeResizeMove(event: MouseEvent) {
    const drag = resizeDrag;
    if (!drag) {
      return;
    }
    resizeDragMoved = true;
    // widthCm/heightCm run along the shape's OWN axes, so a rotated shape's
    // handles are being dragged in a frame turned by the same angle: the
    // pointer delta has to come back out of it before it can size the box
    // (drag the "e" handle of a 90° shape and the pointer travels down the
    // screen, but what grows is still the width).
    const local = unrotateDelta(
      (event.clientX - drag.startClientX) * CM_PER_PX,
      (event.clientY - drag.startClientY) * CM_PER_PX,
      drag.rotationDeg,
    );
    const size = resizedImageSize(drag.start, asImageHandle(drag.handle), local.dx, local.dy, event.shiftKey);
    drag.currentWidthCm = size.widthCm;
    drag.currentHeightCm = size.heightCm;
    // A grouped member previews only on release, same as a grouped move
    // (finishShapeResize's resizeGroup scales every member about the
    // group's own bounding-box centre — live-previewing just the dragged
    // element here would show the rest of the group left behind).
    if (drag.groupOrdinal === undefined) {
      previewImageSize(drag.el, size.widthCm, size.heightCm);
      const shape = doc.value ? shapeAt(doc.value, drag.location) : null;
      if (shape) {
        previewShapeCanvas(drag.el, shape, size.widthCm, size.heightCm);
      }
      // The handles are what the pointer is dragging — they have to track
      // the box they are sizing, not sit at the size it started.
      remeasureShapeHandles();
    }
  }

  function finishShapeResize() {
    window.removeEventListener("mousemove", handleShapeResizeMove);
    window.removeEventListener("mouseup", finishShapeResize);
    const drag = resizeDrag;
    resizeDrag = null;
    if (!drag || !resizeDragMoved) {
      return;
    }
    if (drag.groupOrdinal !== undefined) {
      const model = doc.value;
      const root = contentRoot.value as HTMLElement | null;
      if (!model || !root) {
        return;
      }
      const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
      const sx = drag.start.widthCm > 0 ? drag.currentWidthCm / drag.start.widthCm : 1;
      const sy = drag.start.heightCm > 0 ? drag.currentHeightCm / drag.start.heightCm : 1;
      const bbox = groupBoundingBoxCm(model, drag.groupOrdinal);
      // Scale about the group's own bounding-box centre (AC-37): translate
      // that centre to the origin, scale from there (resizeGroup's own
      // formula scales offsets from whatever origin is current), then
      // translate back — the centre lands exactly where it started.
      let result: CommandResult;
      if (bbox) {
        const cx = bbox.leftCm + bbox.widthCm / 2;
        const cy = bbox.topCm + bbox.heightCm / 2;
        const centred = moveGroup(model, selection, drag.groupOrdinal, -cx, -cy).model;
        const scaled = resizeGroup(centred, selection, drag.groupOrdinal, sx, sy).model;
        result = moveGroup(scaled, selection, drag.groupOrdinal, cx, cy);
      } else {
        result = resizeGroup(model, selection, drag.groupOrdinal, sx, sy);
      }
      applyEdit(result, selection, "resize group");
      void nextTick().then(remeasureShapeHandles);
      return;
    }
    applyShapeEdit(
      (model, selection) => setShapeSize(model, selection, drag.location, drag.currentWidthCm, drag.currentHeightCm),
      "resize shape",
    );
    void nextTick().then(remeasureShapeHandles);
  }

  // --- drag-to-move / re-anchor (AC-28) -----------------------------------

  let moveDrag: {
    location: ShapeLocation;
    startClientX: number;
    startClientY: number;
    startOffsetXCm: number;
    startOffsetYCm: number;
    startTopCm: number;
    heightCm: number;
    groupOrdinal: number | undefined;
    el: HTMLElement;
    dxCm: number;
    dyCm: number;
    startHandles: ReturnType<typeof noImageHandleGeometry>;
  } | null = null;
  let moveDragMoved = false;
  /** True while a shape is actually being dragged (past the click
   * threshold) — the floating toolbar hides for the duration. */
  const movingShape = ref(false);

  /** Arms a possible move drag on mousedown over a shape's own body
   * (DocumentView's handleCanvasMouseDown, both for a native DOM hit and for
   * its shapeAtClientPoint recovery) — takes an explicit `location` rather
   * than reading shapeSelection, since for a shape that was not already
   * selected the native `click` that will select it has not fired yet at
   * mousedown time. Only commits a move on release past a small pixel
   * threshold (moveDragMoved), so a plain click-to-select is untouched. An
   * `inline` shape has no offset to drag (it sits in text flow); the caller
   * also skips this for a Shift-click (multi-select toggle, not a drag). */
  function startShapeMove(event: MouseEvent, location: ShapeLocation) {
    const model = doc.value;
    const root = contentRoot.value;
    const pageEl = pageRef.value;
    if (!model || !root || !pageEl || !editable.value || editingShapeText.value) {
      return;
    }
    const shape = shapeAt(model, location);
    const el = root.querySelector(
      `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
    );
    if (!shape || shape.wrap === "inline" || !(el instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    moveDragMoved = false;
    moveDrag = {
      location,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startOffsetXCm: shape.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm,
      startOffsetYCm: shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm,
      startTopCm: (rect.top - pageRect.top) * CM_PER_PX,
      heightCm: shape.heightCm,
      groupOrdinal: shape.groupOrdinal,
      el,
      dxCm: 0,
      dyCm: 0,
      startHandles: shapeHandles.value ? { ...shapeHandles.value } : null,
    };
    window.addEventListener("mousemove", handleShapeMoveMove);
    window.addEventListener("mouseup", finishShapeMove);
  }

  function handleShapeMoveMove(event: MouseEvent) {
    const drag = moveDrag;
    if (!drag) {
      return;
    }
    const dxPx = event.clientX - drag.startClientX;
    const dyPx = event.clientY - drag.startClientY;
    if (!moveDragMoved && (Math.abs(dxPx) > 3 || Math.abs(dyPx) > 3)) {
      moveDragMoved = true;
    }
    if (!moveDragMoved) {
      return;
    }
    drag.dxCm = dxPx * CM_PER_PX;
    drag.dyCm = dyPx * CM_PER_PX;
    // A grouped member previews only on release (finishShapeMove's
    // moveGroup moves every member at once) — live-previewing just the
    // dragged element here would show the rest of the group left behind.
    if (drag.groupOrdinal === undefined) {
      drag.el.style.left = `${drag.startOffsetXCm + drag.dxCm}cm`;
      drag.el.style.top = `${drag.startOffsetYCm + drag.dyCm}cm`;
      // The handles belong to the shape, so they travel with it — the
      // element moves by exactly this delta, so the measured box does too.
      if (drag.startHandles) {
        shapeHandles.value = {
          ...drag.startHandles,
          left: drag.startHandles.left + dxPx,
          top: drag.startHandles.top + dyPx,
        };
      }
      // The toolbar stays put wherever it was opened, which during a drag
      // means it sits across the page in the way of the shape being
      // dragged; it comes back on release, repositioned (finishShapeMove's
      // remeasureShapeHandles).
      movingShape.value = true;
    }
  }

  function finishShapeMove() {
    window.removeEventListener("mousemove", handleShapeMoveMove);
    window.removeEventListener("mouseup", finishShapeMove);
    const drag = moveDrag;
    moveDrag = null;
    movingShape.value = false;
    if (!drag || !moveDragMoved) {
      return;
    }
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    if (drag.groupOrdinal !== undefined) {
      const result = moveGroup(model, selection, drag.groupOrdinal, drag.dxCm, drag.dyCm);
      applyEdit(result, selection, "move group");
      void nextTick().then(remeasureShapeHandles);
      return;
    }
    // The raw drag delta lands the shape exactly where it was dropped,
    // relative to its CURRENT anchor — moveShape's own job, and the whole
    // effect for a drag that stays within the same anchor's territory.
    const moved = moveShape(model, selection, drag.location, drag.dxCm, drag.dyCm);
    // Re-anchor across a paragraph boundary on top of that (AC-28,
    // "Re-anchoring — the nearest paragraph"): reanchorShape re-expresses
    // the same dropped position relative to whichever paragraph is now
    // nearest the dragged object's vertical centre — page-relative cm, the
    // same basis candidateParagraphsFor's spans use.
    const candidates = candidateParagraphsFor(drag.location);
    const anchorBlockIndex = anchorParagraphBlockIndex(drag.location);
    const anchorSpan = anchorBlockIndex === null ? undefined : candidates.find((span) => span.block === anchorBlockIndex);
    let result: CommandResult = moved;
    let newLocation = drag.location;
    if (anchorSpan) {
      // candidates is non-empty (it contains anchorSpan) and every entry is
      // a real paragraph — never the shape's own block — so reanchorShape's
      // "nearest === the shape's own location" no-op guard can never trigger
      // here; its landing selection always points just past the shape's new
      // position (see its own doc comment).
      result = reanchorShape(
        moved.model,
        selection,
        drag.location,
        candidates,
        drag.startTopCm + drag.dyCm + drag.heightCm / 2,
        anchorSpan.topCm,
      );
      newLocation = { section: result.selection.focus.section, block: result.selection.focus.block - 1 };
    }
    applyEdit(result, selection, "move shape");
    shapeSelection.value = newLocation;
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
    rotateDrag = {
      location,
      centerClientX,
      centerClientY,
      startRotation: shape.rotationDeg ?? 0,
      startAngle,
      groupOrdinal: shape.groupOrdinal,
      appliedDeg: 0,
    };
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
    if (drag.groupOrdinal !== undefined) {
      const model = doc.value;
      const root = contentRoot.value as HTMLElement | null;
      if (!model || !root) {
        return;
      }
      const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
      // rotateGroup adds a delta to each member's CURRENT rotationDeg, so
      // each mousemove applies only the increment since the last one —
      // applying the full deltaDeg fresh every time would compound.
      const increment = deltaDeg - drag.appliedDeg;
      drag.appliedDeg = deltaDeg;
      const result = rotateGroup(model, selection, drag.groupOrdinal, increment);
      applyEdit(result, selection, "rotate group");
      return;
    }
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
    drawPreview.value = null;
  }

  /** The Shape a draw gesture will land. The preview paints this exact
   * object and the insertion takes its fill/outline from it, so what is
   * dragged out and what appears cannot disagree. A text box gets neither
   * fill nor outline (AC-9). */
  function newDrawnShape(
    geometry: ShapeGeometry,
    isTextBox: boolean,
    widthCm: number,
    heightCm: number,
  ): Shape {
    const shape: Shape = { type: "shape", geometry, widthCm, heightCm, wrap: "square" };
    if (isTextBox) {
      return shape;
    }
    shape.outline = DEFAULT_NEW_SHAPE_OUTLINE;
    if (!LINE_GEOMETRIES.has(geometry)) {
      shape.fill = DEFAULT_NEW_SHAPE_FILL;
    }
    return shape;
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

  /** The page-relative cm box the pointer currently describes, and the
   * geometry it will be drawn as — the single answer both the live preview
   * and the committed insertion read. */
  function drawBoxNow(
    drag: NonNullable<typeof drawDrag>,
    event: MouseEvent,
    geometry: ShapeGeometry | "textbox",
  ): { geometry: ShapeGeometry; isTextBox: boolean; box: ReturnType<typeof drawGestureBox> } | null {
    const pageEl = pageRef.value;
    if (!pageEl) {
      return null;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const isTextBox = geometry === "textbox";
    const realGeometry: ShapeGeometry = isTextBox ? "rect" : (geometry as ShapeGeometry);
    return {
      geometry: realGeometry,
      isTextBox,
      box: drawGestureBox(
        realGeometry,
        {
          xCm: (drag.startClientX - pageRect.left) * CM_PER_PX,
          yCm: (drag.startClientY - pageRect.top) * CM_PER_PX,
        },
        { xCm: (event.clientX - pageRect.left) * CM_PER_PX, yCm: (event.clientY - pageRect.top) * CM_PER_PX },
        { shiftKey: event.shiftKey, forceDefaultSize: isTextBox },
      ),
    };
  }

  /** Repaints the rubber band from wherever the pointer is now (AC-3). */
  function handleShapeDrawMove(event: MouseEvent) {
    const drag = drawDrag;
    const geometry = drawGeometry.value;
    if (!drag || !geometry) {
      return;
    }
    const now = drawBoxNow(drag, event, geometry);
    if (!now) {
      return;
    }
    drawPreview.value = {
      shape: newDrawnShape(now.geometry, now.isTextBox, now.box.widthCm, now.box.heightCm),
      left: cmToPx(now.box.leftCm),
      top: cmToPx(now.box.topCm),
      width: cmToPx(now.box.widthCm),
      height: cmToPx(now.box.heightCm),
    };
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
    // Preview from the press, not from the first move: before the pointer
    // passes the drag threshold the box shown is the default click size,
    // which is exactly what releasing here would insert.
    handleShapeDrawMove(event);
    window.addEventListener("mousemove", handleShapeDrawMove);
    window.addEventListener("mouseup", finishShapeDraw);
  }

  function finishShapeDraw(event: MouseEvent) {
    window.removeEventListener("mousemove", handleShapeDrawMove);
    window.removeEventListener("mouseup", finishShapeDraw);
    const drag = drawDrag;
    const geometry = drawGeometry.value;
    drawDrag = null;
    drawPreview.value = null;
    if (!drag || !geometry) {
      return;
    }
    const now = drawBoxNow(drag, event, geometry);
    if (!now) {
      return;
    }
    const { isTextBox, geometry: realGeometry } = now;
    const { leftCm, topCm, widthCm, heightCm } = now.box;

    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = caretSelection(drag.anchor);
    const offsetXCm = leftCm - drag.anchorLeft;
    const offsetYCm = topCm - drag.anchorTop;
    // The very shape the preview has been painting all along.
    const drawn = newDrawnShape(realGeometry, isTextBox, widthCm, heightCm);
    const result = isTextBox
      ? insertTextBox(model, selection)
      : insertShape(model, selection, {
          geometry: drawn.geometry,
          widthCm: drawn.widthCm,
          heightCm: drawn.heightCm,
          offsetXCm,
          offsetYCm,
          ...(drawn.outline ? { outline: drawn.outline } : {}),
          ...(drawn.fill ? { fill: drawn.fill } : {}),
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
      // AC-9: land already in text-editing mode with the caret inside,
      // matching the double-click gesture (handleShapeEditText also
      // populates Shape.text, which AC-51's role="group" depends on).
      handleShapeEditText(inserted.section, inserted.block);
    }
  }

  return {
    shapeSelection,
    shapeMultiSelection,
    shapeHandles,
    shapeAnchorMarker,
    drawGeometry,
    drawPreview,
    movingShape,
    editingShapeText,
    shapeTextRemountEpoch,
    clearShapeSelection,
    handleShapeSelect,
    handleShapeEditText,
    exitShapeTextEditing,
    handleShapeTextBeforeInput,
    handleShapeTextKeyDown,
    handleShapeTextCompositionStart,
    handleShapeTextCompositionEnd,
    shapeAtClientPoint,
    canGroupSelection,
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
    startShapeMove,
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
