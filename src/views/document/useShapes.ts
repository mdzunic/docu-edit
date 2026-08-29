// Shapes and text boxes (SPEC-051): the selected shape's state (selection,
// measured handles), draw-mode insertion (gallery pick -> drag/click on the
// canvas), the shape toolbar's actions, the eight-handle resize drag, the
// rotation-handle drag, keyboard nudge/duplicate, and z-order/grouping.
// Mirrors useImages.ts's shape closely — see its own header comment for the
// pattern this follows.
import { computed, nextTick, ref, watch } from "vue";
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
import { MIN_IMAGE_SIZE_CM } from "../../lib/core/media";
import { cmToPx } from "../../lib/core/units";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import { CM_PER_PX } from "../../lib/core/units";
import { modelRangeFromSelection, setDomCaret, setDomSelection } from "../../lib/domCaret";
import { findShortcut } from "../../lib/shortcuts";
import { blockLength, noCaret } from "../../lib/textEditing";
import { recordEdit as recordHistoryEdit } from "../../lib/undoHistory";
import { replaceDocumentModel } from "../../lib/documentStore";
import {
  DEFAULT_DRAWING_POSITION,
  deleteShape,
  duplicateShape,
  convertShapeToFreeform,
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
  reanchorShape,
  reorderShapeZ,
  resizeGroup,
  rotateGroup,
  setShapeAdjustments,
  setShapeAltText,
  setShapeFill,
  setShapeOutline,
  setDrawingGeometry,
  type DrawingGeometry,
  setShapePath,
  setShapeRotation,
  setShapeSize,
  setShapeWrap,
  shapeAt,
  shapeHasVisibleText,
  shapesInPaintOrder,
  ungroupShapes,
  type ParagraphSpan,
  type ShapeLocation,
  type ZOrderCommand,
} from "../../lib/shapes";
import {
  DEFAULT_NEW_SHAPE_FILL,
  DEFAULT_NEW_SHAPE_OUTLINE,
  drawGestureBox,
  normalizeRotationDeg,
  type NudgeDirection,
} from "../../lib/core/shapes";
import {
  HIT_TOLERANCE_CM,
  pointHitsShape,
  rotatePoint,
  rotatedResizeCursor,
  rotatedBounds,
  unrotateDelta,
  unrotatePoint,
} from "../../lib/core/shapeGeometry";
import { adjustmentHandles, adjustmentValueFromPoint, nudgedAdjustments } from "../../lib/shapeAdjustHandles";
import {
  CLOSE_DISTANCE_CM,
  SCRIBBLE_TOLERANCE_CM,
  catmullRomPath,
  drawnShapeBox,
  polylinePath,
  presetToPath,
} from "../../lib/core/shapeDraw";
import { pathToSvgD, simplifyPolyline, type ShapePath } from "../../lib/core/shapePath";
import {
  deleteVertex,
  moveControl,
  moveVertexTo,
  nudgeVertex,
  pathVertices,
  renormaliseAfterEdit,
  splitSegmentAfter,
  toggleVertexSmooth,
} from "../../lib/core/pointEdit";
import type { Point } from "../../lib/core/shapeGeometry";
import { resolveAnchorBlockIndex } from "../../lib/drawingAnchors";
import { asImageWrapMode, isShape, type DocumentModel, type Shape } from "../../lib/documentModel";
import { LINE_GEOMETRIES, asShapeGeometry, type ShapeGeometry, type ShapeOutline } from "../../lib/core/shapes";
import { useDrawingGuides } from "./useDrawingGuides";
import type { CaretPosition } from "../../lib/textEditing";
import type { CommandResult } from "../../lib/commands";
import type { EditorContext } from "./context";

function sameShapeLocation(a: ShapeLocation, b: ShapeLocation): boolean {
  return a.section === b.section && a.block === b.block;
}

function dist(a: Point, b: Point): number {
  return Math.hypot(a.x - b.x, a.y - b.y);
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

/** The shape text-editing pipeline's own commit path (applyShapeTextEdit)
 * bypasses the generic `applyEdit` — like useNotes.ts's applyNoteEdit, it
 * must fold a synthetic withShapeModel result back into the real document
 * (commitShapeModel) before it can go through the normal history/store
 * pipeline, so it needs `markDirty`/`refreshUndoRedoState` directly. */
export interface ShapesDeps extends EditorContext {
  /** Late-binding: useFields' surface-scoped F9 update (set up later), which
   * runs the command against the text frame's synthetic model — see
   * `updateFieldInShapeText`. */
  getUpdateFieldInSurface: () => (
    surfaceModel: DocumentModel,
    selection: ModelSelection,
  ) => CommandResult | null;
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
    getUpdateFieldInSurface,
    getCaret,
    markDirty,
    refreshUndoRedoState,
    zoom,
    clientToContentPx,
    contentPxToClient,
    onDrawingGeometryChanged,
  } = deps;

  // Alignment guides + snapping while a drawing is dragged or created (SPEC-072).
  const {
    snapGuides,
    applySnap: applyDrawingSnap,
    applyResizeSnap: applyDrawingResizeSnap,
    clearGuides: clearDrawingGuides,
  } = useDrawingGuides({ pageRef, contentRoot, clientToContentPx, zoom });

  const shapeSelection = ref<ShapeLocation | null>(null);
  /** Shift-click multi-selection (AC-34): empty for an ordinary single
   * selection, 2+ entries while a multi-object selection is active —
   * `shapeSelection` tracks the most recently (de)selected member so the
   * handles/toolbar still attach somewhere sensible. Shape-only (an Image
   * has its own separate selection ref); grouping across the two kinds is
   * out of scope for this gesture. */
  const shapeMultiSelection = ref<ShapeLocation[]>([]);
  const shapeHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  /** The numeric "Position & size" dialog (AC-42) — the touch/coarse-pointer
   * fallback for dragging a drawing handle. `shapePositionInitial` is the
   * captured prefill the dialog opens with. */
  const shapePositionOpen = ref(false);
  const shapePositionInitial = ref({
    offsetXCm: 0,
    offsetYCm: 0,
    widthCm: 0,
    heightCm: 0,
    rotationDeg: 0,
    relativeToH: "column",
    relativeToV: "paragraph",
    inline: false,
    isLine: false,
  });
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

  /** F9 inside a shape's text frame (SPEC-053 AC-15 with AC-33, which puts
   * fields in shape text boxes): the field at the frame's caret is recomputed
   * and the result folded back like any other shape text edit. It is
   * dispatched here rather than by the canvas shortcut registry because the
   * text frame stops keydown before the canvas sees it — which is right: a
   * registry command reads the *body* caret, so anything it did here would
   * edit the wrong surface. Only this one command is forwarded; the rest stay
   * body-scoped until they grow a frame-aware path of their own. Mirrors
   * useHeadersFooters' updateFieldInBand and useNotes' updateFieldInNote. */
  function updateFieldInShapeText(): void {
    const location = shapeSelection.value;
    const model = doc.value;
    const selection = currentShapeTextSelection();
    if (!location || !model || !editable.value || !selection) {
      return;
    }
    const synthetic = withShapeModel(model, location);
    const updated = getUpdateFieldInSurface()(synthetic, selection);
    if (updated) {
      applyShapeTextEdit(updated.model, updated.selection, "update field");
    }
  }

  function handleShapeTextKeyDown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      event.stopPropagation();
      exitShapeTextEditing();
      return;
    }
    if (findShortcut(event)?.id === "field.update") {
      event.preventDefault();
      updateFieldInShapeText();
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
  /** SPEC-071: the armed freeform tool, or null. Freeform/Curve place
   * vertices click by click; Scribble samples a drag. */
  const drawFreeformTool = ref<"freeform" | "curve" | "scribble" | null>(null);
  /** Placed vertices (freeform/curve) or live samples (scribble), in
   * page-relative model cm. */
  const freeformPoints = ref<Point[]>([]);
  /** The pointer's current page-cm position, for the rubber-band segment to
   * the next click (freeform/curve). */
  const freeformCursor = ref<Point | null>(null);
  let freeformScribbling = false;
  let freeformListening = false;

  /** SPEC-071 point-edit mode: the shape under edit, a live working copy of
   * its path (previewed by the overlay; the model is committed once on
   * exit), and the selected vertex. */
  const pointEditLocation = ref<ShapeLocation | null>(null);
  const pointEditPath = ref<ShapePath | null>(null);
  const pointEditSelectedVertex = ref(0);
  let pointEditPrevSelection: ShapeLocation | null = null;
  // Serialised snapshot of the path as point-edit was entered, so a session
  // that changes nothing commits nothing (no spurious undo step, no needless
  // drop of a verbatim `pathSource`).
  let pointEditOriginalJson: string | null = null;
  let pointEditDrag: { kind: "vertex" | "in" | "out"; index: number } | null = null;
  let pointEditListening = false;

  let resizeDrag: {
    location: ShapeLocation;
    handle: string;
    startClientX: number;
    startClientY: number;
    start: { widthCm: number; heightCm: number };
    el: HTMLElement;
    groupOrdinal: number | undefined;
    rotationDeg: number;
    /** The drawing's `.pages`-relative top-left in cm at drag start (SPEC-072):
     * the fixed offset corner the resize grows from, so the moving right/bottom
     * edges can be snapped to page candidates (AC-27). */
    startLeftCm: number;
    startTopCm: number;
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
    shapeHandles.value = el ? measureImageGeometry(pageEl, el, zoom.value) : null;
    shapeAnchorMarker.value = measureShapeAnchorMarker(location, root);
  }

  /** The paragraph a shape at `location` is anchored to, as a block index
   * in the same section (AC-27, AC-28) — the shape's explicit
   * `anchorOrdinal` pairing when it carries one, else the adjacency
   * fallback, both via the shared SPEC-078 resolver (the ONLY place an
   * anchor is derived — see lib/drawingAnchors.ts). Null when nothing
   * resolves (e.g. the shape sits between tables in a paragraph-less
   * stretch). */
  function anchorParagraphBlockIndex(location: ShapeLocation): number | null {
    const blocks = doc.value?.sections[location.section]?.blocks;
    return blocks ? resolveAnchorBlockIndex(blocks, location.block) : null;
  }

  /** The DOM position of the paragraph a selected shape is anchored to
   * (AC-27); null when there is none to mark (anchorParagraphBlockIndex) or
   * it isn't measurable. */
  function measureShapeAnchorMarker(
    location: ShapeLocation,
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
    const rect = el.getBoundingClientRect();
    // The marker is painted INSIDE .pages (the scale transform re-applies the
    // zoom), so it carries unscaled, .pages-relative px (SPEC-063 AC-25).
    const point = clientToContentPx(rect.left, rect.top);
    return { left: point.x, top: point.y };
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
      // Route through the zoom-aware helper so the page-relative cm is the
      // unscaled model cm at any zoom (SPEC-063 AC-22).
      const top = clientToContentPx(rect.left, rect.top);
      const bottom = clientToContentPx(rect.left, rect.bottom);
      spans.push({
        section: location.section,
        block: index,
        topCm: top.y * CM_PER_PX,
        bottomCm: bottom.y * CM_PER_PX,
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
      // The offset from the shape's own top-left, in unscaled model px, so
      // the hit test compares like-for-like with the shape's cm geometry at
      // any zoom (SPEC-063 AC-22/AC-24).
      const origin = clientToContentPx(rect.left, rect.top);
      const pointer = clientToContentPx(clientX, clientY);
      const point = {
        x: (pointer.x - origin.x) * CM_PER_PX,
        y: (pointer.y - origin.y) * CM_PER_PX,
      };
      const hit = pointHitsShape(point, {
        geometry: shape.geometry,
        widthCm: shape.widthCm,
        heightCm: shape.heightCm,
        rotationDeg: shape.rotationDeg,
        filled: !!shape.fill,
        hasText: shapeHasVisibleText(shape),
        outlineToleranceCm: HIT_TOLERANCE_CM,
        path: shape.path,
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

  // --- adjustment handles (SPEC-069 AC-5) ---------------------------------

  /** The in-progress adjustment vector during a drag/keyboard nudge, so the
   * handle positions and canvas preview reflect the uncommitted value; null
   * when not adjusting. */
  const adjustPreview = ref<{ location: ShapeLocation; adjustments: number[] } | null>(null);
  let adjustDrag: { location: ShapeLocation; index: number; el: HTMLElement; rotationDeg: number; moved: boolean } | null = null;

  /** Whether the selected shape's container is sealed (SPEC-057): it holds a
   * preserved member, so it offers no adjustment (nor resize/rotation)
   * handles. Flat groupOrdinal model — a preserved member anywhere in the
   * group seals it. */
  function isSelectionSealed(shape: Shape): boolean {
    const model = doc.value;
    if (!model || shape.groupOrdinal === undefined) {
      return false;
    }
    return groupMembers(model, shape.groupOrdinal).some(
      (m) => m.block.type === "shape" && m.block.unsupportedGeometryXml !== undefined,
    );
  }

  /** The selected shape's adjustment handles (AC-5a), each with its pixel
   * style — empty for a preserved shape (its geometry is the rect stand-in,
   * nothing to write to) or a member of a sealed container. Uses the live
   * preview value mid-drag so the handle tracks the pointer. */
  function adjustmentHandleList(): { index: number; label: string; value: number; style: Record<string, string> }[] {
    const model = doc.value;
    const location = shapeSelection.value;
    const geometry = shapeHandles.value;
    if (!model || !location || !geometry) {
      return [];
    }
    const shape = shapeAt(model, location);
    if (!shape || shape.unsupportedGeometryXml !== undefined || isSelectionSealed(shape)) {
      return [];
    }
    const stored =
      adjustPreview.value && sameShapeLocation(adjustPreview.value.location, location)
        ? adjustPreview.value.adjustments
        : shape.adjustments;
    const rotationDeg = selectedShapeRotationDeg();
    return adjustmentHandles(shape.geometry, stored).map((handle) => {
      const point = rotatePoint(
        { x: geometry.width * handle.xFrac, y: geometry.height * handle.yFrac },
        geometry.width,
        geometry.height,
        rotationDeg,
      );
      return {
        index: handle.index,
        label: handle.label,
        value: handle.value,
        style: {
          left: `${geometry.left + point.x - HANDLE_SIZE / 2}px`,
          top: `${geometry.top + point.y - HANDLE_SIZE / 2}px`,
        },
      };
    });
  }

  /** The unit-square fraction (in the shape's own unrotated frame) a client
   * point maps to, for the adjustment drag/mapping. */
  function pointerAdjustFraction(clientX: number, clientY: number): { x: number; y: number } | null {
    const geometry = shapeHandles.value;
    if (!geometry || geometry.width <= 0 || geometry.height <= 0) {
      return null;
    }
    const pointer = clientToContentPx(clientX, clientY);
    const local = unrotatePoint(
      { x: pointer.x - geometry.left, y: pointer.y - geometry.top },
      geometry.width,
      geometry.height,
      selectedShapeRotationDeg(),
    );
    return { x: local.x / geometry.width, y: local.y / geometry.height };
  }

  function startShapeAdjust(event: MouseEvent, index = 0) {
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
    event.stopPropagation();
    adjustDrag = { location, index, el, rotationDeg: shape.rotationDeg ?? 0, moved: false };
    window.addEventListener("mousemove", handleShapeAdjustMove);
    window.addEventListener("mouseup", finishShapeAdjust);
  }

  function handleShapeAdjustMove(event: MouseEvent) {
    const drag = adjustDrag;
    const model = doc.value;
    if (!drag || !model) {
      return;
    }
    const shape = shapeAt(model, drag.location);
    if (!shape) {
      return;
    }
    const frac = pointerAdjustFraction(event.clientX, event.clientY);
    if (!frac) {
      return;
    }
    const value = adjustmentValueFromPoint(shape.geometry, drag.index, frac.x, frac.y);
    if (value === null) {
      return;
    }
    drag.moved = true;
    const next = adjustmentHandles(shape.geometry, shape.adjustments).map((h) => h.value);
    next[drag.index] = value;
    adjustPreview.value = { location: drag.location, adjustments: next };
    // Live preview: repaint the shape's own canvas at the new adjustment
    // without touching the model (committed once on release, AC-5d/AC-5e).
    previewShapeCanvas(drag.el, { ...shape, adjustments: next }, shape.widthCm, shape.heightCm);
  }

  function finishShapeAdjust() {
    window.removeEventListener("mousemove", handleShapeAdjustMove);
    window.removeEventListener("mouseup", finishShapeAdjust);
    const drag = adjustDrag;
    adjustDrag = null;
    const preview = adjustPreview.value;
    adjustPreview.value = null;
    if (!drag || !drag.moved || !preview) {
      remeasureShapeHandles();
      return;
    }
    applyShapeEdit(
      (model, selection) => setShapeAdjustments(model, selection, drag.location, preview.adjustments),
      "adjust shape",
    );
    void nextTick().then(remeasureShapeHandles);
  }

  /** Keyboard adjustment (AC "Adjustment handles"): arrows nudge the focused
   * handle by 1% of its range, Shift by 5%; each keystroke commits one
   * "adjust shape" undo entry. */
  function handleAdjustmentKeydown(event: KeyboardEvent, index = 0) {
    const model = doc.value;
    const location = shapeSelection.value;
    if (!model || !location) {
      return;
    }
    const shape = shapeAt(model, location);
    if (!shape) {
      return;
    }
    let steps = 0;
    if (event.key === "ArrowRight" || event.key === "ArrowUp") {
      steps = 1;
    } else if (event.key === "ArrowLeft" || event.key === "ArrowDown") {
      steps = -1;
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const next = nudgedAdjustments(shape.geometry, shape.adjustments, index, steps, event.shiftKey);
    applyShapeEdit((m, selection) => setShapeAdjustments(m, selection, location, next), "adjust shape");
    void nextTick().then(remeasureShapeHandles);
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
    // The toolbar is a position:fixed element outside the scale transform, so
    // its anchor is the shape's model-px geometry mapped back into client
    // space through the zoom-aware helper (SPEC-063 AC-22/AC-57). The +8/px
    // clearances are screen constants and stay unscaled.
    const topLeft = contentPxToClient(geometry.left - bounds.padX, geometry.top - bounds.padY);
    const shapeTop = topLeft.y;
    const bottomEdge = shapeTop - SHAPE_TOOLBAR_CLEARANCE_PX;
    const left = topLeft.x;
    if (bottomEdge - SHAPE_TOOLBAR_MAX_HEIGHT_PX >= 0) {
      return { left, top: bottomEdge, placement: "above" };
    }
    const below = contentPxToClient(geometry.left - bounds.padX, geometry.top + geometry.height + bounds.padY);
    return { left, top: below.y + 8, placement: "below" };
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
    // A wrap change moves the shape more than any other edit — between text
    // flow and an absolute offset — so the handles/toolbar must re-find it.
    void nextTick().then(remeasureShapeHandles);
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

  /** Captures the selected shape's geometry into the dialog prefill. */
  function seedShapePosition(shape: Shape) {
    shapePositionInitial.value = {
      offsetXCm: shape.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm,
      offsetYCm: shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm,
      widthCm: shape.widthCm,
      heightCm: shape.heightCm,
      rotationDeg: shape.rotationDeg ?? 0,
      relativeToH: shape.relativeToH ?? DEFAULT_DRAWING_POSITION.relativeToH,
      relativeToV: shape.relativeToV ?? DEFAULT_DRAWING_POSITION.relativeToV,
      inline: shape.wrap === "inline",
      isLine: LINE_GEOMETRIES.has(shape.geometry),
    };
  }

  /** Opens the numeric "Position & size" dialog for the selected shape
   * (SPEC-072): the keyboard/coarse-pointer path for every geometric property
   * a drag adjusts. Prefills from the shape's current geometry. */
  function openShapePosition() {
    const model = doc.value;
    const location = shapeSelection.value;
    const shape = model && location ? shapeAt(model, location) : null;
    if (!shape) {
      return;
    }
    seedShapePosition(shape);
    shapePositionOpen.value = true;
  }

  // Reflect the selection live (AC-51): if the selection moves to a different
  // shape while the dialog is open, re-seed the prefill so it shows the new
  // shape's geometry.
  watch(shapeSelection, (location) => {
    if (!shapePositionOpen.value) {
      return;
    }
    const model = doc.value;
    const shape = model && location ? shapeAt(model, location) : null;
    if (shape) {
      seedShapePosition(shape);
    } else {
      shapePositionOpen.value = false;
    }
  });

  /** Applies one live field change from the dialog as a single undo entry
   * (SPEC-072 AC-43/AC-45) — the whole geometry set at once so the entry is
   * self-contained. The dialog stays open (live editing); dismiss closes it. */
  function handleShapePositionApply(
    geometry: DrawingGeometry = { widthCm: 0, heightCm: 0, offsetXCm: 0, offsetYCm: 0, rotationDeg: 0 },
  ) {
    const location = shapeSelection.value;
    if (!location) {
      return;
    }
    applyShapeEdit(
      (model, selection) => setDrawingGeometry(model, selection, location, geometry),
      "shape position",
    );
    void nextTick().then(remeasureShapeHandles);
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
    const rect = el.getBoundingClientRect();
    // Unscaled, .pages-relative top-left in model cm (SPEC-063 AC-22) — the
    // fixed corner the resize snap grows from (AC-27).
    const startTopLeft = clientToContentPx(rect.left, rect.top);
    resizeDrag = {
      location,
      handle,
      startClientX: event.clientX,
      startClientY: event.clientY,
      start: { widthCm: shape.widthCm, heightCm: shape.heightCm },
      el,
      groupOrdinal: shape.groupOrdinal,
      rotationDeg: shape.rotationDeg ?? 0,
      startLeftCm: startTopLeft.x * CM_PER_PX,
      startTopCm: startTopLeft.y * CM_PER_PX,
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
      shape,
      Math.max(1, Math.round(cmToPx(widthCm))),
      Math.max(1, Math.round(cmToPx(Math.max(heightCm, 0.05)))),
      zoom.value,
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
    // The pointer delta is measured in unscaled content px (both endpoints
    // through the zoom-aware helper) so a resize tracks the pointer 1:1 at
    // any zoom (SPEC-063 AC-24).
    const start = clientToContentPx(drag.startClientX, drag.startClientY);
    const now = clientToContentPx(event.clientX, event.clientY);
    const local = unrotateDelta(
      (now.x - start.x) * CM_PER_PX,
      (now.y - start.y) * CM_PER_PX,
      drag.rotationDeg,
    );
    const size = resizedImageSize(drag.start, asImageHandle(drag.handle), local.dx, local.dy, event.shiftKey);
    let widthCm = size.widthCm;
    let heightCm = size.heightCm;
    // Snapping + guides on resize (AC-27): the model resize keeps the top-left
    // offset and grows the size, so the right/bottom edges snap onto page
    // candidates. Only for a single, unrotated shape — a grouped member previews
    // on release, and snapping an axis-aligned page rectangle is meaningless once
    // the shape's own axes are turned. Alt suppresses it (AC-25).
    if (drag.groupOrdinal === undefined && drag.rotationDeg === 0) {
      const snapped = applyDrawingResizeSnap(
        { xCm: drag.startLeftCm, yCm: drag.startTopCm, wCm: widthCm, hCm: heightCm },
        drag.el,
        event.altKey,
      );
      widthCm = Math.max(MIN_IMAGE_SIZE_CM, snapped.widthCm);
      heightCm = Math.max(MIN_IMAGE_SIZE_CM, snapped.heightCm);
    }
    drag.currentWidthCm = widthCm;
    drag.currentHeightCm = heightCm;
    // A grouped member previews only on release, same as a grouped move
    // (finishShapeResize's resizeGroup scales every member about the
    // group's own bounding-box centre — live-previewing just the dragged
    // element here would show the rest of the group left behind).
    if (drag.groupOrdinal === undefined) {
      previewImageSize(drag.el, widthCm, heightCm);
      const shape = doc.value ? shapeAt(doc.value, drag.location) : null;
      if (shape) {
        previewShapeCanvas(drag.el, shape, widthCm, heightCm);
      }
      // The handles are what the pointer is dragging — they have to track
      // the box they are sizing, not sit at the size it started.
      remeasureShapeHandles();
    }
    // SPEC-070: re-route any connector attached to this shape, live.
    onDrawingGeometryChanged?.();
  }

  function finishShapeResize() {
    window.removeEventListener("mousemove", handleShapeResizeMove);
    window.removeEventListener("mouseup", finishShapeResize);
    // Guides are transient — they vanish on release (AC-56).
    clearDrawingGuides();
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
    startLeftCm: number;
    widthCm: number;
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
    const rect = el.getBoundingClientRect();
    // Unscaled, .pages-relative top in model cm (SPEC-063 AC-22).
    const startTop = clientToContentPx(rect.left, rect.top);
    moveDragMoved = false;
    moveDrag = {
      location,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startOffsetXCm: shape.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm,
      startOffsetYCm: shape.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm,
      startTopCm: startTop.y * CM_PER_PX,
      startLeftCm: startTop.x * CM_PER_PX,
      widthCm: shape.widthCm,
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
    // The move must track the pointer 1:1 at any zoom: the drag delta and the
    // handle preview are in unscaled content px, not raw (scaled) screen px
    // (SPEC-063 AC-24/AC-25). The 3px threshold above stays in screen px.
    const start = clientToContentPx(drag.startClientX, drag.startClientY);
    const now = clientToContentPx(event.clientX, event.clientY);
    const dxContent = now.x - start.x;
    const dyContent = now.y - start.y;
    drag.dxCm = dxContent * CM_PER_PX;
    drag.dyCm = dyContent * CM_PER_PX;
    // Snapping + alignment guides (SPEC-072). Only single shapes snap; a group
    // move previews on release (below), and Alt suppresses snapping (AC-25).
    // The snap adjusts the delta by however far the drawing was pulled onto a
    // candidate, so the offset preview and the commit both use it.
    if (drag.groupOrdinal === undefined) {
      const snapped = applyDrawingSnap(
        { xCm: drag.startLeftCm + drag.dxCm, yCm: drag.startTopCm + drag.dyCm, wCm: drag.widthCm, hCm: drag.heightCm },
        drag.el,
        event.altKey,
      );
      drag.dxCm += snapped.xCm - (drag.startLeftCm + drag.dxCm);
      drag.dyCm += snapped.yCm - (drag.startTopCm + drag.dyCm);
    }
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
          left: drag.startHandles.left + dxContent,
          top: drag.startHandles.top + dyContent,
        };
      }
      // The toolbar stays put wherever it was opened, which during a drag
      // means it sits across the page in the way of the shape being
      // dragged; it comes back on release, repositioned (finishShapeMove's
      // remeasureShapeHandles).
      movingShape.value = true;
      // SPEC-070: re-route any connector attached to this shape, live.
      onDrawingGeometryChanged?.();
    }
  }

  function finishShapeMove() {
    window.removeEventListener("mousemove", handleShapeMoveMove);
    window.removeEventListener("mouseup", finishShapeMove);
    // Guides are transient — they vanish on release (AC-56).
    clearDrawingGuides();
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
    // The shape's centre in client space: its model-px centre mapped through
    // the zoom-aware helper so the rotation pivot is where the shape actually
    // sits on screen at any zoom (SPEC-063 AC-24).
    const center = contentPxToClient(geometry.left + geometry.width / 2, geometry.top + geometry.height / 2);
    const centerClientX = center.x;
    const centerClientY = center.y;
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
    // SPEC-070: re-route any connector attached to this shape, live.
    onDrawingGeometryChanged?.();
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

  /** "Insert → Shape" gallery pick (SPEC-051 AC-3): arms draw mode so the
   * next drag/click on the canvas draws the chosen geometry. */
  function handleShapePick(geometry = "rect") {
    armShapeDraw(asShapeGeometry(geometry));
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
    const rect = el.getBoundingClientRect();
    // Unscaled, .pages-relative anchor position in model cm (SPEC-063 AC-22).
    const point = clientToContentPx(rect.left, rect.top);
    return { anchor, left: point.x * CM_PER_PX, top: point.y * CM_PER_PX };
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
    const isTextBox = geometry === "textbox";
    const realGeometry: ShapeGeometry = isTextBox ? "rect" : (geometry as ShapeGeometry);
    // Both corners in unscaled, .pages-relative model cm so the drawn box
    // matches the pointer at any zoom (SPEC-063 AC-24).
    const startPoint = clientToContentPx(drag.startClientX, drag.startClientY);
    const nowPoint = clientToContentPx(event.clientX, event.clientY);
    const box = drawGestureBox(
      realGeometry,
      {
        xCm: startPoint.x * CM_PER_PX,
        yCm: startPoint.y * CM_PER_PX,
      },
      { xCm: nowPoint.x * CM_PER_PX, yCm: nowPoint.y * CM_PER_PX },
      { shiftKey: event.shiftKey, forceDefaultSize: isTextBox },
    );
    // Snapping applies to the creation drag too (AC-28); there is no element
    // yet, so nothing is excluded. Alt suppresses it (AC-25).
    const snapped = applyDrawingSnap(
      { xCm: box.leftCm, yCm: box.topCm, wCm: box.widthCm, hCm: box.heightCm },
      null,
      event.altKey,
    );
    return {
      geometry: realGeometry,
      isTextBox,
      box: { ...box, leftCm: snapped.xCm, topCm: snapped.yCm },
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
    clearDrawingGuides();
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

  // --- freeform / curve / scribble drawing (SPEC-071) ---------------------

  /** A pointer event's position in page-relative model cm — the space
   * freeformPoints and currentAnchor both use. */
  function pageCmPoint(event: MouseEvent): Point {
    const p = clientToContentPx(event.clientX, event.clientY);
    return { x: p.x * CM_PER_PX, y: p.y * CM_PER_PX };
  }

  /** "Insert → Freeform" pick (AC-6): arms a freeform tool, disarming any
   * shape/connector draw and clearing selection so the canvas draws it. */
  function handleFreeformPick(tool = "freeform") {
    disarmShapeDraw();
    clearShapeSelection();
    drawFreeformTool.value = tool as "freeform" | "curve" | "scribble";
    freeformPoints.value = [];
    freeformCursor.value = null;
  }

  function disarmFreeform() {
    stopFreeformListeners();
    drawFreeformTool.value = null;
    freeformPoints.value = [];
    freeformCursor.value = null;
    freeformScribbling = false;
  }

  function startFreeformListeners() {
    if (freeformListening) {
      return;
    }
    freeformListening = true;
    window.addEventListener("mousemove", handleFreeformMove);
    window.addEventListener("keydown", handleFreeformKey, true);
    window.addEventListener("dblclick", handleFreeformDblClick, true);
  }

  function stopFreeformListeners() {
    if (!freeformListening) {
      return;
    }
    freeformListening = false;
    window.removeEventListener("mousemove", handleFreeformMove);
    window.removeEventListener("keydown", handleFreeformKey, true);
    window.removeEventListener("dblclick", handleFreeformDblClick, true);
    window.removeEventListener("mouseup", handleScribbleUp);
  }

  /** Canvas mousedown while a freeform tool is armed (routed from
   * useCanvasGestures ahead of every other canvas behaviour). */
  function startFreeformGesture(event: MouseEvent) {
    if (!drawFreeformTool.value || !editable.value) {
      return;
    }
    event.preventDefault();
    const point = pageCmPoint(event);
    if (drawFreeformTool.value === "scribble") {
      freeformScribbling = true;
      freeformPoints.value = [point];
      startFreeformListeners();
      window.addEventListener("mouseup", handleScribbleUp);
      return;
    }
    // Freeform / Curve: place a vertex. A click within CLOSE_DISTANCE_CM of
    // the origin (with at least two vertices) closes the subpath and ends
    // the draw (AC-7).
    const pts = freeformPoints.value;
    if (pts.length >= 2 && dist(point, pts[0]) <= CLOSE_DISTANCE_CM) {
      finishFreeform(true);
      return;
    }
    // Ignore a duplicate placed on top of the last vertex (the second click
    // of a double-click), so a dbl-click adds one vertex, not two.
    if (pts.length > 0 && dist(point, pts[pts.length - 1]) < 0.01) {
      return;
    }
    freeformPoints.value = [...pts, point];
    freeformCursor.value = point;
    startFreeformListeners();
  }

  function handleFreeformMove(event: MouseEvent) {
    if (!drawFreeformTool.value) {
      return;
    }
    const point = pageCmPoint(event);
    if (freeformScribbling) {
      freeformPoints.value = [...freeformPoints.value, point];
    } else {
      freeformCursor.value = point;
    }
  }

  function handleScribbleUp() {
    if (drawFreeformTool.value !== "scribble") {
      return;
    }
    freeformScribbling = false;
    finishFreeform(false);
  }

  function handleFreeformKey(event: KeyboardEvent) {
    if (!drawFreeformTool.value) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      disarmFreeform();
    } else if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      finishFreeform(false);
    } else if (event.key === "Backspace") {
      event.preventDefault();
      event.stopPropagation();
      const pts = freeformPoints.value;
      if (pts.length <= 1) {
        disarmFreeform();
      } else {
        freeformPoints.value = pts.slice(0, -1);
      }
    }
  }

  function handleFreeformDblClick(event: MouseEvent) {
    if (!drawFreeformTool.value || drawFreeformTool.value === "scribble") {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    finishFreeform(false);
  }

  /** Commits the drawn path as a single "draw freeform" undo entry (AC-12),
   * or nothing when it has too few vertices (AC-7/AC-9). */
  function finishFreeform(closed: boolean) {
    const tool = drawFreeformTool.value;
    let points = freeformPoints.value;
    if (!tool) {
      disarmFreeform();
      return;
    }
    if (tool === "scribble") {
      points = simplifyPolyline(points, SCRIBBLE_TOLERANCE_CM);
    }
    if (points.length < 2) {
      disarmFreeform();
      return;
    }
    const rawPath =
      tool === "freeform" ? polylinePath(points, closed) : catmullRomPath(points, tool === "scribble" ? false : closed);
    const { path, widthCm, heightCm, offset } = drawnShapeBox(rawPath);
    const model = doc.value;
    const anchorInfo = currentAnchor();
    if (!model || !anchorInfo) {
      disarmFreeform();
      return;
    }
    const selection = caretSelection(anchorInfo.anchor);
    const result = insertShape(model, selection, {
      geometry: "rect",
      widthCm,
      heightCm,
      offsetXCm: offset.x - anchorInfo.left,
      offsetYCm: offset.y - anchorInfo.top,
      outline: DEFAULT_NEW_SHAPE_OUTLINE,
      // A closed freeform fills; an open one is outline-only (AC-13).
      ...(closed && tool !== "scribble" ? { fill: DEFAULT_NEW_SHAPE_FILL } : {}),
      path,
    });
    applyEdit(result, selection, "draw freeform");
    disarmFreeform();
    const inserted = shapeInsertedLocation(result.model, anchorInfo.anchor);
    if (inserted) {
      shapeSelection.value = inserted;
      void nextTick().then(remeasureShapeHandles);
    }
  }

  /** The in-flight freeform preview as an SVG `d` string in page px, or null
   * when nothing is being drawn — the overlay strokes it live. */
  const freeformPreviewD = computed(() => {
    const tool = drawFreeformTool.value;
    if (!tool) {
      return null;
    }
    const pts =
      tool === "scribble" || !freeformCursor.value ? freeformPoints.value : [...freeformPoints.value, freeformCursor.value];
    if (pts.length < 1) {
      return null;
    }
    if (pts.length === 1) {
      return `M ${cmToPx(pts[0].x)} ${cmToPx(pts[0].y)}`;
    }
    const preview = tool === "curve" ? catmullRomPath(pts, false) : polylinePath(pts, false);
    const pxPerCm = cmToPx(1);
    return pathToSvgD(preview, pxPerCm, pxPerCm);
  });

  // --- point-edit mode (SPEC-071) -----------------------------------------

  /** Whether "Edit points" can act on the current selection: a shape with a
   * path already, or a preset whose outline can produce one (AC-14; a line
   * or unsupported geometry has nothing to edit). */
  function canEditPoints(): boolean {
    const model = doc.value;
    const location = shapeSelection.value;
    if (!model || !location) {
      return false;
    }
    const shape = shapeAt(model, location);
    if (!shape || shape.unsupportedGeometryXml !== undefined || isSelectionSealed(shape)) {
      return false;
    }
    return !!shape.path || !!presetToPath(shape.geometry, shape.adjustments);
  }

  const pointEditActive = computed(() => pointEditLocation.value !== null);

  /** Enters point-edit mode (AC-14). A preset is first converted to a
   * freeform path as its own "convert to freeform" undo entry (AC-15); a
   * shape that already has a path is entered directly (AC-16). */
  function enterPointEditMode() {
    const model = doc.value;
    const location = shapeSelection.value;
    if (!model || !location) {
      return;
    }
    const shape = shapeAt(model, location);
    if (!shape || shape.unsupportedGeometryXml !== undefined || isSelectionSealed(shape)) {
      return;
    }
    if (!shape.path) {
      // A separate undo entry, independently undoable from the edits (AC-15).
      applyShapeEdit((m, sel, loc) => convertShapeToFreeform(m, sel, loc), "convert to freeform");
    }
    const convertedModel = doc.value;
    const converted = convertedModel ? shapeAt(convertedModel, location) : null;
    if (!converted?.path) {
      return;
    }
    pointEditPrevSelection = location;
    pointEditLocation.value = location;
    pointEditPath.value = clonePath(converted.path);
    pointEditOriginalJson = JSON.stringify(converted.path);
    pointEditSelectedVertex.value = 0;
    if (!pointEditListening) {
      pointEditListening = true;
      window.addEventListener("keydown", handlePointEditKey, true);
    }
    void nextTick().then(remeasureShapeHandles);
  }

  function clonePath(path: ShapePath): ShapePath {
    return {
      subpaths: path.subpaths.map((s) => ({ steps: s.steps.map((step) => ({ ...step })), closed: s.closed })),
      smooth: path.smooth ? [...path.smooth] : undefined,
    };
  }

  /** Leaves point-edit mode, committing the whole session as one "edit
   * points" undo entry with the box renormalised (AC-22/AC-23), or deleting
   * the shape when its path emptied (AC-21). Restores the prior selection. */
  function exitPointEditMode(commit = true) {
    const location = pointEditLocation.value;
    const path = pointEditPath.value;
    const originalJson = pointEditOriginalJson;
    stopPointEditListeners();
    pointEditLocation.value = null;
    pointEditPath.value = null;
    pointEditOriginalJson = null;
    pointEditDrag = null;
    if (!location) {
      return;
    }
    // A session that left the path untouched commits nothing: no spurious undo
    // step, and the shape's verbatim `pathSource` (if any) is preserved.
    const unchanged = path !== null && originalJson !== null && JSON.stringify(path) === originalJson;
    if (commit && !unchanged) {
      if (!path) {
        // Every subpath went (AC-21): delete the shape.
        applyShapeEdit((m, sel) => deleteShape(m, sel, location), "edit points");
      } else {
        const model = doc.value;
        const shape = model ? shapeAt(model, location) : null;
        if (shape) {
          const { path: fitted, offset, scale } = renormaliseAfterEdit(path);
          const box = {
            widthCm: shape.widthCm * (scale.x || 1),
            heightCm: shape.heightCm * (scale.y || 1),
            offsetXCm: (shape.offsetXCm ?? 0) + offset.x * shape.widthCm,
            offsetYCm: (shape.offsetYCm ?? 0) + offset.y * shape.heightCm,
          };
          applyShapeEdit((m, sel, loc) => setShapePath(m, sel, loc, fitted, box), "edit points");
        }
      }
    }
    shapeSelection.value = pointEditPrevSelection;
    pointEditPrevSelection = null;
    void nextTick().then(remeasureShapeHandles);
  }

  function stopPointEditListeners() {
    if (pointEditListening) {
      pointEditListening = false;
      window.removeEventListener("keydown", handlePointEditKey, true);
    }
    window.removeEventListener("mousemove", handlePointEditMove);
    window.removeEventListener("mouseup", finishPointEditDrag);
  }

  /** The edited outline as an SVG `d` in the shape's local px (the overlay
   * scales/rotates it to sit on the shape). */
  const pointEditOutlineD = computed(() => {
    const path = pointEditPath.value;
    const geometry = shapeHandles.value;
    if (!path || !geometry) {
      return null;
    }
    return pathToSvgD(path, geometry.width, geometry.height);
  });

  /** The vertex handles for the overlay: each with its screen style, whether
   * it is a corner or smooth (square vs circle, AC-19) and whether selected. */
  function pointEditVertexList(): { index: number; smooth: boolean; selected: boolean; style: Record<string, string> }[] {
    const path = pointEditPath.value;
    const geometry = shapeHandles.value;
    if (!path || !geometry) {
      return [];
    }
    const rotationDeg = selectedShapeRotationDeg();
    return pathVertices(path).map((v) => {
      const p = rotatePoint(
        { x: geometry.width * v.point.x, y: geometry.height * v.point.y },
        geometry.width,
        geometry.height,
        rotationDeg,
      );
      return {
        index: v.index,
        smooth: v.smooth,
        selected: v.index === pointEditSelectedVertex.value,
        style: { left: `${geometry.left + p.x - HANDLE_SIZE / 2}px`, top: `${geometry.top + p.y - HANDLE_SIZE / 2}px` },
      };
    });
  }

  /** Control-point handles + leader lines for the selected vertex and its
   * neighbours (AC-18/AC-19). */
  function pointEditControlList(): {
    index: number;
    which: "in" | "out";
    dot: Record<string, string>;
    line: Record<string, string>;
  }[] {
    const path = pointEditPath.value;
    const geometry = shapeHandles.value;
    if (!path || !geometry) {
      return [];
    }
    const rotationDeg = selectedShapeRotationDeg();
    const vertices = pathVertices(path);
    const selected = pointEditSelectedVertex.value;
    const shown = new Set([selected, selected - 1, selected + 1]);
    const toScreen = (fx: number, fy: number): Point =>
      rotatePoint({ x: geometry.width * fx, y: geometry.height * fy }, geometry.width, geometry.height, rotationDeg);
    const out: { index: number; which: "in" | "out"; dot: Record<string, string>; line: Record<string, string> }[] = [];
    for (const v of vertices) {
      if (!shown.has(v.index)) {
        continue;
      }
      const anchor = toScreen(v.point.x, v.point.y);
      for (const [which, ctrl] of [
        ["in", v.incoming],
        ["out", v.outgoing],
      ] as const) {
        if (!ctrl) {
          continue;
        }
        const c = toScreen(ctrl.x, ctrl.y);
        out.push({
          index: v.index,
          which,
          dot: { left: `${geometry.left + c.x - 4}px`, top: `${geometry.top + c.y - 4}px` },
          line: leaderLineStyle(geometry.left + anchor.x, geometry.top + anchor.y, geometry.left + c.x, geometry.top + c.y),
        });
      }
    }
    return out;
  }

  function leaderLineStyle(x1: number, y1: number, x2: number, y2: number): Record<string, string> {
    const length = Math.hypot(x2 - x1, y2 - y1);
    const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
    return {
      left: `${x1}px`,
      top: `${y1}px`,
      width: `${length}px`,
      transform: `rotate(${angle}deg)`,
    };
  }

  function selectPointEditVertex(index = 0) {
    pointEditSelectedVertex.value = index;
  }

  function startPointEditVertexDrag(event: MouseEvent, index = 0) {
    if (!pointEditPath.value) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    pointEditSelectedVertex.value = index;
    pointEditDrag = { kind: "vertex", index };
    window.addEventListener("mousemove", handlePointEditMove);
    window.addEventListener("mouseup", finishPointEditDrag);
  }

  function startPointEditControlDrag(event: MouseEvent, index = 0, which = "out") {
    if (!pointEditPath.value) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    pointEditSelectedVertex.value = index;
    pointEditDrag = { kind: which === "in" ? "in" : "out", index };
    window.addEventListener("mousemove", handlePointEditMove);
    window.addEventListener("mouseup", finishPointEditDrag);
  }

  function handlePointEditMove(event: MouseEvent) {
    const drag = pointEditDrag;
    const path = pointEditPath.value;
    if (!drag || !path) {
      return;
    }
    const fraction = pointerAdjustFraction(event.clientX, event.clientY);
    if (!fraction) {
      return;
    }
    if (drag.kind === "vertex") {
      pointEditPath.value = moveVertexTo(path, drag.index, fraction);
    } else {
      pointEditPath.value = moveControl(path, drag.index, drag.kind, fraction);
    }
  }

  function finishPointEditDrag() {
    pointEditDrag = null;
    window.removeEventListener("mousemove", handlePointEditMove);
    window.removeEventListener("mouseup", finishPointEditDrag);
  }

  /** Double-click a vertex toggles it corner/smooth (AC-17). */
  function togglePointEditVertex(index = 0) {
    if (!pointEditPath.value) {
      return;
    }
    pointEditSelectedVertex.value = index;
    pointEditPath.value = toggleVertexSmooth(pointEditPath.value, index);
  }

  /** The overlay frame: the shape's box + rotation and the live outline `d`
   * (in the shape's local px), so the overlay paints the edited outline and
   * lays its handles over the same rotated box. */
  const pointEditFrame = computed(() => {
    const g = shapeHandles.value;
    const d = pointEditOutlineD.value;
    if (!g || d === null) {
      return null;
    }
    return { left: g.left, top: g.top, width: g.width, height: g.height, rotationDeg: selectedShapeRotationDeg(), d };
  });

  /** Click on a segment adds a vertex there (AC-17): splits the segment
   * leaving the vertex nearest the click. */
  function pointEditSplitAtPointer(event: MouseEvent) {
    const path = pointEditPath.value;
    const fraction = pointerAdjustFraction(event.clientX, event.clientY);
    if (!path || !fraction) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    let best = 0;
    let bestDist = Infinity;
    for (const v of pathVertices(path)) {
      const d = Math.hypot(v.point.x - fraction.x, v.point.y - fraction.y);
      if (d < bestDist) {
        bestDist = d;
        best = v.index;
      }
    }
    pointEditPath.value = splitSegmentAfter(path, best, 0.5);
  }

  function deleteSelectedVertex() {
    const path = pointEditPath.value;
    if (!path) {
      return;
    }
    const next = deleteVertex(path, pointEditSelectedVertex.value);
    if (next === null) {
      // Every subpath is gone (AC-21): null the path first so exitPointEditMode
      // reads no path and takes its delete-the-shape branch instead of
      // re-committing the stale outline.
      pointEditPath.value = null;
      exitPointEditMode(true);
      return;
    }
    pointEditPath.value = next;
    pointEditSelectedVertex.value = Math.max(0, Math.min(pointEditSelectedVertex.value, pathVertices(next).length - 1));
  }

  function handlePointEditKey(event: KeyboardEvent) {
    const path = pointEditPath.value;
    const location = pointEditLocation.value;
    if (!path || !location) {
      return;
    }
    const vertices = pathVertices(path);
    const count = vertices.length;
    const shape = doc.value ? shapeAt(doc.value, location) : null;
    const widthCm = shape?.widthCm || 1;
    const heightCm = shape?.heightCm || 1;
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      exitPointEditMode(true);
    } else if (event.key === "Tab") {
      event.preventDefault();
      event.stopPropagation();
      const step = event.shiftKey ? -1 : 1;
      pointEditSelectedVertex.value = ((pointEditSelectedVertex.value + step) % count + count) % count;
    } else if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      togglePointEditVertex(pointEditSelectedVertex.value);
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      event.stopPropagation();
      deleteSelectedVertex();
    } else if (event.key.startsWith("Arrow")) {
      event.preventDefault();
      event.stopPropagation();
      const stepCm = event.shiftKey ? 0.1 : 0.01;
      const dx = (event.key === "ArrowLeft" ? -stepCm : event.key === "ArrowRight" ? stepCm : 0) / widthCm;
      const dy = (event.key === "ArrowUp" ? -stepCm : event.key === "ArrowDown" ? stepCm : 0) / heightCm;
      pointEditPath.value = nudgeVertex(path, pointEditSelectedVertex.value, dx, dy);
    }
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
    snapGuides,
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
    adjustmentHandleList,
    startShapeAdjust,
    handleAdjustmentKeydown,
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
    shapePositionOpen,
    shapePositionInitial,
    openShapePosition,
    handleShapePositionApply,
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
    handleShapePick,
    disarmShapeDraw,
    startShapeDraw,
    finishShapeDraw,
    drawFreeformTool,
    freeformPreviewD,
    handleFreeformPick,
    startFreeformGesture,
    disarmFreeform,
    pointEdit: {
      active: pointEditActive,
      location: pointEditLocation,
      frame: pointEditFrame,
      canEdit: canEditPoints,
      enter: enterPointEditMode,
      exit: exitPointEditMode,
      vertices: pointEditVertexList,
      controls: pointEditControlList,
      select: selectPointEditVertex,
      onVertexDown: startPointEditVertexDrag,
      onVertexToggle: togglePointEditVertex,
      onControlDown: startPointEditControlDrag,
      onSegmentDown: pointEditSplitAtPointer,
    },
    insertTextBoxAction,
    freshGroupOrdinal: () => (doc.value ? freshGroupOrdinal(doc.value) : 1),
    moveGroup,
    resizeGroup,
    rotateGroup,
  };
}
