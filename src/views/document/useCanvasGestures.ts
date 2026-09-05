// The canvas gesture hub: every mouse gesture the page canvas dispatches on —
// the SPEC-022 cell-range drag selection and the table hover handles, shape
// draw/move and dblclick recovery (SPEC-051), chart recovery (SPEC-052),
// image/shape/chart/TOC deselect on a plain mousedown (SPEC-024, SPEC-033),
// note-reference and hyperlink clicks (SPEC-031, SPEC-025), image
// drag-and-drop (SPEC-024), the format painter's drag tracking (SPEC-020)
// and margin-click caret placement (SPEC-011). One hub, because a single
// mousedown/mousemove/mouseup on the canvas fans out into all of them in a
// defined order.
//
// The `cellSelection` and `tableHover` refs stay in DocumentView — useTables'
// wiring and the template consume them — and come in as deps; the drag
// origin and drag-tracking variables live here as plain non-reactive lets
// (nothing renders from them).
import { computed } from "vue";
import { caretSelection } from "../../lib/commands";
import {
  caretAtClientPoint,
  caretNearClientPoint,
  measureTableHandles,
} from "../../lib/caretGeometry";
import { isAcceptedImageType } from "../../lib/core/media";
import { cellRangeBounds, noCellRange } from "../../lib/core/tables";
import {
  noEventTarget,
  noTableCellTarget,
  tableCellFromEventTarget,
} from "../../lib/domCaret";
import { noteById } from "../../lib/notes";
import { noShapeLocation } from "../../lib/shapes";
import type { Ref } from "vue";
import type { TableHandleGeometry } from "../../lib/caretGeometry";
import type { ChartLocation } from "../../lib/charts";
import type { ModelSelection } from "../../lib/commands";
import type { ShapeGeometry } from "../../lib/core/shapes";
import type { CellRange } from "../../lib/core/tables";
import type { NoteKind } from "../../lib/documentModel";
import type { LinkTarget } from "../../lib/links";
import type { ShapeLocation } from "../../lib/shapes";
import type { EditorContext } from "./context";

export interface CanvasGesturesDeps extends EditorContext {
  // --- shared selection state owned by the view -----------------------------
  /** The rectangular cell-range selection (SPEC-022 AC-15, AC-16). */
  cellSelection: Ref<CellRange | null>;
  /** The hover handles' measured geometry (SPEC-022 AC-20, AC-21). */
  tableHover: Ref<TableHandleGeometry | null>;
  // --- from useShapes (SPEC-051) --------------------------------------------
  /** The armed draw-mode geometry — read-only here; useShapes owns it. */
  drawGeometry: Readonly<Ref<ShapeGeometry | "textbox" | null>>;
  startShapeDraw: (event: MouseEvent) => void;
  /** The armed freeform tool (SPEC-071) — read-only here; useShapes owns it. */
  drawFreeformTool: Readonly<Ref<"freeform" | "curve" | "scribble" | null>>;
  startFreeformGesture: (event: MouseEvent) => void;
  /** Point-edit mode (SPEC-071): active flag + commit-and-exit, so a click
   * outside the overlay leaves the mode (AC-22). */
  pointEditActive: Readonly<Ref<boolean>>;
  exitPointEdit: (commit?: boolean) => void;
  handleShapeSelect: (section?: number, block?: number, shiftKey?: boolean) => void;
  handleShapeEditText: (section?: number, block?: number) => void;
  shapeAtClientPoint: (clientX: number, clientY: number) => ShapeLocation | null;
  clearShapeSelection: () => void;
  // --- from useConnectors (SPEC-070) ----------------------------------------
  /** Consumes a mousedown when a connector kind is armed (two-click draw);
   * returns true when it handled the event. */
  handleConnectorMouseDown: (event: MouseEvent) => boolean;
  /** Tracks the pointer for the draw rubber band and site hints. */
  handleConnectorMouseMove: (event: MouseEvent) => void;
  /** Clears any connector selection (a plain canvas click deselects). */
  clearConnectorSelection: () => void;
  startShapeMove: (event: MouseEvent, location: ShapeLocation) => void;
  // --- from useCharts (SPEC-052) --------------------------------------------
  handleChartSelect: (section?: number, block?: number) => void;
  handleChartEditData: (section?: number, block?: number) => void;
  chartAtClientPoint: (clientX: number, clientY: number) => ChartLocation | null;
  clearChartSelection: () => void;
  // --- from useNotes (SPEC-031) ---------------------------------------------
  exitNoteEditing: () => void;
  enterNoteEditing: (noteId?: string, kind?: NoteKind, atEnd?: boolean) => void;
  // --- from useImages (SPEC-024) --------------------------------------------
  clearImageSelection: () => void;
  insertImageFromFile: (file: File, selection: ModelSelection | null) => Promise<void>;
  // --- from useLinks (SPEC-025, SPEC-033) -----------------------------------
  clearTocSelection: () => void;
  jumpToAnchor: (anchor?: string) => void;
  // --- from useCaptions (SPEC-055) ------------------------------------------
  /** Clears a selected table of figures — its click-away lifecycle mirrors
     the TOC's clearTocSelection, so a stale selection can't misroute a later
     Delete/Backspace into deleting the whole ToF. */
  clearTofSelection: () => void;
  // --- from useFields (SPEC-053) --------------------------------------------
  followCrossReference: (target: string) => void;
  /** Selects the whole one-character atom of a clicked field (AC-7); returns
     true when the click landed on a (non-cross-reference) field. */
  selectField: (target: EventTarget | null) => boolean;
  openLinkTarget: (target?: LinkTarget | null) => void;
  // --- from useFormatPainter (SPEC-020) -------------------------------------
  maybeApplyPainter: () => void;
  // --- from useTableTools (SPEC-076) ----------------------------------------
  /** Consumes a mousedown that landed on a REPEATED table header row,
     placing the caret in the real header cell it copies (AC-8); returns
     true when it handled the event. */
  redirectRepeatedHeaderClick: (event: MouseEvent) => boolean;
}

/** Whether a drag carries a file of an accepted image type — checked from
 * `dataTransfer.items` (available during dragenter/dragover, unlike the
 * actual file, which only `drop` exposes). Used to steal the drag away from
 * AppShell's window-level "drop to open a document" handling (SPEC-009) only
 * when it is actually an image (AC-1); anything else — including a .docx/
 * .odt file dragged over the canvas — falls through to that existing
 * behaviour untouched. */
function isImageDrag(event = new DragEvent("dragover")) {
  const items = event.dataTransfer?.items;
  if (!items) {
    return false;
  }
  for (let i = 0; i < items.length; i++) {
    if (items[i].kind === "file" && isAcceptedImageType(items[i].type)) {
      return true;
    }
  }
  return false;
}

export function useCanvasGestures(deps: CanvasGesturesDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    applyCaretMove,
    onSessionClose,
    cellSelection,
    tableHover,
    drawGeometry,
    startShapeDraw,
    drawFreeformTool,
    startFreeformGesture,
    pointEditActive,
    exitPointEdit,
    handleShapeSelect,
    handleShapeEditText,
    shapeAtClientPoint,
    clearShapeSelection,
    handleConnectorMouseDown,
    handleConnectorMouseMove,
    clearConnectorSelection,
    startShapeMove,
    handleChartSelect,
    handleChartEditData,
    chartAtClientPoint,
    clearChartSelection,
    exitNoteEditing,
    enterNoteEditing,
    clearImageSelection,
    insertImageFromFile,
    clearTocSelection,
    clearTofSelection,
    jumpToAnchor,
    followCrossReference,
    selectField,
    openLinkTarget,
    maybeApplyPainter,
    zoom,
    clientToContentPx,
  } = deps;

  // --- tables (SPEC-022) ------------------------------------------------------
  // The rectangular cell-range selection (AC-15, AC-16): activated by
  // dragging from one cell into another of the same table; the DOM text
  // selection is suppressed while it is active. `cellDragOrigin` remembers
  // the cell the primary button went down in until the button is released.
  let cellDragOrigin = noTableCellTarget();

  /** Whether the active cell-range selection spans more than one cell — the
   * context menu's "Merge cells" gate (SPEC-023 AC-6). */
  const canMergeCells = computed(() => {
    const range = cellSelection.value;
    if (!range) {
      return false;
    }
    const bounds = cellRangeBounds(range);
    return bounds.top !== bounds.bottom || bounds.left !== bounds.right;
  });

  function clearCellSelection() {
    if (cellSelection.value) {
      cellSelection.value = noCellRange();
    }
  }

  /** The table cell an event target sits in — see lib/domCaret. */
  function tableCellAtTarget(eventTarget = noEventTarget()) {
    const root = contentRoot.value;
    return root ? tableCellFromEventTarget(root, eventTarget) : null;
  }

  /** Clicks on rendered hyperlinks (SPEC-025 AC-12). Browsers do not navigate
   * a plain click on an <a> inside a contenteditable region — the anchors are
   * real links (href/target/rel) purely for fidelity and Ctrl/Cmd+click — but
   * preventDefault keeps that guaranteed; with Ctrl/Cmd held the target opens
   * through the same sanitized path as the chip. */
  function handleContentClick(event = new MouseEvent("click")) {
    const eventTarget = event.target;
    // Clicking a cross-reference scrolls to its target (SPEC-053 AC-26).
    const ref = eventTarget instanceof Element ? eventTarget.closest(".doc-field[data-ref-target]") : null;
    if (ref) {
      event.preventDefault();
      followCrossReference(ref.getAttribute("data-ref-target") ?? "");
      return;
    }
    // Clicking any other field selects its whole atom (SPEC-053 AC-7), since a
    // plain click cannot land the browser caret inside the field's
    // contenteditable=false atom. Only intercept when the selection actually
    // resolved (a body/cell/shape field), so a header/footer field — whose
    // band has its own activation gesture — falls through untouched.
    if (eventTarget instanceof Element && eventTarget.closest(".doc-field") && selectField(eventTarget)) {
      event.preventDefault();
      return;
    }
    // Clicking a note reference mark jumps to its note (SPEC-031 AC-15).
    const mark = eventTarget instanceof Element ? eventTarget.closest(".doc-note-ref[data-note-id]") : null;
    if (mark) {
      event.preventDefault();
      const noteId = mark.getAttribute("data-note-id") ?? "";
      const note = doc.value ? noteById(doc.value, noteId) : undefined;
      if (note) {
        enterNoteEditing(note.noteId, note.kind, true);
      }
      return;
    }
    const anchor = eventTarget instanceof Element ? eventTarget.closest("a[data-doc-link]") : null;
    if (!anchor) {
      return;
    }
    event.preventDefault();
    if (!event.ctrlKey && !event.metaKey) {
      return;
    }
    const href = anchor.getAttribute("href") ?? "";
    if (href.startsWith("#")) {
      jumpToAnchor(href.slice(1));
    } else if (href !== "") {
      openLinkTarget({ href });
    }
  }

  // Tracks the primary button across a text-selection drag (SPEC-020): the
  // format painter must not paint a partial selection mid-drag, only the
  // range the user settles on at mouseup — see maybeApplyPainter.
  let mouseButtonDown = false;

  /** Mousedown on the canvas: remembers the cell under the pointer as a
   * potential drag origin (AC-15) and dismisses any existing cell-range
   * highlight — a plain click always returns to a normal text caret. */
  function handleCanvasMouseDown(event = new MouseEvent("mousedown")) {
    if (event.button !== 0) {
      return;
    }
    // Point-edit mode (SPEC-071 AC-22): a mousedown that reaches the canvas
    // is outside the overlay (its handles/segment stop propagation), so it
    // commits and leaves the mode, keeping the shape selected.
    if (pointEditActive.value) {
      exitPointEdit(true);
      return;
    }
    // Connector draw mode (SPEC-070 AC-3): a gallery pick armed a connector
    // kind, so this mousedown places an endpoint. Takes precedence over every
    // other canvas behaviour, like shape draw mode below.
    if (handleConnectorMouseDown(event)) {
      return;
    }
    // SPEC-076 AC-8: a repeated header row is uneditable, so the browser
    // would drop the caret at whatever editable position lies next to it.
    // Redirect to the real header cell the repeat copies instead.
    if (deps.redirectRepeatedHeaderClick(event)) {
      return;
    }
    // A plain canvas click (not on the connector overlay, which stops
    // propagation) deselects any selected connector.
    clearConnectorSelection();
    // Draw mode (SPEC-051 AC-3, AC-4): a gallery pick armed drawGeometry, so
    // this mousedown draws the shape instead of doing anything else the
    // canvas would normally do with a click.
    if (drawGeometry.value) {
      startShapeDraw(event);
      return;
    }
    // Freeform draw mode (SPEC-071 AC-6/AC-7): a Freeform gallery pick armed
    // a tool, so this mousedown places a vertex or starts a scribble.
    if (drawFreeformTool.value) {
      startFreeformGesture(event);
      return;
    }
    // Clicking the body while a note is being edited leaves note editing
    // (SPEC-031) — clicks inside the note area stop propagation before here.
    exitNoteEditing();
    mouseButtonDown = true;
    clearCellSelection();
    cellDragOrigin = editable.value ? tableCellAtTarget(event.target) : null;
    // A click outside a selected image deselects it (SPEC-024) — a click on
    // one is handled by DocImageBlock's own @select, which fires afterward.
    const imageTarget = event.target instanceof Element ? event.target.closest(".doc-image-anchor") : null;
    if (!imageTarget) {
      clearImageSelection();
    }
    // Same for a selected shape (SPEC-051) — a click on one is handled by
    // DocShapeBlock's own @select, which fires afterward. An unfilled,
    // textless shape's proxy is pointer-events: none over its whole box (AC-11
    // — its own interior is not part of its painted region), so `event.target`
    // never lands inside it; shapeAtClientPoint recovers that case by testing
    // every shape's actual painted region (outline included) at this point.
    const shapeTarget = event.target instanceof Element ? event.target.closest(".doc-shape-anchor") : null;
    let shapeMouseDownLocation = noShapeLocation();
    if (shapeTarget instanceof HTMLElement) {
      const section = Number(shapeTarget.dataset.section);
      const block = Number(shapeTarget.dataset.block);
      if (Number.isInteger(section) && Number.isInteger(block)) {
        shapeMouseDownLocation = { section, block };
      }
    } else {
      const recovered = shapeAtClientPoint(event.clientX, event.clientY);
      if (recovered) {
        handleShapeSelect(recovered.section, recovered.block, event.shiftKey);
        shapeMouseDownLocation = recovered;
      } else {
        clearShapeSelection();
      }
    }
    // A drag on the shape's own body moves (and, crossing a paragraph
    // boundary, re-anchors) it (AC-28) — armed here rather than from
    // DocShapeBlock's own click handler so it also covers a shape that was not
    // already selected (mousedown's native hit-test target is the same
    // element a plain click would land on) and the AC-11 recovery case above.
    // Shift is reserved for the multi-select toggle (AC-34), not a drag.
    if (shapeMouseDownLocation && !event.shiftKey) {
      startShapeMove(event, shapeMouseDownLocation);
    }
    // Same for a selected chart (SPEC-052) — a click on one is handled by
    // DocChartBlock's own @select, which fires afterward. A chart's proxy is
    // always opaque over its full box (never pointer-events: none, unlike an
    // unfilled/textless shape), so chartAtClientPoint recovery is mostly
    // defensive here, kept for consistency with the shape path.
    const chartTarget = event.target instanceof Element ? event.target.closest(".doc-chart-anchor") : null;
    if (!chartTarget) {
      const recoveredChart = chartAtClientPoint(event.clientX, event.clientY);
      if (recoveredChart) {
        handleChartSelect(recoveredChart.section, recoveredChart.block);
      } else {
        clearChartSelection();
      }
    }
    // Same for a selected TOC (SPEC-033) — a click on one is handled by
    // DocTocBlock's own @action "select", which fires afterward.
    const tocTarget = event.target instanceof Element ? event.target.closest(".doc-toc") : null;
    if (!tocTarget) {
      clearTocSelection();
    }
    // Same for a selected table of figures (SPEC-055) — a click on one is
    // handled by DocTofBlock's own @action "select", which fires afterward.
    // Without this parity a stale tofSelection would misroute the next
    // Delete/Backspace into deleting the whole ToF (see DocumentView keydown).
    const tofTarget = event.target instanceof Element ? event.target.closest(".doc-tof") : null;
    if (!tofTarget) {
      clearTofSelection();
    }
  }

  /** Dblclick recovery for a shape DocShapeBlock's own proxy couldn't catch
   * natively (AC-11's pointer-events: none case, same reasoning as
   * handleCanvasMouseDown's shapeAtClientPoint call) — otherwise AC-17's
   * "double-clicking a shape puts the caret inside it, populating Shape.text
   * if absent" would be unreachable by mouse for an unfilled, textless shape
   * whose only clickable region is its outline. */
  function handleCanvasDblClick(event = new MouseEvent("dblclick")) {
    const shapeTarget = event.target instanceof Element ? event.target.closest(".doc-shape-anchor") : null;
    if (shapeTarget) {
      return;
    }
    const recovered = shapeAtClientPoint(event.clientX, event.clientY);
    if (recovered) {
      handleShapeEditText(recovered.section, recovered.block);
      return;
    }
    const chartTarget = event.target instanceof Element ? event.target.closest(".doc-chart-anchor") : null;
    if (chartTarget) {
      return;
    }
    const recoveredChart = chartAtClientPoint(event.clientX, event.clientY);
    if (recoveredChart) {
      handleChartEditData(recoveredChart.section, recoveredChart.block);
    }
  }

  /** Mousemove with the primary button down: entering a different cell of the
   * drag origin's table activates (or extends) the rectangular cell range
   * (AC-15). The DOM text selection is collapsed away — the highlighted cells
   * are the selection's presentation (AC-16). */
  function handleCanvasMouseMove(event = new MouseEvent("mousemove")) {
    // SPEC-070: track the pointer for the connector draw rubber band and the
    // connection-site hints (a no-op unless a kind is armed / an endpoint is
    // being dragged).
    handleConnectorMouseMove(event);
    const origin = cellDragOrigin;
    if (!origin || (event.buttons & 1) === 0) {
      return;
    }
    const cell = tableCellAtTarget(event.target);
    if (!cell || cell.section !== origin.section || cell.block !== origin.block) {
      return;
    }
    const active = cellSelection.value;
    if (!active && cell.row === origin.row && cell.col === origin.col) {
      return;
    }
    event.preventDefault();
    const next = {
      section: origin.section,
      block: origin.block,
      anchor: { row: origin.row, col: origin.col },
      focus: { row: cell.row, col: cell.col },
    };
    const unchanged =
      active && active.focus.row === next.focus.row && active.focus.col === next.focus.col;
    if (!unchanged) {
      cellSelection.value = next;
      const domSelection = document.getSelection();
      if (domSelection && domSelection.rangeCount > 0) {
        domSelection.removeAllRanges();
      }
    }
  }

  function handleCanvasMouseUp() {
    cellDragOrigin = null;
    if (mouseButtonDown) {
      mouseButtonDown = false;
      maybeApplyPainter();
    }
  }

  function handleCanvasDragEnter(event = new DragEvent("dragenter")) {
    if (editable.value && isImageDrag(event)) {
      event.stopPropagation();
    }
  }

  function handleCanvasDragOver(event = new DragEvent("dragover")) {
    if (editable.value && isImageDrag(event)) {
      event.preventDefault();
      event.stopPropagation();
    }
  }

  /** Drag-and-drop onto the canvas (AC-1): inserts at the caret nearest the
   * drop point. */
  function handleCanvasDrop(event = new DragEvent("drop")) {
    if (!editable.value) {
      return;
    }
    const file = event.dataTransfer?.files?.[0];
    if (!file || !isAcceptedImageType(file.type)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const root = contentRoot.value;
    const dropCaret = root ? caretAtClientPoint(root, event.clientX, event.clientY) : null;
    void insertImageFromFile(file, dropCaret ? caretSelection(dropCaret) : null);
  }

  /** Mouseover anywhere on the page: hovering a table measures its geometry
   * for the insert/delete handles (AC-20, AC-21); hovering the handles keeps
   * them; anything else hides them — unless the pointer is still inside the
   * handle band around the measured table. The buttons float up to 20px
   * outside the table's box, so a real mouse must cross bare page on its way
   * to them; hiding on that first crossing unmounted the button before the
   * pointer could arrive, making the handles unclickable by hand. (A
   * synthetic pointer that teleports straight onto a button never samples
   * the gap, which is why e2e clicks always landed.) */
  function handlePageMouseOver(event = new MouseEvent("mouseover")) {
    if (!editable.value || !(event.target instanceof Element)) {
      return;
    }
    if (event.target.closest(".table-handles")) {
      return;
    }
    const table = event.target.closest("table.doc-table");
    const pageEl = pageRef.value;
    if (!table || !pageEl) {
      if (tableHover.value && !withinHandleBand(event, pageEl)) {
        tableHover.value = null;
      }
      return;
    }
    tableHover.value = measureTableHandles(pageEl, table, zoom.value);
  }

  // The farthest handle face sits 20px outside the table (buttons occupy
  // -20..-4); the resize strips reach 3px past the right/bottom edges. The
  // extra slack keeps a wobbly approach path inside the band.
  const HANDLE_BAND_PX = 28;

  /** True while the pointer is within the handle band: the hovered table's
   * box grown by the handles' overhang, in the same page-relative frame
   * measureTableHandles records. */
  function withinHandleBand(event: MouseEvent, pageEl: Element | null): boolean {
    const hover = tableHover.value;
    if (!hover || !pageEl) {
      return false;
    }
    // The pointer and the band overhang are both taken into the unscaled,
    // .pages-relative frame `hover` records, so the hit band lines up with the
    // (zoom-independent) geometry at any zoom (SPEC-063 AC-25).
    const point = clientToContentPx(event.clientX, event.clientY);
    const band = HANDLE_BAND_PX / zoom.value;
    return (
      point.x >= hover.left - band &&
      point.x <= hover.left + hover.width + band &&
      point.y >= hover.top - band &&
      point.y <= hover.top + hover.height + band
    );
  }

  /** Clicking anywhere on the page canvas outside the text (margins, empty
   * space) places the caret at the nearest text position. Clicks inside the
   * editable content keep the browser's native precise placement — except a
   * read-only header/footer band (SPEC-029): it is a `contenteditable="false"`
   * island inside the editable root, so the browser treats it as an atomic
   * non-editable unit rather than falling through to nearby text the way an
   * ordinary margin click does, and needs the same explicit fallback a click
   * outside the root entirely gets. */
  function handlePageMouseDown(event = new MouseEvent("mousedown")) {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !editable.value) {
      return;
    }
    const targetEl = event.target instanceof Element ? event.target : null;
    const inReadOnlyBand = !!targetEl?.closest(
      ".page-header-band:not([data-hf-editable]), .page-footer-band:not([data-hf-editable])",
    );
    if (!inReadOnlyBand && event.target instanceof Node && root.contains(event.target)) {
      return;
    }
    // A multi-click on a read-only band is the gesture that enters
    // header/footer editing (SPEC-029 AC-1), and the browser has already
    // resolved which word (or paragraph) it selects. Claiming the click for
    // the body caret here would preventDefault that selection away, so the
    // band would open with the caret parked at its start instead of the word
    // under the pointer. Only the opening click still places the body caret.
    if (inReadOnlyBand && event.detail > 1) {
      return;
    }
    const target = caretNearClientPoint(root, event.clientX, event.clientY);
    if (!target) {
      return;
    }
    event.preventDefault();
    root.focus();
    applyCaretMove(target, event.shiftKey);
  }

  // The session-close resets the view's closeSession used to run inline.
  onSessionClose(() => {
    clearCellSelection();
    cellDragOrigin = null;
    tableHover.value = null;
  });

  return {
    canMergeCells,
    clearCellSelection,
    tableCellAtTarget,
    /** Live read of the plain drag-tracking variable — the view's
     * handleSelectionChange reads it fresh to skip the painter mid-drag
     * (SPEC-020 AC-12); a getter, like EditorContext's getCaret. */
    isMouseButtonDown: () => mouseButtonDown,
    handleContentClick,
    handleCanvasMouseDown,
    handleCanvasDblClick,
    handleCanvasMouseMove,
    handleCanvasMouseUp,
    handleCanvasDragEnter,
    handleCanvasDragOver,
    handleCanvasDrop,
    handlePageMouseOver,
    handlePageMouseDown,
  };
}
