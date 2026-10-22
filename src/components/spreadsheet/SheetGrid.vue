<script setup lang="ts">
// The read-only grid (SPEC-116, styled and navigable by SPEC-117): one sheet,
// DOM-virtualised — only the rows and columns in the viewport plus overscan
// exist as elements (useSheetViewport owns the window and the extent).
//
// Structure: the scroller holds a spacer sized to the sheet's extent × zoom
// and one sticky, viewport-sized box; inside it a single content box is
// scaled by the zoom (transform-origin top left) and everything is placed in
// unscaled px: the sheet paper (SheetCellLayer — the file's cells, in every
// app theme), then the selection overlays, the freeze lines and the headers,
// which follow the app theme. Frozen rows and columns (and their corner)
// stay in place while the rest scrolls.
//
// ARIA: a grid (read-only without cell entry) of MAX_ROWS + 1 rows by MAX_COLS + 1 columns
// (headers included), named by the sheet, focusable, with
// aria-activedescendant naming the active cell's element
// (id="sheet-cell-B3"); every rendered cell carries its 1-based
// aria-rowindex/aria-colindex, offset by the header row/column, and
// aria-selected. Keys and pointer gestures go through useSheetGrid
// (useSheetSelection's keyboard map, GridMetrics hit-testing);
// Ctrl+PageUp/Ctrl+PageDown asks the view for the previous/next sheet — a
// convenience only where the browser does not keep the keys; the sheet tabs
// are the keyboard path that always works. SPEC-119: `showFormulas` (View ▸
// Show formulas) has the paper show formula text, and the grid carries
// data-show-formulas="true". SPEC-122: the grid becomes editable — the
// in-cell editor (CellEditor) is mounted over the active cell in its own
// layer above the overlays and holds the grid's focus (focus landing on the
// grid element itself is handed to it), a double-click opens it in Edit
// mode, and a formula's references on this sheet are outlined
// (`sheet-ref-highlight`, `data-ref-index` = the colour index, the colour
// `--sheet-ref-<index + 1>`, the pointed one dashed). SPEC-123: while the
// format painter is armed the grid carries data-format-painter="once|locked"
// and the copy cursor. SPEC-124: every header carries its 0-based
// `data-index` (the header menu reads it); near a header boundary the grid
// shows the col-resize/row-resize cursor, and while a boundary is dragged the
// guide line (`sheet-resize-guide`) and the size tooltip
// (`sheet-resize-tooltip`). SPEC-125: the marching ants (SheetCopyMarquee),
// the fill handle (SheetFillHandle) and the fill and drag-move previews
// (`sheet-fill-preview`, `sheet-move-preview` with the destination's address
// in `sheet-move-tooltip`) among the overlays, the fill options chip
// (FillOptionsChip) above them, and the move cursor over the selection's
// outline. SPEC-128: the paper draws the autofilter's `filter-button`s (a
// press opens the column's popover, the view's useSortFilter through the
// context's `sortFilter`), and while the filter holds a criterion the row
// headers of its data rows carry data-filtered="true" in the accent colour.
// SPEC-130: data validation's chrome sits in its own layer beside the
// active-cell outline, above the editor layer (not among the aria-hidden
// overlays): the input message (ValidationInputMessage, the active cell's
// aria-describedby while it shows) and the dropdown button and its list
// (ValidationDropdown), placed from the active cell's rect in unscaled px so
// they follow scroll and zoom; Alt+Down opens the list (useSheetGrid).
// SPEC-131: the drawing layer sits between the paper and the overlays — one
// SheetDrawingLayer per pane (useDrawingGestures' panes, each a clipping box of
// the cell area) painting the sheet's drawings in unscaled px, so the content
// box's zoom transform scales them; presses on a drawing never reach the
// grid's own handling. While a shape is armed the grid carries
// data-drawing-armed="true" and the crosshair cursor. SPEC-134: `findMarks`
// (useSheetFind's marks for the sheet on show) has the paper mark every
// matching cell data-find-match="true" and the current one
// data-find-current="true"; null — the find bar closed — marks nothing.
// SPEC-135: the paper draws each cell's note and thread indicators
// (useCellNotes' marksAt); the note layer (SheetNoteLayer — the shown notes' boxes, the note editor, the
// hover popover and the active cell's annotation text for its
// aria-describedby) sits above the validation layer, placed from each cell's
// rect in unscaled px; the pointer resting on a cell or a popover goes to the
// notes through useSheetGrid's notePointer. SPEC-135's links: a linked cell
// carries data-link and its text's span the pointer cursor (useCellLinks'
// isLinked); a Ctrl/Cmd+press over that text follows the link instead of
// adding a range; LinkChip pins under the active linked cell.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed, useTemplateRef } from "vue";
import { MAX_COLS, MAX_ROWS } from "../../lib/spreadsheet/model";
import { asSheetGridContext, useSheetGrid } from "../../views/spreadsheet/useSheetGrid";
import { asSheet, COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX } from "../../views/spreadsheet/useSheetViewport";
import { refColor } from "../../views/spreadsheet/cellEditorDom";
import { asFindMarks } from "../../views/spreadsheet/sheetLayerProps";
import { sortFilterOf } from "../../views/spreadsheet/useSortFilter";
import CellEditor from "./CellEditor.vue";
import FillOptionsChip from "./FillOptionsChip.vue";
import SheetCellLayer from "./SheetCellLayer.vue";
import SheetCopyMarquee from "./SheetCopyMarquee.vue";
import SheetDrawingLayer from "./SheetDrawingLayer.vue";
import SheetFillHandle from "./SheetFillHandle.vue";
import SheetNoteLayer from "./SheetNoteLayer.vue";
import LinkChip from "../LinkChip.vue";
import { linksOf } from "../../views/spreadsheet/useCellLinks";
import ValidationDropdown from "./ValidationDropdown.vue";
import ValidationInputMessage from "./ValidationInputMessage.vue";

const props = defineProps({
  /** The grid Sheet on show. */
  sheet: { type: Object, required: true },
  /** The view-level state of the sheet on show (SheetGridContext). */
  context: { type: Object, required: true },
  /** View ▸ Show formulas for the sheet on show (SPEC-119). */
  showFormulas: { type: Boolean, default: false },
  /** SPEC-134: the find bar's marks (sheetLayerProps' FindMarks); absent = none. */
  findMarks: { type: Object, default: undefined },
});

const emit = defineEmits(["cycle-sheet"]);

const scroller = useTemplateRef("scroller");
const grid = useSheetGrid(
  scroller,
  () => asSheet(props.sheet),
  () => asSheetGridContext(props.context),
  (delta = 1) => emit("cycle-sheet", delta),
  () => props.showFormulas,
);
const vp = grid.viewport;
const RH = ROW_HEADER_WIDTH_PX;
const CH = COLUMN_HEADER_HEIGHT_PX;
const cellEditing = computed(() => asSheetGridContext(props.context).cellEditing);
/** SPEC-137: useRecalc's calc state — "calculating" from a pass's start to
 * its end, the benchmark's end-of-recalc signal. */
const calcState = computed(() => asSheetGridContext(props.context).recalc?.calcState?.value ?? "idle");
/** SPEC-127: the function assists' ARIA on the in-cell editor. */
const cellAria = computed(() => asSheetGridContext(props.context).assist?.editorAria.value ?? {});
/** SPEC-123: the armed format painter ("once" | "locked"), "" when off. */
const painterMode = computed(() => asSheetGridContext(props.context).painter?.mode.value ?? "");
/** SPEC-124: header resizing — the cursor, and while dragging the guide line
 * and the size tooltip (unscaled px from the cell area's edge). */
const resizeCursor = computed(() => grid.resize.value?.cursor.value ?? "");
const resizeGuide = computed(() => grid.resize.value?.guide.value ?? null);
const resizeTooltip = computed(() => grid.resize.value?.tooltip.value ?? null);
/** SPEC-125: the ants, the fill handle, the previews, the chip, the cursor. */
const transfer = computed(() => grid.transfer.value);
/** SPEC-134: the find marks, when they are this sheet's. */
const findMarks = computed(() => {
  const marks = asFindMarks(props.findMarks);
  return marks && marks.sheetId === asSheet(props.sheet).id ? marks : null;
});
/** SPEC-128: sort and filter (null: not wired). */
const sortFilter = computed(() => sortFilterOf(props.context));
const filterButton = computed(() => sortFilter.value?.filterButtonAt);

function onFilterButton(col = 0, anchor = { left: 0, top: 0 }) {
  sortFilter.value?.pressFilterButton(col, anchor);
}

/** SPEC-135: links (null: not wired) — the paper's data-link, a
 * Ctrl/Cmd+press over a linked cell's text following its link ahead of the
 * pointer's own handling (SPEC-117's add-a-range), and the active cell's
 * chip pinned under it in viewport px (LinkChip is fixed; teleported, so its
 * presses never reach the grid), hidden while the cell is out of view. */
const links = computed(() => linksOf(props.context));
function onPointerDown(event = new PointerEvent("pointerdown")) {
  if (!links.value?.pressFollows(event)) {
    grid.pointer.onPointerDown(event);
  }
}
const linkChip = computed(() => {
  const chip = links.value?.chip.value;
  const rect = grid.overlays.activeRect.value;
  const box = chip && rect ? scroller.value?.getBoundingClientRect() : undefined;
  if (!chip || !rect || !box || rect.top + rect.height > vp.viewRows.value) {
    return null;
  }
  const z = vp.z.value;
  return { ...chip, left: box.left + (RH + rect.left) * z, top: box.top + (CH + rect.top + rect.height) * z + 2 };
});

/** SPEC-131: the drawing layer (null: no drawings wired) and its events. */
const drawingLayer = computed(() => grid.drawingGestures.layer.value);
const drawings = computed(() => asSheetGridContext(props.context).drawings ?? null);
function drawingKey(id = "", event = new KeyboardEvent("keydown")) {
  drawings.value?.onDrawingKeydown(id, event);
}
function drawingMenu(id = "", x = 0, y = 0) {
  drawings.value?.openContextMenu(id, x, y);
}
function drawingFocus(id = "") {
  drawings.value?.select(id, false);
}

/** SPEC-135: notes (null: not wired). */
const notes = computed(() => asSheetGridContext(props.context).notes ?? null);

/** SPEC-130: the dropdown's events, to data validation. */
const validationChrome = computed(() => grid.validationChrome.value);
function toggleList() {
  grid.validation.value?.toggleList();
}
function selectItem(index = 0) {
  grid.validation.value?.selectItem(index);
}
function pickItem(index = 0) {
  grid.validation.value?.pick(index);
}
function closeList(focusGrid = false) {
  grid.validation.value?.closeList(focusGrid);
}

/** A row header's data-filtered (SPEC-128). */
function filtered(row = 0) {
  return sortFilter.value?.isRowFiltered(row) ? "true" : undefined;
}

function px(value = 0) {
  return `${value}px`;
}

/** Focus on the grid element itself goes on to the in-cell editor. */
function onFocus(event = new FocusEvent("focus")) {
  if (event.target === scroller.value && cellEditing.value) {
    grid.focusEditor();
  }
}

/** The browser scrolls a clipped box to show a caret; the layout never
 * scrolls these boxes, so any such scroll is undone. */
function unscroll(event = new Event("scroll")) {
  const box = event.target;
  if (box instanceof HTMLElement && (box.scrollTop !== 0 || box.scrollLeft !== 0)) {
    box.scrollTop = 0;
    box.scrollLeft = 0;
  }
}
</script>

<template>
  <div
    ref="scroller"
    class="sheet-grid"
    role="grid"
    :aria-readonly="cellEditing ? undefined : 'true'"
    aria-multiselectable="true"
    :aria-rowcount="MAX_ROWS + 1"
    :aria-colcount="MAX_COLS + 1"
    :aria-label="props.sheet.name"
    :aria-activedescendant="grid.activeDescendant.value"
    :data-show-formulas="props.showFormulas ? 'true' : undefined"
    :data-calc-state="calcState"
    :data-format-painter="painterMode || undefined"
    :data-drawing-armed="drawingLayer?.armed ? 'true' : undefined"
    :style="{ cursor: resizeCursor || transfer.cursor || undefined }"
    data-testid="spreadsheet-grid"
    tabindex="0"
    @scroll="vp.onScroll"
    @focus="onFocus"
    @keydown="grid.onKeydown"
    @dblclick="grid.onDblClick"
    @wheel="grid.onWheel"
    @pointerdown="onPointerDown"
    @pointermove="grid.pointer.onPointerMove"
    @pointerup="grid.pointer.onPointerUp"
    @pointercancel="grid.pointer.onPointerUp"
    @pointerover="grid.notePointer.over"
    @pointerleave="grid.notePointer.leave"
  >
    <div
      class="grid-area"
      :style="grid.areaStyle.value"
    >
      <div
        class="grid-viewport"
        :style="{ width: px(vp.viewW.value * vp.z.value), height: px(vp.viewH.value * vp.z.value) }"
        @scroll="unscroll"
      >
        <div
          class="grid-content"
          :style="grid.contentStyle.value"
        >
          <SheetCellLayer
            :bands="grid.bands.value"
            :grid-lines="grid.showGridLines.value"
            :frozen-width="vp.frozenW.value"
            :stale-title="grid.staleTitle"
            :filter-button="filterButton"
            :described-by="grid.describedBy.value ?? undefined"
            :find-marks="findMarks ?? undefined"
            :annotation-marks="notes?.marksAt"
            :link-at="links?.isLinked"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
            @filter-button="onFilterButton"
          />
          <div
            v-if="drawingLayer"
            class="sheet-drawings"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
          >
            <SheetDrawingLayer
              v-for="pane in drawingLayer.panes"
              :key="pane.id"
              :pane="pane"
              :zoom="drawingLayer.zoom"
              :theme="drawingLayer.theme"
              :image-src="drawingLayer.imageSrc"
              :draw-box="drawingLayer.drawBox ?? undefined"
              @press="grid.drawingGestures.pressDrawing"
              @handle-press="grid.drawingGestures.pressHandle"
              @key="drawingKey"
              @menu="drawingMenu"
              @focus-drawing="drawingFocus"
            />
          </div>
          <div
            class="sheet-overlays"
            aria-hidden="true"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
          >
            <div
              v-for="rect in grid.overlays.selectionRects.value"
              :key="rect.key"
              class="sheet-selection"
              data-testid="sheet-selection"
              :style="{ left: px(rect.left), top: px(rect.top), width: px(rect.width), height: px(rect.height) }"
            />
            <div
              v-if="grid.overlays.activeRect.value"
              class="sheet-active-cell"
              data-testid="sheet-active-cell"
              :style="{
                left: px(grid.overlays.activeRect.value.left),
                top: px(grid.overlays.activeRect.value.top),
                width: px(grid.overlays.activeRect.value.width),
                height: px(grid.overlays.activeRect.value.height),
              }"
            />
            <div
              v-for="mark in grid.overlays.refHighlights.value"
              :key="mark.key"
              class="sheet-ref-highlight"
              :class="{ 'is-dashed': mark.dashed }"
              data-testid="sheet-ref-highlight"
              :data-ref-index="mark.colorIndex"
              :style="{ left: px(mark.left), top: px(mark.top), width: px(mark.width), height: px(mark.height), borderColor: refColor(mark.colorIndex) }"
            />
            <SheetCopyMarquee
              v-if="transfer.ants"
              v-bind="transfer.ants"
            />
            <div
              v-if="transfer.fillPreview"
              class="sheet-transfer-preview"
              data-testid="sheet-fill-preview"
              :style="{ left: px(transfer.fillPreview.left), top: px(transfer.fillPreview.top), width: px(transfer.fillPreview.width), height: px(transfer.fillPreview.height) }"
            />
            <div
              v-if="transfer.movePreview"
              class="sheet-transfer-preview"
              data-testid="sheet-move-preview"
              :title="transfer.movePreview.address"
              :style="{ left: px(transfer.movePreview.left), top: px(transfer.movePreview.top), width: px(transfer.movePreview.width), height: px(transfer.movePreview.height) }"
            />
            <SheetFillHandle
              v-if="transfer.handle"
              :left="transfer.handle.left"
              :top="transfer.handle.top"
            />
            <div
              v-if="grid.overlays.freezeLines.value.vertical !== null"
              class="sheet-freeze-line is-vertical"
              data-testid="sheet-freeze-line"
              :style="{ left: px(grid.overlays.freezeLines.value.vertical - 1) }"
            />
            <div
              v-if="grid.overlays.freezeLines.value.horizontal !== null"
              class="sheet-freeze-line is-horizontal"
              data-testid="sheet-freeze-line"
              :style="{ top: px(grid.overlays.freezeLines.value.horizontal - 1) }"
            />
          </div>
          <div
            v-if="cellEditing"
            class="sheet-editor-layer"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
            @scroll="unscroll"
          >
            <CellEditor
              :editor="cellEditing"
              :aria="cellAria"
              :box-style="grid.editor.value.boxStyle"
              :text-style="grid.editor.value.textStyle"
              :label="grid.editor.value.label"
            />
          </div>
          <div
            v-if="validationChrome"
            class="sheet-validation-layer"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
            @scroll="unscroll"
          >
            <ValidationInputMessage
              v-if="validationChrome.message"
              :title="validationChrome.message.title"
              :text="validationChrome.message.text"
              :cell-rect="validationChrome.rect"
              :bounds="validationChrome.bounds"
            />
            <ValidationDropdown
              v-if="validationChrome.dropdown"
              :cell-rect="validationChrome.rect"
              :bounds="validationChrome.bounds"
              :expanded="validationChrome.dropdown.expanded"
              :items="validationChrome.list?.texts ?? []"
              :selected="validationChrome.list?.selected ?? -1"
              :unavailable="validationChrome.list?.unavailable ?? false"
              @toggle="toggleList"
              @select="selectItem"
              @pick="pickItem"
              @close="closeList(true)"
              @dismiss="closeList(false)"
            />
          </div>
          <SheetNoteLayer
            v-if="notes"
            :controller="notes"
            :anchor="grid.noteAnchor"
            :bounds="{ width: vp.viewCols.value, height: vp.viewRows.value }"
            :style="{ left: px(RH), top: px(CH), width: px(vp.viewCols.value), height: px(vp.viewRows.value) }"
          />
          <div
            class="sheet-col-headers"
            role="row"
            aria-rowindex="1"
            :style="{ width: px(vp.viewW.value), height: px(CH) }"
          >
            <div
              class="sheet-select-all"
              data-testid="sheet-select-all"
              title="Select all"
              aria-hidden="true"
              :style="{ width: px(RH), height: px(CH) }"
            />
            <div
              class="sheet-header-region"
              :style="{ left: px(RH), width: px(vp.frozenW.value), height: px(CH) }"
            >
              <div
                v-for="col in grid.overlays.columnHeaders.value.frozen"
                :key="col.index"
                class="col-header"
                :class="{ 'is-selected': col.selected }"
                role="columnheader"
                data-testid="sheet-col-header"
                :data-index="col.index"
                :aria-colindex="col.index + 2"
                :style="{ left: px(col.start), width: px(col.size) }"
              >
                {{ col.label }}
              </div>
            </div>
            <div
              class="sheet-header-region"
              :style="{ left: px(RH + vp.frozenW.value), width: px(Math.max(0, vp.viewCols.value - vp.frozenW.value)), height: px(CH) }"
            >
              <div
                v-for="col in grid.overlays.columnHeaders.value.scrolled"
                :key="col.index"
                class="col-header"
                :class="{ 'is-selected': col.selected }"
                role="columnheader"
                data-testid="sheet-col-header"
                :data-index="col.index"
                :aria-colindex="col.index + 2"
                :style="{ left: px(col.start), width: px(col.size) }"
              >
                {{ col.label }}
              </div>
            </div>
            <div
              v-for="marker in grid.overlays.hiddenMarkers.value.columns"
              :key="marker.key"
              class="sheet-hidden-marker is-column"
              data-testid="sheet-hidden-marker"
              :title="marker.title"
              :aria-label="marker.title"
              :style="{ left: px(RH + marker.at - 3) }"
            />
          </div>
          <div
            class="sheet-row-headers"
            :style="{ top: px(CH), width: px(RH), height: px(vp.viewRows.value) }"
          >
            <div
              class="sheet-header-region"
              :style="{ top: '0px', width: px(RH), height: px(vp.frozenH.value) }"
            >
              <div
                v-for="row in grid.overlays.rowHeaders.value.frozen"
                :key="row.index"
                class="row-header"
                :class="{ 'is-selected': row.selected }"
                role="rowheader"
                data-testid="sheet-row-header"
                :data-index="row.index"
                :data-filtered="filtered(row.index)"
                aria-colindex="1"
                :style="{ top: px(row.start), height: px(row.size), lineHeight: px(row.size - 1) }"
              >
                {{ row.label }}
              </div>
            </div>
            <div
              class="sheet-header-region"
              :style="{ top: px(vp.frozenH.value), width: px(RH), height: px(Math.max(0, vp.viewRows.value - vp.frozenH.value)) }"
            >
              <div
                v-for="row in grid.overlays.rowHeaders.value.scrolled"
                :key="row.index"
                class="row-header"
                :class="{ 'is-selected': row.selected }"
                role="rowheader"
                data-testid="sheet-row-header"
                :data-index="row.index"
                :data-filtered="filtered(row.index)"
                aria-colindex="1"
                :style="{ top: px(row.start), height: px(row.size), lineHeight: px(row.size - 1) }"
              >
                {{ row.label }}
              </div>
            </div>
            <div
              v-for="marker in grid.overlays.hiddenMarkers.value.rows"
              :key="marker.key"
              class="sheet-hidden-marker is-row"
              data-testid="sheet-hidden-marker"
              :title="marker.title"
              :aria-label="marker.title"
              :style="{ top: px(marker.at - 3) }"
            />
          </div>
          <div
            v-if="resizeGuide"
            class="sheet-resize-guide"
            :class="resizeGuide.axis === 'col' ? 'is-vertical' : 'is-horizontal'"
            data-testid="sheet-resize-guide"
            aria-hidden="true"
            :style="resizeGuide.axis === 'col' ? { left: px(RH + resizeGuide.at), height: px(vp.viewH.value) } : { top: px(CH + resizeGuide.at), width: px(vp.viewW.value) }"
          />
          <div
            v-if="transfer.movePreview"
            class="sheet-resize-tooltip"
            role="status"
            data-testid="sheet-move-tooltip"
            :style="{ left: px(RH + transfer.movePreview.left), top: px(CH + transfer.movePreview.top + transfer.movePreview.height + 4) }"
          >
            {{ transfer.movePreview.address }}
          </div>
          <FillOptionsChip
            v-if="transfer.chip && grid.fillChip.value"
            :x="RH + transfer.chip.x"
            :y="CH + transfer.chip.y"
            :applied="transfer.chip.applied"
            @choose="grid.fillChip.value.choose"
          />
          <div
            v-if="resizeTooltip"
            class="sheet-resize-tooltip"
            role="status"
            data-testid="sheet-resize-tooltip"
            :style="{ left: px(resizeTooltip.x), top: px(resizeTooltip.y) }"
          >
            {{ resizeTooltip.text }}
          </div>
        </div>
      </div>
      <div
        class="grid-canvas"
        :style="grid.canvasStyle.value"
      />
    </div>
    <Teleport to="body">
      <LinkChip
        v-if="links && linkChip"
        :left="linkChip.left"
        :top="linkChip.top"
        :label="linkChip.label"
        :external="linkChip.external"
        @open="links.openFromChip()"
        @edit="links.editFromChip()"
        @copy="links.copyFromChip()"
        @remove="links.removeFromChip()"
        @dismiss="links.dismissChip()"
      />
    </Teleport>
  </div>
</template>

<style scoped>
/* Headers, overlays and freeze lines follow the app theme (tokens); the
   paper underneath is SheetCellLayer's. Geometry is unscaled px inside the
   scaled content box. */
.sheet-grid {
  position: relative;
  height: 100%;
  overflow: auto;
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
}

.sheet-grid:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: -2px;
}

.grid-area {
  position: relative;
}

/* The viewport-sized box that does not scroll: everything is placed in it
   from the scroll position. */
.grid-viewport {
  position: sticky;
  top: 0;
  left: 0;
  z-index: 1;
  overflow: hidden;
}

.grid-content {
  position: absolute;
  top: 0;
  left: 0;
  transform-origin: 0 0;
}

/* The spacer: the sheet's extent (headers excluded), behind the viewport. */
.grid-canvas {
  position: absolute;
  pointer-events: none;
}

.sheet-overlays {
  position: absolute;
  overflow: hidden;
  pointer-events: none;
}

/* SPEC-131: the drawing panes' box, over the paper and under the overlays;
   only the drawings take the pointer. */
.sheet-drawings {
  position: absolute;
  overflow: hidden;
  pointer-events: none;
}

.sheet-selection {
  position: absolute;
  box-sizing: border-box;
  background: color-mix(in srgb, var(--accent) 12%, transparent);
  border: var(--border-width) solid var(--accent);
}

/* A formula's reference (SPEC-122): outlined in its --sheet-ref-N colour
   (set inline), the one being pointed at dashed. */
.sheet-ref-highlight {
  position: absolute;
  box-sizing: border-box;
  border: var(--space-1) solid;
}

.sheet-ref-highlight.is-dashed {
  border-style: dashed;
}

/* The in-cell editor's layer: over the paper and the overlays, under the
   headers; it only clips — the editor takes the pointer while editing. */
.sheet-editor-layer {
  position: absolute;
  z-index: 1;
  overflow: hidden;
  pointer-events: none;
}

/* SPEC-130: data validation's input message and dropdown, over the editor
   layer and under the headers; only the button and the list take the
   pointer. */
.sheet-validation-layer {
  position: absolute;
  z-index: 1;
  overflow: hidden;
  pointer-events: none;
}

/* The active cell: a 2 px (--space-1) accent outline. */
.sheet-active-cell {
  position: absolute;
  box-sizing: border-box;
  border: var(--space-1) solid var(--accent);
}

.sheet-freeze-line {
  position: absolute;
  background: var(--border-default);
}

/* SPEC-125: the target of a fill drag and the destination of a drag-move. */
.sheet-transfer-preview {
  position: absolute;
  box-sizing: border-box;
  border: var(--space-1) dashed var(--accent);
}

/* SPEC-124: the boundary being dragged, across the whole viewport, and its
   size tooltip. */
.sheet-resize-guide {
  position: absolute;
  z-index: 6;
  pointer-events: none;
  border: 0 dashed var(--accent);
}

.sheet-resize-guide.is-vertical {
  top: 0;
  width: 0;
  border-left-width: var(--border-width);
}

.sheet-resize-guide.is-horizontal {
  left: 0;
  height: 0;
  border-top-width: var(--border-width);
}

.sheet-resize-tooltip {
  position: absolute;
  z-index: 7;
  padding: var(--space-1) var(--space-2);
  white-space: nowrap;
  pointer-events: none;
  font-size: var(--text-sm);
  font-variant-numeric: tabular-nums;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.sheet-freeze-line.is-vertical {
  top: 0;
  bottom: 0;
  width: 1px;
}

.sheet-freeze-line.is-horizontal {
  left: 0;
  right: 0;
  height: 1px;
}

.sheet-col-headers,
.sheet-row-headers {
  position: absolute;
  top: 0;
  left: 0;
  overflow: hidden;
  background: var(--surface-sunken);
}

.sheet-col-headers {
  z-index: 2;
  border-bottom: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

.sheet-row-headers {
  z-index: 2;
  border-right: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

.sheet-select-all {
  position: absolute;
  top: 0;
  left: 0;
  z-index: 1;
  box-sizing: border-box;
  background: var(--surface-sunken);
  border-right: var(--border-width) solid var(--border-default);
  cursor: cell;
}

.sheet-header-region {
  position: absolute;
  overflow: hidden;
}

.col-header,
.row-header {
  position: absolute;
  color: var(--text-secondary);
  font-size: var(--text-xs);
  text-align: center;
  overflow: hidden;
  white-space: nowrap;
  box-sizing: border-box;
  user-select: none;
}

.col-header {
  top: 0;
  height: 100%;
  line-height: 19px;
  border-right: var(--border-width) solid var(--border-subtle);
}

.row-header {
  left: 0;
  width: 100%;
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.col-header.is-selected,
.row-header.is-selected {
  background: color-mix(in srgb, var(--accent) 16%, transparent);
  color: var(--text-primary);
  font-weight: var(--weight-medium);
}

/* SPEC-128: the row numbers of a filtered range's data rows. */
.row-header[data-filtered="true"] {
  color: var(--text-accent);
}

/* A hidden run's header boundary: a short double tick where the run was. */
.sheet-hidden-marker {
  position: absolute;
  z-index: 2;
  box-sizing: border-box;
  border: var(--border-width) solid var(--text-secondary);
}

.sheet-hidden-marker.is-column {
  top: 3px;
  width: 6px;
  height: 14px;
  border-top: none;
  border-bottom: none;
}

.sheet-hidden-marker.is-row {
  left: 3px;
  width: 50px;
  height: 6px;
  border-left: none;
  border-right: none;
}

/* SPEC-131: an armed shape's crosshair, over every grid layer. */
.sheet-grid[data-drawing-armed],
.sheet-grid[data-drawing-armed] :deep(*) {
  cursor: crosshair;
}

/* SPEC-123: the armed format painter's cursor, over every grid layer. */
.sheet-grid[data-format-painter],
.sheet-grid[data-format-painter] :deep(*) {
  cursor: copy;
}
</style>
