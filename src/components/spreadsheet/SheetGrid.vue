<script setup lang="ts">
// The read-only grid (SPEC-116): one sheet's values, DOM-virtualised — only
// the rows and columns in the viewport plus overscan exist as elements
// (useSheetViewport owns the window and the scroll extent), positioned
// absolutely inside a canvas sized to the extent, under sticky lettered
// column headers and numbered row headers. A value is a text node through
// the General display, clipped to its 64 × 20 px cell; each row paints its
// own gridlines, so an empty cell needs no element at all.
//
// ARIA: a read-only grid of MAX_ROWS + 1 rows by MAX_COLS + 1 columns
// (headers included), named by the sheet; every rendered cell carries its
// 1-based aria-rowindex/aria-colindex, offset by the header row/column.
// Ctrl+PageUp/Ctrl+PageDown asks the view for the previous/next sheet — a
// convenience only where the browser does not keep the keys for its own tab
// switching; the sheet tabs are the keyboard path that always works.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { useTemplateRef } from "vue";
import { MAX_COLS, MAX_ROWS } from "../../lib/spreadsheet/model";
import { asSheet, useSheetViewport } from "../../views/spreadsheet/useSheetViewport";

const props = defineProps({
  /** The grid Sheet on show. */
  sheet: { type: Object, required: true },
});

const emit = defineEmits(["cycle-sheet"]);

const scroller = useTemplateRef("scroller");
const viewport = useSheetViewport(scroller, () => asSheet(props.sheet));

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (!event.ctrlKey || event.altKey || event.metaKey || event.shiftKey) {
    return;
  }
  if (event.key === "PageDown" || event.key === "PageUp") {
    event.preventDefault();
    emit("cycle-sheet", event.key === "PageDown" ? 1 : -1);
  }
}
</script>

<template>
  <div
    ref="scroller"
    class="sheet-grid"
    role="grid"
    aria-readonly="true"
    :aria-rowcount="MAX_ROWS + 1"
    :aria-colcount="MAX_COLS + 1"
    :aria-label="props.sheet.name"
    data-testid="spreadsheet-grid"
    tabindex="0"
    @scroll="viewport.onScroll"
    @keydown="onKeydown"
  >
    <div
      class="grid-head"
      role="row"
      aria-rowindex="1"
      :style="{ width: `${viewport.canvasWidth.value}px` }"
    >
      <div
        class="grid-corner"
        aria-hidden="true"
      />
      <div
        v-for="col in viewport.visibleCols.value"
        :key="col.index"
        class="col-header"
        role="columnheader"
        :aria-colindex="col.index + 2"
        :style="{ left: `${col.left}px` }"
      >
        {{ col.letters }}
      </div>
    </div>
    <div
      class="grid-canvas"
      :style="{ width: `${viewport.canvasWidth.value}px`, height: `${viewport.canvasHeight.value}px` }"
    >
      <div
        v-for="row in viewport.visibleRows.value"
        :key="row.index"
        class="grid-row"
        role="row"
        :aria-rowindex="row.index + 2"
        :style="{ top: `${row.top}px` }"
      >
        <div
          class="row-header"
          role="rowheader"
          aria-colindex="1"
        >
          {{ row.index + 1 }}
        </div>
        <div
          v-for="cell in row.cells"
          :key="cell.col"
          class="sheet-cell"
          :class="`is-${cell.align}`"
          role="gridcell"
          :aria-rowindex="row.index + 2"
          :aria-colindex="cell.col + 2"
          data-testid="sheet-cell"
          :data-ref="cell.ref"
          :style="{ left: `${cell.left}px` }"
        >
          {{ cell.text }}
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
/* The geometry here is useSheetViewport's: 64 × 20 px cells, a 56 px row
   header and a 20 px letter row. */
.sheet-grid {
  position: relative;
  height: 100%;
  overflow: auto;
  background: var(--surface-raised);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
}

.sheet-grid:focus-visible {
  outline: 2px solid var(--focus-ring);
  outline-offset: -2px;
}

.grid-head {
  position: sticky;
  top: 0;
  z-index: 2;
  height: 20px;
  background: var(--surface-sunken);
  border-bottom: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

.grid-corner {
  position: sticky;
  left: 0;
  z-index: 3;
  width: 56px;
  height: 100%;
  background: var(--surface-sunken);
  border-right: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

.col-header,
.row-header {
  color: var(--text-secondary);
  font-size: var(--text-xs);
  line-height: 19px;
  text-align: center;
  overflow: hidden;
  white-space: nowrap;
  box-sizing: border-box;
}

.col-header {
  position: absolute;
  top: 0;
  width: 64px;
  height: 100%;
  border-right: var(--border-width) solid var(--border-subtle);
}

.grid-canvas {
  position: relative;
}

/* One row: its own gridlines (a bottom border, and a column line every
   64 px past the row header), so empty cells need no elements. */
.grid-row {
  position: absolute;
  left: 0;
  width: 100%;
  height: 20px;
  border-bottom: var(--border-width) solid var(--border-subtle);
  box-sizing: border-box;
  background-image: linear-gradient(to right, transparent 63px, var(--border-subtle) 63px);
  background-position: 56px 0;
  background-size: 64px 100%;
  background-repeat: repeat-x;
}

.row-header {
  position: sticky;
  left: 0;
  z-index: 1;
  width: 56px;
  height: 100%;
  background: var(--surface-sunken);
  border-right: var(--border-width) solid var(--border-default);
}

.sheet-cell {
  position: absolute;
  top: 0;
  width: 64px;
  height: 100%;
  padding: 0 3px;
  overflow: hidden;
  white-space: pre;
  line-height: 19px;
  box-sizing: border-box;
}

.sheet-cell.is-number {
  text-align: right;
}

.sheet-cell.is-text {
  text-align: left;
}

.sheet-cell.is-centred {
  text-align: center;
}
</style>
