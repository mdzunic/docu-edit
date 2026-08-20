<script setup lang="ts">
// The table hover handles (SPEC-022 AC-20, AC-21): a page-relative overlay
// of small insert (+) and delete (×) buttons above the columns and beside
// the rows of the hovered table, plus the column/row resize strips
// (SPEC-023 AC-1..5). Purely presentational: the view measures the geometry
// (caretGeometry.measureTableHandles) and owns every action — the buttons
// emit `handle` with the same (kind, index) payload handleTableHandle
// takes, and the resize strips forward their raw mousedown MouseEvent ahead
// of the index so startColumnResize/startRowResize see it unchanged.
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
defineProps({
  /** The measured caretGeometry.TableHandleGeometry of the hovered table. */
  geometry: { type: Object, required: true },
});

const emit = defineEmits(["handle", "column-resize", "row-resize"]);
</script>

<template>
  <div
    class="table-handles print-exclude"
    data-testid="table-handles"
  >
    <!-- Column handles above the table: an insert button at every
       column boundary, a delete button over every column centre
       (AC-20, AC-21). -->
    <button
      v-for="boundary in geometry.cols.length + 1"
      :key="`ci${boundary}`"
      type="button"
      class="table-handle insert"
      :data-testid="`table-handle-insert-col-${boundary - 1}`"
      :aria-label="`Insert column at position ${boundary}`"
      :style="{
        left: `${(boundary - 1 < geometry.cols.length
          ? geometry.cols[boundary - 1].left
          : geometry.left + geometry.width) - 8}px`,
        top: `${geometry.top - 20}px`,
      }"
      @mousedown.prevent
      @click="emit('handle', 'insert-col', boundary - 1)"
    >
      +
    </button>
    <button
      v-for="(col, colIndex) in geometry.cols"
      :key="`cd${colIndex}`"
      type="button"
      class="table-handle delete"
      :data-testid="`table-handle-delete-col-${colIndex}`"
      :aria-label="`Delete column ${colIndex + 1}`"
      :style="{
        left: `${col.left + col.width / 2 - 8}px`,
        top: `${geometry.top - 20}px`,
      }"
      @mousedown.prevent
      @click="emit('handle', 'delete-col', colIndex)"
    >
      ×
    </button>
    <!-- Row handles left of the table: same pattern. -->
    <button
      v-for="boundary in geometry.rows.length + 1"
      :key="`ri${boundary}`"
      type="button"
      class="table-handle insert"
      :data-testid="`table-handle-insert-row-${boundary - 1}`"
      :aria-label="`Insert row at position ${boundary}`"
      :style="{
        left: `${geometry.left - 20}px`,
        top: `${(boundary - 1 < geometry.rows.length
          ? geometry.rows[boundary - 1].top
          : geometry.top + geometry.height) - 8}px`,
      }"
      @mousedown.prevent
      @click="emit('handle', 'insert-row', boundary - 1)"
    >
      +
    </button>
    <button
      v-for="(row, rowIndex) in geometry.rows"
      :key="`rd${rowIndex}`"
      type="button"
      class="table-handle delete"
      :data-testid="`table-handle-delete-row-${rowIndex}`"
      :aria-label="`Delete row ${rowIndex + 1}`"
      :style="{
        left: `${geometry.left - 20}px`,
        top: `${row.top + row.height / 2 - 8}px`,
      }"
      @mousedown.prevent
      @click="emit('handle', 'delete-row', rowIndex)"
    >
      ×
    </button>
    <!-- Column/row resize strips (SPEC-023 AC-1..5): a thin draggable
       strip along each column's right edge (spanning the table's
       height) and each row's bottom edge (spanning its width). -->
    <div
      v-for="(col, colIndex) in geometry.cols"
      :key="`cr${colIndex}`"
      class="table-resize-handle column"
      :data-testid="`table-resize-col-${colIndex}`"
      :style="{
        left: `${col.left + col.width - 3}px`,
        top: `${geometry.top}px`,
        height: `${geometry.height}px`,
      }"
      @mousedown="emit('column-resize', $event, colIndex)"
    />
    <div
      v-for="(row, rowIndex) in geometry.rows"
      :key="`rr${rowIndex}`"
      class="table-resize-handle row"
      :data-testid="`table-resize-row-${rowIndex}`"
      :style="{
        left: `${geometry.left}px`,
        top: `${row.top + row.height - 3}px`,
        width: `${geometry.width}px`,
      }"
      @mousedown="emit('row-resize', $event, rowIndex)"
    />
  </div>
</template>

<style scoped>
/* The table hover handles (SPEC-022 AC-20, AC-21): a page-relative overlay
   of small insert (+) and delete (×) buttons above the columns and beside
   the rows of the hovered table. The layer itself ignores the pointer so
   text editing underneath is unaffected; only the buttons are clickable. */
.table-handles {
  position: absolute;
  inset: 0;
  z-index: 15;
  pointer-events: none;
}

.table-handle {
  position: absolute;
  pointer-events: auto;
  width: 16px;
  height: 16px;
  padding: 0;
  border: 1px solid var(--border-strong);
  border-radius: 50%;
  background: var(--surface-base);
  color: var(--text-primary);
  font-size: 11px;
  line-height: 1;
  cursor: pointer;
  box-shadow: var(--shadow-1);
}

.table-handle.insert:hover {
  background: var(--accent-subtle);
  border-color: var(--accent);
  color: var(--accent);
}

.table-handle.delete:hover {
  background: #fce8e6;
  border-color: var(--danger);
  color: var(--danger);
}

/* Column/row resize strips (SPEC-023 AC-1..5): thin transparent bands
   along a column's right edge or a row's bottom edge, widened just enough
   to be a comfortable drag target without covering neighbouring cells. */
.table-resize-handle {
  position: absolute;
  pointer-events: auto;
  background: transparent;
}

.table-resize-handle.column {
  width: 6px;
  cursor: col-resize;
}

.table-resize-handle.row {
  height: 6px;
  cursor: row-resize;
}

.table-resize-handle:hover {
  background: var(--accent-subtle);
}
</style>
