<script setup lang="ts">
import { computed, nextTick, onMounted, ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// The chart data sheet (SPEC-052 AC-4..AC-13): an editable grid of
// categories and series. Holds its own raw grid (row 0 = corner + series
// names, column 0 = categories, everything else a value) — cell edits, row/
// column insert/delete and pasted TSV/CSV all mutate this raw grid; "First
// row as series names", "First column as categories" and "Switch rows/
// columns" are local toggles that change how the raw grid is INTERPRETED
// into categories/series (deriveChartData below) without touching the grid
// itself, exactly as the spec requires ("changing the chart without
// changing the sheet"). Every interaction re-derives and emits "preview" so
// the caller can re-render the live chart; "Done" commits the same data via
// "apply". BaseDialog owns the shell and Escape; @dismiss maps onto cancel.
const props = defineProps({
  /** { categories: string[], series: {name,color,values}[], singleSeriesOnly: boolean }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["preview", "apply", "cancel"]);

/** A raw grid row: category/series-name text in column 0, numeric text
 * elsewhere. Declared once so every function below infers the same shape
 * from its default value, the espree-parsed-script convention this whole
 * codebase's `.vue` files use in place of type annotations (see
 * ImageOptionsDialog.vue's `event = new Event(...)` defaults). */
const EMPTY_ROW = [""];
const EMPTY_GRID = [EMPTY_ROW];
const DEFAULT_TOGGLES = { firstRowAsNames: true, firstColAsCategories: true, switchRowsColumns: false };

function buildInitialGrid() {
  const categories = props.initial.categories || [];
  const series = props.initial.series || [];
  const header = ["", ...series.map((s = { name: "", values: [] }) => s.name)];
  const rows = categories.map((cat = "", i = 0) => [
    cat,
    ...series.map((s = { name: "", values: [] }) => (s.values[i] === null || s.values[i] === undefined ? "" : String(s.values[i]))),
  ]);
  return [header, ...rows];
}

const grid = ref(buildInitialGrid());
const firstRowAsNames = ref(true);
const firstColAsCategories = ref(true);
const switchRowsColumns = ref(false);
const singleSeriesOnly = !!props.initial.singleSeriesOnly;

function transpose(g = EMPTY_GRID) {
  if (g.length === 0) {
    return g;
  }
  const cols = g[0].length;
  return Array.from({ length: cols }, (_, c) => g.map((row = EMPTY_ROW) => row[c] ?? ""));
}

/** Derives {categories, series} from the raw grid + toggles — pure, called
 * on every edit for the live preview and again on "Done" to commit. */
function deriveChartData(g = EMPTY_GRID, opts = DEFAULT_TOGGLES) {
  const working = opts.switchRowsColumns ? transpose(g) : g;
  if (working.length === 0) {
    return { categories: [], series: [] };
  }
  let seriesNames = [""];
  let dataRows = [EMPTY_ROW];
  if (opts.firstRowAsNames) {
    seriesNames = working[0].slice(1);
    dataRows = working.slice(1);
  } else {
    const colCount = working[0].length - 1;
    seriesNames = Array.from({ length: colCount }, (_, i) => `Series ${i + 1}`);
    dataRows = working;
  }
  const categories = [""];
  categories.length = 0;
  const values = seriesNames.map(() => {
    const empty = [0, null];
    empty.length = 0;
    return empty;
  });
  dataRows.forEach((row = EMPTY_ROW, rowIndex = 0) => {
    categories.push(opts.firstColAsCategories ? row[0] ?? "" : `Category ${rowIndex + 1}`);
    seriesNames.forEach((_name = "", seriesIndex = 0) => {
      const raw = (row[seriesIndex + 1] ?? "").trim();
      const n = Number(raw);
      values[seriesIndex].push(raw === "" ? null : Number.isFinite(n) ? n : null);
    });
  });
  const originalSeries = props.initial.series || [];
  const series = seriesNames.map((name = "", i = 0) => ({
    name,
    values: values[i],
    color: originalSeries[i]?.color,
  }));
  return { categories, series };
}

const toggles = computed(() => ({
  firstRowAsNames: firstRowAsNames.value,
  firstColAsCategories: firstColAsCategories.value,
  switchRowsColumns: switchRowsColumns.value,
}));

function commitPreview() {
  emit("preview", deriveChartData(grid.value, toggles.value));
}

onMounted(() => commitPreview());

function onCellInput(rowIndex = 0, colIndex = 0, event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    grid.value[rowIndex][colIndex] = target.value;
  }
}

function onCellCommit() {
  commitPreview();
}

const colCount = computed(() => grid.value[0]?.length ?? 1);

function addRow() {
  grid.value.push(Array.from({ length: colCount.value }, () => ""));
  commitPreview();
}

function removeRow(rowIndex = 0) {
  if (grid.value.length <= 2 || rowIndex === 0) {
    return;
  }
  grid.value.splice(rowIndex, 1);
  commitPreview();
}

function addColumn() {
  if (singleSeriesOnly && grid.value[0].length >= 2) {
    return;
  }
  for (const row of grid.value) {
    row.push("");
  }
  commitPreview();
}

function removeColumn(colIndex = 0) {
  if (colIndex === 0 || grid.value[0].length <= 2) {
    return;
  }
  for (const row of grid.value) {
    row.splice(colIndex, 1);
  }
  commitPreview();
}

/** Pasted TSV/CSV (AC-10): tab-delimited is checked first (TSV is what a
 * real spreadsheet's copy puts on the clipboard); a single-line paste with
 * commas and no tabs is treated as CSV. Replaces the grid from the focused
 * cell's position, growing it if the pasted block is larger. */
function parseDelimited(text = "") {
  const rows = text.replace(/\r\n/g, "\n").split("\n").filter((line = "", i = 0, arr = [""]) => !(i === arr.length - 1 && line === ""));
  const delimiter = rows.some((row = "") => row.includes("\t")) ? "\t" : ",";
  return rows.map((row = "") => row.split(delimiter));
}

function onPaste(rowIndex = 0, colIndex = 0, event = new ClipboardEvent("paste")) {
  const text = event.clipboardData?.getData("text/plain");
  if (!text || (!text.includes("\t") && !text.includes(",") && !text.includes("\n"))) {
    return; // A single plain value pastes natively into the focused input.
  }
  event.preventDefault();
  const pasted = parseDelimited(text);
  pasted.forEach((row = [""], r = 0) => {
    const targetRow = rowIndex + r;
    while (grid.value.length <= targetRow) {
      grid.value.push(Array.from({ length: colCount.value }, () => ""));
    }
    row.forEach((cell = "", c = 0) => {
      const targetCol = colIndex + c;
      while (grid.value[targetRow].length <= targetCol) {
        for (const gridRow of grid.value) {
          gridRow.push("");
        }
      }
      grid.value[targetRow][targetCol] = cell;
    });
  });
  void nextTick().then(commitPreview);
}

function apply() {
  emit("apply", deriveChartData(grid.value, toggles.value));
}
</script>

<template>
  <BaseDialog
    title="Chart data"
    testid="chart-data-sheet-dialog"
    width="32rem"
    @dismiss="emit('cancel')"
  >
    <p
      v-if="singleSeriesOnly"
      class="dialog-note"
      data-testid="chart-data-single-series-note"
    >
      This chart type renders only the first series.
    </p>

    <div class="toggles">
      <label class="field-inline">
        <input
          v-model="firstRowAsNames"
          type="checkbox"
          data-testid="chart-data-first-row-as-names"
          @change="commitPreview"
        >
        <span>First row as series names</span>
      </label>
      <label class="field-inline">
        <input
          v-model="firstColAsCategories"
          type="checkbox"
          data-testid="chart-data-first-col-as-categories"
          @change="commitPreview"
        >
        <span>First column as categories</span>
      </label>
      <label class="field-inline">
        <input
          v-model="switchRowsColumns"
          type="checkbox"
          data-testid="chart-data-switch-rows-columns"
          @change="commitPreview"
        >
        <span>Switch rows/columns</span>
      </label>
    </div>

    <div class="grid-scroll">
      <table
        class="data-grid"
        data-testid="chart-data-grid"
      >
        <tbody>
          <tr
            v-for="(row, rowIndex) in grid"
            :key="rowIndex"
          >
            <td
              v-for="(cell, colIndex) in row"
              :key="colIndex"
            >
              <input
                class="cell-input"
                :data-testid="`chart-data-cell-${rowIndex}-${colIndex}`"
                :value="cell"
                @input="onCellInput(rowIndex, Number(colIndex), $event)"
                @change="onCellCommit"
                @paste="onPaste(rowIndex, Number(colIndex), $event)"
              >
            </td>
            <td class="row-actions">
              <button
                v-if="rowIndex > 0"
                type="button"
                class="icon-button"
                :data-testid="`chart-data-remove-row-${rowIndex}`"
                title="Delete row"
                @click="removeRow(rowIndex)"
              >
                ✕
              </button>
            </td>
          </tr>
          <tr>
            <td
              v-for="(_, colIndex) in grid[0]"
              :key="colIndex"
              class="col-actions"
            >
              <button
                v-if="Number(colIndex) > 0"
                type="button"
                class="icon-button"
                :data-testid="`chart-data-remove-col-${colIndex}`"
                title="Delete column"
                @click="removeColumn(Number(colIndex))"
              >
                ✕
              </button>
            </td>
            <td />
          </tr>
        </tbody>
      </table>
    </div>

    <div class="grid-add-buttons">
      <button
        type="button"
        data-testid="chart-data-add-row"
        @click="addRow"
      >
        + Row
      </button>
      <button
        type="button"
        data-testid="chart-data-add-col"
        :disabled="singleSeriesOnly && grid[0].length >= 2"
        @click="addColumn"
      >
        + Column
      </button>
    </div>

    <template #actions>
      <button
        type="button"
        data-testid="chart-data-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="chart-data-apply"
        @click="apply"
      >
        Done
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* The grid can be tall, so the panel keeps its own scroll cap (the base's
   .dialog is the short-dialog variant without one). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

.dialog-note {
  margin: 0 0 var(--space-4);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.toggles {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-5);
  margin-bottom: var(--space-4);
}

.field-inline {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-md);
}

.grid-scroll {
  overflow: auto;
  max-height: 40vh;
  border: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-3);
}

.data-grid {
  border-collapse: collapse;
}

.data-grid td {
  border: var(--border-width) solid var(--border-default);
  padding: 0;
}

.cell-input {
  font: inherit;
  font-size: var(--text-md);
  width: 6rem;
  padding: var(--space-2) var(--space-3);
  border: none;
  box-sizing: border-box;
}

.cell-input:focus {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.icon-button {
  font: inherit;
  font-size: var(--text-sm);
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
  padding: var(--space-1) var(--space-3);
}

.icon-button:hover {
  color: var(--text-danger);
}

.grid-add-buttons {
  display: flex;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
}

.grid-add-buttons button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
