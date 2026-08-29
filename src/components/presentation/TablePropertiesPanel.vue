<script setup lang="ts">
// The Table properties panel (SPEC-091 AC-10): edits the selected table's
// cells (fill, borders, padding, vertical alignment) and its table-level look
// (header row, banded rows), plus the row/column/merge/split structure commands
// (AC-7/AC-9). Driven by the `useTables` controller passed as `:controller`;
// runtime prop declarations per the espree constraint.
import { ref } from "vue";
import { asBorderLineStyle, asBorderPreset, asCellVerticalAlign } from "../../lib/core/tables";

const props = defineProps({
  /** The `useTables` controller. */
  controller: { type: Object, required: true },
  /** 1-based slide number, for the panel heading. */
  slideNumber: { type: Number, default: 1 },
});

const c = props.controller;

const fillColor = ref("#ffff00");
const borderPreset = ref("all");
const borderStyle = ref("single");
const borderWidth = ref(1);
const borderColor = ref("#000000");
const paddingCm = ref(0.13);
const valign = ref("top");

function applyFill() {
  c.setFill({ type: "solid", color: fillColor.value });
}
function applyBorder() {
  c.applyBorder(asBorderPreset(borderPreset.value), {
    style: asBorderLineStyle(borderStyle.value),
    widthPt: Number(borderWidth.value) || 1,
    color: borderColor.value,
  });
}
function applyPadding() {
  c.setPadding(Number(paddingCm.value) || 0);
}
function applyValign() {
  c.setValign(asCellVerticalAlign(valign.value));
}
function onHeaderChange(event = new Event("change")) {
  const input = event.target;
  if (input instanceof HTMLInputElement) c.setHeaderRow(input.checked);
}
function onBandChange(event = new Event("change")) {
  const input = event.target;
  if (input instanceof HTMLInputElement) c.setBandRows(input.checked);
}
</script>

<template>
  <section
    class="table-panel"
    aria-label="Table properties"
    data-testid="slides-table-panel"
  >
    <h2 class="panel-title">
      Table · slide {{ slideNumber }}
    </h2>

    <fieldset class="group">
      <legend>Rows &amp; columns</legend>
      <div class="btn-row">
        <button
          type="button"
          data-testid="table-insert-row-above"
          @click="c.insertRowAbove()"
        >
          Row above
        </button>
        <button
          type="button"
          data-testid="table-insert-row-below"
          @click="c.insertRowBelow()"
        >
          Row below
        </button>
        <button
          type="button"
          data-testid="table-insert-column-left"
          @click="c.insertColumnLeft()"
        >
          Col left
        </button>
        <button
          type="button"
          data-testid="table-insert-column-right"
          @click="c.insertColumnRight()"
        >
          Col right
        </button>
      </div>
      <div class="btn-row">
        <button
          type="button"
          data-testid="table-delete-row"
          @click="c.deleteRow()"
        >
          Delete row
        </button>
        <button
          type="button"
          data-testid="table-delete-column"
          @click="c.deleteColumn()"
        >
          Delete column
        </button>
        <button
          type="button"
          data-testid="table-delete-table"
          @click="c.deleteTable()"
        >
          Delete table
        </button>
      </div>
      <div class="btn-row">
        <button
          type="button"
          data-testid="table-merge-cells"
          @click="c.mergeRight()"
        >
          Merge right
        </button>
        <button
          type="button"
          data-testid="table-merge-down"
          @click="c.mergeDown()"
        >
          Merge down
        </button>
        <button
          type="button"
          data-testid="table-split-cell"
          @click="c.splitCurrent()"
        >
          Split cell
        </button>
      </div>
    </fieldset>

    <fieldset class="group">
      <legend>Cell fill</legend>
      <div class="btn-row">
        <input
          v-model="fillColor"
          type="color"
          aria-label="Cell fill colour"
        >
        <button
          type="button"
          data-testid="table-cell-fill"
          @click="applyFill()"
        >
          Fill
        </button>
        <button
          type="button"
          data-testid="table-cell-no-fill"
          @click="c.clearFill()"
        >
          No fill
        </button>
      </div>
    </fieldset>

    <fieldset class="group">
      <legend>Borders</legend>
      <div class="btn-row">
        <select
          v-model="borderPreset"
          data-testid="table-border-preset"
          aria-label="Border preset"
        >
          <option value="all">
            All
          </option>
          <option value="outer">
            Outer
          </option>
          <option value="inner">
            Inner
          </option>
          <option value="none">
            None
          </option>
        </select>
        <select
          v-model="borderStyle"
          data-testid="table-border-style"
          aria-label="Border style"
        >
          <option value="single">
            Single
          </option>
          <option value="double">
            Double
          </option>
          <option value="dashed">
            Dashed
          </option>
          <option value="dotted">
            Dotted
          </option>
          <option value="none">
            None
          </option>
        </select>
        <input
          v-model.number="borderWidth"
          type="number"
          min="0"
          step="0.25"
          data-testid="table-border-width"
          aria-label="Border width (pt)"
        >
        <input
          v-model="borderColor"
          type="color"
          data-testid="table-border-color"
          aria-label="Border colour"
        >
        <button
          type="button"
          @click="applyBorder()"
        >
          Apply
        </button>
      </div>
    </fieldset>

    <fieldset class="group">
      <legend>Cell</legend>
      <label class="field">
        Padding (cm)
        <input
          v-model.number="paddingCm"
          type="number"
          min="0"
          step="0.05"
          data-testid="table-cell-padding"
          @change="applyPadding()"
        >
      </label>
      <label class="field">
        Vertical align
        <select
          v-model="valign"
          data-testid="table-cell-valign"
          @change="applyValign()"
        >
          <option value="top">Top</option>
          <option value="center">Middle</option>
          <option value="bottom">Bottom</option>
        </select>
      </label>
    </fieldset>

    <fieldset class="group">
      <legend>Table style</legend>
      <label class="field check">
        <input
          type="checkbox"
          data-testid="table-header-row"
          :checked="c.table.value?.firstRow === true"
          @change="onHeaderChange"
        >
        Header row
      </label>
      <label class="field check">
        <input
          type="checkbox"
          data-testid="table-band-rows"
          :checked="c.table.value?.bandRow === true"
          @change="onBandChange"
        >
        Banded rows
      </label>
    </fieldset>
  </section>
</template>

<style scoped>
.table-panel {
  padding: var(--space-3);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.panel-title {
  margin: 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.group {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-2);
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.group legend {
  font-size: var(--text-xs, 0.75rem);
  color: var(--text-secondary);
  padding: 0 var(--space-1);
}

.btn-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}

.btn-row button {
  height: 1.9rem;
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-sm);
}

.btn-row button:hover {
  background: var(--surface-sunken);
}

.field {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.field.check {
  cursor: pointer;
}

.field input[type="number"] {
  width: 4.5rem;
}
</style>
