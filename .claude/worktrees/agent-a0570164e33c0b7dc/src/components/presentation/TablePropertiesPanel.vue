<script setup lang="ts">
// The Table properties panel (SPEC-091 AC-10): edits the selected table's
// cells (fill, borders, padding, vertical alignment) and its table-level look
// (header row, banded rows), plus the row/column/merge/split structure commands
// (AC-7/AC-9). Driven by the `useTables` controller passed as `:controller`;
// runtime prop declarations per the espree constraint.
import { computed, ref } from "vue";
import { asBorderLineStyle, asBorderPreset, asCellVerticalAlign } from "../../lib/core/tables";

const props = defineProps({
  /** The `useTables` controller. */
  controller: { type: Object, required: true },
  /** 1-based slide number, for the panel heading. */
  slideNumber: { type: Number, default: 1 },
});

const c = props.controller;

// Every command in this panel routes through useTables.edit(), which returns
// immediately when no table frame is selected. The panel used to be hidden
// entirely in that state (SlidePanelStack gated it on `hasTable`), so opening
// the Table tab with nothing selected lit the tab and showed nothing, and
// clicking off the table made the panel disappear mid-edit. It stays put now
// and says what to do — but the controls go disabled with it, because a button
// that looks live and silently does nothing is the worse half of the bug.
const hasTable = computed(() => c.hasTable.value === true);

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

    <p
      v-if="!hasTable"
      class="panel-empty"
      data-testid="slides-table-empty"
    >
      Select a table on the slide to edit it — or add one from Insert → Table.
    </p>

    <div class="group">
      <span class="group-kicker">Rows &amp; columns</span>
      <div class="btn-row">
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-insert-row-above"
          @click="c.insertRowAbove()"
        >
          Row above
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-insert-row-below"
          @click="c.insertRowBelow()"
        >
          Row below
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-insert-column-left"
          @click="c.insertColumnLeft()"
        >
          Col left
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-insert-column-right"
          @click="c.insertColumnRight()"
        >
          Col right
        </button>
        <span class="row-rule" />
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-delete-row"
          @click="c.deleteRow()"
        >
          Delete row
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-delete-column"
          @click="c.deleteColumn()"
        >
          Delete column
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-delete-table"
          @click="c.deleteTable()"
        >
          Delete table
        </button>
        <span class="row-rule" />
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-merge-cells"
          @click="c.mergeRight()"
        >
          Merge right
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-merge-down"
          @click="c.mergeDown()"
        >
          Merge down
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-split-cell"
          @click="c.splitCurrent()"
        >
          Split cell
        </button>
      </div>
    </div>

    <div class="group">
      <span class="group-kicker">Cell fill</span>
      <div class="btn-row">
        <input
          v-model="fillColor"
          class="swatch"
          type="color"
          :disabled="!hasTable"
          aria-label="Cell fill colour"
        >
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-cell-fill"
          @click="applyFill()"
        >
          Fill
        </button>
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-cell-no-fill"
          @click="c.clearFill()"
        >
          No fill
        </button>
      </div>
    </div>

    <div class="group">
      <span class="group-kicker">Borders</span>
      <div class="btn-row">
        <select
          v-model="borderPreset"
          class="field-box"
          :disabled="!hasTable"
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
          class="field-box"
          :disabled="!hasTable"
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
          class="field-box"
          type="number"
          min="0"
          step="0.25"
          :disabled="!hasTable"
          data-testid="table-border-width"
          aria-label="Border width (pt)"
        >
        <input
          v-model="borderColor"
          class="swatch"
          type="color"
          :disabled="!hasTable"
          data-testid="table-border-color"
          aria-label="Border colour"
        >
        <button
          type="button"
          :disabled="!hasTable"
          data-testid="table-border-apply"
          @click="applyBorder()"
        >
          Apply
        </button>
      </div>
    </div>

    <div class="group">
      <span class="group-kicker">Cell</span>
      <div class="btn-row">
        <label class="field">
          <span class="field-kicker">Padding (cm)</span>
          <input
            v-model.number="paddingCm"
            class="field-box"
            type="number"
            min="0"
            step="0.05"
            :disabled="!hasTable"
            data-testid="table-cell-padding"
            @change="applyPadding()"
          >
        </label>
        <label class="field">
          <span class="field-kicker">Vertical align</span>
          <select
            v-model="valign"
            class="field-box"
            :disabled="!hasTable"
            data-testid="table-cell-valign"
            @change="applyValign()"
          >
            <option value="top">Top</option>
            <option value="center">Middle</option>
            <option value="bottom">Bottom</option>
          </select>
        </label>
      </div>
    </div>

    <div class="group">
      <span class="group-kicker">Table style</span>
      <div class="btn-row">
        <label class="field check">
          <input
            type="checkbox"
            :disabled="!hasTable"
            data-testid="table-header-row"
            :checked="c.table.value?.firstRow === true"
            @change="onHeaderChange"
          >
          Header row
        </label>
        <label class="field check">
          <input
            type="checkbox"
            :disabled="!hasTable"
            data-testid="table-band-rows"
            :checked="c.table.value?.bandRow === true"
            @change="onBandChange"
          >
          Banded rows
        </label>
      </div>
    </div>
  </section>
</template>

<style scoped>
.table-panel {
  padding: var(--space-3) var(--space-5);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

/* The panel's own heading and the group headings share the handoff's kicker:
   condensed, uppercase, tracked, in the accent ink. The <fieldset>/<legend>
   pairs these replaced drew a browser-default box around every group — five
   nested rectangles in a strip whose whole grammar is one hairline. */
.panel-title {
  margin: 0;
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  line-height: var(--line-tight);
  text-transform: uppercase;
  color: var(--text-secondary);
}

.panel-empty {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-md);
}

.group {
  display: flex;
  align-items: baseline;
  gap: var(--space-4);
  padding-block: var(--space-1);
}

.group + .group {
  border-top: var(--border-width) solid var(--border-subtle);
}

.group-kicker {
  flex: none;
  min-width: 7rem;
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  line-height: var(--line-tight);
  text-transform: uppercase;
  color: var(--text-accent);
}

.btn-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}

/* The strip's control grammar: transparent until touched, hairline box on
   hover, square in the palettes that ask for it. */
.btn-row button {
  height: var(--height-toolbar-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-md);
  white-space: nowrap;
}

.btn-row button:hover:not(:disabled) {
  background: var(--state-hover);
  border-color: var(--border-subtle);
}

.btn-row button:disabled,
.field-box:disabled,
.swatch:disabled {
  color: var(--text-disabled);
  cursor: default;
}

/* Separates the three command clusters (insert / delete / merge) that used to
   be three wrapped rows. */
.row-rule {
  width: var(--border-width);
  height: var(--space-5);
  margin: 0 var(--space-2);
  background: var(--border-subtle);
}

.field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.field.check {
  cursor: pointer;
}

.field-kicker {
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  color: var(--text-accent);
}

.field-box {
  height: var(--height-toolbar-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-md);
}

input.field-box[type="number"] {
  width: 4.5rem;
}

.swatch {
  width: var(--height-toolbar-control);
  height: var(--height-toolbar-control);
  padding: var(--space-1);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}
</style>
