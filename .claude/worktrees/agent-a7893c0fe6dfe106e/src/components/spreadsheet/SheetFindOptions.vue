<script setup lang="ts">
// The sheet's own find options (SPEC-134), rendered in FindBar's `options`
// slot after its regex toggle: "Match entire cell" (`sheet-find-entire-cell`,
// a toggle), Within (`sheet-find-within`: Sheet / Workbook), Look in
// (`sheet-find-look-in`: Values / Formulas) and "Find all"
// (`sheet-find-all`). Each writes useSheetFind's state, which keeps it for
// the view session. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import { asSheetFind } from "../../views/spreadsheet/useSheetFind";
import UiSelect from "../ui/UiSelect.vue";

const props = defineProps({
  /** useSheetFind's return. */
  find: { type: Object, required: true },
});

const f = computed(() => asSheetFind(props.find));

const WITHIN = [
  { value: "sheet", label: "Sheet" },
  { value: "workbook", label: "Workbook" },
];
const LOOK_IN = [
  { value: "values", label: "Values" },
  { value: "formulas", label: "Formulas" },
];

function toggleEntireCell() {
  f.value.entireCell.value = !f.value.entireCell.value;
}

function setWithin(value = "sheet") {
  f.value.within.value = value === "workbook" ? "workbook" : "sheet";
}

function setLookIn(value = "values") {
  f.value.lookIn.value = value === "formulas" ? "formulas" : "values";
}
</script>

<template>
  <button
    type="button"
    class="sheet-find-option"
    data-testid="sheet-find-entire-cell"
    :aria-pressed="f.entireCell.value ? 'true' : 'false'"
    title="Match entire cell contents"
    @click="toggleEntireCell"
  >
    Match entire cell
  </button>
  <label class="sheet-find-field">
    <span>Within</span>
    <UiSelect
      data-testid="sheet-find-within"
      label="Within"
      :model-value="f.within.value"
      :options="WITHIN"
      @update:model-value="setWithin"
    />
  </label>
  <label class="sheet-find-field">
    <span>Look in</span>
    <UiSelect
      data-testid="sheet-find-look-in"
      label="Look in"
      :model-value="f.lookIn.value"
      :options="LOOK_IN"
      @update:model-value="setLookIn"
    />
  </label>
  <button
    type="button"
    class="sheet-find-option"
    data-testid="sheet-find-all"
    title="List every matching cell"
    @click="f.findAll()"
  >
    Find all
  </button>
</template>

<style scoped>
.sheet-find-option {
  display: inline-flex;
  align-items: center;
  min-height: var(--height-control);
  font: inherit;
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.sheet-find-option[aria-pressed="true"] {
  background: var(--surface-selected);
  border-color: var(--accent-subtle);
  color: var(--text-accent);
}

.sheet-find-field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}
</style>
