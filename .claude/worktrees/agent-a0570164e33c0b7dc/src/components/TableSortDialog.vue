<script setup lang="ts">
import { computed, ref } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import { asSortKeys, noSortKeys } from "../lib/tableSort";
import { asTable } from "../lib/tables";

// The Sort dialog (SPEC-076 AC-38..AC-41, AC-49, AC-50): up to three keys,
// each a column + direction + data type, over a range the "My table has a
// header row" toggle decides. The sort itself is pure (lib/tableSort.ts);
// this only collects the keys and refuses the cases the caller reports as
// blocked. Runtime prop declarations per the .vue espree constraint.
const props = defineProps({
  /** The table being sorted (documentModel.Table) — its column count limits
   * every key's picker (AC-38). */
  table: { type: Object, required: true },
  /** The dialog's seed: key one defaults to the caret's column (AC-39). */
  initialKeys: { type: Array, default: () => noSortKeys() },
  /** Whether the table's first row is a header row (AC-40). */
  initialHasHeader: { type: Boolean, default: false },
  /** Why OK is disabled for the CURRENT header toggle, or "" — the caller
   * recomputes it as the toggle moves the range (AC-49, AC-50). */
  blockedReason: { type: String, default: "" },
});

const emit = defineEmits(["apply", "cancel", "header-change"]);

const TYPES = [
  { id: "text", label: "Text" },
  { id: "number", label: "Number" },
  { id: "date", label: "Date" },
];

const hasHeader = ref(props.initialHasHeader);

/** Three key slots. Slot 0 is seeded from the caller; slots 1 and 2 start at
 * "(none)", which is columnIndex -1 (AC-39). */
const slots = ref(
  [0, 1, 2].map((index) => {
    const seed = asSortKeys(props.initialKeys)[index];
    return {
      columnIndex: String(seed ? seed.columnIndex : -1),
      direction: seed ? seed.direction : "asc",
      type: seed ? seed.type : "text",
    };
  }),
);

const columns = computed(() =>
  asTable(props.table).columnWidthsCm.map((_unused, index) => ({
    index,
    // 1-based in the UI: "Column 1" is what the user sees in the table.
    label: `Column ${index + 1}`,
  })),
);

const keys = computed(() =>
  asSortKeys(
    slots.value
      .filter((slot) => Number(slot.columnIndex) >= 0)
      .map((slot) => ({
        columnIndex: Number(slot.columnIndex),
        direction: slot.direction,
        type: slot.type,
      })),
  ),
);

const reason = computed(() => props.blockedReason || (keys.value.length === 0 ? "choose a column to sort by" : ""));

function onHeaderChange() {
  emit("header-change", hasHeader.value);
}

function apply() {
  if (reason.value) {
    return;
  }
  emit("apply", keys.value, hasHeader.value);
}
</script>

<template>
  <BaseDialog
    title="Sort"
    testid="table-sort-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <div
      v-for="(slot, index) in slots"
      :key="index"
      class="sort-key"
    >
      <span class="sort-key-label">{{ index === 0 ? "Sort by" : "Then by" }}</span>
      <select
        v-model="slot.columnIndex"
        class="field-input"
        :data-testid="`sort-column-${index}`"
        :aria-label="index === 0 ? 'Sort by column' : `Then by column ${index}`"
      >
        <option value="-1">
          (none)
        </option>
        <option
          v-for="column in columns"
          :key="column.index"
          :value="String(column.index)"
        >
          {{ column.label }}
        </option>
      </select>
      <select
        v-model="slot.type"
        class="field-input"
        :data-testid="`sort-type-${index}`"
        aria-label="Sort data type"
      >
        <option
          v-for="type in TYPES"
          :key="type.id"
          :value="type.id"
        >
          {{ type.label }}
        </option>
      </select>
      <select
        v-model="slot.direction"
        class="field-input"
        :data-testid="`sort-direction-${index}`"
        aria-label="Sort direction"
      >
        <option value="asc">
          Ascending
        </option>
        <option value="desc">
          Descending
        </option>
      </select>
    </div>
    <label class="field-check">
      <input
        v-model="hasHeader"
        type="checkbox"
        data-testid="sort-has-header"
        @change="onHeaderChange"
      >
      <span>My table has a header row</span>
    </label>
    <p
      v-if="reason"
      class="dialog-error"
      data-testid="sort-blocked-reason"
    >
      {{ reason }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sort-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sort-confirm"
        :disabled="!!reason"
        @click="apply"
      >
        Sort
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.sort-key {
  display: grid;
  grid-template-columns: 4.5rem 1fr 1fr 1fr;
  align-items: center;
  gap: var(--space-2);
  margin-bottom: var(--space-3);
}

.sort-key-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.field-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  font-size: var(--text-md);
  color: var(--danger);
}
</style>
