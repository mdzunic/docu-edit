<script setup lang="ts">
// Insert ▸ Table's dialog (SPEC-133 "Creating a table"): `sheet-create-table`
// on BaseDialog — the range (`sheet-create-table-range`, typed, or pointed at
// through "Select cells" `sheet-create-table-pick`, which collapses the dialog
// to `sheet-create-table-picking` while the grid's selection fills it; Done
// `sheet-create-table-pick-done` — SPEC-132's Select Data pattern), "My table
// has headers" (`sheet-create-table-headers`), the refusal
// (`sheet-create-table-error`), `sheet-create-table-ok` and
// `sheet-create-table-cancel`. State is useSheetTables'. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The prefilled range (`$A$1:$D$5`). */
  range: { type: String, required: true },
  /** The prefilled "My table has headers" (detectHeaderRow). */
  headers: { type: Boolean, default: true },
  /** The inline refusal ("" for none). */
  error: { type: String, default: "" },
  /** Collapsed while the grid's selection is pointed at. */
  picking: { type: Boolean, default: false },
  /** The grid's selection as a range. */
  selectionText: { type: String, default: "" },
});

const emit = defineEmits(["ok", "cancel", "pick", "pick-done"]);

const text = ref(props.range);
const hasHeaders = ref(props.headers);

function pickDone() {
  if (props.selectionText) {
    text.value = props.selectionText;
  }
  emit("pick-done");
}
</script>

<template>
  <div
    v-if="props.picking"
    class="table-picking"
    role="region"
    aria-label="Select the table's range"
    data-testid="sheet-create-table-picking"
  >
    <span>{{ props.selectionText || "Select the table's cells" }}</span>
    <UiButton
      data-testid="sheet-create-table-pick-done"
      @click="pickDone"
    >
      Done
    </UiButton>
  </div>
  <BaseDialog
    v-else
    title="Create table"
    testid="sheet-create-table"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span>Where is the data for your table?</span>
      <span class="row">
        <input
          v-model="text"
          type="text"
          autofocus
          data-testid="sheet-create-table-range"
          :aria-invalid="props.error ? 'true' : undefined"
          @keydown.enter.prevent="emit('ok', text, hasHeaders)"
        >
        <UiButton
          variant="ghost"
          data-testid="sheet-create-table-pick"
          @click="emit('pick')"
        >
          Select cells
        </UiButton>
      </span>
    </label>
    <label class="check"><input
      v-model="hasHeaders"
      type="checkbox"
      data-testid="sheet-create-table-headers"
    > My table has headers</label>
    <p
      v-if="props.error"
      class="error"
      role="alert"
      data-testid="sheet-create-table-error"
    >
      {{ props.error }}
    </p>
    <template #actions>
      <UiButton
        variant="ghost"
        data-testid="sheet-create-table-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="sheet-create-table-ok"
        @click="emit('ok', text, hasHeaders)"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  margin: 0 0 var(--space-3);
}

.row {
  display: flex;
  gap: var(--space-2);
}

.check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.error {
  color: var(--danger);
  font-size: var(--text-sm);
}

.table-picking {
  position: fixed;
  top: var(--space-4);
  left: 50%;
  z-index: 50;
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  transform: translateX(-50%);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-3);
}
</style>
