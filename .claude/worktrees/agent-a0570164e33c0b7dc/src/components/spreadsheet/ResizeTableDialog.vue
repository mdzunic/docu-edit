<script setup lang="ts">
// The design bar's Resize table (SPEC-133): `sheet-resize-table` on
// BaseDialog — the new range (`sheet-resize-table-range`, typed, or pointed
// at through "Select cells" `sheet-resize-table-pick`, collapsing the dialog
// to `sheet-resize-table-picking` while the grid's selection fills it; Done
// `sheet-resize-table-pick-done`), the refusal (`sheet-resize-table-error`:
// the header row must stay in the same row and the new range must overlap
// the table), `sheet-resize-table-ok` and `sheet-resize-table-cancel`. State
// is useSheetTables'. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The table's current range (`$A$1:$D$5`). */
  range: { type: String, required: true },
  /** The inline refusal ("" for none). */
  error: { type: String, default: "" },
  /** Collapsed while the grid's selection is pointed at. */
  picking: { type: Boolean, default: false },
  /** The grid's selection as a range. */
  selectionText: { type: String, default: "" },
});

const emit = defineEmits(["ok", "cancel", "pick", "pick-done"]);

const text = ref(props.range);

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
    aria-label="Select the table's new range"
    data-testid="sheet-resize-table-picking"
  >
    <span>{{ props.selectionText || "Select the table's new range" }}</span>
    <UiButton
      data-testid="sheet-resize-table-pick-done"
      @click="pickDone"
    >
      Done
    </UiButton>
  </div>
  <BaseDialog
    v-else
    title="Resize table"
    testid="sheet-resize-table"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span>Select the new data range for your table</span>
      <span class="row">
        <input
          v-model="text"
          type="text"
          autofocus
          data-testid="sheet-resize-table-range"
          :aria-invalid="props.error ? 'true' : undefined"
          @keydown.enter.prevent="emit('ok', text)"
        >
        <UiButton
          variant="ghost"
          data-testid="sheet-resize-table-pick"
          @click="emit('pick')"
        >
          Select cells
        </UiButton>
      </span>
    </label>
    <p
      v-if="props.error"
      class="error"
      role="alert"
      data-testid="sheet-resize-table-error"
    >
      {{ props.error }}
    </p>
    <template #actions>
      <UiButton
        variant="ghost"
        data-testid="sheet-resize-table-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="sheet-resize-table-ok"
        @click="emit('ok', text)"
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
