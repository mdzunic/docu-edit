<script setup lang="ts">
// SPEC-124's size dialogs: "Row height…" (`sheet-row-height-dialog`,
// `sheet-row-height-input`) edits points, 0–409, rounded to the nearest
// 0.75 pt (0 hides); "Column width…" (`sheet-column-width-dialog`,
// `sheet-column-width-input`) edits Excel characters, 0–255 with two
// decimals, converted through SPEC-117's MDW formula (0 hides). Each shows
// the pixel equivalent ("= 64 pixels") as the value is typed and opens blank
// when the selected rows or columns differ. Emits confirm(amount) / cancel;
// useStructure runs the command. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { MAX_COLUMN_WIDTH_CHARS, MAX_ROW_HEIGHT_PT } from "../../lib/spreadsheet/bands";
import { sizePixels } from "../../views/spreadsheet/sheetSizes";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** "row" | "col". */
  axis: { type: String, default: "row" },
  /** The value the dialog opens with ("" when the sizes differ). */
  initial: { type: String, default: "" },
  /** The workbook's maximum digit width, for columns. */
  mdw: { type: Number, default: 7 },
});

const emit = defineEmits(["confirm", "cancel"]);

const rows = computed(() => props.axis === "row");
const text = ref(props.initial);
const amount = computed(() => (text.value.trim() === "" ? Number.NaN : Number(text.value.trim().replace(",", "."))));
const valid = computed(() => Number.isFinite(amount.value) && amount.value >= 0 && amount.value <= (rows.value ? MAX_ROW_HEIGHT_PT : MAX_COLUMN_WIDTH_CHARS));
const pixels = computed(() => (valid.value ? sizePixels(rows.value ? "row" : "col", amount.value, props.mdw) : null));

function confirm() {
  if (valid.value) {
    emit("confirm", rows.value ? amount.value : Math.round(amount.value * 100) / 100);
  }
}
</script>

<template>
  <BaseDialog
    :title="rows ? 'Row height' : 'Column width'"
    :testid="rows ? 'sheet-row-height-dialog' : 'sheet-column-width-dialog'"
    width="20rem"
    @dismiss="emit('cancel')"
  >
    <form
      class="size-form"
      @submit.prevent="confirm"
    >
      <label class="size-field">
        <span>{{ rows ? "Row height (points)" : "Column width (characters)" }}</span>
        <input
          v-model="text"
          class="size-input"
          inputmode="decimal"
          autofocus
          :aria-invalid="valid || text === '' ? undefined : 'true'"
          :data-testid="rows ? 'sheet-row-height-input' : 'sheet-column-width-input'"
        >
      </label>
      <p
        class="size-pixels"
        data-testid="sheet-size-pixels"
        aria-live="polite"
      >
        {{ pixels === null ? (text === "" ? "" : `Enter 0 to ${rows ? MAX_ROW_HEIGHT_PT : MAX_COLUMN_WIDTH_CHARS}`) : `= ${pixels} pixels` }}
      </p>
    </form>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-size-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-size-ok"
        :disabled="!valid"
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.size-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.size-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-md);
}

.size-input {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.size-input[aria-invalid="true"] {
  border-color: var(--danger);
}

.size-pixels {
  min-height: 1.25em;
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-variant-numeric: tabular-nums;
}
</style>
