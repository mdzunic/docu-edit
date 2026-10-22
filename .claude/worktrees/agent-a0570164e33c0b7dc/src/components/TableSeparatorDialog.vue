<script setup lang="ts">
import { computed, ref } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import {
  asConvertSeparatorId,
  CONVERT_SEPARATORS,
  isValidCustomSeparator,
  separatorCharacter,
  textToTableSize,
} from "../lib/tableConvert";
import { asParagraphBlocks } from "../lib/documentModel";

// Both conversions' dialog (SPEC-076 AC-53..AC-55, AC-61): the same
// separator set in each direction — Tab (default), Comma, Semicolon and a
// single-character Custom. In the text→table direction it also previews the
// resulting row and column count, live as the separator changes (AC-54).
// Runtime prop declarations per the .vue espree constraint.
const props = defineProps({
  /** "text-to-table" or "table-to-text" — which title, verb and preview. */
  direction: { type: String, default: "text-to-table" },
  /** The paragraphs a text→table conversion would consume, for the live
   * row/column preview. Ignored in the other direction. */
  paragraphs: { type: Array, default: () => [] },
});

const emit = defineEmits(["apply", "cancel"]);

const separator = ref("tab");
const custom = ref("");

const toTable = computed(() => props.direction === "text-to-table");

const customError = computed(() =>
  separator.value === "custom" && !isValidCustomSeparator(custom.value)
    ? "Enter a single, non-space character"
    : "",
);

const preview = computed(() => {
  if (!toTable.value || customError.value) {
    return null;
  }
  return textToTableSize(
    asParagraphBlocks(props.paragraphs),
    separatorCharacter(asConvertSeparatorId(separator.value), custom.value),
  );
});

function apply() {
  if (customError.value) {
    return;
  }
  emit("apply", separator.value, custom.value);
}
</script>

<template>
  <BaseDialog
    :title="toTable ? 'Convert text to table' : 'Convert table to text'"
    testid="table-separator-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <fieldset class="separators">
      <legend>Separate text at</legend>
      <label
        v-for="option in CONVERT_SEPARATORS"
        :key="option.id"
        class="separator-option"
      >
        <input
          v-model="separator"
          type="radio"
          name="table-separator"
          :value="option.id"
          :data-testid="`separator-${option.id}`"
          :autofocus="option.id === 'tab'"
        >
        <span>{{ option.label }}</span>
      </label>
      <input
        v-model="custom"
        class="field-input custom-input"
        data-testid="separator-custom-char"
        aria-label="Custom separator character"
        :disabled="separator !== 'custom'"
        maxlength="1"
      >
    </fieldset>
    <p
      v-if="preview"
      class="preview"
      data-testid="convert-preview"
    >
      {{ preview.rows }} rows × {{ preview.columns }} columns
    </p>
    <p
      v-if="customError"
      class="dialog-error"
      data-testid="separator-error"
    >
      {{ customError }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="convert-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="convert-confirm"
        :disabled="!!customError"
        @click="apply"
      >
        Convert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.separators {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  margin: 0;
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.separator-option {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.custom-input {
  width: 4rem;
  margin-left: var(--space-5);
}

.preview {
  margin: var(--space-3) 0 0;
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  font-size: var(--text-md);
  color: var(--danger);
}
</style>
