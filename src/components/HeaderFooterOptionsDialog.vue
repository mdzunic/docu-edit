<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// "Header/footer options" (SPEC-029 AC-11, AC-12, AC-22, AC-23): "Different
// first page" and "Different odd and even pages" toggles, plus "Start
// numbering at" (a blank value means "continue from the previous
// section"). Applies to the section passed in `initial.sectionIndex` — the
// section currently being edited, or the caret's own section otherwise.
// BaseDialog owns the shell, Escape and initial focus (the [autofocus]
// first checkbox).
const props = defineProps({
  /** Prefill: { sectionIndex, titlePage, evenAndOddHeaders,
   * pageNumberStart } — pageNumberStart is null for "continue". */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const titlePage = ref(props.initial.titlePage === true);
const evenAndOddHeaders = ref(props.initial.evenAndOddHeaders === true);
const restartNumbering = ref(props.initial.pageNumberStart !== null && props.initial.pageNumberStart !== undefined);
const startAt = ref(String(props.initial.pageNumberStart ?? 1));

function apply() {
  const parsed = Number(startAt.value);
  emit("apply", {
    sectionIndex: props.initial.sectionIndex,
    titlePage: titlePage.value,
    evenAndOddHeaders: evenAndOddHeaders.value,
    pageNumberStart: restartNumbering.value && Number.isFinite(parsed) ? Math.max(1, Math.floor(parsed)) : null,
  });
}
</script>

<template>
  <BaseDialog
    title="Header/footer options"
    testid="header-footer-options-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <label class="checkbox-field">
      <input
        v-model="titlePage"
        type="checkbox"
        data-testid="hf-option-title-page"
        autofocus
      >
      <span>Different first page</span>
    </label>
    <label class="checkbox-field">
      <input
        v-model="evenAndOddHeaders"
        type="checkbox"
        data-testid="hf-option-even-odd"
      >
      <span>Different odd and even pages</span>
    </label>
    <label class="checkbox-field">
      <input
        v-model="restartNumbering"
        type="checkbox"
        data-testid="hf-option-restart"
      >
      <span>Start numbering at</span>
    </label>
    <input
      v-model="startAt"
      class="field-input start-input"
      data-testid="hf-option-start-at"
      type="number"
      min="1"
      step="1"
      :disabled="!restartNumbering"
    >
    <p class="hint">
      Leave unchecked to continue numbering from the previous section.
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="hf-options-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="hf-options-apply"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.checkbox-field {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-md);
  color: var(--text-primary);
  margin-bottom: var(--space-3);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.start-input {
  width: 6rem;
  margin-left: var(--space-7);
}

.hint {
  margin: var(--space-3) 0 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
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
</style>
