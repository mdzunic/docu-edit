<script setup lang="ts">
import { nextTick, onMounted, ref, useTemplateRef } from "vue";

// "Header/footer options" (SPEC-029 AC-11, AC-12, AC-22, AC-23): "Different
// first page" and "Different odd and even pages" toggles, plus "Start
// numbering at" (a blank value means "continue from the previous
// section"). Applies to the section passed in `initial.sectionIndex` — the
// section currently being edited, or the caret's own section otherwise.
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

const firstCheckbox = useTemplateRef("firstCheckbox");

onMounted(async () => {
  await nextTick();
  firstCheckbox.value?.focus();
});

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
  <div
    class="dialog-backdrop"
    data-testid="header-footer-options-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Header/footer options"
    >
      <h2 class="dialog-title">
        Header/footer options
      </h2>
      <label class="checkbox-field">
        <input
          ref="firstCheckbox"
          v-model="titlePage"
          type="checkbox"
          data-testid="hf-option-title-page"
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
      <div class="dialog-actions">
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
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  min-width: 22rem;
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.8125rem;
  color: #3c4043;
  margin-bottom: 0.5rem;
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
}

.start-input {
  width: 6rem;
  margin-left: 1.75rem;
}

.hint {
  margin: 0.5rem 0 0;
  font-size: 0.75rem;
  color: #5f6368;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1.25rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}
</style>
