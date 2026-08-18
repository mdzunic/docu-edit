<script setup lang="ts">
import { computed, ref } from "vue";
import { validateBookmarkName } from "../lib/links";
import BaseDialog from "./BaseDialog.vue";

// The "Insert bookmark" dialog (SPEC-025 AC-22, AC-26..28): names the caret
// location. Validation (no spaces, unique against the document's existing
// names) lives in lib/links.ts; a conflict shows the inline error (AC-28).
// BaseDialog owns the shell, Escape and initial focus (the [autofocus] name
// input). Runtime prop declarations — see AppShell.vue's espree comment.
const props = defineProps({
  /** The document's current bookmark names, for uniqueness validation. */
  existingNames: { type: Array, default: () => [] },
});

const emit = defineEmits(["save", "cancel"]);

const draft = ref("");

const error = computed(() =>
  draft.value === "" ? "" : validateBookmarkName(draft.value, props.existingNames.map(String)),
);

const canSave = computed(() => draft.value !== "" && error.value === "");

function save() {
  if (!canSave.value) {
    return;
  }
  emit("save", draft.value);
}
</script>

<template>
  <BaseDialog
    title="Insert bookmark"
    testid="bookmark-dialog"
    @dismiss="emit('cancel')"
  >
    <label
      class="dialog-label"
      for="bookmark-name-input"
    >Bookmark name</label>
    <input
      id="bookmark-name-input"
      v-model="draft"
      class="dialog-input"
      data-testid="bookmark-name-input"
      autofocus
      @keydown.enter="save"
    >
    <p
      v-if="error"
      class="dialog-error"
      data-testid="bookmark-error"
    >
      {{ error }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="bookmark-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="bookmark-save"
        :disabled="!canSave"
        @click="save"
      >
        Add
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-label {
  display: block;
  margin: var(--space-3) 0 var(--space-2);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.dialog-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
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

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
