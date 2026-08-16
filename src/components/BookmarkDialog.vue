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
  margin: 0.5rem 0 0.25rem;
  font-size: 0.8125rem;
  color: #5f6368;
}

.dialog-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
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

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
