<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { MAX_TITLE_LENGTH } from "../lib/documentStore";
import BaseDialog from "./BaseDialog.vue";

// Modal for renaming a library document. Owns the input-level validation
// (non-empty, max length — the maxlength attribute already blocks typing
// past the limit); deduplication against other titles happens in the store
// when the parent applies the rename. BaseDialog owns the shell and Escape —
// which also fixes the old bug where Escape was bound to the input only and
// did nothing while a button had focus.
const props = defineProps({
  /** The document's current title; prefills the input. */
  currentTitle: { type: String, required: true },
  /** Dialog heading and accessible name (SPEC-046 reuses this dialog for
   * naming a new template — "Save as template" — where "Rename document"
   * would be misleading). Defaults to the original behaviour. */
  heading: { type: String, default: "Rename document" },
  /** The submit button's label. */
  confirmLabel: { type: String, default: "Rename" },
});

const emit = defineEmits(["save", "cancel"]);

const draft = ref(props.currentTitle);
const input = useTemplateRef("input");

// BaseDialog focuses the [autofocus] input; selecting the prefilled title on
// top of that (type-to-replace) is this dialog's own one-shot behaviour.
onMounted(async () => {
  await nextTick();
  input.value?.select();
});

const trimmed = computed(() => draft.value.trim());
const error = computed(() => {
  if (trimmed.value.length === 0) {
    return "Title cannot be empty";
  }
  if (trimmed.value.length > MAX_TITLE_LENGTH) {
    return `Title must be at most ${MAX_TITLE_LENGTH} characters`;
  }
  return "";
});

function save() {
  if (error.value) {
    return;
  }
  emit("save", trimmed.value);
}
</script>

<template>
  <BaseDialog
    :title="props.heading"
    testid="rename-dialog"
    @dismiss="emit('cancel')"
  >
    <input
      ref="input"
      v-model="draft"
      class="dialog-input"
      data-testid="rename-input"
      aria-label="Document title"
      autofocus
      :maxlength="MAX_TITLE_LENGTH"
      @keydown.enter="save"
    >
    <p
      v-if="error"
      class="dialog-error"
      data-testid="rename-error"
    >
      {{ error }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="rename-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="rename-save"
        :disabled="!!error"
        @click="save"
      >
        {{ props.confirmLabel }}
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
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
