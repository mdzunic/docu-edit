<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { MAX_TITLE_LENGTH } from "../lib/documentStore";

// Modal for renaming a library document. Owns the input-level validation
// (non-empty, max length — the maxlength attribute already blocks typing
// past the limit); deduplication against other titles happens in the store
// when the parent applies the rename.
const props = defineProps({
  /** The document's current title; prefills the input. */
  currentTitle: { type: String, required: true },
});

const emit = defineEmits(["save", "cancel"]);

const draft = ref(props.currentTitle);
const input = useTemplateRef("input");

onMounted(async () => {
  await nextTick();
  input.value?.focus();
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
  <div
    class="dialog-backdrop"
    data-testid="rename-dialog"
    @click.self="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Rename document"
    >
      <h2 class="dialog-title">
        Rename document
      </h2>
      <input
        ref="input"
        v-model="draft"
        class="dialog-input"
        data-testid="rename-input"
        aria-label="Document title"
        :maxlength="MAX_TITLE_LENGTH"
        @keydown.enter="save"
        @keydown.escape="emit('cancel')"
      >
      <p
        v-if="error"
        class="dialog-error"
        data-testid="rename-error"
      >
        {{ error }}
      </p>
      <div class="dialog-actions">
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
          Rename
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
  min-width: 20rem;
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
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

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
