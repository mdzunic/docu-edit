<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { validateBookmarkName } from "../lib/links";

// The "Insert bookmark" dialog (SPEC-025 AC-22, AC-26..28): names the caret
// location. Validation (no spaces, unique against the document's existing
// names) lives in lib/links.ts; a conflict shows the inline error (AC-28).
// Runtime prop declarations — see AppShell.vue's espree comment.
const props = defineProps({
  /** The document's current bookmark names, for uniqueness validation. */
  existingNames: { type: Array, default: () => [] },
});

const emit = defineEmits(["save", "cancel"]);

const draft = ref("");
const input = useTemplateRef("input");

onMounted(async () => {
  await nextTick();
  input.value?.focus();
});

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
  <div
    class="dialog-backdrop"
    data-testid="bookmark-dialog"
    @click.self="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Insert bookmark"
      @keydown.escape.stop="emit('cancel')"
    >
      <h2 class="dialog-title">
        Insert bookmark
      </h2>
      <label
        class="dialog-label"
        for="bookmark-name-input"
      >Bookmark name</label>
      <input
        id="bookmark-name-input"
        ref="input"
        v-model="draft"
        class="dialog-input"
        data-testid="bookmark-name-input"
        @keydown.enter="save"
      >
      <p
        v-if="error"
        class="dialog-error"
        data-testid="bookmark-error"
      >
        {{ error }}
      </p>
      <div class="dialog-actions">
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
