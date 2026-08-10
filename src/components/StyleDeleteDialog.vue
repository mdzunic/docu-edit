<script setup lang="ts">
// "Delete style" confirmation (SPEC-042 AC-7): asks before proceeding, and
// when the style is in use (AC-2), warns that its paragraphs/runs/tables
// will revert to the parent style (or Normal/none) first (AC-8). Runtime
// prop/emit declarations — see AppShell.vue's espree comment.
defineProps({
  styleName: { type: String, required: true },
  inUse: { type: Boolean, default: false },
});

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="style-delete-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Delete style"
    >
      <h2 class="dialog-title">
        Delete style
      </h2>
      <p
        class="dialog-message"
        data-testid="style-delete-message"
      >
        Delete "{{ styleName }}"? This cannot be undone.
      </p>
      <p
        v-if="inUse"
        class="dialog-warning"
        data-testid="style-delete-in-use-warning"
      >
        This style is currently in use. Paragraphs, text or tables using it
        will revert to its parent style, or to the default if it has none.
      </p>
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="style-delete-cancel"
          @click="emit('cancel')"
        >
          Cancel
        </button>
        <button
          type="button"
          class="danger"
          data-testid="style-delete-confirm"
          @click="emit('confirm')"
        >
          Delete
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

.dialog-message {
  margin: 0;
  color: #3c4043;
  font-size: 0.9375rem;
  overflow-wrap: anywhere;
}

.dialog-warning {
  margin: 0.75rem 0 0;
  color: #b06000;
  background: #fef7e0;
  border-radius: 4px;
  padding: 0.5rem 0.75rem;
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

.dialog-actions button.danger {
  background: #b3261e;
  border-color: #b3261e;
  color: #fff;
}
</style>
