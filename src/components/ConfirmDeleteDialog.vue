<script setup lang="ts">
// Confirmation modal for deleting a document; names the document (AC-14) so
// the user knows exactly what is being removed. Emits confirm/cancel only —
// the caller owns the actual deletion.
defineProps({
  /** Title of the document about to be deleted; shown in the message. */
  documentTitle: { type: String, required: true },
});

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="delete-dialog"
    @click.self="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Delete document"
    >
      <h2 class="dialog-title">
        Delete document
      </h2>
      <p
        class="dialog-message"
        data-testid="delete-message"
      >
        Delete "{{ documentTitle }}"? This cannot be undone.
      </p>
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="delete-cancel"
          @click="emit('cancel')"
        >
          Cancel
        </button>
        <button
          type="button"
          class="danger"
          data-testid="delete-confirm"
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
