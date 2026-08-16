<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";

// Confirmation modal for deleting a document; names the document (AC-14) so
// the user knows exactly what is being removed. Emits confirm/cancel only —
// the caller owns the actual deletion. BaseDialog owns the shell; through it
// this dialog also gains Escape-to-cancel, which it never had.
const props = defineProps({
  /** Title of the document (or template — see `kind`) about to be deleted;
   * shown in the message. */
  documentTitle: { type: String, required: true },
  /** What's being deleted, for the heading/message noun (SPEC-046 reuses
   * this dialog for user templates). Defaults to the original behaviour. */
  kind: { type: String, default: "document" },
});

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    :title="`Delete ${props.kind}`"
    testid="delete-dialog"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="delete-message"
    >
      Delete "{{ documentTitle }}"? This cannot be undone.
    </p>
    <template #actions>
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
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-message {
  margin: 0;
  color: #3c4043;
  font-size: 0.9375rem;
  overflow-wrap: anywhere;
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
