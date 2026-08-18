<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

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
      <UiButton
        variant="secondary"
        data-testid="delete-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="danger"
        data-testid="delete-confirm"
        @click="emit('confirm')"
      >
        Delete
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-message {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-lg);
  overflow-wrap: anywhere;
}
</style>
