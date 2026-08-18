<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";

// "Delete style" confirmation (SPEC-042 AC-7): asks before proceeding, and
// when the style is in use (AC-2), warns that its paragraphs/runs/tables
// will revert to the parent style (or Normal/none) first (AC-8). BaseDialog
// owns the shell and Escape — no control is marked [autofocus] (this dialog
// never focused one), so the panel itself takes focus and Escape works
// immediately. Runtime prop/emit declarations — see AppShell.vue's espree
// comment.
defineProps({
  styleName: { type: String, required: true },
  inUse: { type: Boolean, default: false },
});

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Delete style"
    testid="style-delete-dialog"
    @dismiss="emit('cancel')"
  >
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
    <template #actions>
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

.dialog-warning {
  margin: var(--space-4) 0 0;
  color: var(--warning);
  background: var(--surface-sunken);
  border-radius: var(--radius-sm);
  padding: var(--space-3) var(--space-4);
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

.dialog-actions button.danger {
  background: var(--danger);
  border-color: var(--danger);
  color: var(--text-on-brand);
}
</style>
