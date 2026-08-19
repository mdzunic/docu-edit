<script setup lang="ts">
// The "delete its caption too?" prompt (SPEC-055 AC-20, AC-21): raised when a
// captioned object is deleted. Defaults to deleting both (the primary button),
// but keeping the caption or cancelling are offered. The parent owns the actual
// deletion. Runtime prop/emit declarations per the espree constraint.
import BaseDialog from "./BaseDialog.vue";

defineProps({
  /** The object being deleted, e.g. "image" / "table" / "chart". */
  objectKind: { type: String, default: "object" },
});

const emit = defineEmits(["delete-both", "keep-caption", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Delete caption?"
    testid="delete-caption-dialog"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="delete-caption-message"
    >
      This {{ objectKind }} has a caption. Delete the caption too?
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="delete-caption-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        data-testid="delete-caption-keep"
        @click="emit('keep-caption')"
      >
        Keep caption
      </button>
      <button
        type="button"
        class="primary"
        data-testid="delete-caption-both"
        @click="emit('delete-both')"
      >
        Delete caption too
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-message {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-lg);
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
</style>
