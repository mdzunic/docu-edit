<script setup lang="ts">
// SPEC-124 "Merges": shown before a merge (or the format painter's merges)
// would clear a value or formula outside the kept upper-left cell. Merge
// (`sheet-merge-confirm`) goes ahead — the other cells' values and formulas
// are cleared, their styles kept; Cancel (`sheet-merge-cancel`) holds the
// initial focus, and BaseDialog's dismissal is a cancel. useMerges owns the
// command. No props.
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Merge cells"
    testid="sheet-merge-warning"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="sheet-merge-warning-message"
    >
      Merging keeps only the upper-left value
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-merge-cancel"
        autofocus
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-merge-confirm"
        @click="emit('confirm')"
      >
        Merge
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.dialog-message {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-md);
}
</style>
