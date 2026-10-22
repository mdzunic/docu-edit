<script setup lang="ts">
// SPEC-125 drag-move: asked before a drop onto non-empty cells outside the
// dragged block (`sheet-move-overwrite`): "There's already data here. Do you
// want to replace it?" Replace (`sheet-move-overwrite-replace`) goes ahead;
// Cancel (`sheet-move-overwrite-cancel`) holds the initial focus, and
// BaseDialog's dismissal is a cancel. useRangeMove owns the command. No props.
import { MOVE_OVERWRITE_MESSAGE } from "../../lib/spreadsheet/clipboard";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Move cells"
    testid="sheet-move-overwrite"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="sheet-move-overwrite-message"
    >
      {{ MOVE_OVERWRITE_MESSAGE }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-move-overwrite-cancel"
        autofocus
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-move-overwrite-replace"
        @click="emit('confirm')"
      >
        Replace
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
