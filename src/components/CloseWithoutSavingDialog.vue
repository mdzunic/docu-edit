<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

// The one exit that can lose work, made explicit (SPEC-060): when storage is
// unavailable (SPEC-003) the open document lives only in memory, so closing
// discards it. This dialog is the confirmation that precedes that loss. It
// emits its three choices and owns none of the behaviour — Download runs the
// existing download flow and deliberately leaves the dialog open, so a user
// can save the file first and then close on purpose. BaseDialog supplies the
// shell (and, through it, Escape-to-cancel).
const emit = defineEmits(["download", "confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Close without saving?"
    testid="close-without-saving-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="close-without-saving-message"
    >
      Storage is unavailable in this browser, so this document has not been
      saved. Closing it will lose your changes.
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="close-without-saving-download"
        @click="emit('download')"
      >
        Download
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="close-without-saving-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="danger"
        data-testid="close-without-saving-confirm"
        @click="emit('confirm')"
      >
        Close anyway
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
