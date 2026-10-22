<script setup lang="ts">
// SPEC-121 "Deleting a sheet": the confirmation shown before deleting a sheet
// that has content. Emits confirm/cancel only — the view's useSheetCommands
// owns the deletion (which is one undo step). Cancel holds the initial focus,
// so an Enter that opened the dialog never deletes; BaseDialog's dismissal
// (Escape, backdrop) is a cancel. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The sheet's name, quoted in the message. */
  name: { type: String, default: "" },
});

const emit = defineEmits(["confirm", "cancel"]);
</script>

<template>
  <BaseDialog
    title="Delete sheet"
    testid="sheet-tab-delete-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <p
      class="dialog-message"
      data-testid="sheet-tab-delete-message"
    >
      Delete "{{ props.name }}"? Its cells and anything on it will be removed.
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-tab-delete-cancel"
        autofocus
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="danger"
        data-testid="sheet-tab-delete-confirm"
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
  font-size: var(--text-md);
  overflow-wrap: anywhere;
}
</style>
