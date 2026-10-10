<script setup lang="ts">
// SPEC-124's insert/delete dialog: Ctrl+Alt+= / Ctrl+Alt+- (Cmd+Option on
// macOS) or the cell menu's Insert…/Delete… on a selection that is not whole
// rows or columns offers "Entire row" and "Entire column"
// (`sheet-insert-delete-entire-row|entire-column`) — inserting above/left or
// deleting the active range's rows or columns. Emits choose(axis) / cancel;
// useStructure runs the command. BaseDialog's dismissal is a cancel.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** "insert" | "delete". */
  mode: { type: String, default: "insert" },
});

const emit = defineEmits(["choose", "cancel"]);
</script>

<template>
  <BaseDialog
    :title="props.mode === 'delete' ? 'Delete' : 'Insert'"
    :testid="props.mode === 'delete' ? 'sheet-delete-dialog' : 'sheet-insert-dialog'"
    width="20rem"
    @dismiss="emit('cancel')"
  >
    <div
      class="choices"
      role="group"
      :aria-label="props.mode === 'delete' ? 'Delete' : 'Insert'"
    >
      <UiButton
        variant="secondary"
        data-testid="sheet-insert-delete-entire-row"
        autofocus
        @click="emit('choose', 'row')"
      >
        Entire row
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="sheet-insert-delete-entire-column"
        @click="emit('choose', 'col')"
      >
        Entire column
      </UiButton>
    </div>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-insert-delete-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.choices {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}
</style>
