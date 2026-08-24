<script setup lang="ts">
// Bulk action bar (SPEC-067 AC-27..32). Pure-lib actions (download,
// duplicate, star, clear) run directly on the controller; "Move to trash"
// emits so HomeView can raise the Undo toast it owns.
import UiButton from "../../components/ui/UiButton.vue";

const props = defineProps({
  lib: { type: Object, required: true },
});

const emit = defineEmits(["trash"]);
</script>

<template>
  <div
    class="selection-bar"
    data-testid="selection-bar"
    role="toolbar"
    aria-label="Bulk actions"
  >
    <span
      class="selection-count"
      data-testid="selection-count"
    >
      {{ props.lib.selectionCount.value }} selected
    </span>
    <UiButton
      size="sm"
      data-testid="bulk-download"
      @click="props.lib.downloadDocuments(props.lib.selectedIds.value)"
    >
      Download
    </UiButton>
    <UiButton
      size="sm"
      data-testid="bulk-duplicate"
      @click="props.lib.duplicateDocuments(props.lib.selectedIds.value)"
    >
      Duplicate
    </UiButton>
    <UiButton
      size="sm"
      data-testid="bulk-star"
      @click="props.lib.starDocuments(props.lib.selectedIds.value)"
    >
      Star
    </UiButton>
    <UiButton
      size="sm"
      variant="danger"
      data-testid="bulk-trash"
      @click="emit('trash')"
    >
      Move to trash
    </UiButton>
    <UiButton
      size="sm"
      variant="ghost"
      data-testid="clear-selection"
      @click="props.lib.clearSelection()"
    >
      Clear selection
    </UiButton>
  </div>
</template>

<style scoped>
.selection-bar {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
  margin-bottom: var(--space-4);
  border-radius: var(--radius-md);
  background: var(--surface-selected);
}

.selection-count {
  font-weight: var(--weight-bold);
  color: var(--text-accent);
  margin-right: auto;
}
</style>
