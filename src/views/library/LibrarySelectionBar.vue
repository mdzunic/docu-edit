<script setup lang="ts">
// Bulk action bar (SPEC-067 AC-27..32). Pure-lib actions (download,
// duplicate, star, clear) run directly on the controller; "Move to trash"
// emits so HomeView can raise the Undo toast it owns, and "Convert to PDF…"
// (SPEC-108) emits the one selected id for HomeView's dialog. That action
// shows only while every selected entry is an image, and converts exactly
// one: with more selected it stays visible but disabled, saying why.
// "Combine…" (SPEC-109) emits the selected image ids for HomeView to open
// the assembly workspace with; it is enabled from two image entries up —
// and (SPEC-111) from one PDF, whose pages the workspace then organises,
// the PDFs' ids travelling with the images' in selection order.
import { computed } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import { isImageFormat } from "../../lib/libraryMenu";
import { isCombinableFormat } from "../assembly/useAssemblyTray";

const props = defineProps({
  lib: { type: Object, required: true },
});

const emit = defineEmits(["trash", "convert-pdf", "combine"]);

const allImages = computed(() => {
  const ids = props.lib.selectedIds.value;
  return ids.length > 0 && ids.every((id = "") => isImageFormat(props.lib.formatOf(id)));
});
const convertOne = computed(() => allImages.value && props.lib.selectedIds.value.length === 1);
const combinableIds = computed(() => props.lib.selectedIds.value.filter((id = "") => isCombinableFormat(props.lib.formatOf(id))));
const canCombine = computed(() => {
  const pdfs = combinableIds.value.filter((id = "") => props.lib.formatOf(id) === "pdf").length;
  return pdfs > 0 || combinableIds.value.length >= 2;
});
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
      v-if="allImages"
      size="sm"
      data-testid="bulk-convert-pdf"
      :disabled="!convertOne"
      :title="convertOne ? undefined : 'Convert to PDF works on one image at a time'"
      @click="emit('convert-pdf', props.lib.selectedIds.value[0])"
    >
      Convert to PDF…
    </UiButton>
    <UiButton
      size="sm"
      data-testid="bulk-combine"
      :disabled="!canCombine"
      :title="canCombine ? undefined : 'Select two or more images, or a PDF, to combine them'"
      @click="emit('combine', combinableIds)"
    >
      Combine…
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
