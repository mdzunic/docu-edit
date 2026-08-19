<script setup lang="ts">
import UiButton from "./ui/UiButton.vue";
// The floating chart toolbar (SPEC-052): shown near a selected chart,
// offering "Edit data" (opens the data sheet, AC-5), "Chart options" (opens
// the options panel) and delete/duplicate. Mirrors ShapeToolbar.vue's own
// positioning convention, trimmed to this spec's own actions (no fill/
// outline/rotation/z-order/group — a chart has none of those). Disabled for
// an unsupported chart type (AC-30), whose reason shows instead (AC-31).
defineProps({
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  unsupported: { type: Boolean, default: false },
});

defineEmits(["edit-data", "options", "delete", "duplicate"]);

function toolbarStyle(left = 0, top = 0) {
  return { left: `${left}px`, top: `${top}px` };
}
</script>

<template>
  <div
    class="chart-toolbar"
    data-testid="chart-toolbar"
    role="toolbar"
    aria-label="Chart formatting"
    :style="toolbarStyle(left, top)"
  >
    <template v-if="unsupported">
      <span
        class="chart-toolbar-notice"
        data-testid="chart-unsupported-reason"
      >
        We can't edit this chart type yet — it's kept exactly as it was.
      </span>
    </template>
    <template v-else>
      <UiButton
        variant="ghost"
        data-testid="chart-edit-data-button"
        @mousedown.prevent
        @click="$emit('edit-data')"
      >
        Edit data
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="chart-options-button"
        @mousedown.prevent
        @click="$emit('options')"
      >
        Chart options
      </UiButton>
    </template>
    <span class="chart-toolbar-separator" />
    <UiButton
      variant="ghost"
      data-testid="chart-duplicate-button"
      title="Duplicate"
      @mousedown.prevent
      @click="$emit('duplicate')"
    >
      Duplicate
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="chart-delete-button"
      title="Delete chart"
      @mousedown.prevent
      @click="$emit('delete')"
    >
      Delete
    </UiButton>
  </div>
</template>

<style scoped>
.chart-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2);
  max-width: 480px;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.chart-toolbar-notice {
  font-size: var(--text-sm);
  color: var(--text-secondary);
  padding: var(--space-2) var(--space-3);
}

.chart-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--border-default);
}
</style>
