<script setup lang="ts">
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
      <button
        type="button"
        class="chart-toolbar-option"
        data-testid="chart-edit-data-button"
        @mousedown.prevent
        @click="$emit('edit-data')"
      >
        Edit data
      </button>
      <button
        type="button"
        class="chart-toolbar-option"
        data-testid="chart-options-button"
        @mousedown.prevent
        @click="$emit('options')"
      >
        Chart options
      </button>
    </template>
    <span class="chart-toolbar-separator" />
    <button
      type="button"
      class="chart-toolbar-option"
      data-testid="chart-duplicate-button"
      title="Duplicate"
      @mousedown.prevent
      @click="$emit('duplicate')"
    >
      Duplicate
    </button>
    <button
      type="button"
      class="chart-toolbar-option"
      data-testid="chart-delete-button"
      title="Delete chart"
      @mousedown.prevent
      @click="$emit('delete')"
    >
      Delete
    </button>
  </div>
</template>

<style scoped>
.chart-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.25rem;
  padding: 0.25rem;
  max-width: 480px;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.chart-toolbar-notice {
  font-size: 0.75rem;
  color: #5f6368;
  padding: 0.25rem 0.5rem;
}

.chart-toolbar-option {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  white-space: nowrap;
}

.chart-toolbar-option:hover {
  background: #f1f3f4;
}

.chart-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: #dadce0;
}
</style>
