<script setup lang="ts">
import { computed } from "vue";
import UiButton from "./ui/UiButton.vue";
import { useTier } from "../composables/useResponsiveTier";
// The chart toolbar's controls (SPEC-052): "Edit data" (AC-5), "Chart
// options" and delete/duplicate — disabled for an unsupported chart type
// (AC-30), whose reason shows instead (AC-31). Positioning, chrome, ARIA and
// dismissal live on the shared ContextualToolbar (SPEC-061 AC-22, AC-23); this
// supplies only its own controls. SPEC-132: an optional `reason` replaces
// the unsupported notice's default text (a read-only sheet chart says why,
// "Pivot charts can't be edited yet"); without one the text is unchanged.
defineProps({
  unsupported: { type: Boolean, default: false },
  reason: { type: String, default: "" },
});

defineEmits(["edit-data", "options", "delete", "duplicate"]);

// SPEC-068 AC-12/13 — the chart data sheet is one of the commands disabled at
// the compact tier (a spreadsheet grid at 390px is not usable). The chart
// itself still renders and can be moved/deleted; only "Edit data" is off,
// shown with a stated reason rather than removed.
const tier = useTier();
const dataSheetUnavailable = computed(() => tier.value === "compact");
const DATA_SHEET_UNAVAILABLE_REASON = "not available on this screen size";
</script>

<template>
  <template v-if="unsupported">
    <span
      class="chart-toolbar-notice"
      data-testid="chart-unsupported-reason"
    >
      <template v-if="reason">{{ reason }}</template>
      <template v-else>We can't edit this chart type yet — it's kept exactly as it was.</template>
    </span>
  </template>
  <template v-else>
    <UiButton
      variant="ghost"
      data-testid="chart-edit-data-button"
      :disabled="dataSheetUnavailable"
      :title="dataSheetUnavailable ? DATA_SHEET_UNAVAILABLE_REASON : 'Edit data'"
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
</template>

<style scoped>
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
