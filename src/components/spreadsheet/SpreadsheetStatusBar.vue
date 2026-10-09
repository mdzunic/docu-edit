<script setup lang="ts">
// The workbook's status bar (SPEC-117 "Aggregates", "Zoom"): while the
// selection covers more than one cell, "Average: …", "Count: …" and
// "Sum: …" (computed by useSheetSelection through lib/spreadsheet/
// aggregates.ts, over stored cells only — Average and Sum only when numbers
// are selected and no error is), and SPEC-063's ZoomControl at its right,
// with its fit entries hidden (a sheet has no page to fit). Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import ZoomControl from "../ZoomControl.vue";

const props = defineProps({
  /** "Average: 20" — empty hides it. */
  average: { type: String, default: "" },
  /** "Count: 3" — empty hides it. */
  count: { type: String, default: "" },
  /** "Sum: 60" — empty hides it. */
  sum: { type: String, default: "" },
  /** The sheet's zoom in percent. */
  zoomPercent: { type: Number, default: 100 },
  /** No grid on show (a foreign sheet): the zoom is disabled. */
  zoomDisabled: { type: Boolean, default: false },
});

const emit = defineEmits(["zoom-in", "zoom-out", "apply-zoom", "set-zoom"]);
</script>

<template>
  <div
    class="spreadsheet-status-bar"
    data-testid="spreadsheet-status-bar"
  >
    <div
      class="sheet-status-aggregates"
      role="status"
      aria-live="polite"
    >
      <span
        v-if="props.average"
        data-testid="sheet-status-average"
      >{{ props.average }}</span>
      <span
        v-if="props.count"
        data-testid="sheet-status-count"
      >{{ props.count }}</span>
      <span
        v-if="props.sum"
        data-testid="sheet-status-sum"
      >{{ props.sum }}</span>
    </div>
    <ZoomControl
      :zoom-percent="props.zoomPercent"
      :show-fit="false"
      :disabled="props.zoomDisabled"
      @zoom-in="emit('zoom-in')"
      @zoom-out="emit('zoom-out')"
      @apply-input="(raw) => emit('apply-zoom', raw)"
      @set-preset="(percent) => emit('set-zoom', percent)"
    />
  </div>
</template>

<style scoped>
.spreadsheet-status-bar {
  display: flex;
  flex: 1 1 auto;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  min-width: 0;
}

.sheet-status-aggregates {
  display: flex;
  gap: var(--space-4);
  min-width: 0;
  overflow: hidden;
  white-space: nowrap;
}
</style>
