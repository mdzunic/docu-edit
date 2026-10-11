<script setup lang="ts">
// The workbook's status bar (SPEC-117 "Aggregates", "Zoom"): while the
// selection covers more than one cell, "Average: …", "Count: …" and
// "Sum: …" (computed by useSheetSelection through lib/spreadsheet/
// aggregates.ts, over stored cells only — Average and Sum only when numbers
// are selected and no error is, in the number format of the first numeric
// cell, SPEC-118), then the number-locale control (SheetLocaleControl,
// SPEC-118) and SPEC-063's ZoomControl at its right, with its fit entries
// hidden (a sheet has no page to fit). SPEC-120 adds two regions after the
// aggregates: the circular-reference warning ("Circular reference:
// Loop!A1") while the workbook has a cycle, and "Calculating…" while a long
// recalculation runs (useRecalc). SPEC-122 puts the cell editor's mode at
// its left, as Excel does: "Ready", "Enter", "Edit" or "Point"
// (`sheet-status-mode`, useCellEditing's modeLabel; empty hides it).
// SPEC-128: after the mode, while the sheet's autofilter holds a criterion,
// "8 of 12 records found" (`sheet-status-filter`, useSortFilter's). Runtime
// prop declarations per the espree constraint (see AppShell.vue).
import { CALCULATING_TEXT } from "../../views/spreadsheet/useRecalc";
import ZoomControl from "../ZoomControl.vue";
import SheetLocaleControl from "./SheetLocaleControl.vue";

const props = defineProps({
  /** The cell editor's mode (SPEC-122): "Ready", "Enter", "Edit", "Point" — empty hides it. */
  mode: { type: String, default: "" },
  /** "8 of 12 records found" (SPEC-128) — empty hides it. */
  filter: { type: String, default: "" },
  /** "Average: 20" — empty hides it. */
  average: { type: String, default: "" },
  /** "Count: 3" — empty hides it. */
  count: { type: String, default: "" },
  /** "Sum: 60" — empty hides it. */
  sum: { type: String, default: "" },
  /** "Circular reference: Loop!A1" (SPEC-120) — empty hides it. */
  circular: { type: String, default: "" },
  /** A recalculation has run past CALCULATING_STATUS_DELAY_MS (SPEC-120). */
  calculating: { type: Boolean, default: false },
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
    <span
      v-if="props.mode"
      class="sheet-status-mode"
      data-testid="sheet-status-mode"
    >{{ props.mode }}</span>
    <span
      v-if="props.filter"
      class="sheet-status-filter"
      role="status"
      data-testid="sheet-status-filter"
    >{{ props.filter }}</span>
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
    <div
      class="sheet-status-calc"
      role="status"
      aria-live="polite"
    >
      <span
        v-if="props.circular"
        class="sheet-status-circular"
        data-testid="sheet-status-circular"
      >{{ props.circular }}</span>
      <span
        v-if="props.calculating"
        data-testid="sheet-status-calculating"
      >{{ CALCULATING_TEXT }}</span>
    </div>
    <SheetLocaleControl class="sheet-status-locale" />
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

/* The editing mode: one word, never squeezed. */
.sheet-status-mode {
  flex: none;
  white-space: nowrap;
}

/* The filter's record count: one phrase, never squeezed. */
.sheet-status-filter {
  flex: none;
  white-space: nowrap;
}

/* The aggregates take what the locale and the zoom leave. */
.sheet-status-aggregates {
  display: flex;
  flex: 1 1 auto;
  gap: var(--space-4);
  min-width: 0;
  overflow: hidden;
  white-space: nowrap;
}

.sheet-status-calc {
  display: flex;
  flex: none;
  gap: var(--space-4);
  white-space: nowrap;
}

.sheet-status-circular {
  color: var(--text-danger);
}

.sheet-status-locale {
  flex: none;
}
</style>
