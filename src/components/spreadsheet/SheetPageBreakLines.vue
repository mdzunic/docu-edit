<script setup lang="ts">
// View ▸ Page breaks (SPEC-136): a line at each page break of the active
// sheet — `data-axis` row|col, `data-index` the first row/column of the next
// page, `data-kind` auto (dashed) or manual (solid) — inside the grid's
// overlay layer, positioned by the grid (`position(axis, index)`: the
// boundary's offset in the layer, null when scrolled under a frozen pane).
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { placedBreakLines } from "../../views/spreadsheet/usePageSetup";

const props = defineProps({
  /** PageBreakLine[]. */
  lines: { type: Array, required: true },
  position: { type: Function, required: true },
});

function placed() {
  return placedBreakLines(props.lines, props.position);
}
</script>

<template>
  <div
    v-for="line in placed()"
    :key="`${line.axis}-${line.index}`"
    class="sheet-page-break-line"
    :class="[`is-${line.axis}`, `is-${line.kind}`]"
    data-testid="sheet-page-break-line"
    :data-axis="line.axis"
    :data-index="line.index"
    :data-kind="line.kind"
    :style="line.axis === 'col' ? { left: `${line.at - 1}px` } : { top: `${line.at - 1}px` }"
  />
</template>

<style scoped>
.sheet-page-break-line {
  position: absolute;
  pointer-events: none;
  z-index: 3;
}

.sheet-page-break-line.is-col {
  top: 0;
  bottom: 0;
  width: 0;
  border-left: var(--border-width) dashed var(--accent);
}

.sheet-page-break-line.is-row {
  left: 0;
  right: 0;
  height: 0;
  border-top: var(--border-width) dashed var(--accent);
}

.sheet-page-break-line.is-manual.is-col {
  border-left: calc(var(--border-width) * 2) solid var(--accent);
}

.sheet-page-break-line.is-manual.is-row {
  border-top: calc(var(--border-width) * 2) solid var(--accent);
}
</style>
