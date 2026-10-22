<script setup lang="ts">
// The selected shape's geometry adjustment handles (SPEC-069 AC-5): one
// draggable yellow diamond per declared adjustment, visually distinct from
// the square resize handles and round rotation handle. Extracted from
// DocumentView so its wiring stays a single element there (DocumentView.size
// budget); the positions, drag mapping and keyboard nudge all live in
// useShapes.ts, which passes the already-styled handle list in.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asAdjustmentHandleViews } from "../lib/shapeAdjustHandles";

const props = defineProps({
  /** The handle list from useShapes.adjustmentHandleList(): each carries its
   * index, label, value and an absolute { left, top } pixel style. */
  handles: { type: Array, default: () => [] },
});

defineEmits(["adjust-start", "adjust-keydown"]);

// The runtime Array prop carries no element type (espree bars PropType here);
// this cast gives the template typed handles.
const handleList = computed(() => asAdjustmentHandleViews(props.handles));
</script>

<template>
  <div
    v-for="handle in handleList"
    :key="`adjust-${handle.index}`"
    class="shape-adjust-handle"
    :data-testid="`shape-adjust-${handle.index}`"
    role="slider"
    tabindex="0"
    :aria-label="`${handle.label} adjustment`"
    :aria-valuenow="Math.round(handle.value * 100)"
    :aria-valuemin="0"
    :aria-valuemax="100"
    :title="handle.label"
    :style="handle.style"
    @mousedown="$emit('adjust-start', $event, handle.index)"
    @keydown="$emit('adjust-keydown', $event, handle.index)"
  />
</template>

<style scoped>
/* A small yellow diamond (rotated square), deliberately distinct from the
   square resize handles and round rotation handle (AC-5b). */
.shape-adjust-handle {
  position: absolute;
  pointer-events: auto;
  width: 8px;
  height: 8px;
  border: 1px solid var(--accent);
  background: var(--warning, #f4b400);
  transform: rotate(45deg);
  cursor: crosshair;
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

.shape-adjust-handle:focus-visible {
  outline: 2px solid var(--focus-ring);
  outline-offset: 1px;
}
</style>
