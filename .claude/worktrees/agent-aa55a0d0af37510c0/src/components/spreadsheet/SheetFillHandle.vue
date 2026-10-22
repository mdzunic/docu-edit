<script setup lang="ts">
// SPEC-125: the fill handle (`sheet-fill-handle`) — a 7×7 px square on the
// selection's bottom-right corner, `cursor: crosshair`. It takes the pointer
// (the overlays around it do not): the grid's pointer handler starts a fill
// drag from a press on it and a double-click fills down (useFillHandle).
// Placed in paper px by useSheetGrid. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
const props = defineProps({
  left: { type: Number, default: 0 },
  top: { type: Number, default: 0 },
});
</script>

<template>
  <div
    class="sheet-fill-handle"
    data-testid="sheet-fill-handle"
    title="Drag to fill; double-click to fill down"
    :style="{ left: `${props.left}px`, top: `${props.top}px` }"
  />
</template>

<style scoped>
.sheet-fill-handle {
  position: absolute;
  z-index: 1;
  box-sizing: border-box;
  width: 7px;
  height: 7px;
  background: var(--accent);
  border: var(--border-width) solid var(--surface-raised);
  cursor: crosshair;
  pointer-events: auto;
}
</style>
