<script setup lang="ts">
// SPEC-125: the marching ants around a copied or cut range
// (`sheet-copy-marquee`, `data-mode="copy|cut"`) — view state from
// useSheetClipboard, never in the undo stack. Dashes travel round the
// outline; under prefers-reduced-motion they stand still. Placed in paper px
// by useSheetGrid; no pointer events. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
const props = defineProps({
  left: { type: Number, default: 0 },
  top: { type: Number, default: 0 },
  width: { type: Number, default: 0 },
  height: { type: Number, default: 0 },
  /** "copy" | "cut". */
  mode: { type: String, default: "copy" },
});
</script>

<template>
  <div
    class="sheet-copy-marquee"
    data-testid="sheet-copy-marquee"
    :data-mode="props.mode"
    :style="{ left: `${props.left}px`, top: `${props.top}px`, width: `${props.width}px`, height: `${props.height}px` }"
  />
</template>

<style scoped>
/* Four dashed edges drawn as repeating gradients, so the dashes can move:
   4 px of accent, 4 px of nothing. */
.sheet-copy-marquee {
  position: absolute;
  box-sizing: border-box;
  pointer-events: none;
  background-image:
    linear-gradient(90deg, var(--accent) 50%, transparent 50%),
    linear-gradient(90deg, var(--accent) 50%, transparent 50%),
    linear-gradient(0deg, var(--accent) 50%, transparent 50%),
    linear-gradient(0deg, var(--accent) 50%, transparent 50%);
  background-repeat: repeat-x, repeat-x, repeat-y, repeat-y;
  background-size: 8px 2px, 8px 2px, 2px 8px, 2px 8px;
  background-position: 0 0, 0 100%, 0 0, 100% 0;
  animation: sheet-marching-ants 0.8s linear infinite;
}

@keyframes sheet-marching-ants {
  to {
    background-position: 8px 0, -8px 100%, 0 -8px, 100% 8px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .sheet-copy-marquee {
    animation: none;
  }
}
</style>
