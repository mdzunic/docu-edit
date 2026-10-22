<script setup lang="ts">
// A selected sheet drawing's eight resize handles (SPEC-131 "The drawing
// layer"): `sheet-drawing-handle-<nw|n|ne|e|se|s|sw|w>`, on its corners and
// edge midpoints, each with its resize cursor. They sit inside the drawing's
// element (so they follow its position and rotation) and are a constant
// on-screen size: 8 px divided by the grid's zoom, since the zoom transform
// scales everything in the content box. A press emits press(handle, event) —
// useDrawingGestures resizes from it. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { DRAWING_HANDLES } from "../../views/spreadsheet/sheetDrawingLayout";

const props = defineProps({
  /** The grid's zoom as a fraction. */
  zoom: { type: Number, default: 1 },
});

const emit = defineEmits(["press"]);

/** Where each handle's centre sits, as fractions of the drawing's box. */
const AT = {
  nw: [0, 0],
  n: [0.5, 0],
  ne: [1, 0],
  e: [1, 0.5],
  se: [1, 1],
  s: [0.5, 1],
  sw: [0, 1],
  w: [0, 0.5],
};
const CURSORS = { nw: "nwse-resize", se: "nwse-resize", ne: "nesw-resize", sw: "nesw-resize", n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize" };

const size = computed(() => 8 / Math.max(0.01, props.zoom));

function handleStyle(handle = "se") {
  const [x, y] = Reflect.get(AT, handle) ?? [1, 1];
  const half = size.value / 2;
  return {
    left: `calc(${x * 100}% - ${half}px)`,
    top: `calc(${y * 100}% - ${half}px)`,
    width: `${size.value}px`,
    height: `${size.value}px`,
    borderWidth: `${1 / Math.max(0.01, props.zoom)}px`,
    cursor: Reflect.get(CURSORS, handle),
  };
}
</script>

<template>
  <div
    class="sheet-drawing-handles"
    aria-hidden="true"
  >
    <div
      v-for="handle in DRAWING_HANDLES"
      :key="handle"
      class="sheet-drawing-handle"
      :data-testid="`sheet-drawing-handle-${handle}`"
      :data-handle="handle"
      :style="handleStyle(handle)"
      @pointerdown="emit('press', handle, $event)"
    />
  </div>
</template>

<style scoped>
.sheet-drawing-handles {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.sheet-drawing-handle {
  position: absolute;
  box-sizing: border-box;
  pointer-events: auto;
  background: var(--surface-base);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
}
</style>
