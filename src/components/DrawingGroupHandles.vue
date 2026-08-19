<script setup lang="ts">
// The resize + rotation handles for a selected drawing group or member
// (SPEC-057 AC-19/20/21, AC-28/29). Positioned over the selected drawing's
// measured box (page-relative px) and turned by its rotation about the box
// centre, exactly as a single shape's handles are. A sealed container shows
// none (DocumentView gates on selectedHasHandles). Runtime prop/emit
// declarations (espree).
const props = defineProps({
  /** The selected drawing's measured box, page-relative px. */
  handles: { type: Object, default: null },
  rotationDeg: { type: Number, default: 0 },
});

defineEmits(["resize-start", "rotate-start"]);

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

function boxStyle() {
  const h = props.handles || { left: 0, top: 0, width: 0, height: 0 };
  return {
    left: `${h.left}px`,
    top: `${h.top}px`,
    width: `${h.width}px`,
    height: `${h.height}px`,
    transform: `rotate(${props.rotationDeg}deg)`,
  };
}

function handleStyle(handle = "se") {
  const x = handle === "nw" || handle === "w" || handle === "sw" ? 0 : handle === "n" || handle === "s" ? 50 : 100;
  const y = handle === "nw" || handle === "n" || handle === "ne" ? 0 : handle === "w" || handle === "e" ? 50 : 100;
  return { left: `${x}%`, top: `${y}%` };
}
</script>

<template>
  <div
    class="group-handles print-exclude"
    data-testid="group-handles"
    :style="boxStyle()"
  >
    <div
      v-for="handle in HANDLES"
      :key="handle"
      class="group-resize-handle"
      :data-testid="`group-resize-${handle}`"
      :style="handleStyle(handle)"
      @mousedown="$emit('resize-start', handle, $event)"
    />
    <div
      class="group-rotate-handle"
      data-testid="group-rotation-handle"
      @mousedown="$emit('rotate-start', $event)"
    />
  </div>
</template>

<style scoped>
.group-handles {
  position: absolute;
  z-index: 25;
  pointer-events: none;
}

.group-resize-handle {
  position: absolute;
  width: 10px;
  height: 10px;
  margin: -5px 0 0 -5px;
  background: #fff;
  border: 1px solid #1a73e8;
  border-radius: 2px;
  pointer-events: auto;
  cursor: pointer;
}

.group-rotate-handle {
  position: absolute;
  left: 50%;
  top: -22px;
  width: 12px;
  height: 12px;
  margin-left: -6px;
  background: #fff;
  border: 1px solid #1a73e8;
  border-radius: 50%;
  pointer-events: auto;
  cursor: grab;
}
</style>
