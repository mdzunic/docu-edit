<script setup lang="ts">
// The multi-selection chrome (SPEC-096): the union box over every selected
// frame's ROTATED bounds, deliberately carrying NO resize, rotate or adjust
// handles — those gestures are unavailable on a multi-selection (group first,
// or collapse the selection). Rendered by SlideStage in slide space. Runtime
// prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSlideFrames, cmPx } from "./slideRender";
import { frameRotatedBox } from "../../lib/presentation/frames";

const props = defineProps({
  /** The selected SlideFrames (two or more). */
  frames: { type: Array, required: true },
});

const boxStyle = computed(() => {
  const frames = asSlideFrames(props.frames);
  if (frames.length === 0) return null;
  const boxes = frames.map(frameRotatedBox);
  const minX = Math.min(...boxes.map((b) => b.xCm));
  const minY = Math.min(...boxes.map((b) => b.yCm));
  const maxX = Math.max(...boxes.map((b) => b.xCm + b.widthCm));
  const maxY = Math.max(...boxes.map((b) => b.yCm + b.heightCm));
  return {
    left: cmPx(minX),
    top: cmPx(minY),
    width: cmPx(maxX - minX),
    height: cmPx(maxY - minY),
  };
});
</script>

<template>
  <div
    class="overlay print-exclude"
    aria-hidden="true"
  >
    <div
      v-if="boxStyle"
      class="bounds"
      data-testid="selection-bounds"
      :style="boxStyle"
    />
  </div>
</template>

<style scoped>
.overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.bounds {
  position: absolute;
  outline: var(--border-width) dashed var(--accent);
  background: color-mix(in srgb, var(--accent) 4%, transparent);
}
</style>
