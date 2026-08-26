<script setup lang="ts">
// The interactive selection chrome for the top-level selected frame (SPEC-082):
// eight resize handles positioned over the frame box, plus the centre/edge
// alignment guides that fire while dragging. Rendered by SlideStage in slide
// space (absolute cm → px), above the frames. Pointer-driven with touch
// equivalents (pointer events, no hover-only affordances, SPEC-068). Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSlideFrame, cmPx } from "./slideRender";
import { RESIZE_HANDLES } from "../../views/presentation/stageMath";

const props = defineProps({
  /** The selected SlideFrame (presentation model). */
  frame: { type: Object, required: true },
  /** Active alignment guides: slide x-positions (vertical) / y-positions. */
  guides: { type: Object, default: () => ({ vertical: [], horizontal: [] }) },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
});

const emit = defineEmits(["handle-pointerdown", "edit-request"]);

const frame = computed(() => asSlideFrame(props.frame));

const boxStyle = computed(() => ({
  left: cmPx(frame.value.xCm),
  top: cmPx(frame.value.yCm),
  width: cmPx(frame.value.widthCm),
  height: cmPx(frame.value.heightCm),
}));

const CURSORS = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

function onHandleDown(handle = "nw", event = new PointerEvent("pointerdown")) {
  emit("handle-pointerdown", handle, event);
}

// A double-click on the selection chrome (over a text frame's edge, where a
// handle would otherwise swallow it) still opens a text-editing session at the
// click point — the parent decides whether the frame is editable.
function onOverlayDblclick(event = new MouseEvent("dblclick")) {
  emit("edit-request", { x: event.clientX, y: event.clientY });
}
</script>

<template>
  <div
    class="overlay"
    aria-hidden="true"
  >
    <!-- Alignment guides span the whole surface -->
    <div
      v-for="(x, i) in guides.vertical"
      :key="`v${i}`"
      class="guide guide-v"
      data-testid="align-guide"
      :style="{ left: cmPx(x) }"
    />
    <div
      v-for="(y, i) in guides.horizontal"
      :key="`h${i}`"
      class="guide guide-h"
      data-testid="align-guide"
      :style="{ top: cmPx(y) }"
    />
    <!-- Selection box + handles over the frame -->
    <div
      class="sel-box"
      data-testid="frame-selection"
      :style="boxStyle"
    >
      <span
        v-for="h in RESIZE_HANDLES"
        :key="h"
        class="handle"
        :class="`handle-${h}`"
        :style="{ cursor: CURSORS[h] }"
        :data-handle="h"
        data-testid="resize-handle"
        @mousedown.stop
        @pointerdown="onHandleDown(h, $event)"
        @dblclick="onOverlayDblclick($event)"
      />
    </div>
  </div>
</template>

<style scoped>
.overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.guide {
  position: absolute;
  background: var(--accent, #3b82f6);
  opacity: 0.9;
}

.guide-v {
  top: 0;
  bottom: 0;
  width: 1px;
}

.guide-h {
  left: 0;
  right: 0;
  height: 1px;
}

.sel-box {
  position: absolute;
  outline: 1px solid var(--accent, #3b82f6);
  pointer-events: none;
}

.handle {
  position: absolute;
  width: 10px;
  height: 10px;
  margin: -5px;
  background: #fff;
  border: 1px solid var(--accent, #3b82f6);
  border-radius: 1px;
  pointer-events: auto;
  touch-action: none;
}

.handle-nw { top: 0; left: 0; }
.handle-n { top: 0; left: 50%; }
.handle-ne { top: 0; left: 100%; }
.handle-e { top: 50%; left: 100%; }
.handle-se { top: 100%; left: 100%; }
.handle-s { top: 100%; left: 50%; }
.handle-sw { top: 100%; left: 0; }
.handle-w { top: 50%; left: 0; }
</style>
