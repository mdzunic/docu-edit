<script setup lang="ts">
// The interactive selection chrome for the top-level selected frame
// (SPEC-082, rotation + adjustments SPEC-096): eight resize handles over the
// frame box — the box itself rotated with the shape so handles sit on its own
// corners — plus a ninth rotate-handle above the n handle, one diamond
// adjust-handle per declared shape adjustment, the aria-live rotation
// readout, and the centre/edge alignment guides that fire while dragging
// (now tagged data-guide-kind="slide"|"frame"). Rendered by SlideStage in
// slide space (absolute cm → px), above the frames. Pointer-driven with touch
// equivalents (pointer events, no hover-only affordances, SPEC-068). Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSlideFrame, asStrings, cmPx } from "./slideRender";
import { RESIZE_HANDLES } from "../../views/presentation/stageMath";
import { rotatedResizeCursor } from "../../lib/core/shapeGeometry";
import { adjustmentHandles } from "../../lib/shapeAdjustHandles";

const props = defineProps({
  /** The selected SlideFrame (presentation model). */
  frame: { type: Object, required: true },
  /** Active alignment guides: slide x-positions (vertical) / y-positions. */
  guides: { type: Object, default: () => ({ vertical: [], horizontal: [] }) },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  /** The live angle while a rotate drag runs (the readout), else null. */
  rotatingDeg: { type: Number, default: null },
});

const emit = defineEmits([
  "handle-pointerdown",
  "edit-request",
  "rotate-pointerdown",
  "adjust-pointerdown",
  "adjust-keydown",
]);

const frame = computed(() => asSlideFrame(props.frame));
const rotationDeg = computed(() => frame.value.rotationDeg ?? 0);

const boxStyle = computed(() => {
  const style = {
    left: cmPx(frame.value.xCm),
    top: cmPx(frame.value.yCm),
    width: cmPx(frame.value.widthCm),
    height: cmPx(frame.value.heightCm),
    transform: "",
    transformOrigin: "",
  };
  // The sel-box rotates with the shape so handles sit on its own corners.
  if (rotationDeg.value) {
    style.transform = `rotate(${rotationDeg.value}deg)`;
    style.transformOrigin = "center";
  }
  return style;
});

function cursorFor(handle = "nw") {
  return rotatedResizeCursor(handle, rotationDeg.value);
}

// One diamond per declared shape adjustment (SPEC-069's specs), positioned
// in unit-square fractions of the frame box.
const adjustHandleList = computed(() => {
  const f = frame.value;
  if (f.kind !== "shape") return [];
  return adjustmentHandles(f.geometry, f.adjustments).map((h) => ({
    ...h,
    style: { left: `${(h.xFrac * 100).toFixed(2)}%`, top: `${(h.yFrac * 100).toFixed(2)}%` },
  }));
});

const readout = computed(() => (props.rotatingDeg === null ? "" : `${props.rotatingDeg}°`));

function guideKind(kinds = undefined, index = 0) {
  return asStrings(kinds)[Number(index)] || "slide";
}

function onHandleDown(handle = "nw", event = new PointerEvent("pointerdown")) {
  emit("handle-pointerdown", handle, event);
}

function onRotateDown(event = new PointerEvent("pointerdown")) {
  emit("rotate-pointerdown", event);
}

function onAdjustDown(index = 0, event = new PointerEvent("pointerdown")) {
  emit("adjust-pointerdown", index, event);
}

function onAdjustKeydown(index = 0, event = new KeyboardEvent("keydown")) {
  emit("adjust-keydown", index, event);
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
    class="overlay print-exclude"
    aria-hidden="false"
  >
    <!-- Alignment guides span the whole surface -->
    <div
      v-for="(x, i) in guides.vertical"
      :key="`v${i}`"
      class="guide guide-v"
      data-testid="align-guide"
      :data-guide-kind="guideKind(guides.verticalKinds, Number(i))"
      :style="{ left: cmPx(x) }"
    />
    <div
      v-for="(y, i) in guides.horizontal"
      :key="`h${i}`"
      class="guide guide-h"
      data-testid="align-guide"
      :data-guide-kind="guideKind(guides.horizontalKinds, Number(i))"
      :style="{ top: cmPx(y) }"
    />
    <!-- Selection box + handles over the frame, rotated with the shape -->
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
        :style="{ cursor: cursorFor(h) }"
        :data-handle="h"
        data-testid="resize-handle"
        @mousedown.stop
        @pointerdown="onHandleDown(h, $event)"
        @dblclick="onOverlayDblclick($event)"
      />
      <!-- The ninth handle: rotate, above the n handle (SPEC-096) -->
      <span
        class="rotate-handle"
        data-testid="rotate-handle"
        role="slider"
        tabindex="0"
        aria-label="Rotate"
        :aria-valuenow="frame.rotationDeg ?? 0"
        :aria-valuemin="0"
        :aria-valuemax="359"
        @mousedown.stop
        @pointerdown="onRotateDown($event)"
      />
      <!-- One diamond per declared shape adjustment (SPEC-096) -->
      <span
        v-for="h in adjustHandleList"
        :key="`adj-${h.index}`"
        class="adjust-handle"
        data-testid="adjust-handle"
        role="slider"
        tabindex="0"
        :aria-label="h.label"
        :aria-valuenow="Math.round(h.value * 100)"
        :aria-valuemin="0"
        :aria-valuemax="100"
        :style="h.style"
        @mousedown.stop
        @pointerdown="onAdjustDown(h.index, $event)"
        @keydown="onAdjustKeydown(h.index, $event)"
      />
    </div>
    <!-- The live rotation angle (SPEC-096) -->
    <div
      class="rotation-readout"
      data-testid="rotation-readout"
      aria-live="polite"
      :class="{ 'is-visible': readout !== '' }"
    >
      {{ readout }}
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
  background: var(--accent);
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
  outline: var(--border-width) solid var(--accent);
  pointer-events: none;
}

.handle {
  position: absolute;
  width: 10px;
  height: 10px;
  /* translate centres the handle on its corner/edge point (no length token
     needed for the -50% offset). */
  transform: translate(-50%, -50%);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
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

/* The rotate handle: a round grip floating above the n handle, joined to the
   box by a short stem. */
.rotate-handle {
  position: absolute;
  top: -22px;
  left: 50%;
  width: 12px;
  height: 12px;
  transform: translate(-50%, -50%);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--accent);
  border-radius: 50%;
  pointer-events: auto;
  touch-action: none;
  cursor: grab;
}

/* A small yellow diamond, deliberately distinct from the square resize
   handles (SPEC-096 — the SPEC-069 look). */
.adjust-handle {
  position: absolute;
  width: 9px;
  height: 9px;
  transform: translate(-50%, -50%) rotate(45deg);
  background: var(--warning, #f4b400);
  border: var(--border-width) solid var(--accent);
  pointer-events: auto;
  touch-action: none;
  cursor: crosshair;
}

.adjust-handle:focus-visible,
.rotate-handle:focus-visible {
  outline: 2px solid var(--focus-ring);
  outline-offset: 1px;
}

.rotation-readout {
  position: absolute;
  top: var(--space-2);
  left: 50%;
  transform: translateX(-50%);
  padding: 0 var(--space-2);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  color: var(--text-primary);
  font-size: var(--text-sm);
  visibility: hidden;
}

.rotation-readout.is-visible {
  visibility: visible;
}
</style>
