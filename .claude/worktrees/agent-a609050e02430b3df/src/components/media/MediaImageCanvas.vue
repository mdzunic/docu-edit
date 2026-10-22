<script setup lang="ts">
// The image viewer surface (SPEC-107): an <img> over the blob URL, opened
// fit-to-window, with the SPEC-063 zoom vocabulary — fit, 100%, a zoom
// control, Cmd/Ctrl +/-/0, drag to pan, double-click to toggle fit/100%.
// The zoom/pan STATE lives in useImageViewport (handed in as the `viewport`
// prop, the same controller-prop pattern the library components use); this
// component owns only the DOM — sizing observers, pointer events, keys.
// Runtime prop declarations — the .vue ESLint parser rejects TS-only syntax
// (see AppShell.vue).
import { onBeforeUnmount, onMounted, useTemplateRef } from "vue";

const props = defineProps({
  /** The useImageViewport controller (see views/media/useImageViewport). */
  viewport: { type: Object, required: true },
  /** blob: URL of the verbatim source bytes. */
  src: { type: String, required: true },
  /** The file name — the image's accessible name (its alt). */
  alt: { type: String, required: true },
});

const stageRef = useTemplateRef("stage");

// A disconnectable stub as the initial value gives `observer` an inferred
// { disconnect(): void } type without a TS annotation (barred in .vue).
let observer = { disconnect() {} };

function measure() {
  const el = stageRef.value;
  if (el) {
    props.viewport.setContainerSize(el.clientWidth, el.clientHeight);
  }
}

onMounted(() => {
  measure();
  if (typeof ResizeObserver === "function" && stageRef.value) {
    const ro = new ResizeObserver(measure);
    ro.observe(stageRef.value);
    observer = ro;
  }
  window.addEventListener("keydown", handleKeydown);
});

onBeforeUnmount(() => {
  observer.disconnect();
  window.removeEventListener("keydown", handleKeydown);
});

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (props.viewport.handleZoomKey(event)) {
    event.preventDefault();
  }
}

function onImageLoad(event = new Event("load")) {
  const image = event.target;
  if (image instanceof HTMLImageElement) {
    props.viewport.setImageSize(image.naturalWidth, image.naturalHeight);
  }
}

// Drag to pan: pointer capture so the drag survives leaving the stage.
let dragging = false;
let lastX = 0;
let lastY = 0;

function onPointerDown(event = new PointerEvent("pointerdown")) {
  dragging = true;
  lastX = event.clientX;
  lastY = event.clientY;
  const el = event.currentTarget;
  if (el instanceof Element && typeof el.setPointerCapture === "function") {
    el.setPointerCapture(event.pointerId);
  }
}

function onPointerMove(event = new PointerEvent("pointermove")) {
  if (!dragging) {
    return;
  }
  props.viewport.panBy(event.clientX - lastX, event.clientY - lastY);
  lastX = event.clientX;
  lastY = event.clientY;
}

function onPointerUp() {
  dragging = false;
}
</script>

<template>
  <div class="image-viewer">
    <div class="zoom-bar">
      <button
        type="button"
        class="zoom-button"
        data-testid="media-zoom-fit"
        :class="{ 'is-active': props.viewport.mode.value === 'fit' }"
        :aria-pressed="props.viewport.mode.value === 'fit'"
        @click="props.viewport.zoomFit()"
      >
        Fit
      </button>
      <button
        type="button"
        class="zoom-button"
        data-testid="media-zoom-actual"
        :class="{ 'is-active': props.viewport.mode.value === 'manual' && props.viewport.percent.value === 100 }"
        @click="props.viewport.zoomActual()"
      >
        100%
      </button>
      <button
        type="button"
        class="zoom-button"
        data-testid="media-zoom-out"
        aria-label="Zoom out"
        @click="props.viewport.zoomOut()"
      >
        −
      </button>
      <span
        class="zoom-level"
        data-testid="media-zoom-level"
        aria-live="polite"
      >{{ props.viewport.percent.value }}%</span>
      <button
        type="button"
        class="zoom-button"
        data-testid="media-zoom-in"
        aria-label="Zoom in"
        @click="props.viewport.zoomIn()"
      >
        +
      </button>
    </div>
    <div
      ref="stage"
      class="image-stage"
      data-testid="media-image-stage"
      @pointerdown="onPointerDown"
      @pointermove="onPointerMove"
      @pointerup="onPointerUp"
      @pointercancel="onPointerUp"
      @dblclick="props.viewport.toggleFitActual()"
    >
      <img
        class="image-surface"
        data-testid="media-image"
        :src="props.src"
        :alt="props.alt"
        draggable="false"
        :style="{
          transform: `translate(${props.viewport.pan.value.x}px, ${props.viewport.pan.value.y}px) scale(${props.viewport.scale.value})`,
        }"
        @load="onImageLoad"
      >
    </div>
  </div>
</template>

<style scoped>
.image-viewer {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  flex-direction: column;
}

.zoom-bar {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
}

.zoom-button {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.zoom-button.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.zoom-level {
  min-width: 3.5rem;
  text-align: center;
  color: var(--text-secondary);
  font-variant-numeric: tabular-nums;
}

.image-stage {
  flex: 1 1 auto;
  min-height: 0;
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-sunken);
  cursor: grab;
  touch-action: none;
}

.image-stage:active {
  cursor: grabbing;
}

.image-surface {
  flex: none;
  user-select: none;
  transform-origin: center center;
}
</style>
