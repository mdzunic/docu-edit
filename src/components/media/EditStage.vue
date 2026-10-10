<script setup lang="ts">
// The image editor's stage (SPEC-112): the preview canvas — painted by
// useEditPreview through the shared renderer, never a CSS filter — with the
// active tool's overlay on top. The crop tool shows the whole croppable
// area (the straightened source) with the crop rectangle's handles:
// focusable buttons that an arrow key moves 1 px (Shift: 10 px) and a drag
// moves freely, one undo entry per drag. The redact and annotate tools turn
// the stage into a drawing surface in output pixels.
// Runtime prop declarations — see AppShell.vue.
import { computed, onBeforeUnmount, onMounted, useTemplateRef } from "vue";
import { CROP_HANDLES } from "../../lib/core/imageEdits";
import { asImageEditor } from "../../views/media/useImageEditor";
import { useEditPreview } from "../../views/media/useEditPreview";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const stageRef = useTemplateRef("stage");
const canvasRef = useTemplateRef("canvas");
const preview = useEditPreview(asImageEditor(props.editor), canvasRef);

// A disconnectable stub gives `observer` an inferred type without a TS
// annotation (barred in .vue scripts).
let observer = { disconnect() {} };
function measure() {
  const el = stageRef.value;
  if (el) {
    preview.setBox(el.clientWidth - 32, el.clientHeight - 32);
  }
}
onMounted(() => {
  measure();
  if (typeof ResizeObserver === "function" && stageRef.value) {
    const ro = new ResizeObserver(measure);
    ro.observe(stageRef.value);
    observer = ro;
  }
});
onBeforeUnmount(() => observer.disconnect());

const scale = computed(() => preview.scale.value);
const drawing = computed(() => e.value.tool.value === "redact" || e.value.tool.value === "annotate");

// The crop rectangle on the preview: the preview shows the inscribed area,
// so the rectangle is offset by its origin and scaled.
const cropBox = computed(() => {
  const rect = e.value.crop.rect.value;
  const bounds = e.value.crop.bounds.value;
  if (e.value.tool.value !== "crop" || !rect || !bounds) {
    return null;
  }
  return {
    left: (rect.x - bounds.x) * scale.value,
    top: (rect.y - bounds.y) * scale.value,
    width: rect.width * scale.value,
    height: rect.height * scale.value,
  };
});
const shadePath = computed(() => {
  const box = cropBox.value;
  const w = preview.size.value.widthPx;
  const h = preview.size.value.heightPx;
  return box
    ? `M0 0H${w}V${h}H0Z M${box.left} ${box.top}h${box.width}v${box.height}h${-box.width}Z`
    : "";
});

function onHandleKey(handle = "move", event = new KeyboardEvent("keydown")) {
  const known = CROP_HANDLES.find((entry) => entry === handle);
  if (known && e.value.crop.handleKey(known, event)) {
    event.preventDefault();
  }
}

// Pointer drags: captured, measured from where they began, in source px.
let dragStart = { x: 0, y: 0 };
let dragging = "";

function onHandleDown(handle = "move", event = new PointerEvent("pointerdown")) {
  const known = CROP_HANDLES.find((entry) => entry === handle);
  if (!known) return;
  event.preventDefault();
  const target = event.currentTarget;
  if (target instanceof Element && typeof target.setPointerCapture === "function") {
    target.setPointerCapture(event.pointerId);
  }
  if (target instanceof HTMLElement) target.focus();
  dragStart = { x: event.clientX, y: event.clientY };
  dragging = "crop";
  e.value.crop.startDrag(known);
}

function outputPoint(event = new PointerEvent("pointermove")) {
  const surface = event.currentTarget;
  const rect = surface instanceof Element ? surface.getBoundingClientRect() : { left: 0, top: 0 };
  return { x: (event.clientX - rect.left) / scale.value, y: (event.clientY - rect.top) / scale.value };
}

function onSurfaceDown(event = new PointerEvent("pointerdown")) {
  const tool = e.value.tool.value;
  if (tool !== "redact" && tool !== "annotate") return;
  const target = event.currentTarget;
  if (target instanceof Element && typeof target.setPointerCapture === "function") {
    target.setPointerCapture(event.pointerId);
  }
  const point = outputPoint(event);
  dragging = "draw";
  e.value.objects.startDraw(tool, point.x, point.y);
}

function onPointerMove(event = new PointerEvent("pointermove")) {
  if (dragging === "crop") {
    e.value.crop.dragBy(Math.round((event.clientX - dragStart.x) / scale.value), Math.round((event.clientY - dragStart.y) / scale.value));
  } else if (dragging === "draw") {
    const point = outputPoint(event);
    e.value.objects.drawTo(point.x, point.y);
  }
}

function onPointerUp() {
  if (dragging === "crop") {
    e.value.crop.endDrag();
  } else if (dragging === "draw") {
    e.value.objects.endDraw();
  }
  dragging = "";
}
</script>

<template>
  <div
    ref="stage"
    class="edit-stage"
    data-testid="edit-stage"
  >
    <div
      class="frame"
      :style="{ width: `${preview.size.value.widthPx}px`, height: `${preview.size.value.heightPx}px` }"
    >
      <canvas
        ref="canvas"
        class="preview"
        data-testid="edit-preview"
        role="img"
        :aria-label="`Preview, ${e.geometry.value?.widthPx ?? 0} × ${e.geometry.value?.heightPx ?? 0} px`"
        :data-painted="preview.painted.value"
      />
      <template v-if="cropBox">
        <svg
          class="crop-shade"
          aria-hidden="true"
          :viewBox="`0 0 ${preview.size.value.widthPx} ${preview.size.value.heightPx}`"
        >
          <path
            :d="shadePath"
            fill-rule="evenodd"
          />
        </svg>
        <div
          class="crop-box"
          data-testid="crop-box"
          :style="{ left: `${cropBox.left}px`, top: `${cropBox.top}px`, width: `${cropBox.width}px`, height: `${cropBox.height}px` }"
        >
          <button
            v-for="handle in CROP_HANDLES"
            :key="handle"
            type="button"
            class="crop-handle"
            :class="`at-${handle}`"
            :data-testid="`crop-handle-${handle}`"
            :aria-label="`${e.crop.handleName(handle)} — ${e.crop.announcement.value}`"
            @keydown="onHandleKey(handle, $event)"
            @pointerdown="onHandleDown(handle, $event)"
            @pointermove="onPointerMove"
            @pointerup="onPointerUp"
            @pointercancel="onPointerUp"
          />
        </div>
      </template>
      <div
        v-if="drawing"
        class="draw-surface"
        data-testid="edit-draw-surface"
        :aria-label="e.tool.value === 'redact' ? 'Drag to cover an area' : 'Drag to draw'"
        @pointerdown="onSurfaceDown"
        @pointermove="onPointerMove"
        @pointerup="onPointerUp"
        @pointercancel="onPointerUp"
      />
    </div>
  </div>
</template>

<style scoped>
.edit-stage {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  background: var(--surface-sunken);
}

.frame {
  position: relative;
  flex: none;
}

.preview {
  display: block;
  width: 100%;
  height: 100%;
}

.crop-shade {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  pointer-events: none;
  fill: color-mix(in srgb, var(--surface-sunken) 60%, transparent);
}

.crop-box {
  position: absolute;
  border: var(--border-width) dashed var(--accent);
}

.crop-handle {
  position: absolute;
  width: 0.875rem;
  height: 0.875rem;
  margin: -0.4375rem 0 0 -0.4375rem;
  padding: 0;
  border: var(--border-width) solid var(--surface-base);
  border-radius: 50%;
  background: var(--accent);
  touch-action: none;
  cursor: pointer;
  z-index: 1;
}

.crop-handle:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.crop-handle.at-move {
  inset: 0;
  width: auto;
  height: auto;
  margin: 0;
  border: none;
  border-radius: 0;
  background: transparent;
  cursor: move;
  z-index: 0;
}

.crop-handle.at-nw { left: 0; top: 0; cursor: nwse-resize; }
.crop-handle.at-n { left: 50%; top: 0; cursor: ns-resize; }
.crop-handle.at-ne { left: 100%; top: 0; cursor: nesw-resize; }
.crop-handle.at-e { left: 100%; top: 50%; cursor: ew-resize; }
.crop-handle.at-se { left: 100%; top: 100%; cursor: nwse-resize; }
.crop-handle.at-s { left: 50%; top: 100%; cursor: ns-resize; }
.crop-handle.at-sw { left: 0; top: 100%; cursor: nesw-resize; }
.crop-handle.at-w { left: 0; top: 50%; cursor: ew-resize; }

.draw-surface {
  position: absolute;
  inset: 0;
  cursor: crosshair;
  touch-action: none;
}
</style>
