<script setup lang="ts">
// The assembly preview (SPEC-109): the sheet at fit-to-window on a canvas
// drawn by the same composeToCanvas the export uses, with the placed
// rectangles beside it as an ARIA-labelled list — what a screen reader (and
// a test) reads in place of the pixels. Each list entry selects its image;
// in the overlay layout the selected image moves by dragging it on the
// canvas (Shift keeps one axis) or with the arrow keys (1 mm; Shift: 1 cm).
// Logic is useAssemblyPreview's and useAssembly's. Runtime prop declarations
// (espree lints .vue scripts).
import { computed } from "vue";
import { formatLength } from "../../lib/core/units";

const props = defineProps({
  /** The useAssembly() controller. */
  assembly: { type: Object, required: true },
  /** The useAssemblyPreview() controller. */
  preview: { type: Object, required: true },
});

const a = props.assembly;
const p = props.preview;

const unit = computed(() => a.sheet.unit.value);
const sheetLabel = computed(() => {
  const sheet = a.geometry.value;
  if (!sheet) {
    return "Preview of the combined image";
  }
  const count = a.placed.value.length;
  return (
    `Preview of the combined image: ${formatLength(sheet.widthCm, unit.value)} × ` +
    `${formatLength(sheet.heightCm, unit.value)} ${unit.value}, ${count} ${count === 1 ? "image" : "images"}`
  );
});

function entryAt(index = 0) {
  return a.tray.entries.value[index] ?? null;
}

function rectLabel(item = { index: 0, xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 }) {
  const entry = entryAt(item.index);
  const u = unit.value;
  return (
    `${item.index + 1}. ${entry ? entry.title : ""} — ${formatLength(item.widthCm, u)} × ` +
    `${formatLength(item.heightCm, u)} ${u} at ${formatLength(item.xCm, u)}, ${formatLength(item.yCm, u)} ${u}`
  );
}
</script>

<template>
  <section
    class="assembly-preview"
    aria-label="Preview"
  >
    <div
      :ref="p.attachContainer"
      class="preview-stage"
    >
      <canvas
        :ref="p.attachCanvas"
        class="preview-canvas"
        :class="{ 'preview-canvas--overlay': a.isOverlay.value }"
        data-testid="assembly-preview"
        role="img"
        tabindex="0"
        :aria-label="sheetLabel"
        aria-describedby="assembly-placed-list"
        :style="p.display.value ? { width: `${p.display.value.widthPx}px`, height: `${p.display.value.heightPx}px` } : undefined"
        @pointerdown="p.onPointerDown"
        @pointermove="p.onPointerMove"
        @pointerup="p.onPointerUp"
        @pointercancel="p.onPointerUp"
        @keydown="p.onKeydown"
      />
    </div>
    <p
      v-if="a.isOverlay.value"
      class="preview-hint"
    >
      Drag an image to move it (hold Shift to keep one direction); arrow keys nudge it by 1 mm, Shift+arrow by 1 cm.
    </p>
    <ol
      id="assembly-placed-list"
      class="placed-list"
      data-testid="assembly-placed"
      aria-label="Placed images"
    >
      <li
        v-for="item in a.placed.value"
        :key="item.index"
        data-testid="assembly-placed-item"
        :data-index="item.index"
        :data-x-cm="item.xCm"
        :data-y-cm="item.yCm"
        :data-width-cm="item.widthCm"
        :data-height-cm="item.heightCm"
      >
        <button
          type="button"
          class="placed-button"
          :aria-pressed="entryAt(item.index)?.key === a.tray.selectedKey.value ? 'true' : 'false'"
          @click="a.select(entryAt(item.index)?.key ?? '')"
          @keydown="p.onKeydown"
        >
          {{ rectLabel(item) }}
        </button>
      </li>
    </ol>
    <p
      v-if="a.placed.value.length === 0"
      class="preview-empty"
      data-testid="assembly-preview-empty"
    >
      Nothing is placed yet: add images to the tray.
    </p>
  </section>
</template>

<style scoped>
.assembly-preview {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  min-width: 0;
  min-height: 0;
  font-size: var(--text-sm);
  color: var(--text-primary);
}

/* The stage is measured (fit-to-window) and the canvas sized from it, so the
   canvas sits out of the stage's flow: its size can never feed back into
   the size it was computed from. */
.preview-stage {
  position: relative;
  flex: 1 1 0;
  min-height: 12rem;
  overflow: hidden;
}

/* A checkerboard behind the sheet shows where it is transparent. */
.preview-canvas {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  display: block;
  box-sizing: border-box;
  border: var(--border-width) solid var(--border-default);
  background-color: var(--surface-base);
  background-image:
    linear-gradient(45deg, var(--surface-sunken) 25%, transparent 25%, transparent 75%, var(--surface-sunken) 75%),
    linear-gradient(45deg, var(--surface-sunken) 25%, transparent 25%, transparent 75%, var(--surface-sunken) 75%);
  background-size: 1rem 1rem;
  background-position: 0 0, 0.5rem 0.5rem;
  touch-action: none;
}

.preview-canvas--overlay {
  cursor: move;
}

.preview-canvas:focus-visible,
.placed-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.preview-hint,
.preview-empty {
  margin: 0;
  color: var(--text-secondary);
}

.placed-list {
  margin: 0;
  padding: 0;
  list-style: none;
  max-height: 8rem;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.placed-button {
  font: inherit;
  width: 100%;
  text-align: left;
  padding: var(--space-1) var(--space-2);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
}

.placed-button[aria-pressed="true"] {
  background: var(--surface-selected);
  color: var(--text-accent);
}
</style>
