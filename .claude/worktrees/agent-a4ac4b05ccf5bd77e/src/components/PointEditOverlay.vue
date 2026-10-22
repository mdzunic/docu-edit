<script setup lang="ts">
// SPEC-071 point-edit mode overlay: paints the live edited outline over the
// shape and lays a handle on every vertex (a square for a corner, a circle
// for a smooth one), plus control-point dots on leader lines for the
// selected vertex and its neighbours. It reports handle drags/toggles and
// segment clicks; useShapes' `pointEdit` controller owns the model edits.
// The whole point-edit API arrives as one `controller` object so
// DocumentView's template wiring stays a single binding (its size budget).
// Runtime prop declaration per the .vue espree constraint.
defineProps({ controller: { type: Object, required: true } });
</script>

<template>
  <div
    class="point-edit-overlay"
    data-testid="point-edit-overlay"
  >
    <svg
      class="point-edit-outline"
      :style="{
        left: `${controller.frame.value.left}px`,
        top: `${controller.frame.value.top}px`,
        width: `${controller.frame.value.width}px`,
        height: `${controller.frame.value.height}px`,
        transform: `rotate(${controller.frame.value.rotationDeg}deg)`,
      }"
      :viewBox="`0 0 ${controller.frame.value.width} ${controller.frame.value.height}`"
      @mousedown="controller.onSegmentDown($event)"
    >
      <path :d="controller.frame.value.d" />
    </svg>
    <div
      v-for="ctrl in controller.controls()"
      :key="`line-${ctrl.index}-${ctrl.which}`"
      class="point-edit-leader"
      :style="ctrl.line"
    />
    <button
      v-for="ctrl in controller.controls()"
      :key="`ctrl-${ctrl.index}-${ctrl.which}`"
      type="button"
      class="point-edit-control"
      :data-testid="`point-edit-control-${ctrl.index}-${ctrl.which}`"
      :style="ctrl.dot"
      aria-label="Control point"
      @mousedown="controller.onControlDown($event, ctrl.index, ctrl.which)"
    />
    <button
      v-for="vertex in controller.vertices()"
      :key="`vertex-${vertex.index}`"
      type="button"
      class="point-edit-vertex"
      :class="{ 'is-smooth': vertex.smooth, 'is-selected': vertex.selected }"
      :data-testid="`point-edit-vertex-${vertex.index}`"
      :style="vertex.style"
      :aria-label="`Vertex ${vertex.index + 1}`"
      :aria-pressed="vertex.selected"
      @mousedown="controller.onVertexDown($event, vertex.index)"
      @dblclick="controller.onVertexToggle(vertex.index)"
    />
  </div>
</template>

<style scoped>
.point-edit-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 6;
}

.point-edit-outline {
  position: absolute;
  overflow: visible;
  pointer-events: none;
  transform-origin: center;
}

.point-edit-outline path {
  fill: none;
  stroke: var(--accent);
  stroke-width: 1;
  pointer-events: stroke;
  cursor: copy;
}

.point-edit-leader {
  position: absolute;
  height: 0;
  border-top: var(--border-width) solid var(--accent);
  transform-origin: left center;
  pointer-events: none;
  opacity: 0.6;
}

.point-edit-vertex,
.point-edit-control {
  position: absolute;
  padding: 0;
  background: var(--surface-base);
  border: var(--border-width) solid var(--accent);
  cursor: grab;
  pointer-events: auto;
}

.point-edit-vertex {
  width: 8px;
  height: 8px;
}

.point-edit-vertex.is-smooth {
  border-radius: 50%;
}

.point-edit-vertex.is-selected {
  background: var(--accent);
}

.point-edit-control {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: var(--accent);
}
</style>
