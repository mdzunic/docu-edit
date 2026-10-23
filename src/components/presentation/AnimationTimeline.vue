<script setup lang="ts">
// The animation timeline (SPEC-105): a second view of the Animations panel's
// list, not a replacement — a ruler at 60px per second with one tick per
// second, one band per click group (the main sequence's, then each
// interactive sequence's) and one bar per build step, offset by its start and
// sized by its playback duration so with-previous overlaps are visible.
// Presentation only: the geometry is components/presentation/animationTimeline.ts
// (via the controller's `timeline`), and every edit — selecting a bar, a
// left-edge (delay) or right-edge (duration) drag recorded once on pointer-up,
// a Left/Right arrow step of 100ms — goes through the useAnimations controller.
// Runtime prop declarations, not defineProps<{...}>() generics (espree).
const props = defineProps({
  /** The useAnimations() controller object (timeline + edge commands). */
  controller: { type: Object, required: true },
});

const c = props.controller;

function barStyle(bar = { left: 0, width: 0 }) {
  return { left: `${bar.left}px`, width: `${bar.width}px` };
}
function laneStyle(widthPx = 0) {
  return { width: `${widthPx}px` };
}
function seconds(ms = 0) {
  return `${Number((ms / 1000).toFixed(2))}s`;
}
function groupLabel(group = { index: 0, triggerFrameId: "" }) {
  return group.triggerFrameId ? `On click of ${c.frameName(group.triggerFrameId)}` : `Click ${group.index + 1}`;
}
function barLabel(bar = { animationIndex: 0, paragraphIndex: undefined, startMs: 0, durationMs: 0 }) {
  const paragraph = bar.paragraphIndex === undefined ? "" : `, paragraph ${bar.paragraphIndex + 1}`;
  return `Animation ${bar.animationIndex + 1}${paragraph}: starts ${seconds(bar.startMs)}, runs ${seconds(bar.durationMs)}`;
}
</script>

<template>
  <div
    v-if="c.timeline.value"
    class="timeline"
    role="group"
    aria-label="Animation timeline"
    data-testid="animation-timeline"
  >
    <div class="lane-scroll">
      <div
        class="lanes"
        :style="laneStyle(c.timeline.value.widthPx)"
      >
        <div
          class="ruler"
          aria-hidden="true"
          data-testid="animation-timeline-ruler"
        >
          <span
            v-for="tick in c.timeline.value.ticks"
            :key="tick.seconds"
            class="tick"
            :style="{ left: `${tick.x}px` }"
            :data-testid="`animation-timeline-tick-${tick.seconds}`"
          >{{ tick.label }}</span>
        </div>
        <div
          v-for="group in c.timeline.value.groups"
          :key="group.index"
          class="group"
          :class="{ 'is-interactive': !!group.triggerFrameId }"
          role="group"
          :aria-label="groupLabel(group)"
          :data-testid="`animation-timeline-group-${group.index}`"
        >
          <span class="group-label">{{ groupLabel(group) }}</span>
          <div
            v-for="bar in group.bars"
            :key="bar.barIndex"
            class="bar-row"
          >
            <div
              class="bar"
              :class="{ 'is-selected': bar.animationIndex === c.selectedIndex.value }"
              :style="barStyle(bar)"
              :data-testid="`animation-timeline-bar-${bar.barIndex}`"
              :data-animation-index="bar.animationIndex"
            >
              <span
                class="edge edge-start"
                role="slider"
                tabindex="0"
                aria-label="Delay"
                aria-valuemin="0"
                aria-valuemax="10000"
                :aria-valuenow="bar.delayMs"
                :aria-valuetext="seconds(bar.delayMs)"
                :data-testid="`animation-timeline-delay-${bar.barIndex}`"
                @pointerdown.stop.prevent="c.startEdgeDrag(bar, 'delay', $event)"
                @pointermove="c.moveEdgeDrag($event)"
                @pointerup="c.endEdgeDrag()"
                @pointercancel="c.endEdgeDrag()"
                @keydown="c.edgeKeydown(bar, 'delay', $event)"
              />
              <button
                type="button"
                class="bar-body"
                :aria-label="barLabel(bar)"
                :aria-pressed="bar.animationIndex === c.selectedIndex.value ? 'true' : 'false'"
                @click="c.selectBar(bar)"
              />
              <span
                class="edge edge-end"
                role="slider"
                tabindex="0"
                aria-label="Duration"
                aria-valuemin="100"
                aria-valuemax="5000"
                :aria-valuenow="bar.singleMs"
                :aria-valuetext="seconds(bar.singleMs)"
                :data-testid="`animation-timeline-duration-${bar.barIndex}`"
                @pointerdown.stop.prevent="c.startEdgeDrag(bar, 'duration', $event)"
                @pointermove="c.moveEdgeDrag($event)"
                @pointerup="c.endEdgeDrag()"
                @pointercancel="c.endEdgeDrag()"
                @keydown="c.edgeKeydown(bar, 'duration', $event)"
              />
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.timeline {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  max-height: 14rem;
  overflow-y: auto;
}

.lane-scroll {
  overflow-x: auto;
  padding: var(--space-2);
}

.lanes {
  position: relative;
  min-width: 100%;
}

.ruler {
  position: relative;
  height: 1.25rem;
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-secondary);
  font-size: var(--text-xs);
  font-variant-numeric: tabular-nums;
}

.tick {
  position: absolute;
  top: 0;
  border-left: var(--border-width) solid var(--border-subtle);
  padding-left: var(--space-1);
  white-space: nowrap;
}

.group {
  position: relative;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  padding: var(--space-1) 0;
  border-bottom: var(--border-width) dashed var(--border-subtle);
}

.group.is-interactive {
  background: var(--surface-default);
}

.group-label {
  color: var(--text-secondary);
  font-size: var(--text-xs);
}

.bar-row {
  position: relative;
  height: 1rem;
}

.bar {
  position: absolute;
  top: 0;
  height: 100%;
  display: flex;
  border-radius: var(--radius-sm);
  background: var(--accent-soft, var(--surface-default));
  border: var(--border-width) solid var(--accent);
}

.bar.is-selected {
  background: var(--accent);
}

.bar-body {
  flex: 1 1 auto;
  min-width: 0;
  border: none;
  padding: 0;
  background: transparent;
  cursor: pointer;
}

.edge {
  flex: 0 0 auto;
  width: 0.375rem;
  cursor: ew-resize;
  background: var(--accent);
  touch-action: none;
}
</style>
