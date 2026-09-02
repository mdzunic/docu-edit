<script setup lang="ts">
// The audio/video side panel (SPEC-095): start mode (on click vs
// automatically), loop, mute, volume (0–100), full screen, and Replace poster…
// for the selected media frame. All logic lives in the `useSlideMedia`
// controller passed in; this component is presentation + event forwarding only.
// Runtime prop declarations, not defineProps<{...}>() generics (espree lints
// .vue scripts), matching the Transitions/Animations panel contract.
import { ref } from "vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useSlideMedia() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const c = props.controller;
const posterInput = ref<HTMLInputElement | null>(null);

function onStartMode(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setStartMode(el.value === "auto");
}
function onLoop(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setLoop(el.checked);
}
function onMute(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setMuted(el.checked);
}
function onVolume(event = new Event("input")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setVolume(Number(el.value));
}
function onFullScreen(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setFullScreen(el.checked);
}
function openPosterPicker() {
  posterInput.value?.click();
}
function onPosterChange(event = new Event("change")) {
  const el = event.target;
  if (!(el instanceof HTMLInputElement)) return;
  const file = el.files?.[0];
  if (file) void c.replacePoster(file);
  el.value = ""; // allow re-picking the same file
}
</script>

<template>
  <section
    class="media-panel"
    data-testid="slides-media-panel"
    :aria-label="`Media for slide ${slideNumber}`"
  >
    <p
      v-if="!c.isMedia.value"
      class="media-empty"
    >
      Select a video or audio frame to edit its playback.
    </p>
    <div
      v-else
      class="media-controls"
    >
      <label class="media-row">
        <span>Start</span>
        <select
          data-testid="media-start-mode"
          :value="c.startAutomatically.value ? 'auto' : 'click'"
          @change="onStartMode"
        >
          <option value="click">On click</option>
          <option value="auto">Automatically</option>
        </select>
      </label>

      <label class="media-row checkbox">
        <input
          type="checkbox"
          data-testid="media-loop"
          :checked="c.loop.value"
          @change="onLoop"
        >
        <span>Loop</span>
      </label>

      <label class="media-row checkbox">
        <input
          type="checkbox"
          data-testid="media-mute"
          :checked="c.muted.value"
          @change="onMute"
        >
        <span>Mute</span>
      </label>

      <label class="media-row">
        <span>Volume</span>
        <input
          type="range"
          min="0"
          max="100"
          step="1"
          data-testid="media-volume"
          :value="c.volume.value"
          :aria-valuenow="c.volume.value"
          @input="onVolume"
        >
        <span class="media-volume-value">{{ c.volume.value }}</span>
      </label>

      <label class="media-row checkbox">
        <input
          type="checkbox"
          data-testid="media-full-screen"
          :checked="c.fullScreen.value"
          @change="onFullScreen"
        >
        <span>Play full screen</span>
      </label>

      <div class="media-row">
        <UiButton
          data-testid="media-poster-replace"
          @click="openPosterPicker"
        >
          Replace poster…
        </UiButton>
        <input
          ref="posterInput"
          class="sr-only"
          type="file"
          :accept="c.posterAccept"
          data-testid="media-poster-input"
          @change="onPosterChange"
        >
      </div>
    </div>
  </section>
</template>

<style scoped>
.media-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
  padding: var(--space-4);
}
.media-empty {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-md);
}
.media-controls {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}
.media-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-md);
}
.media-row.checkbox {
  gap: var(--space-2);
}
.media-row > span:first-child {
  min-width: 4rem;
}
.media-row select,
.media-row input[type="range"] {
  flex: 1;
}
.media-volume-value {
  min-width: 2ch;
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
