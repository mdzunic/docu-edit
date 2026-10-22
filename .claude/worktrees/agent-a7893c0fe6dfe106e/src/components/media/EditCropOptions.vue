<script setup lang="ts">
// The crop tool's options bar (SPEC-112): the aspect presets, the straighten
// angle (a native range with a number input beside it), Reset, and the crop
// rectangle announced as "x, y, width × height" in a live region.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { CROP_ASPECT_LABELS, CROP_ASPECTS, MAX_STRAIGHTEN_DEGREES } from "../../lib/core/imageEdits";
import { asImageEditor } from "../../views/media/useImageEditor";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));

function value(event = new Event("input")) {
  const target = event.target;
  return target instanceof HTMLInputElement || target instanceof HTMLSelectElement ? target.value : "";
}

function chooseAspect(event = new Event("change")) {
  const aspect = CROP_ASPECTS.find((entry) => entry === value(event));
  if (aspect) {
    e.value.crop.setAspect(aspect);
  }
}
</script>

<template>
  <div
    class="options-bar"
    role="group"
    aria-label="Crop and straighten options"
  >
    <label class="field">
      <span>Aspect</span>
      <select
        data-testid="crop-aspect"
        :value="e.crop.aspect.value"
        @change="chooseAspect"
      >
        <option
          v-for="aspect in CROP_ASPECTS"
          :key="aspect"
          :value="aspect"
        >
          {{ CROP_ASPECT_LABELS[aspect] }}
        </option>
      </select>
    </label>
    <label class="field">
      <span>Straighten</span>
      <input
        type="range"
        data-testid="edit-straighten"
        aria-label="Straighten angle in degrees"
        :min="-MAX_STRAIGHTEN_DEGREES"
        :max="MAX_STRAIGHTEN_DEGREES"
        step="0.1"
        :value="e.commands.straighten.value"
        @input="e.commands.straightenInput(Number(value($event)))"
        @change="e.commands.release()"
      >
    </label>
    <input
      type="number"
      class="number"
      data-testid="edit-straighten-number"
      aria-label="Straighten angle in degrees"
      :min="-MAX_STRAIGHTEN_DEGREES"
      :max="MAX_STRAIGHTEN_DEGREES"
      step="0.1"
      :value="e.commands.straighten.value"
      @change="e.commands.setStraighten(Number(value($event)))"
    >
    <UiButton
      size="sm"
      data-testid="crop-reset"
      @click="e.crop.reset()"
    >
      Reset crop
    </UiButton>
    <p
      class="readout"
      data-testid="edit-crop-announce"
      role="status"
      aria-label="Crop rectangle"
    >
      {{ e.crop.announcement.value }}
    </p>
  </div>
</template>

<style scoped>
.options-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.number {
  width: 5rem;
  font: inherit;
}

.readout {
  margin: 0;
  margin-left: auto;
  color: var(--text-secondary);
  font-variant-numeric: tabular-nums;
}
</style>
