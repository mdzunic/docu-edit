<script setup lang="ts">
// The adjust tool's options bar (SPEC-112): brightness, contrast, saturation
// and exposure as native ranges (−100..100) with number inputs beside them,
// plus grayscale and invert. The preview applies them through
// core/imageAdjust's pixel functions — never a CSS filter. A slider drag is
// one undo entry ("Adjust brightness"); so is a typed number.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { ADJUST_MAX, ADJUST_MIN } from "../../lib/core/imageAdjust";
import { asImageEditor } from "../../views/media/useImageEditor";
import { ADJUST_SLIDERS as SLIDERS, ADJUST_TOGGLES as TOGGLES } from "../../views/media/useEditCommands";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));

function numberOf(event = new Event("input")) {
  const target = event.target;
  const number = target instanceof HTMLInputElement ? Number(target.value) : 0;
  return Number.isFinite(number) ? Math.round(Math.min(ADJUST_MAX, Math.max(ADJUST_MIN, number))) : 0;
}

function checked(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement && target.checked;
}
</script>

<template>
  <div
    class="options-bar"
    role="group"
    aria-label="Adjust options"
  >
    <div
      v-for="slider in SLIDERS"
      :key="slider.channel"
      class="field"
    >
      <label :for="`adjust-${slider.channel}`">{{ slider.label }}</label>
      <input
        :id="`adjust-${slider.channel}`"
        type="range"
        :data-testid="`adjust-${slider.channel}`"
        :min="ADJUST_MIN"
        :max="ADJUST_MAX"
        step="1"
        :value="e.commands.adjustValue(slider.channel)"
        @input="e.commands.adjustInput(slider.channel, numberOf($event))"
        @change="e.commands.release()"
      >
      <input
        type="number"
        class="number"
        :data-testid="`adjust-${slider.channel}-number`"
        :aria-label="slider.label"
        :min="ADJUST_MIN"
        :max="ADJUST_MAX"
        step="1"
        :value="e.commands.adjustValue(slider.channel)"
        @change="e.commands.setAdjust(slider.channel, numberOf($event))"
      >
    </div>
    <label
      v-for="toggle in TOGGLES"
      :key="toggle.channel"
      class="field"
    >
      <input
        type="checkbox"
        :data-testid="`adjust-${toggle.channel}`"
        :checked="e.commands.adjustValue(toggle.channel) === true"
        @change="e.commands.setAdjust(toggle.channel, checked($event))"
      >
      <span>{{ toggle.label }}</span>
    </label>
  </div>
</template>

<style scoped>
.options-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-3) var(--space-4);
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
  width: 4.5rem;
  font: inherit;
}
</style>
