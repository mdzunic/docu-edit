<script setup lang="ts">
import AppIcon from "../AppIcon.vue";
import UiControlCaption from "./UiControlCaption.vue";

// Icon-only button primitive (SPEC-058). `label` is required and becomes both
// the aria-label and the native title tooltip, so an icon-only control is
// never unlabelled. The AppIcon inside is aria-hidden, so the button's own
// name is the only thing announced. 32×32 minimum hit target (AC-48).
//
// Toolbar Redesign: `caption` prints the tiny uppercase word under the icon in
// the banded strip, hidden through the inherited --toolbar-caption-display
// property so the strip can drop every caption at once. It is decorative —
// aria-hidden — because `label` is already the button's accessible name.
const props = defineProps({
  /** Icon id passed straight to AppIcon. */
  icon: { type: String, required: true },
  /** Accessible name + tooltip. Required — an icon-only button must be named. */
  label: { type: String, required: true },
  /** The toolbar strip's tiny uppercase caption. Shorter than `label`. */
  caption: { type: String, default: "" },
  disabled: { type: Boolean, default: false },
  /** In-flight state: dims the control, marks it aria-busy and blocks a second
   * activation while the action it triggered runs (SPEC-060's close flush). */
  busy: { type: Boolean, default: false },
});
</script>

<template>
  <button
    type="button"
    class="ui-icon-button"
    :class="{ 'is-busy': props.busy }"
    :aria-label="props.label"
    :title="props.label"
    :aria-busy="props.busy || undefined"
    :disabled="props.disabled || props.busy"
  >
    <AppIcon :name="props.icon" />
    <UiControlCaption
      v-if="props.caption"
      :text="props.caption"
    />
  </button>
</template>

<style scoped>
.ui-icon-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: var(--control-min-width, var(--height-control));
  height: var(--control-height, var(--height-control));
  padding: 0;
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
  transition: background-color var(--motion-fast) var(--easing),
    color var(--motion-fast) var(--easing);
}

/* Captioned: the icon sits over its word, exactly as in UiToggleButton. The
   side padding arrives only WITH the caption — an uncaptioned icon button is
   still exactly --height-control square, which is what every icon button
   outside the toolbar strip depends on. */
.ui-icon-button:has(> .ui-control-caption) {
  flex-direction: column;
  gap: var(--space-1);
  padding: 0 var(--space-2);
}


.ui-icon-button:not(:disabled):hover {
  background: var(--state-hover);
}

.ui-icon-button:not(:disabled):active {
  background: var(--state-active);
}

.ui-icon-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.ui-icon-button:disabled {
  cursor: default;
  opacity: 0.5;
}

/* The busy state pulses so the disabled control still reads as "working",
   not merely inert; stilled under reduced-motion (--motion-base is 0 there). */
.ui-icon-button.is-busy {
  cursor: progress;
  animation: ui-icon-busy-pulse calc(4 * var(--motion-base)) var(--easing) infinite;
}

@keyframes ui-icon-busy-pulse {
  0%,
  100% {
    opacity: 0.5;
  }
  50% {
    opacity: 0.85;
  }
}
</style>
