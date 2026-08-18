<script setup lang="ts">
import AppIcon from "../AppIcon.vue";

// Icon-only button primitive (SPEC-058). `label` is required and becomes both
// the aria-label and the native title tooltip, so an icon-only control is
// never unlabelled. The AppIcon inside is aria-hidden, so the button's own
// name is the only thing announced. 32×32 minimum hit target (AC-48).
const props = defineProps({
  /** Icon id passed straight to AppIcon. */
  icon: { type: String, required: true },
  /** Accessible name + tooltip. Required — an icon-only button must be named. */
  label: { type: String, required: true },
  disabled: { type: Boolean, default: false },
});
</script>

<template>
  <button
    type="button"
    class="ui-icon-button"
    :aria-label="props.label"
    :title="props.label"
    :disabled="props.disabled"
  >
    <AppIcon :name="props.icon" />
  </button>
</template>

<style scoped>
.ui-icon-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: var(--height-control);
  height: var(--height-control);
  padding: 0;
  color: var(--text-secondary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
  transition: background-color var(--motion-fast) var(--easing),
    color var(--motion-fast) var(--easing);
}

.ui-icon-button:not(:disabled):hover {
  background: var(--state-hover);
  color: var(--text-primary);
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
</style>
