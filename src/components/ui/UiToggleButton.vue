<script setup lang="ts">
import AppIcon from "../AppIcon.vue";

// Two-state toggle primitive (SPEC-058) — bold/italic/underline and every
// other sticky formatting control. Renders aria-pressed from `pressed`, and
// the selected treatment (the SPEC's --surface-selected well) is defined here
// once so no consumer re-implements a "this is on" colour. An icon or a text
// label may be supplied; a label falls back to the default slot.
const props = defineProps({
  /** Optional icon id. When set, an AppIcon renders before any slot content. */
  icon: { type: String, default: "" },
  /** Accessible name + tooltip; required whenever the button is icon-only. */
  label: { type: String, default: "" },
  /** The on/off state — rendered as aria-pressed and the selected styling.
   * Accepts the ARIA tri-state "mixed" (an indeterminate selection) as well as
   * a boolean; "mixed" reads as pressed for styling. */
  pressed: { type: [Boolean, String], default: false },
  disabled: { type: Boolean, default: false },
});

</script>

<template>
  <button
    type="button"
    class="ui-toggle-button"
    :class="{ 'is-pressed': props.pressed === true || props.pressed === 'mixed' }"
    :aria-pressed="props.pressed"
    :aria-label="props.label || undefined"
    :title="props.label || undefined"
    :disabled="props.disabled"
  >
    <AppIcon
      v-if="props.icon"
      :name="props.icon"
    />
    <slot />
  </button>
</template>

<style scoped>
.ui-toggle-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-2);
  min-width: var(--height-control);
  height: var(--height-control);
  padding: 0 var(--space-2);
  color: var(--text-secondary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-md);
  font-weight: var(--weight-medium);
  line-height: var(--line-tight);
  cursor: pointer;
  transition: background-color var(--motion-fast) var(--easing),
    color var(--motion-fast) var(--easing);
}

.ui-toggle-button:not(:disabled):hover {
  background: var(--state-hover);
  color: var(--text-primary);
}

.ui-toggle-button:not(:disabled):active {
  background: var(--state-active);
}

/* Selected/pressed state — defined once, applied by every consumer. */
.ui-toggle-button.is-pressed {
  background: var(--surface-selected);
  color: var(--text-accent);
}

.ui-toggle-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.ui-toggle-button:disabled {
  cursor: default;
  opacity: 0.5;
}
</style>
