<script setup lang="ts">
import { computed } from "vue";

// Labelled button primitive (SPEC-058). The hover/active/disabled/focus
// treatment lives here once; every consumer picks a variant and a size and
// re-implements none of it. Runtime prop declaration for the espree
// constraint the other SFCs note.
const props = defineProps({
  /** Visual weight. primary = accent fill, secondary = outline, ghost =
   * transparent until hover, danger = destructive accent. */
  variant: { type: String, default: "secondary" },
  /** md is the default 32px control height; sm is a denser 24px chip. */
  size: { type: String, default: "md" },
  /** Native button type — buttons inside dialogs default to "button" so they
   * do not submit a surrounding form by accident. */
  type: { type: String, default: "button" },
  disabled: { type: Boolean, default: false },
});

// The native button type must be the concrete union, not a bare string — the
// ternary keeps vue-tsc happy without a TS `as` cast (barred in .vue scripts).
const nativeType = computed(() =>
  props.type === "submit" ? "submit" : props.type === "reset" ? "reset" : "button",
);
</script>

<template>
  <button
    :type="nativeType"
    class="ui-button"
    :class="[`variant-${props.variant}`, `size-${props.size}`]"
    :disabled="props.disabled"
  >
    <slot />
  </button>
</template>

<style scoped>
.ui-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-2);
  min-height: var(--height-control);
  padding: 0 var(--space-4);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-button);
  font-weight: var(--weight-button);
  line-height: var(--line-tight);
  cursor: pointer;
  transition: background-color var(--motion-fast) var(--easing),
    border-color var(--motion-fast) var(--easing);
}

.ui-button.size-sm {
  min-height: var(--space-6);
  padding: 0 var(--space-3);
  font-size: var(--text-sm);
}

.ui-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.ui-button:disabled {
  cursor: default;
  opacity: 0.5;
}

/* primary — accent fill */
.variant-primary {
  background: var(--accent);
  color: var(--text-on-brand);
  border-color: var(--accent);
}
.variant-primary:not(:disabled):hover {
  background: var(--accent-hover);
  border-color: var(--accent-hover);
}

/* secondary — outlined surface button */
.variant-secondary {
  background: var(--surface-base);
  color: var(--text-primary);
  border-color: var(--border-default);
}
.variant-secondary:not(:disabled):hover {
  background: var(--state-hover);
  border-color: var(--border-strong);
}

/* ghost — transparent until hover */
.variant-ghost {
  background: transparent;
  color: var(--text-primary);
  border-color: transparent;
}
.variant-ghost:not(:disabled):hover {
  background: var(--state-hover);
}
.variant-ghost:not(:disabled):active {
  background: var(--state-active);
}

/* danger — destructive accent fill */
.variant-danger {
  background: var(--danger);
  color: var(--text-on-brand);
  border-color: var(--danger);
}
.variant-danger:not(:disabled):hover {
  background: var(--text-danger);
  border-color: var(--text-danger);
}
</style>
