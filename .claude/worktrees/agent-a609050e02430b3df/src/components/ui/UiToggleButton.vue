<script setup lang="ts">
import { computed } from "vue";
import AppIcon from "../AppIcon.vue";
import UiControlCaption from "./UiControlCaption.vue";

// Two-state toggle primitive (SPEC-058) — bold/italic/underline and every
// other sticky formatting control. Renders aria-pressed from `pressed`, and
// the selected treatment (the SPEC's --surface-selected well) is defined here
// once so no consumer re-implements a "this is on" colour. An icon or a text
// label may be supplied; a label falls back to the default slot.
//
// Toolbar Redesign: a `caption` is the tiny uppercase word printed UNDER the
// icon in the banded strip. The caption is always rendered when supplied and
// hidden through the inherited --toolbar-caption-display custom property, so
// the strip can drop every caption at once (and measure both geometries)
// without any consumer re-rendering. Size likewise rides two inherited
// properties rather than a prop threaded through six intermediate components.
const props = defineProps({
  /** Optional icon id. When set, an AppIcon renders before any slot content. */
  icon: { type: String, default: "" },
  /** Accessible name + tooltip; required whenever the button is icon-only. */
  label: { type: String, default: "" },
  /** The tiny uppercase caption under the icon in the toolbar strip. Purely
   * decorative — `label` remains the accessible name — so it can be a shorter
   * word than the label ("Strike" under a "Strikethrough (⌘⇧X)" button). */
  caption: { type: String, default: "" },
  /** The on/off state — rendered as aria-pressed and the selected styling.
   * Accepts the ARIA tri-state "mixed" (an indeterminate selection) as well as
   * a boolean; "mixed" gets its own dashed treatment. */
  pressed: { type: [Boolean, String], default: false },
  disabled: { type: Boolean, default: false },
});

// Normalise the tri-state into the exact aria-pressed value: the string
// "mixed", or a real boolean (never a bare string like "true").
const ariaPressed = computed(() =>
  props.pressed === "mixed" ? "mixed" : props.pressed === true,
);
// "on" and "mixed" are now styled apart (the redesign gives an indeterminate
// selection a dashed accent outline instead of the filled well), so they are
// two classes rather than one is-pressed.
const isPressed = computed(() => props.pressed === true);
const isMixed = computed(() => props.pressed === "mixed");
</script>

<template>
  <button
    type="button"
    class="ui-toggle-button"
    :class="{ 'is-pressed': isPressed, 'is-mixed': isMixed }"
    :aria-pressed="ariaPressed"
    :aria-label="props.label || undefined"
    :title="props.label || undefined"
    :disabled="props.disabled"
  >
    <AppIcon
      v-if="props.icon"
      :name="props.icon"
    />
    <slot />
    <UiControlCaption
      v-if="props.caption"
      :text="props.caption"
    />
  </button>
</template>

<style scoped>
.ui-toggle-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-2);
  min-width: var(--control-min-width, var(--height-control));
  height: var(--control-height, var(--height-control));
  padding: 0 var(--space-2);
  /* Full-strength ink: the handoff's control is `color: var(--tx)`, and
     SPEC-059 pins the sibling paragraph buttons to --text-primary — two inks
     in one band is exactly the inconsistency the redesign set out to remove. */
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-md);
  font-weight: var(--weight-medium);
  line-height: var(--line-tight);
  cursor: pointer;
  transition: background-color var(--motion-fast) var(--easing),
    color var(--motion-fast) var(--easing),
    border-color var(--motion-fast) var(--easing);
}

/* With a caption the control stacks icon over word; the gap tightens because
   the two are one object, not two controls side by side. */
.ui-toggle-button:has(> .ui-control-caption) {
  flex-direction: column;
  gap: var(--space-1);
}


.ui-toggle-button:not(:disabled):hover {
  background: var(--state-hover);
}

.ui-toggle-button:not(:disabled):active {
  background: var(--state-active);
}

/* Selected/pressed state — defined once, applied by every consumer. The
   redesign adds the accent outline so a pressed control reads as pressed
   without relying on the fill alone (which is faint in several palettes). */
.ui-toggle-button.is-pressed {
  background: var(--surface-selected);
  color: var(--text-accent);
  border-color: var(--accent);
}

/* Indeterminate ("mixed") — the selection has the format in some runs and not
   others. A dashed accent outline with no fill: visibly not the same thing as
   pressed, and legible without colour alone carrying the meaning. */
.ui-toggle-button.is-mixed {
  color: var(--text-accent);
  border-color: var(--accent);
  border-style: dashed;
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
