<script setup lang="ts">
import { shortcutLabel } from "../lib/shortcuts";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits<{...}>()
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** FormatState ("on"/"off"/"mixed") per CharacterFormatKey, from the
   * caret/selection the parent is tracking. */
  states: { type: Object, required: true },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["toggle"]);

// Static per mount (the modifier key doesn't change mid-session), so this
// runs once rather than as a computed. Shortcut text is read from the
// SPEC-021 registry rather than hardcoded, so it can never drift from what
// the app actually binds (AC-15, AC-16).
const FORMATS = [
  { key: "bold", label: "Bold", shortcut: shortcutLabel("format.bold"), glyph: "B" },
  { key: "italic", label: "Italic", shortcut: shortcutLabel("format.italic"), glyph: "I" },
  { key: "underline", label: "Underline", shortcut: shortcutLabel("format.underline"), glyph: "U" },
  { key: "strikethrough", label: "Strikethrough", shortcut: shortcutLabel("format.strikethrough"), glyph: "S" },
  // SPEC-013: sub/superscript are toggle buttons like the four above;
  // mutual exclusivity lives in the command layer (one vertAlign field).
  { key: "subscript", label: "Subscript", shortcut: shortcutLabel("format.subscript"), glyph: "X₂" },
  { key: "superscript", label: "Superscript", shortcut: shortcutLabel("format.superscript"), glyph: "X²" },
];

function tooltip(format = FORMATS[0]) {
  return `${format.label} (${format.shortcut})`;
}

/** aria-pressed accepts "mixed" alongside the usual booleans — the ARIA
 * spec's way of expressing a toggle button's indeterminate state
 * (SPEC-012 AC-14). */
function ariaPressed(key = "") {
  const state = props.states[key];
  if (state === "mixed") {
    return "mixed";
  }
  return state === "on" ? "true" : "false";
}

function toggle(key = "") {
  if (!props.disabled) {
    emit("toggle", key);
  }
}
</script>

<template>
  <div
    class="format-toolbar"
    role="group"
    aria-label="Character formatting"
  >
    <button
      v-for="format in FORMATS"
      :key="format.key"
      type="button"
      class="format-button"
      :class="`format-${format.key}`"
      :data-testid="`format-${format.key}`"
      :disabled="disabled"
      :title="tooltip(format)"
      :aria-label="tooltip(format)"
      :aria-pressed="ariaPressed(format.key)"
      @mousedown.prevent
      @click="toggle(format.key)"
    >
      {{ format.glyph }}
    </button>
  </div>
</template>

<style scoped>
.format-toolbar {
  display: inline-flex;
  gap: var(--space-2);
}

.format-button {
  font: inherit;
  font-size: var(--text-md);
  width: 1.75rem;
  height: 1.75rem;
  padding: 0;
  line-height: 1;
  text-align: center;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.format-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.format-button[aria-pressed="true"],
.format-button[aria-pressed="mixed"] {
  background: var(--accent-subtle);
  border-color: var(--accent-subtle);
}

.format-bold {
  font-weight: var(--weight-bold);
}

.format-italic {
  font-style: italic;
}

.format-underline {
  text-decoration: underline;
}

.format-strikethrough {
  text-decoration: line-through;
}
</style>
