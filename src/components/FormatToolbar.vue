<script setup lang="ts">
import { shortcutLabel } from "../lib/shortcuts";
import UiToggleButton from "./ui/UiToggleButton.vue";

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
  { key: "bold", label: "Bold", shortcut: shortcutLabel("format.bold"), icon: "bold" },
  { key: "italic", label: "Italic", shortcut: shortcutLabel("format.italic"), icon: "italic" },
  { key: "underline", label: "Underline", shortcut: shortcutLabel("format.underline"), icon: "underline" },
  { key: "strikethrough", label: "Strikethrough", shortcut: shortcutLabel("format.strikethrough"), icon: "strikethrough" },
  // SPEC-013: sub/superscript are toggle buttons like the four above;
  // mutual exclusivity lives in the command layer (one vertAlign field).
  { key: "subscript", label: "Subscript", shortcut: shortcutLabel("format.subscript"), icon: "subscript" },
  { key: "superscript", label: "Superscript", shortcut: shortcutLabel("format.superscript"), icon: "superscript" },
];

function tooltip(format = FORMATS[0]) {
  return `${format.label} (${format.shortcut})`;
}

/** The toggle's pressed value: true/false or the ARIA tri-state "mixed" — the
 * ARIA spec's way of expressing a toggle button's indeterminate state
 * (SPEC-012 AC-14). UiToggleButton renders it straight into aria-pressed. */
function ariaPressed(key = "") {
  const state = props.states[key];
  if (state === "mixed") {
    return "mixed";
  }
  return state === "on";
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
    <UiToggleButton
      v-for="format in FORMATS"
      :key="format.key"
      :icon="format.icon"
      :label="tooltip(format)"
      :pressed="ariaPressed(format.key)"
      :disabled="disabled"
      :data-testid="`format-${format.key}`"
      @mousedown.prevent
      @click="toggle(format.key)"
    />
  </div>
</template>

<style scoped>
.format-toolbar {
  display: inline-flex;
  gap: var(--space-2);
}
</style>
