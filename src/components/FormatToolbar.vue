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
  /** SPEC-068: an optional whitelist of format keys to render, in the fixed
   * FORMATS order. The compact toolbar's essential band shows only bold and
   * italic inline (`['bold','italic']`) and relocates the rest to the overflow;
   * omitted (the default) renders all six. */
  only: { type: Array, default: null },
});

const emit = defineEmits(["toggle"]);

// Static per mount (the modifier key doesn't change mid-session), so this
// runs once rather than as a computed. Shortcut text is read from the
// SPEC-021 registry rather than hardcoded, so it can never drift from what
// the app actually binds (AC-15, AC-16).
// `caption` is the Toolbar Redesign's tiny uppercase word under the icon. It
// is deliberately shorter than `label` — the caption has to fit a 36px control
// without widening the band, while the accessible name stays the full
// "Strikethrough (⌘⇧X)".
const FORMATS = [
  { key: "bold", label: "Bold", caption: "Bold", shortcut: shortcutLabel("format.bold"), icon: "bold" },
  { key: "italic", label: "Italic", caption: "Italic", shortcut: shortcutLabel("format.italic"), icon: "italic" },
  { key: "underline", label: "Underline", caption: "Under", shortcut: shortcutLabel("format.underline"), icon: "underline" },
  { key: "strikethrough", label: "Strikethrough", caption: "Strike", shortcut: shortcutLabel("format.strikethrough"), icon: "strikethrough" },
  // SPEC-013: sub/superscript are toggle buttons like the four above;
  // mutual exclusivity lives in the command layer (one vertAlign field).
  { key: "subscript", label: "Subscript", caption: "Sub", shortcut: shortcutLabel("format.subscript"), icon: "subscript" },
  { key: "superscript", label: "Superscript", caption: "Super", shortcut: shortcutLabel("format.superscript"), icon: "superscript" },
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

// The rendered subset: every format whose key is in `only` (order preserved),
// or all of them when `only` is unset.
function shownFormats() {
  const whitelist = props.only;
  if (!Array.isArray(whitelist)) {
    return FORMATS;
  }
  return FORMATS.filter((format) => whitelist.includes(format.key));
}
</script>

<template>
  <div
    class="format-toolbar"
    role="group"
    aria-label="Character formatting"
  >
    <UiToggleButton
      v-for="format in shownFormats()"
      :key="format.key"
      :icon="format.icon"
      :label="tooltip(format)"
      :caption="format.caption"
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
  gap: 0;
}
</style>
