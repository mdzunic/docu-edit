<script setup lang="ts">
import { shortcutLabel } from "../lib/shortcuts";
import UiIconButton from "./ui/UiIconButton.vue";
import UiToggleButton from "./ui/UiToggleButton.vue";

// The text-selection toolbar's controls (SPEC-061 AC-24): the seven things
// people reach for mid-sentence — bold, italic, underline, text colour,
// highlight, link and comment. Positioning, chrome, ARIA and dismissal live on
// the shared ContextualToolbar; this supplies only its own controls. The
// commands are the same ids the strip and menus dispatch (AC-16), so it just
// re-emits. Runtime prop/emit declarations per the espree constraint.
const props = defineProps({
  /** FormatState ("on"/"off"/"mixed") per key, mirroring the strip's toggles. */
  states: { type: Object, required: true },
});

const emit = defineEmits(["toggle", "text-color", "highlight", "link", "comment"]);

const TOGGLES = [
  { key: "bold", label: "Bold", shortcut: shortcutLabel("format.bold"), icon: "bold" },
  { key: "italic", label: "Italic", shortcut: shortcutLabel("format.italic"), icon: "italic" },
  { key: "underline", label: "Underline", shortcut: shortcutLabel("format.underline"), icon: "underline" },
];

function ariaPressed(key = "") {
  const state = props.states[key];
  if (state === "mixed") {
    return "mixed";
  }
  return state === "on";
}
</script>

<template>
  <UiToggleButton
    v-for="entry in TOGGLES"
    :key="entry.key"
    :icon="entry.icon"
    :label="`${entry.label} (${entry.shortcut})`"
    :pressed="ariaPressed(entry.key)"
    :data-testid="`selection-${entry.key}`"
    @mousedown.prevent
    @click="emit('toggle', entry.key)"
  />
  <span class="selection-toolbar-separator" />
  <UiIconButton
    icon="text-color"
    label="Text colour"
    data-testid="selection-text-color"
    @mousedown.prevent
    @click="emit('text-color')"
  />
  <UiIconButton
    icon="highlight"
    label="Highlight"
    data-testid="selection-highlight"
    @mousedown.prevent
    @click="emit('highlight')"
  />
  <span class="selection-toolbar-separator" />
  <UiIconButton
    icon="link"
    :label="`Insert link (${shortcutLabel('insert.link')})`"
    data-testid="selection-link"
    @mousedown.prevent
    @click="emit('link')"
  />
  <UiIconButton
    icon="comment"
    :label="`Add comment (${shortcutLabel('insert.comment')})`"
    data-testid="selection-comment"
    @mousedown.prevent
    @click="emit('comment')"
  />
</template>

<style scoped>
.selection-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--border-default);
}
</style>
