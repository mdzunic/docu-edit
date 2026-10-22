<script setup lang="ts">
import UiButton from "./ui/UiButton.vue";
import UiToggleButton from "./ui/UiToggleButton.vue";
// The image toolbar's controls (SPEC-024 AC-9, AC-13): the four wrap modes
// plus "Image options…" and Delete. Positioning, chrome, ARIA and dismissal
// now live on the shared ContextualToolbar (SPEC-061 AC-22, AC-23); this
// component supplies only its own controls. Runtime prop/emit declarations per
// the espree constraint (see AppShell.vue).
const props = defineProps({
  wrap: { type: String, required: true },
});

const emit = defineEmits(["wrap", "options", "delete"]);

const WRAP_MODES = [
  { mode: "inline", label: "In line with text" },
  { mode: "square", label: "Wrap text" },
  { mode: "behind", label: "Behind text" },
  { mode: "infront", label: "In front of text" },
];
</script>

<template>
  <UiToggleButton
    v-for="option in WRAP_MODES"
    :key="option.mode"
    :label="option.label"
    :pressed="option.mode === props.wrap"
    :data-testid="`image-wrap-${option.mode}`"
    @mousedown.prevent
    @click="emit('wrap', option.mode)"
  >
    {{ option.label }}
  </UiToggleButton>
  <span class="image-toolbar-separator" />
  <UiButton
    variant="ghost"
    data-testid="image-options-button"
    @mousedown.prevent
    @click="emit('options')"
  >
    Image options…
  </UiButton>
  <UiButton
    variant="ghost"
    data-testid="image-delete-button"
    title="Delete image"
    @mousedown.prevent
    @click="emit('delete')"
  >
    Delete
  </UiButton>
</template>

<style scoped>
.image-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--border-default);
}
</style>
