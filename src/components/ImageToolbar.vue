<script setup lang="ts">
import UiButton from "./ui/UiButton.vue";
import UiToggleButton from "./ui/UiToggleButton.vue";
// The floating image toolbar (SPEC-024 AC-9, AC-13): shown near a selected
// image, offering the four wrap modes plus "Image options…". Positioned the
// same way PasteOptionsChip is — a measured content rect turned into fixed
// left/top, not useFloatingMenu (that composable is for button-anchored
// dropdowns, see useFloatingMenu.ts). Runtime prop/emit declarations per the
// espree constraint (see AppShell.vue).
const props = defineProps({
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  wrap: { type: String, required: true },
});

const emit = defineEmits(["wrap", "options", "delete"]);

const WRAP_MODES = [
  { mode: "inline", label: "In line with text" },
  { mode: "square", label: "Wrap text" },
  { mode: "behind", label: "Behind text" },
  { mode: "infront", label: "In front of text" },
];

function toolbarStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}
</script>

<template>
  <div
    class="image-toolbar"
    data-testid="image-toolbar"
    role="toolbar"
    aria-label="Image formatting"
    :style="toolbarStyle()"
  >
    <UiToggleButton
      v-for="option in WRAP_MODES"
      :key="option.mode"
      :label="option.label"
      :pressed="option.mode === wrap"
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
  </div>
</template>

<style scoped>
.image-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.image-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--border-default);
}
</style>
