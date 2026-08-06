<script setup lang="ts">
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
    <button
      v-for="option in WRAP_MODES"
      :key="option.mode"
      type="button"
      class="image-toolbar-option"
      :class="{ applied: option.mode === wrap }"
      :data-testid="`image-wrap-${option.mode}`"
      :aria-pressed="option.mode === wrap ? 'true' : 'false'"
      :title="option.label"
      @mousedown.prevent
      @click="emit('wrap', option.mode)"
    >
      {{ option.label }}
    </button>
    <span class="image-toolbar-separator" />
    <button
      type="button"
      class="image-toolbar-option"
      data-testid="image-options-button"
      @mousedown.prevent
      @click="emit('options')"
    >
      Image options…
    </button>
    <button
      type="button"
      class="image-toolbar-option"
      data-testid="image-delete-button"
      title="Delete image"
      @mousedown.prevent
      @click="emit('delete')"
    >
      Delete
    </button>
  </div>
</template>

<style scoped>
.image-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  align-items: center;
  gap: 0.25rem;
  padding: 0.25rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.image-toolbar-option {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  white-space: nowrap;
}

.image-toolbar-option:hover {
  background: #f1f3f4;
}

.image-toolbar-option.applied {
  background: #e8f0fe;
  border-color: #aecbfa;
  color: #174ea6;
}

.image-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: #dadce0;
}
</style>
