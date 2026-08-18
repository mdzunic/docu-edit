<script setup lang="ts">
// The floating container toolbar (SPEC-057): shown near a selected drawing
// group. Fill and outline are absent (fanning them to members is out of
// scope); a container offers wrap, rotation, alt text, Duplicate, Delete and
// Ungroup. A sealed container hides rotation and shows alt text read-only
// (its resize/rotation would rewrite preserved coordinates, and its .odt alt
// text lives inside the preserved subtree). Mirrors ShapeToolbar's own
// pattern and positioning. Runtime prop/emit declarations (espree).
import { ref, watch } from "vue";

const props = defineProps({
  left: { type: Number, default: 0 },
  top: { type: Number, default: 0 },
  wrap: { type: String, default: "inline" },
  rotationDeg: { type: Number, default: 0 },
  altText: { type: String, default: "" },
  sealed: { type: Boolean, default: false },
  placement: { type: String, default: "below" },
});

const emit = defineEmits(["wrap", "rotation", "alt-text", "duplicate", "delete", "ungroup"]);

const WRAP_MODES = [
  { mode: "inline", label: "In line with text" },
  { mode: "square", label: "Square" },
  { mode: "topAndBottom", label: "Top and bottom" },
  { mode: "behind", label: "Behind text" },
  { mode: "infront", label: "In front of text" },
];

const rotationInput = ref(String(props.rotationDeg));
watch(
  () => props.rotationDeg,
  (deg) => {
    rotationInput.value = String(deg);
  },
);

function toolbarStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}

function onWrapChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("wrap", target.value);
  }
}

function onRotationChange() {
  const deg = Number(rotationInput.value);
  if (Number.isFinite(deg)) {
    emit("rotation", deg);
  }
}

function onAltTextInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("alt-text", target.value);
  }
}
</script>

<template>
  <div
    class="shape-toolbar"
    :class="{ 'shape-toolbar-above': placement === 'above' }"
    data-testid="group-toolbar"
    role="toolbar"
    aria-label="Group formatting"
    :style="toolbarStyle()"
  >
    <select
      class="shape-toolbar-select"
      data-testid="group-wrap"
      :value="wrap"
      @change="onWrapChange"
      @mousedown.stop
    >
      <option
        v-for="option in WRAP_MODES"
        :key="option.mode"
        :value="option.mode"
      >
        {{ option.label }}
      </option>
    </select>
    <label
      v-if="!sealed"
      class="shape-toolbar-field"
      title="Rotation (degrees)"
    >
      <input
        v-model="rotationInput"
        type="number"
        min="0"
        max="359"
        class="shape-toolbar-number"
        data-testid="group-rotation"
        @change="onRotationChange"
        @mousedown.stop
      >
      °
    </label>
    <span class="shape-toolbar-separator" />
    <label
      class="shape-toolbar-field"
      title="Alt text"
    >
      Alt text
      <input
        type="text"
        class="shape-toolbar-alt"
        data-testid="group-alt-text"
        :value="altText"
        :readonly="sealed"
        @input="onAltTextInput"
        @mousedown.stop
      >
    </label>
    <span class="shape-toolbar-separator" />
    <button
      type="button"
      class="shape-toolbar-option"
      data-testid="group-ungroup-button"
      @mousedown.prevent
      @click="emit('ungroup')"
    >
      Ungroup
    </button>
    <button
      type="button"
      class="shape-toolbar-option"
      data-testid="group-duplicate-button"
      title="Duplicate"
      @mousedown.prevent
      @click="emit('duplicate')"
    >
      Duplicate
    </button>
    <button
      type="button"
      class="shape-toolbar-option"
      data-testid="group-delete-button"
      title="Delete group"
      @mousedown.prevent
      @click="emit('delete')"
    >
      Delete
    </button>
  </div>
</template>

<style scoped>
/* Reuses ShapeToolbar's visual language; the shared classes below are scoped
   copies (Vue scopes per-component), matching ShapeToolbar.vue's own rules. */
.shape-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0.25rem;
  padding: 0.25rem;
  max-width: 640px;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.shape-toolbar-above {
  transform: translateY(-100%);
}

.shape-toolbar-field {
  display: flex;
  align-items: center;
  gap: 0.25rem;
  font-size: 0.75rem;
  color: #202124;
  white-space: nowrap;
}

.shape-toolbar-number {
  width: 3.5em;
}

.shape-toolbar-alt {
  width: 8em;
}

.shape-toolbar-select {
  font-size: 0.8125rem;
}

.shape-toolbar-option {
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

.shape-toolbar-option:hover {
  background: #f1f3f4;
}

.shape-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: #dadce0;
}
</style>
