<script setup lang="ts">
// The floating shape toolbar (SPEC-051 AC-12..16, AC-25, AC-29..32, AC-34,
// AC-39, AC-44): shown near a selected shape, offering fill, outline,
// rotation, wrap mode, alt text, z-order and group/ungroup. Mirrors
// ImageToolbar.vue's own pattern and positioning convention. Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { ref, watch } from "vue";
import UiButton from "./ui/UiButton.vue";

const props = defineProps({
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  fill: { type: String, default: "" },
  outlineColor: { type: String, default: "" },
  outlineWidthPt: { type: Number, default: 1 },
  outlineDash: { type: String, default: "solid" },
  rotationDeg: { type: Number, default: 0 },
  wrap: { type: String, required: true },
  altText: { type: String, default: "" },
  grouped: { type: Boolean, default: false },
  canGroup: { type: Boolean, default: false },
  /** "above" anchors `top` at the toolbar's BOTTOM edge (it lifts itself by
   * its own height in CSS), "below" at its top edge as usual — see
   * useShapes.ts's shapeToolbarStyle for why the caller cannot resolve this
   * to a plain `top` itself. */
  placement: { type: String, default: "below" },
});

const emit = defineEmits([
  "fill",
  "outline",
  "rotation",
  "wrap",
  "alt-text",
  "z-order",
  "group",
  "ungroup",
  "delete",
  "duplicate",
]);

const WRAP_MODES = [
  { mode: "inline", label: "In line with text" },
  { mode: "square", label: "Square" },
  { mode: "topAndBottom", label: "Top and bottom" },
  { mode: "behind", label: "Behind text" },
  { mode: "infront", label: "In front of text" },
];

const Z_ORDER_COMMANDS = [
  { command: "front", label: "Bring to front" },
  { command: "forward", label: "Bring forward" },
  { command: "backward", label: "Send backward" },
  { command: "back", label: "Send to back" },
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

function onFillInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("fill", target.value);
  }
}

function onFillNone() {
  emit("fill", undefined);
}

function onOutlineColorInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("outline", { color: target.value, widthPt: props.outlineWidthPt, dash: props.outlineDash });
  }
}

function onOutlineWidthInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    const widthPt = Number(target.value) || 0;
    emit("outline", { color: props.outlineColor || "#000000", widthPt, dash: props.outlineDash });
  }
}

function onOutlineDashChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("outline", { color: props.outlineColor || "#000000", widthPt: props.outlineWidthPt, dash: target.value });
  }
}

function onOutlineNone() {
  emit("outline", undefined);
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

function onWrapChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("wrap", target.value);
  }
}

function onZOrderChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("z-order", target.value);
    target.selectedIndex = -1;
  }
}
</script>

<template>
  <div
    class="shape-toolbar"
    :class="{ 'shape-toolbar-above': placement === 'above' }"
    data-testid="shape-toolbar"
    role="toolbar"
    aria-label="Shape formatting"
    :style="toolbarStyle()"
  >
    <label
      class="shape-toolbar-field"
      title="Fill colour"
    >
      Fill
      <input
        type="color"
        data-testid="shape-fill-color"
        :value="fill || '#ffffff'"
        @input="onFillInput"
        @mousedown.stop
      >
    </label>
    <UiButton
      variant="ghost"
      data-testid="shape-fill-none"
      @mousedown.prevent
      @click="onFillNone"
    >
      No fill
    </UiButton>
    <span class="shape-toolbar-separator" />
    <label
      class="shape-toolbar-field"
      title="Outline colour"
    >
      Outline
      <input
        type="color"
        data-testid="shape-outline-color"
        :value="outlineColor || '#000000'"
        @input="onOutlineColorInput"
        @mousedown.stop
      >
    </label>
    <label
      class="shape-toolbar-field"
      title="Outline width (pt)"
    >
      <input
        type="number"
        min="0"
        step="0.25"
        class="shape-toolbar-number"
        data-testid="shape-outline-width"
        :value="outlineWidthPt"
        @input="onOutlineWidthInput"
        @mousedown.stop
      >
      pt
    </label>
    <select
      class="shape-toolbar-select"
      data-testid="shape-outline-dash"
      :value="outlineDash"
      @change="onOutlineDashChange"
      @mousedown.stop
    >
      <option value="solid">
        Solid
      </option>
      <option value="dash">
        Dashed
      </option>
      <option value="dot">
        Dotted
      </option>
    </select>
    <UiButton
      variant="ghost"
      data-testid="shape-outline-none"
      @mousedown.prevent
      @click="onOutlineNone"
    >
      No outline
    </UiButton>
    <span class="shape-toolbar-separator" />
    <label
      class="shape-toolbar-field"
      title="Rotation (degrees)"
    >
      <input
        v-model="rotationInput"
        type="number"
        min="0"
        max="359"
        class="shape-toolbar-number"
        data-testid="shape-rotation"
        @change="onRotationChange"
        @mousedown.stop
      >
      °
    </label>
    <span class="shape-toolbar-separator" />
    <select
      class="shape-toolbar-select"
      data-testid="shape-wrap"
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
    <select
      class="shape-toolbar-select"
      data-testid="shape-zorder"
      @change="onZOrderChange"
      @mousedown.stop
    >
      <option
        value=""
        disabled
        selected
      >
        Order…
      </option>
      <option
        v-for="option in Z_ORDER_COMMANDS"
        :key="option.command"
        :value="option.command"
      >
        {{ option.label }}
      </option>
    </select>
    <UiButton
      v-if="canGroup"
      variant="ghost"
      data-testid="shape-group-button"
      @mousedown.prevent
      @click="emit('group')"
    >
      Group
    </UiButton>
    <UiButton
      v-if="grouped"
      variant="ghost"
      data-testid="shape-ungroup-button"
      @mousedown.prevent
      @click="emit('ungroup')"
    >
      Ungroup
    </UiButton>
    <span class="shape-toolbar-separator" />
    <label
      class="shape-toolbar-field"
      title="Alt text"
    >
      Alt text
      <input
        type="text"
        class="shape-toolbar-alt"
        data-testid="shape-alt-text"
        :value="altText"
        @input="onAltTextInput"
        @mousedown.stop
      >
    </label>
    <span class="shape-toolbar-separator" />
    <UiButton
      variant="ghost"
      data-testid="shape-duplicate-button"
      title="Duplicate"
      @mousedown.prevent
      @click="emit('duplicate')"
    >
      Duplicate
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="shape-delete-button"
      title="Delete shape"
      @mousedown.prevent
      @click="emit('delete')"
    >
      Delete
    </UiButton>
  </div>
</template>

<style scoped>
.shape-toolbar {
  position: fixed;
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2);
  max-width: 640px;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

/* Placed above the shape, `top` is where the toolbar's bottom edge belongs:
   lifting by its own rendered height is the one measurement CSS can do that
   the caller cannot, and it is what keeps a two- or three-row toolbar off
   the shape it belongs to. */
.shape-toolbar-above {
  transform: translateY(-100%);
}

.shape-toolbar-field {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-primary);
  white-space: nowrap;
}

.shape-toolbar-number {
  width: 3.5em;
}

.shape-toolbar-alt {
  width: 8em;
}

.shape-toolbar-select {
  font-size: var(--text-md);
}

.shape-toolbar-separator {
  width: 1px;
  align-self: stretch;
  background: var(--border-default);
}
</style>
