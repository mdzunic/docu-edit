<script setup lang="ts">
// The floating shape toolbar (SPEC-051 AC-12..16, AC-25, AC-29..32, AC-34,
// AC-39, AC-44): shown near a selected shape, offering fill, outline,
// rotation, wrap mode, alt text, z-order and group/ungroup. Mirrors
// ImageToolbar.vue's own pattern and positioning convention. Runtime
// prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed, ref, watch } from "vue";
import UiButton from "./ui/UiButton.vue";
import ShapeStylePanel from "./ShapeStylePanel.vue";
import { useTier } from "../composables/useResponsiveTier";

const props = defineProps({
  fill: { type: String, default: "" },
  outlineColor: { type: String, default: "" },
  outlineWidthPt: { type: Number, default: 1 },
  outlineDash: { type: String, default: "solid" },
  rotationDeg: { type: Number, default: 0 },
  wrap: { type: String, default: "inline" },
  altText: { type: String, default: "" },
  grouped: { type: Boolean, default: false },
  canGroup: { type: Boolean, default: false },
  /** SPEC-072: whether align / distribute apply to the current selection. */
  canAlign: { type: Boolean, default: false },
  canDistribute: { type: Boolean, default: false },
  /** SPEC-073: the full styling snapshot ShapeStylePanel renders — { fill,
   * outline, shadow, isLine, isConnector }. */
  styling: { type: Object, default: () => ({}) },
  /** The document's theme (SPEC-077), forwarded to ShapeStylePanel. */
  theme: { type: Object, default: null },
});

const emit = defineEmits([
  "fill",
  "outline",
  "shadow",
  "style-preset",
  "preview-fill",
  "preview-outline",
  "pick-picture",
  "rotation",
  "wrap",
  "alt-text",
  "z-order",
  "group",
  "ungroup",
  "options",
  "delete",
  "duplicate",
  "edit-points",
  "align",
  "distribute",
]);

// SPEC-072 align/distribute on the contextual toolbar (AC-15). The value is the
// command suffix (`drawing.align-<value>` / `drawing.distribute-<value>`).
const ALIGN_COMMANDS = [
  { command: "left", label: "Align left" },
  { command: "center", label: "Align centre" },
  { command: "right", label: "Align right" },
  { command: "top", label: "Align top" },
  { command: "middle", label: "Align middle" },
  { command: "bottom", label: "Align bottom" },
];
const DISTRIBUTE_COMMANDS = [
  { command: "horizontal", label: "Distribute horizontally" },
  { command: "vertical", label: "Distribute vertically" },
];

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

// SPEC-068: the drawing position dialog is one of the commands disabled at
// the compact tier (a 390px canvas cannot host it) — it stays visible with a
// stated reason rather than vanishing (AC-12/13). A shape can still be moved
// by drag and deleted at compact; only this precise numeric editor is off.
const tier = useTier();
const positionUnavailable = computed(() => tier.value === "compact");
const POSITION_UNAVAILABLE_REASON = "not available on this screen size";

// SPEC-073 AC-16/AC-17: the toolbar's full styling picker lives in a
// popover so the quick controls stay one click; the identical panel is also
// hosted by the Format → Drawing → Shape formatting… dialog.
const styleOpen = ref(false);

const rotationInput = ref(String(props.rotationDeg));
watch(
  () => props.rotationDeg,
  (deg) => {
    rotationInput.value = String(deg);
  },
);

function onFillInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("fill", { kind: "solid", color: target.value });
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

/** The outline colour a width/dash edit re-emits: the model's own value from
 * `styling` (a theme reference stays a reference — SPEC-077 AC-18), not the
 * resolved hex the colour input displays. */
function storedOutlineColor() {
  return (props.styling.outline && props.styling.outline.color) || props.outlineColor || "#000000";
}

function onOutlineWidthInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    const widthPt = Number(target.value) || 0;
    emit("outline", { color: storedOutlineColor(), widthPt, dash: props.outlineDash });
  }
}

function onOutlineDashChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("outline", { color: storedOutlineColor(), widthPt: props.outlineWidthPt, dash: target.value });
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

function onAlignChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement && target.value) {
    emit("align", target.value);
    target.value = "";
  }
}

function onDistributeChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement && target.value) {
    emit("distribute", target.value);
    target.value = "";
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
  <div class="shape-toolbar-controls">
    <span class="shape-style-popover-host">
      <UiButton
        variant="ghost"
        data-testid="shape-style-button"
        :aria-expanded="styleOpen"
        @mousedown.prevent
        @click="styleOpen = !styleOpen"
      >
        Style
      </UiButton>
      <div
        v-if="styleOpen"
        class="shape-style-popover"
        data-testid="shape-style-popover"
        @mousedown.stop
      >
        <ShapeStylePanel
          :styling="styling"
          :theme="theme"
          @fill="(value) => emit('fill', value)"
          @outline="(value) => emit('outline', value)"
          @shadow="(value) => emit('shadow', value)"
          @preset="(id) => emit('style-preset', id)"
          @preview-fill="(value) => emit('preview-fill', value)"
          @preview-outline="(value) => emit('preview-outline', value)"
          @pick-picture="emit('pick-picture')"
        />
      </div>
    </span>
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
      data-testid="shape-align"
      :disabled="!canAlign"
      title="Align drawings"
      @change="onAlignChange"
      @mousedown.stop
    >
      <option
        value=""
        disabled
        selected
      >
        Align…
      </option>
      <option
        v-for="option in ALIGN_COMMANDS"
        :key="option.command"
        :value="option.command"
      >
        {{ option.label }}
      </option>
    </select>
    <select
      class="shape-toolbar-select"
      data-testid="shape-distribute"
      :disabled="!canDistribute"
      title="Distribute drawings"
      @change="onDistributeChange"
      @mousedown.stop
    >
      <option
        value=""
        disabled
        selected
      >
        Distribute…
      </option>
      <option
        v-for="option in DISTRIBUTE_COMMANDS"
        :key="option.command"
        :value="option.command"
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
      data-testid="shape-position-button"
      :disabled="positionUnavailable"
      :title="positionUnavailable ? POSITION_UNAVAILABLE_REASON : 'Position and size'"
      @mousedown.prevent
      @click="emit('options')"
    >
      Position &amp; size…
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="shape-edit-points-button"
      title="Edit points"
      @mousedown.prevent
      @click="emit('edit-points')"
    >
      Edit points
    </UiButton>
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
/* The controls are direct flex children of the shared ContextualToolbar; this
   wrapper only groups them in the template (display:contents removes its box
   so the separators' align-self:stretch still resolves against the bar). */
.shape-toolbar-controls {
  display: contents;
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
.shape-style-popover-host {
  position: relative;
  display: inline-flex;
}
.shape-style-popover {
  position: absolute;
  top: calc(100% + var(--space-2));
  left: 0;
  z-index: 40;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-3);
  padding: var(--space-4);
}
</style>
