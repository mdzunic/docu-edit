<script setup lang="ts">
// SPEC-070: the floating toolbar for a selected connector — kind, line
// colour/width/dash (AC-17), start/end arrowheads (AC-18), wrap (AC-19),
// z-order (AC-20), alt text (AC-21) and delete. Mirrors ShapeToolbar's
// convention (a print-excluded overlay positioned by the composable). Runtime
// prop/emit declarations per the espree constraint.
import { computed } from "vue";
import { asConnector } from "../lib/documentModel";

const props = defineProps({
  connector: { type: Object, required: true },
});

const emit = defineEmits([
  "set-kind",
  "set-outline",
  "set-arrowheads",
  "set-wrap",
  "set-alt-text",
  "set-z-order",
  "delete",
]);

const ARROWHEADS = ["none", "triangle", "open", "stealth", "diamond", "oval"];
const ARROW_SIZES = ["small", "medium", "large"];
const KINDS = ["straight", "elbow", "curved"];
const WRAPS = ["square", "topAndBottom", "behind", "infront"];

const model = computed(() => asConnector(props.connector));
const outline = computed(() => model.value.outline ?? { color: "#2e5395", widthPt: 1, dash: "solid" });

/** The value of an input/select event target, or "" (instanceof narrowing —
 * the `.vue` script's espree parser bars a TS `as` cast). */
function valueOf(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement || target instanceof HTMLSelectElement) {
    return target.value;
  }
  return "";
}
function emitOutline(patch = {}) {
  emit("set-outline", { ...outline.value, ...patch });
}
function onKind(event = new Event("change")) {
  emit("set-kind", valueOf(event));
}
function onColor(event = new Event("input")) {
  emitOutline({ color: valueOf(event) });
}
function onWidth(event = new Event("input")) {
  emitOutline({ widthPt: Number(valueOf(event)) || 1 });
}
function onDash(event = new Event("change")) {
  emitOutline({ dash: valueOf(event) });
}
function onStartArrow(event = new Event("change")) {
  emit("set-arrowheads", { startArrow: valueOf(event) });
}
function onEndArrow(event = new Event("change")) {
  emit("set-arrowheads", { endArrow: valueOf(event) });
}
function onStartArrowSize(event = new Event("change")) {
  emit("set-arrowheads", { startArrowSize: valueOf(event) });
}
function onEndArrowSize(event = new Event("change")) {
  emit("set-arrowheads", { endArrowSize: valueOf(event) });
}
function onWrap(event = new Event("change")) {
  emit("set-wrap", valueOf(event));
}
function onAltText(event = new Event("input")) {
  emit("set-alt-text", valueOf(event));
}
</script>

<template>
  <div
    class="connector-toolbar print-exclude"
    data-testid="connector-toolbar"
    @mousedown.stop
    @click.stop
  >
    <label class="connector-toolbar-field">
      Kind
      <select
        data-testid="connector-kind"
        :value="model.kind"
        @change="onKind"
      >
        <option
          v-for="k in KINDS"
          :key="k"
          :value="k"
        >{{ k }}</option>
      </select>
    </label>
    <input
      type="color"
      data-testid="connector-color"
      :value="outline.color"
      aria-label="Line colour"
      @input="onColor"
    >
    <input
      type="number"
      class="connector-toolbar-width"
      data-testid="connector-width"
      min="0.25"
      max="12"
      step="0.25"
      :value="outline.widthPt"
      aria-label="Line width (points)"
      @input="onWidth"
    >
    <select
      data-testid="connector-dash"
      :value="outline.dash"
      aria-label="Line dash"
      @change="onDash"
    >
      <option value="solid">
        solid
      </option>
      <option value="dash">
        dash
      </option>
      <option value="dot">
        dot
      </option>
      <option value="dashDot">
        dashDot
      </option>
      <option value="longDash">
        longDash
      </option>
      <option value="longDashDot">
        longDashDot
      </option>
    </select>
    <select
      data-testid="connector-start-arrow"
      :value="model.startArrow ?? 'none'"
      aria-label="Start arrowhead"
      @change="onStartArrow"
    >
      <option
        v-for="a in ARROWHEADS"
        :key="a"
        :value="a"
      >
        {{ a }}
      </option>
    </select>
    <select
      data-testid="connector-start-arrow-size"
      :value="model.startArrowSize ?? 'medium'"
      aria-label="Start arrowhead size"
      @change="onStartArrowSize"
    >
      <option
        v-for="a in ARROW_SIZES"
        :key="a"
        :value="a"
      >
        {{ a }}
      </option>
    </select>
    <select
      data-testid="connector-end-arrow"
      :value="model.endArrow ?? 'triangle'"
      aria-label="End arrowhead"
      @change="onEndArrow"
    >
      <option
        v-for="a in ARROWHEADS"
        :key="a"
        :value="a"
      >
        {{ a }}
      </option>
    </select>
    <select
      data-testid="connector-end-arrow-size"
      :value="model.endArrowSize ?? 'medium'"
      aria-label="End arrowhead size"
      @change="onEndArrowSize"
    >
      <option
        v-for="a in ARROW_SIZES"
        :key="a"
        :value="a"
      >
        {{ a }}
      </option>
    </select>
    <select
      data-testid="connector-wrap"
      :value="model.wrap"
      aria-label="Wrap"
      @change="onWrap"
    >
      <option
        v-for="w in WRAPS"
        :key="w"
        :value="w"
      >
        {{ w }}
      </option>
    </select>
    <button
      type="button"
      data-testid="connector-front"
      title="Bring to front"
      @click="emit('set-z-order', 'front')"
    >
      ⤒
    </button>
    <button
      type="button"
      data-testid="connector-back"
      title="Send to back"
      @click="emit('set-z-order', 'back')"
    >
      ⤓
    </button>
    <input
      type="text"
      class="connector-toolbar-alt"
      data-testid="connector-alt-text"
      :value="model.altText ?? ''"
      placeholder="Alt text"
      aria-label="Alt text"
      @input="onAltText"
    >
    <button
      type="button"
      data-testid="connector-delete"
      title="Delete connector"
      @click="emit('delete')"
    >
      🗑
    </button>
  </div>
</template>

<style scoped>
.connector-toolbar {
  position: absolute;
  z-index: 32;
  display: flex;
  align-items: center;
  gap: var(--space-1);
  padding: var(--space-1) var(--space-2);
  transform: translate(-50%, -100%);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
  white-space: nowrap;
}

.connector-toolbar-field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.connector-toolbar-width {
  width: 3.5rem;
}

.connector-toolbar-alt {
  width: 7rem;
}

.connector-toolbar select,
.connector-toolbar input,
.connector-toolbar button {
  font-size: var(--text-sm);
}
</style>
