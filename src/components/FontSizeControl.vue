<script setup lang="ts">
import { ref, useTemplateRef, watch } from "vue";
import {
  FONT_SIZE_RANGE_MESSAGE,
  parseFontSizeInput,
  stepFontSize,
} from "../lib/fonts";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The selection's font size in points; null renders the blank/
   * indeterminate state for a mixed selection (SPEC-013 AC). */
  value: { type: Number, default: null },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["change"]);

// The input is free-typed and only committed on Enter/blur; the draft
// mirrors the prop between commits.
const draft = ref(props.value === null ? "" : String(props.value));
watch(
  () => props.value,
  (value) => {
    draft.value = value === null ? "" : String(value);
  },
);

// Shown after out-of-range input was clamped (SPEC-013 AC); cleared on the
// next edit or after a few seconds.
const clampMessage = ref("");
let clampTimer = 0;
const root = useTemplateRef("root");
const tooltip = useTemplateRef("tooltip");

// A fixed-position tooltip, like the toolbar's menus: .toolbar clips on
// overflow-y, and this hint hangs below the control — absolutely positioned
// it was cut off at the strip's edge, leaving the user's out-of-range input
// silently corrected with no explanation.
const { menuStyle } = useFloatingMenu({
  open: () => clampMessage.value !== "",
  anchor: () => root.value,
  menu: () => tooltip.value,
  offsetY: 6,
  align: "center",
});

function showClampTooltip() {
  clampMessage.value = FONT_SIZE_RANGE_MESSAGE;
  clearTimeout(clampTimer);
  clampTimer = window.setTimeout(() => {
    clampMessage.value = "";
  }, 3000);
}

function commit() {
  const { value, clamped } = parseFontSizeInput(draft.value);
  if (value === undefined) {
    // Not a number: restore the last known size rather than guessing.
    draft.value = props.value === null ? "" : String(props.value);
    return;
  }
  if (clamped) {
    showClampTooltip();
  }
  draft.value = String(value);
  emit("change", value);
}

function handleInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    draft.value = target.value;
    clampMessage.value = "";
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter") {
    event.preventDefault();
    commit();
  }
}

function step(direction = 1) {
  if (props.disabled || props.value === null) {
    return;
  }
  const next = stepFontSize(props.value, direction === 1 ? 1 : -1);
  if (next !== props.value) {
    emit("change", next);
  }
}
</script>

<template>
  <div
    ref="root"
    class="font-size-control"
    role="group"
    aria-label="Font size"
  >
    <button
      type="button"
      class="size-step"
      data-testid="font-size-decrease"
      title="Decrease font size"
      aria-label="Decrease font size"
      :disabled="disabled || value === null"
      @mousedown.prevent
      @click="step(-1)"
    >
      −
    </button>
    <input
      class="size-input"
      data-testid="font-size-input"
      type="text"
      inputmode="decimal"
      title="Font size"
      aria-label="Font size in points"
      :disabled="disabled"
      :value="draft"
      @input="handleInput"
      @keydown="handleKeydown"
      @blur="commit"
    >
    <button
      type="button"
      class="size-step"
      data-testid="font-size-increase"
      title="Increase font size"
      aria-label="Increase font size"
      :disabled="disabled || value === null"
      @mousedown.prevent
      @click="step(1)"
    >
      +
    </button>
    <span
      v-if="clampMessage"
      ref="tooltip"
      class="clamp-tooltip"
      role="tooltip"
      data-testid="font-size-tooltip"
      :style="menuStyle"
    >{{ clampMessage }}</span>
  </div>
</template>

<style scoped>
.font-size-control {
  display: inline-flex;
  align-items: center;
}

.size-step {
  font: inherit;
  font-size: 0.875rem;
  width: 1.5rem;
  height: 1.75rem;
  padding: 0;
  line-height: 1;
  border: 1px solid #bdbdbd;
  background: #fff;
  cursor: pointer;
}

.size-step:first-child {
  border-radius: 4px 0 0 4px;
}

.size-step:disabled {
  opacity: 0.5;
  cursor: default;
}

.size-input {
  font: inherit;
  font-size: 0.875rem;
  width: 2.75rem;
  height: 1.75rem;
  padding: 0 0.25rem;
  text-align: center;
  border: 1px solid #bdbdbd;
  border-left: none;
  border-right: none;
  box-sizing: border-box;
}

.size-input:disabled {
  opacity: 0.5;
}

/* The + button closes the group visually. */
.font-size-control > .size-step:nth-of-type(2) {
  border-radius: 0 4px 4px 0;
}

/* position/left/top come inline from useFloatingMenu's menuStyle. */
.clamp-tooltip {
  z-index: 30;
  padding: 0.25rem 0.5rem;
  border-radius: 4px;
  background: #3c4043;
  color: #fff;
  font-size: 0.75rem;
  white-space: nowrap;
}
</style>
