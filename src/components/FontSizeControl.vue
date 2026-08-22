<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef, watch } from "vue";
import {
  COMMON_FONT_SIZES,
  FONT_SIZE_RANGE_MESSAGE,
  parseFontSizeInput,
  stepFontSize,
} from "../lib/fonts";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import AppIcon from "./AppIcon.vue";

// The font-size control (SPEC-013, SPEC-061 AC-21): one unit made of a −
// stepper, an editable numeric field, a + stepper and a dropdown of common
// sizes. Font-size is the one control whose meaning is the number itself
// (AC-2's stated exception), so it stays a text field rather than an icon.
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
const sizeButton = useTemplateRef("sizeButton");
const sizeMenu = useTemplateRef("sizeMenu");
const menuOpen = ref(false);

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

// The common-sizes dropdown is a fixed-position menu escaping the strip's
// clipping — the same pattern as HighlightPicker's swatch popup (SPEC-061
// AC-21: the dropdown must not be painted over by the canvas below).
const { menuStyle: sizeMenuStyle } = useFloatingMenu({
  open: () => menuOpen.value,
  anchor: () => sizeButton.value,
  menu: () => sizeMenu.value,
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

function toggleMenu() {
  if (!props.disabled) {
    menuOpen.value = !menuOpen.value;
  }
}

function chooseSize(size = 11) {
  menuOpen.value = false;
  draft.value = String(size);
  emit("change", size);
}

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (
    menuOpen.value &&
    root.value &&
    sizeMenu.value &&
    event.target instanceof Node &&
    !root.value.contains(event.target) &&
    !sizeMenu.value.contains(event.target)
  ) {
    menuOpen.value = false;
  }
}

function handleMenuKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    menuOpen.value = false;
  }
}

document.addEventListener("mousedown", handleDocumentMousedown);
onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleDocumentMousedown);
});
</script>

<template>
  <div
    ref="root"
    class="font-size-control"
    role="group"
    aria-label="Font size"
    @keydown="handleMenuKeydown"
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
    <button
      ref="sizeButton"
      type="button"
      class="size-step size-menu-button"
      data-testid="font-size-menu-button"
      title="Font size"
      aria-label="Common font sizes"
      aria-haspopup="menu"
      :aria-expanded="menuOpen ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      <AppIcon name="chevron-down" />
    </button>
    <div
      v-if="menuOpen"
      ref="sizeMenu"
      class="size-menu"
      data-testid="font-size-menu"
      role="menu"
      aria-label="Common font sizes"
      :style="sizeMenuStyle"
    >
      <button
        v-for="size in COMMON_FONT_SIZES"
        :key="size"
        type="button"
        role="menuitemradio"
        :aria-checked="value === size ? 'true' : 'false'"
        :class="{ checked: value === size }"
        :data-testid="`font-size-option-${size}`"
        @mousedown.prevent
        @click="chooseSize(size)"
      >
        {{ size }}
      </button>
    </div>
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
  position: relative;
  z-index: 1;
  font: inherit;
  font-size: var(--text-md);
  width: 1.5rem;
  height: var(--height-control);
  padding: 0;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: var(--border-width) solid var(--border-strong);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

/* A transparent out-of-flow hit area lifts each stepper to the 32×32 minimum
   (AC-48) without widening the 24px visual button — so the toolbar stays
   within its 1280px budget (AC-60). Clicks on the pseudo dispatch on the host
   <button>, and z-index:1 lifts the overhang above the sibling input. */
.size-step::before {
  content: "";
  position: absolute;
  top: 50%;
  left: 50%;
  width: var(--height-control);
  height: var(--height-control);
  transform: translate(-50%, -50%);
}

.size-step:first-child {
  border-radius: var(--radius-sm) 0 0 var(--radius-sm);
}

.size-step:disabled {
  opacity: 0.5;
  cursor: default;
}

.size-input {
  font: inherit;
  font-size: var(--text-md);
  width: 2.75rem;
  height: var(--height-control);
  padding: 0 var(--space-2);
  text-align: center;
  border: var(--border-width) solid var(--border-strong);
  border-left: none;
  border-right: none;
  box-sizing: border-box;
}

.size-input:disabled {
  opacity: 0.5;
}

/* The dropdown expander closes the group visually. */
.size-menu-button {
  border-inline-start: none;
  border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   toolbar's vertical clipping). z-index over the canvas so its options are
   clickable, not painted over (AC-21). */
.size-menu {
  z-index: 30;
  display: flex;
  flex-direction: column;
  min-width: 4rem;
  max-height: min(28rem, 80vh);
  overflow-y: auto;
  padding: var(--space-2) 0;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-2);
}

.size-menu button {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: center;
  background: none;
  border: none;
  padding: var(--space-2) var(--space-4);
  cursor: pointer;
  white-space: nowrap;
}

.size-menu button:hover {
  background: var(--surface-canvas);
}

.size-menu button.checked {
  background: var(--accent-subtle);
}

/* position/left/top come inline from useFloatingMenu's menuStyle. */
.clamp-tooltip {
  z-index: 30;
  padding: var(--space-2) var(--space-3);
  border-radius: var(--radius-sm);
  background: var(--text-primary);
  color: var(--text-on-brand);
  font-size: var(--text-sm);
  white-space: nowrap;
}
</style>
