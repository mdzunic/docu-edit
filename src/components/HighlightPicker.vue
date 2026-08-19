<script setup lang="ts">
import { ref, useTemplateRef } from "vue";
import { HIGHLIGHT_PICKER_ITEMS } from "../lib/core/colors";
import { HIGHLIGHT_HEX } from "../lib/core/colors";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The selection's highlight: an OOXML palette name, a "#rrggbb" for an
   * out-of-palette colour (from an .odt), or false for no highlight —
   * explicit or unset, which the picker shows identically. */
  value: { type: [String, Boolean], default: false },
  /** True for a mixed selection: nothing renders active (SPEC-013 AC). */
  mixed: { type: Boolean, default: false },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);

const open = ref(false);
const triggerEl = useTemplateRef("triggerEl");
const menuSurface = useTemplateRef("menuSurface");

const ITEMS = HIGHLIGHT_PICKER_ITEMS;

// The toolbar clips overflow on its vertical axis, so the menu is rendered
// position:fixed (escaped below the anchor) instead of at a clipped offset.
// The measured element is UiMenuSurface's rendered box ($el).
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuSurface.value?.$el ?? null,
});

function toggle() {
  if (!props.disabled) {
    open.value = !open.value;
  }
}

// Opened by the Format menu's "Highlight colour" entry (SPEC-049): the
// control is kept in the toolbar but also reachable from the menu.
function openMenu() {
  if (!props.disabled) {
    open.value = true;
  }
}

defineExpose({ openMenu });

/** Emits the OOXML colour name, or false for "No colour" (SPEC-013: an
 * explicit off, not an absence). */
function pick(name = "") {
  emit("select", name === "" ? false : name);
  open.value = false;
}

/** The button's preview bar colour: the palette hex for a name, a raw hex
 * as-is, nothing for no-highlight or mixed. */
function currentHex() {
  if (props.mixed || props.value === false || props.value === true) {
    return "";
  }
  return props.value.startsWith("#") ? props.value : (HIGHLIGHT_HEX[props.value] ?? "");
}

// Outside-click and Escape dismissal (and focus return) are owned by
// UiMenuSurface — the shared, tested implementation.
</script>

<template>
  <div class="highlight-picker">
    <button
      ref="triggerEl"
      type="button"
      class="picker-button"
      data-testid="highlight-button"
      title="Highlight colour"
      aria-label="Highlight colour"
      aria-haspopup="true"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggle"
    >
      <span class="glyph">ab</span>
      <span
        class="color-bar"
        data-testid="highlight-current"
        :class="{ 'color-bar-mixed': mixed, 'color-bar-none': !mixed && currentHex() === '' }"
        :style="currentHex() === '' ? {} : { background: currentHex() }"
        :data-color="mixed ? 'mixed' : (currentHex() === '' ? 'none' : String(value))"
      />
    </button>
    <UiMenuSurface
      ref="menuSurface"
      :open="open"
      :anchor="triggerEl"
      class="picker-menu"
      data-testid="highlight-menu"
      role="dialog"
      aria-label="Highlight colour"
      :style="menuStyle"
      @dismiss="open = false"
    >
      <div class="swatch-grid">
        <!-- The 15 Word highlight colours (SPEC-013 AC); white is in the
             read/write palette but deliberately not offered here. -->
        <button
          v-for="item in ITEMS"
          :key="item.name"
          type="button"
          class="swatch"
          data-testid="highlight-swatch"
          :data-color="item.name"
          :title="item.label"
          :aria-label="`Highlight ${item.label}`"
          :style="{ background: item.hex }"
          :class="{ active: !mixed && value === item.name }"
          @mousedown.prevent
          @click="pick(item.name)"
        />
      </div>
      <button
        type="button"
        class="no-color"
        data-testid="highlight-none"
        :class="{ active: !mixed && (value === false || value === true) }"
        @mousedown.prevent
        @click="pick('')"
      >
        No colour
      </button>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.highlight-picker {
  position: relative;
  display: inline-flex;
}

.picker-button {
  font: inherit;
  font-size: var(--text-sm);
  width: var(--height-control);
  height: var(--height-control);
  padding: 0;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.picker-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.glyph {
  line-height: 1;
}

.color-bar {
  width: 1.125rem;
  height: 0.25rem;
  margin-top: var(--space-1);
  border-radius: var(--radius-sm);
}

.color-bar-none {
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

.color-bar-mixed {
  background: repeating-linear-gradient(
    45deg,
    var(--border-strong) 0 2px,
    var(--surface-sunken) 2px 4px
  );
}

/* position/left/top come inline from useFloatingMenu; the box is UiMenuSurface's. */
.picker-menu {
  z-index: 30;
}

.swatch-grid {
  display: grid;
  grid-template-columns: repeat(5, 1.125rem);
  gap: var(--space-2);
}

.swatch {
  width: 1.125rem;
  height: 1.125rem;
  padding: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.swatch.active {
  outline: calc(2 * var(--border-width)) solid var(--accent);
  outline-offset: var(--border-width);
}

.no-color {
  font: inherit;
  font-size: var(--text-md);
  width: 100%;
  margin-top: var(--space-3);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
  white-space: nowrap;
}

.no-color.active {
  background: var(--accent-subtle);
  border-color: var(--accent);
}
</style>
