<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { HIGHLIGHT_PICKER_ITEMS } from "../lib/colors";
import { HIGHLIGHT_HEX } from "../lib/odf/properties";
import { useFloatingMenu } from "../composables/useFloatingMenu";

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
const root = useTemplateRef("root");
const triggerEl = useTemplateRef("triggerEl");
const menuEl = useTemplateRef("menuEl");

const ITEMS = HIGHLIGHT_PICKER_ITEMS;

// The toolbar clips overflow on its vertical axis, so the menu is rendered
// position:fixed (escaped below the anchor) instead of at a clipped offset.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuEl.value,
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

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (open.value && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    open.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    open.value = false;
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
    class="highlight-picker"
    @keydown="handleKeydown"
  >
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
    <div
      v-if="open"
      ref="menuEl"
      class="picker-menu"
      data-testid="highlight-menu"
      role="dialog"
      aria-label="Highlight colour"
      :style="menuStyle"
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
    </div>
  </div>
</template>

<style scoped>
.highlight-picker {
  position: relative;
  display: inline-flex;
}

.picker-button {
  font: inherit;
  font-size: 0.75rem;
  width: 1.75rem;
  height: 1.75rem;
  padding: 0;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
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
  margin-top: 0.125rem;
  border-radius: 1px;
}

.color-bar-none {
  background: #fff;
  border: 1px solid #bdbdbd;
  box-sizing: border-box;
}

.color-bar-mixed {
  background: repeating-linear-gradient(45deg, #bbb 0 2px, #eee 2px 4px);
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   toolbar's vertical clipping). */
.picker-menu {
  z-index: 30;
  padding: 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.swatch-grid {
  display: grid;
  grid-template-columns: repeat(5, 1.125rem);
  gap: 0.25rem;
}

.swatch {
  width: 1.125rem;
  height: 1.125rem;
  padding: 0;
  border: 1px solid rgb(0 0 0 / 0.25);
  border-radius: 2px;
  cursor: pointer;
}

.swatch.active {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}

.no-color {
  font: inherit;
  font-size: 0.8125rem;
  width: 100%;
  margin-top: 0.5rem;
  padding: 0.125rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.no-color.active {
  background: #d2e3fc;
  border-color: #8ab4f8;
}
</style>
