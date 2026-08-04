<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { modifierKeyLabel } from "../lib/platform";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// Paragraph geometry controls (SPEC-014): alignment, indent step buttons,
// the line-spacing menu (presets, the space-before toggle, "Custom
// spacing…") and the RTL direction toggle. Runtime prop/emit declarations,
// not the defineProps<{...}>() generics — see AppShell.vue's comment on the
// same espree constraint.
const props = defineProps({
  /** The first touched paragraph's effective alignment ("left" when the
   * paragraph doesn't set one) — drives the pressed state (AC-1). */
  alignment: { type: String, default: "left" },
  /** True when the first touched paragraph has a non-zero spaceBeforePt —
   * the toggle's label source (AC-15..17). */
  spaceBefore: { type: Boolean, default: false },
  /** The first touched paragraph's direction; "ltr" when unset (AC-24). */
  direction: { type: String, default: "ltr" },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits([
  "align",
  "indent",
  "spacing-preset",
  "custom-spacing",
  "toggle-space-before",
  "toggle-direction",
]);

const mod = modifierKeyLabel();

// Static per mount, like FormatToolbar's FORMATS list.
const ALIGNMENTS = [
  { key: "left", label: "Align left", shortcut: `${mod}+L` },
  { key: "center", label: "Align centre", shortcut: `${mod}+E` },
  { key: "right", label: "Align right", shortcut: `${mod}+R` },
  { key: "justify", label: "Justify", shortcut: `${mod}+J` },
];

// The four fixed presets of the line spacing menu (AC-11), as multiples.
const SPACING_PRESETS = [
  { key: "single", label: "Single", value: 1 },
  { key: "115", label: "1.15", value: 1.15 },
  { key: "15", label: "1.5", value: 1.5 },
  { key: "double", label: "Double", value: 2 },
];

const spacingOpen = ref(false);
const root = useTemplateRef("root");
const spacingButton = useTemplateRef("spacingButton");
const spacingMenu = useTemplateRef("spacingMenu");

// Fixed-position menu escaping the toolbar strip's clipping — the same
// pattern as HighlightPicker.
const { menuStyle } = useFloatingMenu({
  open: () => spacingOpen.value,
  anchor: () => spacingButton.value,
  menu: () => spacingMenu.value,
});

function toggleSpacingMenu() {
  if (!props.disabled) {
    spacingOpen.value = !spacingOpen.value;
  }
}

function chooseSpacingPreset(value = 1) {
  spacingOpen.value = false;
  emit("spacing-preset", value);
}

function chooseCustomSpacing() {
  spacingOpen.value = false;
  emit("custom-spacing");
}

function chooseSpaceBeforeToggle() {
  spacingOpen.value = false;
  emit("toggle-space-before");
}

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (spacingOpen.value && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    spacingOpen.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    spacingOpen.value = false;
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
    class="paragraph-toolbar"
    role="group"
    aria-label="Paragraph formatting"
    @keydown="handleKeydown"
  >
    <button
      v-for="entry in ALIGNMENTS"
      :key="entry.key"
      type="button"
      class="paragraph-button"
      :data-testid="`align-${entry.key}`"
      :disabled="disabled"
      :title="`${entry.label} (${entry.shortcut})`"
      :aria-label="`${entry.label} (${entry.shortcut})`"
      :aria-pressed="alignment === entry.key ? 'true' : 'false'"
      @mousedown.prevent
      @click="emit('align', entry.key)"
    >
      <span
        class="align-glyph"
        :class="`align-glyph-${entry.key}`"
        aria-hidden="true"
      />
    </button>
    <button
      type="button"
      class="paragraph-button"
      data-testid="indent-decrease"
      :disabled="disabled"
      title="Decrease indent (Shift+Tab at the start of a paragraph)"
      aria-label="Decrease indent"
      @mousedown.prevent
      @click="emit('indent', -1)"
    >
      ⇤
    </button>
    <button
      type="button"
      class="paragraph-button"
      data-testid="indent-increase"
      :disabled="disabled"
      title="Increase indent (Tab at the start of a paragraph)"
      aria-label="Increase indent"
      @mousedown.prevent
      @click="emit('indent', 1)"
    >
      ⇥
    </button>
    <button
      type="button"
      class="paragraph-button"
      data-testid="paragraph-direction"
      :disabled="disabled"
      title="Right-to-left paragraph direction"
      aria-label="Right-to-left paragraph direction"
      :aria-pressed="direction === 'rtl' ? 'true' : 'false'"
      @mousedown.prevent
      @click="emit('toggle-direction')"
    >
      ¶⇄
    </button>
    <button
      ref="spacingButton"
      type="button"
      class="paragraph-button spacing-button"
      data-testid="line-spacing-button"
      :disabled="disabled"
      title="Line and paragraph spacing"
      aria-label="Line and paragraph spacing"
      aria-haspopup="menu"
      :aria-expanded="spacingOpen ? 'true' : 'false'"
      @mousedown.prevent
      @click="toggleSpacingMenu"
    >
      ↕≡
    </button>
    <div
      v-if="spacingOpen"
      ref="spacingMenu"
      class="spacing-menu"
      data-testid="line-spacing-menu"
      role="menu"
      aria-label="Line and paragraph spacing"
      :style="menuStyle"
    >
      <button
        v-for="preset in SPACING_PRESETS"
        :key="preset.key"
        type="button"
        role="menuitem"
        :data-testid="`line-spacing-${preset.key}`"
        @mousedown.prevent
        @click="chooseSpacingPreset(preset.value)"
      >
        {{ preset.label }}
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="line-spacing-custom"
        @mousedown.prevent
        @click="chooseCustomSpacing"
      >
        Custom spacing…
      </button>
      <hr class="spacing-separator">
      <!-- One entry, two labels: the AC-15/AC-16 table decides which shows,
           from the first touched paragraph only (AC-17). -->
      <button
        type="button"
        role="menuitem"
        data-testid="space-before-toggle"
        @mousedown.prevent
        @click="chooseSpaceBeforeToggle"
      >
        {{ spaceBefore ? "Remove space before paragraph" : "Add space before paragraph" }}
      </button>
    </div>
  </div>
</template>

<style scoped>
.paragraph-toolbar {
  display: inline-flex;
  gap: 0.25rem;
}

.paragraph-button {
  font: inherit;
  font-size: 0.875rem;
  min-width: 1.75rem;
  height: 1.75rem;
  padding: 0 0.25rem;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.paragraph-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.paragraph-button[aria-pressed="true"] {
  background: #d2e3fc;
  border-color: #8ab4f8;
}

/* Alignment glyphs: four text lines whose ragged edge shows the alignment.
   Drawn with gradients so no icon asset or extra dependency is needed. */
.align-glyph {
  width: 14px;
  height: 12px;
  background-repeat: no-repeat;
  background-image:
    linear-gradient(#5f6368, #5f6368),
    linear-gradient(#5f6368, #5f6368),
    linear-gradient(#5f6368, #5f6368),
    linear-gradient(#5f6368, #5f6368);
  background-size: 100% 2px, 70% 2px, 100% 2px, 70% 2px;
}

.align-glyph-left {
  background-position: left 0, left 3.33px, left 6.66px, left 10px;
}

.align-glyph-center {
  background-position: center 0, center 3.33px, center 6.66px, center 10px;
}

.align-glyph-right {
  background-position: right 0, right 3.33px, right 6.66px, right 10px;
}

.align-glyph-justify {
  background-size: 100% 2px, 100% 2px, 100% 2px, 100% 2px;
  background-position: left 0, left 3.33px, left 6.66px, left 10px;
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   toolbar's vertical clipping). */
.spacing-menu {
  z-index: 30;
  min-width: 14rem;
  display: flex;
  flex-direction: column;
  padding: 0.25rem 0;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.spacing-menu button {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
}

.spacing-menu button:hover {
  background: #f1f3f4;
}

.spacing-separator {
  margin: 0.25rem 0;
  border: none;
  border-top: 1px solid #e0e0e0;
}
</style>
