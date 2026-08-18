<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { shortcutLabel } from "../lib/shortcuts";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// Paragraph geometry controls (SPEC-014): alignment, indent step buttons,
// the line-spacing menu (presets, the space-before toggle, "Custom
// spacing…") and the RTL direction toggle — plus the bulleted/numbered
// list toggles with their marker-style dropdowns (SPEC-018). Runtime
// prop/emit declarations, not the defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** The first touched paragraph's effective alignment ("left" when the
   * paragraph doesn't set one) — drives the pressed state (AC-1). */
  alignment: { type: String, default: "left" },
  /** True when the first touched paragraph has a non-zero spaceBeforePt —
   * the toggle's label source (AC-15..17). */
  spaceBefore: { type: Boolean, default: false },
  /** The first touched paragraph's direction; "ltr" when unset (AC-24). */
  direction: { type: String, default: "ltr" },
  /** True when every touched paragraph is a bulleted list item — the
   * bulleted toggle's pressed state (SPEC-018 AC-1). */
  bulleted: { type: Boolean, default: false },
  /** True when every touched paragraph is a numbered list item (AC-2). */
  numbered: { type: Boolean, default: false },
  /** The current level's effective marker style key ("disc", "decimal",
   * …), "" outside a list — checks the matching dropdown entry (AC-12). */
  listMarker: { type: String, default: "" },
  /** The first touched paragraph's pagination flags (SPEC-028 AC-22..28) —
   * commands.paragraphPaginationState's shape; widowControl reads true by
   * default. */
  pagination: {
    type: Object,
    default: () => ({ keepNext: false, keepLines: false, pageBreakBefore: false, widowControl: true }),
  },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits([
  "align",
  "indent",
  "spacing-preset",
  "custom-spacing",
  "toggle-space-before",
  "toggle-direction",
  "toggle-list",
  "list-marker",
  "toggle-pagination",
]);

// Static per mount, like FormatToolbar's FORMATS list. Shortcut text is read
// from the SPEC-021 registry rather than hardcoded (AC-15, AC-16).
const ALIGNMENTS = [
  { key: "left", label: "Align left", shortcut: shortcutLabel("paragraph.alignLeft") },
  { key: "center", label: "Align centre", shortcut: shortcutLabel("paragraph.alignCenter") },
  { key: "right", label: "Align right", shortcut: shortcutLabel("paragraph.alignRight") },
  { key: "justify", label: "Justify", shortcut: shortcutLabel("paragraph.alignJustify") },
];

const decreaseIndentTitle = `Decrease indent (${shortcutLabel("paragraph.outdent")} at the start of a paragraph)`;
const increaseIndentTitle = `Increase indent (${shortcutLabel("paragraph.indent")} at the start of a paragraph)`;

// The four fixed presets of the line spacing menu (AC-11), as multiples.
const SPACING_PRESETS = [
  { key: "single", label: "Single", value: 1 },
  { key: "115", label: "1.15", value: 1.15 },
  { key: "15", label: "1.5", value: 1.5 },
  { key: "double", label: "Double", value: 2 },
];

// The two list buttons' marker dropdowns (SPEC-018 AC-12): every entry is a
// marker style key from lib/lists.ts, previewed by its displayed marker.
const BULLET_STYLES = [
  { key: "disc", label: "•" },
  { key: "circle", label: "◦" },
  { key: "square", label: "▪" },
];
const NUMBER_STYLES = [
  { key: "decimal", label: "1." },
  { key: "lowerAlpha", label: "a." },
  { key: "upperAlpha", label: "A." },
  { key: "lowerRoman", label: "i." },
  { key: "upperRoman", label: "I." },
];

// The pagination menu's four independent toggles (SPEC-028): each row is a
// menuitemcheckbox whose checked state comes from the pagination prop.
const PAGINATION_OPTIONS = [
  { key: "keepNext", label: "Keep with next" },
  { key: "keepLines", label: "Keep lines together" },
  { key: "pageBreakBefore", label: "Page break before" },
  { key: "widowControl", label: "Widow/orphan control" },
];

const spacingOpen = ref(false);
const root = useTemplateRef("root");
const spacingButton = useTemplateRef("spacingButton");
const spacingMenu = useTemplateRef("spacingMenu");
const bulletMenuOpen = ref(false);
const bulletMenuButton = useTemplateRef("bulletMenuButton");
const bulletMenu = useTemplateRef("bulletMenu");
const numberMenuOpen = ref(false);
const numberMenuButton = useTemplateRef("numberMenuButton");
const numberMenu = useTemplateRef("numberMenu");
const paginationOpen = ref(false);
const paginationButton = useTemplateRef("paginationButton");
const paginationMenu = useTemplateRef("paginationMenu");

// Fixed-position menu escaping the toolbar strip's clipping — the same
// pattern as HighlightPicker.
const { menuStyle } = useFloatingMenu({
  open: () => spacingOpen.value,
  anchor: () => spacingButton.value,
  menu: () => spacingMenu.value,
});

const { menuStyle: bulletMenuStyle } = useFloatingMenu({
  open: () => bulletMenuOpen.value,
  anchor: () => bulletMenuButton.value,
  menu: () => bulletMenu.value,
});

const { menuStyle: numberMenuStyle } = useFloatingMenu({
  open: () => numberMenuOpen.value,
  anchor: () => numberMenuButton.value,
  menu: () => numberMenu.value,
});

const { menuStyle: paginationMenuStyle } = useFloatingMenu({
  open: () => paginationOpen.value,
  anchor: () => paginationButton.value,
  menu: () => paginationMenu.value,
});

function toggleSpacingMenu() {
  if (!props.disabled) {
    spacingOpen.value = !spacingOpen.value;
  }
}

function togglePaginationMenu() {
  if (!props.disabled) {
    paginationOpen.value = !paginationOpen.value;
  }
}

function isPaginationChecked(key = "") {
  const state = props.pagination ?? {};
  if (key === "keepNext") {
    return state.keepNext === true;
  }
  if (key === "keepLines") {
    return state.keepLines === true;
  }
  if (key === "pageBreakBefore") {
    return state.pageBreakBefore === true;
  }
  return state.widowControl !== false;
}

function choosePaginationToggle(key = "") {
  // The menu stays open: flipping several pagination flags in one visit is
  // the common flow, and aria-checked updating in place shows the effect.
  emit("toggle-pagination", key);
}

function toggleBulletMenu() {
  if (!props.disabled) {
    numberMenuOpen.value = false;
    bulletMenuOpen.value = !bulletMenuOpen.value;
  }
}

function toggleNumberMenu() {
  if (!props.disabled) {
    bulletMenuOpen.value = false;
    numberMenuOpen.value = !numberMenuOpen.value;
  }
}

function chooseListMarker(style = "") {
  bulletMenuOpen.value = false;
  numberMenuOpen.value = false;
  emit("list-marker", style);
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
  const anyOpen =
    spacingOpen.value || bulletMenuOpen.value || numberMenuOpen.value || paginationOpen.value;
  if (anyOpen && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    spacingOpen.value = false;
    bulletMenuOpen.value = false;
    numberMenuOpen.value = false;
    paginationOpen.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    spacingOpen.value = false;
    bulletMenuOpen.value = false;
    numberMenuOpen.value = false;
    paginationOpen.value = false;
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
      :title="decreaseIndentTitle"
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
      :title="increaseIndentTitle"
      aria-label="Increase indent"
      @mousedown.prevent
      @click="emit('indent', 1)"
    >
      ⇥
    </button>
    <!-- Bulleted/numbered list toggles (SPEC-018 AC-1, AC-2), each with a
         marker-style dropdown on a separate expander next to the toggle
         (AC-12) so a plain click still just toggles. -->
    <div class="list-cluster">
      <button
        type="button"
        class="paragraph-button list-toggle"
        data-testid="list-bulleted"
        :disabled="disabled"
        title="Bulleted list"
        aria-label="Bulleted list"
        :aria-pressed="bulleted ? 'true' : 'false'"
        @mousedown.prevent
        @click="emit('toggle-list', 'bullet')"
      >
        •≡
      </button>
      <button
        ref="bulletMenuButton"
        type="button"
        class="paragraph-button list-menu-button"
        data-testid="list-bulleted-menu-button"
        :disabled="disabled"
        title="Bullet styles"
        aria-label="Bullet styles"
        aria-haspopup="menu"
        :aria-expanded="bulletMenuOpen ? 'true' : 'false'"
        @mousedown.prevent
        @click="toggleBulletMenu"
      >
        ▾
      </button>
      <div
        v-if="bulletMenuOpen"
        ref="bulletMenu"
        class="marker-menu"
        data-testid="bullet-style-menu"
        role="menu"
        aria-label="Bullet styles"
        :style="bulletMenuStyle"
      >
        <button
          v-for="style in BULLET_STYLES"
          :key="style.key"
          type="button"
          role="menuitemradio"
          :aria-checked="listMarker === style.key ? 'true' : 'false'"
          :data-testid="`bullet-style-${style.key}`"
          :class="{ checked: listMarker === style.key }"
          @mousedown.prevent
          @click="chooseListMarker(style.key)"
        >
          {{ style.label }}
        </button>
      </div>
    </div>
    <div class="list-cluster">
      <button
        type="button"
        class="paragraph-button list-toggle"
        data-testid="list-numbered"
        :disabled="disabled"
        title="Numbered list"
        aria-label="Numbered list"
        :aria-pressed="numbered ? 'true' : 'false'"
        @mousedown.prevent
        @click="emit('toggle-list', 'number')"
      >
        1≡
      </button>
      <button
        ref="numberMenuButton"
        type="button"
        class="paragraph-button list-menu-button"
        data-testid="list-numbered-menu-button"
        :disabled="disabled"
        title="Numbering styles"
        aria-label="Numbering styles"
        aria-haspopup="menu"
        :aria-expanded="numberMenuOpen ? 'true' : 'false'"
        @mousedown.prevent
        @click="toggleNumberMenu"
      >
        ▾
      </button>
      <div
        v-if="numberMenuOpen"
        ref="numberMenu"
        class="marker-menu"
        data-testid="number-style-menu"
        role="menu"
        aria-label="Numbering styles"
        :style="numberMenuStyle"
      >
        <button
          v-for="style in NUMBER_STYLES"
          :key="style.key"
          type="button"
          role="menuitemradio"
          :aria-checked="listMarker === style.key ? 'true' : 'false'"
          :data-testid="`number-style-${style.key}`"
          :class="{ checked: listMarker === style.key }"
          @mousedown.prevent
          @click="chooseListMarker(style.key)"
        >
          {{ style.label }}
        </button>
      </div>
    </div>
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
    <button
      ref="paginationButton"
      type="button"
      class="paragraph-button pagination-button"
      data-testid="paragraph-pagination-button"
      :disabled="disabled"
      title="Paragraph pagination"
      aria-label="Paragraph pagination"
      aria-haspopup="menu"
      :aria-expanded="paginationOpen ? 'true' : 'false'"
      @mousedown.prevent
      @click="togglePaginationMenu"
    >
      ¶⇟
    </button>
    <div
      v-if="paginationOpen"
      ref="paginationMenu"
      class="spacing-menu"
      data-testid="paragraph-pagination-menu"
      role="menu"
      aria-label="Paragraph pagination"
      :style="paginationMenuStyle"
    >
      <button
        v-for="option in PAGINATION_OPTIONS"
        :key="option.key"
        type="button"
        role="menuitemcheckbox"
        :aria-checked="isPaginationChecked(option.key) ? 'true' : 'false'"
        :class="{ checked: isPaginationChecked(option.key) }"
        :data-testid="`pagination-${option.key}`"
        @mousedown.prevent
        @click="choosePaginationToggle(option.key)"
      >
        {{ option.label }}
      </button>
    </div>
  </div>
</template>

<style scoped>
.paragraph-toolbar {
  display: inline-flex;
  gap: var(--space-2);
}

.paragraph-button {
  font: inherit;
  font-size: var(--text-md);
  min-width: 1.75rem;
  height: 1.75rem;
  padding: 0 var(--space-2);
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.paragraph-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.paragraph-button[aria-pressed="true"] {
  background: var(--accent-subtle);
  border-color: var(--accent-subtle);
}

/* Alignment glyphs: four text lines whose ragged edge shows the alignment.
   Drawn with gradients so no icon asset or extra dependency is needed. */
.align-glyph {
  width: 14px;
  height: 12px;
  background-repeat: no-repeat;
  background-image:
    linear-gradient(var(--text-secondary), var(--text-secondary)),
    linear-gradient(var(--text-secondary), var(--text-secondary)),
    linear-gradient(var(--text-secondary), var(--text-secondary)),
    linear-gradient(var(--text-secondary), var(--text-secondary));
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
  padding: var(--space-2) 0;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-2);
}

.spacing-menu button {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: left;
  background: none;
  border: none;
  padding: var(--space-3) var(--space-5);
  cursor: pointer;
  white-space: nowrap;
}

.spacing-menu button:hover {
  background: var(--surface-canvas);
}

.spacing-separator {
  margin: var(--space-2) 0;
  border: none;
  border-top: var(--border-width) solid var(--border-default);
}

/* A list toggle and its dropdown expander read as one control. */
.list-cluster {
  display: inline-flex;
}

.list-cluster .list-toggle {
  border-start-end-radius: 0;
  border-end-end-radius: 0;
}

.list-cluster .list-menu-button {
  min-width: 1rem;
  padding: 0;
  font-size: var(--text-xs);
  border-inline-start: none;
  border-start-start-radius: 0;
  border-end-start-radius: 0;
}

/* position/left/top come inline from useFloatingMenu, like .spacing-menu. */
.marker-menu {
  z-index: 30;
  display: flex;
  flex-direction: column;
  min-width: 4rem;
  padding: var(--space-2) 0;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-2);
}

.marker-menu button {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: left;
  background: none;
  border: none;
  padding: var(--space-3) var(--space-5);
  cursor: pointer;
  white-space: nowrap;
}

.marker-menu button:hover {
  background: var(--surface-canvas);
}

.marker-menu button.checked {
  background: var(--accent-subtle);
}
</style>
