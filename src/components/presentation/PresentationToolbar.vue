<script setup lang="ts">
// The presentation formatting toolbar (SPEC-081): font family / size / colour
// reuse the product-neutral document widgets (they already live in
// src/components and read the shared core/fonts + core/colors), and the
// bold/italic/underline/align/bullet/indent controls are toolbar-local buttons.
// Every control reflects the selection's state (mixed → indeterminate) and is
// disabled when nothing is targeted. Alignment/bullet/indent collapse into a ⋯
// overflow menu when the strip is too narrow for them (the e2e-viewport lesson:
// controls must never fall off the edge unreachable).
//
// Runtime prop/emit declarations, not the generic forms — espree lints .vue
// scripts (see AppShell.vue's note).
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import FontFamilySelect from "../FontFamilySelect.vue";
import FontSizeControl from "../FontSizeControl.vue";
import TextColorPicker from "../TextColorPicker.vue";
import { alignIsActive, asFormatState, togglePressed } from "./slideToolbarState";

const props = defineProps({
  /** The selection's formatting state (SelectionFormatState) or null when no
   * text frame is targeted. */
  state: { type: Object, default: null },
  /** Whether the formatting controls act on anything (a selected/edited text
   * frame). */
  enabled: { type: Boolean, default: false },
  /** Recently applied text colours, for the colour picker. */
  recentColors: { type: Array, default: () => [] },
  /** Whether the slide-transition panel is open (SPEC-085). */
  transitionOpen: { type: Boolean, default: false },
  /** Whether the animations panel is open (SPEC-086). */
  animateOpen: { type: Boolean, default: false },
});

const emit = defineEmits([
  "bold",
  "italic",
  "underline",
  "font",
  "size",
  "color",
  "align",
  "bullet",
  "indent",
  "outdent",
  "transition-toggle",
  "animate-toggle",
]);

// SPEC-081 AC-18: the presentation font-size control clamps to 6–96 pt, a
// narrower range than the word processor's 1–409 — passed to the shared
// FontSizeControl so both typed entry and the steppers respect it.
const PRESENTATION_FONT_SIZE_MIN = 6;
const PRESENTATION_FONT_SIZE_MAX = 96;

const s = computed(() => asFormatState(props.state));

// --- control values from the folded selection state ------------------------
const fontValue = computed(() => (typeof s.value?.fontFamily === "string" ? s.value.fontFamily : ""));
const sizeValue = computed(() => (typeof s.value?.sizePt === "number" ? s.value.sizePt : undefined));
const colorValue = computed(() => (typeof s.value?.color === "string" ? s.value.color : ""));
const colorMixed = computed(() => s.value?.color === "mixed");

function pressed(field = "bold") {
  return togglePressed(s.value, field);
}
function alignActive(value = "") {
  return alignIsActive(s.value, value);
}

const ALIGNMENTS = [
  { value: "left", glyph: "⯇", label: "Align left" },
  { value: "center", glyph: "≡", label: "Align centre" },
  { value: "right", glyph: "⯈", label: "Align right" },
  { value: "justify", glyph: "☰", label: "Justify" },
];

// --- overflow (⋯) ----------------------------------------------------------
const root = useTemplateRef("root");
const compact = ref(false);
const overflowOpen = ref(false);
const OVERFLOW_BELOW_PX = 620;

function updateCompact() {
  compact.value = (root.value?.clientWidth ?? 9999) < OVERFLOW_BELOW_PX;
  if (!compact.value) overflowOpen.value = false;
}
const observer = typeof ResizeObserver !== "undefined" ? new ResizeObserver(updateCompact) : null;
onMounted(() => {
  if (observer && root.value) observer.observe(root.value);
});
onBeforeUnmount(() => observer?.disconnect());
</script>

<template>
  <div
    ref="root"
    class="slides-toolbar"
    role="toolbar"
    aria-label="Slide formatting"
    data-testid="slides-toolbar"
  >
    <div
      class="tb-group"
      data-testid="slides-font-family"
    >
      <FontFamilySelect
        :value="fontValue"
        :disabled="!enabled"
        @select="(family) => emit('font', family)"
      />
    </div>

    <div
      class="tb-group"
      data-testid="slides-font-size"
    >
      <FontSizeControl
        :value="sizeValue"
        :min="PRESENTATION_FONT_SIZE_MIN"
        :max="PRESENTATION_FONT_SIZE_MAX"
        :disabled="!enabled"
        @change="(pt) => emit('size', pt)"
      />
    </div>

    <div class="tb-group">
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-bold"
        aria-label="Bold"
        :aria-pressed="pressed('bold')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('bold')"
      >
        <b>B</b>
      </button>
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-italic"
        aria-label="Italic"
        :aria-pressed="pressed('italic')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('italic')"
      >
        <i>I</i>
      </button>
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-underline"
        aria-label="Underline"
        :aria-pressed="pressed('underline')"
        :disabled="!enabled"
        @mousedown.prevent
        @click="emit('underline')"
      >
        <u>U</u>
      </button>
    </div>

    <div
      class="tb-group"
      data-testid="slides-text-color"
    >
      <TextColorPicker
        :value="colorValue"
        :mixed="colorMixed"
        :recent="recentColors"
        :disabled="!enabled"
        @select="(color) => emit('color', color)"
      />
    </div>

    <!-- Alignment / bullet / indent: inline, or in the ⋯ menu when compact -->
    <Teleport
      :to="'#slides-overflow-target'"
      :disabled="!compact"
    >
      <div class="tb-secondary">
        <div
          class="tb-group"
          role="group"
          aria-label="Alignment"
          data-testid="slides-align"
        >
          <button
            v-for="a in ALIGNMENTS"
            :key="a.value"
            type="button"
            class="tb-toggle"
            :aria-label="a.label"
            :aria-pressed="alignActive(a.value) ? 'true' : 'false'"
            :data-align="a.value"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('align', a.value)"
          >
            {{ a.glyph }}
          </button>
        </div>

        <div class="tb-group">
          <button
            type="button"
            class="tb-toggle"
            data-testid="slides-bullet"
            aria-label="Bulleted list"
            :aria-pressed="pressed('bullet')"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('bullet')"
          >
            •
          </button>
          <button
            type="button"
            class="tb-toggle"
            data-testid="slides-outdent"
            aria-label="Decrease indent"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('outdent')"
          >
            ⇤
          </button>
          <button
            type="button"
            class="tb-toggle"
            data-testid="slides-indent"
            aria-label="Increase indent"
            :disabled="!enabled"
            @mousedown.prevent
            @click="emit('indent')"
          >
            ⇥
          </button>
        </div>
      </div>
    </Teleport>

    <div class="tb-group tb-slide-actions">
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-transition-toggle"
        aria-label="Transition"
        :aria-pressed="transitionOpen"
        @click="emit('transition-toggle')"
      >
        Transition
      </button>
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-animate-toggle"
        aria-label="Animate"
        :aria-pressed="animateOpen"
        @click="emit('animate-toggle')"
      >
        Animate
      </button>
    </div>

    <div
      v-show="compact"
      class="tb-group tb-overflow"
    >
      <button
        type="button"
        class="tb-toggle"
        data-testid="slides-overflow"
        aria-label="More formatting"
        aria-haspopup="menu"
        :aria-expanded="overflowOpen"
        @click="overflowOpen = !overflowOpen"
      >
        ⋯
      </button>
      <div
        v-show="overflowOpen"
        class="tb-overflow-menu"
        role="menu"
        data-testid="slides-overflow-menu"
      >
        <div id="slides-overflow-target" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.slides-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
}

.tb-group {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.tb-secondary {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.tb-toggle {
  min-width: 2rem;
  height: 2rem;
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-sm);
}

.tb-toggle:hover:not(:disabled) {
  background: var(--surface-sunken);
}

.tb-toggle[aria-pressed="true"] {
  background: var(--accent-soft, var(--surface-sunken));
  border-color: var(--accent, var(--border-default));
}

.tb-toggle[aria-pressed="mixed"] {
  background: repeating-linear-gradient(
    45deg,
    transparent,
    transparent 3px,
    var(--surface-sunken) 3px,
    var(--surface-sunken) 6px
  );
}

.tb-toggle:disabled {
  opacity: 0.5;
  cursor: default;
}

.tb-slide-actions {
  margin-inline-start: auto;
}

.tb-slide-actions .tb-toggle {
  width: auto;
}

.tb-overflow {
  position: relative;
}

.tb-overflow-menu {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-1));
  z-index: 40;
  padding: var(--space-2);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}
</style>
