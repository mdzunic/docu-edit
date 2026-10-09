<script setup lang="ts">
// The status bar's zoom control (SPEC-063 AC-1..AC-5, AC-55). A − button, an
// editable percentage that doubles as a menu trigger, and a + button — the
// same self-contained button+popover shape as LanguageStatusControl.vue,
// anchored upward with useFloatingMenu because the status bar sits at the
// bottom of the viewport.
//
// The percentage element is BOTH the "menu button" (AC-3) and the "directly
// editable field" (AC-8): it is an <input> the user can type into (Enter or
// blur commits; empty/non-numeric reverts, AC-10), and clicking it opens the
// preset menu. UiMenuSurface leaves focus on the input (return-focus off), so
// the two roles do not fight.
//
// `showFit` (SPEC-117) hides the two fit entries: the workbook view shares
// this control and its ZOOM_PRESETS ladder, but a sheet has no page to fit.
import { computed, ref, useTemplateRef, watch } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { ZOOM_PRESETS } from "../lib/zoom";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// Runtime prop/emit declarations (not the type-generic form): ESLint parses
// .vue scripts with espree, which bars TS-only syntax here — see CLAUDE.md.
const props = defineProps({
  /** The live zoom as a percentage (may be off-ladder, e.g. a fit result). */
  zoomPercent: { type: Number, default: 100 },
  /** Which fit mode, if any, is active — its menu entry shows checked. */
  fitMode: { type: String, default: "none" },
  disabled: { type: Boolean, default: false },
  /** Whether the menu offers Fit width / Fit page. A sheet has no page to
   * fit (SPEC-117 "Zoom"), so the workbook view passes false. */
  showFit: { type: Boolean, default: true },
});

const emit = defineEmits(["zoom-in", "zoom-out", "fit-width", "fit-page", "apply-input", "set-preset"]);

const rounded = computed(() => Math.round(props.zoomPercent));

const open = ref(false);
const editing = ref(false);
const draft = ref("");

// The displayed value: the live percentage while idle, the user's raw text
// while they edit.
const displayValue = computed(() => (editing.value ? draft.value : `${rounded.value}%`));

// Keep the idle display in sync when zoom changes from elsewhere (buttons,
// shortcuts, wheel, fit recompute).
watch(rounded, () => {
  if (!editing.value) {
    draft.value = `${rounded.value}%`;
  }
});

const triggerEl = useTemplateRef("triggerEl");
const menuSurface = useTemplateRef("menuSurface");

const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuSurface.value?.$el ?? null,
  align: "center",
  placement: "above",
});

function openMenu() {
  if (!props.disabled) {
    open.value = true;
  }
}

function onFocus() {
  editing.value = true;
  draft.value = `${rounded.value}`;
  triggerEl.value?.select();
}

function onInput() {
  draft.value = triggerEl.value?.value ?? "";
}

function commit() {
  emit("apply-input", draft.value);
  editing.value = false;
  draft.value = `${rounded.value}%`;
}

function onEnter() {
  commit();
  triggerEl.value?.blur();
}

function pickPreset(percent = 100) {
  emit("set-preset", percent);
  open.value = false;
}

function pickFitWidth() {
  emit("fit-width");
  open.value = false;
}

function pickFitPage() {
  emit("fit-page");
  open.value = false;
}

function isPresetChecked(preset = 0) {
  return props.fitMode === "none" && rounded.value === preset;
}
</script>

<template>
  <div
    class="zoom-control"
    data-testid="zoom-control"
    role="group"
    aria-label="Zoom"
  >
    <button
      type="button"
      class="zoom-step"
      data-testid="zoom-out"
      aria-label="Zoom out"
      :disabled="disabled"
      @click="emit('zoom-out')"
    >
      −
    </button>
    <input
      ref="triggerEl"
      class="zoom-percent"
      data-testid="zoom-percent"
      type="text"
      inputmode="numeric"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      aria-label="Zoom level"
      :disabled="disabled"
      :value="displayValue"
      @focus="onFocus"
      @input="onInput"
      @click="openMenu"
      @keydown.enter.prevent="onEnter"
      @blur="commit"
    >
    <button
      type="button"
      class="zoom-step"
      data-testid="zoom-in"
      aria-label="Zoom in"
      :disabled="disabled"
      @click="emit('zoom-in')"
    >
      +
    </button>
    <UiMenuSurface
      ref="menuSurface"
      :open="open"
      :anchor="triggerEl"
      :return-focus="false"
      class="zoom-menu"
      data-testid="zoom-menu"
      role="menu"
      aria-label="Zoom"
      :style="menuStyle"
      @dismiss="open = false"
    >
      <button
        v-for="preset in ZOOM_PRESETS"
        :key="preset"
        type="button"
        role="menuitemradio"
        class="zoom-option"
        :class="{ active: isPresetChecked(preset) }"
        :aria-checked="isPresetChecked(preset) ? 'true' : 'false'"
        :data-testid="`zoom-preset-${preset}`"
        @click="pickPreset(preset)"
      >
        {{ preset }}%
      </button>
      <template v-if="showFit">
        <div
          class="zoom-menu-separator"
          role="separator"
        />
        <button
          type="button"
          role="menuitemradio"
          class="zoom-option"
          :class="{ active: fitMode === 'fit-width' }"
          :aria-checked="fitMode === 'fit-width' ? 'true' : 'false'"
          data-testid="zoom-fit-width"
          @click="pickFitWidth"
        >
          Fit width
        </button>
        <button
          type="button"
          role="menuitemradio"
          class="zoom-option"
          :class="{ active: fitMode === 'fit-page' }"
          :aria-checked="fitMode === 'fit-page' ? 'true' : 'false'"
          data-testid="zoom-fit-page"
          @click="pickFitPage"
        >
          Fit page
        </button>
      </template>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.zoom-control {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  position: relative;
}

.zoom-step {
  font: inherit;
  font-size: var(--text-md);
  line-height: 1;
  width: 1.4em;
  height: 1.4em;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 0;
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: inherit;
  cursor: pointer;
}

.zoom-step:hover:not(:disabled),
.zoom-step:focus-visible {
  background: var(--state-hover);
}

.zoom-step:disabled {
  opacity: 0.5;
  cursor: default;
}

.zoom-percent {
  font: inherit;
  font-size: inherit;
  width: 3.5em;
  text-align: center;
  padding: 0 var(--space-1);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: inherit;
  cursor: pointer;
}

.zoom-percent:hover,
.zoom-percent:focus-visible {
  border-color: var(--border-subtle, var(--state-hover));
}

.zoom-menu {
  z-index: 30;
  min-width: 8rem;
}

.zoom-option {
  font: inherit;
  font-size: var(--text-md);
  display: block;
  width: 100%;
  padding: var(--space-2) var(--space-3);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  text-align: left;
}

.zoom-option:hover {
  background: var(--state-hover);
}

.zoom-option.active {
  background: var(--surface-selected);
}

.zoom-menu-separator {
  height: 1px;
  margin: var(--space-1) 0;
  background: var(--border-subtle, var(--state-hover));
}
</style>
