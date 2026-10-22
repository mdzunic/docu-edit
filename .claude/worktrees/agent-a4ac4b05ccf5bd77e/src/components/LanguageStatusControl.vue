<script setup lang="ts">
// The status bar's language control (SPEC-035 AC-15, AC-15.1): shows the
// current language, and opens a popover to change it — self-contained
// (button + popover in one component) the same way HighlightPicker.vue is,
// anchored with the same useFloatingMenu escape-the-clipped-strip pattern.
import { computed, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { asLanguageOptions } from "../lib/core/language";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

const props = defineProps({
  currentLanguageTag: { type: String, default: "" },
  currentLanguageLabel: { type: String, default: "" },
  languageOptions: { type: Array, default: () => [] },
  /** Whether a real (non-collapsed) text selection is active — "Do not
   * check spelling" only applies to a selection (SPEC-035 AC-11). */
  selectionActive: { type: Boolean, default: false },
  selectionNoProof: { type: Boolean, default: false },
});

const emit = defineEmits(["select", "manage-dictionary", "toggle-no-proof"]);

const options = asLanguageOptions(props.languageOptions);

const open = ref(false);
const triggerEl = useTemplateRef("triggerEl");
const menuSurface = useTemplateRef("menuSurface");

// The status bar sits at the bottom of the viewport, so the popover opens
// upward (placement: "above") — "below" (the default, right for toolbar
// buttons near the top) would open off-screen here.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuSurface.value?.$el ?? null,
  align: "center",
  placement: "above",
});

function toggle() {
  open.value = !open.value;
}

function pick(tag = "") {
  emit("select", tag);
  open.value = false;
}

const noProofChecked = computed({
  get: () => props.selectionNoProof,
  set: (value) => emit("toggle-no-proof", value),
});

function manageDictionary() {
  open.value = false;
  emit("manage-dictionary");
}

// Outside-click and Escape dismissal (and focus return) are owned by
// UiMenuSurface — the shared, tested implementation.
</script>

<template>
  <div class="language-control">
    <button
      ref="triggerEl"
      type="button"
      class="language-button"
      data-testid="language-status-button"
      aria-haspopup="true"
      :aria-expanded="open ? 'true' : 'false'"
      :aria-label="`Language: ${currentLanguageLabel}`"
      @click="toggle"
    >
      {{ currentLanguageLabel }}
    </button>
    <UiMenuSurface
      ref="menuSurface"
      :open="open"
      :anchor="triggerEl"
      class="language-menu"
      data-testid="language-menu"
      role="menu"
      aria-label="Language"
      :style="menuStyle"
      @dismiss="open = false"
    >
      <button
        v-for="option in options"
        :key="option.tag"
        type="button"
        role="menuitemradio"
        class="language-option"
        :class="{ active: option.tag === currentLanguageTag }"
        :aria-checked="option.tag === currentLanguageTag ? 'true' : 'false'"
        :data-testid="`language-option-${option.tag}`"
        @click="pick(option.tag)"
      >
        {{ option.label }}
      </button>
      <template v-if="selectionActive">
        <div
          class="language-menu-separator"
          role="separator"
        />
        <label class="no-proof-option">
          <input
            v-model="noProofChecked"
            type="checkbox"
            data-testid="language-no-proof"
          >
          Do not check spelling
        </label>
      </template>
      <div
        class="language-menu-separator"
        role="separator"
      />
      <button
        type="button"
        role="menuitem"
        class="language-option"
        data-testid="language-manage-dictionary"
        @click="manageDictionary"
      >
        Custom dictionary…
      </button>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.language-control {
  position: relative;
  display: inline-flex;
}

.language-button {
  font: inherit;
  font-size: inherit;
  padding: 0;
  border: none;
  background: transparent;
  color: inherit;
  cursor: pointer;
}

.language-button:hover,
.language-button:focus-visible {
  text-decoration: underline;
}

/* position/left/top come inline from useFloatingMenu; the box is UiMenuSurface's. */
.language-menu {
  z-index: 30;
  min-width: 10rem;
}

.language-option {
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

.language-option:hover {
  background: var(--state-hover);
}

.language-option.active {
  background: var(--surface-selected);
}

.language-menu-separator {
  height: 1px;
  margin: var(--space-2) 0;
  background: var(--border-subtle);
}

.no-proof-option {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  font-size: var(--text-md);
  color: var(--text-primary);
  cursor: pointer;
}
</style>
