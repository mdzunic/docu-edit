<script setup lang="ts">
// The workbook status bar's number-locale control (SPEC-118 "The
// preference"): shows the locale's tag and opens a menu of one
// `menuitemradio` per locale (`sheet-locale-<tag>`), each labelled with its
// name and a live sample. Choosing one writes the preference through
// useSpreadsheetLocale — every open sheet re-renders, nothing is marked dirty
// and no stored or downloaded byte changes. Self-contained (button + popover
// in one component) like LanguageStatusControl.vue, opening upward with
// useFloatingMenu because the status bar sits at the bottom of the viewport.
import { ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../../composables/useFloatingMenu";
import { isSpreadsheetLocaleTag, SPREADSHEET_LOCALE_TAGS } from "../../lib/spreadsheet/numberFormat/locale";
import { localeOptionLabel } from "../../lib/spreadsheet/numberFormat/render";
import { useSpreadsheetLocale } from "../../views/spreadsheet/useSpreadsheetLocale";
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const { tag, setLocale } = useSpreadsheetLocale();
const options = SPREADSHEET_LOCALE_TAGS.map((value) => ({ tag: value, label: localeOptionLabel(value) }));

const open = ref(false);
const triggerEl = useTemplateRef("triggerEl");
const menuSurface = useTemplateRef("menuSurface");

const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => triggerEl.value,
  menu: () => menuSurface.value?.$el ?? null,
  align: "center",
  placement: "above",
});

function pick(value = "") {
  if (isSpreadsheetLocaleTag(value)) {
    setLocale(value);
  }
  open.value = false;
}
</script>

<template>
  <div class="sheet-locale-control">
    <button
      ref="triggerEl"
      type="button"
      class="sheet-locale-button"
      data-testid="sheet-locale"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :aria-label="`Number and date format: ${tag}`"
      title="Number and date format"
      @click="open = !open"
    >
      {{ tag }}
    </button>
    <UiMenuSurface
      ref="menuSurface"
      :open="open"
      :anchor="triggerEl"
      class="sheet-locale-menu"
      data-testid="sheet-locale-menu"
      role="menu"
      aria-label="Number and date format"
      :style="menuStyle"
      @dismiss="open = false"
    >
      <button
        v-for="option in options"
        :key="option.tag"
        type="button"
        role="menuitemradio"
        class="sheet-locale-option"
        :class="{ active: option.tag === tag }"
        :aria-checked="option.tag === tag ? 'true' : 'false'"
        :data-testid="`sheet-locale-${option.tag}`"
        @click="pick(option.tag)"
      >
        {{ option.label }}
      </button>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.sheet-locale-control {
  position: relative;
  display: inline-flex;
}

.sheet-locale-button {
  font: inherit;
  font-size: inherit;
  padding: 0;
  border: none;
  background: transparent;
  color: inherit;
  cursor: pointer;
  white-space: nowrap;
}

.sheet-locale-button:hover,
.sheet-locale-button:focus-visible {
  text-decoration: underline;
}

/* position/left/top come inline from useFloatingMenu; the box is UiMenuSurface's. */
.sheet-locale-menu {
  z-index: 30;
  min-width: 18rem;
}

.sheet-locale-option {
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
  white-space: nowrap;
}

.sheet-locale-option:hover {
  background: var(--state-hover);
}

.sheet-locale-option.active {
  background: var(--surface-selected);
}
</style>
