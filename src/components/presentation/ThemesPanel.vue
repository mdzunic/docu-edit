<script setup lang="ts">
// The Design (deck theme) panel (SPEC-094): a gallery of theme cards — the
// deck's current theme first and marked, then the six built-ins, each drawing
// its own accents as swatches — plus Apply and a Customise disclosure with the
// twelve scheme-colour inputs, the two font inputs and Reset. All logic lives
// in the `useDeckTheme` controller passed in; this component is presentation +
// event forwarding only. Runtime `defineProps`, no emits (espree bars the
// generic form and the controller carries every command).
import UiButton from "../ui/UiButton.vue";
import { THEME_COLOR_SLOTS, themeSlotLabel } from "../../lib/core/theme";

const props = defineProps({
  /** The useDeckTheme() controller object (refs + command functions). */
  controller: { type: Object, required: true },
});

const c = props.controller;
const slots = THEME_COLOR_SLOTS;

function onSlotInput(slot = "dk1", event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setSchemeColor(slot, el.value);
}
function onFont(which = "major", event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setFont(which, el.value);
}
</script>

<template>
  <section
    v-if="c.open.value"
    class="themes-panel"
    data-testid="slides-theme-panel"
    aria-label="Design"
  >
    <header class="panel-header">
      <h3>Design</h3>
      <p
        class="current-name"
        data-testid="theme-current-name"
      >
        Theme: {{ c.themeName.value }}<span v-if="c.isCustom.value"> (customised)</span>
      </p>
    </header>

    <div
      class="theme-gallery"
      data-testid="theme-gallery"
      role="listbox"
      aria-label="Themes"
    >
      <button
        v-for="card in c.gallery.value"
        :key="card.id"
        type="button"
        class="theme-card"
        :class="{ 'theme-current': card.current, 'is-selected': card.id === c.selectedId.value }"
        data-testid="theme-card"
        role="option"
        :aria-selected="card.id === c.selectedId.value"
        :data-theme-id="card.id"
        :data-current="card.current ? 'true' : undefined"
        @click="c.select(card.id)"
        @dblclick="c.apply(card.id)"
        @mouseenter="c.hoverPreview(card.id)"
        @mouseleave="c.clearPreview()"
        @focus="c.hoverPreview(card.id)"
        @blur="c.clearPreview()"
      >
        <span class="theme-name">{{ card.name }}</span>
        <span
          class="theme-accents"
          aria-hidden="true"
        >
          <span
            v-for="(accent, i) in card.accents"
            :key="i"
            class="accent-swatch"
            :style="{ background: accent }"
          />
        </span>
      </button>
    </div>

    <div class="actions">
      <UiButton
        data-testid="theme-apply"
        @click="c.apply()"
      >
        Apply theme
      </UiButton>
    </div>

    <p
      v-if="c.optedOutSlides.value.length > 0"
      class="helper"
      data-testid="theme-opted-out-note"
    >
      Slide{{ c.optedOutSlides.value.length > 1 ? "s" : "" }} {{ c.optedOutSlides.value.join(", ") }}
      {{ c.optedOutSlides.value.length > 1 ? "carry" : "carries" }} a colour mapping this editor can't
      read and {{ c.optedOutSlides.value.length > 1 ? "are" : "is" }} opted out of theme colours — those
      slides keep their exact colours.
    </p>

    <p
      v-if="c.lossNote.value"
      class="helper"
      data-testid="theme-loss-note"
    >
      {{ c.lossNote.value }}
    </p>

    <details class="customise">
      <summary data-testid="theme-customise">Customise</summary>
      <div class="customise-grid">
        <label
          v-for="slot in slots"
          :key="slot"
          class="slot-field"
        >
          <span>{{ themeSlotLabel(slot) }}</span>
          <input
            type="color"
            :value="c.theme.value.colors[slot]"
            :data-testid="`theme-color-${slot}`"
            :aria-label="`${themeSlotLabel(slot)} colour`"
            @change="onSlotInput(slot, $event)"
          >
        </label>
      </div>
      <div class="font-fields">
        <label class="slot-field">
          <span>Heading font</span>
          <input
            type="text"
            :value="c.theme.value.fonts.major"
            data-testid="theme-font-major"
            aria-label="Heading font"
            @change="onFont('major', $event)"
          >
        </label>
        <label class="slot-field">
          <span>Body font</span>
          <input
            type="text"
            :value="c.theme.value.fonts.minor"
            data-testid="theme-font-minor"
            aria-label="Body font"
            @change="onFont('minor', $event)"
          >
        </label>
      </div>
      <div class="actions">
        <UiButton
          data-testid="theme-reset"
          :disabled="!c.canReset.value"
          @click="c.reset()"
        >
          Reset
        </UiButton>
      </div>
    </details>
  </section>
</template>

<style scoped>
.themes-panel {
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised);
  padding: var(--space-3) var(--space-4);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  overflow: auto;
  max-height: 40vh;
}

.panel-header {
  display: flex;
  align-items: baseline;
  gap: var(--space-3);
}

.panel-header h3 {
  margin: 0;
}

.current-name {
  margin: 0;
  color: var(--text-secondary);
}

.theme-gallery {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.theme-card {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  align-items: flex-start;
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.theme-card.is-selected {
  border-color: var(--accent);
  outline: var(--border-width) solid var(--accent);
}

.theme-card.theme-current .theme-name::after {
  content: " ✓";
  color: var(--accent);
}

.theme-name {
  font-weight: var(--weight-medium);
}

.theme-accents {
  display: inline-flex;
  gap: 2px;
}

.accent-swatch {
  width: 14px;
  height: 14px;
  border-radius: 2px;
  border: var(--border-width) solid var(--border-subtle);
}

.actions {
  display: flex;
  gap: var(--space-3);
}

.helper {
  margin: 0;
  color: var(--text-secondary);
  font-style: italic;
}

.customise summary {
  cursor: pointer;
  font-weight: var(--weight-medium);
}

.customise-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr));
  gap: var(--space-2);
  margin-top: var(--space-2);
}

.slot-field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-2);
}

.font-fields {
  display: flex;
  gap: var(--space-4);
  margin: var(--space-2) 0;
}
</style>
