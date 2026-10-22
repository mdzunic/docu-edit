<script setup lang="ts">
// The Format → Theme… gallery (SPEC-077 AC-10..AC-13, AC-19): the document's
// current theme first, marked "Current", then the six built-ins (Office
// first). Each card previews the open document — its leading paragraphs
// restyled with that card's colours and fonts — falling back to a specimen
// sheet (heading, two body lines, the six accent swatches) when nothing is
// laid out yet. "Customise…" edits a purely local draft of the twelve slot
// colours and the two fonts, with Reset returning to the theme as it was
// when the dialog opened. Nothing here ever writes to the model (AC-19):
// choosing a card or Apply emits `apply` for the parent's one command call;
// Cancel/Escape emits `cancel`. BaseDialog owns the shell, Escape, initial
// focus and the Tab trap (AC-13 — every control is a native button/input).
//
// This script is parsed by espree (see AppShell.vue), so props are runtime-
// declared and helper params are typed by default value, not annotation.
import { computed, ref } from "vue";
import { BUILT_IN_THEMES, DEFAULT_THEME, THEME_COLOR_SLOTS, themeSlotLabel } from "../lib/core/theme";
import { asDocumentTheme } from "../lib/documentModel";
import { asThemePreviewParagraphs } from "../views/document/useThemeGallery";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** documentThemeOf(doc) — the first card, marked current (AC-10). */
  currentTheme: { type: Object, required: true },
  /** ThemePreviewParagraph[] (useThemeGallery); [] shows the specimen. */
  previewParagraphs: { type: Array, default: () => [] },
});
const emit = defineEmits(["apply", "cancel"]);

/** The six accent slots, for the specimen's swatch row (AC-11). */
const ACCENT_SLOTS = THEME_COLOR_SLOTS.slice(4, 10);

const previewLines = computed(() => asThemePreviewParagraphs(props.previewParagraphs));

const cards = computed(() => [
  { key: "__current__", theme: asDocumentTheme(props.currentTheme), current: true },
  ...BUILT_IN_THEMES.map((theme) => ({ key: theme.name, theme, current: false })),
]);

function cloneTheme(theme = DEFAULT_THEME) {
  return { name: theme.name, colors: { ...theme.colors }, fonts: { ...theme.fonts } };
}

// --- Customise… (AC-12): a local draft, never the model (AC-19) -------------
const customising = ref(false);
/** The theme as it was when the dialog opened — what Reset returns to. */
const initial = cloneTheme(asDocumentTheme(props.currentTheme));
const draft = ref(cloneTheme(initial));

function resetDraft() {
  draft.value = cloneTheme(initial);
}

function applyCustomised() {
  emit("apply", cloneTheme(draft.value));
}

// --- per-card preview styling: local inline styles only (AC-19) -------------
function headingStyle(theme = DEFAULT_THEME) {
  return { color: theme.colors.accent1, fontFamily: theme.fonts.major };
}

function bodyStyle(theme = DEFAULT_THEME) {
  return { color: theme.colors.dk1, fontFamily: theme.fonts.minor };
}

function lineStyle(line = { text: "", heading: false }, theme = DEFAULT_THEME) {
  return line.heading ? headingStyle(theme) : bodyStyle(theme);
}

/** The card's accessible name (AC-13). */
function cardName(card = { theme: DEFAULT_THEME, current: false }) {
  return card.current ? `${card.theme.name} (current theme)` : `${card.theme.name} theme`;
}
</script>

<template>
  <BaseDialog
    title="Theme"
    testid="theme-gallery-dialog"
    width="44rem"
    @dismiss="emit('cancel')"
  >
    <div
      class="gallery-grid"
      data-testid="theme-gallery-grid"
    >
      <button
        v-for="card in cards"
        :key="card.key"
        type="button"
        class="theme-card"
        :data-testid="card.current ? 'theme-card-current' : 'theme-card'"
        :data-theme-name="card.theme.name"
        :aria-label="cardName(card)"
        @click="emit('apply', card.theme)"
      >
        <span class="theme-card-header">
          <span class="theme-card-title">{{ card.theme.name }}</span>
          <span
            v-if="card.current"
            class="current-badge"
            data-testid="theme-current-badge"
          >Current</span>
        </span>
        <span
          class="card-preview"
          aria-hidden="true"
          :style="{ backgroundColor: card.theme.colors.lt1 }"
        >
          <template v-if="previewLines.length > 0">
            <span
              v-for="(line, index) in previewLines"
              :key="index"
              class="preview-line"
              :class="{ 'preview-heading': line.heading }"
              :style="lineStyle(line, card.theme)"
            >{{ line.text }}</span>
          </template>
          <template v-else>
            <span
              class="preview-line preview-heading"
              :style="headingStyle(card.theme)"
            >Heading</span>
            <span
              class="preview-line"
              :style="bodyStyle(card.theme)"
            >Body text set in the theme's own face.</span>
            <span
              class="preview-line"
              :style="bodyStyle(card.theme)"
            >A second line completes the specimen.</span>
            <span class="swatch-row">
              <span
                v-for="slot in ACCENT_SLOTS"
                :key="slot"
                class="accent-swatch"
                data-testid="theme-accent-swatch"
                :style="{ backgroundColor: card.theme.colors[slot] }"
              />
            </span>
          </template>
        </span>
      </button>
    </div>
    <button
      type="button"
      class="customise-toggle"
      data-testid="theme-customise"
      :aria-expanded="customising ? 'true' : 'false'"
      @click="customising = !customising"
    >
      Customise…
    </button>
    <section
      v-if="customising"
      class="customise-pane"
      data-testid="theme-customise-pane"
      aria-label="Customise theme"
    >
      <div class="colour-grid">
        <label
          v-for="slot in THEME_COLOR_SLOTS"
          :key="slot"
          class="colour-field"
        >
          <span class="field-label">{{ themeSlotLabel(slot) }}</span>
          <input
            v-model="draft.colors[slot]"
            type="color"
            :data-testid="`theme-colour-${slot}`"
          >
        </label>
      </div>
      <div class="font-row">
        <label class="font-field">
          <span class="field-label">Heading font</span>
          <input
            v-model="draft.fonts.major"
            type="text"
            data-testid="theme-font-major"
          >
        </label>
        <label class="font-field">
          <span class="field-label">Body font</span>
          <input
            v-model="draft.fonts.minor"
            type="text"
            data-testid="theme-font-minor"
          >
        </label>
      </div>
      <div class="customise-actions">
        <button
          type="button"
          data-testid="theme-reset"
          @click="resetDraft"
        >
          Reset
        </button>
        <button
          type="button"
          class="primary-button"
          data-testid="theme-apply-custom"
          @click="applyCustomised"
        >
          Apply customised theme
        </button>
      </div>
    </section>
    <div class="dialog-actions">
      <button
        type="button"
        data-testid="theme-gallery-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
    </div>
  </BaseDialog>
</template>

<style scoped>
/* Fixed-width panel with the tall-dialog scroll cap (template gallery's). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

.gallery-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(11rem, 1fr));
  gap: var(--space-4);
}

.theme-card {
  font: inherit;
  text-align: left;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  cursor: pointer;
}

.theme-card:hover,
.theme-card:focus-visible {
  border-color: var(--surface-brand);
  background: var(--surface-raised);
}

.theme-card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-2);
}

.theme-card-title {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.current-badge {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  color: var(--text-inverse);
  background: var(--surface-brand);
  border-radius: var(--radius-sm);
  padding: 0 var(--space-2);
}

/* The miniature is decorative: never interactive, never announced. */
.card-preview {
  pointer-events: none;
  user-select: none;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  min-height: 6.5rem;
  max-height: 6.5rem;
  overflow: hidden;
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.preview-line {
  /* Specimen text at miniature scale: a fixed fraction of the smallest
   * type token, so the sheet shrinks and grows with the app's type ramp. */
  font-size: calc(var(--text-xs) * 0.8);
  line-height: calc(1.3em);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.preview-heading {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.swatch-row {
  display: flex;
  gap: var(--space-1);
  margin-top: auto;
}

.accent-swatch {
  width: 0.9rem;
  height: 0.9rem;
  border-radius: var(--radius-sm);
}

.customise-toggle {
  font: inherit;
  margin-top: var(--space-5);
  padding: var(--space-2) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.customise-pane {
  margin-top: var(--space-4);
  padding: var(--space-4);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
}

.colour-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(9rem, 1fr));
  gap: var(--space-3);
}

.colour-field,
.font-field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-2);
  font-size: var(--text-md);
}

.colour-field input {
  inline-size: 3rem;
  block-size: 1.8rem;
  padding: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: none;
  cursor: pointer;
}

.font-row {
  display: flex;
  gap: var(--space-4);
  margin-top: var(--space-4);
}

.font-field {
  flex: 1;
}

.font-field input {
  font: inherit;
  flex: 1;
  padding: var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.customise-actions {
  display: flex;
  justify-content: flex-end;
  gap: var(--space-3);
  margin-top: var(--space-4);
}

.customise-actions button,
.dialog-actions button {
  font: inherit;
  padding: var(--space-2) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.primary-button {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-inverse);
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  margin-top: var(--space-5);
}
</style>
