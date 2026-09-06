<script setup lang="ts">
import { ref } from "vue";
import AppIcon from "./AppIcon.vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import UiField from "./ui/UiField.vue";
import UiSelect from "./ui/UiSelect.vue";
import UiToggleButton from "./ui/UiToggleButton.vue";
import { usePreferences } from "../composables/usePreferences";
import {
  AUTOSAVE_DELAY_CHOICES,
  MARGIN_PRESET_NAMES,
  PAGE_SIZE_PRESET_NAMES,
  PALETTE_OPTIONS,
} from "../lib/preferences";
import { SUPPORTED_LANGUAGES } from "../lib/core/language";
import { STANDARD_FONT_FAMILIES } from "../lib/core/fonts";
import { COMMON_FONT_SIZES } from "../lib/fonts";
import { ZOOM_PRESETS } from "../lib/zoom";
import { parseLengthToCm, formatLength, asMeasurementUnit } from "../lib/core/units";

// The Preferences dialog (SPEC-066). Every control reads the live reactive
// preferences and, on change, writes straight through usePreferences —
// there is no OK/Apply step and the Close button just closes (AC-7, AC-8).
// Built on BaseDialog (shell, Escape, Tab trap, initial focus) with SPEC-058's
// UiField/UiSelect primitives; only app-level settings appear here — nothing
// per-document (AC-23).
const emit = defineEmits(["close"]);

const { preferences, setPreference, resetPreferences } = usePreferences();

// Confirmation gate for the destructive reset (AC-28).
const confirmingReset = ref(false);

// The theme switch is a segmented icon control (system / sun / moon).
const themeOptions = [
  { value: "system", label: "System", icon: "system" },
  { value: "light", label: "Light", icon: "sun" },
  { value: "dark", label: "Dark", icon: "moon" },
];
const paletteOptions = PALETTE_OPTIONS;
const unitOptions = [
  { value: "cm", label: "Centimetres (cm)" },
  { value: "mm", label: "Millimetres (mm)" },
  { value: "in", label: "Inches (in)" },
  { value: "pt", label: "Points (pt)" },
];
const zoomOptions = [
  ...ZOOM_PRESETS.map((percent) => ({ value: String(percent), label: `${percent}%` })),
  { value: "fit-width", label: "Fit width" },
  { value: "fit-page", label: "Fit page" },
];
const viewModeOptions = [
  { value: "print", label: "Print" },
  { value: "web", label: "Web" },
];
const languageOptions = SUPPORTED_LANGUAGES.map((option) => ({ value: option.tag, label: option.label }));
const fontOptions = STANDARD_FONT_FAMILIES.map((family) => ({ value: family, label: family }));
const fontSizeOptions = COMMON_FONT_SIZES.map((size) => ({ value: String(size), label: String(size) }));
const pageSizeOptions = PAGE_SIZE_PRESET_NAMES.map((name) => ({ value: name, label: name }));
const marginOptions = MARGIN_PRESET_NAMES.map((name) => ({ value: name, label: name }));
const autosaveDelayOptions = AUTOSAVE_DELAY_CHOICES.map((ms) => ({
  value: String(ms),
  label: `${ms / 1000} s`,
}));
const downloadFormatOptions = [
  { value: "source", label: "Same as source" },
  { value: "docx", label: ".docx" },
  { value: "odt", label: ".odt" },
];

// The .vue script is parsed by espree (no TS-only syntax): these helpers
// narrow an event and coerce a select value with runtime checks rather than
// casts, and every setPreference call passes a literal key inline so its
// value type is inferred without an annotation.
function checked(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement && target.checked;
}

function textValue(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement ? target.value : "";
}

/** Coerces a zoom select value: the two fit modes pass through, everything
 * else is a percentage number. */
function zoomValue(value = "") {
  if (value === "fit-width" || value === "fit-page") {
    return value;
  }
  return Number(value);
}

/** Narrows a theme segment value to the preference union without a cast
 * (espree bars TS-only syntax in .vue scripts). */
function themeValue(value = "") {
  return value === "light" ? "light" : value === "dark" ? "dark" : "system";
}

function confirmReset() {
  resetPreferences();
  confirmingReset.value = false;
}

// The snap-to-grid spacing shows and parses in the current measurement unit
// (SPEC-072 AC-34) while the preference keeps centimetres.
function gridSpacingDisplay() {
  return formatLength(preferences.gridSpacingCm, asMeasurementUnit(preferences.measurementUnit));
}

function commitGridSpacing(event = new Event("change")) {
  const cm = parseLengthToCm(textValue(event), asMeasurementUnit(preferences.measurementUnit));
  if (cm !== null) {
    setPreference("gridSpacingCm", cm);
  }
}
</script>

<template>
  <BaseDialog
    title="Preferences"
    testid="preferences-dialog"
    width="34rem"
    @dismiss="emit('close')"
  >
    <template #header-extra>
      <UiButton
        variant="ghost"
        size="sm"
        data-testid="preferences-close-x"
        aria-label="Close preferences"
        autofocus
        @click="emit('close')"
      >
        ✕
      </UiButton>
    </template>

    <!-- Appearance -->
    <section
      class="section"
      aria-labelledby="preferences-appearance-heading"
    >
      <h3
        id="preferences-appearance-heading"
        class="section-title"
      >
        <AppIcon name="styles" />
        Appearance
      </h3>
      <UiField
        label="Theme"
        hint="Default: System"
      >
        <div
          class="theme-segment"
          role="group"
          aria-label="Theme"
          data-testid="preferences-theme"
        >
          <UiToggleButton
            v-for="option in themeOptions"
            :key="option.value"
            :icon="option.icon"
            :label="option.label"
            :data-testid="`preferences-theme-${option.value}`"
            :pressed="preferences.theme === option.value"
            @click="setPreference('theme', themeValue(option.value))"
          >
            {{ option.label }}
          </UiToggleButton>
        </div>
      </UiField>
      <UiField
        class="field-wide"
        label="Color theme"
        hint="Each works in light and dark. Default: Modernist"
      >
        <div
          class="palette-grid"
          role="group"
          aria-label="Color theme"
          data-testid="preferences-palette"
        >
          <button
            v-for="option in paletteOptions"
            :key="option.value"
            type="button"
            class="palette-chip"
            :data-testid="`preferences-palette-${option.value}`"
            :class="{ 'is-active': preferences.palette === option.value }"
            :aria-pressed="preferences.palette === option.value"
            @click="setPreference('palette', option.value)"
          >
            <span
              class="palette-swatch"
              :style="{ background: option.swatch }"
              aria-hidden="true"
            />
            {{ option.label }}
          </button>
        </div>
      </UiField>
      <UiField
        label="Toolbar labels"
        for="preferences-toolbar-labels"
        hint="Prints the word under each toolbar icon. Dropped automatically when the strip runs out of room. Default: on"
      >
        <input
          id="preferences-toolbar-labels"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-toolbar-labels"
          aria-describedby="preferences-toolbar-labels-hint"
          :checked="preferences.toolbarLabels"
          @change="(event) => setPreference('toolbarLabels', checked(event))"
        >
      </UiField>
      <UiField
        label="Dim page in dark theme"
        for="preferences-dim-page"
        hint="Tones down the bright page against a dark UI. Default: off"
      >
        <input
          id="preferences-dim-page"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-dim-page"
          aria-describedby="preferences-dim-page-hint"
          :checked="preferences.dimPage"
          @change="(event) => setPreference('dimPage', checked(event))"
        >
      </UiField>
      <UiField
        label="Measurement unit"
        for="preferences-measurement-unit"
        hint="Units shown on rulers and in Page setup. Default: cm"
      >
        <UiSelect
          id="preferences-measurement-unit"
          data-testid="preferences-measurement-unit"
          aria-describedby="preferences-measurement-unit-hint"
          :model-value="preferences.measurementUnit"
          :options="unitOptions"
          @update:model-value="(value) => setPreference('measurementUnit', value)"
        />
      </UiField>
      <UiField
        label="Default zoom for new documents"
        for="preferences-default-zoom"
        hint="Default: 100%"
      >
        <UiSelect
          id="preferences-default-zoom"
          data-testid="preferences-default-zoom"
          aria-describedby="preferences-default-zoom-hint"
          :model-value="String(preferences.defaultZoom)"
          :options="zoomOptions"
          @update:model-value="(value) => setPreference('defaultZoom', zoomValue(value))"
        />
      </UiField>
      <UiField
        label="Default view mode"
        for="preferences-default-view-mode"
        hint="Print shows pages; Web is a continuous column. Default: Print"
      >
        <UiSelect
          id="preferences-default-view-mode"
          data-testid="preferences-default-view-mode"
          aria-describedby="preferences-default-view-mode-hint"
          :model-value="preferences.defaultViewMode"
          :options="viewModeOptions"
          @update:model-value="(value) => setPreference('defaultViewMode', value)"
        />
      </UiField>
      <UiField
        label="Show rulers by default"
        for="preferences-show-ruler"
        hint="New documents open with both rulers shown. Default: on"
      >
        <input
          id="preferences-show-ruler"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-show-ruler"
          aria-describedby="preferences-show-ruler-hint"
          :checked="preferences.showRuler"
          @change="(event) => setPreference('showRuler', checked(event))"
        >
      </UiField>
      <UiField
        label="Show formatting marks by default"
        for="preferences-show-formatting-marks"
        hint="Pilcrows, spaces and tabs shown in new documents. Default: off"
      >
        <input
          id="preferences-show-formatting-marks"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-show-formatting-marks"
          aria-describedby="preferences-show-formatting-marks-hint"
          :checked="preferences.showFormattingMarks"
          @change="(event) => setPreference('showFormattingMarks', checked(event))"
        >
      </UiField>
    </section>

    <!-- Editing -->
    <section
      class="section"
      aria-labelledby="preferences-editing-heading"
    >
      <h3
        id="preferences-editing-heading"
        class="section-title"
      >
        <AppIcon name="pencil" />
        Editing
      </h3>
      <UiField
        label="Smart quotes"
        for="preferences-smart-quotes"
        hint="Straight quotes become curly as you type. Default: on"
      >
        <input
          id="preferences-smart-quotes"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-smart-quotes"
          aria-describedby="preferences-smart-quotes-hint"
          :checked="preferences.smartQuotes"
          @change="(event) => setPreference('smartQuotes', checked(event))"
        >
      </UiField>
      <UiField
        label="Em and en dashes"
        for="preferences-smart-dashes"
        hint="Typing -- becomes an en or em dash. Default: on"
      >
        <input
          id="preferences-smart-dashes"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-smart-dashes"
          aria-describedby="preferences-smart-dashes-hint"
          :checked="preferences.smartDashes"
          @change="(event) => setPreference('smartDashes', checked(event))"
        >
      </UiField>
      <UiField
        label="Check spelling as you type"
        for="preferences-spell-check"
        hint="Underlines misspellings while you edit. Default: on"
      >
        <input
          id="preferences-spell-check"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-spell-check"
          aria-describedby="preferences-spell-check-hint"
          :checked="preferences.spellCheck"
          @change="(event) => setPreference('spellCheck', checked(event))"
        >
      </UiField>
      <UiField
        label="Default document language"
        for="preferences-default-language"
        hint="Language new documents are created in. Default: English (US)"
      >
        <UiSelect
          id="preferences-default-language"
          data-testid="preferences-default-language"
          aria-describedby="preferences-default-language-hint"
          :model-value="preferences.defaultLanguage"
          :options="languageOptions"
          @update:model-value="(value) => setPreference('defaultLanguage', value)"
        />
      </UiField>
      <UiField
        label="Author name for comments and tracked changes"
        for="preferences-author-name"
        hint="Attributed to your comments and edits. Default: You"
      >
        <input
          id="preferences-author-name"
          type="text"
          class="pref-text"
          data-testid="preferences-author-name"
          aria-describedby="preferences-author-name-hint"
          :value="preferences.authorName"
          @change="(event) => setPreference('authorName', textValue(event))"
        >
      </UiField>
    </section>

    <!-- New documents -->
    <section
      class="section"
      aria-labelledby="preferences-new-documents-heading"
    >
      <h3
        id="preferences-new-documents-heading"
        class="section-title"
      >
        <AppIcon name="text-box" />
        New documents
      </h3>
      <UiField
        label="Default font family"
        for="preferences-default-font"
        hint="Default: Calibri"
      >
        <UiSelect
          id="preferences-default-font"
          data-testid="preferences-default-font"
          aria-describedby="preferences-default-font-hint"
          :model-value="preferences.defaultFontFamily"
          :options="fontOptions"
          @update:model-value="(value) => setPreference('defaultFontFamily', value)"
        />
      </UiField>
      <UiField
        label="Default font size"
        for="preferences-default-font-size"
        hint="Point size for new documents. Default: 11"
      >
        <UiSelect
          id="preferences-default-font-size"
          data-testid="preferences-default-font-size"
          aria-describedby="preferences-default-font-size-hint"
          :model-value="String(preferences.defaultFontSizePt)"
          :options="fontSizeOptions"
          @update:model-value="(value) => setPreference('defaultFontSizePt', Number(value))"
        />
      </UiField>
      <UiField
        label="Default page size"
        for="preferences-default-page-size"
        hint="Default: A4"
      >
        <UiSelect
          id="preferences-default-page-size"
          data-testid="preferences-default-page-size"
          aria-describedby="preferences-default-page-size-hint"
          :model-value="preferences.defaultPageSizeId"
          :options="pageSizeOptions"
          @update:model-value="(value) => setPreference('defaultPageSizeId', value)"
        />
      </UiField>
      <UiField
        label="Default margins"
        for="preferences-default-margins"
        hint="Margins new documents start with. Default: Normal"
      >
        <UiSelect
          id="preferences-default-margins"
          data-testid="preferences-default-margins"
          aria-describedby="preferences-default-margins-hint"
          :model-value="preferences.defaultMarginsId"
          :options="marginOptions"
          @update:model-value="(value) => setPreference('defaultMarginsId', value)"
        />
      </UiField>
    </section>

    <!-- Saving -->
    <section
      class="section"
      aria-labelledby="preferences-saving-heading"
    >
      <h3
        id="preferences-saving-heading"
        class="section-title"
      >
        <AppIcon name="download" />
        Saving
      </h3>
      <UiField
        label="Autosave"
        for="preferences-autosave"
        hint="Off still keeps crash-recovery drafts — only the committed save stops. Default: on"
      >
        <input
          id="preferences-autosave"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-autosave"
          aria-describedby="preferences-autosave-hint"
          :checked="preferences.autosave"
          @change="(event) => setPreference('autosave', checked(event))"
        >
      </UiField>
      <UiField
        label="Autosave delay"
        for="preferences-autosave-delay"
        hint="How long after a change a save fires. Default: 2 s"
      >
        <UiSelect
          id="preferences-autosave-delay"
          data-testid="preferences-autosave-delay"
          aria-describedby="preferences-autosave-delay-hint"
          :model-value="String(preferences.autosaveDelayMs)"
          :options="autosaveDelayOptions"
          @update:model-value="(value) => setPreference('autosaveDelayMs', Number(value))"
        />
      </UiField>
      <UiField
        label="Keep version history"
        for="preferences-keep-version-history"
        hint="Record periodic versions you can restore. Default: on"
      >
        <input
          id="preferences-keep-version-history"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-keep-version-history"
          aria-describedby="preferences-keep-version-history-hint"
          :checked="preferences.keepVersionHistory"
          @change="(event) => setPreference('keepVersionHistory', checked(event))"
        >
      </UiField>
      <UiField
        label="Default download format"
        for="preferences-download-format"
        hint="Format the one-click “Download a copy” links save in. Default: Same as source"
      >
        <UiSelect
          id="preferences-download-format"
          data-testid="preferences-download-format"
          aria-describedby="preferences-download-format-hint"
          :model-value="preferences.defaultDownloadFormat"
          :options="downloadFormatOptions"
          @update:model-value="(value) => setPreference('defaultDownloadFormat', value)"
        />
      </UiField>
    </section>

    <!-- Drawing (SPEC-072) -->
    <section
      class="section"
      aria-labelledby="preferences-drawing-heading"
    >
      <h3
        id="preferences-drawing-heading"
        class="section-title"
      >
        <AppIcon name="shape" />
        Drawing
      </h3>
      <UiField
        label="Snap to guides"
        for="preferences-snap-guides"
        hint="Snap a dragged drawing to alignment guides. Default: on"
      >
        <input
          id="preferences-snap-guides"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-snap-guides"
          :checked="preferences.snapToGuides"
          @change="(event) => setPreference('snapToGuides', checked(event))"
        >
      </UiField>
      <UiField
        label="Snap to grid"
        for="preferences-snap-grid"
        hint="Snap a dragged drawing to the grid. Default: off"
      >
        <input
          id="preferences-snap-grid"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-snap-grid"
          :checked="preferences.snapToGrid"
          @change="(event) => setPreference('snapToGrid', checked(event))"
        >
      </UiField>
      <UiField
        label="Grid spacing"
        for="preferences-grid-spacing"
        :hint="`Distance between grid lines, in ${preferences.measurementUnit}. Default: 0.25cm`"
      >
        <input
          id="preferences-grid-spacing"
          type="text"
          inputmode="decimal"
          class="pref-text"
          data-testid="preferences-grid-spacing"
          :value="gridSpacingDisplay()"
          @change="commitGridSpacing"
        >
      </UiField>
    </section>

    <!-- Advanced -->
    <section
      class="section"
      aria-labelledby="preferences-advanced-heading"
    >
      <h3
        id="preferences-advanced-heading"
        class="section-title"
      >
        <AppIcon name="settings" />
        Advanced
      </h3>
      <UiField
        label="Show the conversion report after a cross-format save"
        for="preferences-conversion-report"
        hint="Warns what a .docx↔.odt save could not carry over. Default: on"
      >
        <input
          id="preferences-conversion-report"
          type="checkbox"
          class="pref-checkbox"
          data-testid="preferences-conversion-report"
          aria-describedby="preferences-conversion-report-hint"
          :checked="preferences.showConversionReport"
          @change="(event) => setPreference('showConversionReport', checked(event))"
        >
      </UiField>
      <div class="reset-row">
        <template v-if="!confirmingReset">
          <UiButton
            variant="danger"
            data-testid="preferences-reset"
            @click="confirmingReset = true"
          >
            Reset all preferences
          </UiButton>
          <p class="reset-help">
            Restores every setting above to its default. Does not touch your documents,
            drafts, dictionary or version history.
          </p>
        </template>
        <template v-else>
          <p
            class="reset-confirm-text"
            role="alert"
          >
            Reset every preference to its default?
          </p>
          <div class="reset-confirm-actions">
            <UiButton
              variant="secondary"
              data-testid="preferences-reset-cancel"
              @click="confirmingReset = false"
            >
              Cancel
            </UiButton>
            <UiButton
              variant="danger"
              data-testid="preferences-reset-confirm"
              @click="confirmReset"
            >
              Reset
            </UiButton>
          </div>
        </template>
      </div>
    </section>

    <template #actions>
      <UiButton
        variant="primary"
        data-testid="preferences-close"
        @click="emit('close')"
      >
        Close
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* The dialog is tall; cap it and scroll the body (BaseDialog does not). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

.section {
  margin-bottom: var(--space-5);
  padding-bottom: var(--space-4);
  border-bottom: var(--border-width) solid var(--border-default);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
}

/* ---- Compact settings rows -----------------------------------------------
   Each field is a settings-app row: label (and its hint beneath) on the left,
   the control flush right — instead of the stacked label/control/hint column,
   which made the dialog twice as tall. The wide controls (the palette grid)
   opt out via .field-wide and keep the stacked layout. */
.section :deep(.ui-field) {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: center;
  column-gap: var(--space-5);
  row-gap: var(--space-1);
}

.section :deep(.ui-field > .ui-field-label) {
  grid-column: 1;
  grid-row: 1;
  color: var(--text-primary);
  font-size: var(--text-md);
}

.section :deep(.ui-field > .ui-field-hint),
.section :deep(.ui-field > .ui-field-error) {
  grid-column: 1;
  grid-row: 2;
}

.section :deep(.ui-field > *:not(.ui-field-label):not(.ui-field-hint):not(.ui-field-error)) {
  grid-column: 2;
  grid-row: 1 / span 2;
  justify-self: end;
  align-self: center;
}

.section :deep(.field-wide.ui-field) {
  display: flex;
  flex-direction: column;
  align-items: stretch;
}

.section :deep(.field-wide.ui-field > *:not(.ui-field-label):not(.ui-field-hint):not(.ui-field-error)) {
  align-self: stretch;
  justify-self: stretch;
}

/* ---- The theme segment (system / sun / moon) ------------------------------ */
.theme-segment {
  display: inline-flex;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  overflow: hidden;
}

.theme-segment :deep(.ui-toggle-button) {
  border-radius: 0;
}

.theme-segment :deep(.ui-toggle-button + .ui-toggle-button) {
  border-inline-start: var(--border-width) solid var(--border-default);
}

/* ---- The palette chips ---------------------------------------------------- */
.palette-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(9rem, 1fr));
  gap: var(--space-2);
}

.palette-chip {
  font: inherit;
  font-size: var(--text-md);
  display: flex;
  align-items: center;
  gap: var(--space-2);
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  text-align: start;
  cursor: pointer;
}

.palette-chip:hover {
  background: var(--state-hover);
}

.palette-chip.is-active {
  border-color: var(--accent);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

/* The dominant-colour square; its fill is data (PALETTE_OPTIONS.swatch),
   bound inline — palettes are not chrome tokens. */
.palette-swatch {
  flex: none;
  width: var(--space-5);
  height: var(--space-5);
  border-radius: var(--radius-sm);
  border: var(--border-width) solid var(--border-subtle);
}

.pref-checkbox {
  width: 1rem;
  height: 1rem;
  accent-color: var(--accent);
}

.pref-text {
  height: var(--height-control);
  padding: 0 var(--space-3);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-control);
}

.pref-text:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.reset-row {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.reset-help,
.reset-confirm-text {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.reset-confirm-text {
  color: var(--text-danger);
}

.reset-confirm-actions {
  display: flex;
  gap: var(--space-3);
}
</style>
