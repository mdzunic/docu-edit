<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import UiField from "./ui/UiField.vue";
import UiSelect from "./ui/UiSelect.vue";
import { usePreferences } from "../composables/usePreferences";
import {
  AUTOSAVE_DELAY_CHOICES,
  MARGIN_PRESET_NAMES,
  PAGE_SIZE_PRESET_NAMES,
} from "../lib/preferences";
import { SUPPORTED_LANGUAGES } from "../lib/core/language";
import { STANDARD_FONT_FAMILIES } from "../lib/core/fonts";
import { COMMON_FONT_SIZES } from "../lib/fonts";
import { ZOOM_PRESETS } from "../lib/zoom";

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

const themeOptions = [
  { value: "system", label: "System" },
  { value: "light", label: "Light" },
  { value: "dark", label: "Dark" },
];
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

function confirmReset() {
  resetPreferences();
  confirmingReset.value = false;
}
</script>

<template>
  <BaseDialog
    title="Preferences"
    testid="preferences-dialog"
    width="30rem"
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
        Appearance
      </h3>
      <UiField
        label="Theme"
        for="preferences-theme"
        hint="How the app chrome looks. Default: System"
      >
        <UiSelect
          id="preferences-theme"
          data-testid="preferences-theme"
          aria-describedby="preferences-theme-hint"
          :model-value="preferences.theme"
          :options="themeOptions"
          @update:model-value="(value) => setPreference('theme', value)"
        />
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

    <!-- Advanced -->
    <section
      class="section"
      aria-labelledby="preferences-advanced-heading"
    >
      <h3
        id="preferences-advanced-heading"
        class="section-title"
      >
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
  margin-bottom: var(--space-6);
  padding-bottom: var(--space-5);
  border-bottom: var(--border-width) solid var(--border-default);
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0;
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
}

.pref-checkbox {
  width: 1rem;
  height: 1rem;
  align-self: flex-start;
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
  font-size: var(--text-md);
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
