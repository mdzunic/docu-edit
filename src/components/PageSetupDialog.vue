<script setup lang="ts">
import { computed, ref } from "vue";
import { asMeasurementUnit } from "../lib/core/units";
import {
  asOrientation,
  asPaperSizeName,
  MARGINS_TOO_LARGE_MESSAGE,
  matchPaperSize,
  orientationOf,
  PAPER_SIZE_NAMES,
  pageSizeFor,
  textAreaTooNarrow,
} from "../lib/core/pageSetup";
import { fromCm, parseLengthToCm, roundForDisplay } from "../lib/core/units";
import {
  asPageBorderEdge,
  emptyPageBorderDraft,
  pageBackgroundFromDraft,
  pageBorderDraftFrom,
  pageBordersFromDraft,
  eventTargetChecked,
  eventTargetValue,
  PAGE_BORDER_EDGES,
} from "../lib/pageBorders";
import { asDocumentTheme, asPageBorders, asPageBorderAppliesTo, asPageBorderMeasureFrom } from "../lib/documentModel";
import { isThemeColor, resolveColor } from "../lib/core/theme";
import { BORDER_LINE_STYLES, asBorderLineStyleValue } from "../lib/core/borders";
import BaseDialog from "./BaseDialog.vue";
import ThemeSwatchRow from "./ThemeSwatchRow.vue";
import UiButton from "./ui/UiButton.vue";

// The "Page setup" dialog (SPEC-027 AC-1..9): paper size (a standard size or
// Custom width/height), orientation and the four margins, all edited in the
// document's measurement-unit preference (AC-6). BaseDialog owns the shell,
// Escape and initial focus (the [autofocus] paper-size select); this file
// keeps only the tall body and its scroll behaviour. Every field commits
// together on OK as one combined patch the parent applies through
// commands.ts's setPageSetup (SPEC-027, mirroring how TablePropertiesDialog
// batches its fields into one undo step); Cancel discards everything.
//
// Fields are plain strings displayed in the current unit; the unit selector
// converts them (through centimetres, the model's own unit) instead of
// reinterpreting the digits, so switching cm -> in -> cm round-trips the
// same physical size.
const props = defineProps({
  /** Prefill: { pageSize: {widthCm,heightCm}, margins: {topCm,rightCm,
   * bottomCm,leftCm}, measurementUnit, applyTo }. */
  initial: { type: Object, required: true },
  /** The DocumentTheme (SPEC-077 AC-22): when present, a theme swatch row
   * renders under the page-background control; picking a swatch makes OK
   * carry the ThemeColorRef itself, not the hex it displays as. Null (the
   * default) renders no row — existing mounts are unchanged. */
  theme: { type: Object, default: null },
});

const emit = defineEmits(["apply", "cancel"]);

const unit = ref(String(props.initial.measurementUnit || "cm"));
// SPEC-074 AC-19: Page setup gains a Borders tab. Two panels, one form —
// OK still commits everything as one edit, exactly as before.
const activeTab = ref("page");
const borderDraft = ref(emptyPageBorderDraft());
borderDraft.value = pageBorderDraftFrom(
  asPageBorders(props.initial.pageBorders),
  // SPEC-077: a theme-referenced background seeds the draft as the reference
  // itself, so an untouched OK re-emits it — reopening must not flatten it.
  typeof props.initial.pageBackground === "string" || isThemeColor(props.initial.pageBackground)
    ? props.initial.pageBackground
    : undefined,
);

// The colour inputs can only display a hex, so a theme reference resolves
// for display; the draft keeps the reference until the input itself is
// edited (which writes a literal through the handlers below) — SPEC-077.
const borderColorHex = computed(() => resolveColor(borderDraft.value.color, asDocumentTheme(props.theme)));
const backgroundHex = computed(() => resolveColor(borderDraft.value.background, asDocumentTheme(props.theme)));

function setBorderColor(value = "#000000") {
  borderDraft.value.color = value;
}

function setBackgroundColor(value = "#fdf6e3") {
  borderDraft.value.background = value;
}

/** A theme swatch for the page background (SPEC-077 AC-22): the draft holds
 * the reference; the colour input shows its resolved hex. */
function onBackgroundSwatch(reference = Object()) {
  borderDraft.value.background = reference;
  borderDraft.value.backgroundOn = true;
}

const BORDER_STYLE_LABELS = {
  none: "None",
  single: "Single",
  thick: "Thick",
  double: "Double",
  dashed: "Dashed",
  dotted: "Dotted",
  dashDot: "Dash-dot",
  dashDotDot: "Dash-dot-dot",
};

const BORDER_EDGE_LABELS = { top: "Top", bottom: "Bottom", left: "Left", right: "Right" };

function toggleBorderEdge(key = "top", on = false) {
  borderDraft.value[asPageBorderEdge(key)] = on;
}

function setBorderStyle(value = "single") {
  borderDraft.value.style = asBorderLineStyleValue(value);
}

function setMeasureFrom(value = "page") {
  borderDraft.value.measureFrom = asPageBorderMeasureFrom(value);
}

function setAppliesTo(value = "allPages") {
  borderDraft.value.appliesTo = asPageBorderAppliesTo(value);
}
// What the edit targets (SPEC-028 AC-20/21): the caret's section or the
// whole document — commands.PageSetupScope's values.
const applyTo = ref(String(props.initial.applyTo || "wholeDocument"));
const paperSize = ref(String(matchPaperSize(props.initial.pageSize)));
const orientation = ref(String(orientationOf(props.initial.pageSize)));

function displayValue(cm = 0) {
  const u = asMeasurementUnit(unit.value);
  return String(roundForDisplay(fromCm(cm, u), u));
}

const widthInput = ref(displayValue(props.initial.pageSize.widthCm));
const heightInput = ref(displayValue(props.initial.pageSize.heightCm));
const marginTopInput = ref(displayValue(props.initial.margins.topCm));
const marginRightInput = ref(displayValue(props.initial.margins.rightCm));
const marginBottomInput = ref(displayValue(props.initial.margins.bottomCm));
const marginLeftInput = ref(displayValue(props.initial.margins.leftCm));

const PAPER_OPTIONS = PAPER_SIZE_NAMES;
const UNIT_OPTIONS = ["cm", "mm", "in", "pt"];

/** The page size the current fields describe (AC-1, AC-2, AC-4): a standard
 * size's catalogue entry in the chosen orientation, or the Custom width/
 * height fields converted to cm. */
const pageSize = computed(() => {
  if (paperSize.value === "Custom") {
    const u = asMeasurementUnit(unit.value);
    // parseLengthToCm accepts a bare number in the current unit or an explicit
    // suffix ("21cm", "8.5in"); null (unparseable) becomes NaN so sizeError
    // catches it and OK stays disabled (SPEC-064 AC-11/AC-11b).
    return {
      widthCm: parseLengthToCm(widthInput.value, u) ?? NaN,
      heightCm: parseLengthToCm(heightInput.value, u) ?? NaN,
    };
  }
  return pageSizeFor(asPaperSizeName(paperSize.value), asOrientation(orientation.value));
});

const margins = computed(() => {
  const u = asMeasurementUnit(unit.value);
  return {
    topCm: parseLengthToCm(marginTopInput.value, u) ?? NaN,
    rightCm: parseLengthToCm(marginRightInput.value, u) ?? NaN,
    bottomCm: parseLengthToCm(marginBottomInput.value, u) ?? NaN,
    leftCm: parseLengthToCm(marginLeftInput.value, u) ?? NaN,
  };
});

/** Orientation flip (AC-2): a standard size swaps through its own
 * catalogue lookup automatically (pageSize above); Custom has no catalogue
 * entry to re-look-up, so its own width/height fields are swapped here —
 * the only path that actually rotates a custom page. */
function handleOrientationChange() {
  if (paperSize.value === "Custom") {
    const w = widthInput.value;
    widthInput.value = heightInput.value;
    heightInput.value = w;
  }
}

/** Tracks the paper size just before the current change, so switching to
 * Custom can seed its width/height from whatever standard size was active
 * — otherwise the fields would show stale leftovers from whatever the
 * dialog originally opened with. Switching between two standard sizes, or
 * staying on Custom, leaves the Custom fields alone. */
let previousPaperSize = paperSize.value;
function handlePaperSizeChange() {
  if (paperSize.value === "Custom" && previousPaperSize !== "Custom") {
    const seed = pageSizeFor(asPaperSizeName(previousPaperSize), asOrientation(orientation.value));
    widthInput.value = displayValue(seed.widthCm);
    heightInput.value = displayValue(seed.heightCm);
  }
  previousPaperSize = paperSize.value;
}

/** Re-displays every field in the newly chosen unit without changing the
 * physical value it represents (AC-6, AC-7) — converts through centimetres
 * rather than reinterpreting the typed digits. */
let previousUnit = unit.value;
function convert(value = "0") {
  const from = asMeasurementUnit(previousUnit);
  const to = asMeasurementUnit(unit.value);
  // Reinterpret whatever the field holds (a bare number or a suffixed value)
  // through centimetres into the newly chosen unit; leave an unparseable draft
  // untouched so the user can fix it.
  const cm = parseLengthToCm(value, from);
  return cm === null ? value : String(roundForDisplay(fromCm(cm, to), to));
}
function handleUnitChange() {
  widthInput.value = convert(widthInput.value);
  heightInput.value = convert(heightInput.value);
  marginTopInput.value = convert(marginTopInput.value);
  marginRightInput.value = convert(marginRightInput.value);
  marginBottomInput.value = convert(marginBottomInput.value);
  marginLeftInput.value = convert(marginLeftInput.value);
  previousUnit = unit.value;
}

const sizeError = computed(() => {
  if (paperSize.value !== "Custom") {
    return "";
  }
  const { widthCm, heightCm } = pageSize.value;
  return Number.isFinite(widthCm) && widthCm > 0 && Number.isFinite(heightCm) && heightCm > 0
    ? ""
    : "Enter a width and height greater than 0";
});

/** AC-9: rejects margins that would leave a text area narrower than 1cm in
 * either dimension, with the exact required message. */
const marginsError = computed(() => {
  const m = margins.value;
  if (![m.topCm, m.rightCm, m.bottomCm, m.leftCm].every(Number.isFinite)) {
    return "Enter a valid margin";
  }
  return textAreaTooNarrow(pageSize.value, m) ? MARGINS_TOO_LARGE_MESSAGE : "";
});

const error = computed(() => sizeError.value || marginsError.value);

function apply() {
  if (error.value) {
    return;
  }
  emit("apply", {
    pageSize: pageSize.value,
    margins: margins.value,
    measurementUnit: unit.value,
    applyTo: applyTo.value,
    // SPEC-074 AC-19..AC-23 — committed with the rest, as one undo step.
    pageBorders: pageBordersFromDraft(borderDraft.value),
    pageBackground: pageBackgroundFromDraft(borderDraft.value),
  });
}
</script>

<template>
  <BaseDialog
    title="Page setup"
    testid="page-setup-dialog"
    width="28rem"
    @dismiss="emit('cancel')"
  >
    <!-- SPEC-074 AC-19: Page setup gains a Borders tab. -->
    <div
      class="tabs"
      role="tablist"
      aria-label="Page setup sections"
    >
      <button
        type="button"
        class="tab"
        role="tab"
        :class="{ active: activeTab === 'page' }"
        :aria-selected="activeTab === 'page'"
        data-testid="page-setup-tab-page"
        @click="activeTab = 'page'"
      >
        Page
      </button>
      <button
        type="button"
        class="tab"
        role="tab"
        :class="{ active: activeTab === 'borders' }"
        :aria-selected="activeTab === 'borders'"
        data-testid="page-setup-tab-borders"
        @click="activeTab = 'borders'"
      >
        Borders
      </button>
    </div>

    <section
      v-show="activeTab === 'page'"
      class="section"
    >
      <div class="field-row">
        <label class="field">
          <span class="field-label">Paper size</span>
          <select
            v-model="paperSize"
            class="field-input"
            data-testid="page-setup-paper-size"
            autofocus
            @change="handlePaperSizeChange"
          >
            <option
              v-for="name in PAPER_OPTIONS"
              :key="name"
              :value="name"
            >
              {{ name }}
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Orientation</span>
          <select
            v-model="orientation"
            class="field-input"
            data-testid="page-setup-orientation"
            @change="handleOrientationChange"
          >
            <option value="portrait">
              Portrait
            </option>
            <option value="landscape">
              Landscape
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Unit</span>
          <select
            v-model="unit"
            class="field-input"
            data-testid="page-setup-unit"
            @change="handleUnitChange"
          >
            <option
              v-for="u in UNIT_OPTIONS"
              :key="u"
              :value="u"
            >
              {{ u }}
            </option>
          </select>
        </label>
      </div>
      <div
        v-if="paperSize === 'Custom'"
        class="field-row"
      >
        <label class="field">
          <span class="field-label">Width ({{ unit }})</span>
          <input
            v-model="widthInput"
            class="field-input"
            data-testid="page-setup-width"
            type="text"
            inputmode="decimal"
            :aria-label="`Width in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Height ({{ unit }})</span>
          <input
            v-model="heightInput"
            class="field-input"
            data-testid="page-setup-height"
            type="text"
            inputmode="decimal"
            :aria-label="`Height in ${unit}`"
          >
        </label>
      </div>
      <p
        v-if="sizeError"
        class="dialog-error"
        data-testid="page-setup-size-error"
      >
        {{ sizeError }}
      </p>
    </section>

    <section class="section">
      <h3 class="section-title">
        Margins
      </h3>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Top ({{ unit }})</span>
          <input
            v-model="marginTopInput"
            class="field-input"
            data-testid="page-setup-margin-top"
            type="text"
            inputmode="decimal"
            :aria-label="`Top margin in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Bottom ({{ unit }})</span>
          <input
            v-model="marginBottomInput"
            class="field-input"
            data-testid="page-setup-margin-bottom"
            type="text"
            inputmode="decimal"
            :aria-label="`Bottom margin in ${unit}`"
          >
        </label>
      </div>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Left ({{ unit }})</span>
          <input
            v-model="marginLeftInput"
            class="field-input"
            data-testid="page-setup-margin-left"
            type="text"
            inputmode="decimal"
            :aria-label="`Left margin in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Right ({{ unit }})</span>
          <input
            v-model="marginRightInput"
            class="field-input"
            data-testid="page-setup-margin-right"
            type="text"
            inputmode="decimal"
            :aria-label="`Right margin in ${unit}`"
          >
        </label>
      </div>
      <p
        v-if="marginsError"
        class="dialog-error"
        data-testid="page-setup-margins-error"
      >
        {{ marginsError }}
      </p>
    </section>

    <!-- The Borders tab (SPEC-074 AC-19..AC-23): one line, which edges take
         it, what its offset measures from, and the page background. -->
    <section
      v-show="activeTab === 'borders'"
      class="section"
      data-testid="page-setup-borders"
    >
      <div class="field-row">
        <span class="field-label">Edges</span>
        <label
          v-for="edge in PAGE_BORDER_EDGES"
          :key="edge"
          class="edge-toggle"
        >
          <input
            type="checkbox"
            :checked="borderDraft[edge]"
            :data-testid="`page-border-edge-${edge}`"
            @change="toggleBorderEdge(edge, eventTargetChecked($event.target))"
          >
          {{ BORDER_EDGE_LABELS[edge] }}
        </label>
      </div>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Style</span>
          <select
            class="field-input"
            :value="borderDraft.style"
            data-testid="page-border-style"
            @change="setBorderStyle(eventTargetValue($event.target))"
          >
            <option
              v-for="style in BORDER_LINE_STYLES"
              :key="style"
              :value="style"
            >
              {{ BORDER_STYLE_LABELS[style] }}
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Width (pt)</span>
          <input
            v-model.number="borderDraft.widthPt"
            class="field-input"
            type="number"
            min="0.25"
            max="12"
            step="0.25"
            data-testid="page-border-width"
          >
        </label>
        <label class="field">
          <span class="field-label">Colour</span>
          <input
            class="field-input colour-input"
            type="color"
            :value="borderColorHex"
            data-testid="page-border-color"
            @input="setBorderColor(eventTargetValue($event.target))"
          >
        </label>
      </div>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Offset (pt)</span>
          <input
            v-model.number="borderDraft.offsetPt"
            class="field-input"
            type="number"
            min="0"
            max="31"
            step="1"
            data-testid="page-border-offset"
          >
        </label>
        <label class="field">
          <span class="field-label">Measure from</span>
          <select
            class="field-input"
            :value="borderDraft.measureFrom"
            data-testid="page-border-measure-from"
            @change="setMeasureFrom(eventTargetValue($event.target))"
          >
            <option value="page">
              Page edge
            </option>
            <option value="text">
              Text
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Apply border to</span>
          <select
            class="field-input"
            :value="borderDraft.appliesTo"
            data-testid="page-border-applies-to"
            @change="setAppliesTo(eventTargetValue($event.target))"
          >
            <option value="allPages">
              All pages
            </option>
            <option value="firstPage">
              First page only
            </option>
            <option value="exceptFirstPage">
              All but the first
            </option>
          </select>
        </label>
      </div>
      <div class="field-row">
        <label class="edge-toggle">
          <input
            v-model="borderDraft.backgroundOn"
            type="checkbox"
            data-testid="page-background-on"
          >
          Page background
        </label>
        <input
          class="field-input colour-input"
          type="color"
          :value="backgroundHex"
          :disabled="!borderDraft.backgroundOn"
          data-testid="page-background-color"
          aria-label="Page background colour"
          @input="setBackgroundColor(eventTargetValue($event.target))"
        >
      </div>
      <ThemeSwatchRow
        :theme="theme"
        testid="page-background-theme"
        @select="onBackgroundSwatch"
      />
      <p class="hint">
        Page backgrounds print only when "Print background colours" is chosen
        in the print preview; page borders always print.
      </p>
    </section>

    <section class="section">
      <div class="field-row">
        <label class="field">
          <span class="field-label">Apply to</span>
          <select
            v-model="applyTo"
            class="field-input"
            data-testid="page-setup-apply-to"
          >
            <option value="thisSection">
              this section
            </option>
            <option value="wholeDocument">
              whole document
            </option>
          </select>
        </label>
      </div>
    </section>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="page-setup-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="page-setup-apply"
        :disabled="!!error"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.tabs {
  display: flex;
  gap: var(--space-2);
  margin-bottom: var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.tab {
  padding: var(--space-3) var(--space-4);
  border: none;
  border-bottom: var(--space-1) solid transparent;
  background: none;
  color: inherit;
  cursor: pointer;
  font: inherit;
}

.tab.active {
  border-bottom-color: var(--accent);
  color: var(--accent);
}

.edge-toggle {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  white-space: nowrap;
}

.colour-input {
  width: 46px;
  height: 28px;
  padding: 0;
}

.hint {
  margin: var(--space-3) 0 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

/* This dialog's body is tall, so the panel keeps its own scroll cap (the
   base's .dialog is the short-dialog variant without one). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

.section {
  margin-bottom: var(--space-6);
  padding-bottom: var(--space-5);
  border-bottom: var(--border-width) solid var(--border-default);
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
}

.field-row {
  display: flex;
  gap: var(--space-4);
  margin-bottom: var(--space-4);
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.field-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
}
</style>
