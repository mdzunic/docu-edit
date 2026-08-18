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
import { fromCm, roundForDisplay, toCm } from "../lib/core/units";
import BaseDialog from "./BaseDialog.vue";
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
});

const emit = defineEmits(["apply", "cancel"]);

const unit = ref(String(props.initial.measurementUnit || "cm"));
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
    return {
      widthCm: toCm(Number(widthInput.value), u),
      heightCm: toCm(Number(heightInput.value), u),
    };
  }
  return pageSizeFor(asPaperSizeName(paperSize.value), asOrientation(orientation.value));
});

const margins = computed(() => {
  const u = asMeasurementUnit(unit.value);
  return {
    topCm: toCm(Number(marginTopInput.value), u),
    rightCm: toCm(Number(marginRightInput.value), u),
    bottomCm: toCm(Number(marginBottomInput.value), u),
    leftCm: toCm(Number(marginLeftInput.value), u),
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
  return String(roundForDisplay(fromCm(toCm(Number(value), from), to), to));
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
const marginsError = computed(() =>
  textAreaTooNarrow(pageSize.value, margins.value) ? MARGINS_TOO_LARGE_MESSAGE : "",
);

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
    <section class="section">
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
            type="number"
            step="0.01"
            :aria-label="`Width in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Height ({{ unit }})</span>
          <input
            v-model="heightInput"
            class="field-input"
            data-testid="page-setup-height"
            type="number"
            step="0.01"
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
            type="number"
            step="0.01"
            :aria-label="`Top margin in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Bottom ({{ unit }})</span>
          <input
            v-model="marginBottomInput"
            class="field-input"
            data-testid="page-setup-margin-bottom"
            type="number"
            step="0.01"
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
            type="number"
            step="0.01"
            :aria-label="`Left margin in ${unit}`"
          >
        </label>
        <label class="field">
          <span class="field-label">Right ({{ unit }})</span>
          <input
            v-model="marginRightInput"
            class="field-input"
            data-testid="page-setup-margin-right"
            type="number"
            step="0.01"
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
