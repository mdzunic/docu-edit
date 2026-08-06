<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { asMeasurementUnit } from "../lib/documentModel";
import {
  asOrientation,
  asPaperSizeName,
  fromCm,
  MARGINS_TOO_LARGE_MESSAGE,
  matchPaperSize,
  orientationOf,
  PAPER_SIZE_NAMES,
  pageSizeFor,
  roundForDisplay,
  textAreaTooNarrow,
  toCm,
} from "../lib/pageSetup";

// The "Page setup" dialog (SPEC-027 AC-1..9): paper size (a standard size or
// Custom width/height), orientation and the four margins, all edited in the
// document's measurement-unit preference (AC-6). Every field commits
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
   * bottomCm,leftCm}, measurementUnit }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const unit = ref(String(props.initial.measurementUnit || "cm"));
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

const paperSizeSelect = useTemplateRef("paperSizeSelect");

onMounted(async () => {
  await nextTick();
  paperSizeSelect.value?.focus();
});

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
  });
}
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="page-setup-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Page setup"
    >
      <h2 class="dialog-title">
        Page setup
      </h2>

      <section class="section">
        <div class="field-row">
          <label class="field">
            <span class="field-label">Paper size</span>
            <select
              ref="paperSizeSelect"
              v-model="paperSize"
              class="field-input"
              data-testid="page-setup-paper-size"
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

      <div class="dialog-actions">
        <button
          type="button"
          data-testid="page-setup-cancel"
          @click="emit('cancel')"
        >
          Cancel
        </button>
        <button
          type="button"
          class="primary"
          data-testid="page-setup-apply"
          :disabled="!!error"
          @click="apply"
        >
          OK
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  min-width: 28rem;
  max-width: 90vw;
  max-height: 85vh;
  overflow-y: auto;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.section {
  margin-bottom: 1.25rem;
  padding-bottom: 1rem;
  border-bottom: 1px solid #e0e0e0;
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0 0 0.5rem;
  font-size: 0.8125rem;
  font-weight: 600;
  color: #5f6368;
  text-transform: uppercase;
}

.field-row {
  display: flex;
  gap: 0.75rem;
  margin-bottom: 0.75rem;
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
}

.field-label {
  font-size: 0.8125rem;
  color: #3c4043;
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
