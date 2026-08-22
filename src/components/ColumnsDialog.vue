<script setup lang="ts">
// The Format → Columns dialog (SPEC-030 AC-1..AC-5): presets (1/2/3), a
// custom count (1..6) that reverts on an out-of-range entry, per-column
// widths + gutter, "Line between columns", and an "Apply to" scope. OK emits
// the resolved layout and scope for the parent to apply through a command;
// Cancel closes. BaseDialog owns the shell, Escape and initial focus (the
// [autofocus] count input).
import { computed, nextTick, onMounted, ref, useTemplateRef, watch } from "vue";
import {
  MAX_COLUMN_COUNT,
  MIN_COLUMN_COUNT,
  MIN_SECTION_COLUMN_WIDTH_CM,
  asSectionColumns,
  resolveColumnWidthsCm,
} from "../lib/documentModel";
import { asMeasurementUnit, formatLength, parseLengthToCm } from "../lib/core/units";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

const props = defineProps({
  /** { columns: SectionColumns, contentWidthCm: number, applyTo: string }. */
  initial: { type: Object, required: true },
  /** The document's measurement-unit preference (SPEC-064) — the length field
   * text displays and parses in this unit; the model stays in centimetres. */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["apply", "cancel"]);

// The unit the width/spacing fields display/parse in; stored lengths stay cm.
const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const contentWidthCm = Number(props.initial.contentWidthCm) || 16;
const seed = asSectionColumns(props.initial.columns);
const seedGapCm = round(seed.gapCm ?? 1.27);

const count = ref(Math.max(MIN_COLUMN_COUNT, Math.min(MAX_COLUMN_COUNT, seed.count || 1)));
const countDraft = ref(String(count.value));
const countError = ref("");
const equalWidth = ref(seed.equalWidth !== false);
const separator = ref(seed.separator === true);
// Spacing/gap held as display text in the current unit; cm stays canonical.
const gapDraft = ref(formatLength(seedGapCm, measurementUnit.value));
const gapCm = computed(() => {
  const parsed = parseLengthToCm(gapDraft.value, measurementUnit.value);
  return parsed !== null && parsed >= 0 ? round(parsed) : seedGapCm;
});
// Per-column widths held as display text; the parsed cm values stay canonical.
const widthDrafts = ref(seedWidthsCm().map((cm) => formatLength(cm, measurementUnit.value)));
const widthsCm = computed(() =>
  widthDrafts.value.map((text) => {
    const parsed = parseLengthToCm(text, measurementUnit.value);
    return parsed === null ? Number.NaN : round(parsed);
  }),
);
const applyTo = ref(props.initial.applyTo === "selection" ? "selection" : "thisSection");

const countInput = useTemplateRef("countInput");

// BaseDialog focuses the [autofocus] count input; selecting the prefilled
// count on top of that (type-to-replace) is this dialog's own one-shot
// behaviour.
onMounted(async () => {
  await nextTick();
  countInput.value?.select();
});

function round(value = 0) {
  return Math.round(value * 100) / 100;
}

// The per-column widths to seed the editable fields with: the section's own
// widths when it has unequal columns, else the evenly computed widths.
function seedWidthsCm() {
  const resolved = resolveColumnWidthsCm(
    { count: count.value, equalWidth: equalWidth.value, gapCm: gapCm.value, separator: false, widthsCm: seed.widthsCm },
    contentWidthCm,
  );
  return resolved.map(round);
}

// The widths shown in the fields, as display text in the current unit: the
// evenly computed split when equal (read-only), else the editable drafts.
const displayWidths = computed(() => {
  if (equalWidth.value) {
    return resolveColumnWidthsCm(
      { count: count.value, equalWidth: true, gapCm: gapCm.value, separator: false },
      contentWidthCm,
    ).map((cm) => formatLength(cm, measurementUnit.value));
  }
  return widthDrafts.value;
});

function setCount(next = 1) {
  count.value = next;
  countDraft.value = String(next);
  countError.value = "";
  widthDrafts.value = seedWidthsCm().map((cm) => formatLength(cm, measurementUnit.value));
}

function choosePreset(next = 1) {
  setCount(next);
}

// AC-2: a custom count outside 1..6 (or non-integer) is rejected with an
// inline message and the field reverts to the previous valid value.
function commitCount() {
  const parsed = Number(countDraft.value);
  if (!Number.isInteger(parsed) || parsed < MIN_COLUMN_COUNT || parsed > MAX_COLUMN_COUNT) {
    countError.value = `Enter a whole number from ${MIN_COLUMN_COUNT} to ${MAX_COLUMN_COUNT}`;
    countDraft.value = String(count.value);
    return;
  }
  setCount(parsed);
}

// v-model has already flipped `equalWidth` by the time this @change fires;
// when it just became false, seed the now-editable fields from the even split.
function onEqualChange() {
  if (!equalWidth.value) {
    widthDrafts.value = resolveColumnWidthsCm(
      { count: count.value, equalWidth: true, gapCm: gapCm.value, separator: false },
      contentWidthCm,
    ).map((cm) => formatLength(cm, measurementUnit.value));
  }
}

const gapError = computed(() => {
  if (count.value <= 1) {
    return "";
  }
  const parsed = parseLengthToCm(gapDraft.value, measurementUnit.value);
  if (gapDraft.value === "" || parsed === null || parsed < 0) {
    return `Spacing must be a non-negative ${unitLabel.value} value`;
  }
  return "";
});

const widthError = computed(() => {
  if (equalWidth.value || count.value <= 1) {
    return "";
  }
  for (const width of widthsCm.value) {
    if (!Number.isFinite(width) || width < MIN_SECTION_COLUMN_WIDTH_CM) {
      return `Columns are too narrow (minimum ${formatLength(MIN_SECTION_COLUMN_WIDTH_CM, measurementUnit.value)} ${unitLabel.value})`;
    }
  }
  return "";
});

const error = computed(() => gapError.value || widthError.value);

// Reformat the visible field text when the unit changes, re-deriving each from
// its underlying cm value so centimetres stay the source of truth (SPEC-064).
watch(
  () => props.unit,
  (next = "cm", previous = "cm") => {
    const previousUnit = asMeasurementUnit(previous);
    const nextUnit = asMeasurementUnit(next);
    const gap = parseLengthToCm(gapDraft.value, previousUnit);
    if (gap !== null) {
      gapDraft.value = formatLength(gap, nextUnit);
    }
    widthDrafts.value = widthDrafts.value.map((text) => {
      const cm = parseLengthToCm(text, previousUnit);
      return cm === null ? text : formatLength(cm, nextUnit);
    });
  },
);

function apply() {
  if (error.value) {
    return;
  }
  emit("apply", {
    columns: {
      count: count.value,
      equalWidth: equalWidth.value,
      gapCm: gapCm.value,
      separator: separator.value,
      widthsCm: equalWidth.value ? undefined : widthsCm.value.slice(),
    },
    applyTo: applyTo.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Columns"
    testid="columns-dialog"
    @dismiss="emit('cancel')"
  >
    <div class="presets">
      <span class="field-label">Presets</span>
      <div class="preset-buttons">
        <button
          v-for="preset in [1, 2, 3]"
          :key="preset"
          type="button"
          class="preset-button"
          :class="{ active: count === preset }"
          :data-testid="`columns-preset-${preset}`"
          :aria-pressed="count === preset"
          @click="choosePreset(preset)"
        >
          {{ preset }}
        </button>
      </div>
    </div>

    <div class="field-row">
      <label class="field">
        <span class="field-label">Number of columns</span>
        <input
          ref="countInput"
          v-model="countDraft"
          class="field-input"
          data-testid="columns-count"
          type="number"
          :min="MIN_COLUMN_COUNT"
          :max="MAX_COLUMN_COUNT"
          step="1"
          autofocus
          @change="commitCount"
          @keydown.enter="commitCount"
        >
      </label>
      <label class="field">
        <span class="field-label">Spacing ({{ unitLabel }})</span>
        <input
          v-model="gapDraft"
          class="field-input"
          data-testid="columns-spacing"
          type="number"
          min="0"
          step="0.1"
          :disabled="count <= 1"
        >
      </label>
    </div>
    <p
      v-if="countError"
      class="dialog-error"
      data-testid="columns-count-error"
    >
      {{ countError }}
    </p>

    <label
      v-if="count > 1"
      class="checkbox-row"
    >
      <input
        v-model="equalWidth"
        type="checkbox"
        data-testid="columns-equal"
        @change="onEqualChange"
      >
      <span>Equal column widths</span>
    </label>

    <div
      v-if="count > 1"
      class="widths"
    >
      <label
        v-for="(width, index) in displayWidths"
        :key="index"
        class="field width-field"
      >
        <span class="field-label">Column {{ index + 1 }} width ({{ unitLabel }})</span>
        <input
          v-if="equalWidth"
          class="field-input"
          :data-testid="`columns-width-${index}`"
          type="number"
          :value="width"
          disabled
        >
        <input
          v-else
          v-model="widthDrafts[index]"
          class="field-input"
          :data-testid="`columns-width-${index}`"
          type="number"
          :min="MIN_SECTION_COLUMN_WIDTH_CM"
          step="0.1"
        >
      </label>
    </div>

    <label
      v-if="count > 1"
      class="checkbox-row"
    >
      <input
        v-model="separator"
        type="checkbox"
        data-testid="columns-separator"
      >
      <span>Line between columns</span>
    </label>

    <label class="field">
      <span class="field-label">Apply to</span>
      <select
        v-model="applyTo"
        class="field-input"
        data-testid="columns-apply-to"
      >
        <option value="thisSection">Current section</option>
        <option value="selection">Selection</option>
      </select>
    </label>

    <p
      v-if="error"
      class="dialog-error"
      data-testid="columns-error"
    >
      {{ error }}
    </p>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="columns-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="columns-confirm"
        :disabled="!!error"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.presets {
  margin-bottom: var(--space-4);
}

.preset-buttons {
  display: flex;
  gap: var(--space-3);
  margin-top: var(--space-2);
}

.preset-button {
  font: inherit;
  min-width: 2.5rem;
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.preset-button.active {
  border-color: var(--accent);
  background: var(--surface-selected);
  color: var(--text-accent);
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
  margin-bottom: var(--space-4);
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

.field-input:disabled {
  background: var(--surface-canvas);
  color: var(--text-disabled);
}

.widths {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-4);
}

.width-field {
  flex: 1 1 6rem;
  margin-bottom: var(--space-3);
}

.checkbox-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.dialog-error {
  margin: var(--space-2) 0 0;
  font-size: var(--text-md);
  color: var(--text-danger);
}
</style>
