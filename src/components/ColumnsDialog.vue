<script setup lang="ts">
// The Format → Columns dialog (SPEC-030 AC-1..AC-5): presets (1/2/3), a
// custom count (1..6) that reverts on an out-of-range entry, per-column
// widths + gutter, "Line between columns", and an "Apply to" scope. OK emits
// the resolved layout and scope for the parent to apply through a command;
// Cancel closes. BaseDialog owns the shell, Escape and initial focus (the
// [autofocus] count input).
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import {
  MAX_COLUMN_COUNT,
  MIN_COLUMN_COUNT,
  MIN_SECTION_COLUMN_WIDTH_CM,
  asSectionColumns,
  resolveColumnWidthsCm,
} from "../lib/documentModel";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** { columns: SectionColumns, contentWidthCm: number, applyTo: string }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const contentWidthCm = Number(props.initial.contentWidthCm) || 16;
const seed = asSectionColumns(props.initial.columns);

const count = ref(Math.max(MIN_COLUMN_COUNT, Math.min(MAX_COLUMN_COUNT, seed.count || 1)));
const countDraft = ref(String(count.value));
const countError = ref("");
const equalWidth = ref(seed.equalWidth !== false);
const separator = ref(seed.separator === true);
const gapCm = ref(round(seed.gapCm ?? 1.27));
const widths = ref(seedWidths());
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
function seedWidths() {
  const resolved = resolveColumnWidthsCm(
    { count: count.value, equalWidth: equalWidth.value, gapCm: gapCm.value, separator: false, widthsCm: seed.widthsCm },
    contentWidthCm,
  );
  return resolved.map(round);
}

// The widths shown in the (read-only, when equal) width fields.
const displayWidths = computed(() => {
  if (equalWidth.value) {
    return resolveColumnWidthsCm(
      { count: count.value, equalWidth: true, gapCm: gapCm.value, separator: false },
      contentWidthCm,
    ).map(round);
  }
  return widths.value;
});

function setCount(next = 1) {
  count.value = next;
  countDraft.value = String(next);
  countError.value = "";
  widths.value = seedWidths();
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
    widths.value = resolveColumnWidthsCm(
      { count: count.value, equalWidth: true, gapCm: gapCm.value, separator: false },
      contentWidthCm,
    ).map(round);
  }
}

const widthError = computed(() => {
  if (equalWidth.value || count.value <= 1) {
    return "";
  }
  for (const width of widths.value) {
    if (!Number.isFinite(width) || width < MIN_SECTION_COLUMN_WIDTH_CM) {
      return `Columns are too narrow (minimum ${MIN_SECTION_COLUMN_WIDTH_CM} cm)`;
    }
  }
  return "";
});

const error = computed(() => widthError.value);

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
      widthsCm: equalWidth.value ? undefined : widths.value.slice(),
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
        <span class="field-label">Spacing (cm)</span>
        <input
          v-model.number="gapCm"
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
        <span class="field-label">Column {{ index + 1 }} width (cm)</span>
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
          v-model.number="widths[index]"
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
      <button
        type="button"
        data-testid="columns-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="columns-confirm"
        :disabled="!!error"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.presets {
  margin-bottom: 0.75rem;
}

.preset-buttons {
  display: flex;
  gap: 0.5rem;
  margin-top: 0.25rem;
}

.preset-button {
  font: inherit;
  min-width: 2.5rem;
  padding: 0.375rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.preset-button.active {
  border-color: #1a73e8;
  background: #e8f0fe;
  color: #1a73e8;
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
  margin-bottom: 0.75rem;
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

.field-input:disabled {
  background: #f1f3f4;
  color: #80868b;
}

.widths {
  display: flex;
  flex-wrap: wrap;
  gap: 0.75rem;
}

.width-field {
  flex: 1 1 6rem;
  margin-bottom: 0.5rem;
}

.checkbox-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.75rem;
  font-size: 0.875rem;
  color: #3c4043;
}

.dialog-error {
  margin: 0.25rem 0 0;
  font-size: 0.8125rem;
  color: #c5221f;
}

.dialog-actions button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #1a73e8;
  border-color: #1a73e8;
  color: #fff;
}

.dialog-actions button.primary:disabled {
  background: #9e9e9e;
  border-color: #9e9e9e;
  cursor: default;
}
</style>
