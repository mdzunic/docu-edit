<script setup lang="ts">
// The conditional-format rule editor (SPEC-129 "The rule editor"): `cf-editor`
// inside the rules manager, bound to useConditionalFormatEditor. Excel's six
// rule types (`cf-rule-type`) and each type's fields — colour-scale stops
// (`cf-style`, `cf-stop-type-<i>`, `cf-stop-value-<i>`, `cf-stop-color-<i>`),
// data-bar options (`cf-bar-color`, `cf-bar-gradient`,
// `cf-bar-negative-color`, `cf-bar-axis`), icons (`cf-icon-set`,
// `cf-icon-threshold-<i>`, `cf-icon-reverse` — absent for an `.ods`
// workbook, which has no wire form for it — and `cf-show-value`),
// `cf-contain-kind`, `cf-operator`, `cf-value-1`/`cf-value-2`; the format
// (`cf-format`: Excel's six presets and "Custom format…", which opens
// FormatCellsDialog in its differential mode), the live `cf-preview`, and
// `cf-save`/`cf-cancel` (`cf-editor-back` cancels too). An invalid value is
// shown inline (`cf-editor-error`) and nothing is saved. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { ICON_SET_LABELS } from "../../lib/spreadsheet/cfIcons";
import {
  CF_CONTAIN_KINDS,
  CF_OPERATORS,
  CF_RULE_TYPES,
  CF_TEXT_OPERATORS,
  CF_VALUE_TYPES,
  CF_VALUES_STYLES,
} from "../../views/spreadsheet/cfEditorFields";
import { TIME_PERIOD_LABELS } from "../../views/spreadsheet/cfDescriptions";
import { CF_FORMAT_PRESETS } from "../../views/spreadsheet/cfPresets";
import { asConditionalFormatEditor, CUSTOM_FORMAT } from "../../views/spreadsheet/useConditionalFormatEditor";
import CfRulePreview from "./CfRulePreview.vue";
import FormatCellsDialog from "./FormatCellsDialog.vue";

const props = defineProps({
  /** useConditionalFormatEditor's return. */
  editor: { type: Object, required: true },
});

const e = computed(() => asConditionalFormatEditor(props.editor));
const f = computed(() => e.value.fields.value);
const err = computed(() => e.value.error.value);
const twoValues = computed(() => f.value.operator === "between" || f.value.operator === "notBetween");
/** The stop rows a colour scale (or a data bar's min and max) shows: `i` is
 * the stop's index in the rule (its test ids), `at` the field it edits (the
 * fields keep lowest, midpoint, highest). */
const stopRows = computed(() =>
  f.value.style === "3-color"
    ? [
        { i: 0, at: 0, label: "Minimum" },
        { i: 1, at: 1, label: "Midpoint" },
        { i: 2, at: 2, label: "Maximum" },
      ]
    : [
        { i: 0, at: 0, label: "Minimum" },
        { i: 1, at: 2, label: "Maximum" },
      ],
);
const scaleTypes = CF_VALUE_TYPES.filter((type) => type.id !== "autoMin" && type.id !== "autoMax");
const PERIODS = Object.entries(TIME_PERIOD_LABELS).map(([id, label]) => ({ id, label }));
const AXES = [
  { id: "automatic", label: "Automatic" },
  { id: "middle", label: "Cell midpoint" },
  { id: "none", label: "None" },
];

/** An input or select's value, as text. */
function valueOf(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement ? el.value : "";
}

function checked(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement && el.checked;
}

function invalid(field = "") {
  return err.value !== null && err.value.field === field;
}

/** Enter in a value field saves, as a dialog's default button would. */
function onEnter() {
  e.value.save();
}
</script>

<template>
  <section
    class="cf-editor"
    data-testid="cf-editor"
    aria-label="Conditional formatting rule"
  >
    <header class="cf-editor-header">
      <button
        type="button"
        class="cf-editor-back"
        data-testid="cf-editor-back"
        aria-label="Back to the rules"
        @click="e.cancel()"
      >
        ←
      </button>
      <h3 class="cf-editor-title">
        {{ e.ruleId.value ? "Edit formatting rule" : "New formatting rule" }}
      </h3>
    </header>

    <label class="cf-field">Rule type
      <select
        data-testid="cf-rule-type"
        :value="f.type"
        @change="e.set('type', valueOf($event))"
      >
        <option
          v-for="type in CF_RULE_TYPES"
          :key="type.id"
          :value="type.id"
        >{{ type.label }}</option>
      </select>
    </label>

    <template v-if="f.type === 'values'">
      <label class="cf-field">Format style
        <select
          data-testid="cf-style"
          :value="f.style"
          @change="e.set('style', valueOf($event))"
        >
          <option
            v-for="style in CF_VALUES_STYLES"
            :key="style.id"
            :value="style.id"
          >{{ style.label }}</option>
        </select>
      </label>
      <template v-if="f.style !== 'iconSet'">
        <div
          v-for="row in stopRows"
          :key="row.i"
          class="cf-row"
        >
          <span class="cf-row-label">{{ row.label }}</span>
          <select
            :data-testid="`cf-stop-type-${row.i}`"
            :aria-label="`${row.label} type`"
            :value="f.stops[row.at].type"
            @change="e.setStop(row.at, 'type', valueOf($event))"
          >
            <option
              v-for="type in f.style === 'dataBar' ? CF_VALUE_TYPES : scaleTypes"
              :key="type.id"
              :value="type.id"
            >
              {{ type.label }}
            </option>
          </select>
          <input
            type="text"
            :data-testid="`cf-stop-value-${row.i}`"
            :aria-label="`${row.label} value`"
            :aria-invalid="invalid(`stop-${row.at}`)"
            :disabled="['min', 'max', 'autoMin', 'autoMax'].includes(f.stops[row.at].type)"
            :value="f.stops[row.at].value"
            @input="e.setStop(row.at, 'value', valueOf($event))"
          >
          <input
            v-if="f.style !== 'dataBar'"
            type="color"
            :data-testid="`cf-stop-color-${row.i}`"
            :aria-label="`${row.label} colour`"
            :value="e.colorHex(f.stops[row.at].color)"
            @input="e.setStop(row.at, 'color', valueOf($event))"
          >
        </div>
      </template>
      <template v-if="f.style === 'dataBar'">
        <label class="cf-field">Bar colour
          <input
            type="color"
            data-testid="cf-bar-color"
            :value="e.colorHex(f.bar.color)"
            @input="e.setBar('color', valueOf($event))"
          >
        </label>
        <label class="cf-check"><input
          type="checkbox"
          data-testid="cf-bar-gradient"
          :checked="f.bar.gradient"
          @change="e.setBar('gradient', checked($event))"
        > Gradient fill</label>
        <label class="cf-field">Negative bar colour
          <input
            type="color"
            data-testid="cf-bar-negative-color"
            :value="e.colorHex(f.bar.negativeColor ?? '#FF0000')"
            @input="e.setBar('negativeColor', valueOf($event))"
          >
        </label>
        <label class="cf-field">Axis
          <select
            data-testid="cf-bar-axis"
            :value="f.bar.axis"
            @change="e.setBar('axis', valueOf($event))"
          >
            <option
              v-for="axis in AXES"
              :key="axis.id"
              :value="axis.id"
            >{{ axis.label }}</option>
          </select>
        </label>
      </template>
      <template v-if="f.style === 'iconSet'">
        <label class="cf-field">Icon style
          <select
            data-testid="cf-icon-set"
            :value="f.iconSet"
            @change="e.setIconSet(valueOf($event))"
          >
            <option
              v-for="name in e.iconSets.value"
              :key="name"
              :value="name"
            >{{ ICON_SET_LABELS[name] }}</option>
          </select>
        </label>
        <div
          v-for="(threshold, i) in f.iconThresholds"
          :key="i"
          class="cf-row"
        >
          <span class="cf-row-label">Icon {{ i + 1 }}</span>
          <select
            :data-testid="`cf-icon-threshold-op-${i}`"
            :aria-label="`Icon ${i + 1} comparison`"
            :value="String(threshold.gte)"
            @change="e.setIconThreshold(i, 'gte', valueOf($event))"
          >
            <option value="true">
              ≥
            </option>
            <option value="false">
              &gt;
            </option>
          </select>
          <input
            type="text"
            :data-testid="`cf-icon-threshold-${i}`"
            :aria-label="`Icon ${i + 1} threshold`"
            :aria-invalid="invalid(`icon-${i}`)"
            :disabled="threshold.type === 'min' || threshold.type === 'max'"
            :value="threshold.value"
            @input="e.setIconThreshold(i, 'value', valueOf($event))"
          >
          <select
            :data-testid="`cf-icon-threshold-type-${i}`"
            :aria-label="`Icon ${i + 1} type`"
            :value="threshold.type"
            @change="e.setIconThreshold(i, 'type', valueOf($event))"
          >
            <option
              v-for="type in scaleTypes"
              :key="type.id"
              :value="type.id"
            >
              {{ type.label }}
            </option>
          </select>
        </div>
        <label
          v-if="!e.isOds.value"
          class="cf-check"
        ><input
          type="checkbox"
          data-testid="cf-icon-reverse"
          :checked="f.iconReverse"
          @change="e.set('iconReverse', checked($event))"
        > Reverse icon order</label>
      </template>
      <label
        v-if="f.style === 'dataBar' || f.style === 'iconSet'"
        class="cf-check"
      ><input
        type="checkbox"
        data-testid="cf-show-value"
        :checked="f.showValue"
        @change="e.set('showValue', checked($event))"
      > Show value</label>
    </template>

    <template v-else-if="f.type === 'contain'">
      <label class="cf-field">Format only cells with
        <select
          data-testid="cf-contain-kind"
          :value="f.containKind"
          @change="e.set('containKind', valueOf($event))"
        >
          <option
            v-for="kind in CF_CONTAIN_KINDS"
            :key="kind.id"
            :value="kind.id"
          >{{ kind.label }}</option>
        </select>
      </label>
      <template v-if="f.containKind === 'cellValue'">
        <select
          data-testid="cf-operator"
          aria-label="Operator"
          :value="f.operator"
          @change="e.set('operator', valueOf($event))"
        >
          <option
            v-for="op in CF_OPERATORS"
            :key="op.id"
            :value="op.id"
          >
            {{ op.label }}
          </option>
        </select>
        <input
          type="text"
          data-testid="cf-value-1"
          aria-label="Value"
          :aria-invalid="invalid('value1')"
          :value="f.value1"
          @input="e.set('value1', valueOf($event))"
          @keydown.enter.prevent="onEnter"
        >
        <input
          v-if="twoValues"
          type="text"
          data-testid="cf-value-2"
          aria-label="Second value"
          :aria-invalid="invalid('value2')"
          :value="f.value2"
          @input="e.set('value2', valueOf($event))"
          @keydown.enter.prevent="onEnter"
        >
      </template>
      <template v-else-if="f.containKind === 'text'">
        <select
          data-testid="cf-operator"
          aria-label="Operator"
          :value="f.textOperator"
          @change="e.set('textOperator', valueOf($event))"
        >
          <option
            v-for="op in CF_TEXT_OPERATORS"
            :key="op.id"
            :value="op.id"
          >
            {{ op.label }}
          </option>
        </select>
        <input
          type="text"
          data-testid="cf-value-1"
          aria-label="Text"
          :aria-invalid="invalid('value1')"
          :value="f.value1"
          @input="e.set('value1', valueOf($event))"
          @keydown.enter.prevent="onEnter"
        >
      </template>
      <select
        v-else-if="f.containKind === 'dates'"
        data-testid="cf-operator"
        aria-label="Dates occurring"
        :value="f.period"
        @change="e.set('period', valueOf($event))"
      >
        <option
          v-for="period in PERIODS"
          :key="period.id"
          :value="period.id"
        >
          {{ period.label }}
        </option>
      </select>
    </template>

    <div
      v-else-if="f.type === 'rank'"
      class="cf-row"
    >
      <select
        data-testid="cf-rank-direction"
        aria-label="Top or bottom"
        :value="f.bottom ? 'bottom' : 'top'"
        @change="e.set('bottom', valueOf($event) === 'bottom')"
      >
        <option value="top">
          Top
        </option>
        <option value="bottom">
          Bottom
        </option>
      </select>
      <input
        type="text"
        data-testid="cf-rank"
        aria-label="Rank"
        :aria-invalid="invalid('rank')"
        :value="f.rank"
        @input="e.set('rank', valueOf($event))"
        @keydown.enter.prevent="onEnter"
      >
      <label class="cf-check"><input
        type="checkbox"
        data-testid="cf-rank-percent"
        :checked="f.percent"
        @change="e.set('percent', checked($event))"
      > % of the selected range</label>
    </div>

    <div
      v-else-if="f.type === 'average'"
      class="cf-row"
    >
      <select
        data-testid="cf-average-direction"
        aria-label="Above or below"
        :value="f.below ? 'below' : 'above'"
        @change="e.set('below', valueOf($event) === 'below')"
      >
        <option value="above">
          above
        </option>
        <option value="below">
          below
        </option>
      </select>
      <label class="cf-check"><input
        type="checkbox"
        data-testid="cf-average-equal"
        :checked="f.equalAverage"
        @change="e.set('equalAverage', checked($event))"
      > or equal to</label>
      <select
        data-testid="cf-average-std-dev"
        aria-label="Standard deviations"
        :value="String(f.stdDev)"
        @change="e.set('stdDev', valueOf($event))"
      >
        <option value="0">
          the average
        </option>
        <option value="1">
          1 std dev
        </option>
        <option value="2">
          2 std dev
        </option>
        <option value="3">
          3 std dev
        </option>
      </select>
    </div>

    <label
      v-else-if="f.type === 'unique'"
      class="cf-field"
    >Format all
      <select
        data-testid="cf-unique"
        :value="f.unique ? 'unique' : 'duplicate'"
        @change="e.set('unique', valueOf($event) === 'unique')"
      >
        <option value="duplicate">duplicate</option>
        <option value="unique">unique</option>
      </select>
    </label>

    <label
      v-else-if="f.type === 'formula'"
      class="cf-field"
    >Format values where this formula is true
      <input
        type="text"
        data-testid="cf-value-1"
        :aria-invalid="invalid('value1')"
        :value="f.value1"
        placeholder="="
        @input="e.set('value1', valueOf($event))"
        @keydown.enter.prevent="onEnter"
      >
    </label>

    <p
      v-if="err"
      class="cf-editor-error"
      role="alert"
      data-testid="cf-editor-error"
    >
      {{ err.message }}
    </p>

    <label
      v-if="e.formatting.value"
      class="cf-field"
    >Format with
      <select
        data-testid="cf-format"
        :value="e.formatChoice.value"
        @change="e.chooseFormat(valueOf($event))"
      >
        <option
          v-for="(preset, index) in CF_FORMAT_PRESETS"
          :key="preset.testid"
          :value="String(index)"
        >{{ preset.label }}</option>
        <option :value="CUSTOM_FORMAT">Custom format…</option>
      </select>
    </label>

    <div class="cf-field">
      Preview
      <CfRulePreview
        :preview="e.preview.value"
        testid="cf-preview"
      />
    </div>

    <div class="cf-editor-actions">
      <button
        type="button"
        class="cf-button cf-button-primary"
        data-testid="cf-save"
        @click="e.save()"
      >
        Save
      </button>
      <button
        type="button"
        class="cf-button"
        data-testid="cf-cancel"
        @click="e.cancel()"
      >
        Cancel
      </button>
    </div>

    <FormatCellsDialog
      v-if="e.formatDialog.open.value"
      :dialog="e.formatDialog"
    />
  </section>
</template>

<style scoped>
.cf-editor {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-3);
  font-size: var(--text-md);
}
.cf-editor-header {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}
.cf-editor-title {
  margin: 0;
  font-size: var(--text-md);
}
.cf-editor-back {
  border: none;
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-lg);
}
.cf-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}
.cf-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}
.cf-row-label {
  min-width: 4.5rem;
}
.cf-row input[type="text"] {
  width: 5rem;
}
.cf-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}
.cf-editor-error {
  margin: 0;
  color: var(--text-danger);
}
[aria-invalid="true"] {
  border-color: var(--danger);
}
.cf-editor-actions {
  display: flex;
  gap: var(--space-3);
  justify-content: flex-end;
}
.cf-button {
  padding: var(--space-2) var(--space-4);
  cursor: pointer;
}
</style>
