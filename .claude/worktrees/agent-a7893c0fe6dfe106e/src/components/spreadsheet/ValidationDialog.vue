<script setup lang="ts">
// Data ▸ "Data validation…" (SPEC-130 "The dialog"): `validation-dialog` on
// BaseDialog, bound to useValidationDialog. Three tabs (the ARIA tabs
// pattern: `validation-tab-criteria`, `validation-tab-input`,
// `validation-tab-error`; Left/Right, Home and End move between them):
// Criteria — `validation-allow`, `validation-operator` (the types that take
// one), `validation-value-1`/`validation-value-2` labelled per type and
// operator (validationFields.ts's valueLabels), `validation-ignore-blank` and
// `validation-in-cell-dropdown` (a list only); Input message —
// `validation-show-input`, `validation-input-title` (32 characters) and
// `validation-input-text` (255); Error alert — `validation-show-error`,
// `validation-error-style`, `validation-error-title` (32) and
// `validation-error-text` (225). `validation-mixed-notice` says when the
// selection's cells disagree, `validation-dialog-error` shows the dialog's
// own errors (nothing is saved), and `validation-clear-all`,
// `validation-cancel` (Escape too) and `validation-ok` (Enter in a one-line
// field too) close the row. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, nextTick, useTemplateRef } from "vue";
import { asValidationDialog, VALIDATION_DIALOG_TABS } from "../../views/spreadsheet/useValidationDialog";
import {
  ERROR_TEXT_MAX,
  ERROR_TITLE_MAX,
  INPUT_TEXT_MAX,
  INPUT_TITLE_MAX,
  MIXED_SELECTION_NOTICE,
  VALIDATION_ALLOW_OPTIONS,
  VALIDATION_ERROR_STYLE_OPTIONS,
  VALIDATION_OPERATOR_OPTIONS,
} from "../../views/spreadsheet/validationFields";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useValidationDialog's return. */
  dialog: { type: Object, required: true },
});

const d = computed(() => asValidationDialog(props.dialog));
const f = computed(() => d.value.fields.value);
const labels = computed(() => d.value.labels.value);
const err = computed(() => d.value.error.value);
const tablist = useTemplateRef("tablist");

/** An input, textarea or select's value, as text. */
function valueOf(event = new Event("input")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement || el instanceof HTMLTextAreaElement ? el.value : "";
}

function checked(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement && el.checked;
}

function invalid(field = "") {
  return err.value !== null && err.value.field === field ? "true" : undefined;
}

/** The tabs' arrow keys: the neighbouring tab (wrapping), Home / End the
 * first / last — selected and focused. */
async function onTabKey(event = new KeyboardEvent("keydown")) {
  const ids = VALIDATION_DIALOG_TABS.map((tab) => tab.id);
  const at = ids.indexOf(d.value.tab.value);
  const target =
    event.key === "ArrowRight" ? at + 1 : event.key === "ArrowLeft" ? at - 1 + ids.length : event.key === "Home" ? 0 : event.key === "End" ? ids.length - 1 : -1;
  if (target < 0) {
    return;
  }
  event.preventDefault();
  const next = ids[target % ids.length] ?? "criteria";
  d.value.showTab(next);
  await nextTick();
  const button = tablist.value?.querySelector(`[data-testid="validation-tab-${next}"]`);
  if (button instanceof HTMLElement) {
    button.focus();
  }
}
</script>

<template>
  <BaseDialog
    title="Data validation"
    testid="validation-dialog"
    width="32rem"
    @dismiss="d.cancel()"
  >
    <p
      v-if="d.mixed.value"
      class="validation-notice"
      data-testid="validation-mixed-notice"
    >
      {{ MIXED_SELECTION_NOTICE }}
    </p>
    <div
      ref="tablist"
      class="validation-tabs"
      role="tablist"
      aria-label="Data validation settings"
    >
      <button
        v-for="tab in VALIDATION_DIALOG_TABS"
        :id="`validation-tab-${tab.id}`"
        :key="tab.id"
        type="button"
        role="tab"
        class="validation-tab"
        :aria-selected="d.tab.value === tab.id ? 'true' : 'false'"
        aria-controls="validation-tabpanel"
        :tabindex="d.tab.value === tab.id ? 0 : -1"
        :data-testid="`validation-tab-${tab.id}`"
        @click="d.showTab(tab.id)"
        @keydown="onTabKey"
      >
        {{ tab.label }}
      </button>
    </div>
    <div
      id="validation-tabpanel"
      class="validation-panel"
      role="tabpanel"
      :aria-labelledby="`validation-tab-${d.tab.value}`"
    >
      <template v-if="d.tab.value === 'criteria'">
        <div class="validation-field">
          <label for="validation-allow-control">Allow</label>
          <select
            id="validation-allow-control"
            class="validation-control"
            data-testid="validation-allow"
            autofocus
            :value="f.type"
            @change="d.set('type', valueOf($event))"
          >
            <option
              v-for="option in VALIDATION_ALLOW_OPTIONS"
              :key="option.id"
              :value="option.id"
            >
              {{ option.label }}
            </option>
          </select>
        </div>
        <div
          v-if="d.hasOperator.value"
          class="validation-field"
        >
          <label for="validation-operator-control">Data</label>
          <select
            id="validation-operator-control"
            class="validation-control"
            data-testid="validation-operator"
            :value="f.operator"
            @change="d.set('operator', valueOf($event))"
          >
            <option
              v-for="option in VALIDATION_OPERATOR_OPTIONS"
              :key="option.id"
              :value="option.id"
            >
              {{ option.label }}
            </option>
          </select>
        </div>
        <div
          v-if="labels.value1 !== null"
          class="validation-field"
        >
          <label for="validation-value-1-control">{{ labels.value1 }}</label>
          <input
            id="validation-value-1-control"
            type="text"
            class="validation-control"
            data-testid="validation-value-1"
            :aria-invalid="invalid('value1')"
            :value="f.value1"
            @input="d.set('value1', valueOf($event))"
            @keydown.enter.prevent="d.ok()"
          >
        </div>
        <div
          v-if="labels.value2 !== null"
          class="validation-field"
        >
          <label for="validation-value-2-control">{{ labels.value2 }}</label>
          <input
            id="validation-value-2-control"
            type="text"
            class="validation-control"
            data-testid="validation-value-2"
            :aria-invalid="invalid('value2')"
            :value="f.value2"
            @input="d.set('value2', valueOf($event))"
            @keydown.enter.prevent="d.ok()"
          >
        </div>
        <div class="validation-check">
          <input
            id="validation-ignore-blank-control"
            type="checkbox"
            data-testid="validation-ignore-blank"
            :checked="f.allowBlank"
            @change="d.set('allowBlank', checked($event))"
          >
          <label for="validation-ignore-blank-control">Ignore blank</label>
        </div>
        <div
          v-if="f.type === 'list'"
          class="validation-check"
        >
          <input
            id="validation-in-cell-dropdown-control"
            type="checkbox"
            data-testid="validation-in-cell-dropdown"
            :checked="f.showDropdown"
            @change="d.set('showDropdown', checked($event))"
          >
          <label for="validation-in-cell-dropdown-control">In-cell dropdown</label>
        </div>
      </template>
      <template v-else-if="d.tab.value === 'input'">
        <div class="validation-check">
          <input
            id="validation-show-input-control"
            type="checkbox"
            data-testid="validation-show-input"
            :checked="f.inputShow"
            @change="d.set('inputShow', checked($event))"
          >
          <label for="validation-show-input-control">Show input message when cell is selected</label>
        </div>
        <div class="validation-field">
          <label for="validation-input-title-control">Title</label>
          <input
            id="validation-input-title-control"
            type="text"
            class="validation-control"
            data-testid="validation-input-title"
            :maxlength="INPUT_TITLE_MAX"
            :value="f.inputTitle"
            @input="d.set('inputTitle', valueOf($event))"
            @keydown.enter.prevent="d.ok()"
          >
        </div>
        <div class="validation-field">
          <label for="validation-input-text-control">Input message</label>
          <textarea
            id="validation-input-text-control"
            class="validation-control validation-text"
            data-testid="validation-input-text"
            rows="4"
            :maxlength="INPUT_TEXT_MAX"
            :value="f.inputText"
            @input="d.set('inputText', valueOf($event))"
          />
        </div>
      </template>
      <template v-else>
        <div class="validation-check">
          <input
            id="validation-show-error-control"
            type="checkbox"
            data-testid="validation-show-error"
            :checked="f.errorShow"
            @change="d.set('errorShow', checked($event))"
          >
          <label for="validation-show-error-control">Show error alert after invalid data is entered</label>
        </div>
        <div class="validation-field">
          <label for="validation-error-style-control">Style</label>
          <select
            id="validation-error-style-control"
            class="validation-control"
            data-testid="validation-error-style"
            :value="f.errorStyle"
            @change="d.set('errorStyle', valueOf($event))"
          >
            <option
              v-for="option in VALIDATION_ERROR_STYLE_OPTIONS"
              :key="option.id"
              :value="option.id"
            >
              {{ option.label }}
            </option>
          </select>
        </div>
        <div class="validation-field">
          <label for="validation-error-title-control">Title</label>
          <input
            id="validation-error-title-control"
            type="text"
            class="validation-control"
            data-testid="validation-error-title"
            :maxlength="ERROR_TITLE_MAX"
            :value="f.errorTitle"
            @input="d.set('errorTitle', valueOf($event))"
            @keydown.enter.prevent="d.ok()"
          >
        </div>
        <div class="validation-field">
          <label for="validation-error-text-control">Error message</label>
          <textarea
            id="validation-error-text-control"
            class="validation-control validation-text"
            data-testid="validation-error-text"
            rows="4"
            :maxlength="ERROR_TEXT_MAX"
            :value="f.errorText"
            @input="d.set('errorText', valueOf($event))"
          />
        </div>
      </template>
    </div>
    <p
      v-if="err"
      class="validation-error"
      role="alert"
      data-testid="validation-dialog-error"
    >
      {{ err.message }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        class="validation-clear-all"
        data-testid="validation-clear-all"
        @click="d.clearAll()"
      >
        Clear All
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="validation-cancel"
        @click="d.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="validation-ok"
        :disabled="d.busy.value"
        @click="d.ok()"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.validation-notice {
  margin: 0 0 var(--space-3);
  padding: var(--space-2) var(--space-3);
  color: var(--text-primary);
  font-size: var(--text-sm);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.validation-tabs {
  display: flex;
  gap: var(--space-1);
  border-bottom: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-3);
}

.validation-tab {
  font: inherit;
  color: var(--text-primary);
  padding: var(--space-1) var(--space-3);
  background: transparent;
  border: 0;
  border-bottom: calc(2 * var(--border-width)) solid transparent;
  cursor: pointer;
}

.validation-tab[aria-selected="true"] {
  border-bottom-color: var(--accent);
}

.validation-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  min-height: 16rem;
}

.validation-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.validation-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.validation-control {
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-sizing: border-box;
}

.validation-text {
  padding: var(--space-1) var(--space-2);
  resize: vertical;
}

.validation-control[aria-invalid="true"] {
  border-color: var(--danger);
}

.validation-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
}

/* Clear All sits apart, at the row's start, as in Excel. */
.validation-clear-all {
  margin-right: auto;
}
</style>
