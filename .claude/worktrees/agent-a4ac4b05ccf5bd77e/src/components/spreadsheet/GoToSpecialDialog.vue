<script setup lang="ts">
// Go To Special (SPEC-134), on BaseDialog: what to select — Blanks,
// Constants, Formulas, Errors, Last cell, Visible cells only
// (`sheet-goto-special-blanks`, `-constants`, `-formulas`, `-errors`,
// `-last-cell`, `-visible`, one radio group) — the type filters
// (`sheet-goto-type-numbers`, `-text`, `-logicals`, `-errors`; all checked by
// default, enabled for Constants and Formulas), Cancel and OK
// (`sheet-goto-special-ok`); a selection that finds nothing shows why in
// `sheet-goto-special-message` and leaves the dialog open. Shown while
// useSheetGoTo's `specialOpen`. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetGoTo } from "../../views/spreadsheet/useSheetGoTo";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useSheetGoTo's return. */
  goTo: { type: Object, required: true },
});

const g = computed(() => asSheetGoTo(props.goTo));

const KINDS = [
  { kind: "blanks", label: "Blanks", testid: "sheet-goto-special-blanks" },
  { kind: "constants", label: "Constants", testid: "sheet-goto-special-constants" },
  { kind: "formulas", label: "Formulas", testid: "sheet-goto-special-formulas" },
  { kind: "errors", label: "Errors", testid: "sheet-goto-special-errors" },
  { kind: "lastCell", label: "Last cell", testid: "sheet-goto-special-last-cell" },
  { kind: "visible", label: "Visible cells only", testid: "sheet-goto-special-visible" },
];

const TYPES = [
  { type: "numbers", label: "Numbers", testid: "sheet-goto-type-numbers" },
  { type: "text", label: "Text", testid: "sheet-goto-type-text" },
  { type: "logicals", label: "Logicals", testid: "sheet-goto-type-logicals" },
  { type: "errors", label: "Errors", testid: "sheet-goto-type-errors" },
];

function setKind(kind = "blanks") {
  if (kind === "constants" || kind === "formulas" || kind === "errors" || kind === "lastCell" || kind === "visible") {
    g.value.setKind(kind);
  } else {
    g.value.setKind("blanks");
  }
}

function setType(type = "numbers", event = new Event("change")) {
  const on = event.target instanceof HTMLInputElement && event.target.checked;
  if (type === "text" || type === "logicals" || type === "errors") {
    g.value.setType(type, on);
  } else {
    g.value.setType("numbers", on);
  }
}

function typeChecked(type = "numbers") {
  return Reflect.get(g.value.specialTypes.value, type) === true;
}
</script>

<template>
  <BaseDialog
    v-if="g.specialOpen.value"
    title="Go To Special"
    testid="sheet-goto-special-dialog"
    width="22rem"
    @dismiss="g.cancelSpecial()"
  >
    <fieldset class="special-group">
      <legend class="special-legend">
        Select
      </legend>
      <label
        v-for="(option, index) in KINDS"
        :key="option.kind"
        class="special-option"
      >
        <input
          type="radio"
          name="sheet-goto-special-kind"
          :value="option.kind"
          :data-testid="option.testid"
          :checked="g.specialKind.value === option.kind"
          :autofocus="index === 0 ? true : undefined"
          @change="setKind(option.kind)"
        >
        {{ option.label }}
      </label>
    </fieldset>
    <fieldset
      class="special-group special-types"
      :disabled="!g.typesEnabled.value"
    >
      <legend class="special-legend">
        Of type (constants and formulas)
      </legend>
      <label
        v-for="option in TYPES"
        :key="option.type"
        class="special-option"
      >
        <input
          type="checkbox"
          :data-testid="option.testid"
          :checked="typeChecked(option.type)"
          @change="setType(option.type, $event)"
        >
        {{ option.label }}
      </label>
    </fieldset>
    <p
      v-if="g.specialMessage.value"
      class="special-message"
      data-testid="sheet-goto-special-message"
      role="alert"
    >
      {{ g.specialMessage.value }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-goto-special-cancel"
        @click="g.cancelSpecial()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-goto-special-ok"
        @click="g.confirmSpecial()"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.special-group {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-2) var(--space-4);
  margin: 0 0 var(--space-3);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.special-legend {
  padding: 0 var(--space-1);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.special-option {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-primary);
  cursor: pointer;
}

.special-types:disabled .special-option {
  color: var(--text-disabled);
  cursor: default;
}

.special-message {
  margin: 0;
  color: var(--text-danger);
}
</style>
