<script setup lang="ts">
// The custom-filter dialog (SPEC-128 "UI"): `filter-custom-dialog` on
// BaseDialog, opened from the filter popover's conditions submenu pre-set
// (Equals…, Between…, Custom Filter… on the column's current conditions).
// Two rows — `filter-custom-op-1` / `filter-custom-value-1`, then
// `filter-custom-op-2` / `filter-custom-value-2` (its operator may be left
// empty) — joined by `filter-custom-join` (And / Or). The operators are
// Excel's twelve, worded for the column's type (filterConditions.ts — "is
// before" on a date column). OK (`filter-custom-ok`, disabled until the first
// value is typed) emits confirm(CustomDraft); the view turns it into a
// `custom` criterion through parseInput. Cancel (`filter-custom-cancel`) and
// Escape emit cancel. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { computed, ref } from "vue";
import { asColumnKind, asCustomDraft, CUSTOM_OPERATORS, customOperatorLabel } from "../../views/spreadsheet/filterConditions";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The column's header text ("Region"). */
  title: { type: String, default: "" },
  /** "text" | "number" | "date": the operators' wording. */
  kind: { type: String, default: "text" },
  /** The pre-set conditions (CustomDraft). */
  draft: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["confirm", "cancel"]);

const initial = asCustomDraft(props.draft);
const op1 = ref(initial.op1);
const value1 = ref(initial.value1);
const join = ref(initial.join);
const op2 = ref(initial.op2);
const value2 = ref(initial.value2);

const operators = computed(() => CUSTOM_OPERATORS.map((op) => ({ op, label: customOperatorLabel(op, asColumnKind(props.kind)) })));
const valid = computed(() => value1.value !== "");

function confirm() {
  if (valid.value) {
    emit("confirm", asCustomDraft({ op1: op1.value, value1: value1.value, join: join.value, op2: op2.value, value2: value2.value }));
  }
}
</script>

<template>
  <BaseDialog
    title="Custom filter"
    testid="filter-custom-dialog"
    width="30rem"
    @dismiss="emit('cancel')"
  >
    <form
      class="custom-form"
      @submit.prevent="confirm"
    >
      <p class="custom-caption">
        Show rows where {{ props.title }}:
      </p>
      <div class="custom-row">
        <select
          v-model="op1"
          class="custom-control"
          aria-label="First condition"
          data-testid="filter-custom-op-1"
        >
          <option
            v-for="item in operators"
            :key="item.op"
            :value="item.op"
          >
            {{ item.label }}
          </option>
        </select>
        <input
          v-model="value1"
          class="custom-control"
          aria-label="First value"
          autofocus
          data-testid="filter-custom-value-1"
        >
      </div>
      <select
        v-model="join"
        class="custom-control custom-join"
        aria-label="Join"
        data-testid="filter-custom-join"
      >
        <option value="and">
          And
        </option>
        <option value="or">
          Or
        </option>
      </select>
      <div class="custom-row">
        <select
          v-model="op2"
          class="custom-control"
          aria-label="Second condition"
          data-testid="filter-custom-op-2"
        >
          <option value="" />
          <option
            v-for="item in operators"
            :key="item.op"
            :value="item.op"
          >
            {{ item.label }}
          </option>
        </select>
        <input
          v-model="value2"
          class="custom-control"
          aria-label="Second value"
          data-testid="filter-custom-value-2"
        >
      </div>
      <p class="custom-hint">
        Use ? to represent any single character, * to represent any series of characters.
      </p>
    </form>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="filter-custom-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="filter-custom-ok"
        :disabled="!valid"
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.custom-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.custom-caption,
.custom-hint {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.custom-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-2);
}

.custom-control {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.custom-join {
  justify-self: start;
  width: 6rem;
}
</style>
