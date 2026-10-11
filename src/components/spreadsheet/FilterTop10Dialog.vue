<script setup lang="ts">
// The Top 10 dialog (SPEC-128 "UI"): `filter-top10-dialog` on BaseDialog,
// opened from Number Filters ▸ Top 10…. Show the `filter-top10-direction`
// (Top / Bottom) `filter-top10-value` (a whole number, 1–500 items or 1–100
// percent — out of range says so and disables OK) `filter-top10-unit`
// (Items / Percent). OK (`filter-top10-ok`) emits confirm({ top, percent,
// value }) — the view writes the `top10` criterion; Cancel
// (`filter-top10-cancel`) and Escape emit cancel. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { top10Criteria, TOP10_MAX } from "../../views/spreadsheet/filterConditions";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The column's header text ("Amount"). */
  title: { type: String, default: "" },
});

const emit = defineEmits(["confirm", "cancel"]);

const direction = ref("top");
const value = ref("10");
const unit = ref("items");

const max = computed(() => (unit.value === "percent" ? TOP10_MAX.percent : TOP10_MAX.items));
const valid = computed(() => top10Criteria(direction.value === "top", unit.value === "percent", value.value) !== null);

function confirm() {
  if (valid.value) {
    emit("confirm", { top: direction.value === "top", percent: unit.value === "percent", value: value.value });
  }
}
</script>

<template>
  <BaseDialog
    title="Top 10 filter"
    testid="filter-top10-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <form
      class="top10-form"
      @submit.prevent="confirm"
    >
      <p class="top10-caption">
        Show the rows of {{ props.title }}:
      </p>
      <div class="top10-row">
        <select
          v-model="direction"
          class="top10-control"
          aria-label="Top or bottom"
          data-testid="filter-top10-direction"
        >
          <option value="top">
            Top
          </option>
          <option value="bottom">
            Bottom
          </option>
        </select>
        <input
          v-model="value"
          class="top10-control top10-value"
          inputmode="numeric"
          aria-label="How many"
          autofocus
          :aria-invalid="valid ? undefined : 'true'"
          data-testid="filter-top10-value"
        >
        <select
          v-model="unit"
          class="top10-control"
          aria-label="Items or percent"
          data-testid="filter-top10-unit"
        >
          <option value="items">
            Items
          </option>
          <option value="percent">
            Percent
          </option>
        </select>
      </div>
      <p
        class="top10-hint"
        data-testid="filter-top10-hint"
        aria-live="polite"
      >
        {{ valid ? "" : `Enter a whole number from 1 to ${max}` }}
      </p>
    </form>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="filter-top10-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="filter-top10-ok"
        :disabled="!valid"
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.top10-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.top10-caption {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.top10-row {
  display: flex;
  gap: var(--space-2);
}

.top10-control {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.top10-value {
  width: 5rem;
}

.top10-control[aria-invalid="true"] {
  border-color: var(--danger);
}

.top10-hint {
  min-height: 1.25em;
  margin: 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
}
</style>
