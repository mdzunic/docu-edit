<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import BaseDialog from "./BaseDialog.vue";

// The "Custom spacing…" dialog (SPEC-014 AC-12..14): a line-spacing mode
// (multiple / exactly / at least) with its value, plus space before/after
// in points. OK emits the values for the parent to apply to every touched
// paragraph; Cancel just closes — the draft state lives here and dies here.
// BaseDialog owns the shell, Escape and initial focus (the [autofocus]
// value input).
const props = defineProps({
  /** Prefill from the first touched paragraph: { mode, value, beforePt,
   * afterPt } with mode "multiple" | "exact" | "atLeast". */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const mode = ref(String(props.initial.mode ?? "multiple"));
const value = ref(String(props.initial.value ?? 1));
const beforePt = ref(String(props.initial.beforePt ?? 0));
const afterPt = ref(String(props.initial.afterPt ?? 0));

const valueInput = useTemplateRef("valueInput");

// BaseDialog focuses the [autofocus] value input; selecting the prefilled
// value on top of that (type-to-replace) is this dialog's own one-shot
// behaviour.
onMounted(async () => {
  await nextTick();
  valueInput.value?.select();
});

const isMultiple = computed(() => mode.value === "multiple");

// Word's own bounds: multiples 0.05..132 lines, absolute values and
// paragraph spacing 0..1584 pt (absolute line heights at least 1 pt).
const valueError = computed(() => {
  const parsed = Number(value.value);
  if (!Number.isFinite(parsed)) {
    return "Enter a number";
  }
  if (isMultiple.value && (parsed < 0.05 || parsed > 132)) {
    return "Line spacing must be between 0.05 and 132";
  }
  if (!isMultiple.value && (parsed < 1 || parsed > 1584)) {
    return "Line height must be between 1 and 1584 pt";
  }
  return "";
});

const spacingError = computed(() => {
  for (const raw of [beforePt.value, afterPt.value]) {
    const parsed = Number(raw);
    if (!Number.isFinite(parsed) || parsed < 0 || parsed > 1584) {
      return "Space before/after must be between 0 and 1584 pt";
    }
  }
  return "";
});

const error = computed(() => valueError.value || spacingError.value);

/** When the mode flips between a multiple and an absolute height the old
 * value is meaningless in the new unit, so it resets to that unit's
 * default; switching between "exactly" and "at least" keeps it — both are
 * point heights. */
let previousWasMultiple = mode.value === "multiple";
function handleModeChange() {
  if (isMultiple.value !== previousWasMultiple) {
    value.value = isMultiple.value ? "1" : "12";
    previousWasMultiple = isMultiple.value;
  }
}

function apply() {
  if (error.value) {
    return;
  }
  emit("apply", {
    mode: mode.value,
    value: Number(value.value),
    beforePt: Number(beforePt.value),
    afterPt: Number(afterPt.value),
  });
}
</script>

<template>
  <BaseDialog
    title="Custom spacing"
    testid="custom-spacing-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Line spacing</span>
        <select
          v-model="mode"
          class="field-input"
          data-testid="spacing-mode"
          @change="handleModeChange"
        >
          <option value="multiple">Multiple</option>
          <option value="exact">Exactly</option>
          <option value="atLeast">At least</option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">{{ isMultiple ? "Lines" : "Points" }}</span>
        <input
          ref="valueInput"
          v-model="value"
          class="field-input value-input"
          data-testid="spacing-value"
          type="number"
          :step="isMultiple ? 0.05 : 0.5"
          :aria-label="isMultiple ? 'Line spacing in lines' : 'Line height in points'"
          autofocus
          @keydown.enter="apply"
        >
      </label>
    </div>
    <div class="field-row">
      <label class="field">
        <span class="field-label">Space before (pt)</span>
        <input
          v-model="beforePt"
          class="field-input value-input"
          data-testid="spacing-before"
          type="number"
          min="0"
          step="0.5"
          @keydown.enter="apply"
        >
      </label>
      <label class="field">
        <span class="field-label">Space after (pt)</span>
        <input
          v-model="afterPt"
          class="field-input value-input"
          data-testid="spacing-after"
          type="number"
          min="0"
          step="0.5"
          @keydown.enter="apply"
        >
      </label>
    </div>
    <p
      v-if="error"
      class="dialog-error"
      data-testid="spacing-error"
    >
      {{ error }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="spacing-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="spacing-apply"
        :disabled="!!error"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
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

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
