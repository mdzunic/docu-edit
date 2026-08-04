<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";

// The "Custom spacing…" dialog (SPEC-014 AC-12..14): a line-spacing mode
// (multiple / exactly / at least) with its value, plus space before/after
// in points. OK emits the values for the parent to apply to every touched
// paragraph; Cancel just closes — the draft state lives here and dies here.
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

onMounted(async () => {
  await nextTick();
  valueInput.value?.focus();
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
  <div
    class="dialog-backdrop"
    data-testid="custom-spacing-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Custom spacing"
    >
      <h2 class="dialog-title">
        Custom spacing
      </h2>
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
      <div class="dialog-actions">
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
  min-width: 22rem;
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
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
  margin-top: 1.25rem;
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
