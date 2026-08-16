<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// "Insert page number" (SPEC-029 AC-14..19): position (header/footer,
// left/centre/right), numbering format (1, 2, 3 / i, ii, iii / a, b, c)
// and an optional "Page X of Y". OK emits the settings for the parent to
// build the field runs and write them into the target section's
// header/footer (see commands.ts's insertPageNumberCommand); Cancel just
// closes. BaseDialog owns the shell, Escape and initial focus (the
// [autofocus] position select).
const props = defineProps({
  /** Prefill: { part: "header"|"footer" }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const part = ref(props.initial.part === "footer" ? "footer" : "header");
const position = ref("center");
const format = ref("decimal");
const pageXofY = ref(false);

function apply() {
  emit("apply", {
    part: part.value,
    position: position.value,
    format: format.value,
    pageXofY: pageXofY.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Insert page number"
    testid="page-number-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Position</span>
        <select
          v-model="part"
          class="field-input"
          data-testid="page-number-part"
          autofocus
        >
          <option value="header">Header</option>
          <option value="footer">Footer</option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Alignment</span>
        <select
          v-model="position"
          class="field-input"
          data-testid="page-number-position"
        >
          <option value="left">Left</option>
          <option value="center">Centre</option>
          <option value="right">Right</option>
        </select>
      </label>
    </div>
    <div class="field-row">
      <label class="field">
        <span class="field-label">Format</span>
        <select
          v-model="format"
          class="field-input"
          data-testid="page-number-format"
        >
          <option value="decimal">1, 2, 3</option>
          <option value="lowerRoman">i, ii, iii</option>
          <option value="lowerAlpha">a, b, c</option>
        </select>
      </label>
    </div>
    <label class="checkbox-field">
      <input
        v-model="pageXofY"
        type="checkbox"
        data-testid="page-number-of-total"
      >
      <span>Page X of Y</span>
    </label>
    <template #actions>
      <button
        type="button"
        data-testid="page-number-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="page-number-apply"
        @click="apply"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
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

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.8125rem;
  color: #3c4043;
  margin-bottom: 0.5rem;
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
</style>
