<script setup lang="ts">
// "Insert → Caption" (SPEC-055 AC-2, AC-4..12): choose a label (Figure/Table/
// Equation or a custom one via "New label…"), a position (above/below), a
// separator, whether to exclude the label, and a number format. The parent
// inserts the Caption-styled paragraph. Runtime prop/emit declarations per the
// espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { asStringList } from "../lib/documentModel";
import { isValidNewLabel } from "../lib/captions";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** Every label the dialog offers: the three built-ins plus any custom
   * labels already used in the document. */
  labels: { type: Array, default: () => ["Figure", "Table", "Equation"] },
});

const emit = defineEmits(["insert", "cancel"]);

const labelList = ref([...asStringList(props.labels)]);
const label = ref(labelList.value[0] ?? "Figure");
const position = ref("below");
const separator = ref(": ");
const excludeLabel = ref(false);
const format = ref("decimal");

const addingLabel = ref(false);
const newLabel = ref("");
const newLabelValid = computed(() => isValidNewLabel(newLabel.value, labelList.value));

function startNewLabel() {
  addingLabel.value = true;
  newLabel.value = "";
}

function confirmNewLabel() {
  if (!newLabelValid.value) {
    return;
  }
  const name = newLabel.value.trim();
  labelList.value = [...labelList.value, name];
  label.value = name;
  addingLabel.value = false;
}

function insert() {
  emit("insert", {
    label: label.value,
    position: position.value,
    separator: separator.value,
    excludeLabel: excludeLabel.value,
    format: format.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Caption"
    testid="caption-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <label class="row">
      <span class="row-label">Label</span>
      <select
        v-model="label"
        class="row-input"
        data-testid="caption-label"
      >
        <option
          v-for="name in labelList"
          :key="name"
          :value="name"
        >
          {{ name }}
        </option>
      </select>
    </label>

    <div
      v-if="!addingLabel"
      class="row"
    >
      <span class="row-label" />
      <button
        type="button"
        class="link-button"
        data-testid="caption-new-label"
        @click="startNewLabel"
      >
        New label…
      </button>
    </div>
    <div
      v-else
      class="row"
    >
      <span class="row-label">New label</span>
      <span class="new-label-controls">
        <input
          v-model="newLabel"
          class="row-input"
          data-testid="caption-new-label-input"
          type="text"
          @keydown.enter.prevent="confirmNewLabel"
        >
        <button
          type="button"
          data-testid="caption-new-label-add"
          :disabled="!newLabelValid"
          @click="confirmNewLabel"
        >
          Add
        </button>
      </span>
    </div>

    <label class="row">
      <span class="row-label">Position</span>
      <select
        v-model="position"
        class="row-input"
        data-testid="caption-position"
      >
        <option value="below">
          Below selected item
        </option>
        <option value="above">
          Above selected item
        </option>
      </select>
    </label>

    <label class="row">
      <span class="row-label">Separator</span>
      <input
        v-model="separator"
        class="row-input"
        data-testid="caption-separator"
        type="text"
      >
    </label>

    <label class="row">
      <span class="row-label">Numbering</span>
      <select
        v-model="format"
        class="row-input"
        data-testid="caption-format"
      >
        <option value="decimal">
          1, 2, 3
        </option>
        <option value="lowerRoman">
          i, ii, iii
        </option>
        <option value="lowerAlpha">
          a, b, c
        </option>
      </select>
    </label>

    <label class="checkbox-row">
      <input
        v-model="excludeLabel"
        type="checkbox"
        data-testid="caption-exclude-label"
      >
      <span>Exclude label from caption</span>
    </label>

    <template #actions>
      <button
        type="button"
        data-testid="caption-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="caption-insert"
        @click="insert"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  font-size: 0.9375rem;
  margin-bottom: 0.75rem;
}

.row-label {
  color: #5f6368;
  min-width: 6rem;
}

.row-input {
  font: inherit;
  min-width: 12rem;
  flex: 1 1 auto;
}

.new-label-controls {
  display: flex;
  gap: 0.5rem;
  align-items: center;
  flex: 1 1 auto;
}

.new-label-controls .row-input {
  min-width: 0;
}

.link-button {
  font: inherit;
  background: none;
  border: none;
  color: #1a73e8;
  cursor: pointer;
  padding: 0;
}

.checkbox-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.9375rem;
  margin-bottom: 0.25rem;
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

.dialog-actions button:disabled,
.new-label-controls button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
