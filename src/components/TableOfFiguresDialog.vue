<script setup lang="ts">
// "Insert → Table of figures" (SPEC-055): choose which caption label to list
// and the entry options (page numbers, dot leader, hyperlinks), mirroring
// TocOptionsDialog. The parent generates the block. Runtime prop/emit
// declarations per the espree constraint.
import { ref } from "vue";
import { asStringList } from "../lib/documentModel";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** The caption labels the document has, e.g. ["Figure", "Table"]. */
  labels: { type: Array, default: () => ["Figure"] },
});

const emit = defineEmits(["apply", "cancel"]);

const labelList = asStringList(props.labels);
const label = ref(labelList[0] ?? "Figure");
const pageNumbers = ref(true);
const dotLeader = ref(true);
const hyperlinks = ref(true);

function apply() {
  emit("apply", {
    label: label.value,
    options: {
      pageNumbers: pageNumbers.value,
      dotLeader: dotLeader.value,
      hyperlinks: hyperlinks.value,
    },
  });
}
</script>

<template>
  <BaseDialog
    title="Table of figures"
    testid="table-of-figures-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label class="row">
      <span class="row-label">Caption label</span>
      <select
        v-model="label"
        class="row-input"
        data-testid="tof-label"
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

    <label class="checkbox-row">
      <input
        v-model="pageNumbers"
        type="checkbox"
        data-testid="tof-page-numbers"
      >
      <span>Show page numbers</span>
    </label>
    <label class="checkbox-row">
      <input
        v-model="dotLeader"
        type="checkbox"
        data-testid="tof-dot-leader"
      >
      <span>Right-align with dot leader</span>
    </label>
    <label class="checkbox-row">
      <input
        v-model="hyperlinks"
        type="checkbox"
        data-testid="tof-hyperlinks"
      >
      <span>Entries link to their captions</span>
    </label>

    <template #actions>
      <button
        type="button"
        data-testid="tof-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="tof-insert"
        @click="apply"
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
}

.row-input {
  font: inherit;
  min-width: 12rem;
}

.checkbox-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.9375rem;
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
