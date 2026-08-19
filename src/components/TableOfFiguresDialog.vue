<script setup lang="ts">
// "Insert → Table of figures" (SPEC-055): choose which caption label to list
// and the entry options (page numbers, dot leader, hyperlinks), mirroring
// TocOptionsDialog. The parent generates the block. Runtime prop/emit
// declarations per the espree constraint.
import { ref } from "vue";
import { asStringList } from "../lib/documentModel";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

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
      <UiButton
        variant="secondary"
        data-testid="tof-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="tof-insert"
        @click="apply"
      >
        Insert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  font-size: var(--text-lg);
  margin-bottom: var(--space-4);
}

.row-label {
  color: var(--text-secondary);
}

.row-input {
  font: inherit;
  min-width: 12rem;
}

.checkbox-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-lg);
  margin-bottom: var(--space-3);
}
</style>
