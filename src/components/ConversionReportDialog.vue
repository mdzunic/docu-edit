<script setup lang="ts">
// SPEC-008: the conversion report shown once after a cross-format save,
// listing every item the writer could not represent in the target format.
// The list renders the writer's `dropped: DropRecord[]` directly — it is
// never derived by comparing models, so the report and the round-trip tests
// read the same array. Emits close only; the caller owns the "don't show
// again" persistence. BaseDialog owns the shell and Escape (which this
// dialog never handled); its title-testid prop carries the
// "conversion-report-title" testid this dialog's spec asserts on the h2.
import { computed, ref } from "vue";
import { asDropRecords, dropKindLabel } from "../lib/opaque";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** The writer's DropRecord[] — { kind, location, reason } per item. */
  items: { type: Array, required: true },
  /** SPEC-052: true for the one same-format case a save can still report —
   * an edited chart's lost properties (chartFeature). "Conversion report"
   * is the wrong heading/message when nothing converted. */
  sameFormat: { type: Boolean, default: false },
});

const emit = defineEmits(["close"]);

const dontShowAgain = ref(false);

const records = computed(() => asDropRecords(props.items));
</script>

<template>
  <BaseDialog
    :title="sameFormat ? 'Chart changes report' : 'Conversion report'"
    testid="conversion-report"
    title-testid="conversion-report-title"
    width="24rem"
    @dismiss="emit('close', dontShowAgain)"
  >
    <p class="dialog-message">
      <template v-if="sameFormat">
        Editing this chart rewrote its underlying data — some of its
        original formatting could not be carried over and was left out of
        the saved file.
      </template>
      <template v-else>
        Some content could not be converted to the saved format and was left
        out of the downloaded file. Your open document still has it.
      </template>
    </p>
    <ul
      class="report-list"
      data-testid="conversion-report-items"
    >
      <li
        v-for="(item, index) in records"
        :key="index"
        class="report-item"
        data-testid="conversion-report-item"
      >
        <span class="report-kind">{{ dropKindLabel(item.kind) }}</span>
        <span class="report-location">{{ item.location }}</span>
        <span class="report-reason">{{ item.reason }}</span>
      </li>
    </ul>
    <label class="report-suppress">
      <input
        v-model="dontShowAgain"
        type="checkbox"
        data-testid="conversion-report-dont-show"
      >
      Don't show again for this document
    </label>
    <template #actions>
      <button
        type="button"
        data-testid="conversion-report-close"
        @click="emit('close', dontShowAgain)"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* This dialog capped its panel narrower than the base's 90vw. */
:deep(.dialog) {
  max-width: min(36rem, 90vw);
}

.dialog-message {
  margin: 0;
  color: #3c4043;
  font-size: 0.9375rem;
}

.report-list {
  margin: 1rem 0 0;
  padding: 0;
  list-style: none;
  max-height: 40vh;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}

.report-item {
  display: flex;
  flex-direction: column;
  gap: 0.125rem;
  padding: 0.5rem 0.75rem;
  border-bottom: 1px solid #eee;
  font-size: 0.875rem;
}

.report-item:last-child {
  border-bottom: none;
}

.report-kind {
  font-weight: 600;
  color: #202124;
}

.report-location {
  color: #3c4043;
}

.report-reason {
  color: #5f6368;
  font-size: 0.8125rem;
}

.report-suppress {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-top: 1rem;
  font-size: 0.875rem;
  color: #3c4043;
  cursor: pointer;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
</style>
