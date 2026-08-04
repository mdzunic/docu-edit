<script setup lang="ts">
// SPEC-008: the conversion report shown once after a cross-format save,
// listing every item the writer could not represent in the target format.
// The list renders the writer's `dropped: DropRecord[]` directly — it is
// never derived by comparing models, so the report and the round-trip tests
// read the same array. Emits close only; the caller owns the "don't show
// again" persistence.
import { computed, ref } from "vue";
import { asDropRecords, dropKindLabel } from "../lib/opaque";

const props = defineProps({
  /** The writer's DropRecord[] — { kind, location, reason } per item. */
  items: { type: Array, required: true },
});

const emit = defineEmits(["close"]);

const dontShowAgain = ref(false);

const records = computed(() => asDropRecords(props.items));
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="conversion-report"
    @click.self="emit('close', dontShowAgain)"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Conversion report"
    >
      <h2 class="dialog-title">
        Conversion report
      </h2>
      <p class="dialog-message">
        Some content could not be converted to the saved format and was left
        out of the downloaded file. Your open document still has it.
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
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="conversion-report-close"
          @click="emit('close', dontShowAgain)"
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
  min-width: 24rem;
  max-width: min(36rem, 90vw);
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
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
</style>
