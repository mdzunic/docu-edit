<script setup lang="ts">
// "Insert → Date and time" (SPEC-053 AC-2..5): offers the closed catalogue of
// date/time formats, each previewed in the document's language, with an
// "Update automatically" toggle. Off inserts the rendered text as plain text
// (AC-4); on inserts a live `date`/`time` field (AC-5). The parent owns the
// clock instant, so the preview it computes and the value it inserts agree.
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { asDateTimePreviews } from "../lib/fields";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** Each catalogue format already rendered for the preview column, in the
     document's language at the parent's clock instant:
     `[{ format, label, isTime }]`. */
  previews: { type: Array, default: () => [] },
});

const emit = defineEmits(["insert", "cancel"]);

const previews = computed(() => asDateTimePreviews(props.previews));
const selected = ref(previews.value.length > 0 ? previews.value[0].format : "");
const updateAutomatically = ref(false);

function insert() {
  if (!selected.value) {
    return;
  }
  emit("insert", { format: selected.value, updateAutomatically: updateAutomatically.value });
}
</script>

<template>
  <BaseDialog
    title="Date and time"
    testid="date-time-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <p class="hint">
      Choose a format:
    </p>
    <ul
      class="format-list"
      role="listbox"
      aria-label="Date and time formats"
      data-testid="date-time-formats"
    >
      <li
        v-for="preview in previews"
        :key="preview.format"
      >
        <label class="format-row">
          <input
            v-model="selected"
            type="radio"
            name="date-time-format"
            :value="preview.format"
            :data-testid="`date-time-format-${preview.format}`"
          >
          <span class="format-label">{{ preview.label }}</span>
        </label>
      </li>
    </ul>

    <label class="auto-row">
      <input
        v-model="updateAutomatically"
        type="checkbox"
        data-testid="date-time-update-automatically"
      >
      Update automatically
    </label>

    <template #actions>
      <button
        type="button"
        data-testid="date-time-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="date-time-insert"
        :disabled="!selected"
        @click="insert"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.hint {
  margin: 0 0 0.5rem;
  font-size: 0.875rem;
  color: #5f6368;
}

.format-list {
  list-style: none;
  margin: 0 0 0.75rem;
  padding: 0;
  max-height: 16rem;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}

.format-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.35rem 0.6rem;
  cursor: pointer;
  font-size: 0.9375rem;
}

.format-row:hover {
  background: #f1f3f4;
}

.auto-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.9375rem;
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
