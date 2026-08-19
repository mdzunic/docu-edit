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
import UiButton from "./ui/UiButton.vue";

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
      <UiButton
        variant="secondary"
        data-testid="date-time-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="date-time-insert"
        :disabled="!selected"
        @click="insert"
      >
        Insert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.hint {
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.format-list {
  list-style: none;
  margin: 0 0 var(--space-4);
  padding: 0;
  max-height: 16rem;
  overflow-y: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.format-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-3);
  cursor: pointer;
  font-size: var(--text-lg);
}

.format-row:hover {
  background: var(--surface-canvas);
}

.auto-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-lg);
}
</style>
