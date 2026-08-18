<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// "Insert table of contents" options (SPEC-033 AC-5..7): which heading
// levels to include, whether to show page numbers, and whether entries are
// plain dot-leader-and-page-number rows or hyperlinks only (Word's own "Use
// hyperlinks instead of page numbers" — no dot leader, no page number).
// Apply emits the settings for the parent to generate the block from
// (lib/toc.ts's insertTableOfContents); Cancel just closes. BaseDialog owns
// the shell and Escape — no control is marked [autofocus] (this dialog
// never focused one), so the panel itself takes focus and Escape works
// immediately.
const props = defineProps({
  /** Prefill: { levels: number[], showPageNumbers: boolean, hyperlinksOnly: boolean }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const LEVELS = [1, 2, 3, 4, 5, 6];
const initialLevels = new Set(
  Array.isArray(props.initial.levels) && props.initial.levels.length > 0 ? props.initial.levels : [1, 2, 3],
);
const levelChecked = ref(LEVELS.map((level) => initialLevels.has(level)));
const showPageNumbers = ref(props.initial.showPageNumbers !== false);
const hyperlinksOnly = ref(props.initial.hyperlinksOnly === true);

function toggleLevel(index = 0) {
  const next = [...levelChecked.value];
  next[index] = !next[index];
  levelChecked.value = next;
}

function apply() {
  const levels = LEVELS.filter((_, index) => levelChecked.value[index]);
  emit("apply", {
    levels: levels.length > 0 ? levels : [1],
    showPageNumbers: showPageNumbers.value,
    hyperlinksOnly: hyperlinksOnly.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Table of contents"
    testid="toc-options-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <fieldset class="field-group">
      <legend class="field-label">
        Show levels
      </legend>
      <label
        v-for="(level, index) in LEVELS"
        :key="level"
        class="checkbox-field level-field"
      >
        <input
          type="checkbox"
          :checked="levelChecked[index]"
          :data-testid="`toc-level-${level}`"
          @change="toggleLevel(index)"
        >
        <span>Heading {{ level }}</span>
      </label>
    </fieldset>
    <label class="checkbox-field">
      <input
        v-model="showPageNumbers"
        type="checkbox"
        data-testid="toc-show-page-numbers"
      >
      <span>Show page numbers</span>
    </label>
    <label class="checkbox-field">
      <input
        v-model="hyperlinksOnly"
        type="checkbox"
        data-testid="toc-hyperlinks-only"
      >
      <span>Use hyperlinks instead of page numbers</span>
    </label>
    <template #actions>
      <button
        type="button"
        data-testid="toc-options-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="toc-options-apply"
        @click="apply"
      >
        Insert
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field-group {
  border: none;
  padding: 0;
  margin: 0 0 var(--space-4);
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-2) var(--space-4);
}

.field-label {
  font-size: var(--text-md);
  color: var(--text-primary);
  padding: 0;
  grid-column: 1 / -1;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  font-size: var(--text-md);
  color: var(--text-primary);
  margin-bottom: var(--space-3);
}

.level-field {
  margin-bottom: var(--space-2);
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
</style>
