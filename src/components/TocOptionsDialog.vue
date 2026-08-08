<script setup lang="ts">
import { ref } from "vue";

// "Insert table of contents" options (SPEC-033 AC-5..7): which heading
// levels to include, whether to show page numbers, and whether entries are
// plain dot-leader-and-page-number rows or hyperlinks only (Word's own "Use
// hyperlinks instead of page numbers" — no dot leader, no page number).
// Apply emits the settings for the parent to generate the block from
// (lib/toc.ts's insertTableOfContents); Cancel just closes.
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
  <div
    class="dialog-backdrop"
    data-testid="toc-options-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Table of contents"
    >
      <h2 class="dialog-title">
        Table of contents
      </h2>
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
      <div class="dialog-actions">
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

.field-group {
  border: none;
  padding: 0;
  margin: 0 0 0.75rem;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0.25rem 0.75rem;
}

.field-label {
  font-size: 0.8125rem;
  color: #3c4043;
  padding: 0;
  grid-column: 1 / -1;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.8125rem;
  color: #3c4043;
  margin-bottom: 0.5rem;
}

.level-field {
  margin-bottom: 0.25rem;
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
</style>
