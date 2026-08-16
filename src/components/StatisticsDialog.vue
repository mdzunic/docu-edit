<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";

// The statistics dialog (SPEC-036 AC-4..12): pages, words, characters (with
// and without spaces), paragraphs, sentences and a separate CJK character
// count, opened by clicking the status bar's live word count. `stats` is
// lib/wordCount's CountStats for the current "Include footnotes and
// endnotes" setting — the parent recomputes it whenever `includeNotes`
// changes, which is what makes the toggle update the numbers live without
// closing the dialog (AC-12) with no extra state in here. BaseDialog owns
// the shell, Escape and initial focus (the [autofocus] Close button).
defineProps({
  stats: { type: Object, required: true },
  pages: { type: Number, required: true },
  includeNotes: { type: Boolean, default: false },
});

const emit = defineEmits(["update-include-notes", "close"]);

function onToggleIncludeNotes(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update-include-notes", target.checked);
  }
}
</script>

<template>
  <BaseDialog
    title="Word count"
    testid="statistics-dialog"
    @dismiss="emit('close')"
  >
    <dl class="stats-grid">
      <dt>Pages</dt>
      <dd data-testid="statistics-pages">
        {{ pages }}
      </dd>
      <dt>Words</dt>
      <dd data-testid="statistics-words">
        {{ stats.words }}
      </dd>
      <dt>Characters (no spaces)</dt>
      <dd data-testid="statistics-characters-without-spaces">
        {{ stats.charactersWithoutSpaces }}
      </dd>
      <dt>Characters (with spaces)</dt>
      <dd data-testid="statistics-characters-with-spaces">
        {{ stats.charactersWithSpaces }}
      </dd>
      <dt>Paragraphs</dt>
      <dd data-testid="statistics-paragraphs">
        {{ stats.paragraphs }}
      </dd>
      <dt>Sentences</dt>
      <dd data-testid="statistics-sentences">
        {{ stats.sentences }}
      </dd>
      <dt>CJK characters</dt>
      <dd data-testid="statistics-cjk-characters">
        {{ stats.cjkCharacters }}
      </dd>
    </dl>
    <label class="checkbox-field">
      <input
        :checked="includeNotes"
        type="checkbox"
        data-testid="statistics-include-notes"
        @change="onToggleIncludeNotes"
      >
      <span>Include footnotes and endnotes</span>
    </label>
    <template #actions>
      <button
        type="button"
        class="primary"
        data-testid="statistics-close"
        autofocus
        @click="emit('close')"
      >
        Close
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.stats-grid {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 0.375rem 1rem;
  margin: 0 0 1rem;
  font-size: 0.8125rem;
}

.stats-grid dt {
  color: #5f6368;
}

.stats-grid dd {
  margin: 0;
  text-align: right;
  color: #202124;
  font-variant-numeric: tabular-nums;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  font-size: 0.8125rem;
  color: #3c4043;
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
