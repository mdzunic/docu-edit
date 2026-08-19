<script setup lang="ts">
// The custom dictionary settings dialog (SPEC-035 AC-8.1..AC-8.3): view,
// search, add, remove and filter-by-language the persisted custom
// dictionary. Mirrors BookmarkManagerDialog.vue's list-with-add/remove
// shape; the add form's own validation (non-empty, no duplicate for the
// selected language) mirrors what lib/customDictionary.ts's addCustomWord
// already enforces, surfaced here as an inline error instead of a silent
// no-op. BaseDialog owns the shell and Escape — no field is marked
// [autofocus] (this dialog never focused one), so the panel itself takes
// focus and Escape works immediately.
import { computed, ref } from "vue";
import { asCustomDictionaryEntries } from "../lib/customDictionary";
import { asLanguageOptions } from "../lib/core/language";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** Every custom-dictionary entry, in insertion order. */
  entries: { type: Array, default: () => [] },
  /** The language options offered by the add form / filter (SUPPORTED_LANGUAGES). */
  languageOptions: { type: Array, default: () => [] },
  /** Preselects the add form's language — the document's current language
   * when opened from the status bar, so "Add word" defaults sensibly. */
  defaultLanguage: { type: String, default: "" },
});

const emit = defineEmits(["add", "remove", "close"]);

const entries = computed(() => asCustomDictionaryEntries(props.entries));
const languageOptions = computed(() => asLanguageOptions(props.languageOptions));

const searchText = ref("");
const languageFilter = ref("");
const newWord = ref("");
const newWordLanguage = ref(props.defaultLanguage);
const addError = ref("");

const filteredEntries = computed(() => {
  const search = searchText.value.trim().toLowerCase();
  const lang = languageFilter.value;
  return entries.value.filter((entry) => {
    if (lang && entry.lang !== lang) {
      return false;
    }
    if (search && !entry.word.toLowerCase().includes(search)) {
      return false;
    }
    return true;
  });
});

function submitAdd() {
  const word = newWord.value.trim();
  if (word.length === 0) {
    addError.value = "Enter a word to add";
    return;
  }
  const lang = newWordLanguage.value || props.defaultLanguage;
  const exists = entries.value.some(
    (entry) => entry.word.toLowerCase() === word.toLowerCase() && entry.lang === lang,
  );
  if (exists) {
    addError.value = "That word is already in the dictionary for this language";
    return;
  }
  addError.value = "";
  newWord.value = "";
  emit("add", word, lang);
}
</script>

<template>
  <BaseDialog
    title="Custom dictionary"
    testid="custom-dictionary-dialog"
    width="24rem"
    @dismiss="emit('close')"
  >
    <form
      class="add-form"
      @submit.prevent="submitAdd"
    >
      <input
        v-model="newWord"
        type="text"
        placeholder="Add a word…"
        aria-label="Word to add"
        data-testid="custom-dictionary-add-input"
      >
      <select
        v-model="newWordLanguage"
        aria-label="Language"
        data-testid="custom-dictionary-add-language"
      >
        <option
          v-for="option in languageOptions"
          :key="option.tag"
          :value="option.tag"
        >
          {{ option.label }}
        </option>
      </select>
      <button
        type="submit"
        data-testid="custom-dictionary-add-submit"
      >
        Add
      </button>
    </form>
    <p
      v-if="addError"
      class="error"
      data-testid="custom-dictionary-add-error"
    >
      {{ addError }}
    </p>

    <div class="filters">
      <input
        v-model="searchText"
        type="search"
        placeholder="Search…"
        aria-label="Search custom dictionary"
        data-testid="custom-dictionary-search"
      >
      <select
        v-model="languageFilter"
        aria-label="Filter by language"
        data-testid="custom-dictionary-language-filter"
      >
        <option value="">
          All languages
        </option>
        <option
          v-for="option in languageOptions"
          :key="option.tag"
          :value="option.tag"
        >
          {{ option.label }}
        </option>
      </select>
    </div>

    <p
      v-if="entries.length === 0"
      class="empty"
      data-testid="custom-dictionary-empty"
    >
      No custom words yet
    </p>
    <p
      v-else-if="filteredEntries.length === 0"
      class="empty"
      data-testid="custom-dictionary-no-matches"
    >
      No words match
    </p>
    <ul
      v-else
      class="entry-list"
    >
      <li
        v-for="entry in filteredEntries"
        :key="`${entry.lang}:${entry.word}`"
        class="entry-row"
        data-testid="custom-dictionary-row"
      >
        <span class="entry-word">{{ entry.word }}</span>
        <span class="entry-lang">{{ entry.lang }}</span>
        <button
          type="button"
          :aria-label="`Remove ${entry.word}`"
          data-testid="custom-dictionary-remove"
          @click="emit('remove', entry.word, entry.lang)"
        >
          Remove
        </button>
      </li>
    </ul>

    <template #actions>
      <button
        type="button"
        data-testid="custom-dictionary-close"
        @click="emit('close')"
      >
        Close
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.add-form {
  display: flex;
  gap: var(--space-3);
  margin-bottom: var(--space-3);
}

.add-form input[type="text"] {
  flex: 1;
}

.add-form input,
.add-form select,
.filters input,
.filters select {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.add-form button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.error {
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  color: var(--text-danger);
}

.filters {
  display: flex;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
}

.filters input {
  flex: 1;
}

.empty {
  margin: var(--space-3) 0;
  font-size: var(--text-md);
  color: var(--text-secondary);
}

.entry-list {
  list-style: none;
  margin: 0;
  padding: 0;
  max-height: 14rem;
  overflow-y: auto;
}

.entry-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) 0;
}

.entry-word {
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: var(--text-md);
}

.entry-lang {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.entry-row button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.entry-row button:hover {
  background: var(--surface-canvas);
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}
</style>
