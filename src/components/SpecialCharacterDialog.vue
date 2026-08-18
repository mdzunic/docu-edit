<script setup lang="ts">
import { computed, ref } from "vue";
import {
  CHARACTER_CATEGORIES,
  displayName,
  isGlyphless,
  loadRecentCharacters,
  matchesQuery,
  pushRecentCharacter,
  findCharacter,
} from "../lib/core/specialCharacters";
import BaseDialog from "./BaseDialog.vue";

// The "Insert → Special characters" picker (SPEC-026). A searchable grid of
// the fixed catalogue grouped by category, a most-recent-first Recent row and
// direct-insert buttons for the four hyphen/space specials. Clicking a
// character emits it for the parent to insert at the caret and keeps the
// dialog open (AC-5, AC-6). BaseDialog owns the shell, Escape and initial
// focus (the [autofocus] search input); its #header-extra slot carries the
// × close button on the title row. Runtime emit declaration per the espree
// constraint (see DownloadMenu.vue).
const emit = defineEmits(["insert", "cancel"]);

const query = ref("");
const recent = ref(loadRecentCharacters());

// The four direct-insert specials (AC-11..14). Non-breaking space also has a
// keyboard shortcut, wired in DocumentView.
const DIRECT_INSERTS = [
  { cp: 0x00a0, label: "Non-breaking space" },
  { cp: 0x2013, label: "En dash" },
  { cp: 0x2014, label: "Em dash" },
  { cp: 0x2011, label: "Non-breaking hyphen" },
];

/** Categories with their characters filtered by the current query; empty
 * categories drop out. */
const filteredCategories = computed(() =>
  CHARACTER_CATEGORIES.map((category) => ({
    ...category,
    chars: category.chars.filter((c) => matchesQuery(c, query.value)),
  })).filter((category) => category.chars.length > 0),
);

const recentChars = computed(() =>
  recent.value.flatMap((cp) => {
    const c = findCharacter(cp);
    return c ? [c] : [];
  }),
);

const searching = computed(() => query.value.trim().length > 0);

function hex(cp = 0) {
  return cp.toString(16).toUpperCase();
}

function glyph(cp = 0) {
  return String.fromCodePoint(cp);
}

function pick(cp = 0) {
  emit("insert", String.fromCodePoint(cp));
  recent.value = pushRecentCharacter(cp, recent.value);
}
</script>

<template>
  <BaseDialog
    title="Special characters"
    testid="special-character-dialog"
    @dismiss="emit('cancel')"
  >
    <template #header-extra>
      <button
        type="button"
        class="close-button"
        data-testid="special-character-close"
        aria-label="Close"
        @click="emit('cancel')"
      >
        ×
      </button>
    </template>

    <input
      v-model="query"
      class="search-input"
      type="search"
      placeholder="Search by name, keyword or code point (e.g. 2014)"
      aria-label="Search special characters"
      data-testid="special-character-search"
      autofocus
    >

    <div
      class="direct-inserts"
      data-testid="special-character-direct"
    >
      <button
        v-for="entry in DIRECT_INSERTS"
        :key="entry.cp"
        type="button"
        class="direct-button"
        :data-testid="'special-char-direct-' + hex(entry.cp)"
        :title="entry.label"
        @mousedown.prevent
        @click="pick(entry.cp)"
      >
        {{ entry.label }}
      </button>
    </div>

    <div class="scroll-area">
      <section
        v-if="!searching && recentChars.length > 0"
        class="category"
        data-testid="special-character-recent"
      >
        <h3 class="category-label">
          Recent
        </h3>
        <div class="grid">
          <button
            v-for="c in recentChars"
            :key="'recent-' + c.cp"
            type="button"
            class="char-cell"
            :class="{ glyphless: isGlyphless(c.cp) }"
            :data-testid="'recent-char-' + hex(c.cp)"
            :title="displayName(c.name)"
            :aria-label="displayName(c.name)"
            @mousedown.prevent
            @click="pick(c.cp)"
          >
            <span
              v-if="isGlyphless(c.cp)"
              class="glyphless-box"
              aria-hidden="true"
            />
            <span
              v-else
              aria-hidden="true"
            >{{ glyph(c.cp) }}</span>
          </button>
        </div>
      </section>

      <p
        v-if="filteredCategories.length === 0"
        class="no-results"
        data-testid="special-character-empty"
      >
        No characters match “{{ query }}”.
      </p>

      <section
        v-for="category in filteredCategories"
        :key="category.id"
        class="category"
        :data-testid="'special-character-category-' + category.id"
      >
        <h3 class="category-label">
          {{ category.label }}
        </h3>
        <div class="grid">
          <button
            v-for="c in category.chars"
            :key="c.cp"
            type="button"
            class="char-cell"
            :class="{ glyphless: isGlyphless(c.cp) }"
            :data-testid="'special-char-' + hex(c.cp)"
            :title="displayName(c.name)"
            :aria-label="displayName(c.name)"
            @mousedown.prevent
            @click="pick(c.cp)"
          >
            <span
              v-if="isGlyphless(c.cp)"
              class="glyphless-box"
              aria-hidden="true"
            />
            <span
              v-else
              aria-hidden="true"
            >{{ glyph(c.cp) }}</span>
          </button>
        </div>
      </section>
    </div>
  </BaseDialog>
</template>

<style scoped>
/* Fixed 34rem panel, laid out as a column so .scroll-area flex-scrolls
   under the fixed search/direct-insert rows; keeps its slightly wider
   92vw cap (the base's is 90vw). */
:deep(.dialog) {
  width: 34rem;
  max-width: 92vw;
  max-height: 82vh;
  display: flex;
  flex-direction: column;
}

/* The title row sat tighter above the search input than the base's 1rem. */
:deep(.dialog-header) {
  margin-bottom: var(--space-4);
}

.close-button {
  font: inherit;
  font-size: var(--text-xl);
  line-height: 1;
  border: none;
  background: none;
  color: var(--text-secondary);
  cursor: pointer;
  padding: 0 var(--space-2);
}

.search-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  margin-bottom: var(--space-4);
}

.direct-inserts {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
}

.direct-button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
  white-space: nowrap;
}

.direct-button:hover {
  background: var(--surface-canvas);
}

.scroll-area {
  overflow-y: auto;
  flex: 1;
}

.category {
  margin-bottom: var(--space-4);
}

.category-label {
  margin: 0 0 var(--space-3);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-secondary);
}

.grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(2rem, 1fr));
  gap: var(--space-2);
}

.char-cell {
  font: inherit;
  font-size: var(--text-xl);
  aspect-ratio: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
  padding: 0;
  overflow: hidden;
}

.char-cell:hover {
  background: var(--surface-selected);
  border-color: var(--accent);
}

.glyphless-box {
  width: 1rem;
  height: 1rem;
  border: var(--border-width) dotted var(--border-strong);
  border-radius: var(--radius-sm);
  display: block;
}

.no-results {
  margin: var(--space-3) 0;
  font-size: var(--text-md);
  color: var(--text-secondary);
}
</style>
