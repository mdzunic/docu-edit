<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import {
  CHARACTER_CATEGORIES,
  displayName,
  isGlyphless,
  loadRecentCharacters,
  matchesQuery,
  pushRecentCharacter,
  findCharacter,
} from "../lib/core/specialCharacters";

// The "Insert → Special characters" picker (SPEC-026). A searchable grid of
// the fixed catalogue grouped by category, a most-recent-first Recent row and
// direct-insert buttons for the four hyphen/space specials. Clicking a
// character emits it for the parent to insert at the caret and keeps the
// dialog open (AC-5, AC-6). Runtime emit declaration per the espree
// constraint (see DownloadMenu.vue).
const emit = defineEmits(["insert", "cancel"]);

const query = ref("");
const recent = ref(loadRecentCharacters());
const searchInput = useTemplateRef("searchInput");

// The four direct-insert specials (AC-11..14). Non-breaking space also has a
// keyboard shortcut, wired in DocumentView.
const DIRECT_INSERTS = [
  { cp: 0x00a0, label: "Non-breaking space" },
  { cp: 0x2013, label: "En dash" },
  { cp: 0x2014, label: "Em dash" },
  { cp: 0x2011, label: "Non-breaking hyphen" },
];

onMounted(async () => {
  await nextTick();
  searchInput.value?.focus();
});

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
  <div
    class="dialog-backdrop"
    data-testid="special-character-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Special characters"
    >
      <div class="dialog-header">
        <h2 class="dialog-title">
          Special characters
        </h2>
        <button
          type="button"
          class="close-button"
          data-testid="special-character-close"
          aria-label="Close"
          @click="emit('cancel')"
        >
          ×
        </button>
      </div>

      <input
        ref="searchInput"
        v-model="query"
        class="search-input"
        type="search"
        placeholder="Search by name, keyword or code point (e.g. 2014)"
        aria-label="Search special characters"
        data-testid="special-character-search"
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
  width: 34rem;
  max-width: 92vw;
  max-height: 82vh;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}

.dialog-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.dialog-title {
  margin: 0 0 0.75rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.close-button {
  font: inherit;
  font-size: 1.25rem;
  line-height: 1;
  border: none;
  background: none;
  color: #5f6368;
  cursor: pointer;
  padding: 0 0.25rem;
}

.search-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  margin-bottom: 0.75rem;
}

.direct-inserts {
  display: flex;
  flex-wrap: wrap;
  gap: 0.375rem;
  margin-bottom: 0.75rem;
}

.direct-button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.direct-button:hover {
  background: #f1f3f4;
}

.scroll-area {
  overflow-y: auto;
  flex: 1;
}

.category {
  margin-bottom: 0.75rem;
}

.category-label {
  margin: 0 0 0.375rem;
  font-size: 0.75rem;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: #5f6368;
}

.grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(2rem, 1fr));
  gap: 0.25rem;
}

.char-cell {
  font: inherit;
  font-size: 1.125rem;
  aspect-ratio: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  padding: 0;
  overflow: hidden;
}

.char-cell:hover {
  background: #e8f0fe;
  border-color: #1a73e8;
}

.glyphless-box {
  width: 1rem;
  height: 1rem;
  border: 1px dotted #9e9e9e;
  border-radius: 2px;
  display: block;
}

.no-results {
  margin: 0.5rem 0;
  font-size: 0.875rem;
  color: #5f6368;
}
</style>
