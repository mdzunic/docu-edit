<script setup lang="ts">
// The library's header controls (SPEC-067/068): heading, search, sort, the
// list/grid toggle, the trash link, the format/starred filter chips — the
// product-neutral chrome of the library module. The product's own create/
// open actions render through the `actions` slot, so this component never
// knows what kinds of documents the app can make.
// In the full-tier rail layout (the modernist handoff) the rail owns the
// trash link and the create/open actions, and the app bar owns the search
// box; the `rail` prop suppresses those here so each renders exactly once.
// `lib` is the LibraryController from useLibrary (runtime-typed Object —
// the .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
import { computed, ref, watch } from "vue";
import { chipActive, chipCount, chipUnavailable, FORMAT_CHIPS } from "../../lib/libraryView";
import LibrarySearchBox from "./LibrarySearchBox.vue";

const props = defineProps({
  lib: { type: Object, required: true },
  rail: { type: Boolean, default: false },
});

// A plain string ref drives the sort <select> (v-model), avoiding a TS `as`
// cast on the change event (barred in .vue). String(...) forces Ref<string> so
// v-model can assign to it; setSortKey validates the value.
const sortKeyChoice = ref(String(props.lib.sort.value.key));
watch(sortKeyChoice, (value) => props.lib.setSortKey(value));

// The handoff's section subtitle: how many entries the current search/filters
// leave visible, out of everything the library holds.
const subtitle = computed(() => {
  const shown = props.lib.visible.value.length;
  const counts = props.lib.counts.value.formats;
  const total = Object.keys(counts).reduce((sum, key) => sum + counts[key], 0);
  return props.lib.filterActive.value ? `${shown} of ${total} shown` : `${total} in the library`;
});
</script>

<template>
  <div class="library-header">
    <div class="header-line">
      <div class="heading-block">
        <h2 class="library-heading">
          Documents
        </h2>
        <div
          class="library-subtitle"
          data-testid="library-subtitle"
        >
          {{ subtitle }}
        </div>
      </div>
      <LibrarySearchBox
        v-if="!props.rail"
        :lib="props.lib"
      />
      <!-- In the rail layout (search in the app bar, actions in the rail) the
           sort and view controls sit flush right, as in the handoff. -->
      <div
        v-if="props.rail"
        class="header-spacer"
      />
      <div class="sort-control">
        <label
          class="visually-hidden"
          for="library-sort"
        >Sort by</label>
        <select
          id="library-sort"
          v-model="sortKeyChoice"
          class="sort-select"
          data-testid="sort-key"
        >
          <option value="modified">
            Last modified
          </option>
          <option value="title">
            Title
          </option>
          <option value="size">
            Size
          </option>
        </select>
        <button
          type="button"
          class="sort-direction"
          data-testid="sort-direction"
          :aria-label="props.lib.sort.value.direction === 'asc' ? 'Ascending' : 'Descending'"
          @click="props.lib.toggleSortDirection()"
        >
          {{ props.lib.sort.value.direction === "asc" ? "↑" : "↓" }}
        </button>
      </div>
      <div
        class="view-toggle"
        role="group"
        aria-label="View"
      >
        <button
          type="button"
          class="toggle-button"
          data-testid="view-list"
          :class="{ 'is-active': props.lib.view.value === 'list' }"
          :aria-pressed="props.lib.view.value === 'list'"
          @click="props.lib.setView('list')"
        >
          List
        </button>
        <button
          type="button"
          class="toggle-button"
          data-testid="view-grid"
          :class="{ 'is-active': props.lib.view.value === 'grid' }"
          :aria-pressed="props.lib.view.value === 'grid'"
          @click="props.lib.setView('grid')"
        >
          Grid
        </button>
      </div>
      <div
        v-if="!props.rail"
        class="header-spacer"
      />
      <button
        v-if="!props.rail"
        type="button"
        class="trash-link"
        data-testid="trash-link"
        @click="props.lib.setShowTrash(true)"
      >
        Trash <span
          class="trash-count"
          data-testid="trash-count"
        >({{ props.lib.trashCount.value }})</span>
      </button>
      <div
        v-if="!props.rail"
        class="library-actions"
      >
        <slot name="actions" />
      </div>
    </div>
    <div
      class="filter-chips"
      role="group"
      aria-label="Filters"
    >
      <!-- One chip per document format, plus the KIND-grouped media chips
           ("Images", "PDFs") so each is one click (SPEC-107). -->
      <button
        v-for="chip in FORMAT_CHIPS"
        :key="chip.id"
        type="button"
        class="chip"
        :data-testid="`filter-${chip.id}`"
        :class="{ 'is-active': chipActive(chip, props.lib.activeFormats.value) }"
        :aria-pressed="chipActive(chip, props.lib.activeFormats.value)"
        :disabled="chipUnavailable(chip, props.lib.activeFormats.value, props.lib.counts.value.formats)"
        @click="props.lib.toggleFormatChip(chip)"
      >
        {{ chip.label }}
        <span class="chip-count">{{ chipCount(chip, props.lib.counts.value.formats) }}</span>
      </button>
      <button
        type="button"
        class="chip"
        data-testid="filter-starred"
        :class="{ 'is-active': props.lib.starredOnly.value }"
        :aria-pressed="props.lib.starredOnly.value"
        :disabled="props.lib.counts.value.starred === 0 && !props.lib.starredOnly.value"
        @click="props.lib.toggleStarredFilter()"
      >
        ★ Starred
        <span class="chip-count">{{ props.lib.counts.value.starred }}</span>
      </button>
    </div>
  </div>
</template>

<style scoped>
.library-header {
  margin-bottom: var(--space-5);
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.header-line {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.header-spacer {
  flex: 1 1 auto;
}

.heading-block {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.library-heading {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.library-subtitle {
  font-size: var(--text-sm);
  color: var(--text-secondary);
  white-space: nowrap;
}

.sort-control {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.sort-select {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.sort-direction,
.toggle-button,
.trash-link {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.view-toggle {
  display: inline-flex;
}

.view-toggle .toggle-button:first-child {
  border-top-right-radius: 0;
  border-bottom-right-radius: 0;
}

.view-toggle .toggle-button:last-child {
  border-top-left-radius: 0;
  border-bottom-left-radius: 0;
  border-left: none;
}

.toggle-button.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.trash-count {
  color: var(--text-secondary);
}

.library-actions {
  display: flex;
  gap: var(--space-3);
}

.filter-chips {
  display: flex;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.chip {
  font: inherit;
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-pill);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

/* A filter that stands for nothing in this library: still listed, so the set
   of filters stays a stable map of what the library can hold, but no longer
   offering a click whose only outcome is an empty list. */
.chip:disabled {
  color: var(--text-disabled);
  border-color: var(--border-subtle);
  cursor: default;
}

.chip.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  border-color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.chip-count {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.chip.is-active .chip-count {
  color: var(--text-accent);
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
