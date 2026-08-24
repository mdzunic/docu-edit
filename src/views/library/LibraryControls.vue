<script setup lang="ts">
// The library's header controls (SPEC-067/068): heading, search, sort, the
// list/grid toggle, the trash link, the format/starred filter chips — the
// product-neutral chrome of the library module. The product's own create/
// open actions render through the `actions` slot, so this component never
// knows what kinds of documents the app can make.
// `lib` is the LibraryController from useLibrary (runtime-typed Object —
// the .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
import { ref, watch } from "vue";
import { formatBadgeLabel } from "../../lib/format";
import { FILTERABLE_FORMATS } from "../../lib/libraryView";

const props = defineProps({
  lib: { type: Object, required: true },
});

// A plain string ref drives the sort <select> (v-model), avoiding a TS `as`
// cast on the change event (barred in .vue). String(...) forces Ref<string> so
// v-model can assign to it; setSortKey validates the value.
const sortKeyChoice = ref(String(props.lib.sort.value.key));
watch(sortKeyChoice, (value) => props.lib.setSortKey(value));
</script>

<template>
  <div class="library-header">
    <div class="header-line">
      <h2 class="library-heading">
        Documents
      </h2>
      <input
        v-model="props.lib.search.value"
        type="search"
        class="library-search"
        data-testid="library-search"
        placeholder="Search documents…"
        aria-label="Search documents"
      >
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
      <div class="header-spacer" />
      <button
        type="button"
        class="trash-link"
        data-testid="trash-link"
        @click="props.lib.showTrash.value = true"
      >
        Trash <span
          class="trash-count"
          data-testid="trash-count"
        >({{ props.lib.trashCount.value }})</span>
      </button>
      <div class="library-actions">
        <slot name="actions" />
      </div>
    </div>
    <div
      class="filter-chips"
      role="group"
      aria-label="Filters"
    >
      <button
        v-for="format in FILTERABLE_FORMATS"
        :key="format"
        type="button"
        class="chip"
        :data-testid="`filter-${format}`"
        :class="{ 'is-active': props.lib.activeFormats.value.has(format) }"
        :aria-pressed="props.lib.activeFormats.value.has(format)"
        @click="props.lib.toggleFormat(format)"
      >
        {{ formatBadgeLabel(format) }}
        <span class="chip-count">{{ props.lib.counts.value.formats[format] }}</span>
      </button>
      <button
        type="button"
        class="chip"
        data-testid="filter-starred"
        :class="{ 'is-active': props.lib.starredOnly.value }"
        :aria-pressed="props.lib.starredOnly.value"
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

.library-heading {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.library-search {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  min-width: 12rem;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
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

/* SPEC-068 AC-37/39 — at the compact tier (<640px, matching
   useResponsiveTier) the search box takes its own full-width row so it
   always stays visible. */
@media (max-width: 639.98px) {
  .library-search {
    flex: 1 1 100%;
    min-width: 0;
    order: 1;
  }
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
