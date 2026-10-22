<script setup lang="ts">
// The library's left rail (the modernist handoff's full-tier layout): the
// product's create/open actions on top (through the `actions` slot, so this
// component never knows what kinds of documents the app can make), the
// section navigation (All files / Starred / Trash) with counts and the
// handoff's inset accent mark on the active section, and the storage meter
// at the bottom (through the `storage` slot). Rendered only at the full
// tier — HomeView keeps the single-column layout below it — so the trash
// link and the actions render exactly once whichever layout is active.
// `lib` is the LibraryController from useLibrary (runtime-typed Object —
// the .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
import { computed } from "vue";

const props = defineProps({
  lib: { type: Object, required: true },
});

const totalCount = computed(() => {
  const counts = props.lib.counts.value.formats;
  return Object.keys(counts).reduce((sum, key) => sum + counts[key], 0);
});

const allActive = computed(() => !props.lib.starredOnly.value);

function showAll() {
  if (props.lib.starredOnly.value) {
    props.lib.toggleStarredFilter();
  }
}

function showStarred() {
  if (!props.lib.starredOnly.value) {
    props.lib.toggleStarredFilter();
  }
}
</script>

<template>
  <aside
    class="library-rail"
    data-testid="library-rail"
    aria-label="Library sections"
  >
    <div class="rail-actions">
      <slot name="actions" />
    </div>

    <nav
      class="rail-sections"
      aria-label="Sections"
    >
      <button
        type="button"
        class="rail-item"
        data-testid="rail-all"
        :class="{ 'is-active': allActive }"
        :aria-pressed="allActive"
        @click="showAll"
      >
        <span>All files</span>
        <span class="rail-count">{{ totalCount }}</span>
      </button>
      <button
        type="button"
        class="rail-item"
        data-testid="rail-starred"
        :class="{ 'is-active': props.lib.starredOnly.value }"
        :aria-pressed="props.lib.starredOnly.value"
        @click="showStarred"
      >
        <span>Starred</span>
        <span class="rail-count">{{ props.lib.counts.value.starred }}</span>
      </button>
      <button
        type="button"
        class="rail-item trash-link"
        data-testid="trash-link"
        @click="props.lib.setShowTrash(true)"
      >
        <span>Trash</span>
        <span
          class="rail-count"
          data-testid="trash-count"
        >({{ props.lib.trashCount.value }})</span>
      </button>
    </nav>

    <div class="rail-storage">
      <slot name="storage" />
    </div>
  </aside>
</template>

<style scoped>
.library-rail {
  display: flex;
  flex-direction: column;
  gap: var(--space-6);
  padding-inline-end: var(--space-5);
  /* The handoff's heavier column rule. The pane's height comes from the
     layout that mounts it (HomeView pins it viewport-tall and sticky), not
     from the grid row — a full-row pane would leave sticky no travel. */
  border-inline-end: calc(2 * var(--border-width)) solid var(--border-default);
}

.rail-actions {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  align-items: stretch;
}

/* The slotted product actions stack full-width, reading as the handoff's
   left-aligned rail buttons whatever buttons the product provides. */
.rail-actions :slotted(button) {
  justify-content: flex-start;
  width: 100%;
}

.rail-sections {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.rail-item {
  font: inherit;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: none;
  /* Reserves the active mark's width so activation never shifts the label. */
  border-inline-start: calc(2 * var(--border-width)) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  text-align: start;
  cursor: pointer;
}

.rail-item:hover {
  background: var(--state-hover);
}

/* The handoff's active-section treatment: a tinted row with an accent mark
   on the reading edge. */
.rail-item.is-active {
  background: var(--surface-selected);
  border-inline-start-color: var(--accent);
  font-weight: var(--weight-bold);
}

.rail-count {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.rail-item.is-active .rail-count {
  color: var(--text-accent);
}

.rail-storage {
  margin-top: auto;
}
</style>
