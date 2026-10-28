<script setup lang="ts">
// The library's search field (SPEC-067), extracted so the full-tier layout
// can host it in the app bar (the modernist handoff's centred header search)
// while the medium/compact tiers keep it in the controls row — exactly one
// instance renders at a time, so `library-search` stays a unique testid.
// `lib` is the LibraryController from useLibrary (runtime-typed Object — the
// .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
// SPEC-114's scope control (Titles / Everything) follows the field, behind
// its experiment's gate: with library-search off it renders nothing and its
// code is never loaded.
import { computed } from "vue";
import ExperimentGate from "../../components/ExperimentGate.vue";

const loadScope = () => import("./LibrarySearchScope.vue");

const props = defineProps({
  lib: { type: Object, required: true },
});

// v-model routes through the controller (vue/no-mutating-props) and stays in
// sync when the controller clears the search (e.g. "Clear filters").
const searchText = computed({
  get: () => String(props.lib.search.value),
  set: (value) => props.lib.setSearch(value),
});
</script>

<template>
  <input
    v-model="searchText"
    type="search"
    class="library-search"
    data-testid="library-search"
    placeholder="Search documents…"
    aria-label="Search documents"
  >
  <ExperimentGate
    experiment="library-search"
    :load="loadScope"
    :lib="props.lib"
  />
</template>

<style scoped>
.library-search {
  font: inherit;
  font-size: var(--text-control);
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  min-width: 12rem;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
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
</style>
