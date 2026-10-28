<script setup lang="ts">
// SPEC-114 (experimental): the library search field's scope control —
// Titles (the default, SPEC-067's title matcher) or Everything (full-text,
// lib/libraryIndex.ts). Loaded only by LibrarySearchBox's ExperimentGate
// while the library-search experiment is on, so with the flag off it does
// not render and nothing of the search index is fetched. In Everything
// scope the library's own rows, chips and counts stop filtering by the
// field (`lib.setTitleSearch`), because the field now searches inside
// documents and the results replace the rows.
// `lib` is the LibraryController (runtime-typed Object — see useLibrary.ts).
import { computed, onBeforeUnmount, onMounted } from "vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import { useExperiments } from "../../composables/useExperiments";
import { useLibrarySearch } from "../../composables/useLibrarySearch";

const props = defineProps({
  lib: { type: Object, required: true },
});

const search = useLibrarySearch();
const { isEnabled } = useExperiments();

const scopeChoice = computed({
  get: () => search.scope.value,
  set: (value) => {
    const next = value === "everything" ? "everything" : "titles";
    props.lib.setTitleSearch(next === "titles");
    search.setScope(next);
  },
});

// The scope outlives this control (the tiers move the field between the app
// bar and the controls row): re-apply it to this library controller.
onMounted(() => props.lib.setTitleSearch(search.scope.value === "titles"));
// Switched off in Labs: back to SPEC-067's library, exactly.
onBeforeUnmount(() => {
  if (!isEnabled("library-search")) {
    search.setScope("titles");
    props.lib.setTitleSearch(true);
  }
});
</script>

<template>
  <span class="search-scope-wrap">
    <label
      class="visually-hidden"
      for="library-search-scope"
    >Search in</label>
    <select
      id="library-search-scope"
      v-model="scopeChoice"
      class="search-scope"
      data-testid="search-scope"
      title="Search titles, or everything inside your documents"
    >
      <option value="titles">
        Titles
      </option>
      <option value="everything">
        Everything
      </option>
    </select>
    <ExperimentBadge experiment="library-search" />
  </span>
</template>

<style scoped>
.search-scope-wrap {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.search-scope {
  font: inherit;
  font-size: var(--text-control);
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
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
